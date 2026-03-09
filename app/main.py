"""FastAPI アプリケーション: メール抽出 & CSV アップロード → ネットワーク可視化."""

import json
import logging
import sys
from pathlib import Path

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

from app.core import (
    build_graph,
    analyze_graph,
    generate_vis_data,
    load_config,
    load_csv,
)

# === Log to file (server.log) ===
LOG_FILE = Path(__file__).resolve().parent.parent / "server.log"
_logfile = open(LOG_FILE, "w", encoding="utf-8")

def _log(msg):
    """Print to both console and log file."""
    line = str(msg)
    print(line, flush=True)
    _logfile.write(line + "\n")
    _logfile.flush()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

app = FastAPI(title="Dot-connect", description="メールネットワーク可視化")


@app.middleware("http")
async def log_all_requests(request: Request, call_next):
    _log(f"[REQ] {request.method} {request.url.path}")
    response = await call_next(request)
    _log(f"[RES] {request.method} {request.url.path} -> {response.status_code}")
    return response


@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    _log(f"[VALIDATION ERROR] {request.method} {request.url.path}: {exc.errors()}")
    return RedirectResponse(url=f"/?error=入力エラー: {exc.errors()}", status_code=303)


TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "templates"
templates = Jinja2Templates(directory=str(TEMPLATE_DIR))


def _default_template_vars() -> dict:
    """upload.html に渡すデフォルト値."""
    default_config = load_config()
    thresholds = default_config.get("thresholds", {})
    return {
        "domains": ", ".join(default_config.get("company_domains", [])),
        "cc_key_person_threshold": thresholds.get("cc_key_person_threshold", 0.30),
        "min_edge_weight": thresholds.get("min_edge_weight", 1),
        "hub_degree_weight": thresholds.get("hub_degree_weight", 0.5),
        "hub_betweenness_weight": thresholds.get("hub_betweenness_weight", 0.5),
    }


def _build_config(
    company_domains: str,
    cc_key_person_threshold: float,
    min_edge_weight: int,
    hub_degree_weight: float,
    hub_betweenness_weight: float,
) -> dict:
    """フォーム値から config dict を構築."""
    domains = [d.strip() for d in company_domains.split(",") if d.strip()]
    return {
        "company_domains": domains,
        "thresholds": {
            "cc_key_person_threshold": cc_key_person_threshold,
            "min_edge_weight": min_edge_weight,
            "hub_degree_weight": hub_degree_weight,
            "hub_betweenness_weight": hub_betweenness_weight,
        },
    }


def _run_analysis(df, config, request):
    """DataFrame → 分析 → network.html レスポンス."""
    G = build_graph(df, config)
    analysis = analyze_graph(G, len(df), config)
    graph_data = generate_vis_data(G, analysis, config)

    log.info(
        "分析完了: ノード=%d, エッジ=%d, コミュニティ=%d",
        graph_data["analysis"]["total_nodes"],
        graph_data["analysis"]["total_edges"],
        len(graph_data["communities"]),
    )

    return templates.TemplateResponse("network.html", {
        "request": request,
        "graph_data": json.dumps(graph_data, ensure_ascii=False),
    })


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def upload_page(request: Request, error: str = ""):
    """アップロード画面を表示（同期: Outlook COM でフォルダ取得）."""
    _log("[GET /] Loading page...")

    # Outlook フォルダ一覧を取得（失敗してもページは表示する）
    folders = []
    outlook_error = ""
    try:
        from app.extract import get_outlook_folders
        folders = get_outlook_folders()
        _log(f"[GET /] Outlook folders: {len(folders)}")
    except ImportError:
        outlook_error = "pywin32 がインストールされていません"
        _log(f"[GET /] {outlook_error}")
    except Exception as e:
        outlook_error = str(e)
        _log(f"[GET /] Outlook error: {e}")

    response = templates.TemplateResponse("upload.html", {
        "request": request,
        "error": error,
        "folders": folders,
        "outlook_error": outlook_error,
        **_default_template_vars(),
    })
    response.headers["Cache-Control"] = "no-store"
    return response


@app.get("/api/folders")
def get_folders():
    """Outlook フォルダ一覧を返す（COM 操作のため同期エンドポイント）."""
    try:
        from app.extract import get_outlook_folders
        folders = get_outlook_folders()
        return JSONResponse(content={"folders": folders})
    except ImportError:
        return JSONResponse(
            content={"error": "Outlook 連携に必要なパッケージ (pywin32) がインストールされていません"},
            status_code=501,
        )
    except Exception as e:
        log.exception("Outlook フォルダ取得エラー")
        return JSONResponse(content={"error": str(e)}, status_code=500)


@app.post("/extract-and-analyze", response_class=HTMLResponse)
def extract_and_analyze(
    request: Request,
    folder_paths: list[str] = Form(...),
    start_date: str = Form(...),
    end_date: str = Form(...),
    company_domains: str = Form(""),
    cc_key_person_threshold: float = Form(0.30),
    min_edge_weight: int = Form(1),
    hub_degree_weight: float = Form(0.5),
    hub_betweenness_weight: float = Form(0.5),
):
    """Outlook からメール抽出 → 分析 → 可視化."""
    _log("=== extract-and-analyze START ===")
    _log(f"folder_paths={folder_paths}, start={start_date}, end={end_date}")
    try:
        from app.extract import run_extraction
        _log("app.extract imported OK")

        config = _build_config(
            company_domains, cc_key_person_threshold,
            min_edge_weight, hub_degree_weight, hub_betweenness_weight,
        )

        # Outlook から抽出（extract.py の config も必要）
        from extract import load_config as extract_load_config
        extract_config = extract_load_config()
        _log("extract config loaded OK")

        df = run_extraction(folder_paths, start_date, end_date, extract_config)
        _log(f"extraction done: {len(df)} rows")
        return _run_analysis(df, config, request)

    except Exception as e:
        import traceback
        _log("=== extract-and-analyze ERROR ===")
        _log(traceback.format_exc())
        return RedirectResponse(
            url=f"/?error=エラー: {e}",
            status_code=303,
        )


@app.post("/analyze", response_class=HTMLResponse)
async def analyze(
    request: Request,
    file: UploadFile = File(...),
    company_domains: str = Form(""),
    cc_key_person_threshold: float = Form(0.30),
    min_edge_weight: int = Form(1),
    hub_degree_weight: float = Form(0.5),
    hub_betweenness_weight: float = Form(0.5),
):
    """CSV アップロード → 分析 → 可視化."""
    if not file.filename or not file.filename.lower().endswith(".csv"):
        return RedirectResponse(url="/?error=CSVファイルを選択してください", status_code=303)

    try:
        df = load_csv(file.file)

        if df.empty:
            return RedirectResponse(url="/?error=CSVファイルが空です", status_code=303)

        config = _build_config(
            company_domains, cc_key_person_threshold,
            min_edge_weight, hub_degree_weight, hub_betweenness_weight,
        )

        return _run_analysis(df, config, request)

    except Exception as e:
        log.exception("分析中にエラーが発生しました")
        return RedirectResponse(
            url=f"/?error=分析中にエラーが発生しました: {e}",
            status_code=303,
        )
