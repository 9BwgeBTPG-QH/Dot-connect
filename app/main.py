"""FastAPI アプリケーション: メール抽出 & CSV アップロード → ネットワーク可視化."""

import json
import logging
import shutil
import sys
import uuid
from collections import OrderedDict
from pathlib import Path
from typing import Optional

import yaml

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
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

# --- Auto-create config.yaml from template if missing ---
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_CONFIG_PATH = _PROJECT_ROOT / "config.yaml"
_CONFIG_EXAMPLE = _PROJECT_ROOT / "config.yaml.example"
if not _CONFIG_PATH.exists() and _CONFIG_EXAMPLE.exists():
    shutil.copy2(_CONFIG_EXAMPLE, _CONFIG_PATH)
    _log("[Config] config.yaml を config.yaml.example からコピーしました")

app = FastAPI(title="Dot-connect", description="メールネットワーク可視化")

# --- In-memory result store (UUID → rendered HTML, max 50 entries) ---
MAX_RESULTS = 50
_results_store: OrderedDict[str, str] = OrderedDict()

# --- Graph API auth (lazy-initialized) ---
_graph_auth_instance: Optional["GraphAuth"] = None  # noqa: F821
_graph_auth_checked = False


def _get_graph_auth():
    """config.yaml の graph_api セクションから GraphAuth を遅延初期化."""
    global _graph_auth_instance, _graph_auth_checked
    if _graph_auth_checked:
        return _graph_auth_instance

    _graph_auth_checked = True
    cfg = load_config()
    graph_cfg = cfg.get("graph_api", {})
    client_id = (graph_cfg.get("client_id") or "").strip()
    tenant_id = (graph_cfg.get("tenant_id") or "").strip()

    if not client_id or not tenant_id:
        _graph_auth_instance = None
        return None

    try:
        from app.graph_auth import GraphAuth
        redirect_uri = (graph_cfg.get("redirect_uri") or "http://localhost:8000/auth/callback").strip()
        _graph_auth_instance = GraphAuth(
            client_id=client_id,
            tenant_id=tenant_id,
            redirect_uri=redirect_uri,
        )
        _log(f"[Graph API] Initialized: client_id={client_id[:8]}...")
    except Exception as e:
        _log(f"[Graph API] Init failed: {e}")
        _graph_auth_instance = None

    return _graph_auth_instance


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

# --- Outlook COM 結果キャッシュ（起動後1回だけ試行） ---
_outlook_cache: dict = {"checked": False, "folders": [], "error": ""}


def _get_outlook_folders_cached() -> tuple[list, str]:
    """Outlook COM を1回だけ試行し、結果をキャッシュ."""
    if _outlook_cache["checked"]:
        return _outlook_cache["folders"], _outlook_cache["error"]

    folders = []
    error = ""
    try:
        from app.extract import get_outlook_folders
        folders = get_outlook_folders()
        _log(f"[Outlook] Connected: {len(folders)} folders")
    except ImportError:
        error = "pywin32 がインストールされていません"
        _log(f"[Outlook] {error}")
    except Exception as e:
        error = str(e)
        _log(f"[Outlook] Error: {e}")

    _outlook_cache["checked"] = True
    _outlook_cache["folders"] = folders
    _outlook_cache["error"] = error
    return folders, error


@app.get("/", response_class=HTMLResponse)
def upload_page(request: Request, error: str = ""):
    """アップロード画面を表示."""
    _log("[GET /] Loading page...")

    folders, outlook_error = _get_outlook_folders_cached()

    # Graph API availability
    graph_auth = _get_graph_auth()
    graph_available = graph_auth is not None
    graph_authenticated = graph_auth.is_authenticated() if graph_available else False

    response = templates.TemplateResponse("upload.html", {
        "request": request,
        "error": error,
        "folders": folders,
        "outlook_error": outlook_error,
        "graph_available": graph_available,
        "graph_authenticated": graph_authenticated,
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


# ---------------------------------------------------------------------------
# Graph API Routes
# ---------------------------------------------------------------------------

@app.get("/auth/login")
def auth_login(request: Request):
    """Microsoft ログインページへリダイレクト."""
    graph_auth = _get_graph_auth()
    if graph_auth is None:
        return RedirectResponse(url="/?error=Graph API が設定されていません", status_code=303)

    state = str(uuid.uuid4())
    auth_url = graph_auth.get_auth_url(state)
    _log(f"[Graph API] Login redirect: state={state[:8]}...")
    return RedirectResponse(url=auth_url)


@app.get("/auth/callback")
def auth_callback(request: Request):
    """OAuth2 コールバック処理."""
    graph_auth = _get_graph_auth()
    if graph_auth is None:
        return RedirectResponse(url="/?error=Graph API が設定されていません", status_code=303)

    # クエリパラメータを dict に変換
    auth_response = dict(request.query_params)
    result = graph_auth.acquire_token_by_auth_code(auth_response)

    if "error" in result:
        error_desc = result.get("error_description", result.get("error", "認証に失敗しました"))
        _log(f"[Graph API] Auth callback error: {error_desc}")
        return RedirectResponse(url=f"/?error=認証エラー: {error_desc}", status_code=303)

    _log("[Graph API] Authentication successful")
    return RedirectResponse(url="/")


@app.get("/auth/logout")
def auth_logout():
    """トークンクリア."""
    graph_auth = _get_graph_auth()
    if graph_auth is not None:
        graph_auth.sign_out()
        _log("[Graph API] Signed out")
    return RedirectResponse(url="/")


@app.get("/api/graph-folders")
def get_graph_folders_api():
    """Graph API でフォルダ一覧取得."""
    graph_auth = _get_graph_auth()
    if graph_auth is None:
        return JSONResponse(content={"error": "Graph API が設定されていません"}, status_code=501)

    access_token = graph_auth.get_access_token()
    if access_token is None:
        return JSONResponse(content={"error": "認証が必要です"}, status_code=401)

    try:
        from app.graph_extract import get_graph_folders
        folders = get_graph_folders(access_token)
        return JSONResponse(content={"folders": folders})
    except Exception as e:
        log.exception("Graph API フォルダ取得エラー")
        return JSONResponse(content={"error": str(e)}, status_code=500)


@app.post("/graph-extract-and-analyze", response_class=HTMLResponse)
def graph_extract_and_analyze(
    request: Request,
    folder_ids: list[str] = Form(...),
    start_date: str = Form(...),
    end_date: str = Form(...),
    company_domains: str = Form(""),
    cc_key_person_threshold: float = Form(0.30),
    min_edge_weight: int = Form(1),
    hub_degree_weight: float = Form(0.5),
    hub_betweenness_weight: float = Form(0.5),
):
    """Graph API 抽出 → 分析 → 可視化."""
    _log("=== graph-extract-and-analyze START ===")
    _log(f"folder_ids={folder_ids}, start={start_date}, end={end_date}")
    try:
        graph_auth = _get_graph_auth()
        if graph_auth is None:
            return RedirectResponse(url="/?error=Graph API が設定されていません", status_code=303)

        access_token = graph_auth.get_access_token()
        if access_token is None:
            return RedirectResponse(url="/?error=再認証が必要です。サインインしてください。", status_code=303)

        config = _build_config(
            company_domains, cc_key_person_threshold,
            min_edge_weight, hub_degree_weight, hub_betweenness_weight,
        )

        # extract.py の config も必要（exclude_addresses, alias_map 等）
        from extract import load_config as extract_load_config
        extract_config = extract_load_config()

        from app.graph_extract import run_graph_extraction
        df = run_graph_extraction(access_token, folder_ids, start_date, end_date, extract_config)
        _log(f"Graph extraction done: {len(df)} rows")

        if df.empty:
            return RedirectResponse(url="/?error=抽出されたメールがありません", status_code=303)

        return _run_analysis(df, config, request)

    except Exception as e:
        import traceback
        _log("=== graph-extract-and-analyze ERROR ===")
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


# ---------------------------------------------------------------------------
# API: ローカル抽出ツール連携
# ---------------------------------------------------------------------------

@app.post("/api/upload-csv")
async def api_upload_csv(
    request: Request,
    file: UploadFile = File(...),
    company_domains: str = Form(""),
    cc_key_person_threshold: float = Form(0.30),
    min_edge_weight: int = Form(1),
    hub_degree_weight: float = Form(0.5),
    hub_betweenness_weight: float = Form(0.5),
):
    """ローカル抽出ツールからの CSV アップロード → 分析 → 結果URL返却."""
    try:
        df = load_csv(file.file)
        if df.empty:
            return JSONResponse({"error": "CSVファイルが空です"}, status_code=400)

        config = _build_config(
            company_domains, cc_key_person_threshold,
            min_edge_weight, hub_degree_weight, hub_betweenness_weight,
        )

        G = build_graph(df, config)
        analysis = analyze_graph(G, len(df), config)
        graph_data = generate_vis_data(G, analysis, config)

        log.info(
            "API分析完了: ノード=%d, エッジ=%d",
            graph_data["analysis"]["total_nodes"],
            graph_data["analysis"]["total_edges"],
        )

        # Render HTML to string and store
        graph_json = json.dumps(graph_data, ensure_ascii=False)
        html = templates.get_template("network.html").render(
            graph_data=graph_json,
        )

        result_id = str(uuid.uuid4())
        _results_store[result_id] = html
        # Evict oldest if over limit
        while len(_results_store) > MAX_RESULTS:
            _results_store.popitem(last=False)

        return JSONResponse({"result_url": f"/results/{result_id}"})

    except Exception as e:
        log.exception("API upload-csv エラー")
        return JSONResponse({"error": str(e)}, status_code=500)


@app.get("/results/{result_id}", response_class=HTMLResponse)
def get_result(result_id: str):
    """保存済み分析結果を表示."""
    html = _results_store.get(result_id)
    if html is None:
        return HTMLResponse(
            content="<h1>結果が見つかりません</h1><p>結果の有効期限が切れたか、URLが無効です。</p>",
            status_code=404,
        )
    return HTMLResponse(content=html)


@app.get("/api/extract-script")
def serve_extract_script():
    """extract_and_upload.py をダウンロード提供."""
    script_path = Path(__file__).resolve().parent.parent / "extract_and_upload.py"
    return FileResponse(
        script_path,
        media_type="text/plain; charset=utf-8",
        filename="extract_and_upload.py",
    )


@app.get("/download-extractor")
def download_extractor(
    request: Request,
    start_date: str = "",
    end_date: str = "",
    company_domains: str = "",
    cc_key_person_threshold: float = 0.30,
    min_edge_weight: int = 1,
    hub_degree_weight: float = 0.5,
    hub_betweenness_weight: float = 0.5,
):
    """パラメータ埋め込み済みの .bat ファイルを生成してダウンロード."""
    host = request.headers.get("host", "localhost:8000")
    server_url = f"http://{host}"

    # config.yaml から共有パスを取得
    cfg = load_config()
    share_path = cfg.get("network_share_path", "")

    bat_content = f'''@echo off
setlocal enabledelayedexpansion

echo ========================================
echo   Dot-connect - Email Extractor
echo   Period: {start_date} ~ {end_date}
echo   Server: {server_url}
echo ========================================
echo.

set "SERVER_URL={server_url}"
set "TEMP_SCRIPT=%TEMP%\\extract_and_upload.py"
set "LOCAL_PYTHON_DIR=%TEMP%\\dotconnect_python"
set "PYTHON_VER=3.12.8"
set "PYTHON_ZIP=python-%PYTHON_VER%-embed-amd64.zip"
set "PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VER%/%PYTHON_ZIP%"

REM --- Find Python ---
set "PYTHON_CMD="

REM 1) Check network share embedded Python
if exist "{share_path}\\python\\python.exe" (
    set "PYTHON_CMD={share_path}\\python\\python.exe"
    echo [OK] Python found: network share
    goto :check_pywin32
)

REM 2) Check previously downloaded local Python
if exist "%LOCAL_PYTHON_DIR%\\python.exe" (
    set "PYTHON_CMD=%LOCAL_PYTHON_DIR%\\python.exe"
    echo [OK] Python found: local cache
    goto :check_pywin32
)

REM 3) Check python in PATH (verify it's real, not Windows Store stub)
where python >nul 2>&1
if %errorlevel%==0 (
    python -c "import sys" >nul 2>&1
    if !errorlevel!==0 (
        set "PYTHON_CMD=python"
        echo [OK] Python found: PATH
        goto :check_pywin32
    )
)

REM 4) Try py launcher
where py >nul 2>&1
if %errorlevel%==0 (
    py -c "import sys" >nul 2>&1
    if !errorlevel!==0 (
        set "PYTHON_CMD=py"
        echo [OK] Python found: py launcher
        goto :check_pywin32
    )
)

REM 5) Auto-download embedded Python
echo [SETUP] Python not found. Downloading portable Python %PYTHON_VER%...
echo   (This is a one-time setup, please wait...)
echo.

if not exist "%LOCAL_PYTHON_DIR%" mkdir "%LOCAL_PYTHON_DIR%"

powershell -Command "Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%TEMP%\\%PYTHON_ZIP%'" 2>nul
if not exist "%TEMP%\\%PYTHON_ZIP%" (
    echo [ERROR] Python download failed. Check internet connection.
    pause
    exit /b 1
)

echo   Extracting...
powershell -Command "Expand-Archive -Path '%TEMP%\\%PYTHON_ZIP%' -DestinationPath '%LOCAL_PYTHON_DIR%' -Force"
del "%TEMP%\\%PYTHON_ZIP%" >nul 2>&1

REM Enable pip in embedded Python
echo import site>> "%LOCAL_PYTHON_DIR%\\python312._pth"
echo   Installing pip...
powershell -Command "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%LOCAL_PYTHON_DIR%\\get-pip.py'" 2>nul
"%LOCAL_PYTHON_DIR%\\python.exe" "%LOCAL_PYTHON_DIR%\\get-pip.py" --no-warn-script-location -q
del "%LOCAL_PYTHON_DIR%\\get-pip.py" >nul 2>&1

set "PYTHON_CMD=%LOCAL_PYTHON_DIR%\\python.exe"
echo [OK] Python %PYTHON_VER% installed to %LOCAL_PYTHON_DIR%

:check_pywin32
REM --- Check pywin32 dependency ---
"%PYTHON_CMD%" -c "import win32com.client" >nul 2>&1
if %errorlevel% neq 0 (
    echo [SETUP] Installing pywin32...
    "%PYTHON_CMD%" -m pip install pywin32 --no-warn-script-location -q
    "%PYTHON_CMD%" -c "import win32com.client" >nul 2>&1
    if %errorlevel% neq 0 (
        echo [ERROR] pywin32 installation failed.
        pause
        exit /b 1
    )
    echo [OK] pywin32 installed.
)

:download_script
echo [1/3] Downloading script...
powershell -Command "Invoke-WebRequest -Uri '%SERVER_URL%/api/extract-script' -OutFile '%TEMP_SCRIPT%'" 2>nul
if not exist "%TEMP_SCRIPT%" (
    echo [ERROR] Failed to download script.
    echo   Check connection to %SERVER_URL%
    pause
    exit /b 1
)

echo [2/3] Connecting to Outlook...
echo   Select mail folders when prompted.
echo.

"%PYTHON_CMD%" "%TEMP_SCRIPT%" ^
    --server_url "%SERVER_URL%" ^
    --start_date "{start_date}" ^
    --end_date "{end_date}" ^
    --company_domains "{company_domains}" ^
    --cc_key_person_threshold {cc_key_person_threshold} ^
    --min_edge_weight {min_edge_weight} ^
    --hub_degree_weight {hub_degree_weight} ^
    --hub_betweenness_weight {hub_betweenness_weight}

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Email extraction failed.
    echo   Check that Outlook is running and accessible.
    pause
    exit /b 1
)

echo.
echo [3/3] Done.
del "%TEMP_SCRIPT%" >nul 2>&1
pause
'''

    return Response(
        content=bat_content.encode("ascii", errors="replace"),
        media_type="application/octet-stream",
        headers={"Content-Disposition": "attachment; filename=extract_emails.bat"},
    )


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

@app.get("/settings", response_class=HTMLResponse)
def settings_page(request: Request, saved: str = ""):
    """設定画面を表示."""
    cfg = load_config()
    graph_cfg = cfg.get("graph_api", {})
    return templates.TemplateResponse("settings.html", {
        "request": request,
        "saved": saved,
        "client_id": graph_cfg.get("client_id") or "",
        "tenant_id": graph_cfg.get("tenant_id") or "",
        "redirect_uri": graph_cfg.get("redirect_uri") or "http://localhost:8000/auth/callback",
    })


@app.post("/settings", response_class=HTMLResponse)
def save_settings(
    request: Request,
    client_id: str = Form(""),
    tenant_id: str = Form(""),
    redirect_uri: str = Form("http://localhost:8000/auth/callback"),
):
    """設定を config.yaml に保存."""
    global _graph_auth_instance, _graph_auth_checked

    try:
        # config.yaml を読み込み
        cfg = load_config()

        # graph_api セクションを更新
        cfg["graph_api"] = {
            "client_id": client_id.strip(),
            "tenant_id": tenant_id.strip(),
            "redirect_uri": redirect_uri.strip(),
        }

        # config.yaml に書き戻し
        with open(_CONFIG_PATH, "w", encoding="utf-8") as f:
            yaml.dump(cfg, f, default_flow_style=False, allow_unicode=True, sort_keys=False)

        _log(f"[Settings] config.yaml updated: client_id={client_id[:8]}..." if client_id else "[Settings] config.yaml updated: Graph API cleared")

        # Graph API 認証を再初期化
        _graph_auth_instance = None
        _graph_auth_checked = False

        return RedirectResponse(url="/settings?saved=1", status_code=303)

    except Exception as e:
        _log(f"[Settings] Error saving config: {e}")
        return templates.TemplateResponse("settings.html", {
            "request": request,
            "saved": "",
            "error": f"設定の保存に失敗しました: {e}",
            "client_id": client_id,
            "tenant_id": tenant_id,
            "redirect_uri": redirect_uri,
        })
