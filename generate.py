"""CSV → ネットワーク分析 → HTML可視化 生成.

Usage:
    python generate.py --input output/emails_20250101.csv
"""

import argparse
import json
import logging
import sys
from pathlib import Path

from jinja2 import Environment, FileSystemLoader

from app.core import (
    build_graph,
    analyze_graph,
    generate_vis_data,
    load_config,
    load_csv,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# HTML generation
# ---------------------------------------------------------------------------

def render_html(graph_data: dict, output_dir: str = "output"):
    """Jinja2テンプレートにデータを注入してHTML生成."""
    template_dir = Path(__file__).parent / "templates"
    if not (template_dir / "network.html").exists():
        log.error("templates/network.html が見つかりません。")
        sys.exit(1)

    env = Environment(
        loader=FileSystemLoader(str(template_dir)),
        autoescape=False,
    )
    template = env.get_template("network.html")

    html = template.render(graph_data=json.dumps(graph_data, ensure_ascii=False))

    out = Path(__file__).parent / output_dir
    out.mkdir(exist_ok=True)
    output_path = out / "index.html"
    output_path.write_text(html, encoding="utf-8")

    log.info("HTML生成完了: %s", output_path)
    return str(output_path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="メールネットワーク可視化 HTML 生成"
    )
    parser.add_argument(
        "--input", required=True, help="入力CSVファイルパス"
    )
    parser.add_argument(
        "--output", default="output", help="出力ディレクトリ (default: output)"
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        log.error("入力ファイルが見つかりません: %s", args.input)
        sys.exit(1)

    config = load_config()
    df = load_csv(str(input_path))

    if df.empty:
        log.warning("CSVが空です。")
        sys.exit(0)

    G = build_graph(df, config)
    analysis = analyze_graph(G, len(df), config)
    graph_data = generate_vis_data(G, analysis, config)
    html_path = render_html(graph_data, args.output)

    print(f"\n完了！ ネットワーク可視化を生成しました。")
    print(f"HTML: {html_path}")
    print(f"ノード: {graph_data['analysis']['total_nodes']}, "
          f"エッジ: {graph_data['analysis']['total_edges']}, "
          f"コミュニティ: {len(graph_data['communities'])}")


if __name__ == "__main__":
    main()
