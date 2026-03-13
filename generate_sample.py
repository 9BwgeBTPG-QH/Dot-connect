"""サンプルデータからスタンドアロン HTML を生成し docs/index.html に配置.

Usage:
    python generate_sample.py
    -> docs/index.html が出力される (GitHub Pages 用)
"""

import json
import os
import re

import pandas as pd

from app.core import build_graph, analyze_graph, generate_vis_data


def load_sample_csv() -> pd.DataFrame:
    """sample/emails_sample.csv を読み込む."""
    path = os.path.join(os.path.dirname(__file__), "sample", "emails_sample.csv")
    return pd.read_csv(path, encoding="utf-8-sig", quoting=1)


def read_template_css() -> str:
    """templates/network.html から <style> タグの中身を抽出."""
    path = os.path.join(os.path.dirname(__file__), "templates", "network.html")
    with open(path, encoding="utf-8") as f:
        content = f.read()
    start = content.find("<style>") + len("<style>")
    end = content.find("</style>")
    return content[start:end]


def read_template_script() -> str:
    """templates/network.html から <script> タグの中身を抽出.

    Jinja2 テンプレート変数 ({{ graph_data | safe }}) の行と
    エクスポート関連の関数は除外する。
    """
    path = os.path.join(os.path.dirname(__file__), "templates", "network.html")
    with open(path, encoding="utf-8") as f:
        content = f.read()
    start = content.find("<script>") + len("<script>")
    end = content.rfind("</script>")
    script = content[start:end]

    # Jinja2 テンプレート行を除去
    lines = script.split("\n")
    filtered = []
    skip_export = False
    for line in lines:
        # Jinja2 injection line
        if "{{ graph_data" in line or "graph_data |" in line:
            continue
        # Export functions - skip until next section header
        if "// Export:" in line or "function exportPNG" in line or "function exportHTML" in line or "function exportCSV" in line:
            skip_export = True
            continue
        if skip_export and line.startswith("// =="):
            skip_export = False
        if skip_export:
            continue
        # Export button event listeners
        if "btn-export-" in line:
            continue
        filtered.append(line)

    return "\n".join(filtered)


def build_standalone_html(vis_data: dict) -> str:
    """vis_data からスタンドアロン HTML を構築."""
    css = read_template_css()
    script = read_template_script()
    data_json = json.dumps(vis_data, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dot-connect Demo — Email Network Visualization</title>
<script src="https://unpkg.com/vis-network@9.1.6/standalone/umd/vis-network.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/wordcloud2.js/1.1.2/wordcloud2.min.js"></script>
<style>
{css}
</style>
</head>
<body>

<!-- Toolbar -->
<div id="toolbar">
  <div class="left">
    <div style="font-size:13px;color:#8899aa;margin-right:8px;">
      <span style="color:#6366f1;font-weight:600;">Dot-connect</span> / Sample Data (30 emails)
    </div>
    <div id="breadcrumb">
      <span class="current">All</span>
    </div>
    <button class="view-btn active" id="btn-network">Network</button>
    <button class="view-btn" id="btn-wordcloud">Word Cloud</button>
  </div>
  <div class="right">
    <span id="stat-nodes"></span>
    <span id="stat-edges"></span>
    <span id="stat-communities"></span>
  </div>
</div>

<!-- Main -->
<div id="main">
  <div id="network-container"></div>
  <div id="wordcloud-container">
    <canvas id="wordcloud-canvas"></canvas>
  </div>
  <div id="side-panel"></div>
</div>

<!-- Legend -->
<div id="legend"></div>

<script>
var DATA = {data_json};
{script}
</script>
</body>
</html>"""
    return html


def main():
    print("サンプルCSVを読み込み中...")
    df = load_sample_csv()
    print(f"  {len(df)} レコード")

    print("分析パイプラインを実行中...")
    config = {
        "company_domains": ["example.co.jp"],
        "thresholds": {
            "cc_key_person_threshold": 0.30,
            "min_edge_weight": 1,
            "hub_degree_weight": 0.5,
            "hub_betweenness_weight": 0.5,
        },
    }
    G = build_graph(df, config)
    analysis = analyze_graph(G, len(df), config)
    vis_data = generate_vis_data(G, analysis, config)

    n = vis_data["analysis"]["total_nodes"]
    e = vis_data["analysis"]["total_edges"]
    c = len(vis_data["communities"])
    print(f"  ノード {n} / エッジ {e} / コミュニティ {c}")

    print("スタンドアロン HTML を生成中...")
    html = build_standalone_html(vis_data)

    out_dir = os.path.join(os.path.dirname(__file__), "docs")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "index.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"完了！ {out_path}")
    print(f"  ファイルサイズ: {os.path.getsize(out_path) // 1024} KB")


if __name__ == "__main__":
    main()
