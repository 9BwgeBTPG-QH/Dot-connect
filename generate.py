"""CSV → ネットワーク分析 → HTML可視化 生成.

Usage:
    python generate.py --input output/emails_20250101.csv
"""

import argparse
import json
import logging
import re
import sys
from collections import defaultdict
from pathlib import Path

import networkx as nx
import pandas as pd
import yaml
from jinja2 import Environment, FileSystemLoader

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# コミュニティカラーパレット
COMMUNITY_COLORS = [
    "#6366f1",  # indigo
    "#f59e0b",  # amber
    "#10b981",  # emerald
    "#ef4444",  # red
    "#8b5cf6",  # violet
    "#06b6d4",  # cyan
    "#f97316",  # orange
    "#ec4899",  # pink
    "#14b8a6",  # teal
    "#a855f7",  # purple
    "#84cc16",  # lime
    "#e11d48",  # rose
]


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def load_config(path: str = "config.yaml") -> dict:
    config_path = Path(__file__).parent / path
    if not config_path.exists():
        return {
            "company_domains": [],
            "thresholds": {
                "cc_key_person_threshold": 0.30,
                "min_edge_weight": 1,
                "hub_degree_weight": 0.5,
                "hub_betweenness_weight": 0.5,
            },
        }
    with open(config_path, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


# ---------------------------------------------------------------------------
# CSV parsing
# ---------------------------------------------------------------------------

def parse_address_field(field: str) -> list[tuple[str, str]]:
    """'Name <email>; Name2 <email2>' 形式のフィールドをパース.

    Returns:
        list of (email, name)
    """
    if not field or pd.isna(field):
        return []
    results = []
    for entry in field.split("; "):
        entry = entry.strip()
        if not entry:
            continue
        match = re.match(r"^(.*?)\s*<(.+?)>$", entry)
        if match:
            name, email = match.group(1).strip(), match.group(2).strip().lower()
            results.append((email, name))
        elif "@" in entry:
            results.append((entry.strip().lower(), ""))
    return results


def load_csv(filepath: str) -> pd.DataFrame:
    """CSV読み込み."""
    df = pd.read_csv(filepath, encoding="utf-8-sig", quoting=1)
    log.info("CSV読込: %d 件", len(df))
    return df


# ---------------------------------------------------------------------------
# Graph construction
# ---------------------------------------------------------------------------

def build_graph(df: pd.DataFrame, config: dict) -> nx.DiGraph:
    """メールデータから有向グラフを構築."""
    G = nx.DiGraph()
    company_domains = [d.lower() for d in config.get("company_domains", [])]

    # ノード情報を集計
    node_stats = defaultdict(lambda: {
        "name": "", "email": "", "domain": "",
        "sent": 0, "received": 0, "cc_count": 0,
    })

    for _, row in df.iterrows():
        from_email = row["from_email"].strip().lower() if pd.notna(row.get("from_email")) else ""
        from_name = row.get("from_name", "") if pd.notna(row.get("from_name")) else ""

        if not from_email:
            continue

        # 送信者のノード情報更新
        domain = from_email.split("@")[-1] if "@" in from_email else ""
        node_stats[from_email]["name"] = from_name or node_stats[from_email]["name"]
        node_stats[from_email]["email"] = from_email
        node_stats[from_email]["domain"] = domain
        node_stats[from_email]["sent"] += 1

        # To 受信者
        to_addrs = parse_address_field(row.get("to", ""))
        for to_email, to_name in to_addrs:
            to_domain = to_email.split("@")[-1] if "@" in to_email else ""
            node_stats[to_email]["name"] = to_name or node_stats[to_email]["name"]
            node_stats[to_email]["email"] = to_email
            node_stats[to_email]["domain"] = to_domain
            node_stats[to_email]["received"] += 1

            # To エッジ
            if G.has_edge(from_email, to_email):
                G[from_email][to_email]["to_weight"] += 1
            else:
                G.add_edge(from_email, to_email, to_weight=1, cc_weight=0)

        # CC 受信者
        cc_addrs = parse_address_field(row.get("cc", ""))
        for cc_email, cc_name in cc_addrs:
            cc_domain = cc_email.split("@")[-1] if "@" in cc_email else ""
            node_stats[cc_email]["name"] = cc_name or node_stats[cc_email]["name"]
            node_stats[cc_email]["email"] = cc_email
            node_stats[cc_email]["domain"] = cc_domain
            node_stats[cc_email]["cc_count"] += 1

            # CC エッジ
            if G.has_edge(from_email, cc_email):
                G[from_email][cc_email]["cc_weight"] += 1
            else:
                G.add_edge(from_email, cc_email, to_weight=0, cc_weight=1)

    # ノード属性を設定
    for email, stats in node_stats.items():
        if email in G.nodes:
            is_internal = any(stats["domain"].endswith(d) for d in company_domains)
            G.nodes[email].update({
                "name": stats["name"] or email.split("@")[0],
                "email": email,
                "domain": stats["domain"],
                "is_internal": is_internal,
                "sent": stats["sent"],
                "received": stats["received"],
                "cc_count": stats["cc_count"],
            })

    log.info("グラフ構築: %d ノード, %d エッジ", G.number_of_nodes(), G.number_of_edges())
    return G


# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------

def analyze_graph(G: nx.DiGraph, config: dict) -> dict:
    """ネットワーク分析を実行."""
    thresholds = config.get("thresholds", {})
    total_mails = sum(
        data.get("to_weight", 0) for _, _, data in G.edges(data=True)
    )
    if total_mails == 0:
        total_mails = 1

    # --- CC キーマン ---
    cc_threshold = thresholds.get("cc_key_person_threshold", 0.30)
    cc_key_persons = []
    for node, data in G.nodes(data=True):
        cc_count = data.get("cc_count", 0)
        if cc_count / total_mails >= cc_threshold:
            cc_key_persons.append({
                "email": node,
                "name": data.get("name", node),
                "cc_count": cc_count,
                "ratio": round(cc_count / total_mails, 3),
            })
    cc_key_persons.sort(key=lambda x: x["cc_count"], reverse=True)

    # --- ハブ (centrality) ---
    hub_dw = thresholds.get("hub_degree_weight", 0.5)
    hub_bw = thresholds.get("hub_betweenness_weight", 0.5)

    undirected = G.to_undirected()
    degree_c = nx.degree_centrality(undirected)
    betweenness_c = nx.betweenness_centrality(undirected)

    hub_scores = {}
    for node in G.nodes:
        hub_scores[node] = (
            hub_dw * degree_c.get(node, 0)
            + hub_bw * betweenness_c.get(node, 0)
        )

    top_hubs = sorted(hub_scores.items(), key=lambda x: x[1], reverse=True)[:20]
    hubs = []
    for email, score in top_hubs:
        data = G.nodes[email]
        hubs.append({
            "email": email,
            "name": data.get("name", email),
            "score": round(score, 4),
            "degree_centrality": round(degree_c.get(email, 0), 4),
            "betweenness_centrality": round(betweenness_c.get(email, 0), 4),
        })

    # --- Louvain コミュニティ ---
    communities = list(nx.community.louvain_communities(undirected, seed=42))
    community_map = {}
    for idx, comm in enumerate(communities):
        for node in comm:
            community_map[node] = idx

    # ノードにコミュニティIDを付与
    for node in G.nodes:
        G.nodes[node]["community"] = community_map.get(node, 0)

    community_info = []
    for idx, comm in enumerate(communities):
        members = [{"email": n, "name": G.nodes[n].get("name", n)} for n in comm if n in G.nodes]
        community_info.append({
            "id": idx,
            "color": COMMUNITY_COLORS[idx % len(COMMUNITY_COLORS)],
            "size": len(comm),
            "members": members,
        })

    log.info(
        "分析完了: CCキーマン %d人, コミュニティ %d個",
        len(cc_key_persons), len(communities),
    )

    return {
        "total_mails": total_mails,
        "cc_key_persons": cc_key_persons,
        "hubs": hubs,
        "communities": community_info,
        "community_map": community_map,
    }


# ---------------------------------------------------------------------------
# vis.js data generation
# ---------------------------------------------------------------------------

def generate_vis_data(G: nx.DiGraph, analysis: dict, config: dict) -> dict:
    """vis.js 用のJSONデータを生成."""
    thresholds = config.get("thresholds", {})
    min_weight = thresholds.get("min_edge_weight", 1)
    community_map = analysis["community_map"]
    communities = analysis["communities"]

    # ノードデータ
    nodes = []
    cc_key_emails = {p["email"] for p in analysis["cc_key_persons"]}
    hub_emails = {h["email"] for h in analysis["hubs"][:10]}

    for node, data in G.nodes(data=True):
        comm_id = community_map.get(node, 0)
        color = COMMUNITY_COLORS[comm_id % len(COMMUNITY_COLORS)]

        total_activity = data.get("sent", 0) + data.get("received", 0) + data.get("cc_count", 0)
        size = max(8, min(40, 8 + total_activity * 0.5))

        label = data.get("name", node)
        if len(label) > 15:
            label = label[:14] + "…"

        node_entry = {
            "id": node,
            "label": label,
            "name": data.get("name", node),
            "email": node,
            "domain": data.get("domain", ""),
            "is_internal": data.get("is_internal", False),
            "sent": data.get("sent", 0),
            "received": data.get("received", 0),
            "cc_count": data.get("cc_count", 0),
            "community": comm_id,
            "color": color,
            "size": size,
            "is_cc_key": node in cc_key_emails,
            "is_hub": node in hub_emails,
        }
        nodes.append(node_entry)

    # エッジデータ
    edges = []
    for u, v, data in G.edges(data=True):
        to_w = data.get("to_weight", 0)
        cc_w = data.get("cc_weight", 0)
        total = to_w + cc_w
        if total < min_weight:
            continue
        edges.append({
            "from": u,
            "to": v,
            "to_weight": to_w,
            "cc_weight": cc_w,
            "weight": total,
            "width": max(1, min(8, total * 0.3)),
        })

    # ワードクラウドデータ
    wordcloud_data = []
    for node, data in G.nodes(data=True):
        total = data.get("sent", 0) + data.get("received", 0) + data.get("cc_count", 0)
        if total > 0:
            comm_id = community_map.get(node, 0)
            wordcloud_data.append({
                "text": data.get("name", node),
                "size": total,
                "email": node,
                "color": COMMUNITY_COLORS[comm_id % len(COMMUNITY_COLORS)],
            })
    wordcloud_data.sort(key=lambda x: x["size"], reverse=True)

    return {
        "nodes": nodes,
        "edges": edges,
        "communities": communities,
        "analysis": {
            "total_mails": analysis["total_mails"],
            "total_nodes": len(nodes),
            "total_edges": len(edges),
            "cc_key_persons": analysis["cc_key_persons"],
            "hubs": analysis["hubs"],
        },
        "wordcloud_data": wordcloud_data,
    }


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
    analysis = analyze_graph(G, config)
    graph_data = generate_vis_data(G, analysis, config)
    html_path = render_html(graph_data, args.output)

    print(f"\n完了！ ネットワーク可視化を生成しました。")
    print(f"HTML: {html_path}")
    print(f"ノード: {graph_data['analysis']['total_nodes']}, "
          f"エッジ: {graph_data['analysis']['total_edges']}, "
          f"コミュニティ: {len(graph_data['communities'])}")


if __name__ == "__main__":
    main()
