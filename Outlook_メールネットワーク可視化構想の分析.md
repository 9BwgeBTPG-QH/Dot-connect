# Outlook メールネットワーク可視化構想の分析

## 1. アイデアの本質的価値

### 核心的な問い

> [!NOTE]
>
> 「組織内の情報フローは誰が握っているのか？」を可視化する試み

### 実現できること

* **暗黙知の可視化**: 「誰が実際に意思決定に関わっているか」が見える

* **引継ぎの革命**: 組織図ではなく「実際のコミュニケーション構造」を伝承

* **権力構造の発見**: CCに必ず入る人 = キーパーソン

* **情報のハブ特定**: ネットワーク分析で中心人物を数値化

### 私の思考プロセスとの親和性

```
受容: Outlookのメールデータという既存資産
 ↓
適応: Obsidianという使い慣れたツールで可視化
 ↓
創造: 引継ぎ資料という新しい価値
```

***

## 2. 技術的実現可能性（確信度：高）

### 実装の流れ

#### Phase 1: データ抽出

```python
# Microsoft Graph API または win32com (Outlook COM)
- フォルダ指定
- メール一覧取得
- From / To / CC の抽出
- タイムスタンプ取得
```

#### Phase 2: ネットワークデータ構築

```python
# NetworkX または独自処理
nodes = {
    "user@example.com": {
        "name": "山田太郎",
        "sent_count": 45,
        "received_count": 120,
        "cc_count": 89
    }
}

edges = [
    ("sender@example.com", "receiver@example.com", {"weight": 5, "type": "to"}),
    ("sender@example.com", "cc@example.com", {"weight": 3, "type": "cc"})
]
```

#### Phase 3: Obsidian形式出力

```markdown
# 選択肢A: Markdown + Dataview plugin
---
type: person
email: hide@company.com
sent: 45
received: 120
---

# 選択肢B: Obsidian Graph View用のリンク形式
[[山田太郎]] → [[佐藤花子]] (送信: 5通)
[[山田太郎]] → [[鈴木一郎]] (CC: 3通)
```

***

## 3. 可視化の設計思想

### ノード（人）の属性

| 属性         | 意味           | 視覚表現               |
| ---------- | ------------ | ------------------ |
| **ノードサイズ** | 総メール数（送受信合計） | 大きいほど活発            |
| **ノード色**   | 役割区分         | 送信者=赤、受信者=青、CC常連=緑 |
| **ノード形状**  | 社内/社外        | 丸=社内、四角=社外         |

### エッジ（関係）の属性

| 属性        | 意味                | 視覸表現           |
| --------- | ----------------- | -------------- |
| **線の太さ**  | やり取り頻度            | 太いほど密接         |
| **線の色**   | To(実線赤) / CC(点線緑) | 意思決定者 vs 情報共有者 |
| **矢印の向き** | 情報フロー             | A→B = AがBに送信   |

### 実装時の工夫

```python
# CCに"必ず入っている人"を特定
cc_frequency = Counter(cc_addresses)
always_cc = [addr for addr, count in cc_frequency.items()
             if count > total_emails * 0.7]  # 70%以上のメールに登場

# この人たちを「情報のゲートキーパー」として強調表示
```

***

## 4. 引継ぎ資料としての実用性

### 従来の引継ぎ資料との違い

| 従来              | この可視化          |
| --------------- | -------------- |
| 「○○さんに相談してください」 | 実際に誰が誰に相談しているか |
| 「この案件は△△部が担当」   | 実際に誰がCCに入っているか |
| 組織図の形式知         | コミュニケーションの暗黙知  |

### 配布形式の選択肢

1. **Obsidian Vault丸ごと配布**

   * Graph Viewをインタラクティブに操作可能

   * 受け手もObsidian必須

2. **静的HTML出力**

   * vis.js や D3.js で可視化

   * ブラウザで開くだけ、誰でも見られる

3. **PDF + 注釈**

   * NetworkXでグラフ描画 → matplotlib → PDF

   * 印刷可能、会議資料に

### 実装例（vis.js）

```html
<!DOCTYPE html>
<html>
<head>
  <script type="text/javascript" src="https://unpkg.com/vis-network/standalone/umd/vis-network.min.js"></script>
</head>
<body>
<div id="mynetwork"></div>
<script>
  var nodes = new vis.DataSet([
    {id: 1, label: '山田太郎\nhide@company.com', color: '#ff6b6b'},
    {id: 2, label: '佐藤花子\nsato@company.com', color: '#4ecdc4'}
  ]);
  var edges = new vis.DataSet([
    {from: 1, to: 2, arrows: 'to', label: '送信5通', width: 3}
  ]);
  var container = document.getElementById('mynetwork');
  var data = {nodes: nodes, edges: edges};
  var network = new vis.Network(container, data, {});
</script>
</body>
</html>
```

***

## 5. 実装時の注意点

### プライバシー・コンプライアンス

* **個人情報の扱い**: メールアドレス＋氏名のペアは個人情報

* **対策**:

  * 匿名化オプション（ID番号のみ表示）

  * 社内規定確認

  * 本人の同意取得（特に社外配布時）

### データの前処理

```python
# メーリングリスト除外
exclude_addresses = ['all@company.com', 'team@company.com']

# 自動送信メール除外
exclude_senders = ['noreply@', 'donotreply@']

# 同一人物の複数アドレス統合
alias_map = {
    'hide@company.com': 'hide@company.com',
    'hide.yamada@company.com': 'hide@company.com'  # エイリアス統合
}
```

### パフォーマンス

* メール数が数千件を超えるとグラフが複雑化

* **対策**: 期間指定（直近3ヶ月など）、重要度でフィルタ

***

## 6. 発展的な活用

### 時系列分析

```python
# 月ごとのコミュニケーション変化
monthly_graphs = []
for month in date_range:
    graph = build_network(emails_in_month)
    monthly_graphs.append(graph)

# アニメーション化 → 組織再編やプロジェクト開始の影響を可視化
```

### コミュニティ検出

```python
import networkx as nx
from networkx.algorithms import community

G = nx.Graph(edges)
communities = community.greedy_modularity_communities(G)

# 発見できること:
# - 部署を超えた非公式なチーム
# - 情報が孤立している部門
# - ブリッジパーソン（複数コミュニティをつなぐ人）
```

### Expert Network的活用

```python
# 特定トピックの専門家マップ
topic_keywords = ['ダイボンダ', 'void解析', 'X線']
expert_graph = filter_by_subject(emails, topic_keywords)

# 「この技術について知りたいなら、この3人に聞け」を可視化
```

***

## 7. 推奨される実装順序

### MVP（最小実用版）

1. ✅ win32comでOutlookフォルダ読み込み
2. ✅ From/To/CC抽出 → CSVエクスポート
3. ✅ NetworkXでグラフ構築
4. ✅ matplotlib で静的画像出力

### Phase 2（Obsidian連携）

1. ✅ Markdown形式でノート生成
2. ✅ Dataview pluginで集計
3. ✅ Graph Viewで可視化

### Phase 3（配布可能化）

1. ✅ vis.js でインタラクティブHTML
2. ✅ 匿名化オプション実装
3. ✅ 時系列アニメーション

***

## 8. 具体的提案

### このアイデアが刺さる理由

* **暗黙知の言語化**: いつもやっていること

* **独立志向との親和性**: 一人で完結できるツール

* **尊敬できる人との出会い**: このツールで社内外の専門家を発見

* **引継ぎ資料**: 替えが効かない状態からの脱却

### 最初の一歩

```bash
# 1. テスト用フォルダで小規模実験
python outlook_network_analyzer.py --folder "送信済みアイテム" --limit 100

# 2. 結果をObsidianにインポート
# 3. Graph Viewで「意外な発見」を探す
# 4. Slackで社内共有 → フィードバック収集
```

### 差別化ポイント

* 既存のメール分析ツール（例: Microsoft Viva Insights）は**個人の生産性**に焦点

* 私のツールは**組織の暗黙知**を可視化

* 半導体業界の技術伝承にも応用可能

***

## 9. 確信度の整理

| 項目          | 確信度   | 根拠                         |
| ----------- | ----- | -------------------------- |
| 技術的実現可能性    | ★★★★★ | win32com/Graph APIは実績多数    |
| Obsidian連携  | ★★★★☆ | Dataview/Graph Viewの活用事例あり |
| 引継ぎ資料としての価値 | ★★★★☆ | 暗黙知可視化は需要高い                |
| 配布形式の実現     | ★★★★☆ | vis.js等の成熟したライブラリ存在        |
| プライバシー対応    | ★★★☆☆ | 社内規定次第、要確認                 |

***

## 10. 次のアクション

### 今すぐできること

```python
# サンプルコード（概念実証）
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 受信トレイ

for message in inbox.Items[:10]:  # 最初の10件
    print(f"From: {message.SenderName}")
    print(f"To: {message.To}")
    print(f"CC: {message.CC}")
    print("---")
```

### 判断が必要なこと

1. **Obsidian vs HTML**: どちらをメイン出力にするか
2. **匿名化の程度**: フルネーム表示 or イニシャル or ID
3. **公開範囲**: 個人用 or チーム共有 or 全社展開

***

## 結論

**このアイデアは実装する価値がある（確信度：★★★★☆）**

理由:

* 私の「受容適応創造」サイクルに完全適合

* 既存ツール（Outlook, Obsidian）の新しい使い方

* 半導体業界の技術伝承という社会的意義

* Expert Network的な知見の蓄積にも貢献

懸念点:

* プライバシー対応の社内調整が必要

* グラフが複雑化した際のUI/UX設計

**まず小さく始めて、**私**自身の受信トレイで実験するのが吉**

***

## なぜ実現可能と断言できるか

### 1. 必要な技術スタックは既に習得済み

私**が既に使える技術：**

* ✅ Python（Streamlitアプリ開発経験）

* ✅ Obsidian連携（メール→Markdown変換を今まさに構築中）

* ✅ データ可視化（生産管理ダッシュボード、ガントチャート）

* ✅ Slack連携（メンションマップ、Timeline）

**追加で必要な技術（学習容易）：**

* `win32com.client`（OutlookのCOM操作）→ VBAマクロと同じ概念

* `NetworkX`（ネットワーク分析）→ 使いやすいライブラリ

* `vis.js` or `Plotly`（インタラクティブ可視化）→ Streamlit経験があれば容易

### 2. 段階的実装が可能

**Phase 1: データ抽出（1-2日）**

```python
import win32com.client
import pandas as pd
from datetime import datetime

def extract_email_network(folder_name="受信トレイ", limit=None):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = outlook.GetDefaultFolder(6)  # 受信トレイ
  
    network_data = []
  
    for i, message in enumerate(folder.Items):
        if limit and i >= limit:
            break
  
        try:
            # From/To/CCを抽出
            sender = message.SenderEmailAddress
            sender_name = message.SenderName
  
            # To（複数の場合がある）
            recipients_to = []
            for recipient in message.Recipients:
                if recipient.Type == 1:  # olTo
                    recipients_to.append({
                        'email': recipient.Address,
                        'name': recipient.Name
                    })
  
            # CC
            recipients_cc = []
            for recipient in message.Recipients:
                if recipient.Type == 2:  # olCC
                    recipients_cc.append({
                        'email': recipient.Address,
                        'name': recipient.Name
                    })
  
            # 件名、日時
            subject = message.Subject
            sent_on = message.ReceivedTime
  
            network_data.append({
                'sender_email': sender,
                'sender_name': sender_name,
                'recipients_to': recipients_to,
                'recipients_cc': recipients_cc,
                'subject': subject,
                'datetime': sent_on
            })
  
        except Exception as e:
            print(f"Error processing email {i}: {e}")
            continue
  
    return pd.DataFrame(network_data)

# 実行例
df = extract_email_network(limit=100)
df.to_csv('email_network_raw.csv', index=False, encoding='utf-8-sig')
print(f"抽出完了: {len(df)}件")
```

**Phase 2: ネットワーク構築（1-2日）**

```python
import networkx as nx
from collections import Counter
import json

def build_network_graph(df):
    G = nx.DiGraph()  # 有向グラフ
  
    # エッジの重み計算用
    edge_weights = Counter()
  
    for idx, row in df.iterrows():
        sender = row['sender_email']
  
        # ノードの属性を追加
        if sender not in G:
            G.add_node(sender,
                      name=row['sender_name'],
                      sent_count=0,
                      received_count=0,
                      cc_count=0)
  
        G.nodes[sender]['sent_count'] += 1
  
        # To宛先
        for recipient in row['recipients_to']:
            recipient_email = recipient['email']
  
            if recipient_email not in G:
                G.add_node(recipient_email,
                          name=recipient['name'],
                          sent_count=0,
                          received_count=0,
                          cc_count=0)
  
            G.nodes[recipient_email]['received_count'] += 1
  
            # エッジを追加
            edge_key = (sender, recipient_email, 'to')
            edge_weights[edge_key] += 1
  
        # CC宛先
        for recipient in row['recipients_cc']:
            recipient_email = recipient['email']
  
            if recipient_email not in G:
                G.add_node(recipient_email,
                          name=recipient['name'],
                          sent_count=0,
                          received_count=0,
                          cc_count=0)
  
            G.nodes[recipient_email]['cc_count'] += 1
  
            edge_key = (sender, recipient_email, 'cc')
            edge_weights[edge_key] += 1
  
    # エッジを追加（重み付き）
    for (src, dst, edge_type), weight in edge_weights.items():
        G.add_edge(src, dst, weight=weight, type=edge_type)
  
    return G

# 実行例
G = build_network_graph(df)
print(f"ノード数: {G.number_of_nodes()}")
print(f"エッジ数: {G.number_of_edges()}")

# NetworkXのグラフをJSON形式で保存
data = nx.node_link_data(G)
with open('email_network.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
```

**Phase 3: 可視化（2-3日）**

選択肢A: **Streamlitで簡易ダッシュボード**

```python
import streamlit as st
import plotly.graph_objects as go
import networkx as nx

st.title("📧 メールネットワーク可視化")

# データ読み込み
G = nx.read_gpickle('email_network.gpickle')

# フィルタリング
min_weight = st.slider("最小メール数", 1, 20, 3)

# グラフをフィルタ
G_filtered = G.copy()
edges_to_remove = [(u, v) for u, v, d in G_filtered.edges(data=True)
                   if d['weight'] < min_weight]
G_filtered.remove_edges_from(edges_to_remove)

# レイアウト計算
pos = nx.spring_layout(G_filtered, k=0.5, iterations=50)

# Plotlyで可視化
edge_trace = []
for edge in G_filtered.edges(data=True):
    x0, y0 = pos[edge[0]]
    x1, y1 = pos[edge[1]]
  
    edge_trace.append(
        go.Scatter(
            x=[x0, x1, None],
            y=[y0, y1, None],
            mode='lines',
            line=dict(
                width=edge[2]['weight'],
                color='red' if edge[2]['type'] == 'to' else 'green'
            ),
            hoverinfo='none'
        )
    )

# ノード
node_trace = go.Scatter(
    x=[pos[node][0] for node in G_filtered.nodes()],
    y=[pos[node][1] for node in G_filtered.nodes()],
    mode='markers+text',
    text=[G_filtered.nodes[node]['name'] for node in G_filtered.nodes()],
    textposition='top center',
    marker=dict(
        size=[G_filtered.nodes[node]['sent_count'] +
              G_filtered.nodes[node]['received_count']
              for node in G_filtered.nodes()],
        color='lightblue',
        line_width=2
    )
)

fig = go.Figure(data=edge_trace + [node_trace])
fig.update_layout(showlegend=False, hovermode='closest')

st.plotly_chart(fig, use_container_width=True)

# 統計情報
st.subheader("📊 統計")
col1, col2, col3 = st.columns(3)
col1.metric("総ノード数", G_filtered.number_of_nodes())
col2.metric("総エッジ数", G_filtered.number_of_edges())
col3.metric("平均次数", f"{sum(dict(G_filtered.degree()).values()) / G_filtered.number_of_nodes():.1f}")
```

選択肢B: **Obsidian Graph View用Markdown生成**

```python
def export_to_obsidian(G, output_folder="email_network"):
    import os
    os.makedirs(output_folder, exist_ok=True)
  
    # 各人物のノート作成
    for node in G.nodes():
        node_data = G.nodes[node]
  
        # ファイル名（メールアドレスから安全な名前を生成）
        safe_name = node.replace('@', '_at_').replace('.', '_')
        filename = f"{output_folder}/{safe_name}.md"
  
        # Markdown生成
        content = f"""---
type: person
email: {node}
sent: {node_data['sent_count']}
received: {node_data['received_count']}
cc: {node_data['cc_count']}
---

# {node_data['name']}

## メール統計
- 送信: {node_data['sent_count']}通
- 受信: {node_data['received_count']}通
- CC: {node_data['cc_count']}通

## 送信先
"""
  
        # 送信先リンク
        for successor in G.successors(node):
            successor_name = G.nodes[successor]['name']
            weight = G[node][successor]['weight']
            edge_type = G[node][successor]['type']
  
            content += f"- [[{successor_name}]] ({weight}通, {edge_type})\n"
  
        content += "\n## 受信元\n"
  
        # 受信元リンク
        for predecessor in G.predecessors(node):
            predecessor_name = G.nodes[predecessor]['name']
            weight = G[predecessor][node]['weight']
  
            content += f"- [[{predecessor_name}]] ({weight}通)\n"
  
        # ファイル書き込み
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(content)
  
    print(f"Obsidianノート生成完了: {G.number_of_nodes()}ファイル")

# 実行
export_to_obsidian(G, "C:/Users/YourName/Documents/Obsidian/EmailNetwork")
```

## 実装上の重要ポイント

### 1. **プライバシー対策（必須）**

```python
# 匿名化オプション
def anonymize_email(email, mapping=None):
    if mapping is None:
        mapping = {}
  
    if email not in mapping:
        mapping[email] = f"User_{len(mapping) + 1}"
  
    return mapping[email]

# 社外メール除外
def is_internal_email(email, company_domain="example.co.jp"):
    return company_domain in email
```

### 2. **パフォーマンス最適化**

```python
# 大量メール処理時
from tqdm import tqdm  # プログレスバー

for i in tqdm(range(len(folder.Items)), desc="メール処理中"):
    message = folder.Items[i]
    # 処理...
```

### 3. **エラーハンドリング**

```python
try:
    # Outlook操作
except Exception as e:
    logging.error(f"Error: {e}")
    continue  # スキップして続行
```

## 私に最適な理由

### 強みとの完全な適合

1. **「受容適応創造」サイクル**

   * 受容: Outlookという既存資産

   * 適応: Obsidianという使い慣れたツール

   * 創造: 暗黙知の可視化という新価値

2. **独立志向**

   * 一人で完結できるツール

   * 外部依存なし

3. **替えが効かない状態からの脱却**

   * 引継ぎ資料の自動生成

   * 組織の暗黙知を形式知化

4. **Expert Network的活用**

   * 社内の技術専門家マップ

   * 「誰に聞けばいいか」の可視化

## 次のアクション

```bash
# 1. 環境準備
pip install pywin32 networkx pandas plotly streamlit

# 2. MVP実装（Phase 1のみ）
python email_extractor.py --limit 50  # まず50件で試す

# 3. 結果確認
# email_network_raw.csv を Excel で開いて構造確認
```

**来週できること：**

* Phase 2実装（ネットワーク構築）

* Streamlitダッシュボード作成

* Obsidianへのエクスポート

