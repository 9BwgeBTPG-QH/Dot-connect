# Dot-connect 設計仕様書 — 他プラットフォームへの移植ガイド

Slack 等の他メッセージングプラットフォームに Dot-connect の可視化方式を移植する際の、設計上のポイントをまとめたドキュメント。

---

## 1. なぜサンキーダイアグラムよりネットワークグラフか

| 観点 | サンキーダイアグラム | Dot-connect (Force-directed Network) |
|------|---------------------|--------------------------------------|
| **表現できる関係** | 1→多 / 多→1 の流量（一方向） | 多↔多の双方向関係、CC/同報を含む |
| **コミュニティ発見** | 不可（流れの太さしか見えない） | Louvain 法で自動クラスタ検出 + Convex Hull 描画 |
| **キーマン特定** | 流量の多い人がわかる程度 | Degree/Betweenness centrality によるハブ検出 + CC キーマン検出 |
| **探索的分析** | 静的（フィルタ操作が限定的） | ノードクリックで接続関係にドリルダウン（Workflowy 式） |
| **スケーラビリティ** | 50人以上でラベルが重なり破綻 | 500+ノードに対応（物理エンジン + フィルタリング + クラスタ折りたたみ） |
| **インサイト** | 「誰→誰にたくさん送っている」 | 「この人は組織のハブ」「このグループは密に連携」「このCCは形骸化」 |

**結論**: サンキーは「流量の可視化」に特化しており、**組織構造・暗黙の関係性・キーマン発見**には不向き。Force-directed network + centrality 分析が適切。

---

## 2. アーキテクチャ概要

```
┌─────────────────────────────────────────────────────────────┐
│                    Data Source Layer                         │
│  Outlook COM / Slack API / CSV import                       │
└───────────────┬─────────────────────────────────────────────┘
                │ DataFrame (date, from, to, cc, subject)
                ▼
┌─────────────────────────────────────────────────────────────┐
│                    Analysis Pipeline                         │
│  1. build_graph()  → NetworkX DiGraph                       │
│  2. analyze_graph() → CC keymen, Hubs, Louvain communities  │
│  3. generate_vis_data() → JSON for frontend                 │
└───────────────┬─────────────────────────────────────────────┘
                │ JSON (nodes, edges, communities, analysis)
                ▼
┌─────────────────────────────────────────────────────────────┐
│                    Visualization Layer                       │
│  vis.js (network) + wordcloud2.js + Canvas (convex hull)    │
└─────────────────────────────────────────────────────────────┘
```

**ポイント**: Data Source Layer だけを差し替えれば、Analysis Pipeline と Visualization Layer はそのまま再利用できる。

---

## 3. データモデル（共通インターフェース）

### 3.1 入力: メッセージ DataFrame

移植時に合わせるべき**正規化された入力形式**。どのプラットフォームでも、このスキーマに変換すればパイプラインに乗る。

| カラム | 型 | 説明 | 例 |
|--------|------|------|------|
| `date` | str (ISO 8601) | メッセージ日時 | `2025-06-15 14:30:00` |
| `from_email` | str | 送信者の一意識別子 | `taro@example.co.jp` |
| `from_name` | str | 送信者の表示名 | `山田太郎` |
| `to` | str | 受信者リスト（セミコロン区切り） | `Name1 <email1>; Name2 <email2>` |
| `cc` | str | CC受信者リスト（同上） | `Name3 <email3>` |
| `subject` | str | 件名（分析には未使用、参考情報） | `Re: 会議の件` |

**Slack への適用**:
- `from_email` → Slack user ID またはメールアドレス
- `to` → DM: 相手 / Channel: チャンネル内でメンションされた人 / Thread reply: 元投稿者
- `cc` → チャンネルの他参加者（メンション無し）、またはリアクションした人

### 3.2 中間: NetworkX DiGraph

```python
# ノード属性
G.nodes["taro@example.co.jp"] = {
    "name": "山田太郎",
    "email": "taro@example.co.jp",
    "domain": "example.co.jp",        # Slack: workspace名
    "is_internal": True,               # 社内/社外判定
    "sent": 150,                       # 送信数
    "received": 200,                   # 受信数（To宛先として）
    "cc_count": 80,                    # CC出現数
    "community": 2,                    # Louvain コミュニティID
}

# エッジ属性（有向: from → to）
G.edges["taro@example.co.jp"]["hanako@example.co.jp"] = {
    "to_weight": 45,    # 直接送信回数
    "cc_weight": 12,    # CC に含まれた回数
}
```

### 3.3 出力: フロントエンド JSON

```json
{
  "nodes": [
    {
      "id": "taro@example.co.jp",
      "label": "山田太郎",
      "name": "山田太郎",
      "email": "taro@example.co.jp",
      "domain": "example.co.jp",
      "is_internal": true,
      "sent": 150,
      "received": 200,
      "cc_count": 80,
      "community": 2,
      "color": "#10b981",
      "size": 25,
      "is_cc_key": false,
      "is_hub": true
    }
  ],
  "edges": [
    {
      "from": "taro@example.co.jp",
      "to": "hanako@example.co.jp",
      "to_weight": 45,
      "cc_weight": 12,
      "weight": 57,
      "width": 4.5
    }
  ],
  "communities": [
    {
      "id": 0,
      "color": "#6366f1",
      "size": 12,
      "members": [{"email": "...", "name": "..."}]
    }
  ],
  "analysis": {
    "total_mails": 5000,
    "total_nodes": 150,
    "total_edges": 430,
    "cc_key_persons": [...],
    "hubs": [...]
  },
  "wordcloud_data": [
    {"text": "山田太郎", "size": 350, "email": "...", "color": "#10b981"}
  ]
}
```

---

## 4. 分析パイプライン詳細

### 4.1 グラフ構築 (`build_graph`)

```
入力 DataFrame → 行ごとに走査:
  1. from_email → ノード作成/更新 (sent++)
  2. to フィールドをパース → 各宛先のノード作成/更新 (received++)
     → エッジ追加: from → to (to_weight++)
  3. cc フィールドをパース → 各CC先のノード作成/更新 (cc_count++)
     → エッジ追加: from → cc (cc_weight++)
  4. ノードに社内/社外フラグを付与 (company_domains で判定)
```

**設計判断**:
- **有向グラフ** (DiGraph): 送信方向を保持。分析時に無向化して centrality 計算
- **エッジ重みの分離**: `to_weight` と `cc_weight` を分けて保持。CCだけの関係とToの直接関係を区別できる
- **ノードサイズ**: `sent + received + cc_count` の合計に比例（8〜40px, `8 + total * 0.5`）

### 4.2 分析 (`analyze_graph`)

#### CC キーマン検出
```
各ノードの cc_count / total_mails >= threshold (default: 0.30)
→ 閾値を超えた人物を CC Key Person としてマーク
```
**意味**: 「メール総数の30%以上のCCに含まれる = 形式的に情報共有されている人」。承認者・管理職層が検出されやすい。

#### ハブ検出 (Centrality 分析)
```
undirected = G.to_undirected()
degree_centrality    = nx.degree_centrality(undirected)
betweenness_centrality = nx.betweenness_centrality(undirected)

hub_score = degree_weight * degree_c + betweenness_weight * betweenness_c
→ 上位20名をハブとしてマーク
```
- **Degree centrality**: 接続数の多さ（多くの人とやり取り）
- **Betweenness centrality**: 最短経路の仲介度（異なるグループを橋渡し）
- 重み配分はユーザーが調整可能（default: 各0.5）

#### コミュニティ検出 (Louvain 法)
```
communities = nx.community.louvain_communities(undirected, seed=42)
→ 各ノードに community ID を付与
→ コミュニティごとに色を割り当て (12色パレット)
```
**seed固定**: 同じデータなら同じ結果を保証

### 4.3 可視化データ生成 (`generate_vis_data`)

- **エッジフィルタ**: `to_weight + cc_weight < min_edge_weight` のエッジを除外（ノイズ除去）
- **エッジ太さ**: `max(1, min(8, total * 0.3))` — 重みに比例、1〜8px の範囲
- **ラベル省略**: 15文字超の名前は14文字+`…` に切り詰め
- **ワードクラウドデータ**: ノードごとの総活動量をサイズに変換

---

## 5. フロントエンド可視化の設計

### 5.1 vis.js ネットワーク設定

| 設定 | 値 | 理由 |
|------|------|------|
| 物理エンジン | `barnesHut` | 大規模グラフに適したO(n log n)アルゴリズム |
| `gravitationalConstant` | `-3000` | ノード間の反発力（大きいほど広がる） |
| `centralGravity` | `0.1` | 中心への引力（散らばりすぎ防止） |
| `springLength` | `120` | エッジの自然長 |
| `damping` | `0.3` | 振動の減衰率 |
| `stabilization.iterations` | `200` | 初期安定化の反復回数 |

### 5.2 インタラクション設計

```
クリック (1回)   → ノードの接続関係にフォーカス
                   接続ノードだけを残して再描画
                   サイドパネルに詳細表示
                   パンくずナビ更新: All > 山田太郎

クリック (空白)  → 全体表示にリセット

ダブルクリック   → コミュニティの折りたたみ/展開
                   折りたたみ: 同コミュニティのノードを1つのクラスターノードに
                   展開: クラスターを解除

凡例クリック     → コミュニティの全メンバーを選択 & フィット表示
```

### 5.3 Convex Hull (コミュニティ境界)

```
vis.js の afterDrawing イベントで Canvas 直接描画:
  1. 各コミュニティのメンバーの現在座標を取得
  2. Graham Scan で凸包を計算
  3. 重心から外側に 40px オフセット（パディング）
  4. 半透明の塗り (opacity ~0.08) + 境界線 (opacity ~0.15)
```
**ポイント**: vis.js のレイヤーの上に Canvas で直接描画するため、物理エンジンの更新と同期する。

### 5.4 サイドパネル

ノードクリック時に表示。以下の情報を含む:
- 名前 + メールアドレス（コピーボタン）
- バッジ: CC Key Person / Hub
- 統計: 送信数・受信数・CC数
- 展開可能リスト: To先・CC先・受信元（各人の通数を表示、クリックでそのノードにフォーカス）

### 5.5 ワードクラウドビュー

- wordcloud2.js で人名をフォントサイズ = メール頻度で表示
- コミュニティカラーと連動
- クリックでネットワークビューの該当ノードにジャンプ

---

## 6. Slack 移植時の変換マッピング

### 6.1 概念マッピング

| Outlook の概念 | Slack の概念 | マッピング方法 |
|----------------|-------------|----------------|
| メール1通 | メッセージ1件 | 1メッセージ = 1レコード |
| 送信者 (from) | メッセージ投稿者 | `user_id` → `from_email` |
| To受信者 | メンション先 | `<@USER_ID>` をパースして `to` に |
| CC受信者 | チャンネル参加者（メンション無し） | チャンネルメンバー - メンション先 = `cc` |
| メールフォルダ | チャンネル / DM | 抽出対象の選択単位 |
| メールアドレス | Slack user ID | ノードの一意識別子 |
| ドメイン (社内/社外判定) | ワークスペース | `is_internal` の判定基準 |
| 件名 | スレッドの最初のメッセージ | 参考情報 |

### 6.2 Slack 特有の考慮点

**スレッド応答の扱い**:
```
# 案A: スレッド応答 → 元投稿者への返信として扱う
if message.thread_ts != message.ts:
    to = [thread_parent_author]

# 案B: スレッド参加者全員を to として扱う
if message.thread_ts:
    to = [thread内の他の投稿者全員]
```
推奨は**案B**。スレッド内のやり取りは実質的な対話関係を反映する。

**リアクション の扱い**:
```
# リアクション = 軽量なCC（受動的な参加表明）
for reaction in message.reactions:
    for user in reaction.users:
        cc_list.append(user)
```

**チャンネル参加者の扱い**:
- 全参加者をCCにすると大規模チャンネルでノイズが増大
- **推奨**: メンション + スレッド参加 + リアクション のみを関係として抽出

### 6.3 Slack API での抽出フロー

```python
# 1. チャンネル一覧取得
channels = slack_client.conversations_list()

# 2. メッセージ取得（期間指定）
messages = slack_client.conversations_history(
    channel=channel_id,
    oldest=start_timestamp,
    latest=end_timestamp,
)

# 3. スレッド応答取得
for msg in messages:
    if msg.get("reply_count", 0) > 0:
        replies = slack_client.conversations_replies(
            channel=channel_id,
            ts=msg["ts"],
        )

# 4. ユーザー情報取得（ID → 名前/メール変換）
user_info = slack_client.users_info(user=user_id)
email = user_info["user"]["profile"]["email"]
name = user_info["user"]["real_name"]
```

---

## 7. 設定パラメータ一覧

| パラメータ | デフォルト | 説明 | 調整の指針 |
|-----------|-----------|------|-----------|
| `company_domains` | `[]` | 社内ドメインリスト | ノードの色分け（社内/社外）に使用 |
| `cc_key_person_threshold` | `0.30` | CC出現率の閾値 | 小さくすると多くの人がキーマンに |
| `min_edge_weight` | `1` | エッジ表示の最小重み | 大きくするとノイズ除去、グラフが疎に |
| `hub_degree_weight` | `0.5` | ハブスコアの degree 重み | 接続数重視なら上げる |
| `hub_betweenness_weight` | `0.5` | ハブスコアの betweenness 重み | 仲介役重視なら上げる |

---

## 8. 移植チェックリスト

### Phase 1: データ抽出レイヤー
- [ ] Slack API 接続・認証 (OAuth2 Bot Token)
- [ ] チャンネル一覧取得 → フォルダ選択UIに相当
- [ ] メッセージ抽出 → DataFrame 変換
- [ ] メンション / スレッド / リアクション のパース
- [ ] ユーザー ID → 表示名・メールの解決

### Phase 2: パイプライン統合
- [ ] DataFrame スキーマを Dot-connect 形式に合わせる
- [ ] `build_graph()` → `analyze_graph()` → `generate_vis_data()` をそのまま呼び出し
- [ ] Web UI のデータソース選択を追加（Outlook / Slack / CSV）

### Phase 3: UI 調整
- [ ] サイドパネルの表示項目を Slack 向けに調整（メールアドレス → Slack handle）
- [ ] 「CC Key Person」のラベルを「Passive Observer」等に変更（Slack の文脈に合わせる）
- [ ] チャンネル情報をコミュニティの補足情報として表示

### Phase 4: 運用
- [ ] Slack Bot のスコープ設定 (`channels:history`, `users:read`, etc.)
- [ ] Rate limit 対応（Slack API は Tier制）
- [ ] 定期実行 / 差分更新の仕組み

---

## 9. コアモジュールの再利用ガイド

移植時に**そのまま再利用できるファイル**と**差し替えが必要なファイル**:

| ファイル | 再利用 | 備考 |
|---------|--------|------|
| `app/core.py` | **そのまま** | `build_graph`, `analyze_graph`, `generate_vis_data` は入力DFのスキーマさえ合えば動く |
| `templates/network.html` | **そのまま** | JSON 構造が同じなら変更不要 |
| `templates/upload.html` | **一部変更** | データソース選択UIを追加 |
| `app/main.py` | **一部変更** | 抽出エンドポイントを Slack 用に追加 |
| `extract.py` | **差し替え** | Outlook COM → Slack API に差し替え |
| `app/extract.py` | **差し替え** | Web用ラッパーを Slack API 用に差し替え |
| `extract_and_upload.py` | **不要** | Slack API はサーバーから直接アクセス可能（ローカル実行不要） |
| `config.yaml` | **拡張** | Slack Bot Token, 対象チャンネル等を追加 |
