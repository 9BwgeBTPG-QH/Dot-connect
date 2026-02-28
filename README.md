# Dot-connect — Outlook メールネットワーク可視化ツール

Outlookメールの送受信・CC関係をネットワークグラフで可視化し、**CCキーマン**や**ハブ人物**を自動特定するツール。組織内の暗黙的なコミュニケーション構造を「見える化」し、引き継ぎ資料やナレッジトランスファーに活用できる。

## パイプライン

```
Outlook → [extract.py] → CSV → [generate.py] → index.html（ブラウザで開く）
               ↑
          config.yaml
```

## ファイル構成

```
Dot-connect/
├── extract.py          # Outlook COM → CSV抽出
├── generate.py         # CSV → NetworkX分析 → HTML生成
├── config.yaml         # 除外設定・エイリアス・閾値
├── templates/
│   └── network.html    # vis.js + wordcloud2.js テンプレート
├── requirements.txt
└── output/             # 生成物（gitignore対象）
    ├── emails_YYYYMMDD.csv
    └── index.html
```

## セットアップ

```bash
pip install -r requirements.txt
```

依存パッケージ: pywin32, networkx (>=3.2), pandas, pyyaml, jinja2, tqdm

> **注意**: `extract.py` は Outlook COM を使うため **Windows Python** が必須。`generate.py` 以降は WSL/Mac/Linux でも動作する。

## 使い方

### 1. メール抽出

```bash
python extract.py --start 2025-01-01 --end 2025-12-31
```

- 対話的にメールフォルダを選択（カンマ区切りで複数可）
- DASL フィルタで Outlook 側で日付絞り込み（高速）
- Exchange DN (`/o=Org/...`) は SMTP アドレスに自動変換（多段フォールバック）
- 出力: `output/emails_YYYYMMDD.csv`（UTF-8 BOM、Excel対応）

### 2. HTML生成

```bash
python generate.py --input output/emails_20250101.csv
```

- NetworkX で有向グラフ構築・分析
- 出力: `output/index.html`（自己完結型、サーバ不要）

### 3. ブラウザで開く

`output/index.html` をブラウザで開くだけ。

## 可視化の機能

### ネットワークビュー
- **Obsidian Graph View 風ダークテーマ** (#1a1a2e)
- **vis.js** barnesHut 物理エンジンによるフォースレイアウト
- **ノードクリック → ズームイン**: その人の接続のみに絞り込み表示（Workflowy 式）
- **パンくずナビ**: `◀ All > 山田太郎` で階層遷移
- **情報パネル**: 氏名、メール（コピーボタン）、送受信/CC 数、展開リスト（To先/CC先/受信元）
- **クラスタ境界**: Louvain コミュニティを convex hull で描画
- **クラスタ折りたたみ**: ダブルクリックでコミュニティ単位の折りたたみ/展開
- **凡例**: コミュニティ色・CCキーマン一覧

### ワードクラウドビュー
- **wordcloud2.js** で人名表示（フォントサイズ = メール頻度）
- コミュニティ色に対応
- クリックでネットワークビューの該当ノードにジャンプ

## 分析機能

| 分析 | 説明 |
|------|------|
| **CC キーマン** | CC 出現率が閾値（デフォルト 30%）を超える人物 |
| **ハブ** | degree centrality + betweenness centrality の加重スコア上位 |
| **コミュニティ** | Louvain 法による自動クラスタ検出 |

## config.yaml 設定

```yaml
# 社内ドメイン（色分けに使用）
company_domains:
  - example.co.jp

# 除外（完全一致）
exclude_addresses:
  - noreply@example.co.jp

# 除外（正規表現）
exclude_patterns:
  - "^no-?reply@"

# エイリアス統合（同一人物の複数アドレスを正規化）
alias_map:
  taro.yamada@example.co.jp:
    - yamada.taro@old-domain.co.jp

# 分析閾値
thresholds:
  cc_key_person_threshold: 0.30  # CC出現割合
  min_edge_weight: 1             # 最小エッジ重み
  hub_degree_weight: 0.5         # ハブスコア重み
  hub_betweenness_weight: 0.5
```

## 技術的な注意点

| 課題 | 対策 |
|------|------|
| Exchange DN → SMTP 変換 | ExchangeUser → AddressEntry → PropertyAccessor → ダミー生成 の4段フォールバック |
| 大規模グラフ (500+ ノード) | `min_edge_weight` フィルタ、小クラスタ自動折りたたみ |
| CSV 内セミコロン衝突 | pandas `QUOTE_ALL` でフィールドを引用符囲い |
| pywin32 は WSL 不可 | `extract.py` のみ Windows Python 必須、`generate.py` は環境不問 |
