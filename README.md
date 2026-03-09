# Dot-connect — Outlook メールネットワーク可視化ツール

Outlookメールの送受信・CC関係をネットワークグラフで可視化し、**CCキーマン**や**ハブ人物**を自動特定するツール。組織内の暗黙的なコミュニケーション構造を「見える化」し、引き継ぎ資料やナレッジトランスファーに活用できる。

## 使い方（かんたん）

**Python や CLI の知識は不要です。**

### 初回セットアップ（1回だけ）

1. `setup.bat` をダブルクリック
2. Python と必要なパッケージが自動でインストールされる（数分）

### 日常の操作

1. `start.bat` をダブルクリック → ブラウザが自動で開く
2. 「Outlook から抽出」タブで:
   - メールフォルダを選択（チェックボックス、ページ読込時に自動取得）
   - 期間を指定
   - 「抽出 & 分析する」を押す
3. ネットワーク可視化が表示される

> 既に抽出済みの CSV がある場合は「CSV アップロード」タブからも分析可能。

### 共有フォルダで配布する場合

フォルダごとコピーするだけ。各ユーザーの PC で `start.bat` を実行すれば、そのユーザーの Outlook メールが可視化される。

```
\\server\share\dot-connect\
├── setup.bat     ← 初回ダブルクリック
├── start.bat     ← 毎回ダブルクリック
├── python\       ← 自動生成される（配布時は含めてもOK）
└── ...
```

> **既知の制限**: 埋め込み Python 環境では、Microsoft 365 Click-to-Run 版 Outlook の COM 登録が見えないケースがあります（`クラス文字列が無効です` エラー）。この場合はユーザーの PC にインストール済みの Python を使うか、CSV アップロード方式を利用してください。詳細は [Issue #1](../../issues) を参照。

---

## パイプライン

```
方法A (Web UI):   start.bat → ブラウザで Outlook フォルダ選択 → 可視化
方法B (Web UI):   start.bat → ブラウザで CSV アップロード → 可視化
方法C (CLI):      extract.py → CSV → generate.py → index.html
```

## ファイル構成

```
Dot-connect/
├── setup.bat              # 初回セットアップ（Python自動DL）
├── start.bat              # サーバー起動（ダブルクリック）
├── app/
│   ├── __init__.py        # パッケージ初期化
│   ├── core.py            # 分析コアロジック（CLI/Web共通）
│   ├── extract.py         # Outlook COM ラッパー（Web用）
│   ├── main.py            # FastAPI アプリ
│   └── models.py          # Pydantic バリデーションモデル
├── templates/
│   ├── upload.html        # Web UI: トップページ（抽出 / アップロード）
│   └── network.html       # 可視化テンプレート（vis.js + wordcloud2.js）
├── extract.py             # CLI: Outlook → CSV抽出
├── generate.py            # CLI: CSV → HTML生成
├── config.yaml            # 除外設定・エイリアス・閾値
├── requirements.txt       # 依存パッケージ
├── requirements-extract.txt # CLI用 pywin32
├── python/                # 埋め込みPython（setup.batで自動生成）
└── output/                # 生成物（gitignore対象）
```

## 開発者向けセットアップ

既に Python 環境がある場合は `setup.bat` を使わず直接インストールできる。

```bash
pip install -r requirements.txt
pip install pywin32               # Windows でメール抽出する場合
uvicorn app.main:app --reload     # 開発サーバー起動
```

### CLI での使い方

```bash
# メール抽出
python extract.py --start 2025-01-01 --end 2025-12-31

# HTML 生成
python generate.py --input output/emails_20250101.csv
# → output/index.html をブラウザで開く
```

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
| Outlook COM 制限 | 埋め込み Python では M365 C2R 版 Outlook の COM が見えない場合あり。通常インストール Python を推奨 |
| CSV エンコーディング | utf-8-sig → cp932 → latin-1 の順で自動判定 |
