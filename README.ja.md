# Dot-connect — Outlook メールネットワーク可視化ツール

Outlookメールの送受信・CC関係をネットワークグラフで可視化し、**CCキーマン**や**ハブ人物**を自動特定するツール。組織内の暗黙的なコミュニケーション構造を「見える化」し、引き継ぎ資料やナレッジトランスファーに活用できる。

> **[デモを見る（サンプルデータ）](https://9BwgeBTPG-QH.github.io/Dot-connect/)** — 架空のチームデータで実際のダッシュボードを操作できます。

## 使い方（かんたん）

**Python や CLI の知識は不要です。**

### 初回セットアップ（1回だけ）

1. `setup.bat` をダブルクリック
2. Python と必要なパッケージが自動でインストールされる（数分）

### 方法A: ローカルPCで起動（最もシンプル）

1. `start.bat` をダブルクリック → ブラウザが自動で開く
2. 「Outlook から抽出」タブで:
   - メールフォルダを選択（チェックボックス、ページ読込時に自動取得）
   - 期間を指定
   - 「抽出 & 分析する」を押す
3. ネットワーク可視化が表示される

> 既に抽出済みの CSV がある場合は「CSV アップロード」タブからも分析可能。

### 方法B: ファイルサーバーで起動（複数人で共有）

ファイルサーバーで `start.bat` を実行し、各ユーザーはブラウザからアクセスする。
**ユーザーのPCに Python のインストールは不要** — 共有フォルダの Python、またはポータブル Python の自動ダウンロードにより動作する。

**サーバー側（管理者が1回だけ）:**

1. 共有フォルダに Dot-connect を配置（例: `\\SERVER\share\Dot-connect`）
2. `setup.bat` → `start.bat` を実行（サーバー上で常時起動）
3. `config.yaml` の `network_share_path` を実際の共有パスに設定
4. Windows Firewall でポート 8000 の受信を許可

```
netsh advfirewall firewall add rule name="Dot-connect" dir=in action=allow protocol=TCP localport=8000
```

**各ユーザー（毎回）:**

1. ブラウザで `http://<サーバー名>:8000` にアクセス
2. 「Outlook から抽出」タブで期間を指定し「抽出ツールをダウンロード (.bat)」をクリック
3. ダウンロードした `.bat` をダブルクリックで実行
4. Outlook フォルダを選択 → 自動でメール抽出 & サーバーにアップロード
5. ブラウザに分析結果が自動表示される

```
\\server\share\dot-connect\
├── setup.bat                ← 初回セットアップ
├── start.bat                ← サーバー起動
├── extract_and_upload.py    ← ローカル抽出スクリプト（自動DL）
├── config.yaml              ← network_share_path を設定
├── python\                  ← 埋め込みPython（setup.batで生成、ユーザーPCから共有利用）
└── ...
```

> `.bat` はサーバーからスクリプトをダウンロードして実行する。Python は共有フォルダ → PATH → 自動ダウンロードの順に検索され、見つからない場合はポータブル版 Python を `%TEMP%` に自動インストールする。pywin32 も未導入なら自動インストールされるため、ユーザーのPCに事前準備は不要。

---

## パイプライン

```
方法A (ローカル):    start.bat → ブラウザで Outlook フォルダ選択 → 可視化
方法B (サーバー):    ブラウザ → .bat DL → ローカル実行 → サーバーで分析 → 可視化
方法C (CSV):        start.bat → ブラウザで CSV アップロード → 可視化
方法D (Graph API):  ブラウザ → Microsoft サインイン → フォルダ選択 → 可視化
方法E (CLI):        extract.py → CSV → generate.py → index.html
```

## ファイル構成

```
Dot-connect/
├── setup.bat                # 初回セットアップ（Python自動DL）
├── start.bat                # サーバー起動（ダブルクリック）
├── extract_and_upload.py    # ローカル抽出 & サーバーアップロード
├── app/
│   ├── __init__.py          # パッケージ初期化
│   ├── core.py              # 分析コアロジック（CLI/Web共通）
│   ├── extract.py           # Outlook COM ラッパー（Web用）
│   ├── graph_auth.py        # Graph API OAuth2 認証（MSAL + PKCE）
│   ├── graph_extract.py     # Graph API メール抽出
│   ├── main.py              # FastAPI アプリ
│   └── models.py            # Pydantic バリデーションモデル
├── templates/
│   ├── upload.html          # Web UI: トップページ（抽出 / アップロード）
│   └── network.html         # 可視化テンプレート（vis.js + wordcloud2.js）
├── docs/
│   └── GRAPH_API_SETUP.md   # Graph API セットアップガイド
├── extract.py               # CLI: Outlook → CSV抽出
├── generate.py              # CLI: CSV → HTML生成
├── config.yaml.example      # 設定テンプレート（config.yaml にコピーして使用）
├── requirements.txt         # 依存パッケージ
├── requirements-extract.txt # CLI用 pywin32
├── python/                  # 埋め込みPython（setup.batで自動生成）
└── output/                  # 生成物（gitignore対象）
```

## 開発者向けセットアップ

既に Python 環境がある場合は `setup.bat` を使わず直接インストールできる。

```bash
cp config.yaml.example config.yaml   # 初回のみ
pip install -r requirements.txt
pip install pywin32                   # Windows でメール抽出する場合
uvicorn app.main:app --reload         # 開発サーバー起動
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

### 方法D: Microsoft 365（Graph API）

Outlook COM が使えない環境（New Outlook、Outlook on the web、macOS/Linux）向け:

1. Microsoft Entra ID にアプリ登録 — 詳細は [Graph API セットアップガイド](docs/GRAPH_API_SETUP.md) を参照
2. Web UI 右上の「設定」→ `Client ID` と `Tenant ID` を入力して保存
3. トップページに戻る →「Microsoft 365」タブが表示される
4. Microsoft アカウントでサインイン → フォルダ選択 → 抽出 & 分析

> Exchange Online ライセンスと管理者の同意が必要です。詳細は [docs/GRAPH_API_SETUP.md](docs/GRAPH_API_SETUP.md) を参照してください。

---

## 動作要件と制約

本ツールは2つの抽出方法をサポートしています:

| | COM（方法A/B） | Graph API（方法D） |
|--|--|--|
| 管理者の承認 | **不要** | Microsoft Entra ID アプリ登録 + 管理者の同意 |
| 認証 | なし（ローカルOutlookに接続） | OAuth2（PKCE、クライアントシークレット不要） |
| 対応Outlook | Classic（デスクトップ版）のみ | New Outlook / Web / Classic |
| ネットワーク | 不要（ローカル処理） | Microsoft 365 API呼び出し |
| 取得範囲 | 自分のメールボックス | 自分のメールボックス |
| 必要なライセンス | Outlook Classic がインストール済み | Exchange Online |

**COM** が最もシンプル — セットアップ不要、`start.bat` を実行するだけ。Outlook Classic がある環境ではこちらを推奨。

**Graph API** は COM が使えない環境をカバー（New Outlook、Web、macOS/Linux）。初回のみ Azure セットアップが必要。詳細は [docs/GRAPH_API_SETUP.md](docs/GRAPH_API_SETUP.md) を参照。

> M365 C2R 環境などサーバー上で COM が制限される場合は、方法Bを使用してください。各ユーザーのPCの Outlook Classic からメールを抽出し、サーバーにアップロードします。

## config.yaml 設定

初回は `config.yaml.example` をコピーして `config.yaml` を作成してください:

```bash
cp config.yaml.example config.yaml
```

> `config.yaml` は `.gitignore` に含まれており、Git にコミットされません。

```yaml
# ネットワーク共有パス（ファイルサーバー運用時に設定）
network_share_path: "\\\\SERVER\\Dot-connect"

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
| Outlook COM 制限 | M365 C2R 環境では COM が使えないため、ローカル抽出ツール方式で回避（方法B）。COM 結果はキャッシュされ、2回目以降のアクセスは即座に応答 |
| CSV エンコーディング | utf-8-sig → cp932 → latin-1 の順で自動判定 |
| config.yaml エンコーディング | BOM / NULバイト / UTF-16 残骸を自動除去して読み込み。どのエディタで保存しても動作 |

## プライバシーとデータの取り扱い

### 収集するデータ

本ツールは Outlook COM 経由で以下のメタデータを抽出します:

- 送信者のメールアドレスと表示名
- To/CC 受信者のメールアドレスと表示名
- 受信日時と件名

**メール本文は一切収集・保存されません。**

### データの処理と保存

- すべての処理は**ローカルPC（または自社サーバー）上で完結**します。外部サービスへのデータ送信はありません
- ファイルサーバーモード（方法B）では、抽出した CSV は自社の内部サーバーにのみアップロードされます
- 分析結果はメモリ上に保持され、サーバー停止時に破棄されます
- エクスポートした HTML ファイルには集計済みのネットワークデータ（氏名、メールアドレス、通信回数）が含まれます — **共有先にご注意ください**

### 管理者への推奨事項

- 本ツールの導入前に、メール通信パターンが可視化されることを**従業員に周知**してください
- 出力結果は組織のコミュニケーション構造を明らかにするため、**機密情報として扱ってください**
- 組織のデータポリシーおよび関連法規（GDPR、個人情報保護法等）に準拠した運用をお願いします
