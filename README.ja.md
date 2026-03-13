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
**ユーザーのPCに Python のインストールは不要** — 共有フォルダの埋め込み Python が自動的に使用される。

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

> `.bat` はサーバーからスクリプトをダウンロードし、共有フォルダの Python で実行する。ユーザーのPCに Python がなくても動作する。

---

## パイプライン

```
方法A (ローカル):    start.bat → ブラウザで Outlook フォルダ選択 → 可視化
方法B (サーバー):    ブラウザ → .bat DL → ローカル実行 → サーバーで分析 → 可視化
方法C (CSV):        start.bat → ブラウザで CSV アップロード → 可視化
方法D (CLI):        extract.py → CSV → generate.py → index.html
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
│   ├── main.py              # FastAPI アプリ
│   └── models.py            # Pydantic バリデーションモデル
├── templates/
│   ├── upload.html          # Web UI: トップページ（抽出 / アップロード）
│   └── network.html         # 可視化テンプレート（vis.js + wordcloud2.js）
├── extract.py               # CLI: Outlook → CSV抽出
├── generate.py              # CLI: CSV → HTML生成
├── config.yaml              # 除外設定・エイリアス・閾値・共有パス
├── requirements.txt         # 依存パッケージ
├── requirements-extract.txt # CLI用 pywin32
├── python/                  # 埋め込みPython（setup.batで自動生成）
└── output/                  # 生成物（gitignore対象）
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

## 動作要件と制約

本ツールは **Outlook COM オートメーション（MAPI）** でローカルのメールボックスに直接アクセスします。

**管理者の承認は不要です** — Microsoft Graph API と異なり、Azure AD アプリ登録もテナント管理者の同意も OAuth2 認証も必要ありません。自分のPCで `start.bat` を実行するだけで使えます。

| | COM（本ツール） | Graph API |
|--|--|--|
| 管理者の承認 | **不要** | Azure AD アプリ登録 + テナント管理者の同意が必要 |
| 認証 | なし（ローカルOutlookに接続） | OAuth2 フロー |
| 対応Outlook | Classic（デスクトップ版）のみ | New Outlook / Web / Classic |
| ネットワーク | 不要（ローカル処理） | Microsoft 365 API呼び出し |
| 取得範囲 | 自分のメールボックス | 権限設定次第 |

**対応:** Outlook Classic（MAPI対応のデスクトップ版）
**非対応:** New Outlook（ストアアプリ版）、Outlook on the web

> M365 C2R 環境などサーバー上で COM が制限される場合は、方法Bを使用してください。各ユーザーのPCの Outlook Classic からメールを抽出し、サーバーにアップロードします。

## config.yaml 設定

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
