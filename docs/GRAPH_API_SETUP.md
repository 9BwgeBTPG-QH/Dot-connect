# Microsoft Graph API セットアップガイド

## 概要

Dot-connect は通常 Outlook COM (MAPI) を使ってローカルのメールボックスからメールを抽出します。しかし、以下の環境では COM が利用できません:

- **New Outlook** (Windows)
- **Outlook on the web** (OWA)
- **サーバー上の M365 C2R 環境**
- **macOS / Linux**

Microsoft Graph API を使えば、これらの環境でもメール抽出が可能になります。

### COM vs Graph API

| 項目 | COM (従来) | Graph API (新規) |
|------|-----------|-----------------|
| 対応環境 | Outlook Classic (デスクトップ版) のみ | すべての Microsoft 365 環境 |
| 必要な準備 | なし | Microsoft Entra ID アプリ登録 + 管理者承認 |
| メール本文 | 収集しない | 収集しない |
| 認証 | 不要 | Microsoft アカウントでサインイン |

---

## 1. Microsoft Entra ID アプリ登録手順

> **補足**: Azure Active Directory は **Microsoft Entra ID** に名称変更されました。Azure Portal 上では「Microsoft Entra ID」と表示されます。

### 1.1 Azure Portal にアクセス

1. [Azure Portal](https://portal.azure.com) にサインイン
2. 上部の検索バーで「**Microsoft Entra ID**」を検索してクリック
3. 左メニューの **App registrations**（アプリの登録）をクリック
4. **New registration**（新規登録）をクリック

> **注意**: 「Azure AD B2C」は別のサービスです。必ず「**Microsoft Entra ID**」を選んでください。

### 1.2 アプリ情報を入力

| 項目 | 値 |
|------|-----|
| 名前 | `Dot-connect Mail Analyzer` |
| サポートされているアカウントの種類 | 「この組織ディレクトリのみに含まれるアカウント」(Accounts in this organizational directory only) |
| リダイレクト URI | **この時点では空欄のまま**（次のステップで設定します） |

**Register** をクリックしてアプリを登録します。

### 1.3 リダイレクト URI の設定

> **重要**: プラットフォームは必ず **Mobile and desktop applications** を選んでください。**Web** を選ぶとクライアントシークレットが要求されエラーになります。

1. 登録したアプリの左メニューで **Authentication** をクリック
2. **Add a platform** をクリック
3. **Mobile and desktop applications** を選択
4. カスタム URI の欄に `http://localhost:8000/auth/callback` を入力
5. **Configure** をクリックして保存

> **注意**: 本番環境ではリダイレクト URI を実際のサーバー URL に変更してください。

### 1.4 アプリケーション ID を取得

登録後の「概要」ページで以下をメモ:
- **アプリケーション (クライアント) ID** → `config.yaml` の `client_id`
- **ディレクトリ (テナント) ID** → `config.yaml` の `tenant_id`

---

## 2. API 権限設定

### 2.1 権限を追加

1. 登録したアプリの左メニューで **Call APIs** → **View API permissions** をクリック
2. **Add a permission** をクリック
3. **Microsoft Graph** → **Delegated permissions**（委任されたアクセス許可）を選択
4. 以下を検索して追加:
   - `Mail.Read` — メールの読み取り (読み取り専用)
   - `User.Read` — サインインとプロフィールの読み取り

### 2.2 管理者の同意を付与

組織の管理者が「管理者の同意を与えます」ボタンをクリックする必要があります。

> **ヒント**: 管理者でない場合は、下記の「管理者承認リクエストテンプレート」を使って申請してください。

---

## 3. Web UI で接続設定

1. Dot-connect サーバーを起動（`config.yaml` は初回起動時に自動生成されます）
2. ブラウザで `http://localhost:8000` にアクセス
3. 画面右上の **「設定」** リンクをクリック
4. 以下を入力して **「保存する」** をクリック:
   - **Client ID**: 1.4 でメモしたアプリケーション (クライアント) ID
   - **Tenant ID**: 1.4 でメモしたディレクトリ (テナント) ID
   - **Redirect URI**: 通常は `http://localhost:8000/auth/callback` のまま
5. トップページに戻ると **「Microsoft 365」タブ** が表示されます

> **補足**: 設定は `config.yaml` に保存されます。このファイルは `.gitignore` に含まれているため、Git にコミットされません。`config.yaml` を直接編集することもできますが、Web UI からの設定が推奨です。

> `client_id` と `tenant_id` が空の場合、Microsoft 365 タブは表示されません。

---

## 4. 管理者承認リクエストテンプレート

### 日本語版

```
件名: Microsoft Entra ID アプリ登録申請 — Dot-connect メールネットワーク可視化

申請者: [あなたの名前]
日付: [申請日]

■ 申請内容

Microsoft Entra ID にアプリケーションを登録し、管理者の同意をお願いいたします。

- アプリ名: Dot-connect Mail Analyzer
- 用途: メール送受信パターンの可視化による組織コミュニケーション分析
- 必要な権限:
  - Mail.Read（メールの読み取り — 読み取り専用）
  - User.Read（サインインとプロフィールの読み取り）

■ セキュリティに関する補足

- メール本文は一切収集しません（メタデータのみ: 日時、送信者、受信者、件名）
- すべてのデータ処理は社内サーバーで完結します
- 外部サービスへのデータ送信はありません
- ユーザーは自身のメールボックスのみアクセス可能です（他人のメールは読めません）

■ アカウントの種類

「この組織ディレクトリのみに含まれるアカウント」（シングルテナント）

ご確認のほど、よろしくお願いいたします。
```

### English Version

```
Subject: Microsoft Entra ID App Registration Request — Dot-connect Email Network Visualization

Requester: [Your Name]
Date: [Request Date]

■ Request Details

I would like to register an application in Microsoft Entra ID and request admin consent.

- App name: Dot-connect Mail Analyzer
- Purpose: Organizational communication analysis through email network visualization
- Required permissions:
  - Mail.Read (Read user mail — read-only)
  - User.Read (Sign in and read user profile)

■ Security Notes

- Email body content is NEVER collected (metadata only: date, sender, recipients, subject)
- All data processing occurs on internal servers
- No data is sent to external services
- Users can only access their own mailbox (no access to other users' mail)

■ Account Type

"Accounts in this organizational directory only" (single tenant)

Thank you for your consideration.
```

---

## 5. トラブルシューティング

### サインインできない

| 症状 | 対処 |
|------|------|
| `AADSTS65001: The user or administrator has not consented` | 管理者の同意が未付与。管理者に依頼してください |
| `AADSTS700016: Application not found` | `client_id` が正しいか確認してください |
| `AADSTS90002: Tenant not found` | `tenant_id` が正しいか確認してください |
| リダイレクト後に白い画面 | `redirect_uri` がアプリ登録と `config.yaml` で一致しているか確認 |

### フォルダが取得できない

- `Mail.Read` 権限が付与されているか確認
- 管理者の同意が付与されているか確認
- サインアウトして再度サインインを試行

### トークンの有効期限

- アクセストークンは通常1時間で期限切れ
- リフレッシュトークンにより自動更新されます
- 長期間使用しない場合はリフレッシュトークンも期限切れになるため、再サインインが必要です
- トークンキャッシュは `~/.dot-connect/token_cache.bin` に保存されます

### キャッシュのクリア

問題が解決しない場合:
1. Web UI の「サインアウト」ボタンをクリック
2. または `~/.dot-connect/token_cache.bin` を手動で削除
