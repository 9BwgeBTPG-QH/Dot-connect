"""Outlook メール抽出 & サーバーアップロード (自己完結型スクリプト).

ローカルPCで実行し、Outlook COM からメールを抽出して
Dot-connect サーバーにアップロードする。

依存: pywin32 のみ (stdlib + win32com.client)
"""

import argparse
import csv
import io
import json
import re
import sys
import webbrowser
from datetime import datetime
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


# ---------------------------------------------------------------------------
# Outlook COM helpers (extract.py から移植, yaml/tqdm 依存を除去)
# ---------------------------------------------------------------------------

def connect_outlook():
    """Outlook MAPI 名前空間に接続."""
    try:
        import win32com.client
    except ImportError:
        print("[ERROR] pywin32 が必要です。")
        print("  pip install pywin32")
        sys.exit(1)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return namespace
    except Exception as e:
        print(f"[ERROR] Outlook に接続できません: {e}")
        sys.exit(1)


def list_mail_folders(folder, prefix="", results=None):
    """メールフォルダを再帰的に列挙 (DefaultItemType == 0 のみ)."""
    if results is None:
        results = []
    try:
        if folder.DefaultItemType == 0:
            results.append((f"{prefix}{folder.Name}", folder))
    except Exception:
        pass
    try:
        for i in range(1, folder.Folders.Count + 1):
            subfolder = folder.Folders.Item(i)
            list_mail_folders(subfolder, prefix=f"{prefix}{folder.Name}/", results=results)
    except Exception:
        pass
    return results


def choose_folder(namespace):
    """対話的にメールフォルダを選択."""
    all_folders = []
    for i in range(1, namespace.Folders.Count + 1):
        store = namespace.Folders.Item(i)
        list_mail_folders(store, results=all_folders)

    if not all_folders:
        print("[ERROR] メールフォルダが見つかりません。")
        sys.exit(1)

    print("\n--- メールフォルダ一覧 ---")
    for idx, (path, _) in enumerate(all_folders, 1):
        print(f"  {idx:3d}: {path}")
    print()

    while True:
        raw = input("フォルダ番号を入力 (カンマ区切りで複数可, 例: 1,3,5): ").strip()
        if not raw:
            continue
        try:
            indices = [int(x.strip()) for x in raw.split(",")]
            selected = []
            for idx in indices:
                if 1 <= idx <= len(all_folders):
                    selected.append(all_folders[idx - 1])
                else:
                    print(f"  番号 {idx} は範囲外です。")
            if selected:
                for path, _ in selected:
                    print(f"  -> {path}")
                return [folder for _, folder in selected]
        except ValueError:
            print("  数字を入力してください。")


# ---------------------------------------------------------------------------
# Address resolution (extract.py から移植)
# ---------------------------------------------------------------------------

def resolve_address(recipient):
    """受信者オブジェクトからメールアドレスと表示名を解決."""
    display_name = ""
    try:
        display_name = recipient.Name or ""
    except Exception:
        pass

    try:
        entry = recipient.AddressEntry
        if entry.AddressEntryUserType in (0, 5):
            exuser = entry.GetExchangeUser()
            if exuser and exuser.PrimarySmtpAddress:
                return exuser.PrimarySmtpAddress.lower(), display_name
    except Exception:
        pass

    try:
        addr = recipient.AddressEntry.Address
        if addr and "@" in addr:
            return addr.lower(), display_name
    except Exception:
        pass

    try:
        PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        smtp = recipient.PropertyAccessor.GetProperty(PR_SMTP)
        if smtp and "@" in smtp:
            return smtp.lower(), display_name
    except Exception:
        pass

    if display_name:
        sanitized = re.sub(r"[^\w.-]", "_", display_name)
        return f"{sanitized}@unresolved.local".lower(), display_name

    return "unknown@unresolved.local", "Unknown"


def resolve_sender(mail_item):
    """送信者のメールアドレスと表示名を解決."""
    display_name = ""
    try:
        display_name = mail_item.SenderName or ""
    except Exception:
        pass

    try:
        if mail_item.SenderEmailType == "EX":
            sender_entry = mail_item.Sender
            if sender_entry:
                exuser = sender_entry.GetExchangeUser()
                if exuser and exuser.PrimarySmtpAddress:
                    return exuser.PrimarySmtpAddress.lower(), display_name
    except Exception:
        pass

    try:
        addr = mail_item.SenderEmailAddress
        if addr and "@" in addr:
            return addr.lower(), display_name
    except Exception:
        pass

    try:
        PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
        smtp = mail_item.PropertyAccessor.GetProperty(PR_SENDER_SMTP)
        if smtp and "@" in smtp:
            return smtp.lower(), display_name
    except Exception:
        pass

    if display_name:
        sanitized = re.sub(r"[^\w.-]", "_", display_name)
        return f"{sanitized}@unresolved.local".lower(), display_name

    return "unknown@unresolved.local", "Unknown"


# ---------------------------------------------------------------------------
# Extraction (extract.py から移植, tqdm を除去)
# ---------------------------------------------------------------------------

def extract_emails(folders, start_date, end_date):
    """選択フォルダからメールメタデータを抽出."""
    restriction = (
        f"[ReceivedTime] >= '{start_date} 00:00' "
        f"AND [ReceivedTime] <= '{end_date} 23:59'"
    )

    records = []
    for folder in folders:
        print(f"\nフォルダ処理中: {folder.Name}")
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            filtered = items.Restrict(restriction)
            count = filtered.Count
        except Exception as e:
            print(f"  フォルダ '{folder.Name}' のアクセスに失敗: {e}")
            continue

        if count == 0:
            print("  対象メールなし")
            continue

        print(f"  {count} 件のメールを処理中...")

        for i in range(1, count + 1):
            if i % 50 == 0 or i == count:
                print(f"\r  {i}/{count} 件処理済み", end="", flush=True)

            try:
                mail = filtered.Item(i)
                if mail.Class != 43:  # olMail = 43
                    continue
            except Exception:
                continue

            from_email, from_name = resolve_sender(mail)

            try:
                received = mail.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                received = ""
            try:
                subject = mail.Subject or ""
            except Exception:
                subject = ""

            to_list = []
            try:
                for j in range(1, mail.Recipients.Count + 1):
                    recip = mail.Recipients.Item(j)
                    if recip.Type == 1:  # olTo
                        email, name = resolve_address(recip)
                        to_list.append(f"{name} <{email}>")
            except Exception:
                pass

            cc_list = []
            try:
                for j in range(1, mail.Recipients.Count + 1):
                    recip = mail.Recipients.Item(j)
                    if recip.Type == 2:  # olCC
                        email, name = resolve_address(recip)
                        cc_list.append(f"{name} <{email}>")
            except Exception:
                pass

            records.append({
                "date": received,
                "from_name": from_name,
                "from_email": from_email,
                "to": "; ".join(to_list),
                "cc": "; ".join(cc_list),
                "subject": subject,
            })

        print()  # newline after progress

    return records


# ---------------------------------------------------------------------------
# Upload
# ---------------------------------------------------------------------------

def build_csv_bytes(records):
    """レコードを CSV バイト列 (UTF-8 BOM) に変換."""
    output = io.StringIO()
    writer = csv.DictWriter(
        output,
        fieldnames=["date", "from_name", "from_email", "to", "cc", "subject"],
        quoting=csv.QUOTE_ALL,
    )
    writer.writeheader()
    writer.writerows(records)
    return output.getvalue().encode("utf-8-sig")


def upload_csv(csv_bytes, server_url, config_params):
    """CSV をサーバーにアップロードし、結果URLを返す."""
    boundary = "----DotConnectBoundary8192"

    parts = []

    # CSV file part
    parts.append(f"--{boundary}\r\n".encode())
    parts.append(b'Content-Disposition: form-data; name="file"; filename="extract.csv"\r\n')
    parts.append(b"Content-Type: text/csv\r\n\r\n")
    parts.append(csv_bytes)
    parts.append(b"\r\n")

    # Config params
    for key, value in config_params.items():
        parts.append(f"--{boundary}\r\n".encode())
        parts.append(f'Content-Disposition: form-data; name="{key}"\r\n\r\n'.encode())
        parts.append(str(value).encode())
        parts.append(b"\r\n")

    parts.append(f"--{boundary}--\r\n".encode())

    body = b"".join(parts)
    url = f"{server_url.rstrip('/')}/api/upload-csv"

    req = Request(
        url,
        data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        method="POST",
    )

    try:
        with urlopen(req, timeout=120) as resp:
            result = json.loads(resp.read().decode())
            return result.get("result_url", "")
    except HTTPError as e:
        print(f"[ERROR] サーバーエラー: {e.code} {e.reason}")
        try:
            print(e.read().decode())
        except Exception:
            pass
        return ""
    except URLError as e:
        print(f"[ERROR] サーバーに接続できません: {e.reason}")
        return ""


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Outlook メール抽出 & Dot-connect サーバーアップロード"
    )
    parser.add_argument("--server_url", required=True, help="サーバーURL (例: http://yourserver:8000)")
    parser.add_argument("--start_date", required=True, help="開始日 (YYYY-MM-DD)")
    parser.add_argument("--end_date", required=True, help="終了日 (YYYY-MM-DD)")
    parser.add_argument("--company_domains", default="", help="社内ドメイン (カンマ区切り)")
    parser.add_argument("--cc_key_person_threshold", type=float, default=0.30)
    parser.add_argument("--min_edge_weight", type=int, default=1)
    parser.add_argument("--hub_degree_weight", type=float, default=0.5)
    parser.add_argument("--hub_betweenness_weight", type=float, default=0.5)
    args = parser.parse_args()

    # Validate dates
    for label, val in [("開始日", args.start_date), ("終了日", args.end_date)]:
        try:
            datetime.strptime(val, "%Y-%m-%d")
        except ValueError:
            print(f"[ERROR] {label} は YYYY-MM-DD 形式で入力してください: {val}")
            sys.exit(1)

    print("========================================")
    print("  Dot-connect - メール抽出ツール")
    print(f"  期間: {args.start_date} ~ {args.end_date}")
    print(f"  サーバー: {args.server_url}")
    print("========================================")

    # 1. Outlook に接続
    print("\nOutlook に接続中...")
    namespace = connect_outlook()

    # 2. フォルダ選択
    folders = choose_folder(namespace)

    # 3. メール抽出
    print(f"\n抽出開始: {args.start_date} ~ {args.end_date}")
    records = extract_emails(folders, args.start_date, args.end_date)

    if not records:
        print("\n対象期間にメールがありません。")
        sys.exit(0)

    print(f"\n{len(records)} 件のメールを抽出しました。")

    # 4. サーバーにアップロード
    print("サーバーにアップロード中...")
    csv_bytes = build_csv_bytes(records)

    config_params = {
        "company_domains": args.company_domains,
        "cc_key_person_threshold": args.cc_key_person_threshold,
        "min_edge_weight": args.min_edge_weight,
        "hub_degree_weight": args.hub_degree_weight,
        "hub_betweenness_weight": args.hub_betweenness_weight,
    }

    result_url = upload_csv(csv_bytes, args.server_url, config_params)

    if result_url:
        full_url = f"{args.server_url.rstrip('/')}{result_url}"
        print(f"\n分析完了！ ブラウザで結果を表示します。")
        print(f"  URL: {full_url}")
        webbrowser.open(full_url)
    else:
        print("\n[ERROR] アップロードに失敗しました。")
        sys.exit(1)


if __name__ == "__main__":
    main()
