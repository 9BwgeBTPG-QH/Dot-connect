"""Outlook メールメタデータ抽出 → CSV出力.

Usage:
    python extract.py --start 2024-01-01 --end 2024-12-31
"""

import argparse
import csv
import logging
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import yaml
from tqdm import tqdm

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def load_config(path: str = "config.yaml") -> dict:
    config_path = Path(__file__).parent / path
    if not config_path.exists():
        log.warning("config.yaml が見つかりません。デフォルト設定を使用します。")
        return {
            "company_domains": [],
            "exclude_addresses": [],
            "exclude_patterns": [],
            "alias_map": {},
            "thresholds": {
                "cc_key_person_threshold": 0.30,
                "min_edge_weight": 1,
            },
        }
    with open(config_path, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


# ---------------------------------------------------------------------------
# Outlook COM helpers
# ---------------------------------------------------------------------------

def connect_outlook():
    """Outlook MAPI 名前空間に接続."""
    try:
        import win32com.client
    except ImportError:
        log.error("pywin32 が必要です。 pip install pywin32 を実行してください。")
        sys.exit(1)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return namespace
    except Exception as e:
        log.error("Outlook に接続できません: %s", e)
        sys.exit(1)


def list_mail_folders(folder, prefix="", results=None):
    """メールフォルダを再帰的に列挙.

    DefaultItemType == 0 のフォルダのみ（メールアイテム）.
    """
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


def choose_folder(namespace) -> list:
    """対話的にメールフォルダを選択."""
    all_folders = []
    for i in range(1, namespace.Folders.Count + 1):
        store = namespace.Folders.Item(i)
        list_mail_folders(store, results=all_folders)

    if not all_folders:
        log.error("メールフォルダが見つかりません。")
        sys.exit(1)

    print("\n--- メールフォルダ一覧 ---")
    for idx, (path, _) in enumerate(all_folders, 1):
        print(f"  {idx:3d}: {path}")
    print()

    while True:
        raw = input("フォルダ番号を入力（カンマ区切りで複数可, 例: 1,3,5）: ").strip()
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
                    print(f"  ✓ {path}")
                return [folder for _, folder in selected]
        except ValueError:
            print("  数字を入力してください。")


# ---------------------------------------------------------------------------
# Address resolution
# ---------------------------------------------------------------------------

def resolve_address(recipient) -> tuple[str, str]:
    """受信者オブジェクトからメールアドレスと表示名を解決.

    Exchange DN (/o=Org/...) を SMTP アドレスに変換する多段フォールバック。

    Returns:
        (email, display_name)
    """
    display_name = ""
    try:
        display_name = recipient.Name or ""
    except Exception:
        pass

    # 1) ExchangeUser → PrimarySmtpAddress
    try:
        entry = recipient.AddressEntry
        if entry.AddressEntryUserType in (0, 5):  # olExchangeUserAddressEntry, olExchangeRemoteUserAddressEntry
            exuser = entry.GetExchangeUser()
            if exuser and exuser.PrimarySmtpAddress:
                return exuser.PrimarySmtpAddress.lower(), display_name
    except Exception:
        pass

    # 2) AddressEntry.Address (SMTP の場合そのまま使える)
    try:
        addr = recipient.AddressEntry.Address
        if addr and "@" in addr:
            return addr.lower(), display_name
    except Exception:
        pass

    # 3) PropertyAccessor で SMTP アドレスを取得
    try:
        PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        smtp = recipient.PropertyAccessor.GetProperty(PR_SMTP)
        if smtp and "@" in smtp:
            return smtp.lower(), display_name
    except Exception:
        pass

    # 4) 最終フォールバック: 表示名からダミーアドレス生成
    if display_name:
        sanitized = re.sub(r"[^\w.-]", "_", display_name)
        fallback = f"{sanitized}@unresolved.local"
        log.debug("アドレス解決不能: %s → %s", display_name, fallback)
        return fallback.lower(), display_name

    return "unknown@unresolved.local", "Unknown"


def resolve_sender(mail_item) -> tuple[str, str]:
    """送信者のメールアドレスと表示名を解決."""
    display_name = ""
    try:
        display_name = mail_item.SenderName or ""
    except Exception:
        pass

    # 1) SenderEmailType が EX の場合 → Exchange解決
    try:
        if mail_item.SenderEmailType == "EX":
            sender_entry = mail_item.Sender
            if sender_entry:
                exuser = sender_entry.GetExchangeUser()
                if exuser and exuser.PrimarySmtpAddress:
                    return exuser.PrimarySmtpAddress.lower(), display_name
    except Exception:
        pass

    # 2) SenderEmailAddress (SMTPならそのまま)
    try:
        addr = mail_item.SenderEmailAddress
        if addr and "@" in addr:
            return addr.lower(), display_name
    except Exception:
        pass

    # 3) PropertyAccessor
    try:
        PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
        smtp = mail_item.PropertyAccessor.GetProperty(PR_SENDER_SMTP)
        if smtp and "@" in smtp:
            return smtp.lower(), display_name
    except Exception:
        pass

    # 4) フォールバック
    if display_name:
        sanitized = re.sub(r"[^\w.-]", "_", display_name)
        return f"{sanitized}@unresolved.local".lower(), display_name

    return "unknown@unresolved.local", "Unknown"


# ---------------------------------------------------------------------------
# Filtering
# ---------------------------------------------------------------------------

def build_exclude_set(config: dict) -> tuple[set, list]:
    """除外アドレスセットと除外パターンリストを構築."""
    addresses = {a.lower() for a in config.get("exclude_addresses", [])}
    patterns = []
    for p in config.get("exclude_patterns", []):
        try:
            patterns.append(re.compile(p, re.IGNORECASE))
        except re.error as e:
            log.warning("無効な除外パターン '%s': %s", p, e)
    return addresses, patterns


def is_excluded(email: str, exclude_addrs: set, exclude_patterns: list) -> bool:
    if email in exclude_addrs:
        return True
    return any(p.search(email) for p in exclude_patterns)


def apply_alias(email: str, alias_map: dict) -> str:
    """エイリアスを正規アドレスに変換."""
    if not alias_map:
        return email
    for canonical, aliases in alias_map.items():
        if email == canonical.lower():
            return canonical.lower()
        if isinstance(aliases, list) and email in [a.lower() for a in aliases]:
            return canonical.lower()
    return email


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------

def extract_emails(folders, start_date: str, end_date: str, config: dict) -> list[dict]:
    """選択フォルダからメールメタデータを抽出."""
    # DASL フィルタ（Outlook側で日付絞り込み）
    restriction = (
        f"[ReceivedTime] >= '{start_date} 00:00' "
        f"AND [ReceivedTime] <= '{end_date} 23:59'"
    )

    exclude_addrs, exclude_patterns = build_exclude_set(config)
    alias_map = config.get("alias_map") or {}

    records = []
    for folder in folders:
        log.info("フォルダ処理中: %s", folder.Name)
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            filtered = items.Restrict(restriction)
            count = filtered.Count
        except Exception as e:
            log.warning("フォルダ '%s' のアクセスに失敗: %s", folder.Name, e)
            continue

        if count == 0:
            log.info("  対象メールなし")
            continue

        log.info("  %d 件のメールを処理", count)

        for i in tqdm(range(1, count + 1), desc=f"  {folder.Name}", unit="mail"):
            try:
                mail = filtered.Item(i)
                # MailItem 以外 (会議招集等) はスキップ
                if mail.Class != 43:  # olMail = 43
                    continue
            except Exception:
                continue

            # 送信者
            from_email, from_name = resolve_sender(mail)
            from_email = apply_alias(from_email, alias_map)
            if is_excluded(from_email, exclude_addrs, exclude_patterns):
                continue

            # 日時・件名
            try:
                received = mail.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                received = ""
            try:
                subject = mail.Subject or ""
            except Exception:
                subject = ""

            # To 受信者
            to_list = []
            try:
                for j in range(1, mail.Recipients.Count + 1):
                    recip = mail.Recipients.Item(j)
                    if recip.Type == 1:  # olTo
                        email, name = resolve_address(recip)
                        email = apply_alias(email, alias_map)
                        if not is_excluded(email, exclude_addrs, exclude_patterns):
                            to_list.append(f"{name} <{email}>")
            except Exception:
                pass

            # CC 受信者
            cc_list = []
            try:
                for j in range(1, mail.Recipients.Count + 1):
                    recip = mail.Recipients.Item(j)
                    if recip.Type == 2:  # olCC
                        email, name = resolve_address(recip)
                        email = apply_alias(email, alias_map)
                        if not is_excluded(email, exclude_addrs, exclude_patterns):
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

    return records


# ---------------------------------------------------------------------------
# CSV output
# ---------------------------------------------------------------------------

def save_csv(records: list[dict], output_dir: str = "output"):
    """CSVファイルに保存 (UTF-8 BOM)."""
    out = Path(__file__).parent / output_dir
    out.mkdir(exist_ok=True)

    today = datetime.now().strftime("%Y%m%d")
    filepath = out / f"emails_{today}.csv"

    with open(filepath, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["date", "from_name", "from_email", "to", "cc", "subject"],
            quoting=csv.QUOTE_ALL,
        )
        writer.writeheader()
        writer.writerows(records)

    log.info("CSV出力完了: %s (%d件)", filepath, len(records))
    return str(filepath)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Outlook メールメタデータ抽出ツール"
    )
    parser.add_argument(
        "--start", required=True, help="開始日 (YYYY-MM-DD)"
    )
    parser.add_argument(
        "--end", required=True, help="終了日 (YYYY-MM-DD)"
    )
    parser.add_argument(
        "--output", default="output", help="出力ディレクトリ (default: output)"
    )
    args = parser.parse_args()

    # 日付バリデーション
    try:
        datetime.strptime(args.start, "%Y-%m-%d")
        datetime.strptime(args.end, "%Y-%m-%d")
    except ValueError:
        log.error("日付は YYYY-MM-DD 形式で入力してください。")
        sys.exit(1)

    config = load_config()
    namespace = connect_outlook()
    folders = choose_folder(namespace)
    records = extract_emails(folders, args.start, args.end, config)

    if not records:
        log.warning("抽出されたメールがありません。")
        sys.exit(0)

    csv_path = save_csv(records, args.output)
    print(f"\n完了！ {len(records)} 件のメールを抽出しました。")
    print(f"CSV: {csv_path}")


if __name__ == "__main__":
    main()
