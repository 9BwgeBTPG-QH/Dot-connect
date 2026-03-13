"""Microsoft Graph API メール抽出モジュール.

COM extract.py と同じ DataFrame 形式 [date, from_name, from_email, to, cc, subject] を返す。
"""

import logging
import time
from typing import Optional

import pandas as pd
import requests

# extract.py からフィルタ関数を再利用
from extract import apply_alias, build_exclude_set, is_excluded

log = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_PAGE_SIZE = 250
_MAX_RETRIES = 3


# ---------------------------------------------------------------------------
# Folder listing
# ---------------------------------------------------------------------------

def get_graph_folders(access_token: str) -> list[dict]:
    """Graph API でメールフォルダ一覧を取得.

    Returns:
        [{index, path, id}, ...] — COM の get_outlook_folders() と同形式。
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    folders: list[dict] = []

    def _recurse(parent_id: Optional[str], prefix: str) -> None:
        if parent_id is None:
            url = f"{GRAPH_BASE}/me/mailFolders"
        else:
            url = f"{GRAPH_BASE}/me/mailFolders/{parent_id}/childFolders"

        while url:
            resp = _request_with_retry("GET", url, headers=headers)
            if not resp.ok:
                _raise_graph_error(resp)
            data = resp.json()
            for item in data.get("value", []):
                path = f"{prefix}{item['displayName']}"
                folders.append({
                    "index": len(folders),
                    "path": path,
                    "id": item["id"],
                })
                # 子フォルダを再帰展開
                if item.get("childFolderCount", 0) > 0:
                    _recurse(item["id"], f"{path}/")
            url = data.get("@odata.nextLink")

    _recurse(None, "")
    return folders


# ---------------------------------------------------------------------------
# Mail extraction
# ---------------------------------------------------------------------------

def run_graph_extraction(
    access_token: str,
    folder_ids: list[str],
    start_date: str,
    end_date: str,
    config: dict,
) -> pd.DataFrame:
    """Graph API でメールを抽出し、COM と同じ DataFrame 形式で返す.

    Args:
        access_token: 有効な Graph API アクセストークン
        folder_ids: 抽出対象フォルダ ID リスト
        start_date: 開始日 (YYYY-MM-DD)
        end_date: 終了日 (YYYY-MM-DD)
        config: extract.py 互換の設定 dict (exclude_addresses, exclude_patterns, alias_map)

    Returns:
        DataFrame[date, from_name, from_email, to, cc, subject]
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    exclude_addrs, exclude_patterns = build_exclude_set(config)
    alias_map = config.get("alias_map") or {}

    records: list[dict] = []

    for folder_id in folder_ids:
        log.info("Graph API: フォルダ %s を処理中", folder_id)
        url = (
            f"{GRAPH_BASE}/me/mailFolders/{folder_id}/messages"
            f"?$filter=receivedDateTime ge {start_date}T00:00:00Z"
            f" and receivedDateTime le {end_date}T23:59:59Z"
            f"&$select=receivedDateTime,subject,from,toRecipients,ccRecipients"
            f"&$top={_PAGE_SIZE}"
            f"&$orderby=receivedDateTime desc"
        )

        while url:
            resp = _request_with_retry("GET", url, headers=headers)
            if not resp.ok:
                _raise_graph_error(resp)
            data = resp.json()

            for msg in data.get("value", []):
                record = _parse_message(msg, exclude_addrs, exclude_patterns, alias_map)
                if record is not None:
                    records.append(record)

            url = data.get("@odata.nextLink")

    log.info("Graph API: 合計 %d 件のメールを抽出", len(records))

    if not records:
        return pd.DataFrame(columns=["date", "from_name", "from_email", "to", "cc", "subject"])

    return pd.DataFrame(records)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _raise_graph_error(resp: requests.Response) -> None:
    """Graph API エラーレスポンスから詳細メッセージを抽出して例外を送出."""
    try:
        err = resp.json().get("error", {})
        code = err.get("code", resp.status_code)
        msg = err.get("message", resp.text)
    except Exception:
        code = resp.status_code
        msg = resp.text

    if resp.status_code == 404:
        raise RuntimeError(
            f"メールボックスが見つかりません ({code})。"
            "このアカウントに Exchange Online ライセンスが割り当てられているか確認してください。"
        )

    raise RuntimeError(f"Graph API エラー ({code}): {msg}")


def _parse_message(
    msg: dict,
    exclude_addrs: set,
    exclude_patterns: list,
    alias_map: dict,
) -> Optional[dict]:
    """Graph API メッセージ JSON → レコード dict に変換."""
    # 送信者
    from_data = msg.get("from", {}).get("emailAddress", {})
    from_email = (from_data.get("address") or "").lower()
    from_name = from_data.get("name") or ""

    if not from_email:
        return None

    from_email = apply_alias(from_email, alias_map)
    if is_excluded(from_email, exclude_addrs, exclude_patterns):
        return None

    # 日時
    received = msg.get("receivedDateTime", "")
    if received:
        # ISO 8601 → "YYYY-MM-DD HH:MM:SS"
        received = received.replace("T", " ")[:19]

    # 件名
    subject = msg.get("subject") or ""

    # To 受信者
    to_list = _format_recipients(
        msg.get("toRecipients", []), exclude_addrs, exclude_patterns, alias_map
    )

    # CC 受信者
    cc_list = _format_recipients(
        msg.get("ccRecipients", []), exclude_addrs, exclude_patterns, alias_map
    )

    return {
        "date": received,
        "from_name": from_name,
        "from_email": from_email,
        "to": "; ".join(to_list),
        "cc": "; ".join(cc_list),
        "subject": subject,
    }


def _format_recipients(
    recipients: list[dict],
    exclude_addrs: set,
    exclude_patterns: list,
    alias_map: dict,
) -> list[str]:
    """Graph API recipients リスト → "Name <email>; ..." 形式のリスト."""
    result = []
    for recip in recipients:
        addr_data = recip.get("emailAddress", {})
        email = (addr_data.get("address") or "").lower()
        name = addr_data.get("name") or ""
        if not email:
            continue
        email = apply_alias(email, alias_map)
        if is_excluded(email, exclude_addrs, exclude_patterns):
            continue
        if name:
            result.append(f"{name} <{email}>")
        else:
            result.append(email)
    return result


def _request_with_retry(method: str, url: str, **kwargs) -> requests.Response:
    """HTTP リクエスト + 429 リトライ (最大3回)."""
    for attempt in range(_MAX_RETRIES):
        resp = requests.request(method, url, timeout=30, **kwargs)
        if resp.status_code == 429:
            retry_after = int(resp.headers.get("Retry-After", 5))
            log.warning(
                "Graph API rate limited (429). %d秒後にリトライ (attempt %d/%d)",
                retry_after, attempt + 1, _MAX_RETRIES,
            )
            time.sleep(retry_after)
            continue
        return resp
    return resp  # Return last response even if still 429
