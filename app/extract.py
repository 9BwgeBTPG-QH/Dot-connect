"""Outlook COM 操作の Web UI 用ラッパー.

各関数は COM を初期化/解放するためスレッドセーフ。
FastAPI の同期エンドポイントから呼び出すことを想定。
"""

import logging

import pandas as pd

log = logging.getLogger(__name__)


def get_outlook_folders() -> list[dict]:
    """Outlook に接続してメールフォルダ一覧を返す."""
    import pythoncom
    import win32com.client
    from extract import list_mail_folders

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        all_folders = []
        for i in range(1, namespace.Folders.Count + 1):
            store = namespace.Folders.Item(i)
            list_mail_folders(store, results=all_folders)

        return [{"index": idx, "path": path} for idx, (path, _) in enumerate(all_folders)]
    finally:
        pythoncom.CoUninitialize()


def run_extraction(
    folder_paths: list[str],
    start_date: str,
    end_date: str,
    config: dict,
) -> pd.DataFrame:
    """指定フォルダからメールを抽出し DataFrame を返す."""
    import pythoncom
    import win32com.client
    from extract import extract_emails, list_mail_folders

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # 全フォルダを列挙してパスで照合
        all_folders = []
        for i in range(1, namespace.Folders.Count + 1):
            store = namespace.Folders.Item(i)
            list_mail_folders(store, results=all_folders)

        path_to_folder = {path: folder for path, folder in all_folders}
        selected = []
        for fp in folder_paths:
            if fp in path_to_folder:
                selected.append(path_to_folder[fp])
            else:
                log.warning("フォルダが見つかりません: %s", fp)

        if not selected:
            raise ValueError("選択されたフォルダが見つかりません")

        log.info("抽出開始: %d フォルダ, 期間 %s ~ %s", len(selected), start_date, end_date)
        records = extract_emails(selected, start_date, end_date, config)

        if not records:
            raise ValueError("対象期間にメールがありません")

        log.info("抽出完了: %d 件", len(records))
        return pd.DataFrame(records)
    finally:
        pythoncom.CoUninitialize()
