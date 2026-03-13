"""Unit tests for app.graph_extract — Graph API mail extraction."""

import json
from pathlib import Path
from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

# Fixtures path
FIXTURES = Path(__file__).parent / "fixtures" / "graph_api_responses"


def _load_fixture(name):
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def _mock_response(json_data, status_code=200):
    resp = MagicMock()
    resp.status_code = status_code
    resp.json.return_value = json_data
    resp.raise_for_status.return_value = None
    return resp


def _default_config():
    return {
        "exclude_addresses": [],
        "exclude_patterns": [],
        "alias_map": {},
    }


# ======================================================================
# get_graph_folders
# ======================================================================

class TestGetGraphFolders:
    """Tests for get_graph_folders()."""

    @patch("app.graph_extract._request_with_retry")
    def test_basic_folder_listing(self, mock_req):
        from app.graph_extract import get_graph_folders

        folders_data = _load_fixture("mail_folders.json")
        child_data = _load_fixture("child_folders.json")

        # First call: top-level folders, Second: child folders of 受信トレイ
        mock_req.side_effect = [
            _mock_response(folders_data),
            _mock_response(child_data),
        ]

        folders = get_graph_folders("test-token")
        assert len(folders) == 3  # 受信トレイ + 送信済み + プロジェクトA
        assert folders[0]["path"] == "受信トレイ"
        assert folders[1]["path"] == "受信トレイ/プロジェクトA"
        assert folders[2]["path"] == "送信済みアイテム"

    @patch("app.graph_extract._request_with_retry")
    def test_folder_structure(self, mock_req):
        from app.graph_extract import get_graph_folders

        mock_req.return_value = _mock_response({
            "value": [{
                "id": "folder-1",
                "displayName": "TestFolder",
                "childFolderCount": 0,
            }]
        })

        folders = get_graph_folders("test-token")
        assert len(folders) == 1
        assert folders[0]["index"] == 0
        assert folders[0]["path"] == "TestFolder"
        assert folders[0]["id"] == "folder-1"

    @patch("app.graph_extract._request_with_retry")
    def test_pagination(self, mock_req):
        from app.graph_extract import get_graph_folders

        page1 = {
            "value": [{"id": "f1", "displayName": "Folder1", "childFolderCount": 0}],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/mailFolders?$skip=1",
        }
        page2 = {
            "value": [{"id": "f2", "displayName": "Folder2", "childFolderCount": 0}],
        }
        mock_req.side_effect = [_mock_response(page1), _mock_response(page2)]

        folders = get_graph_folders("test-token")
        assert len(folders) == 2
        assert folders[0]["path"] == "Folder1"
        assert folders[1]["path"] == "Folder2"

    @patch("app.graph_extract._request_with_retry")
    def test_empty_mailbox(self, mock_req):
        from app.graph_extract import get_graph_folders

        mock_req.return_value = _mock_response({"value": []})
        folders = get_graph_folders("test-token")
        assert folders == []

    @patch("app.graph_extract._request_with_retry")
    def test_auth_header(self, mock_req):
        from app.graph_extract import get_graph_folders

        mock_req.return_value = _mock_response({"value": []})
        get_graph_folders("my-access-token")

        call_kwargs = mock_req.call_args
        assert call_kwargs[1]["headers"]["Authorization"] == "Bearer my-access-token"


# ======================================================================
# run_graph_extraction
# ======================================================================

class TestRunGraphExtraction:
    """Tests for run_graph_extraction()."""

    @patch("app.graph_extract._request_with_retry")
    def test_basic_extraction(self, mock_req):
        from app.graph_extract import run_graph_extraction

        messages_data = _load_fixture("messages.json")
        mock_req.return_value = _mock_response(messages_data)

        df = run_graph_extraction(
            "test-token", ["folder-1"], "2024-06-01", "2024-06-30", _default_config()
        )

        assert isinstance(df, pd.DataFrame)
        assert len(df) == 3
        assert list(df.columns) == ["date", "from_name", "from_email", "to", "cc", "subject"]

    @patch("app.graph_extract._request_with_retry")
    def test_date_format(self, mock_req):
        from app.graph_extract import run_graph_extraction

        messages_data = _load_fixture("messages.json")
        mock_req.return_value = _mock_response(messages_data)

        df = run_graph_extraction(
            "test-token", ["folder-1"], "2024-06-01", "2024-06-30", _default_config()
        )

        # ISO 8601 → "YYYY-MM-DD HH:MM:SS"
        assert df.iloc[0]["date"] == "2024-06-15 09:30:00"

    @patch("app.graph_extract._request_with_retry")
    def test_from_fields(self, mock_req):
        from app.graph_extract import run_graph_extraction

        messages_data = _load_fixture("messages.json")
        mock_req.return_value = _mock_response(messages_data)

        df = run_graph_extraction(
            "test-token", ["folder-1"], "2024-06-01", "2024-06-30", _default_config()
        )

        assert df.iloc[0]["from_name"] == "山田太郎"
        assert df.iloc[0]["from_email"] == "taro.yamada@example.co.jp"

    @patch("app.graph_extract._request_with_retry")
    def test_recipients_format(self, mock_req):
        from app.graph_extract import run_graph_extraction

        messages_data = _load_fixture("messages.json")
        mock_req.return_value = _mock_response(messages_data)

        df = run_graph_extraction(
            "test-token", ["folder-1"], "2024-06-01", "2024-06-30", _default_config()
        )

        # First message has 2 to recipients and 1 cc
        to_field = df.iloc[0]["to"]
        assert "hanako.suzuki@example.co.jp" in to_field
        assert "ichiro.sato@example.co.jp" in to_field
        assert ";" in to_field  # semicolon separator

        cc_field = df.iloc[0]["cc"]
        assert "jiro.tanaka@example.co.jp" in cc_field

    @patch("app.graph_extract._request_with_retry")
    def test_exclude_addresses(self, mock_req):
        from app.graph_extract import run_graph_extraction

        messages_data = _load_fixture("messages.json")
        mock_req.return_value = _mock_response(messages_data)

        config = {
            "exclude_addresses": ["partner@external.com"],
            "exclude_patterns": [],
            "alias_map": {},
        }

        df = run_graph_extraction(
            "test-token", ["folder-1"], "2024-06-01", "2024-06-30", config
        )

        # Third message (from partner@external.com) should be excluded
        assert len(df) == 2
        assert "partner@external.com" not in df["from_email"].values

    @patch("app.graph_extract._request_with_retry")
    def test_exclude_patterns(self, mock_req):
        from app.graph_extract import run_graph_extraction

        mock_req.return_value = _mock_response({
            "value": [{
                "receivedDateTime": "2024-06-15T09:00:00Z",
                "subject": "Auto reply",
                "from": {"emailAddress": {"name": "Noreply", "address": "noreply@example.com"}},
                "toRecipients": [{"emailAddress": {"name": "User", "address": "user@example.com"}}],
                "ccRecipients": [],
            }]
        })

        config = {
            "exclude_addresses": [],
            "exclude_patterns": ["^noreply@"],
            "alias_map": {},
        }

        df = run_graph_extraction(
            "test-token", ["folder-1"], "2024-06-01", "2024-06-30", config
        )
        assert len(df) == 0

    @patch("app.graph_extract._request_with_retry")
    def test_alias_mapping(self, mock_req):
        from app.graph_extract import run_graph_extraction

        mock_req.return_value = _mock_response({
            "value": [{
                "receivedDateTime": "2024-06-15T09:00:00Z",
                "subject": "Test",
                "from": {"emailAddress": {"name": "Taro", "address": "old-taro@example.com"}},
                "toRecipients": [{"emailAddress": {"name": "User", "address": "user@example.com"}}],
                "ccRecipients": [],
            }]
        })

        config = {
            "exclude_addresses": [],
            "exclude_patterns": [],
            "alias_map": {
                "taro@example.com": ["old-taro@example.com"],
            },
        }

        df = run_graph_extraction(
            "test-token", ["folder-1"], "2024-06-01", "2024-06-30", config
        )
        assert df.iloc[0]["from_email"] == "taro@example.com"

    @patch("app.graph_extract._request_with_retry")
    def test_empty_folder(self, mock_req):
        from app.graph_extract import run_graph_extraction

        mock_req.return_value = _mock_response({"value": []})

        df = run_graph_extraction(
            "test-token", ["folder-1"], "2024-06-01", "2024-06-30", _default_config()
        )

        assert isinstance(df, pd.DataFrame)
        assert len(df) == 0
        assert list(df.columns) == ["date", "from_name", "from_email", "to", "cc", "subject"]

    @patch("app.graph_extract._request_with_retry")
    def test_multiple_folders(self, mock_req):
        from app.graph_extract import run_graph_extraction

        msg1 = {
            "value": [{
                "receivedDateTime": "2024-06-15T09:00:00Z",
                "subject": "Mail 1",
                "from": {"emailAddress": {"name": "A", "address": "a@example.com"}},
                "toRecipients": [{"emailAddress": {"name": "B", "address": "b@example.com"}}],
                "ccRecipients": [],
            }]
        }
        msg2 = {
            "value": [{
                "receivedDateTime": "2024-06-16T10:00:00Z",
                "subject": "Mail 2",
                "from": {"emailAddress": {"name": "C", "address": "c@example.com"}},
                "toRecipients": [{"emailAddress": {"name": "D", "address": "d@example.com"}}],
                "ccRecipients": [],
            }]
        }
        mock_req.side_effect = [_mock_response(msg1), _mock_response(msg2)]

        df = run_graph_extraction(
            "test-token", ["folder-1", "folder-2"], "2024-06-01", "2024-06-30", _default_config()
        )
        assert len(df) == 2

    @patch("app.graph_extract._request_with_retry")
    def test_message_without_from(self, mock_req):
        from app.graph_extract import run_graph_extraction

        mock_req.return_value = _mock_response({
            "value": [{
                "receivedDateTime": "2024-06-15T09:00:00Z",
                "subject": "No sender",
                "from": {"emailAddress": {"name": "", "address": ""}},
                "toRecipients": [],
                "ccRecipients": [],
            }]
        })

        df = run_graph_extraction(
            "test-token", ["folder-1"], "2024-06-01", "2024-06-30", _default_config()
        )
        assert len(df) == 0

    @patch("app.graph_extract._request_with_retry")
    def test_pagination(self, mock_req):
        from app.graph_extract import run_graph_extraction

        page1 = {
            "value": [{
                "receivedDateTime": "2024-06-15T09:00:00Z",
                "subject": "Page1",
                "from": {"emailAddress": {"name": "A", "address": "a@example.com"}},
                "toRecipients": [{"emailAddress": {"name": "B", "address": "b@example.com"}}],
                "ccRecipients": [],
            }],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/mailFolders/f1/messages?$skip=1",
        }
        page2 = {
            "value": [{
                "receivedDateTime": "2024-06-16T10:00:00Z",
                "subject": "Page2",
                "from": {"emailAddress": {"name": "C", "address": "c@example.com"}},
                "toRecipients": [{"emailAddress": {"name": "D", "address": "d@example.com"}}],
                "ccRecipients": [],
            }],
        }
        mock_req.side_effect = [_mock_response(page1), _mock_response(page2)]

        df = run_graph_extraction(
            "test-token", ["folder-1"], "2024-06-01", "2024-06-30", _default_config()
        )
        assert len(df) == 2


# ======================================================================
# _request_with_retry
# ======================================================================

class TestRequestWithRetry:
    """Tests for rate limit retry logic."""

    @patch("app.graph_extract.requests.request")
    @patch("app.graph_extract.time.sleep")
    def test_retry_on_429(self, mock_sleep, mock_request):
        from app.graph_extract import _request_with_retry

        # First call: 429, second call: 200
        resp_429 = MagicMock()
        resp_429.status_code = 429
        resp_429.headers = {"Retry-After": "2"}

        resp_200 = MagicMock()
        resp_200.status_code = 200

        mock_request.side_effect = [resp_429, resp_200]

        result = _request_with_retry("GET", "https://graph.microsoft.com/test")
        assert result.status_code == 200
        mock_sleep.assert_called_once_with(2)

    @patch("app.graph_extract.requests.request")
    @patch("app.graph_extract.time.sleep")
    def test_max_retries_exceeded(self, mock_sleep, mock_request):
        from app.graph_extract import _request_with_retry

        resp_429 = MagicMock()
        resp_429.status_code = 429
        resp_429.headers = {"Retry-After": "1"}

        mock_request.return_value = resp_429

        result = _request_with_retry("GET", "https://graph.microsoft.com/test")
        assert result.status_code == 429
        assert mock_sleep.call_count == 3  # _MAX_RETRIES = 3
