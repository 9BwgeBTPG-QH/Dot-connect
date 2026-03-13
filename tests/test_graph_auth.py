"""Unit tests for app.graph_auth — MSAL OAuth2 authentication."""

import time
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from app.graph_auth import GraphAuth, _FLOW_TTL_SECONDS


# ======================================================================
# Helpers
# ======================================================================

def _make_auth(tmp_path):
    """Create a GraphAuth instance with mocked MSAL."""
    with patch("app.graph_auth.msal") as mock_msal:
        mock_app = MagicMock()
        mock_msal.PublicClientApplication.return_value = mock_app
        mock_msal.SerializableTokenCache.return_value = MagicMock(
            has_state_changed=False,
        )

        auth = GraphAuth(
            client_id="test-client-id",
            tenant_id="test-tenant-id",
            redirect_uri="http://localhost:8000/auth/callback",
            cache_path=str(tmp_path / "token_cache.bin"),
        )
        auth._app = mock_app
        return auth, mock_app


# ======================================================================
# Tests
# ======================================================================

class TestGraphAuth:
    """Tests for GraphAuth class."""

    def test_init_creates_instance(self, tmp_path):
        auth, _ = _make_auth(tmp_path)
        assert auth._client_id == "test-client-id"
        assert auth._tenant_id == "test-tenant-id"

    def test_get_auth_url(self, tmp_path):
        auth, mock_app = _make_auth(tmp_path)
        mock_app.initiate_auth_code_flow.return_value = {
            "auth_uri": "https://login.microsoftonline.com/authorize?...",
            "state": "test-state",
        }
        url = auth.get_auth_url("test-state")
        assert "login.microsoftonline.com" in url
        assert "test-state" in auth._pending_flows
        mock_app.initiate_auth_code_flow.assert_called_once()

    def test_acquire_token_success(self, tmp_path):
        auth, mock_app = _make_auth(tmp_path)
        # Setup pending flow
        auth._pending_flows["my-state"] = {
            "flow": {"state": "my-state"},
            "created_at": time.time(),
        }
        mock_app.acquire_token_by_auth_code_flow.return_value = {
            "access_token": "test-token-123",
        }
        result = auth.acquire_token_by_auth_code({"state": "my-state", "code": "auth-code"})
        assert result["access_token"] == "test-token-123"
        assert "my-state" not in auth._pending_flows

    def test_acquire_token_invalid_state(self, tmp_path):
        auth, _ = _make_auth(tmp_path)
        result = auth.acquire_token_by_auth_code({"state": "unknown", "code": "auth-code"})
        assert "error" in result
        assert result["error"] == "invalid_state"

    def test_get_access_token_from_cache(self, tmp_path):
        auth, mock_app = _make_auth(tmp_path)
        mock_app.get_accounts.return_value = [{"username": "user@example.com"}]
        mock_app.acquire_token_silent.return_value = {"access_token": "cached-token"}
        token = auth.get_access_token()
        assert token == "cached-token"

    def test_get_access_token_no_accounts(self, tmp_path):
        auth, mock_app = _make_auth(tmp_path)
        mock_app.get_accounts.return_value = []
        token = auth.get_access_token()
        assert token is None

    def test_get_access_token_silent_fails(self, tmp_path):
        auth, mock_app = _make_auth(tmp_path)
        mock_app.get_accounts.return_value = [{"username": "user@example.com"}]
        mock_app.acquire_token_silent.return_value = {"error": "interaction_required"}
        token = auth.get_access_token()
        assert token is None

    def test_is_authenticated_true(self, tmp_path):
        auth, mock_app = _make_auth(tmp_path)
        mock_app.get_accounts.return_value = [{"username": "user@example.com"}]
        mock_app.acquire_token_silent.return_value = {"access_token": "token"}
        assert auth.is_authenticated() is True

    def test_is_authenticated_false(self, tmp_path):
        auth, mock_app = _make_auth(tmp_path)
        mock_app.get_accounts.return_value = []
        assert auth.is_authenticated() is False

    def test_sign_out(self, tmp_path):
        auth, mock_app = _make_auth(tmp_path)
        account = {"username": "user@example.com"}
        mock_app.get_accounts.return_value = [account]
        auth.sign_out()
        mock_app.remove_account.assert_called_once_with(account)

    def test_expired_flows_cleanup(self, tmp_path):
        auth, mock_app = _make_auth(tmp_path)
        # Add an expired flow
        auth._pending_flows["old-state"] = {
            "flow": {},
            "created_at": time.time() - _FLOW_TTL_SECONDS - 100,
        }
        # Add a valid flow
        auth._pending_flows["new-state"] = {
            "flow": {},
            "created_at": time.time(),
        }
        mock_app.initiate_auth_code_flow.return_value = {
            "auth_uri": "https://login.microsoftonline.com/authorize",
            "state": "fresh-state",
        }
        auth.get_auth_url("fresh-state")
        assert "old-state" not in auth._pending_flows
        assert "new-state" in auth._pending_flows

    def test_cache_save_on_token_acquire(self, tmp_path):
        auth, mock_app = _make_auth(tmp_path)
        auth._cache.has_state_changed = True

        auth._pending_flows["s"] = {
            "flow": {"state": "s"},
            "created_at": time.time(),
        }
        mock_app.acquire_token_by_auth_code_flow.return_value = {
            "access_token": "new-token",
        }
        auth.acquire_token_by_auth_code({"state": "s", "code": "c"})
        # Cache should have been serialized
        auth._cache.serialize.assert_called()
