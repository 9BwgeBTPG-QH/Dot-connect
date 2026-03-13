"""Microsoft Graph API OAuth2 認証管理 (MSAL + PKCE)."""

import logging
import os
import stat
import time
from pathlib import Path
from typing import Optional

import msal

log = logging.getLogger(__name__)

# Pending auth flows expire after 10 minutes
_FLOW_TTL_SECONDS = 600


class GraphAuth:
    """MSAL PublicClientApplication ベースの OAuth2 認証."""

    SCOPES = ["Mail.Read", "User.Read"]

    def __init__(
        self,
        client_id: str,
        tenant_id: str,
        redirect_uri: str = "http://localhost:8000/auth/callback",
        cache_path: Optional[str] = None,
    ):
        self._client_id = client_id
        self._tenant_id = tenant_id
        self._redirect_uri = redirect_uri
        self._cache_path = Path(
            cache_path or os.path.join(Path.home(), ".dot-connect", "token_cache.bin")
        )

        # Pending auth code flows keyed by state
        self._pending_flows: dict[str, dict] = {}

        # Initialize token cache
        self._cache = msal.SerializableTokenCache()
        self._load_cache()

        self._app = msal.PublicClientApplication(
            client_id=self._client_id,
            authority=f"https://login.microsoftonline.com/{self._tenant_id}",
            token_cache=self._cache,
        )

    # ------------------------------------------------------------------
    # Cache persistence
    # ------------------------------------------------------------------

    def _load_cache(self) -> None:
        """ディスクからトークンキャッシュを読み込む."""
        if self._cache_path.exists():
            try:
                self._cache.deserialize(self._cache_path.read_text(encoding="utf-8"))
            except Exception as e:
                log.warning("トークンキャッシュの読み込みに失敗: %s", e)

    def _save_cache(self) -> None:
        """トークンキャッシュをディスクに保存 (パーミッション 0600)."""
        if self._cache.has_state_changed:
            try:
                self._cache_path.parent.mkdir(parents=True, exist_ok=True)
                self._cache_path.write_text(
                    self._cache.serialize(), encoding="utf-8"
                )
                # POSIX のみ: パーミッション制限
                try:
                    os.chmod(self._cache_path, stat.S_IRUSR | stat.S_IWUSR)
                except OSError:
                    pass
            except Exception as e:
                log.warning("トークンキャッシュの保存に失敗: %s", e)

    # ------------------------------------------------------------------
    # Auth flow
    # ------------------------------------------------------------------

    def get_auth_url(self, state: str) -> str:
        """認証コードフロー (PKCE) を開始し、ログインURLを返す."""
        self._cleanup_expired_flows()
        flow = self._app.initiate_auth_code_flow(
            scopes=self.SCOPES,
            redirect_uri=self._redirect_uri,
            state=state,
        )
        self._pending_flows[state] = {
            "flow": flow,
            "created_at": time.time(),
        }
        return flow["auth_uri"]

    def acquire_token_by_auth_code(self, auth_response: dict) -> dict:
        """コールバックの応答からトークンを取得.

        Args:
            auth_response: コールバックのクエリパラメータ dict (code, state 等)

        Returns:
            MSAL トークンレスポンス dict。成功時は "access_token" キーを含む。
        """
        state = auth_response.get("state", "")
        pending = self._pending_flows.pop(state, None)
        if pending is None:
            return {"error": "invalid_state", "error_description": "認証フローが見つかりません。再度サインインしてください。"}

        result = self._app.acquire_token_by_auth_code_flow(
            pending["flow"],
            auth_response,
        )
        self._save_cache()
        return result

    def get_access_token(self) -> Optional[str]:
        """キャッシュからアクセストークンを取得 (自動リフレッシュ).

        Returns:
            有効なアクセストークン文字列。取得できない場合は None。
        """
        accounts = self._app.get_accounts()
        if not accounts:
            return None

        result = self._app.acquire_token_silent(
            scopes=self.SCOPES,
            account=accounts[0],
        )
        if result and "access_token" in result:
            self._save_cache()
            return result["access_token"]
        return None

    def is_authenticated(self) -> bool:
        """有効なトークンがキャッシュにあるか."""
        return self.get_access_token() is not None

    def sign_out(self) -> None:
        """全アカウントをサインアウトし、キャッシュをクリア."""
        accounts = self._app.get_accounts()
        for account in accounts:
            self._app.remove_account(account)
        self._save_cache()
        # キャッシュファイルも削除
        if self._cache_path.exists():
            try:
                self._cache_path.unlink()
            except OSError:
                pass

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def _cleanup_expired_flows(self) -> None:
        """TTL を超えた pending flows を削除."""
        now = time.time()
        expired = [
            s for s, v in self._pending_flows.items()
            if now - v["created_at"] > _FLOW_TTL_SECONDS
        ]
        for s in expired:
            del self._pending_flows[s]
