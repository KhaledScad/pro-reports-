"""Drop-in replacement for starlette.middleware.sessions.SessionMiddleware
that signs the session cookie with stdlib hmac instead of itsdangerous
(itsdangerous isn't installable in this environment)."""
from __future__ import annotations

import base64
import hashlib
import hmac
import json

from starlette.datastructures import MutableHeaders
from starlette.requests import HTTPConnection
from starlette.types import ASGIApp, Message, Receive, Scope, Send


def _sign(secret: bytes, body: bytes) -> bytes:
    return hmac.new(secret, body, hashlib.sha256).digest()


class SessionMiddleware:
    def __init__(self, app: ASGIApp, secret_key: str, session_cookie: str = "session",
                 max_age: int = 14 * 24 * 60 * 60, path: str = "/", same_site: str = "lax",
                 https_only: bool = False) -> None:
        self.app = app
        self.secret_key = secret_key.encode("utf-8")
        self.session_cookie = session_cookie
        self.max_age = max_age
        self.path = path
        security_flags = "httponly; samesite=" + same_site
        if https_only:
            security_flags += "; secure"
        self.security_flags = security_flags

    def _encode(self, data: dict) -> str:
        body = base64.urlsafe_b64encode(json.dumps(data, separators=(",", ":")).encode("utf-8")).rstrip(b"=")
        sig = base64.urlsafe_b64encode(_sign(self.secret_key, body)).rstrip(b"=")
        return (body + b"." + sig).decode("ascii")

    def _decode(self, cookie_value: str) -> dict:
        try:
            body_b64, sig_b64 = cookie_value.encode("ascii").split(b".", 1)
            expected_sig = base64.urlsafe_b64encode(_sign(self.secret_key, body_b64)).rstrip(b"=")
            if not hmac.compare_digest(expected_sig, sig_b64):
                return {}
            padded = body_b64 + b"=" * (-len(body_b64) % 4)
            return json.loads(base64.urlsafe_b64decode(padded))
        except Exception:  # noqa: BLE001
            return {}

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] not in ("http", "websocket"):
            await self.app(scope, receive, send)
            return

        connection = HTTPConnection(scope)
        initial_session_was_empty = True

        if self.session_cookie in connection.cookies:
            scope["session"] = self._decode(connection.cookies[self.session_cookie])
            initial_session_was_empty = not scope["session"]
        else:
            scope["session"] = {}

        async def send_wrapper(message: Message) -> None:
            if message["type"] == "http.response.start":
                if scope["session"]:
                    data = self._encode(scope["session"])
                    header_value = "{}={}; path={}; Max-Age={}; {}".format(
                        self.session_cookie, data, self.path, self.max_age, self.security_flags)
                    MutableHeaders(scope=message).append("Set-Cookie", header_value)
                elif not initial_session_was_empty:
                    header_value = "{}=null; path={}; expires=Thu, 01 Jan 1970 00:00:00 GMT; {}".format(
                        self.session_cookie, self.path, self.security_flags)
                    MutableHeaders(scope=message).append("Set-Cookie", header_value)
            await send(message)

        await self.app(scope, receive, send_wrapper)
