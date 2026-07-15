"""Automated access_token extraction via a persistent, logged-in browser.

The Copilot backend access_token (appid c0ab8ce9-..., aud substrate/sydney) is
generated server-side via an OBO exchange that third parties cannot perform.
The only place it surfaces client-side is the WebSocket URL that Outlook opens
to substrate.office.com the moment the user types into the Copilot input box.

This module drives a persistent (logged-in) browser with Playwright, focuses
the Copilot editor and types a single character (WITHOUT sending) to trigger
that WebSocket, captures the URL, extracts and validates the token, and writes
it back to token.txt / token_cache.json.

Playwright is an optional dependency. Install with:
    pip install -e '.[browser]'
    playwright install chromium
"""
import base64
import json
import os
import time
import urllib.parse
from datetime import datetime

# The Outlook Copilot host page. Hardcoded and stable across accounts per the
# reverse-engineering notes; overridable via env for resilience to UI changes.
DEFAULT_HOST_URL = os.environ.get(
    "OUTLOOK_COPILOT_HOST_URL",
    "https://outlook.cloud.microsoft/host/b5abf2ae-c16b-4310-8f8a-d3bcdb52f162/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429",
)

# Selectors for the Lexical rich-text editor input box (in priority order).
EDITOR_SELECTORS = [
    "#m365-chat-editor-target-element",
    '[aria-label="向 Copilot 发送消息"]',
    '[data-lexical-editor="true"]',
    '[role="textbox"][contenteditable="true"]',
]

# The WebSocket we care about: substrate Chathub carries access_token in its URL.
_WS_HOST = "substrate.office.com"
_WS_PATH_MARKER = "/m365Copilot/Chathub/"

# When set, dump the current URL, all frame URLs, and a screenshot on failure.
DEBUG = os.environ.get("OUTLOOK_COPILOT_BROWSER_DEBUG", "0").strip() in ("1", "true", "yes")

# Expected token claims (used to reject wrong-audience / wrong-client tokens).
_EXPECTED_AUD = "https://substrate.office.com/sydney"
_EXPECTED_APPID = "c0ab8ce9-e9a0-42e7-b064-33d422df41f1"


class BrowserAuthError(Exception):
    pass


class LoginRequiredError(BrowserAuthError):
    """Raised when the persistent profile is not logged in (headless landed on
    the Microsoft login page). The caller should re-run in headed mode."""


def decode_jwt_claims(token):
    """Decode a JWT payload segment without verifying the signature. Returns a
    dict of claims, or None if the token is not a parseable JWT."""
    try:
        parts = token.split(".")
        if len(parts) < 2:
            return None
        padded = parts[1] + "=" * (-len(parts[1]) % 4)
        return json.loads(base64.urlsafe_b64decode(padded))
    except Exception:
        return None


def extract_token_from_ws_url(ws_url):
    """Pull the access_token query parameter out of a substrate Chathub WS URL.
    Returns the token string, or None if absent."""
    try:
        parsed = urllib.parse.urlparse(ws_url)
        qs = urllib.parse.parse_qs(parsed.query)
        vals = qs.get("access_token")
        return vals[0] if vals else None
    except Exception:
        return None


def is_target_ws(ws_url):
    """True if the URL is the substrate Chathub WebSocket that carries a token."""
    try:
        parsed = urllib.parse.urlparse(ws_url)
        return parsed.hostname == _WS_HOST and _WS_PATH_MARKER in parsed.path
    except Exception:
        return False


def validate_token(token, min_seconds_left=0):
    """Validate that a captured token is the Copilot backend token we want.

    Returns the exp (unix ts) on success; raises BrowserAuthError otherwise."""
    claims = decode_jwt_claims(token)
    if not claims:
        raise BrowserAuthError("Captured token is not a parseable JWT")
    aud = claims.get("aud", "")
    if aud != _EXPECTED_AUD:
        raise BrowserAuthError(f"Unexpected token audience: {aud!r}")
    appid = claims.get("appid", "")
    if appid != _EXPECTED_APPID:
        raise BrowserAuthError(
            f"Unexpected token appid: {appid!r} (need the Copilot backend client)"
        )
    exp = claims.get("exp", 0)
    if exp <= time.time() + min_seconds_left:
        raise BrowserAuthError("Captured token is already expired or too close to expiry")
    return exp


def _save_token(token, exp, token_file, cache_file):
    os.makedirs(os.path.dirname(token_file), exist_ok=True)
    with open(token_file, "w") as f:
        f.write(token)
    with open(cache_file, "w") as f:
        json.dump({"access_token": token, "expires_at": exp}, f)


class BrowserTokenExtractor:
    """Drives a persistent-context browser to capture a fresh Copilot token."""

    def __init__(self, profile_dir, token_file, cache_file,
                 host_url=DEFAULT_HOST_URL, headless=True,
                 nav_timeout=60000, capture_timeout=45.0, editor_timeout=30.0):
        self.profile_dir = profile_dir
        self.token_file = token_file
        self.cache_file = cache_file
        self.host_url = host_url
        self.headless = headless
        self.nav_timeout = nav_timeout
        self.capture_timeout = capture_timeout
        self.editor_timeout = editor_timeout

    def _find_editor(self, page):
        """Locate the Copilot editor in the top frame or any child frame.
        Returns a Playwright Locator or None."""
        for sel in EDITOR_SELECTORS:
            loc = page.locator(sel)
            try:
                if loc.count() > 0:
                    return loc.first
            except Exception:
                pass
        # Fall back to searching inside iframes (the host page embeds Copilot).
        for frame in page.frames:
            if frame == page.main_frame:
                continue
            for sel in EDITOR_SELECTORS:
                try:
                    loc = frame.locator(sel)
                    if loc.count() > 0:
                        return loc.first
                except Exception:
                    pass
        return None

    def fetch_token(self):
        """Launch the browser, trigger the WebSocket, capture and persist the
        token. Returns (token, exp). Raises LoginRequiredError / BrowserAuthError."""
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as e:
            raise BrowserAuthError(
                "Playwright is not installed. Run: pip install -e '.[browser]' "
                "&& playwright install chromium"
            ) from e

        os.makedirs(self.profile_dir, exist_ok=True)
        captured = {"url": None}

        with sync_playwright() as p:
            ctx = p.chromium.launch_persistent_context(
                user_data_dir=self.profile_dir,
                headless=self.headless,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-first-run",
                ],
            )
            try:
                page = ctx.pages[0] if ctx.pages else ctx.new_page()

                def on_ws(ws):
                    if captured["url"] is None and is_target_ws(ws.url):
                        captured["url"] = ws.url

                page.on("websocket", on_ws)

                page.goto(self.host_url, wait_until="domcontentloaded",
                          timeout=self.nav_timeout)

                if "login.microsoftonline.com" in page.url:
                    if self.headless:
                        raise LoginRequiredError(
                            "Browser profile is not logged in. Re-run with "
                            "headless=False (outlook-copilot --login) to sign in."
                        )
                    # Headed: wait for the user to finish signing in, then
                    # navigate back to the Copilot host page (login usually
                    # lands on the Outlook mailbox, not the Copilot page).
                    self._wait_for_login(page)
                    page.goto(self.host_url, wait_until="domcontentloaded",
                              timeout=self.nav_timeout)

                try:
                    token = self._trigger_and_capture(page, captured)
                except BrowserAuthError:
                    self._dump_debug(page)
                    raise
                exp = validate_token(token)
                _save_token(token, exp, self.token_file, self.cache_file)
                return token, exp
            finally:
                try:
                    ctx.close()
                except Exception:
                    pass

    def _dump_debug(self, page):
        if not DEBUG:
            return
        try:
            data_dir = os.path.dirname(self.token_file)
            print(f"[debug] current url: {page.url}")
            for i, frame in enumerate(page.frames):
                print(f"[debug] frame[{i}]: {frame.url}")
            shot = os.path.join(data_dir, "browser_debug.png")
            page.screenshot(path=shot, full_page=True)
            print(f"[debug] screenshot saved: {shot}")
        except Exception as e:
            print(f"[debug] dump failed: {e}")

    def _wait_for_login(self, page, timeout=300):
        deadline = time.time() + timeout
        print("请在打开的浏览器窗口中登录 Microsoft 账户（勾选“保持登录”）...")
        while time.time() < deadline:
            if "login.microsoftonline.com" not in page.url:
                page.wait_for_timeout(3000)
                return
            page.wait_for_timeout(1000)
        raise LoginRequiredError("登录超时")

    def _wait_for_editor(self, page):
        """Wait for the Copilot editor to appear. It normally lives in the top
        document, but we fall back to child frames for layout variants.
        Returns a Locator or None."""
        # Let the SPA settle before probing.
        try:
            page.wait_for_load_state("networkidle", timeout=self.editor_timeout * 1000)
        except Exception:
            pass
        primary = EDITOR_SELECTORS[0]
        try:
            page.wait_for_selector(primary, timeout=self.editor_timeout * 1000,
                                   state="attached")
        except Exception:
            pass
        return self._find_editor(page)

    def _trigger_and_capture(self, page, captured):
        """Focus the Lexical editor, type one char (no send) to open the WS,
        and wait for its URL to be captured."""
        editor = self._wait_for_editor(page)
        deadline = time.time() + self.editor_timeout
        while editor is None and time.time() < deadline:
            page.wait_for_timeout(1000)
            editor = self._find_editor(page)
        if editor is None:
            raise BrowserAuthError(
                "Could not find the Copilot input editor. The page UI may have "
                "changed; set OUTLOOK_COPILOT_HOST_URL or update EDITOR_SELECTORS."
            )

        editor.click()
        page.wait_for_timeout(300)
        # Lexical needs real key events, not fill(); a single char is enough to
        # open the substrate WebSocket without sending a message.
        page.keyboard.type("a")

        capture_deadline = time.time() + self.capture_timeout
        while captured["url"] is None and time.time() < capture_deadline:
            page.wait_for_timeout(250)

        if captured["url"] is None:
            raise BrowserAuthError(
                "Typed into the editor but no substrate WebSocket was observed "
                f"within {self.capture_timeout}s."
            )

        token = extract_token_from_ws_url(captured["url"])
        if not token:
            raise BrowserAuthError("Captured WebSocket URL had no access_token parameter")
        return token


def fetch_token_blocking(profile_dir, token_file, cache_file,
                         host_url=DEFAULT_HOST_URL, headless=True):
    """Convenience wrapper. Returns (token, exp)."""
    extractor = BrowserTokenExtractor(
        profile_dir, token_file, cache_file, host_url=host_url, headless=headless
    )
    token, exp = extractor.fetch_token()
    exp_str = datetime.fromtimestamp(exp).strftime("%Y-%m-%d %H:%M:%S")
    mins = int((exp - time.time()) / 60)
    print(f"Token 已刷新（有效期至 {exp_str}，约 {mins} 分钟）")
    return token, exp
