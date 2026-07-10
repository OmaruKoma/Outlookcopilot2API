"""Cookie-based authentication for the M365 REST API"""
import json, os, time
from pathlib import Path

COOKIE_FILE = Path(__file__).resolve().parent.parent.parent / 'data' / 'tokens' / 'm365_cookies.json'

class CookieStore:
    def __init__(self, cookie_path=None):
        self.cookie_path = Path(cookie_path or COOKIE_FILE)

    def _load(self):
        if not self.cookie_path.exists():
            return None
        with open(self.cookie_path) as f:
            return json.load(f)

    def get_cookie_string(self):
        data = self._load()
        if not data:
            return None
        cookies = data.get('cookies', [])
        return '; '.join([f"{c['name']}={c['value']}" for c in cookies])

    def get_cookie_dict(self):
        data = self._load()
        if not data:
            return {}
        return {c['name']: c['value'] for c in data.get('cookies', [])}

    def has_cookies(self):
        data = self._load()
        if not data:
            return False
        return len(data.get('cookies', [])) > 0

    def save(self, cookies, domain='m365.cloud.microsoft'):
        self.cookie_path.parent.mkdir(parents=True, exist_ok=True)
        data = {
            'domain': domain,
            'extracted_at': time.time(),
            'cookies': cookies,
        }
        with open(self.cookie_path, 'w') as f:
            json.dump(data, f, indent=2)

    def build_opener(self):
        import urllib.request, http.cookiejar
        cookie_str = self.get_cookie_string()
        if not cookie_str:
            return None
        cookie_jar = http.cookiejar.CookieJar()
        for c in self._load().get('cookies', []):
            ck = http.cookiejar.Cookie(
                version=0, name=c['name'], value=c['value'],
                port=None, port_specified=False,
                domain=c.get('domain', 'm365.cloud.microsoft'),
                domain_specified=True,
                domain_initial_dot=c.get('domain', '').startswith('.'),
                path=c.get('path', '/'), path_specified=True,
                secure=c.get('secure', True),
                expires=None, discard=False,
                comment=None, comment_url=None,
                rest={'HttpOnly': c.get('httpOnly', True)},
                rfc2109=False)
            cookie_jar.set_cookie(ck)
        return urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(cookie_jar))

    @staticmethod
    def extract_from_browser():
        """Launch Edge via Playwright to extract m365.cloud.microsoft cookies."""
        import asyncio
        from playwright.async_api import async_playwright

        async def _extract():
            user_data = os.path.expanduser(
                '~\\AppData\\Local\\Microsoft\\Edge\\User Data')
            async with async_playwright() as p:
                ctx = await p.chromium.launch_persistent_context(
                    user_data_dir=user_data,
                    channel='msedge',
                    headless=False,
                    args=['--disable-blink-features=AutomationControlled',
                          '--no-first-run'],
                )
                page = ctx.pages[0] if ctx.pages else await ctx.new_page()
                await page.goto('https://m365.cloud.microsoft',
                                wait_until='networkidle', timeout=60000)
                await asyncio.sleep(5)

                if 'login.microsoftonline.com' in page.url:
                    print('请在弹出的浏览器窗口中登录您的 Microsoft 账户...')
                    for _ in range(180):
                        await asyncio.sleep(1)
                        if 'login.microsoftonline.com' not in page.url:
                            await asyncio.sleep(3)
                            break

                cookies = await ctx.cookies()
                target = [c for c in cookies
                          if any(d in c.get('domain', '')
                                 for d in ['m365.cloud.microsoft',
                                           '.microsoft.com',
                                           'login.microsoftonline.com'])]
                await ctx.close()
                return target

        return asyncio.run(_extract())
