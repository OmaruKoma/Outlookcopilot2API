import json, urllib.request, urllib.parse, uuid, time, os, ssl, hashlib, base64
from datetime import datetime

TOKEN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
AUTHORIZE_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize"


class TokenRefreshError(Exception):
    pass


class TokenManager:
    def __init__(self, tenant, client_id, scope, rt_file, cache_file, token_file=None):
        self.tenant = tenant
        self.client_id = client_id
        self.scope = scope
        self.rt_file = rt_file
        self.cache_file = cache_file
        self.token_file = token_file
        self._token_url = TOKEN_URL.format(tenant=tenant)

    def _read_rt(self):
        with open(self.rt_file) as f:
            raw = f.read().strip()
        if not raw:
            raise TokenRefreshError("Refresh token file is empty")
        from .scripts.crypto import decrypt
        try:
            return decrypt(raw)
        except Exception as e:
            raise TokenRefreshError(
                f"无法解密 refresh token: {e}\n"
                 f"可能原因: 加密密钥文件 (~/.m365-copilot/encryption.key) 不匹配或已损坏\n"
                 f"解决方法: 重新运行 outlook-copilot-setup"
            )

    def _write_rt(self, token):
        from .scripts.crypto import encrypt
        with open(self.rt_file, 'w') as f:
            f.write(encrypt(token))

    def refresh(self):
        if not os.path.exists(self.rt_file):
            raise TokenRefreshError(f"Refresh token not found: {self.rt_file}")
        rt = self._read_rt()
        data = urllib.parse.urlencode({
            'client_id': self.client_id, 'refresh_token': rt,
            'grant_type': 'refresh_token', 'scope': self.scope,
            'redirect_uri': 'https://outlook.cloud.microsoft',
        }).encode()
        req = urllib.request.Request(self._token_url, data=data)
        req.add_header('Content-Type', 'application/x-www-form-urlencoded')
        req.add_header('User-Agent', 'Mozilla/5.0')
        try:
            with urllib.request.urlopen(req, context=ssl.create_default_context()) as resp:
                result = json.loads(resp.read())
        except urllib.error.HTTPError as e:
            err = json.loads(e.read())
            desc = err.get('error_description', '')[:300]
            code = err.get('error', '')
            if 'AADSTS700082' in desc or 'AADSTS700022' in desc:
                hint = "Refresh token 已过期。请重新运行 outlook-copilot-setup"
            else:
                hint = f"请重新运行 outlook-copilot-setup"
            raise TokenRefreshError(
                f"Refresh failed: {code}: {desc}\n{hint}"
            )
        if 'refresh_token' in result:
            self._write_rt(result['refresh_token'])
        cache = {
            'access_token': result['access_token'],
            'expires_at': time.time() + result.get('expires_in', 3600),
        }
        with open(self.cache_file, 'w') as f:
            json.dump(cache, f)
        return result['access_token']

    def get(self):
        if self.token_file and os.path.exists(self.token_file):
            token = open(self.token_file).read().strip()
            if token:
                payload = self._decode_token(token)
                if payload and time.time() < payload.get('exp', 0) - 60:
                    return token
        try:
            with open(self.cache_file) as f:
                cache = json.load(f)
            if cache['expires_at'] > time.time() + 60:
                return cache['access_token']
        except Exception:
            pass
        if not os.path.exists(self.rt_file):
            raise TokenRefreshError(
                "Refresh token not found. "
                "If your access token expired, run: outlook-copilot --refresh"
            )
        return self.refresh()

    @staticmethod
    def _decode_token(token):
        try:
            parts = token.split('.')
            padded = parts[1] + '=' * (4 - len(parts[1]) % 4)
            return json.loads(base64.urlsafe_b64decode(padded))
        except Exception:
            return None

    def refresh_interactive(self):
        if not self.token_file:
            print("错误: 未配置 token_file")
            return False
        print()
        print("Token 已过期。获取新 token:")
        print()
        print("  1. 打开 https://outlook.cloud.microsoft 并登录")
        print("  2. F12 → Network → 筛选 ws")
        print("  3. 刷新页面（F5）")
        print("  4. 点击 substrate.office.com 的 WS 请求")
        print("  5. 从 URL 参数中复制 access_token=...（eyJ 开头）")
        print()
        print("  也可用 outlook-copilot-setup 重新走完整流程")
        print()
        token = input("粘贴 Token => ").strip()
        if not token:
            print("未输入 token")
            return False
        if not token.startswith("eyJ"):
            print("错误: Token 应以 eyJ 开头，请确认复制的是完整的 access_token")
            return False
        payload = self._decode_token(token)
        if not payload:
            print("错误: 无法解析 token")
            return False
        with open(self.token_file, 'w') as f:
            f.write(token)
        exp_str = datetime.fromtimestamp(payload['exp']).strftime('%Y-%m-%d %H:%M:%S')
        cache = {'access_token': token, 'expires_at': payload['exp']}
        with open(self.cache_file, 'w') as f:
            json.dump(cache, f)
        print(f"Token 已更新（有效期至 {exp_str}，约 {int((payload['exp']-time.time())/60)} 分钟）")
        return True

    def setup(self):
        if self._try_wam():
            return True
        return self._try_auth_code()

    def _try_wam(self):
        try:
            from msal.broker import _signin_interactively
            print("正在打开 Windows 账户选择窗口...")
            result = _signin_interactively(
                f"https://login.microsoftonline.com/{self.tenant}",
                self.client_id,
                [self.scope, "openid", "profile", "offline_access"],
                parent_window_handle=None,
                prompt="select_account",
            )
            if "access_token" in result:
                self._save_tokens(result)
                return True
            err = result.get("error", "?")
            desc = result.get("error_description", "")
            print(f"  WAM 失败: {err}" + (f" {desc[:200]}" if desc else ""))
        except ImportError:
            print("  WAM broker 未安装 (pip install msal[broker])")
        except Exception as e:
            print(f"  WAM 不可用: {e}")
        return False

    def _try_auth_code(self):
        verifier = base64.urlsafe_b64encode(os.urandom(64)).rstrip(b'=').decode()
        challenge = base64.urlsafe_b64encode(
            hashlib.sha256(verifier.encode()).digest()
        ).rstrip(b'=').decode()
        params = {
            'client_id': self.client_id, 'response_type': 'code',
            'response_mode': 'fragment',
            'scope': f'{self.scope} openid profile offline_access',
            'redirect_uri': 'https://outlook.cloud.microsoft',
            'state': uuid.uuid4().hex, 'nonce': uuid.uuid4().hex,
            'code_challenge': challenge, 'code_challenge_method': 'S256',
            'prompt': 'select_account',
        }
        url = AUTHORIZE_URL.format(tenant=self.tenant) + "?" + urllib.parse.urlencode(params)
        print("\n" + "=" * 60)
        print("浏览器认证:")
        print("1. 打开链接并登录")
        print("2. 从地址栏复制 code=... 的值")
        print("\n链接:")
        print(url)
        print("=" * 60)
        code = input("\n请输入 authorization code: ").strip()
        if not code:
            print("未输入 code")
            return False
        if 'code=' in code:
            import urllib.parse as _up
            if '#' in code:
                code = code.split('#')[1]
            code = _up.parse_qs(code).get('code', [code])[0]
        data = urllib.parse.urlencode({
            'client_id': self.client_id, 'code': code,
            'redirect_uri': 'https://outlook.cloud.microsoft',
            'grant_type': 'authorization_code', 'code_verifier': verifier,
        }).encode()
        req = urllib.request.Request(self._token_url, data=data)
        req.add_header('Content-Type', 'application/x-www-form-urlencoded')
        try:
            with urllib.request.urlopen(req, context=ssl.create_default_context()) as resp:
                result = json.loads(resp.read())
            if 'refresh_token' in result:
                self._save_tokens(result)
                return True
            error_desc = result.get('error_description', '')[:200]
            print(f"\n失败: {result.get('error')}: {error_desc}")
            if 'AADSTS700084' in error_desc:
                print("提示: SPA client 的 refresh token 有效期约 24h，已过期，需重新登录获取")
        except urllib.error.HTTPError as e:
            err = json.loads(e.read())
            print(f"\n失败: {err.get('error')}: {err.get('error_description', '')[:200]}")
        return False

    def _save_tokens(self, result):
        if 'refresh_token' in result:
            from .scripts.crypto import encrypt
            with open(self.rt_file, 'w') as f:
                f.write(encrypt(result['refresh_token']))
        cache = {
            'access_token': result['access_token'],
            'expires_at': time.time() + result.get('expires_in', 3600),
        }
        with open(self.cache_file, 'w') as f:
            json.dump(cache, f)
        print("认证成功！Refresh token 已保存")
