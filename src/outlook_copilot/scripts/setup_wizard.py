"""
Outlook Copilot 一键配置向导
运行: python -m outlook_copilot.scripts.setup_wizard
"""
import os, sys, json, re, time, urllib.parse

if sys.stdout.encoding and sys.stdout.encoding.lower() not in ('utf-8', 'utf8'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

CLIENT_ID = "4765445b-32c6-49b0-83e6-1d93765276ca"

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
DATA_DIR = os.path.join(BASE_DIR, "data", "tokens")
ENV_FILE = os.path.join(BASE_DIR, ".env")

HOST_URL = os.environ.get(
    "OUTLOOK_COPILOT_HOST_URL",
    "https://outlook.cloud.microsoft/host/b5abf2ae-c16b-4310-8f8a-d3bcdb52f162/entity1-d870f6cd-4aa5-4d42-9626-ab690c041429",
)

EDITOR_SELECTORS = [
    "#m365-chat-editor-target-element",
    '[aria-label="向 Copilot 发送消息"]',
    '[data-lexical-editor="true"]',
    '[role="textbox"][contenteditable="true"]',
]


def step(title):
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")


def save_env(tenant, oid):
    env_content = f"""# Outlook Copilot Configuration
M365_TENANT_ID={tenant}
M365_USER_OID={oid}
M365_CLIENT_ID={CLIENT_ID}
"""
    with open(ENV_FILE, "w") as f:
        f.write(env_content)
    print(f"  环境变量已保存 → {ENV_FILE}")


# ── Manual fallback helpers ──────────────────────────────────────────────

def manual_oid_tenant_step():
    """手动从 DevTools Console 获取 OID/Tenant"""
    step("手动获取 OID / Tenant")
    print()
    print("请在浏览器中完成以下操作:")
    print("  1. 打开 https://outlook.cloud.microsoft 并登录")
    print("  2. 按 F12 打开 DevTools → Console")
    print("  3. 粘贴运行下面这行代码:")
    print()
    print("-" * 60)
    js_snippet = (
        "(() => {"
        "const k = Object.keys(localStorage).find(k => k.includes('|refreshtoken|'));"
        "if (!k) return JSON.stringify({error:'NOT_FOUND: 请先登录 outlook.cloud.microsoft'});"
        "const parts = k.split('|');"
        "const idParts = parts[1].split('.');"
        "return JSON.stringify({oid:idParts[0], tenant:idParts[1]});"
        "})()"
    )
    print(js_snippet)
    print("-" * 60)
    print()
    print("  请复制控制台输出的 JSON（从 { 到 } 的完整内容）")
    print("  如果显示 error 字段，请确认已在 outlook.cloud.microsoft 登录后重试")
    print()

    raw = input("粘贴 => ").strip()
    if not raw:
        print("错误: 未输入任何内容")
        sys.exit(1)

    raw = re.sub(r'^粘贴\s*=>\s*', '', raw)
    raw = re.sub(r'^PS\s+[^>]+>\s*', '', raw)
    raw = re.sub(r'^>\s*', '', raw)

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        print("错误: 无法解析输出，请确认 Console 输出的是完整 JSON")
        print("提示: 只粘贴从 { 开始到 } 结束的部分")
        sys.exit(1)

    if "error" in data:
        print(f"错误: {data['error']}")
        sys.exit(1)

    tenant = data.get("tenant")
    oid = data.get("oid")
    if not tenant or not oid:
        print("错误: 无法获取 Tenant ID 或 User OID")
        sys.exit(1)

    return tenant, oid


def token_extract_step():
    """手动从 DevTools Network 提取 access_token"""
    step("手动提取 access_token")
    print()
    print("操作步骤（共 7 步）:")
    print()
    print("  ┌─────────────────────────────────────────────────────────┐")
    print("  │  1. 打开 https://outlook.cloud.microsoft 并登录         │")
    print("  │  2. 按 F12 打开 DevTools                                │")
    print("  │  3. 点击 Network（网络）选项卡                           │")
    print("  │  4. 在筛选输入框输入: ws （过滤 WebSocket 请求）        │")
    print("  │  5. 刷新页面（F5 / Ctrl+R）                              │")
    print("  │  6. 列表中出现 substrate.office.com 请求 → 点击它       │")
    print("  │  7. 在右侧 Headers 面板的 URL 最后找到:                 │")
    print("  │     access_token=eyJ...（以 eyJ 开头的一长串）         │")
    print("  └─────────────────────────────────────────────────────────┘")
    print()
    print("  也可从请求 URL 直接提取:")
    print("  │  wss://substrate.office.com/m365Copilot/Chathub/...")
    print("  │      ?chatsessionid=...")
    print("  │      &access_token=eyJ0eXAiOiJKV1Q...← 复制这个值")
    print()
    print("  注意：")
    print("  · access_token 是 eyJ 开头的一长串（约 2000+ 字符）")
    print("  · 请确保复制完整，不要截断（可从 URL 的 &access_token= 后面开始选）")
    print("  · Token 有效期约 1 小时，到期后重复本步骤")
    print()
    print("  粘贴 Token (以 eyJ 开头）到下面:")
    print()

    token = input("粘贴 Token => ").strip()
    if not token:
        print("错误: 未输入 Token")
        sys.exit(1)

    if not token.startswith("eyJ"):
        print("错误: Token 应以 eyJ 开头，请确认复制的是完整的 access_token")
        sys.exit(1)

    token_file = os.path.join(DATA_DIR, "token.txt")
    with open(token_file, 'w') as f:
        f.write(token)
    print(f"  Token 已保存 → {token_file}")
    print(f"  （{len(token)} 字符，有效期约 1 小时）")
    print(f"\n  过期后重新运行: outlook-copilot --refresh")


# ── Browser auto-login ──────────────────────────────────────────────────

def _is_target_ws(ws_url):
    try:
        parsed = urllib.parse.urlparse(ws_url)
        return parsed.hostname == "substrate.office.com" and "/m365Copilot/Chathub/" in parsed.path
    except Exception:
        return False


def _extract_token_from_ws_url(ws_url):
    try:
        parsed = urllib.parse.urlparse(ws_url)
        qs = urllib.parse.parse_qs(parsed.query)
        vals = qs.get("access_token")
        return vals[0] if vals else None
    except Exception:
        return None


def _extract_oid_tenant_from_ws_url(ws_url):
    """The substrate Chathub URL path ends with '{oid}@{tenant}'. Return
    (oid, tenant) or (None, None)."""
    try:
        parsed = urllib.parse.urlparse(ws_url)
        marker = "/m365Copilot/Chathub/"
        idx = parsed.path.find(marker)
        if idx < 0:
            return None, None
        tail = urllib.parse.unquote(parsed.path[idx + len(marker):])
        # tail looks like 'oid@tenant' (possibly with a trailing '/')
        tail = tail.split("/")[0]
        if "@" not in tail:
            return None, None
        oid, tenant = tail.split("@", 1)
        return (oid or None), (tenant or None)
    except Exception:
        return None, None


def _find_editor(page):
    for sel in EDITOR_SELECTORS:
        loc = page.locator(sel)
        try:
            if loc.count() > 0:
                return loc.first
        except Exception:
            pass
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


def browser_full_setup():
    """Playwright 全自动：打开浏览器 → 登录 → 提取 OID/Tenant → 抓取 Token"""
    import base64

    step("浏览器自动登录")
    print()
    print("  即将打开浏览器窗口，请登录 Microsoft 账户并勾选'保持登录'。")
    print("  登录后会自动提取全部配置，无需任何手动操作。")
    print()
    input("  按回车继续（确保已安装 playwright + chromium）...")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        print(f"  导入失败: {e}")
        print("  请先运行: pip install -e '.[browser]' && playwright install chromium")
        sys.exit(1)

    token_file = os.path.join(DATA_DIR, "token.txt")
    cache_file = os.path.join(DATA_DIR, "token_cache.json")
    profile_dir = os.path.join(BASE_DIR, "data", "browser_profile")
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(profile_dir, exist_ok=True)

    captured = {"url": None}

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=profile_dir,
            headless=False,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-first-run",
            ],
        )
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()

            def on_ws(ws):
                if captured["url"] is None and _is_target_ws(ws.url):
                    captured["url"] = ws.url

            page.on("websocket", on_ws)

            # Navigate straight to the Copilot host page (same as --login).
            print()
            print("  正在打开 Copilot 页面 ...")
            page.goto(HOST_URL, wait_until="domcontentloaded", timeout=60000)

            # If not signed in, the page redirects to the Microsoft login. Wait
            # for the user to finish, then re-navigate to the Copilot host page.
            if "login.microsoftonline.com" in page.url:
                print("  请在浏览器中登录 Microsoft 账户（勾选'保持登录'）...")
                deadline = time.time() + 300
                while time.time() < deadline:
                    if "login.microsoftonline.com" not in page.url:
                        page.wait_for_timeout(3000)
                        break
                    page.wait_for_timeout(1000)
                else:
                    print("  错误: 登录超时")
                    sys.exit(1)
                page.goto(HOST_URL, wait_until="domcontentloaded", timeout=60000)

            if "login.microsoftonline.com" in page.url:
                print("  错误: 浏览器未成功登录，请重试")
                sys.exit(1)

            print("  正在抓取 access_token ...")
            try:
                page.wait_for_load_state("networkidle", timeout=30000)
            except Exception:
                pass
            try:
                page.wait_for_selector(EDITOR_SELECTORS[0], timeout=30000, state="attached")
            except Exception:
                pass

            editor = _find_editor(page)
            deadline = time.time() + 30
            while editor is None and time.time() < deadline:
                page.wait_for_timeout(1000)
                editor = _find_editor(page)

            if editor is None:
                print("  错误: 找不到 Copilot 输入框，页面 UI 可能已改变")
                sys.exit(1)

            editor.click()
            page.wait_for_timeout(300)
            page.keyboard.type("a")

            capture_deadline = time.time() + 45
            while captured["url"] is None and time.time() < capture_deadline:
                page.wait_for_timeout(250)

            if captured["url"] is None:
                print("  错误: 未捕获到 Copilot WebSocket 连接")
                sys.exit(1)

            token = _extract_token_from_ws_url(captured["url"])
            if not token:
                print("  错误: WebSocket URL 中未找到 access_token")
                sys.exit(1)

            # Decode the JWT once: it gives us exp for the cache and oid/tid as a
            # reliable cross-check for the values parsed out of the WS URL path.
            claims = {}
            try:
                padded = token.split(".")[1] + "==="
                claims = json.loads(base64.urlsafe_b64decode(padded))
            except Exception:
                claims = {}
            exp = claims.get("exp", 0)
            if exp <= time.time():
                print("  错误: 抓取的 token 已过期")
                sys.exit(1)

            # OID/tenant come straight from the WS URL path ({oid}@{tenant}),
            # falling back to the JWT claims (oid / tid). No localStorage needed.
            oid, tenant = _extract_oid_tenant_from_ws_url(captured["url"])
            oid = oid or claims.get("oid")
            tenant = tenant or claims.get("tid")
            if not oid or not tenant:
                print("  错误: 未能从 WebSocket URL 或 token 中解析 OID/Tenant")
                sys.exit(1)

            print(f"  Tenant ID: {tenant}")
            print(f"  User OID:  {oid}")
            save_env(tenant, oid)

            with open(token_file, "w") as f:
                f.write(token)
            with open(cache_file, "w") as f:
                json.dump({"access_token": token, "expires_at": exp}, f)

            print(f"  Token 已保存 → {token_file}")
            exp_str = time.strftime("%H:%M:%S", time.localtime(exp))
            print(f"  （{len(token)} 字符，有效期至 {exp_str}）")
        finally:
            try:
                ctx.close()
            except Exception:
                pass


# ── Main ────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  Outlook Copilot 配置向导 v2.0")
    print("=" * 60)
    print()
    print("  [1] 浏览器自动登录（推荐）— 打开浏览器，登录后自动提取全部配置")
    print("      需要: pip install -e '.[browser]' && playwright install chromium")
    print("  [2] 手动配置 — 从浏览器 DevTools 复制信息")
    print()
    choice = input("选择 [1/2]（默认 1）=> ").strip() or "1"

    if choice == "1":
        browser_full_setup()
    else:
        tenant, oid = manual_oid_tenant_step()
        save_env(tenant, oid)
        os.makedirs(DATA_DIR, exist_ok=True)
        token_extract_step()

    step("配置完成!")
    print()
    print("使用方式:")
    print("  outlook-copilot \"你好\"              # CLI 提问")
    print("  outlook-copilot -i                  # 交互模式")
    print("  outlook-copilot --list-models       # 列出模型")
    print("  outlook-copilot-server --port 8000  # 启动 API 服务")
    print()
    print(f"Token 存储: {DATA_DIR}")
    print(f"配置文件:   {ENV_FILE}")
    print()
    if choice == "1":
        print("自动刷新（无人值守）:")
        print("  outlook-copilot-server --auto-refresh --port 8000")
        print("  或设置 OUTLOOK_COPILOT_AUTO_REFRESH=1")
        print()
        print("登录态失效后重新登录:")
        print("  outlook-copilot --login")
    else:
        print("Token 过期后:")
        print("  1. F12 → Network → ws → 刷新 → substrate 请求 → 复制 access_token")
        print("  2. outlook-copilot --refresh")


if __name__ == "__main__":
    main()
