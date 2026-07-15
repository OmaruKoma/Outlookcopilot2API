"""
Outlook Copilot 一键配置向导
运行: python -m outlook_copilot.scripts.setup_wizard
"""
import os, sys, json, re

if sys.stdout.encoding and sys.stdout.encoding.lower() not in ('utf-8', 'utf8'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

CLIENT_ID = "4765445b-32c6-49b0-83e6-1d93765276ca"

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
DATA_DIR = os.path.join(BASE_DIR, "data", "tokens")
ENV_FILE = os.path.join(BASE_DIR, ".env")


def step(title):
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")


def extract_from_console_output(raw: str):
    """Parse tenant, oid from browser console output."""
    raw = re.sub(r'^粘贴\s*=>\s*', '', raw)
    raw = re.sub(r'^PS\s+[^>]+>\s*', '', raw)
    raw = re.sub(r'^>\s*', '', raw)

    tenant = oid = None

    m_oid = re.search(r"OID:\s*([a-f0-9-]+)", raw)
    m_tenant = re.search(r"TENANT:\s*([a-f0-9-]+)", raw)
    if m_oid and m_tenant:
        oid = m_oid.group(1)
        tenant = m_tenant.group(1)

    if not tenant or not oid:
        m = re.search(r"\{", raw)
        if m:
            start = m.start()
            depth = 0
            in_string = False
            escape = False
            for i in range(start, len(raw)):
                c = raw[i]
                if escape:
                    escape = False
                    continue
                if c == '\\' and in_string:
                    escape = True
                    continue
                if c == '"' and not escape:
                    in_string = not in_string
                    continue
                if not in_string:
                    if c == '{':
                        depth += 1
                    elif c == '}':
                        depth -= 1
                        if depth == 0:
                            try:
                                data = json.loads(raw[start:i+1])
                                tenant = tenant or data.get("tenant")
                                oid = oid or data.get("oid")
                            except json.JSONDecodeError:
                                pass
                            break

    return tenant, oid


def get_config_from_browser():
    step("步骤 1: 从浏览器获取配置")
    print()
    print("请在浏览器中完成以下操作:")
    print("  1. 打开 https://outlook.cloud.microsoft 并登录")
    print("  2. 按 F12 打开 DevTools → Console")
    print("  3. 粘贴运行下面这行代码:")
    print()
    print("  (复制下面完整的一行)")
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
    print("=" * 60)
    print("【请复制从这里开始 ==================================】")
    print("=" * 60)
    print()
    print("  提示: 只粘贴 JSON 部分（从 { 开始到 } 结束），不要带 === 标记")
    print()

    raw = input("粘贴 => ").strip()
    if not raw:
        print("错误: 未输入任何内容")
        sys.exit(1)

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        tenant, oid = extract_from_console_output(raw)
        if not tenant or not oid:
            print("错误: 无法解析输出，请确认 Console 输出的是完整 JSON")
            print("提示: 只粘贴从 { 开始到 } 结束的部分")
            sys.exit(1)
    else:
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
    step("步骤 2: 从浏览器 Network 提取 access_token")
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


def save_env(tenant, oid):
    env_content = f"""# Outlook Copilot Configuration
M365_TENANT_ID={tenant}
M365_USER_OID={oid}
M365_CLIENT_ID={CLIENT_ID}
"""
    with open(ENV_FILE, "w") as f:
        f.write(env_content)
    print(f"  环境变量已保存 → {ENV_FILE}")


def main():
    print("=" * 60)
    print("  Outlook Copilot 配置向导 v1.0")
    print("=" * 60)
    print()

    tenant, oid = get_config_from_browser()
    save_env(tenant, oid)
    os.makedirs(DATA_DIR, exist_ok=True)

    step("步骤 2: 选择 Token 获取方式")
    print()
    print("  [1] 自动模式（推荐）— 用 Playwright 浏览器登录一次，之后自动刷新")
    print("      需要: pip install -e '.[browser]' && playwright install chromium")
    print("  [2] 手动模式 — 从 DevTools 复制 access_token（每小时过期需重复）")
    print()
    choice = input("选择 [1/2]（默认 1）=> ").strip() or "1"

    if choice == "1":
        auto_login_step()
    else:
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


def auto_login_step():
    step("步骤 2: 浏览器登录并抓取首个 Token")
    print()
    print("  即将打开浏览器窗口，请登录 Microsoft 账户并勾选“保持登录”。")
    print("  登录后会自动抓取 access_token 并保存。")
    print()
    input("  按回车继续（确保已安装 playwright + chromium）...")
    try:
        from outlook_copilot.browser_auth import fetch_token_blocking, BrowserAuthError
    except ImportError as e:
        print(f"  导入失败: {e}")
        print("  请先运行: pip install -e '.[browser]' && playwright install chromium")
        sys.exit(1)
    token_file = os.path.join(DATA_DIR, "token.txt")
    cache_file = os.path.join(DATA_DIR, "token_cache.json")
    profile_dir = os.path.join(BASE_DIR, "data", "browser_profile")
    try:
        fetch_token_blocking(profile_dir, token_file, cache_file, headless=False)
    except BrowserAuthError as e:
        print(f"  抓取失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
