# 技术文档

## 概述

M365 Copilot2API 是一个代理层，将 Microsoft 365 Copilot 的私有 SignalR WebSocket 协议转换为标准的 OpenAI 兼容 HTTP API。

## 架构

```
  浏览器 (持久登录态)
       │
       │  Playwright WebSocket 侦听
       ▼
  browser_auth.py ───► token.txt / token_cache.json
       │                       │
       │  TokenRefresher       │ TokenManager.get()
       │  (后台守护线程)         ▼
       │                  M365Client (client.py)
       │                       │
       │                       ▼    SignalR + WebSocket
       │                  substrate.office.com (Copilot 后端)
       │
用户
  │
  ├─ CLI: outlook-copilot
  │      ↓                         OpenAI / Anthropic SDK
  ├─ API: outlook-copilot-server ─────────────────────► 第三方应用
  │      ↓
  ├─ --login: headed 首次登录
  ├─ --auto-refresh: headless 自动刷新
  └─ --session-id: 服务端默认会话（零配置多轮上下文）
```

## 认证机制

### 核心问题

Microsoft 365 Copilot 使用 `c0ab8ce9-e9a0-42e7-b064-33d422df41f1` 作为客户端 ID。这是一个**机密客户端**（confidential client），要求 `client_secret` 才能获取 token。因此无法通过标准的 OAuth 授权码流或设备代码流直接从 CLI 获取可用的 access_token。

### 可行路径

第三方应用能拿到的 token 集合：

| 客户端 ID | 类型 | Copilot 权限 | 说明 |
|-----------|------|-------------|------|
| `9199bf20-a13f-4107-85dc-02114787ef48` | 公开 (SPA) | ❌ 403 | Outlook Web 前端 |
| `4765445b-32c6-49b0-83e6-1d93765276ca` | 公开 | ❌ 403 | M365 公共客户端 |
| `c0ab8ce9-e9a0-42e7-b064-33d422df41f1` | 机密 | ✅ | **真正的 Copilot 后端（需 OBO）** |

### 浏览器中的令牌流

```
Outlook SPA                   Outlook API                      Azure AD
   │                              │                               │
   │  1. 登录获取 id_token         │                               │
   │◄─────────────────────────────│                               │
   │                              │                               │
   │  2. acquireTokenSilent()     │                               │
   │  (client: 9199bf20-...)      │                               │
   │─────────────────────────────►│                               │
   │  3. token for substrate      │                               │
   │◄─────────────────────────────│                               │
   │                              │                               │
   │  4. 调用 Copilot API         │                               │
   │  (含 9199bf20-... token)     │                               │
   │─────────────────────────────►│                               │
   │                              │  5. OBO 交换                  │
   │                              │  (client: c0ab8ce9-...)       │
   │                              │──────────────────────────────►│
   │                              │  6. Copilot access_token      │
   │                              │◄──────────────────────────────│
   │  7. 返回 access_token        │                               │
   │◄─────────────────────────────│                               │
   │                              │                               │
   │  8. 构造 WS URL: wss://...&access_token=...                  │
   │─────────────────────────────► substrate.office.com           │
```

步骤 5-6 的 OBO 交换需要 `client_secret`，只在微软后端服务器上进行，第三方无法直接调用。

### 当前方案

**主要路径：Playwright 浏览器自动化（推荐，可无人值守）**

`c0ab8ce9-...` token 只在已登录浏览器建立 Copilot WebSocket 时出现在 WS URL 的 `access_token=` 参数里，且从不落盘缓存。通过 Playwright 驱动一个持久登录的浏览器，在 Copilot 输入框输入一个字符（不发消息）即可触发 substrate WS，从 WS URL 中提取 token。

1. 首次运行 `outlook-copilot --login`（headed），弹出浏览器窗口，用户登录并勾选"保持登录"。
2. 之后用 `outlook-copilot-server --auto-refresh` 启动服务，后台线程以 headless 模式复用已登录的 profile，在 token 到期前自动抓取新 token。
3. 登录态约 90 天有效（基于 AAD KMSI 机制）；失效后重新 `--login` 即可。

关键点：
- Playwright 的 `page.on("websocket")` 在 CDP 层、导航前注册，能稳定捕获 WS URL（比页面内 JS 注入可靠）。
- Copilot 输入框是 Lexical 富文本编辑器（`#m365-chat-editor-target-element`），需用 `click()` + `keyboard.type()`（不能用 `fill()`）。
- 输入框在顶层文档而非 iframe 内（已实测确认）。

**备用路径：DevTools 手动复制**

```bash
outlook-copilot --refresh
```
F12 → Network → 筛选 ws → 刷新 → substrate 请求 → 复制 access_token。

### 自动化的可能性

| 方法 | 可行 | 原因 |
|------|------|------|
| IndexedDB (`msal.db`) | ❌ | `c0ab8ce9-...` 的 token 不在持久缓存中 |
| `localStorage` | ❌ | 只在存 `9199bf20-...` 的 token |
| MSAL `acquireTokenSilent()` | ❌ | MSAL 实例被 webpack 闭包封装，未暴露到 window |
| WebSocket 构造函数拦截（页面内 JS） | ❌ | WS 在页面加载时创建，脚本注入时已错过 |
| fetch 拦截 | ❌ | OBO 走的是后端 API，不经过前端 fetch |
| **Playwright WS 侦听** | ✅ | CDP 层捕获，时序对了；Copilot 输入单字符即触发 |
| DevTools Network 手动复制 | ✅ | 100% 可靠（手动） |

## WebSocket / SignalR 协议

### 端点

```
wss://substrate.office.com/m365Copilot/Chathub/{oid}@{tenant}
  ?chatsessionid={uuid}
  &XRoutingParameterSessionKey={uuid}
  &clientrequestid={uuid}
  &X-SessionId={uuid}
  &ConversationId={uuid}
  &access_token={jwt}
  &variants={...}
  &source="owahub"
  &product=OwaHub
  &agentHost=Bizchat.FullScreen
  &licenseType=Starter
  &isEdu=false
  &agent=none
  &scenario=owahub
```

### URL 参数说明

| 参数 | 说明 | 每次请求 |
|------|------|---------|
| `chatsessionid` | 本次 WebSocket 会话 ID | ✅ 唯一 |
| `XRoutingParameterSessionKey` | 路由密钥（同 chatsessionid） | ✅ 唯一 |
| `clientrequestid` | 客户端请求 ID（同 chatsessionid） | ✅ 唯一 |
| `X-SessionId` | 浏览器会话 ID | ❌ 固定 |
| `ConversationId` | 对话 ID（多次对话共享上下文） | ❌ 固定 |

### SignalR 消息格式

使用 JSON 协议，消息以 `\x1e`（Record Separator）分隔。

**客户端 → 服务器：**

```json
{
  "type": 4,
  "invocationId": "uuid",
  "target": "chat",
  "arguments": [{
    "source": "owahub",
    "clientCorrelationId": "hex_sid",
    "sessionId": "uuid_sid",
    "traceId": "hex_sid",
    "optionsSets": [...],
    "streamingMode": "ConciseWithPadding",
    "message": {
      "author": "user",
      "inputMethod": "Keyboard",
      "text": "用户消息",
      "messageType": "Chat"
    },
    "tone": "Magic",
    ...
  }]
}
```

**服务器 → 客户端（流式）：**

```json
{
  "type": 1,
  "target": "update",
  "arguments": [{
    "writeAtCursor": "文本增量",
    "messages": [{"text": "累计文本"}]
  }]
}
```

**服务器 → 客户端（完成）：**

```json
{"type": 3}
```

### tone 参数

| CLI 参数 | tone 值 | 功能 |
|----------|---------|------|
| `auto` | `Magic` | 默认，自动平衡 |
| `quick` | `Chat` | 短回复，快速 |
| `reasoning` | `Reasoning` | 启用深度思考 |
| `opus` | `Claude_Opus` | 启用 Claude Opus 4.8 模式 |
| `gpt-5.5` | `Gpt_5_5_Chat` | 启用 GPT-5.5 快速答复 |
| `gpt-5.6` | `Gpt_5_6_Reasoning` | 启用 GPT-5.6 Think Deeper |

### 注意事项

- `source`、`product`、`agentHost`、`scenario` 等参数在 URL 和 payload body 中**必须一致**
- `clientInfo.clientPlatform`、`clientAppName`、`clientEntrypoint` 需与 `source` 匹配
- `Origin` 请求头需设置为 `https://outlook.office.com`（对应 OWA）

## 多轮对话与上下文

### 核心发现

**M365 Copilot 后端不认客户端回放的 `messageHistory`。** 不同于标准 OpenAI API（每次发完整的 `messages` 数组），M365 后端无视 SignalR payload 中回放的历史消息。将多轮 `messages` 塞入 `messageHistory` 字段后，后端仍可能用其自身的用户信息（如 token 中的 `given_name`）回答，而非使用对话中约定的上下

### 唯一起作用的上下文机制：ConversationId

M365 后端仅在 WebSocket URL 的 `ConversationId` 相同时才维持上下文。本服务通过「会话标识」（session）映射到 M365 的 `ConversationId`：

```
客户端请求 (_get_conv_id)
   │
   ├─ 优先级 1: 请求体 session_id 字段
   ├─ 优先级 2: 请求头 X-Session-Id
   ├─ 优先级 3: 请求体 user 字段（OpenAI SDK 原生支持）
   └─ 优先级 4: 服务端 --session-id 默认值
        │
        ▼
   (session_id, timestamp) 存入 OrderedDict (TTL + LRU)
        │
        ▼
   conversation_id (固定，关联同一 session 的多次请求)
        │
        ▼
   WebSocket URL: &ConversationId={conversation_id}
```

### 服务端默认会话

`--session-id` / `OUTLOK_COPILOT_SESSION_ID` 为所有"不带自己会话标识"的请求提供兜底的默认会话，实现客户端**零配置多轮上下文**。带请求级标识的请求不受影响（进入各自独立对话）。

已验证行为：

| 场景 | 结果 |
|------|------|
| 无任何标识 + 有默认会话 | 同一 `ConversationId`，上下文保留 ✅ |
| 带独立 `session_id` + 有默认会话 | 使用请求级 `session_id`，与默认会话隔离 ✅ |
| 无任何标识 + 无默认会话 | 每次新 `ConversationId`，请求彼此隔离 ✅ |

## Token 格式

```json
{
  "aud": "https://substrate.office.com/sydney",
  "iss": "https://sts.windows.net/{tenant}/",
  "appid": "c0ab8ce9-e9a0-42e7-b064-33d422df41f1",
  "scp": "CopilotPlatformContent.Process.All CopilotPlatformFiles.Read ...",
  "exp": 1783504730,
  "iat": 1783500247
}
```

- 有效期：`exp - iat` ≈ 75 分钟
- 权限（scp）：`CopilotPlatform*.*`、`M365Chat.Read`、`sydney.readwrite`
- 不可通过 refresh token 续期（机密客户端，需要 client_secret）
- 自动刷新：通过 `--auto-refresh` / `OUTLOOK_COPILOT_AUTO_REFRESH=1` 启用，Playwright 在 token 到期前 5 分钟自动用 headless 浏览器重新抓取

## 项目配置

### .env

```ini
M365_TENANT_ID=xxx         # Azure AD 租户 ID
M365_USER_OID=xxx           # 用户 Object ID
M365_CLIENT_ID=xxx          # 客户端 ID（默认 4765445b-...）
```

### 文件路径

```
data/
├── tokens/
│   ├── token.txt               # access_token（手动提取或自动刷新写入）
│   ├── token_cache.json        # 自动缓存（含 expires_at）
│   └── rt_90day.txt            # 加密存储的 refresh token（当前不可用）
└── browser_profile/            # Playwright 持久登录浏览器 profile
```

### 环境变量一览

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `OUTLOOK_COPILOT_API_KEY` | （空） | 设置后所有 `/v1/*` 请求需 Bearer 鉴权 |
| `OUTLOOK_COPILOT_AUTO_REFRESH` | `0` | 设为 `1` 开启后台自动刷新 |
| `OUTLOOK_COPILOT_REFRESH_MARGIN` | `300` | token 到期前多少秒开始刷新 |
| `OUTLOOK_COPILOT_BROWSER_HEADLESS` | `1` | 刷新浏览器是否 headless |
| `OUTLOOK_COPILOT_HOST_URL` | 内置 | Copilot host 页面 URL |
| `OUTLOK_COPILOT_SESSION_ID` | （空） | 服务端默认会话 ID |
| `OUTLOOK_COPILOT_POOL_SIZE` | `8` | HTTP 客户端池大小 |
| `OUTLOOK_COPILOT_SESSION_MAX` | `1000` | 会话映射最大条数 |
| `OUTLOOK_COPILOT_SESSION_TTL` | `3600` | 会话映射过期时间（秒） |
| `OUTLOOK_COPILOT_MAX_BODY_BYTES` | `10MB` | 请求体大小上限 |
