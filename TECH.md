# 技术文档

## 概述

M365 Copilot2API 是一个代理层，将 Microsoft 365 Copilot 的私有 SignalR WebSocket 协议转换为标准的 OpenAI 兼容 HTTP API。

## 架构

```
用户
  │
  ├─ CLI: outlook-copilot
  │      ↓                         OpenAI / Anthropic SDK
  ├─ API: outlook-copilot-server ─────────────────────► 第三方应用
  │      ↓
  ├─ setup: outlook-copilot-setup
  │
  └──► TokenManager (auth.py)
         │
         ▼    access_token
     M365Client (client.py)
         │
         ▼    SignalR + WebSocket
     substrate.office.com (Coplot 后端)
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

通过浏览器 DevTools → Network 手动提取 WebSocket URL 中的 `access_token` 参数。

### 自动化的可能性

| 方法 | 可行 | 原因 |
|------|------|------|
| IndexedDB (`msal.db`) | ❌ | `c0ab8ce9-...` 的 token 不在持久缓存中 |
| `localStorage` | ❌ | 只在存 `9199bf20-...` 的 token |
| MSAL `acquireTokenSilent()` | ❌ | MSAL 实例被 webpack 闭包封装，未暴露到 window |
| WebSocket 构造函数拦截 | ❌ | WS 在页面加载时创建，脚本注入时已错过 |
| fetch 拦截 | ❌ | OBO 走的是后端 API，不经过前端 fetch |
| DevTools Network 手动复制 | ✅ | 100% 可靠 |

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
  &agent=work
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
| `opus` | `Claude_Opus` | 启用 Claude Opus 模式 |

### 注意事项

- `source`、`product`、`agentHost`、`scenario` 等参数在 URL 和 payload body 中**必须一致**
- `clientInfo.clientPlatform`、`clientAppName`、`clientEntrypoint` 需与 `source` 匹配
- `Origin` 请求头需设置为 `https://outlook.office.com`（对应 OWA）

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
- 不可通过 refresh token 续期（机密客户端）

## 项目配置

### .env

```ini
M365_TENANT_ID=xxx         # Azure AD 租户 ID
M365_USER_OID=xxx           # 用户 Object ID
M365_CLIENT_ID=xxx          # 客户端 ID（默认 4765445b-...）
```

### 文件路径

```
data/tokens/
├── token.txt               # 手动提取的 access_token
├── token_cache.json        # 自动缓存
└── rt_90day.txt            # 加密存储的 refresh token（当前不可用）
```
