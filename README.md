# Outlook Copilot 2 API

> 将 Microsoft 365 Outlook Copilot 的 SignalR WebSocket 协议包装成 OpenAI / Anthropic 兼容的 HTTP API。  
> 基于 [M365-Copilot2API](https://github.com/anomalyco/M365-Copilot2API) 修改，适配 Outlook OwaHub 端点，支持 **Claude Opus 4.7** 及 **GPT-5.5 / GPT-5.6** 模型。

---

## 特性

- **OpenAI 兼容** — `/v1/chat/completions`、`/v1/completions`（FIM）
- **Anthropic 兼容** — `/v1/messages`、`/v1/complete`
- **流式 & 非流式** — SSE 流式输出，支持 `stream=true`
- **多模型** — `auto`、`quick`、`reasoning`、`opus`、`gpt-5.5`、`gpt-5.6`
- **对话模式** — 多轮对话共享上下文
- **Tool Calls** — 搜索、代码解释器、图片生成等工具调用透传
- **CLI + API Server** — 命令行直接提问 / 启动独立 HTTP 服务

---

## 目录

- [安装](#安装)
- [快速配置](#快速配置)
- [CLI 用法](#cli-用法)
- [API Server](#api-server)
  - [端点一览](#端点一览)
  - [模型 / Tone 对照](#模型--tone-对照)
  - [OpenAI Chat 调用示例](#openai-chat-调用示例)
  - [OpenAI Completions (FIM) 调用示例](#openai-completions-fim-调用示例)
  - [Anthropic Messages 调用示例](#anthropic-messages-调用示例)
  - [Anthropic Complete 调用示例](#anthropic-complete-调用示例)
- [Token 管理](#token-管理)
- [项目结构](#项目结构)
- [致谢](#致谢)

---

## 安装

```bash
git clone https://github.com/OmaruKoma/Outlookcopilot2API.git
cd Outlookcopilot2API
pip install -e .
```

依赖：`websockets>=12`、`cryptography>=41`（Python >= 3.10）

---

## 快速配置

运行配置向导：

```bash
outlook-copilot-setup
```

向导分两步：

### 步骤 1 — 获取 OID / Tenant

1. 打开 https://outlook.cloud.microsoft 并登录
2. 按 F12 打开 DevTools → **Console**
3. 粘贴运行以下代码：

```js
(() => {const k = Object.keys(localStorage).find(k => k.includes('|refreshtoken|'));if (!k) return JSON.stringify({error:'NOT_FOUND: 请先登录 outlook.cloud.microsoft'});const parts = k.split('|');const idParts = parts[1].split('.');return JSON.stringify({oid:idParts[0], tenant:idParts[1]});})()
```

4. 将控制台输出的 JSON（包含 `oid` 和 `tenant`）粘贴回终端

### 步骤 2 — 获取 access_token

1. 保持 outlook.cloud.microsoft 打开
2. F12 → **Network** → 筛选输入 `ws`
3. 刷新页面（F5）
4. 点击列表中的 `substrate.office.com` WebSocket 请求
5. 在 URL 末尾找到 `access_token=eyJ...`，**完整复制**（约 2000+ 字符）
6. 粘贴回终端

配置完成后，系统会自动保存 `.env` 文件及 Token 到 `data/tokens/` 目录。

---

## CLI 用法

```bash
# 单次提问
outlook-copilot "你好"

# 指定模型
outlook-copilot --model opus "用 Claude Opus 回答"

# 交互模式
outlook-copilot -i

# 对话模式（多轮共享上下文）
outlook-copilot -c -i

# 非流式输出
outlook-copilot --no-stream "你好"

# 列出所有可用模型
outlook-copilot --list-models

# 更新过期 Token
outlook-copilot --refresh

# 重新运行配置向导
outlook-copilot --setup
```

---

## API Server

启动 HTTP 服务：

```bash
# 默认仅监听本机 (127.0.0.1)，安全
outlook-copilot-server --port 8000

# 对外暴露（局域网/公网）——务必同时设置 API Key 鉴权
OUTLOOK_COPILOT_API_KEY=your-secret-key outlook-copilot-server --host 0.0.0.0 --port 8000
```

首次启动需先完成配置，否则会提示运行 `outlook-copilot-setup`。

> **安全提示**：服务默认无鉴权。若绑定到非本机地址（如 `0.0.0.0`），
> 任何能访问该端口的人都能用你的 M365 token 发请求。此时请务必设置环境变量
> `OUTLOOK_COPILOT_API_KEY`，所有 `/v1/*` 请求需携带 `Authorization: Bearer <key>`。
> 未设置且绑定非 loopback 地址时，启动会打印告警。

### 可选环境变量

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `OUTLOOK_COPILOT_API_KEY` | （空） | 设置后所有 `/v1/*` 请求需 Bearer 鉴权 |
| `OUTLOOK_COPILOT_POOL_SIZE` | `8` | 客户端池大小 = 最大并发上游请求数 |
| `OUTLOOK_COPILOT_MAX_BODY_BYTES` | `10485760` | 请求体大小上限（字节） |
| `OUTLOOK_COPILOT_SESSION_MAX` | `1000` | 会话映射最大条数 |
| `OUTLOOK_COPILOT_SESSION_TTL` | `3600` | 会话映射过期时间（秒） |

### 端点一览

| 端点 | 兼容协议 | 方法 | 说明 |
|------|---------|------|------|
| `/v1/chat/completions` | OpenAI Chat | POST | 聊天补全（流式/非流式） |
| `/v1/completions` | OpenAI FIM | POST | 文本补全 / 填充中间 |
| `/v1/messages` | Anthropic Messages | POST | 消息 API |
| `/v1/complete` | Anthropic Text | POST | 文本补全 API |
| `/v1/models` | — | GET | 获取模型列表 |
| `/health` | — | GET | 健康检查 |

### 模型 / Tone 对照

| 模型 Key | Tone 值 | OpenAI 模型 ID | 说明 |
|----------|---------|----------------|------|
| `auto` | `Magic` | `gpt-5.5` | 默认，自动平衡 |
| `quick` | `Chat` | `gpt-5.5` | 快速简短回复 |
| `reasoning` | `Reasoning` | `gpt-5.5` | 深度推理模式 |
| `opus` | `Claude_Opus` | `claude-opus-4.8` | **Claude Opus 4.8**（Outlook Copilot 特有） |
| `gpt-5.5` | `Gpt_5_5_Chat` | `gpt-5.5` | **GPT-5.5 快速答复** |
| `gpt-5.6` | `Gpt_5_6_Reasoning` | `gpt-5.6` | **GPT-5.6 Think Deeper** |

> Anthropic 端点会自动映射模型名：`claude-opus-4.8` / `claude-opus-4.7` → `opus`，`gpt-5.6` → `gpt-5.6`，`gpt-5.5` / `gpt-5` / `gpt-4` → `auto`

### OpenAI Chat 调用示例

```bash
curl http://localhost:8000/v1/chat/completions \
  -H "Content-Type: application/json" \
  -d '{
    "model": "opus",
    "messages": [{"role": "user", "content": "你好"}],
    "stream": true
  }'
```

```python
import openai

client = openai.OpenAI(base_url="http://localhost:8000/v1", api_key="ignored")
response = client.chat.completions.create(
    model="opus",
    messages=[{"role": "user", "content": "你好"}],
    stream=True,
)
for chunk in response:
    if chunk.choices[0].delta.content:
        print(chunk.choices[0].delta.content, end="")
```

**非流式请求**（去掉 `stream` 参数即可）：

```python
response = client.chat.completions.create(
    model="auto",
    messages=[{"role": "user", "content": "帮我写一封邮件"}],
)
print(response.choices[0].message.content)
```

### OpenAI Completions (FIM) 调用示例

```bash
curl http://localhost:8000/v1/completions \
  -H "Content-Type: application/json" \
  -d '{
    "model": "auto",
    "prompt": "def hello():\n    ",
    "suffix": "    return greeting",
    "stream": true
  }'
```

```python
response = client.completions.create(
    model="auto",
    prompt="def hello():\n    ",
    suffix="    return greeting",
)
print(response.choices[0].text)
```

### Anthropic Messages 调用示例

```bash
curl http://localhost:8000/v1/messages \
  -H "Content-Type: application/json" \
  -d '{
    "model": "claude-opus-4.7",
    "messages": [{"role": "user", "content": "你好"}],
    "system": "你是一个助手",
    "stream": true
  }'
```

```python
import anthropic

client = anthropic.Anthropic(base_url="http://localhost:8000", api_key="ignored")
response = client.messages.create(
    model="claude-opus-4.7",
    messages=[{"role": "user", "content": "你好"}],
    system="你是一个助手",
    stream=True,
)
for event in response:
    if event.type == "content_block_delta":
        print(event.delta.text, end="")
```

**非流式请求：**

```python
response = client.messages.create(
    model="claude-opus-4.7",
    messages=[{"role": "user", "content": "你好"}],
)
print(response.content[0].text)
```

### Anthropic Complete 调用示例

```bash
curl http://localhost:8000/v1/complete \
  -H "Content-Type: application/json" \
  -d '{
    "model": "claude-opus-4.7",
    "prompt": "人类：你好\n\n助手："
  }'
```

```python
response = client.completions.create(
    model="claude-opus-4.7",
    prompt="人类：你好\n\n助手：",
)
print(response.completion)
```

### 设置会话 ID（保持上下文）

通过 `X-Session-Id` 请求头或请求体中的 `session_id` 字段，可以让多轮请求共享同一对话上下文：

```bash
curl http://localhost:8000/v1/chat/completions \
  -H "Content-Type: application/json" \
  -H "X-Session-Id: my-session-001" \
  -d '{"model":"auto","messages":[{"role":"user","content":"记住我的名字是小明"}],"stream":true}'
```

---

## Token 管理

access_token 有效期约 **75 分钟**，过期后需要更新。

### 方法 1：交互式刷新（推荐）

```bash
outlook-copilot --refresh
```

按提示从浏览器 DevTools → Network → ws 筛选 → 刷新 → 复制新 token 粘贴。

### 方法 2：从浏览器获取并直接传入

```bash
outlook-copilot --token "eyJ..."
```

### 方法 3：OAuth 授权码流程（实验性）

```bash
outlook-copilot-setup
```

此流程会尝试打开浏览器进行 OAuth 登录，需要手动复制 authorization code。

### 方法 4：Windows WAM（仅 Windows）

安装 `msal[broker]` 后，`outlook-copilot-setup` 会自动尝试 WAM broker 登录：

```bash
pip install msal[broker]
outlook-copilot-setup
```

### Token 文件位置

```
data/tokens/
├── token.txt            # 手动提取的 access_token
├── token_cache.json     # 自动缓存
└── rt_90day.txt         # 加密存储的 refresh token
```

---

## 项目结构

```
Outlookcopilot2API/
├── .env                    # 环境变量（租户 / OID / 客户端 ID）
├── .env.example            # 环境变量模板
├── pyproject.toml          # 项目元数据 + 依赖
├── README.md               # 本文件
├── data/
│   └── tokens/             # Token 存储目录
└── src/outlook_copilot/
    ├── __init__.py         # 自动加载 .env + 导出核心模块
    ├── __main__.py         # python -m outlook_copilot
    ├── auth.py             # Token 管理（获取 / 刷新 / 缓存）
    ├── client.py           # WebSocket / SignalR 客户端
    ├── cookie_store.py     # Cookie 存储
    ├── models.py           # 模型定义 + Tone 配置
    ├── payload.py          # WS URL 构建 + 消息 Payload
    ├── scripts/
    │   ├── crypto.py       # Refresh token 加解密
    │   └── setup_wizard.py # 配置向导
    └── servers/
        ├── cli.py          # CLI 入口
        └── openai.py       # OpenAI / Anthropic 兼容 HTTP API 服务
```

---

## 致谢

- 本项目基于 [M365-Copilot2API](https://github.com/anomalyco/M365-Copilot2API) 修改
- 感谢原作者 reverse engineering 了 M365 Copilot 的 SignalR 协议
- 适配 Outlook OwaHub 端点，补充 Claude Opus 4.7 模型支持
