import json, os, sys, asyncio, time, uuid, logging, http.server, socketserver, threading, queue, contextlib, hmac
from collections import OrderedDict

from .. import __version__
from ..auth import TokenManager
from ..client import M365Client
from ..models import MODELS, lookup_model, TENANT_ID, USER_OID, CLIENT_ID, SCOPE

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
RT_FILE = os.path.join(BASE_DIR, "data", "tokens", "rt_90day.txt")
CACHE_FILE = os.path.join(BASE_DIR, "data", "tokens", "token_cache.json")
TOKEN_FILE = os.path.join(BASE_DIR, "data", "tokens", "token.txt")

ANTHROPIC_MODEL_MAP = {
    "claude-opus-4.8": "opus",
    "claude-opus-4.7": "opus",
    "gpt-5.6": "gpt-5.6",
    "gpt-5.5": "auto",
    "gpt-5": "auto",
    "gpt-4": "auto",
}


def _resolve_anthropic_model(anthropic_id):
    if anthropic_id in ANTHROPIC_MODEL_MAP:
        return ANTHROPIC_MODEL_MAP[anthropic_id]
    for prefix, mapped in ANTHROPIC_MODEL_MAP.items():
        if anthropic_id.startswith(prefix):
            return mapped
    return "auto"


def sse_msg(data, chunk_id=None, model="gpt-5.5"):
    if chunk_id is None:
        chunk_id = f"chatcmpl-{uuid.uuid4().hex}"
    chunk = {
        "id": chunk_id, "object": "chat.completion.chunk",
        "created": int(time.time()), "model": model,
        "choices": [{"index": 0, "delta": data, "finish_reason": None}],
    }
    return f"data: {json.dumps(chunk, ensure_ascii=False)}\n\n"


def sse_done(chunk_id=None, model="gpt-5.5", usage=None, finish_reason="stop"):
    if chunk_id is None:
        chunk_id = f"chatcmpl-{uuid.uuid4().hex}"
    chunk = {
        "id": chunk_id, "object": "chat.completion.chunk",
        "created": int(time.time()), "model": model,
        "choices": [{"index": 0, "delta": {}, "finish_reason": finish_reason}],
    }
    if usage:
        chunk["usage"] = usage
    out = f"data: {json.dumps(chunk, ensure_ascii=False)}\n\n"
    out += "data: [DONE]\n\n"
    return out


def fim_to_chat(prompt, suffix=None):
    if suffix:
        return [
            {"role": "user", "content": f"Complete the middle of the following text naturally.\n\n--- BEGIN TEXT ---\n{prompt}\n--- MIDDLE ---\n{suffix}\n--- END ---\n\nWrite only the middle part that connects the two sections."}
        ]
    return [
        {"role": "user", "content": f"Continue writing from this point:\n\n{prompt}"}
    ]


# API key auth (optional). If set, all /v1/* requests must present a
# matching "Authorization: Bearer <key>" header.
API_KEY = os.environ.get("OUTLOOK_COPILOT_API_KEY", "").strip()

# Maximum accepted request body size (bytes) to avoid unbounded reads / DoS.
MAX_BODY_BYTES = int(os.environ.get("OUTLOOK_COPILOT_MAX_BODY_BYTES", 10 * 1024 * 1024))

# Number of pooled clients == max concurrent upstream requests. Acquiring a
# client blocks when the pool is exhausted, providing natural backpressure.
POOL_SIZE = int(os.environ.get("OUTLOOK_COPILOT_POOL_SIZE", 8))

# Automatic token refresh via a headless logged-in browser (Playwright).
AUTO_REFRESH = os.environ.get("OUTLOOK_COPILOT_AUTO_REFRESH", "0").strip() in ("1", "true", "yes")
# Refresh this many seconds before the token expires.
REFRESH_MARGIN = int(os.environ.get("OUTLOOK_COPILOT_REFRESH_MARGIN", 300))
# Run the refresh browser headless (set to 0 only for debugging).
BROWSER_HEADLESS = os.environ.get("OUTLOOK_COPILOT_BROWSER_HEADLESS", "1").strip() in ("1", "true", "yes")
# How often the refresher wakes up to re-check expiry (seconds).
REFRESH_POLL_INTERVAL = int(os.environ.get("OUTLOOK_COPILOT_REFRESH_POLL", 30))
PROFILE_DIR = os.path.join(BASE_DIR, "data", "browser_profile")

# Server-side default session. The M365 Copilot backend ignores client-replayed
# messageHistory (verified empirically); the only way to keep multi-turn context
# is its own ConversationId, which we map from a session key. Many OpenAI-compatible
# clients cannot set a custom X-Session-Id header or session_id/user field, so a
# server-wide default session lets those clients share one conversation with zero
# client config. Per-request session_id / X-Session-Id / user still override it.
DEFAULT_SESSION_ID = os.environ.get("OUTLOOK_COPILOT_SESSION_ID", "").strip() or None

_SENTINEL = object()


class _BackgroundLoop:
    """A dedicated asyncio event loop running in a daemon thread. Handler
    threads submit coroutines to it via thread-safe bridges, so many HTTP
    requests can be served concurrently against one shared loop."""

    def __init__(self):
        self._loop = asyncio.new_event_loop()
        self._thread = threading.Thread(target=self._run, name="asyncio-loop", daemon=True)
        self._thread.start()

    def _run(self):
        asyncio.set_event_loop(self._loop)
        self._loop.run_forever()

    def run(self, coro):
        """Run a coroutine to completion on the loop, blocking the caller
        thread until it finishes; exceptions propagate."""
        return asyncio.run_coroutine_threadsafe(coro, self._loop).result()

    def iterate(self, agen_factory, queue_maxsize=64):
        """Drive an async generator on the loop and yield its items on the
        calling thread. `agen_factory` is a zero-arg callable returning an
        async generator (deferred so it is created inside the loop)."""
        q = queue.Queue(maxsize=queue_maxsize)

        async def _pump():
            try:
                agen = agen_factory()
                async for item in agen:
                    q.put(item)
            except Exception as e:  # noqa: BLE001 - forward to consumer thread
                q.put((_SENTINEL, e))
            finally:
                q.put(_SENTINEL)

        asyncio.run_coroutine_threadsafe(_pump(), self._loop)
        while True:
            item = q.get()
            if item is _SENTINEL:
                return
            if isinstance(item, tuple) and len(item) == 2 and item[0] is _SENTINEL:
                raise item[1]
            yield item


class ClientPool:
    """Fixed-size pool of M365Client instances sharing one TokenManager.
    Each client owns its own WebSocket, so checked-out clients run
    independently and concurrently."""

    def __init__(self, size, factory):
        self._q = queue.Queue()
        for _ in range(size):
            self._q.put(factory())

    @contextlib.contextmanager
    def acquire(self):
        client = self._q.get()
        try:
            yield client
        finally:
            self._q.put(client)


class TokenRefresher:
    """Refreshes the access_token by driving a logged-in browser. Runs its own
    daemon thread (Playwright's sync API cannot share a thread with an asyncio
    loop) and refreshes proactively before expiry. A lock ensures only one
    browser instance runs at a time, and also lets on-demand callers (the
    TokenManager provider hook) piggyback on the same serialized refresh."""

    def __init__(self, token_file, cache_file, profile_dir):
        self.token_file = token_file
        self.cache_file = cache_file
        self.profile_dir = profile_dir
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread = None

    def _current_exp(self):
        """Return the current token's exp (unix ts), or 0 if none/invalid."""
        try:
            with open(self.cache_file) as f:
                return float(json.load(f).get("expires_at", 0))
        except Exception:
            pass
        try:
            with open(self.token_file) as f:
                token = f.read().strip()
            claims = _decode_jwt(token)
            return float(claims.get("exp", 0)) if claims else 0
        except Exception:
            return 0

    def refresh_now(self):
        """Synchronously refresh the token. Serialized via the lock so that
        concurrent callers wait for the in-flight refresh instead of launching
        multiple browsers. Returns the fresh access_token, or None on failure.

        If another thread refreshed while this caller waited on the lock, the
        already-fresh token is returned without launching a second browser."""
        with self._lock:
            # Another caller may have just refreshed while we waited; reuse it.
            exp = self._current_exp()
            if exp - time.time() > REFRESH_MARGIN:
                try:
                    with open(self.token_file) as f:
                        existing = f.read().strip()
                    if existing:
                        return existing
                except Exception:
                    pass
            from ..browser_auth import fetch_token_blocking
            token, _ = fetch_token_blocking(
                self.profile_dir, self.token_file, self.cache_file,
                headless=BROWSER_HEADLESS,
            )
            return token

    def _loop(self):
        while not self._stop.is_set():
            try:
                exp = self._current_exp()
                now = time.time()
                if exp - now <= REFRESH_MARGIN:
                    logging.info("token near expiry; refreshing via browser")
                    self.refresh_now()
            except Exception:
                logging.exception("background token refresh failed")
            self._stop.wait(REFRESH_POLL_INTERVAL)

    def start(self):
        self._thread = threading.Thread(target=self._loop, name="token-refresher", daemon=True)
        self._thread.start()


def _decode_jwt(token):
    try:
        import base64
        parts = token.split(".")
        padded = parts[1] + "=" * (-len(parts[1]) % 4)
        return json.loads(base64.urlsafe_b64decode(padded))
    except Exception:
        return None


_bg = None
_pool = None
_tm = None
_refresher = None


def _init_runtime():
    global _bg, _pool, _tm, _refresher
    _bg = _BackgroundLoop()
    _tm = TokenManager(TENANT_ID, CLIENT_ID, SCOPE, RT_FILE, CACHE_FILE, TOKEN_FILE)
    if AUTO_REFRESH:
        _refresher = TokenRefresher(TOKEN_FILE, CACHE_FILE, PROFILE_DIR)
        # On-demand hook: if a request finds the token expired before the
        # background thread refreshed it, refresh inline (serialized).
        _tm.token_provider = _refresher.refresh_now
    _pool = ClientPool(POOL_SIZE, lambda: M365Client(_tm))


class OpenAIHandler(http.server.BaseHTTPRequestHandler):
    # session_id -> (conversation_id, last_seen_ts). Shared across handler
    # threads, so guarded by a lock. Bounded by size + TTL to prevent the
    # map from growing without limit on a long-running server.
    _session_conv = OrderedDict()
    _session_lock = threading.Lock()
    _session_max = int(os.environ.get("OUTLOOK_COPILOT_SESSION_MAX", 1000))
    _session_ttl = int(os.environ.get("OUTLOOK_COPILOT_SESSION_TTL", 3600))
    # Server-wide default session key; None means requests are isolated unless
    # they carry their own session identifier. Set via --session-id / env.
    default_session_id = DEFAULT_SESSION_ID

    def _get_conv_id(self, req):
        # Per-request identifiers take precedence; fall back to the server-wide
        # default session so clients that cannot set headers/fields still share
        # one conversation.
        sid = (req.get("session_id") or self.headers.get("X-Session-Id")
               or req.get("user") or self.default_session_id)
        if not sid:
            return None
        now = time.time()
        with self._session_lock:
            self._prune_sessions(now)
            entry = self._session_conv.get(sid)
            if entry:
                conv_id = entry[0]
            else:
                conv_id = uuid.uuid4().hex
            self._session_conv[sid] = (conv_id, now)
            self._session_conv.move_to_end(sid)
            # enforce hard cap (evict oldest)
            while len(self._session_conv) > self._session_max:
                self._session_conv.popitem(last=False)
        return conv_id

    @classmethod
    def _prune_sessions(cls, now):
        """Drop expired entries. Caller must hold _session_lock."""
        ttl = cls._session_ttl
        expired = [k for k, (_, ts) in cls._session_conv.items() if now - ts > ttl]
        for k in expired:
            cls._session_conv.pop(k, None)

    def log_message(self, format, *args):
        logging.info(f"{self.client_address[0]} - {format % args}")

    def _send_json(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _send_error(self, code, msg):
        self._send_json(code, {"error": {"message": msg, "type": "error", "code": code}})

    def _read_body(self):
        length = int(self.headers.get("Content-Length", 0))
        if length > MAX_BODY_BYTES:
            raise ValueError(f"Request body too large (max {MAX_BODY_BYTES} bytes)")
        return json.loads(self.rfile.read(length)) if length else {}

    def _authorized(self):
        """Return True if the request may proceed. When API_KEY is unset,
        the server is open (backward compatible). When set, a matching
        Bearer token is required."""
        if not API_KEY:
            return True
        auth = self.headers.get("Authorization", "")
        expected = f"Bearer {API_KEY}"
        # constant-time compare to avoid timing leaks
        return hmac.compare_digest(auth, expected)

    def _parse_params(self, req):
        model = req.get("model", "auto")
        messages = req.get("messages", [])
        stream = bool(req.get("stream", False))
        cfg = lookup_model(model)
        if not cfg:
            return self._send_error(400, f"Unknown model: {model}")
        return model, messages, stream, cfg

    def do_GET(self):
        if self.path == "/v1/models" and not self._authorized():
            self._send_error(401, "Unauthorized")
            return
        if self.path == "/v1/models":
            models = [
                {"id": v["openai_id"], "object": "model", "created": 1700000000, "owned_by": "microsoft"}
                for v in MODELS.values()
            ]
            self._send_json(200, {"object": "list", "data": models})
        elif self.path == "/health":
            self.send_response(200)
            self.send_header("Content-Type", "text/plain")
            self.end_headers()
            self.wfile.write(b"OK")
        else:
            self._send_error(404, "Not found")

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Session-Id")
        self.end_headers()

    def do_POST(self):
        path = self.path.rstrip("/")
        if not self._authorized():
            self._send_error(401, "Unauthorized")
            return
        length = int(self.headers.get("Content-Length", 0))
        if length > MAX_BODY_BYTES:
            self._send_error(413, "Request body too large")
            return
        try:
            req = self._read_body()
        except Exception as e:
            self._send_error(400, f"Invalid JSON: {e}")
            return
        if path == "/v1/chat/completions":
            self._handle_chat(req)
        elif path == "/v1/completions":
            self._handle_completions(req)
        elif path == "/v1/messages":
            self._handle_anthropic_messages(req)
        elif path == "/v1/complete":
            self._handle_anthropic_complete(req)
        else:
            self._send_error(404, f"Not found: {self.path}")

    def _handle_chat(self, req):
        parsed = self._parse_params(req)
        if parsed is None:
            return
        model, messages, stream, cfg = parsed
        conv_id = self._get_conv_id(req)
        with _pool.acquire() as client:
            if stream:
                self._stream_chat(messages, cfg, client, conv_id)
            else:
                self._non_stream_chat(messages, cfg, client, conv_id)

    def _write_sse(self, data):
        try:
            self.wfile.write(data.encode())
        except (ConnectionAbortedError, ConnectionResetError, BrokenPipeError, OSError):
            pass

    def _stream_chat(self, messages, cfg, client, conv_id=None):
        chunk_id = f"chatcmpl-{uuid.uuid4().hex}"
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Connection", "close")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        openai_model = cfg["openai_id"]
        tone = cfg["tone"]
        has_content = False
        full_text = ""
        try:
            for chunk, is_final in _bg.iterate(
                lambda: client.chat_conversation_stream_gen(messages, tone, conversation_id=conv_id)
            ):
                if is_final:
                    break
                if not chunk:
                    continue
                full_text += chunk
                if not has_content:
                    self._write_sse(sse_msg(
                        {"role": "assistant", "content": chunk}, chunk_id, openai_model))
                    has_content = True
                else:
                    self._write_sse(sse_msg({"content": chunk}, chunk_id, openai_model))
            prompt_str = str(messages)
            usage = {
                "prompt_tokens": len(prompt_str.split()),
                "completion_tokens": len(full_text.split()),
                "total_tokens": len(prompt_str.split()) + len(full_text.split()),
            }
            self._write_sse(sse_done(chunk_id, openai_model, usage))
        except Exception:
            logging.exception("stream chat failed")
            err = {"id": chunk_id, "object": "chat.completion.chunk",
                   "created": int(time.time()), "model": openai_model,
                   "choices": [{"index": 0, "delta": {"content": "Error: internal server error"},
                                "finish_reason": "stop"}]}
            self._write_sse(f"data: {json.dumps(err)}\n\n")
            self._write_sse("data: [DONE]\n\n")

    def _non_stream_chat(self, messages, cfg, client, conv_id=None):
        openai_model = cfg["openai_id"]
        tone = cfg["tone"]
        try:
            result_text, tool_calls, finish_reason = _bg.run(
                client.chat_conversation(messages, tone, conversation_id=conv_id)
            )
        except Exception:
            logging.exception("non-stream chat failed")
            self._send_error(500, "internal server error")
            return
        msg = {"role": "assistant", "content": result_text if result_text else None}
        if tool_calls:
            msg["tool_calls"] = tool_calls
            msg["content"] = None
        prompt_str = str(messages)
        response = {
            "id": f"chatcmpl-{uuid.uuid4().hex}",
            "object": "chat.completion",
            "created": int(time.time()),
            "model": openai_model,
            "choices": [{"index": 0, "message": msg, "finish_reason": finish_reason}],
            "usage": {
                "prompt_tokens": len(prompt_str.split()),
                "completion_tokens": len((result_text or "").split()),
                "total_tokens": len(prompt_str.split()) + len((result_text or "").split()),
            },
        }
        self._send_json(200, response)

    def _handle_completions(self, req):
        model = req.get("model", "auto")
        prompt = req.get("prompt", "")
        suffix = req.get("suffix", None)
        stream = bool(req.get("stream", False))
        cfg = lookup_model(model)
        if not cfg:
            self._send_error(400, f"Unknown model: {model}")
            return
        messages = fim_to_chat(prompt, suffix)
        with _pool.acquire() as client:
            if stream:
                self._stream_completions(messages, cfg, client)
            else:
                self._non_stream_completions(messages, cfg, client)

    def _stream_completions(self, messages, cfg, client):
        openai_model = cfg["openai_id"]
        tone = cfg["tone"]
        comp_id = f"cmpl-{uuid.uuid4().hex}"
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Connection", "close")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        try:
            for chunk, is_final in _bg.iterate(
                lambda: client.chat_conversation_stream_gen(messages, tone)
            ):
                if is_final:
                    break
                cdata = {
                    "id": comp_id, "object": "text_completion",
                    "created": int(time.time()), "model": openai_model,
                    "choices": [{"index": 0, "text": chunk, "finish_reason": None, "logprobs": None}],
                }
                self._write_sse(f"data: {json.dumps(cdata, ensure_ascii=False)}\n\n")
            done = {
                "id": comp_id, "object": "text_completion",
                "created": int(time.time()), "model": openai_model,
                "choices": [{"index": 0, "text": "", "finish_reason": "stop", "logprobs": None}],
            }
            self._write_sse(f"data: {json.dumps(done, ensure_ascii=False)}\n\n")
            self._write_sse("data: [DONE]\n\n")
        except Exception:
            logging.exception("stream completions failed")
            err = {"id": comp_id, "object": "text_completion",
                   "created": int(time.time()), "model": openai_model,
                   "choices": [{"index": 0, "text": "Error: internal server error",
                                "finish_reason": "stop", "logprobs": None}]}
            self._write_sse(f"data: {json.dumps(err)}\n\n")
            self._write_sse("data: [DONE]\n\n")

    def _non_stream_completions(self, messages, cfg, client):
        openai_model = cfg["openai_id"]
        tone = cfg["tone"]
        try:
            result_text, _, _ = _bg.run(client.chat_conversation(messages, tone))
        except Exception:
            logging.exception("non-stream completions failed")
            self._send_error(500, "internal server error")
            return
        response = {
            "id": f"cmpl-{uuid.uuid4().hex}",
            "object": "text_completion",
            "created": int(time.time()),
            "model": openai_model,
            "choices": [{"index": 0, "text": result_text, "finish_reason": "stop", "logprobs": None}],
            "usage": {
                "prompt_tokens": len(str(messages).split()),
                "completion_tokens": len(result_text.split()),
                "total_tokens": len(str(messages).split()) + len(result_text.split()),
            },
        }
        self._send_json(200, response)

    def _handle_anthropic_messages(self, req):
        model = req.get("model", "claude-opus-4.7")
        messages = req.get("messages", [])
        system_prompt = req.get("system", "")
        stream = bool(req.get("stream", False))
        mapped = _resolve_anthropic_model(model)
        cfg = lookup_model(mapped)
        chat_messages = []
        if system_prompt:
            chat_messages.append({"role": "system", "content": system_prompt})
        for m in messages:
            role = m.get("role", "")
            content = m.get("content", "")
            if isinstance(content, list):
                texts = [p.get("text", "") for p in content if p.get("type") == "text"]
                content = " ".join(texts)
            chat_messages.append({"role": role, "content": content})
        with _pool.acquire() as client:
            if stream:
                self._anthropic_stream_messages(chat_messages, cfg, client, model)
            else:
                self._anthropic_non_stream_messages(chat_messages, cfg, client, model)

    def _anthropic_stream_messages(self, chat_messages, cfg, client, anthropic_model):
        tone = cfg["tone"]
        msg_id = f"msg_{uuid.uuid4().hex}"
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Connection", "close")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        full_text = ""
        try:
            header = {
                "type": "message_start",
                "message": {
                    "id": msg_id, "type": "message", "role": "assistant",
                    "content": [], "model": anthropic_model,
                    "stop_reason": None, "stop_sequence": None,
                    "usage": {"input_tokens": len(str(chat_messages).split()), "output_tokens": 0},
                },
            }
            self._write_sse(f"event: message_start\ndata: {json.dumps(header, ensure_ascii=False)}\n\n")
            cb_start = {"type": "content_block_start", "index": 0, "content_block": {"type": "text", "text": ""}}
            self._write_sse(f"event: content_block_start\ndata: {json.dumps(cb_start, ensure_ascii=False)}\n\n")
            for chunk, is_final in _bg.iterate(
                lambda: client.chat_conversation_stream_gen(chat_messages, tone)
            ):
                if is_final:
                    break
                if not chunk:
                    continue
                full_text += chunk
                delta = {"type": "content_block_delta", "index": 0, "delta": {"type": "text_delta", "text": chunk}}
                self._write_sse(f"event: content_block_delta\ndata: {json.dumps(delta, ensure_ascii=False)}\n\n")
            cb_stop = {"type": "content_block_stop", "index": 0}
            self._write_sse(f"event: content_block_stop\ndata: {json.dumps(cb_stop, ensure_ascii=False)}\n\n")
            msg_delta = {
                "type": "message_delta",
                "delta": {"stop_reason": "end_turn", "stop_sequence": None},
                "usage": {"output_tokens": len(full_text.split())},
            }
            self._write_sse(f"event: message_delta\ndata: {json.dumps(msg_delta, ensure_ascii=False)}\n\n")
            msg_stop = {"type": "message_stop"}
            self._write_sse(f"event: message_stop\ndata: {json.dumps(msg_stop, ensure_ascii=False)}\n\n")
        except Exception:
            logging.exception("anthropic stream failed")
            self._write_sse(f"event: error\ndata: {json.dumps({'type': 'error', 'error': {'type': 'server_error', 'message': 'internal server error'}})}\n\n")

    def _anthropic_non_stream_messages(self, chat_messages, cfg, client, anthropic_model):
        tone = cfg["tone"]
        try:
            result_text, tool_calls, finish_reason = _bg.run(
                client.chat_conversation(chat_messages, tone)
            )
        except Exception:
            logging.exception("anthropic non-stream failed")
            self._send_error(500, "internal server error")
            return
        stop_reason = {"tool_calls": "tool_use", "stop": "end_turn"}.get(finish_reason or "stop", "end_turn")
        response = {
            "id": f"msg_{uuid.uuid4().hex}",
            "type": "message", "role": "assistant",
            "content": [{"type": "text", "text": result_text or ""}],
            "model": anthropic_model,
            "stop_reason": stop_reason,
            "stop_sequence": None,
            "usage": {
                "input_tokens": len(str(chat_messages).split()),
                "output_tokens": len((result_text or "").split()),
            },
        }
        if tool_calls:
            for tc in tool_calls:
                response["content"].append({
                    "type": "tool_use",
                    "id": tc.get("id", f"tu_{uuid.uuid4().hex}"),
                    "name": tc["function"]["name"],
                    "input": json.loads(tc["function"]["arguments"]),
                })
        self._send_json(200, response)

    def _handle_anthropic_complete(self, req):
        model = req.get("model", "claude-opus-4.7")
        prompt = req.get("prompt", "")
        stream = bool(req.get("stream", False))
        stop_sequences = req.get("stop_sequences", [])
        mapped = _resolve_anthropic_model(model)
        cfg = lookup_model(mapped)
        messages = fim_to_chat(prompt)
        try:
            with _pool.acquire() as client:
                result_text, _, _ = _bg.run(
                    client.chat_conversation(messages, cfg["tone"])
                )
        except Exception:
            logging.exception("anthropic complete failed")
            self._send_error(500, "internal server error")
            return
        response = {
            "completion": result_text,
            "stop_reason": "stop_sequence" if any(s in result_text for s in stop_sequences) else "end_turn",
            "model": model,
            "stop": None,
            "log_id": f"cmpl_{uuid.uuid4().hex}",
        }
        self._send_json(200, response)


class ThreadedServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True
    allow_reuse_address = True


def main():
    import argparse
    parser = argparse.ArgumentParser(description=f"Outlook Copilot API Server v{__version__}")
    parser.add_argument("--port", type=int, default=8000, help="listen port")
    parser.add_argument("--host", type=str, default="127.0.0.1", help="bind address")
    parser.add_argument("--setup", action="store_true", help="first time setup")
    parser.add_argument("--login", action="store_true",
                        help="open a browser to sign in and capture the first token (headed)")
    parser.add_argument("--auto-refresh", action="store_true",
                        help="auto-refresh the token via a headless browser before expiry")
    parser.add_argument("--session-id", nargs="?", const="__auto__", default=None,
                        metavar="ID",
                        help="share one conversation across requests that carry no "
                             "session identifier. Pass an ID for a stable session, or "
                             "omit the value to generate a random per-run session.")
    args = parser.parse_args()

    global AUTO_REFRESH, DEFAULT_SESSION_ID
    if args.auto_refresh:
        AUTO_REFRESH = True
    if args.session_id is not None:
        DEFAULT_SESSION_ID = uuid.uuid4().hex if args.session_id == "__auto__" else args.session_id
        OpenAIHandler.default_session_id = DEFAULT_SESSION_ID

    if not TENANT_ID or not USER_OID:
        print("Error: M365_TENANT_ID and M365_USER_OID not configured")
        print("Run: outlook-copilot-setup")
        sys.exit(1)

    os.makedirs(os.path.dirname(RT_FILE), exist_ok=True)
    tm = TokenManager(TENANT_ID, CLIENT_ID, SCOPE, RT_FILE, CACHE_FILE, TOKEN_FILE)

    if args.setup:
        from ..scripts.setup_wizard import main as setup_main
        setup_main()
        return

    if args.login:
        from ..browser_auth import fetch_token_blocking, BrowserAuthError
        try:
            fetch_token_blocking(PROFILE_DIR, TOKEN_FILE, CACHE_FILE, headless=False)
            print("登录并抓取首个 token 成功。现在可用 --auto-refresh 或 "
                  "OUTLOOK_COPILOT_AUTO_REFRESH=1 启动自动刷新。")
        except BrowserAuthError as e:
            print(f"登录失败: {e}")
            sys.exit(1)
        return

    token_ok = False
    try:
        tm.get()
        token_ok = True
        print("Token OK")
    except Exception as e:
        if AUTO_REFRESH:
            print(f"No valid token yet ({e}); will fetch via browser on startup.")
        else:
            print(f"Token failed: {e}")
            sys.exit(1)

    _init_runtime()

    if AUTO_REFRESH and _refresher is not None:
        if not token_ok:
            # No usable token yet: fetch one synchronously before serving.
            try:
                _refresher.refresh_now()
                print("Token fetched via browser.")
            except Exception as e:
                print(f"Initial browser token fetch failed: {e}")
                print("Run: outlook-copilot --login  (headed sign-in)")
                sys.exit(1)
        _refresher.start()
        print(f"  auto-refresh: on (margin {REFRESH_MARGIN}s, headless={BROWSER_HEADLESS})")

    if not API_KEY and args.host not in ("127.0.0.1", "localhost", "::1"):
        print("WARNING: binding to a non-loopback address without authentication.")
        print("         Anyone who can reach this port can use your M365 token.")
        print("         Set OUTLOOK_COPILOT_API_KEY to require a Bearer token.")

    server = ThreadedServer((args.host, args.port), OpenAIHandler)
    print(f"Outlook Copilot API Server v{__version__}")
    print(f"  http://{args.host}:{args.port}")
    print(f"  POST /v1/chat/completions  (OpenAI)")
    print(f"  POST /v1/completions        (OpenAI FIM)")
    print(f"  POST /v1/messages           (Anthropic)")
    print(f"  POST /v1/complete           (Anthropic FIM)")
    print(f"  GET  /v1/models             (model list)")
    print(f"  auth: {'Bearer token required' if API_KEY else 'open (no API key set)'}")
    if OpenAIHandler.default_session_id:
        print(f"  default session: {OpenAIHandler.default_session_id} "
              f"(requests without a session id share one conversation)")
    else:
        print(f"  default session: off (requests are isolated unless they carry a session id)")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down...")
        server.shutdown()
