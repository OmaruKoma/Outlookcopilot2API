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


_bg = None
_pool = None
_tm = None


def _init_runtime():
    global _bg, _pool, _tm
    _bg = _BackgroundLoop()
    _tm = TokenManager(TENANT_ID, CLIENT_ID, SCOPE, RT_FILE, CACHE_FILE, TOKEN_FILE)
    _pool = ClientPool(POOL_SIZE, lambda: M365Client(_tm))


class OpenAIHandler(http.server.BaseHTTPRequestHandler):
    # session_id -> (conversation_id, last_seen_ts). Shared across handler
    # threads, so guarded by a lock. Bounded by size + TTL to prevent the
    # map from growing without limit on a long-running server.
    _session_conv = OrderedDict()
    _session_lock = threading.Lock()
    _session_max = int(os.environ.get("OUTLOOK_COPILOT_SESSION_MAX", 1000))
    _session_ttl = int(os.environ.get("OUTLOOK_COPILOT_SESSION_TTL", 3600))

    def _get_conv_id(self, req):
        sid = req.get("session_id") or self.headers.get("X-Session-Id") or req.get("user")
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
    args = parser.parse_args()

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

    if not os.path.exists(RT_FILE) and not os.path.exists(TOKEN_FILE):
        print("First time: outlook-copilot-setup")
        sys.exit(1)

    try:
        tm.get()
        print("Token OK")
    except Exception as e:
        print(f"Token failed: {e}")
        sys.exit(1)

    _init_runtime()

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
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down...")
        server.shutdown()
