from outlook_copilot.servers import openai as srv


def test_resolve_anthropic_exact_matches():
    assert srv._resolve_anthropic_model("claude-opus-4.8") == "opus"
    assert srv._resolve_anthropic_model("claude-opus-4.7") == "opus"
    assert srv._resolve_anthropic_model("gpt-5.6") == "gpt-5.6"
    assert srv._resolve_anthropic_model("gpt-5.5") == "auto"
    assert srv._resolve_anthropic_model("gpt-5") == "auto"
    assert srv._resolve_anthropic_model("gpt-4") == "auto"


def test_resolve_anthropic_prefix_match():
    # Versioned/dated Anthropic ids should map by prefix.
    assert srv._resolve_anthropic_model("claude-opus-4.7-20240101") == "opus"


def test_resolve_anthropic_unknown_defaults_auto():
    assert srv._resolve_anthropic_model("some-unknown-model") == "auto"


def test_fim_to_chat_with_suffix():
    msgs = srv.fim_to_chat("def f():\n    ", "    return x")
    assert len(msgs) == 1
    assert msgs[0]["role"] == "user"
    content = msgs[0]["content"]
    assert "def f():" in content
    assert "return x" in content
    assert "middle" in content.lower()


def test_fim_to_chat_without_suffix():
    msgs = srv.fim_to_chat("once upon a time")
    assert len(msgs) == 1
    assert msgs[0]["role"] == "user"
    assert "once upon a time" in msgs[0]["content"]
    assert "Continue writing" in msgs[0]["content"]


def test_sse_msg_shape():
    import json
    out = srv.sse_msg({"content": "hi"}, chunk_id="cid", model="m")
    assert out.startswith("data: ")
    assert out.endswith("\n\n")
    payload = json.loads(out[len("data: "):].strip())
    assert payload["id"] == "cid"
    assert payload["object"] == "chat.completion.chunk"
    assert payload["model"] == "m"
    assert payload["choices"][0]["delta"] == {"content": "hi"}


def test_sse_done_includes_done_marker_and_usage():
    import json
    out = srv.sse_done(chunk_id="cid", model="m", usage={"total_tokens": 3})
    assert out.strip().endswith("data: [DONE]")
    first = out.split("\n\n")[0]
    payload = json.loads(first[len("data: "):])
    assert payload["choices"][0]["finish_reason"] == "stop"
    assert payload["usage"] == {"total_tokens": 3}


class _StubHandler:
    """Borrows _get_conv_id / _prune_sessions without the socket machinery of
    BaseHTTPRequestHandler."""
    _get_conv_id = srv.OpenAIHandler._get_conv_id
    _prune_sessions = srv.OpenAIHandler._prune_sessions

    def __init__(self, headers=None, default_session_id=None):
        import collections
        import threading
        self.headers = headers or {}
        self.default_session_id = default_session_id
        self._session_conv = collections.OrderedDict()
        self._session_lock = threading.Lock()
        self._session_max = 1000
        self._session_ttl = 3600


def test_conv_id_none_without_session_or_default():
    h = _StubHandler()
    assert h._get_conv_id({}) is None


def test_conv_id_falls_back_to_default_session():
    h = _StubHandler(default_session_id="global")
    cid = h._get_conv_id({})
    assert cid is not None
    # Same default session -> same conversation id across requests.
    assert h._get_conv_id({}) == cid


def test_conv_id_per_request_overrides_default():
    h = _StubHandler(default_session_id="global")
    default_cid = h._get_conv_id({})
    other_cid = h._get_conv_id({"session_id": "other"})
    assert other_cid != default_cid
    # The explicit session stays stable on repeat.
    assert h._get_conv_id({"session_id": "other"}) == other_cid


def test_conv_id_header_and_user_sources():
    h = _StubHandler()
    by_header = _StubHandler(headers={"X-Session-Id": "hdr"})
    assert by_header._get_conv_id({}) is not None
    # body session_id takes precedence over header
    h2 = _StubHandler(headers={"X-Session-Id": "hdr"})
    a = h2._get_conv_id({"session_id": "body"})
    b = h2._get_conv_id({})  # falls to header now
    assert a != b
