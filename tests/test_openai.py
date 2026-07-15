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
