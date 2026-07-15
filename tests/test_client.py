import json

from outlook_copilot import client


def test_clean_citations_removes_marker():
    # Private-use citation markers (\uE000-\uF8FF) wrap "cite" spans.
    text = "answer\ue000cite\ue001source\ue002 tail"
    out = client._clean_citations(text)
    assert "cite" not in out
    assert out.startswith("answer")


def test_clean_citations_leaves_plain_text():
    text = "just normal text, no citations"
    assert client._clean_citations(text) == text


def test_clean_text_strips_control_and_trims():
    text = "  hello\x00\x07world\x1f  "
    out = client.clean_text(text)
    assert out == "helloworld"


def test_clean_text_keeps_newlines_and_tabs():
    text = "line1\nline2\tend"
    assert client.clean_text(text) == "line1\nline2\tend"


def test_clean_text_empty():
    assert client.clean_text("") == ""
    assert client.clean_text(None) == ""


def test_clean_chunk_preserves_surrounding_whitespace():
    # Unlike clean_text, clean_chunk must NOT trim, so chunk boundaries survive.
    assert client.clean_chunk(" mid ") == " mid "


def test_clean_chunk_strips_control_chars():
    assert client.clean_chunk("a\x00b\x1fc") == "abc"


def test_clean_chunk_empty():
    assert client.clean_chunk("") == ""


def test_extract_tool_call_search():
    msg = {"messageType": "InternalSearchQuery",
           "text": "search: weather today", "messageId": "m1"}
    tc = client.extract_tool_call(msg)
    assert tc["function"]["name"] == "search"
    assert json.loads(tc["function"]["arguments"]) == {"query": "weather today"}
    assert tc["id"] == "m1"


def test_extract_tool_call_search_without_prefix():
    msg = {"messageType": "InternalSearchQuery", "text": "raw query"}
    tc = client.extract_tool_call(msg)
    assert json.loads(tc["function"]["arguments"]) == {"query": "raw query"}


def test_extract_tool_call_code():
    msg = {"messageType": "GeneratedCode", "text": "print(1)"}
    tc = client.extract_tool_call(msg)
    assert tc["function"]["name"] == "code_interpreter"
    assert json.loads(tc["function"]["arguments"]) == {"code": "print(1)"}


def test_extract_tool_call_image():
    msg = {"messageType": "GenerateGraphicArt", "text": "a cat"}
    tc = client.extract_tool_call(msg)
    assert tc["function"]["name"] == "generate_image"
    assert json.loads(tc["function"]["arguments"]) == {"prompt": "a cat"}


def test_extract_tool_call_generic_input():
    msg = {"messageType": "TriggerPlugin", "text": "do-thing"}
    tc = client.extract_tool_call(msg)
    assert tc["function"]["name"] == "trigger_plugin"
    assert json.loads(tc["function"]["arguments"]) == {"input": "do-thing"}


def test_extract_tool_call_unknown_type_returns_none():
    assert client.extract_tool_call({"messageType": "Chat", "text": "hi"}) is None


def test_extract_tool_call_missing_fields_returns_none():
    assert client.extract_tool_call({"messageType": "", "text": "x"}) is None
    assert client.extract_tool_call({"messageType": "GeneratedCode", "text": ""}) is None


def test_extract_tool_call_generates_id_when_missing():
    msg = {"messageType": "GeneratedCode", "text": "x"}
    tc = client.extract_tool_call(msg)
    assert tc["id"].startswith("call_")
