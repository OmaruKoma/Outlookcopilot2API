import json

import pytest

from outlook_copilot import payload


@pytest.fixture(autouse=True)
def _set_ids(monkeypatch):
    # payload.py binds USER_OID / TENANT_ID at import time, so patch the
    # module-level names the functions actually read.
    monkeypatch.setattr(payload, "USER_OID", "oid123")
    monkeypatch.setattr(payload, "TENANT_ID", "tenant456")


def test_build_url_contains_required_params():
    url, hex_sid, uuid_sid = payload.build_url("tok")
    assert url.startswith(
        "wss://substrate.office.com/m365Copilot/Chathub/oid123@tenant456"
    )
    for key in ("chatsessionid=", "XRoutingParameterSessionKey=",
                "clientrequestid=", "X-SessionId=", "access_token=tok",
                "variants=", "scenario=owahub"):
        assert key in url
    # session ids share the same hex source
    assert len(hex_sid) == 32
    assert uuid_sid.replace("-", "") == hex_sid


def test_build_url_conversation_id_optional():
    url_no, _, _ = payload.build_url("tok")
    assert "ConversationId=" not in url_no
    url_yes, _, _ = payload.build_url("tok", conversation_id="conv-1")
    assert "ConversationId=conv-1" in url_yes


def test_build_url_requires_ids(monkeypatch):
    monkeypatch.setattr(payload, "USER_OID", "")
    with pytest.raises(ValueError):
        payload.build_url("tok")


def test_build_url_custom_x_session_id():
    url, _, uuid_sid = payload.build_url("tok", x_session_id="custom-sid")
    assert uuid_sid == "custom-sid"
    assert "X-SessionId=custom-sid" in url


def test_build_payload_shape():
    raw = payload.build_payload("hex", "uuid", "hello", tone="Magic")
    p = json.loads(raw)
    assert p["type"] == 4
    assert p["target"] == "chat"
    arg = p["arguments"][0]
    assert arg["tone"] == "Magic"
    assert arg["message"]["text"] == "hello"
    assert arg["clientCorrelationId"] == "hex"
    assert arg["sessionId"] == "uuid"


def test_conversation_payload_uses_last_message_as_current():
    messages = [
        {"role": "user", "content": "first"},
        {"role": "assistant", "content": "reply"},
        {"role": "user", "content": "second"},
    ]
    p = json.loads(payload.build_conversation_payload("hex", "uuid", messages))
    arg = p["arguments"][0]
    assert arg["message"]["text"] == "second"
    history = arg["messageHistory"]
    assert [m["text"] for m in history] == ["first", "reply"]
    assert history[0]["author"] == "user"
    assert history[1]["author"] == "bot"


def test_conversation_payload_empty_user_history_not_backfilled():
    # Regression: an empty historical user message must NOT be replaced with
    # the current message text (previous `content or last_text` bug).
    messages = [
        {"role": "user", "content": ""},
        {"role": "user", "content": "current"},
    ]
    p = json.loads(payload.build_conversation_payload("hex", "uuid", messages))
    arg = p["arguments"][0]
    assert arg["message"]["text"] == "current"
    history = arg["messageHistory"]
    assert len(history) == 1
    assert history[0]["text"] == ""


def test_conversation_payload_list_content_flattened():
    messages = [
        {"role": "user", "content": [
            {"type": "text", "text": "a"},
            {"type": "image", "url": "ignored"},
            {"type": "text", "text": "b"},
        ]},
        {"role": "user", "content": "now"},
    ]
    p = json.loads(payload.build_conversation_payload("hex", "uuid", messages))
    history = p["arguments"][0]["messageHistory"]
    assert history[0]["text"] == "a b"


def test_conversation_payload_no_history_key_when_single_message():
    messages = [{"role": "user", "content": "only"}]
    p = json.loads(payload.build_conversation_payload("hex", "uuid", messages))
    assert "messageHistory" not in p["arguments"][0]


def test_conversation_payload_tool_role():
    messages = [
        {"role": "tool", "content": "result-data"},
        {"role": "user", "content": "now"},
    ]
    p = json.loads(payload.build_conversation_payload("hex", "uuid", messages))
    history = p["arguments"][0]["messageHistory"]
    assert history[0]["text"] == "[Tool result: result-data]"
    assert history[0]["author"] == "user"
