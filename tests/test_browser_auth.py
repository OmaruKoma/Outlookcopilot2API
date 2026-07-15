import base64
import json
import time

import pytest

from outlook_copilot.browser_auth import (
    BrowserAuthError,
    decode_jwt_claims,
    extract_token_from_ws_url,
    is_target_ws,
    validate_token,
    _EXPECTED_AUD,
    _EXPECTED_APPID,
)


def _make_jwt(claims):
    """Build a fake unsigned JWT with the given payload claims."""
    def _seg(obj):
        raw = json.dumps(obj).encode()
        return base64.urlsafe_b64encode(raw).rstrip(b"=").decode()
    header = _seg({"alg": "none", "typ": "JWT"})
    payload = _seg(claims)
    return f"{header}.{payload}.sig"


def _valid_claims(**overrides):
    claims = {
        "aud": _EXPECTED_AUD,
        "appid": _EXPECTED_APPID,
        "exp": int(time.time()) + 3600,
    }
    claims.update(overrides)
    return claims


# --- decode_jwt_claims ---

def test_decode_jwt_claims_roundtrip():
    token = _make_jwt({"aud": "x", "exp": 123})
    claims = decode_jwt_claims(token)
    assert claims["aud"] == "x"
    assert claims["exp"] == 123


def test_decode_jwt_claims_handles_padding():
    # payloads of varying length exercise the urlsafe b64 padding logic
    for n in range(1, 20):
        token = _make_jwt({"k": "a" * n})
        assert decode_jwt_claims(token)["k"] == "a" * n


def test_decode_jwt_claims_bad_input():
    assert decode_jwt_claims("not-a-jwt") is None
    assert decode_jwt_claims("") is None
    # a dotted string whose second segment is not valid base64 JSON -> None
    assert decode_jwt_claims("aaa.!!!.bbb") is None


# --- extract_token_from_ws_url ---

def test_extract_token_from_ws_url():
    url = "wss://substrate.office.com/m365Copilot/Chathub/oid@tid?chatsessionid=x&access_token=eyJabc&variants=y"
    assert extract_token_from_ws_url(url) == "eyJabc"


def test_extract_token_from_ws_url_missing():
    url = "wss://substrate.office.com/m365Copilot/Chathub/oid@tid?chatsessionid=x"
    assert extract_token_from_ws_url(url) is None


# --- is_target_ws ---

def test_is_target_ws_true():
    url = "wss://substrate.office.com/m365Copilot/Chathub/oid@tid?access_token=z"
    assert is_target_ws(url) is True


def test_is_target_ws_false_other_hosts():
    assert is_target_ws("wss://go.trouter.teams.microsoft.com/v4/c") is False
    assert is_target_ws("wss://augloop.svc.cloud.microsoft/") is False
    # right host, wrong path
    assert is_target_ws("wss://substrate.office.com/other") is False


# --- validate_token ---

def test_validate_token_ok():
    token = _make_jwt(_valid_claims())
    exp = validate_token(token)
    assert exp > time.time()


def test_validate_token_rejects_wrong_audience():
    token = _make_jwt(_valid_claims(aud="https://graph.microsoft.com"))
    with pytest.raises(BrowserAuthError, match="audience"):
        validate_token(token)


def test_validate_token_rejects_wrong_appid():
    token = _make_jwt(_valid_claims(appid="9199bf20-a13f-4107-85dc-02114787ef48"))
    with pytest.raises(BrowserAuthError, match="appid"):
        validate_token(token)


def test_validate_token_rejects_expired():
    token = _make_jwt(_valid_claims(exp=int(time.time()) - 10))
    with pytest.raises(BrowserAuthError, match="expired"):
        validate_token(token)


def test_validate_token_rejects_non_jwt():
    with pytest.raises(BrowserAuthError, match="parseable JWT"):
        validate_token("garbage")


def test_validate_token_respects_min_seconds_left():
    # token valid for 100s, but we demand 200s of headroom -> reject
    token = _make_jwt(_valid_claims(exp=int(time.time()) + 100))
    with pytest.raises(BrowserAuthError):
        validate_token(token, min_seconds_left=200)
