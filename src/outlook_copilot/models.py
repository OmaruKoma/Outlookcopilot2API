import os

CLIENT_ID = os.environ.get("M365_CLIENT_ID", "4765445b-32c6-49b0-83e6-1d93765276ca")
TENANT_ID = os.environ.get("M365_TENANT_ID", "")
USER_OID = os.environ.get("M365_USER_OID", "")
SCOPE = "https://substrate.office.com/sydney/.default openid profile offline_access"

MODELS = {
    "auto":      {"tone": "Magic",       "openai_id": "gpt-5.5"},
    "quick":     {"tone": "Chat",        "openai_id": "gpt-5.5"},
    "reasoning": {"tone": "Reasoning",   "openai_id": "gpt-5.5"},
    "opus":      {"tone": "Claude_Opus", "openai_id": "claude-opus-4.8"},
    "gpt-5.5":   {"tone": "Gpt_5_5_Chat",       "openai_id": "gpt-5.5"},
    "gpt-5.6":   {"tone": "Gpt_5_6_Reasoning",   "openai_id": "gpt-5.6"},
}

TOOL_MESSAGE_TYPES = {
    "InternalSearchQuery": "search",
    "GeneratedCode": "code_interpreter",
    "GenerateGraphicArt": "generate_image",
    "TriggerPlugin": "trigger_plugin",
    "InvokeAction": "invoke_action",
}

def lookup_model(model_key):
    if model_key in MODELS:
        return MODELS[model_key]
    for v in MODELS.values():
        if v["openai_id"] == model_key:
            return v
    return MODELS["auto"]
