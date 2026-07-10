import json, uuid

from .models import USER_OID, TENANT_ID

LOCAL_TZ_OFFSET = 8
LOCAL_TZ_NAME = "Asia/Hong_Kong"
LOCAL_LOCALE = "zh-tw"

OUTLOOK_OPTIONS_SETS = [
    "enterprise_flux_web",
    "enterprise_flux_work",
    "enable_request_response_interstitials",
    "enterprise_flux_image_v1",
    "enterprise_toolbox_with_skdsstore_search_message_extensions",
    "enable_ME_auth_interstitial",
    "enable_confirmation_interstitial",
    "enable_plugin_auth_interstitial",
    "enable_response_action_processing",
    "enterprise_pagination_support",
    "search_result_progress_messages_with_search_queries",
    "flux_v3_gptv_enable_upload_multi_image_in_turn_wo_ch",
    "rich_responses",
    "gptvnorm2048",
    "enterprise_flux_work_code_interpreter",
    "cwc_code_interpreter_citation_fix",
    "code_interpreter_interactive_charts",
    "enterprise_code_interpreter_citation_fix",
    "cwc_code_interpreter_interactive_charts_inline_image",
    "code_interpreter_matplotlib_patching",
    "enable_batch_token_processing",
    "disable_cea_message_listener",
    "enable_selective_url_redaction",
    "update_memory_plugin",
    "add_custom_instructions",
    "agent_recommendations",
    "enable_gg_gpt",
    "enable_inferred_memory_read",
    "update_textdoc_response_after_streaming",
    "deepleo_networking_timeout_10minutes_canmore",
    "flux_v3_references",
    "flux_v3_references_entities",
    "flux_v3_image_gen_enable_dimensions",
    "flux_v3_image_gen_enable_non_watermarked_storage",
    "flux_v3_image_gen_enable_icon_dimensions",
    "flux_v3_image_gen_enable_system_text_with_params",
    "flux_v3_image_gen_enable_designer_dimensions_meta_prompting_in_system_prompts",
    "flux_v3_image_gen_enable_story",
    "pages_citations_multiturn",
]

ALLOWED_MSG_TYPES = [
    "Chat", "Suggestion", "InternalSearchQuery", "Disengaged",
    "InternalLoaderMessage", "Progress", "GeneratedCode",
    "RenderCardRequest", "AdsQuery", "SemanticSerp",
    "GenerateContentQuery", "GenerateGraphicArt", "SearchQuery",
    "ConfirmationCard", "AuthError", "DeveloperLogs",
    "TriggerPlugin", "HintInvocation", "MemoryUpdate",
    "EndOfRequest", "TriggerConfirmation",
    "ResumeInvokeAction", "ResumeUserInputRequest",
    "TriggerUserInputRequest", "EscapeHatch",
    "TriggerPluginAuth", "ResumePluginAuth",
    "ReferencesListComplete", "CompleteExtension",
    "TriggerExtension", "SwitchRespondingEndpoint",
]

VARIANTS = (
    "EnableMcpServerWidgets,feature.EnableLuForChatCIQ,feature.enableChatCIQPlugin,"
    "EnableRequestPlugins,feature.EnableSensitivityLabels,EnableUnsupportedUrlDetector,"
    "feature.IsCustomEngineCopilotEnabled,feature.bizchatfluxv3,"
    "feature.enablechatpages,feature.turnOnWorkTabRecommendation,"
    "feature.turnOnDARecommendation,feature.IsStreamingModeInChatRequestEnabled,"
    "IncludeSourceAttributionsConcise,SkipPublishEmptyMessage,"
    "feature.EnableDeduplicatingSourceAttributions,feature.IsCitationsReferencesOutputEnabled,"
    "feature.enableDeltaStreamingForReferences,feature.enableIncludeReferencesInDeltaResponse,"
    "feature.enablereferencesforagents,Enable3PActionProgressMessages,"
    "feature.enableClientWebRtc,feature.EnableMeetingRecapOfSeriesMeetingWithCiq,"
    "feature.EnableReferencesListCompleteSignal,feature.StorageMessageSplitDisabled,"
    "feature.EnableCuaTakeControlApi,SingletonEnvOn,"
    "agt_bizchat_enablePagesCitationsForMultiturn,"
    "agt_module_canvasSetup_enablePagesCitationsForMultiturn,"
    "EnableComposeWidget,feature.EnableMergingPureDeltas,"
    "feature.isExternalEmailEnabled,feature.isExcludedEmailEnabled,"
    "feature.disabledisallowedmsgs,feature.enableCitationsForSynthesisData,"
    "feature.EnableConversationShareApis,feature.enableGenerateGraphicArtOptionsSet,"
    "cdximagen,feature.EnableContentApiandDocTypeHtmlInRichAnswers,"
    "cdxgrounding_api_v2_rich_web_answers_reference_bottom_force,"
    "cdxenablerenderforisocomp,feature.EnableDesignEditorImageGrounding,"
    "feature.EnableDesignerEditor,feature.EnableSkipRehydrationForSpeCIdImages,"
    "feature.sourcescontrolmainline,feature.sourcescontrolmainlineal,"
    "feature.EnableConnectorExecutionControlsAllowlist,"
    "feature.EnableBizchatMainlineExecutionControlsResolution,"
    "feature.EnablePersonalization,cdxentrecapvifluxv3,rich_responses,"
    "feature.EnableBase64DataInMessageAnnotations,feature.EnableStarterLicenseCheckBypass,"
    "feature.DisableMimir3sFlow,feature.EnablePersonalWorkingSetFor3s,"
    "feature.EnableSkipEmittingMessageOnFlush,feature.EnableRemoveEmptySourceAttributions,"
    "feature.EnableRemoveStreamingMode,feature.OfficeWebToHelix,"
    "feature.OfficeDesktopToHelix,feature.M365TeamsHubToHelix,"
    "feature.OwaHubToHelix,feature.MonarchHubToHelix,feature.Win32OutlookHubToHelix,"
    "feature.MacOutlookHubToHelix,Agt_bizchat_enableGpt5ForHelix"
)


def build_url(token, hex_sid=None, conversation_id=None, x_session_id=None):
    if not USER_OID or not TENANT_ID:
        raise ValueError(
            "M365_USER_OID and M365_TENANT_ID environment variables required.\n"
            "Get them from: https://graph.microsoft.com/v1.0/me (id and tenantId)"
        )
    if hex_sid is None:
        hex_sid = uuid.uuid4().hex
    if x_session_id is None:
        uuid_sid = f"{hex_sid[:8]}-{hex_sid[8:12]}-{hex_sid[12:16]}-{hex_sid[16:20]}-{hex_sid[20:32]}"
    else:
        uuid_sid = x_session_id
    url = f"wss://substrate.office.com/m365Copilot/Chathub/{USER_OID}@{TENANT_ID}"
    url += f"?chatsessionid={hex_sid}&XRoutingParameterSessionKey={hex_sid}"
    url += f"&clientrequestid={hex_sid}&X-SessionId={uuid_sid}"
    if conversation_id:
        url += f"&ConversationId={conversation_id}"
    url += f"&access_token={token}"
    url += f"&variants={VARIANTS}"
    url += "&source=%22owahub%22&product=OwaHub&agentHost=Bizchat.FullScreen"
    url += "&licenseType=Starter&isEdu=false&agent=none&scenario=owahub"
    return url, hex_sid, uuid_sid


def build_payload(hex_sid, uuid_sid, text, tone="Magic"):
    inv_id = str(uuid.uuid4())
    p = {
        "type": 4, "invocationId": inv_id, "target": "chat",
        "arguments": [{
            "source": "owahub",
            "clientCorrelationId": hex_sid,
            "sessionId": uuid_sid,
            "traceId": hex_sid,
            "optionsSets": list(OUTLOOK_OPTIONS_SETS),
            "streamingMode": "ConciseWithPadding",
            "options": {},
            "extraExtensionParameters": {},
            "allowedMessageTypes": ALLOWED_MSG_TYPES,
            "sliceIds": [],
            "threadLevelGptId": {},
            "isStartOfSession": False,
            "clientInfo": {
                "clientPlatform": "OwaHub-web",
                "clientAppName": "OwaHub",
                "clientEntrypoint": "owahub",
                "clientSessionId": uuid_sid,
                "clientAppType": "Web",
                "deviceOS": "Linux",
                "deviceType": "Desktop",
                "clientPlatformVersion": "Unknown",
            },
            "message": {
                "author": "user",
                "inputMethod": "Keyboard",
                "text": text,
                "entityAnnotationTypes": ["People", "File", "Event", "Email", "TeamsMessage"],
                "requestId": f"{hex_sid}_0",
                "locationInfo": {"timeZoneOffset": LOCAL_TZ_OFFSET, "timeZone": LOCAL_TZ_NAME},
                "locale": LOCAL_LOCALE,
                "messageType": "Chat",
                "experienceType": "Default",
                "adaptiveCards": [],
                "clientPreferences": {
                    "executionControls": {"web": {}, "work": {}}
                },
            },
            "gpts": [{
                "id": "bizchat-as-gpt-scenario",
                "source": "BuiltInAgents",
                "clientOverrides": {
                    "capabilities": [{"name": "WebSearch"}, {"name": "WorkSearch"}],
                    "deepResearchModels@odata.type": "Collection(String)",
                },
            }],
            "plugins": [{"Id": "BingWebSearch", "Source": "BuiltIn"}],
            "tone": tone,
            "renderReferencesBehindEOS": True,
            "disconnectBehavior": "continue",
        }]
    }
    return json.dumps(p)


def build_conversation_payload(hex_sid, uuid_sid, messages, tone="Magic"):
    inv_id = str(uuid.uuid4())
    m365_history = []
    last_text = messages[-1].get("content", "") if messages else ""

    for m in messages[:-1]:
        role = m.get("role", "")
        content = m.get("content", "")
        if isinstance(content, list):
            texts = [p.get("text", "") for p in content if p.get("type") == "text"]
            content = " ".join(texts)

        if role == "user":
            m365_history.append({
                "author": "user",
                "inputMethod": "Keyboard",
                "text": content or last_text,
                "messageType": "Chat",
                "experienceType": "Default",
                "adaptiveCards": [],
                "clientPreferences": {"executionControls": {"web": {}, "work": {}}},
            })
        elif role == "assistant" and content:
            m365_history.append({
                "author": "bot",
                "text": content,
                "messageType": "Chat",
            })
        elif role == "tool":
            m365_history.append({
                "author": "user",
                "inputMethod": "Keyboard",
                "text": f"[Tool result: {content}]",
                "messageType": "Chat",
                "adaptiveCards": [],
                "clientPreferences": {"executionControls": {"web": {}, "work": {}}},
            })

    p = {
        "type": 4, "invocationId": inv_id, "target": "chat",
        "arguments": [{
            "source": "owahub",
            "clientCorrelationId": hex_sid,
            "sessionId": uuid_sid,
            "traceId": hex_sid,
            "optionsSets": list(OUTLOOK_OPTIONS_SETS),
            "streamingMode": "ConciseWithPadding",
            "options": {},
            "extraExtensionParameters": {},
            "allowedMessageTypes": ALLOWED_MSG_TYPES,
            "sliceIds": [],
            "threadLevelGptId": {},
            "isStartOfSession": False,
            "clientInfo": {
                "clientPlatform": "OwaHub-web",
                "clientAppName": "OwaHub",
                "clientEntrypoint": "owahub",
                "clientSessionId": uuid_sid,
                "clientAppType": "Web",
                "deviceOS": "Linux",
                "deviceType": "Desktop",
                "clientPlatformVersion": "Unknown",
            },
            "message": {
                "author": "user",
                "inputMethod": "Keyboard",
                "text": last_text,
                "entityAnnotationTypes": ["People", "File", "Event", "Email", "TeamsMessage"],
                "requestId": f"{hex_sid}_0",
                "locale": LOCAL_LOCALE,
                "messageType": "Chat",
                "experienceType": "Default",
                "adaptiveCards": [],
                "clientPreferences": {"executionControls": {"web": {}, "work": {}}},
            },
            "gpts": [{
                "id": "bizchat-as-gpt-scenario",
                "source": "BuiltInAgents",
                "clientOverrides": {
                    "capabilities": [{"name": "WebSearch"}, {"name": "WorkSearch"}],
                    "deepResearchModels@odata.type": "Collection(String)",
                },
            }],
            "plugins": [{"Id": "BingWebSearch", "Source": "BuiltIn"}],
            "tone": tone,
            "renderReferencesBehindEOS": True,
            "disconnectBehavior": "continue",
        }]
    }
    if m365_history:
        p["arguments"][0]["messageHistory"] = m365_history
    return json.dumps(p)
