import sys, os, asyncio, argparse, uuid

sys.stdout.reconfigure(encoding="utf-8")

from .. import __version__
from ..auth import TokenManager
from ..client import M365Client
from ..models import MODELS, TENANT_ID, USER_OID, CLIENT_ID, SCOPE

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
RT_FILE = os.path.join(BASE_DIR, "data", "tokens", "rt_90day.txt")
CACHE_FILE = os.path.join(BASE_DIR, "data", "tokens", "token_cache.json")
TOKEN_FILE = os.path.join(BASE_DIR, "data", "tokens", "token.txt")


def main():
    parser = argparse.ArgumentParser(description=f"Outlook Copilot v{__version__}")
    parser.add_argument("prompt", nargs="?", help="question")
    parser.add_argument("--model", default="auto", choices=list(MODELS.keys()), help="model")
    parser.add_argument("-i", "--interactive", action="store_true", help="interactive mode")
    parser.add_argument("--no-stream", action="store_true", help="disable streaming")
    parser.add_argument("--setup", action="store_true", help="first time setup")
    parser.add_argument("--list-models", action="store_true", help="list models")
    parser.add_argument("--token", help="save an access token and connect (extract from WebSocket URL)")
    parser.add_argument("-c", "--conversation", action="store_true", help="conversation mode (same context)")
    parser.add_argument("--refresh", action="store_true", help="update access_token via bookmarklet")
    args = parser.parse_args()

    if args.list_models:
        print("Available models:")
        for k, v in MODELS.items():
            print(f"  {k:12s} - {v['tone']} -> {v['openai_id']}")
        return

    os.makedirs(os.path.dirname(RT_FILE), exist_ok=True)

    if args.token:
        with open(TOKEN_FILE, 'w') as f:
            f.write(args.token.strip())
        print(f"Token saved ({len(args.token.strip())} chars)")
        return

    tm = TokenManager(TENANT_ID, CLIENT_ID, SCOPE, RT_FILE, CACHE_FILE, TOKEN_FILE)

    if args.refresh:
        if not TENANT_ID:
            print("Error: M365_TENANT_ID not configured")
            return
        tm.refresh_interactive()
        return

    if args.setup:
        from ..scripts.setup_wizard import main as setup_main
        setup_main()
        return

    if not TENANT_ID or not USER_OID:
        print("Error: M365_TENANT_ID and M365_USER_OID not configured")
        print("Run: outlook-copilot-setup")
        return

    if not args.prompt and not args.interactive:
        if not os.path.exists(RT_FILE):
            print("First time: outlook-copilot-setup")
            return
        parser.print_help()
        return

    if not os.path.exists(RT_FILE):
        print("First time: outlook-copilot-setup")
        return

    cfg = MODELS[args.model]
    tone = cfg["tone"]
    stream = not args.no_stream
    client = M365Client(tm)

    if args.conversation:
        x_session_id = str(uuid.uuid4())
        conv_id = str(uuid.uuid4())
    else:
        x_session_id = conv_id = None

    async def run():
        if args.interactive:
            print(f"Outlook Copilot v{__version__} (interactive mode)")
            if args.conversation:
                print("  Conversation mode: context shared across messages")
            print(f"Model: {args.model}")
            print("Exit: Ctrl+C")
            try:
                while True:
                    text = input("\n> ")
                    if not text:
                        continue
                    if stream:
                        print()
                        await client.chat_stream(text, tone, conversation_id=conv_id, x_session_id=x_session_id)
                    else:
                        result = await client.chat(text, tone, conversation_id=conv_id, x_session_id=x_session_id)
                        if result:
                            print(result)
                    print()
            except (KeyboardInterrupt, EOFError):
                print("\nBye!")
        else:
            text = args.prompt or ""
            if not text:
                return
            if stream:
                await client.chat_stream(text, tone, conversation_id=conv_id, x_session_id=x_session_id)
            else:
                result = await client.chat(text, tone, conversation_id=conv_id, x_session_id=x_session_id)
                if result:
                    print(result)
            print()
        await client.close()

    asyncio.run(run())
