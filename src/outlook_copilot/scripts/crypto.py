"""Encrypt/decrypt refresh tokens at rest using Fernet (AES-128-CBC).

Key is stored outside the project at ~/.m365-copilot/encryption.key
so that encrypted tokens in the project directory are useless without it.
"""
import os, stat
from pathlib import Path

try:
    from cryptography.fernet import Fernet
except ImportError:
    raise ImportError(
        "需要 cryptography 包。安装: pip install cryptography"
    )

KEY_DIR = Path.home() / ".m365-copilot"
KEY_FILE = KEY_DIR / "encryption.key"


def _load_or_create_key() -> bytes:
    if KEY_FILE.exists():
        return KEY_FILE.read_bytes()
    KEY_DIR.mkdir(parents=True, exist_ok=True)
    key = Fernet.generate_key()
    KEY_FILE.write_bytes(key)
    KEY_FILE.chmod(stat.S_IRUSR | stat.S_IWUSR)
    return key


def encrypt(plaintext: str) -> str:
    f = Fernet(_load_or_create_key())
    return f.encrypt(plaintext.encode()).decode()


def decrypt(ciphertext: str) -> str:
    f = Fernet(_load_or_create_key())
    return f.decrypt(ciphertext.encode()).decode()
