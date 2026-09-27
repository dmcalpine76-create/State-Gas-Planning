"""core/env.py — .env discovery, previously copy-pasted in four scripts."""

from pathlib import Path
from dotenv import load_dotenv, find_dotenv

from .config import BASE_DIR


def load_env() -> bool:
    """
    Load the first .env found in BASE_DIR, its parent, or grandparent,
    falling back to python-dotenv's own search. Returns True if loaded.
    """
    for candidate in [BASE_DIR, BASE_DIR.parent, BASE_DIR.parent.parent]:
        env = candidate / ".env"
        if env.exists():
            load_dotenv(env)
            return True
    found = find_dotenv(usecwd=True)
    if found:
        load_dotenv(found)
        return True
    print("  ⚠️  No .env file found.")
    print("      Create a .env file in the tools folder containing:")
    print("        OUTLOOK_CLIENT_ID=your-client-id")
    print("        OUTLOOK_TENANT_ID=consumers   (or your org tenant ID)")
    print("        ANTHROPIC_API_KEY=sk-ant-...")
    return False
