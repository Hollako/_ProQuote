"""Runtime/deployment helpers for ProQuote."""
from __future__ import annotations

import os
import shutil
from pathlib import Path


def is_streamlit_cloud() -> bool:
    """Best-effort detection for hosted Streamlit Cloud-style runtimes."""
    markers = (
        "STREAMLIT_CLOUD",
        "STREAMLIT_SHARING_MODE",
        "STREAMLIT_RUNTIME_ENV",
        "STREAMLIT_SERVER_HEADLESS",
    )
    if any(os.environ.get(k) for k in markers[:3]):
        return True
    return bool(os.environ.get("STREAMLIT_SERVER_HEADLESS") and os.environ.get("HOME") == "/home/adminuser")


def git_update_available(app_dir: str | os.PathLike) -> tuple[bool, str]:
    app_dir = Path(app_dir)
    if is_streamlit_cloud():
        return False, "Cloud deployments update from GitHub automatically. Use the Streamlit app dashboard to reboot if needed."
    if not (app_dir / ".git").exists():
        return False, "This installation was not installed as a Git checkout, so in-app git updates are unavailable. Install a newer setup file instead."
    if shutil.which("git") is None:
        return False, "Git is not installed or not available on this PC."
    return True, ""