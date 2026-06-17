"""GitHub release checking and local git update helpers."""
from __future__ import annotations

import json
import re
import subprocess
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path


DEFAULT_BRANCH = "main"


@dataclass
class ReleaseInfo:
    tag: str
    name: str
    url: str
    published_at: str
    body: str
    source: str = "release"


def _version_tuple(value: str) -> tuple[int, ...]:
    """Parse v1.2.3-ish tags into tuples that can be compared."""
    value = (value or "").strip().lstrip("vV")
    parts = re.findall(r"\d+", value)
    return tuple(int(p) for p in parts) if parts else (0,)


def is_newer(latest: str, current: str) -> bool:
    latest_tuple = _version_tuple(latest)
    current_tuple = _version_tuple(current)
    width = max(len(latest_tuple), len(current_tuple))
    latest_tuple += (0,) * (width - len(latest_tuple))
    current_tuple += (0,) * (width - len(current_tuple))
    return latest_tuple > current_tuple


def _github_json(url: str, timeout: int):
    req = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "ProQuote-Updater",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def latest_tag(owner: str, repo: str, timeout: int = 10) -> ReleaseInfo:
    url = f"https://api.github.com/repos/{owner}/{repo}/tags?per_page=1"
    data = _github_json(url, timeout)
    if not data:
        raise RuntimeError("No GitHub release or tag was found for this repository.")
    tag = str(data[0].get("name") or "")
    return ReleaseInfo(
        tag=tag,
        name=f"Latest tag {tag}",
        url=f"https://github.com/{owner}/{repo}/tree/{tag}",
        published_at="",
        body="",
        source="tag",
    )


def latest_release(owner: str, repo: str, timeout: int = 10) -> ReleaseInfo:
    owner = (owner or "").strip()
    repo = (repo or "").strip()
    if not owner or not repo:
        raise ValueError("GitHub owner and repository are required.")

    url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    try:
        data = _github_json(url, timeout)
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            try:
                return latest_tag(owner, repo, timeout)
            except urllib.error.URLError as tag_exc:
                raise RuntimeError(f"Could not connect to GitHub: {tag_exc.reason}") from tag_exc
            except urllib.error.HTTPError as tag_exc:
                raise RuntimeError(f"GitHub returned HTTP {tag_exc.code} while checking tags.") from tag_exc
        raise RuntimeError(f"GitHub returned HTTP {exc.code}.") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not connect to GitHub: {exc.reason}") from exc

    return ReleaseInfo(
        tag=str(data.get("tag_name") or ""),
        name=str(data.get("name") or data.get("tag_name") or ""),
        url=str(data.get("html_url") or ""),
        published_at=str(data.get("published_at") or ""),
        body=str(data.get("body") or ""),
        source="release",
    )


def run_git_update(app_dir: Path, branch: str = DEFAULT_BRANCH) -> tuple[bool, str]:
    """Fast-forward the local checkout from origin. Returns success + command output."""
    app_dir = Path(app_dir)
    commands = [
        ["git", "fetch", "--tags", "origin"],
        ["git", "pull", "--ff-only", "origin", branch],
    ]
    output: list[str] = []
    for cmd in commands:
        proc = subprocess.run(
            cmd,
            cwd=app_dir,
            text=True,
            capture_output=True,
            timeout=120,
        )
        output.append(f"$ {' '.join(cmd)}")
        if proc.stdout:
            output.append(proc.stdout.strip())
        if proc.stderr:
            output.append(proc.stderr.strip())
        if proc.returncode != 0:
            return False, "\n".join(part for part in output if part)
    return True, "\n".join(part for part in output if part)