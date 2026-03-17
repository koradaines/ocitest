# github_storage.py
# Handles reading, writing, and deleting client JSON files
# stored in the GitHub repo under the /saved_clients/ folder.
#
# Required Streamlit secret:
#   GITHUB_TOKEN  — a Personal Access Token with repo read/write scope
#   GITHUB_REPO   — e.g. "your-org/pharmaroi-app"
#   GITHUB_BRANCH — e.g. "main"

from __future__ import annotations

import base64
import json
from typing import Optional

import requests
import streamlit as st

# -----------------------------
# Config (pulled from st.secrets)
# -----------------------------
def _get_cfg():
    try:
        token  = st.secrets["GITHUB_TOKEN"]
        repo   = st.secrets["GITHUB_REPO"]
        branch = st.secrets.get("GITHUB_BRANCH", "main")
        return token, repo, branch
    except Exception:
        return None, None, "main"

FOLDER = "saved_clients"

def _headers(token: str) -> dict:
    return {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github+json",
    }

def _file_url(repo: str, client_name: str) -> str:
    safe = client_name.strip().replace(" ", "_")
    return f"https://api.github.com/repos/{repo}/contents/{FOLDER}/{safe}.json"

# -----------------------------
# Public API
# -----------------------------

def list_clients() -> list[str]:
    """Return sorted list of saved client names (without .json extension)."""
    token, repo, branch = _get_cfg()
    if not token:
        return []

    url = f"https://api.github.com/repos/{repo}/contents/{FOLDER}?ref={branch}"
    resp = requests.get(url, headers=_headers(token), timeout=10)

    if resp.status_code == 404:
        return []                          # folder doesn't exist yet — that's fine
    if not resp.ok:
        st.warning(f"GitHub: could not list clients ({resp.status_code})")
        return []

    files = resp.json()
    return sorted(
        f["name"].replace(".json", "")
        for f in files
        if isinstance(f, dict) and f.get("name", "").endswith(".json")
    )


def load_client(client_name: str) -> Optional[dict]:
    """Load and return the JSON payload for a client, or None on failure."""
    token, repo, branch = _get_cfg()
    if not token:
        return None

    url = _file_url(repo, client_name) + f"?ref={branch}"
    resp = requests.get(url, headers=_headers(token), timeout=10)

    if not resp.ok:
        st.error(f"Could not load '{client_name}' ({resp.status_code})")
        return None

    raw = base64.b64decode(resp.json()["content"]).decode("utf-8")
    return json.loads(raw)


def save_client(client_name: str, payload: dict) -> bool:
    """
    Create or update the JSON file for client_name.
    Returns True on success.
    """
    token, repo, branch = _get_cfg()
    if not token:
        st.error("GitHub credentials not configured. Add GITHUB_TOKEN, GITHUB_REPO to Streamlit secrets.")
        return False

    url = _file_url(repo, client_name)
    content_b64 = base64.b64encode(
        json.dumps(payload, indent=2).encode("utf-8")
    ).decode("utf-8")

    # Check if file already exists (need its SHA to update)
    sha = None
    check = requests.get(url + f"?ref={branch}", headers=_headers(token), timeout=10)
    if check.ok:
        sha = check.json().get("sha")

    body: dict = {
        "message": f"Save client: {client_name}",
        "content": content_b64,
        "branch": branch,
    }
    if sha:
        body["sha"] = sha

    resp = requests.put(url, headers=_headers(token), json=body, timeout=15)
    if resp.ok:
        return True
    else:
        st.error(f"GitHub save failed ({resp.status_code}): {resp.text[:200]}")
        return False


def delete_client(client_name: str) -> bool:
    """
    Delete the JSON file for client_name from the repo.
    Returns True on success.
    """
    token, repo, branch = _get_cfg()
    if not token:
        return False

    url = _file_url(repo, client_name)

    # Must supply SHA to delete
    check = requests.get(url + f"?ref={branch}", headers=_headers(token), timeout=10)
    if not check.ok:
        st.error(f"Could not find '{client_name}' to delete.")
        return False

    sha = check.json().get("sha")
    body = {
        "message": f"Delete client: {client_name}",
        "sha": sha,
        "branch": branch,
    }
    resp = requests.delete(url, headers=_headers(token), json=body, timeout=15)
    if resp.ok:
        return True
    else:
        st.error(f"GitHub delete failed ({resp.status_code}): {resp.text[:200]}")
        return False
