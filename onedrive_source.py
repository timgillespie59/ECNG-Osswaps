"""
onedrive_source.py — fetches the live xlsx bytes from a OneDrive/SharePoint
"Anyone with the link can view" share URL, without any interactive login.
"""
import base64
import requests


def _download_variant(share_url: str) -> str:
    sep = "&" if "?" in share_url else "?"
    return f"{share_url}{sep}download=1"


def _graph_shares_url(share_url: str) -> str:
    b64 = base64.urlsafe_b64encode(share_url.encode("utf-8")).decode("utf-8")
    token = "u!" + b64.rstrip("=")
    return f"https://api.onedrive.com/v1.0/shares/{token}/root/content"


def fetch_excel_bytes(share_url: str, timeout: int = 20) -> bytes:
    errors = []
    try:
        url = _download_variant(share_url)
        resp = requests.get(url, timeout=timeout, allow_redirects=True)
        if resp.status_code == 200 and resp.content[:2] == b"PK":
            return resp.content
        errors.append(f"download=1 variant: status {resp.status_code}, "
                       f"content-type {resp.headers.get('content-type')}")
    except requests.RequestException as e:
        errors.append(f"download=1 variant failed: {e}")

    try:
        url = _graph_shares_url(share_url)
        resp = requests.get(url, timeout=timeout, allow_redirects=True)
        if resp.status_code == 200 and resp.content[:2] == b"PK":
            return resp.content
        errors.append(f"graph shares endpoint: status {resp.status_code}, "
                       f"content-type {resp.headers.get('content-type')}")
    except requests.RequestException as e:
        errors.append(f"graph shares endpoint failed: {e}")

    raise RuntimeError(
        "Could not fetch the Excel file from OneDrive. Tried:\n- "
        + "\n- ".join(errors)
        + "\n\nMost likely cause: the tenant's sharing policy doesn't allow "
        "fully anonymous access even with 'Anyone with the link' set, or "
        "the link needs to be regenerated. Falling back to the last known "
        "snapshot if one is available."
    )
