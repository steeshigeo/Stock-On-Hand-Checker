"""
sync_onedrive.py
-----------------
Downloads the latest version of the store's stock/promo Excel file straight
from OneDrive/SharePoint (Microsoft 365) using an Azure AD App Registration
(a "robot" identity, the same idea as Google's Service Account in the
reference architecture).

Required environment variables (set as GitHub Secrets):
  MS_TENANT_ID       - Azure AD tenant ID (or the tenant's domain, e.g. mapcoid.onmicrosoft.com)
  MS_CLIENT_ID       - Azure AD App Registration's Application (client) ID
  MS_CLIENT_SECRET   - Azure AD App Registration's client secret value
  ONEDRIVE_SHARE_URL - the normal "anyone with the link" share URL for the Excel file
                        (the same kind of link you already tested in Incognito)

Output:
  downloaded/stockpositionreportdetail.xls  (raw bytes, whatever the source file is)
"""
import os
import sys
import base64
import pathlib
import requests

TENANT_ID = os.environ["MS_TENANT_ID"]
CLIENT_ID = os.environ["MS_CLIENT_ID"]
CLIENT_SECRET = os.environ["MS_CLIENT_SECRET"]
SHARE_URL = os.environ["ONEDRIVE_SHARE_URL"]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
OUT_DIR = pathlib.Path("downloaded")
OUT_PATH = OUT_DIR / "stockpositionreportdetail.xls"


def get_app_token() -> str:
    """Client-credentials (app-only) OAuth2 flow — no user login involved,
    exactly like a service-account/bot identity."""
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    r = requests.post(url, data=data, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]


def encode_share_url(url: str) -> str:
    """Microsoft Graph needs share links encoded as an opaque 'shares' id:
    base64 the URL, make it URL-safe, strip padding, prefix with 'u!'.
    https://learn.microsoft.com/en-us/graph/api/shares-get
    """
    b = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8")
    b = b.rstrip("=")
    return "u!" + b


def download_file(token: str) -> bytes:
    share_id = encode_share_url(SHARE_URL)
    headers = {"Authorization": f"Bearer {token}"}

    # Resolve the share link to a driveItem, then stream its content.
    meta_url = f"{GRAPH_BASE}/shares/{share_id}/driveItem"
    r = requests.get(meta_url, headers=headers, timeout=30)
    r.raise_for_status()
    item = r.json()
    print(f"Resolved file: {item.get('name')} (id={item.get('id')})")

    content_url = f"{GRAPH_BASE}/shares/{share_id}/driveItem/content"
    r = requests.get(content_url, headers=headers, timeout=60, allow_redirects=True)
    r.raise_for_status()
    return r.content


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    token = get_app_token()
    content = download_file(token)
    OUT_PATH.write_bytes(content)
    print(f"Saved {len(content):,} bytes to {OUT_PATH}")


if __name__ == "__main__":
    try:
        main()
    except requests.HTTPError as e:
        print(f"Graph API error: {e.response.status_code} {e.response.text}", file=sys.stderr)
        sys.exit(1)
