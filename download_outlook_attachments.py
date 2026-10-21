"""Download matching Outlook mail attachments to a local folder.

Uses Microsoft Graph and device-code sign-in, so it works with new Outlook.
Run it repeatedly (for example with Windows Task Scheduler). Downloaded
attachments are recorded locally and will not be downloaded twice.
"""

from __future__ import annotations

import json
import os
import re
import sys
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any

import msal
import requests
from dotenv import load_dotenv

GRAPH = "https://graph.microsoft.com/v1.0"
# MSAL adds OpenID Connect reserved scopes (including offline_access) itself.
# Passing them explicitly raises ValueError in current MSAL versions.
SCOPES = ["User.Read", "Mail.Read"]


def setting(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def safe_filename(name: str) -> str:
    """Make an attachment name safe for Windows, retaining its extension."""
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(". ")
    return cleaned or "attachment"


def unique_destination(folder: Path, filename: str) -> Path:
    candidate = folder / safe_filename(filename)
    if not candidate.exists():
        return candidate
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return folder / f"{candidate.stem}_{stamp}{candidate.suffix}"


def load_state(path: Path) -> set[str]:
    try:
        return set(json.loads(path.read_text(encoding="utf-8")))
    except (FileNotFoundError, json.JSONDecodeError):
        return set()


def save_state(path: Path, processed: set[str]) -> None:
    # Bound the state file while retaining more IDs than a typical lookback needs.
    path.write_text(json.dumps(sorted(processed)[-10000:]), encoding="utf-8")


def access_token(client_id: str, tenant_id: str, cache_path: Path) -> str:
    cache = msal.SerializableTokenCache()
    if cache_path.exists():
        cache.deserialize(cache_path.read_text(encoding="utf-8"))
    app = msal.PublicClientApplication(
        client_id, authority=f"https://login.microsoftonline.com/{tenant_id}", token_cache=cache
    )
    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Could not start sign-in: {flow}")
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)
    if cache.has_state_changed:
        cache_path.write_text(cache.serialize(), encoding="utf-8")
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description", str(result)))
    return result["access_token"]


def graph_get(url: str, headers: dict[str, str]) -> Any:
    response = requests.get(url, headers=headers, timeout=60)
    response.raise_for_status()
    return response.json()


def find_folder_id(headers: dict[str, str], display_name: str) -> str:
    if display_name.casefold() == "inbox":
        return "inbox"
    data = graph_get(f"{GRAPH}/me/mailFolders?$top=100&$select=id,displayName", headers)
    for folder in data["value"]:
        if folder["displayName"].casefold() == display_name.casefold():
            return folder["id"]
    raise RuntimeError(f"Outlook folder {display_name!r} was not found. Create it or set OUTLOOK_FOLDER=Inbox.")


def download_attachment(url: str, headers: dict[str, str], destination: Path) -> None:
    response = requests.get(url, headers=headers, timeout=120)
    response.raise_for_status()
    destination.write_bytes(response.content)


def main() -> int:
    load_dotenv()
    client_id = setting("CLIENT_ID")
    if not client_id or client_id.startswith("replace-"):
        print("Set CLIENT_ID in .env. Start by copying .env.example to .env.", file=sys.stderr)
        return 2

    folder = Path(setting("DOWNLOAD_FOLDER", "OutlookAttachments")).expanduser()
    folder.mkdir(parents=True, exist_ok=True)
    state_path = folder / ".outlook_attachment_state.json"
    token = access_token(client_id, setting("TENANT_ID", "common"), folder / ".msal_token_cache.json")
    headers = {"Authorization": f"Bearer {token}"}
    mail_folder_id = find_folder_id(headers, setting("OUTLOOK_FOLDER", "Inbox"))
    sender = setting("SENDER_EMAIL").casefold()
    subject = setting("SUBJECT_CONTAINS").casefold()
    extensions = {x.strip().casefold() for x in setting("ALLOWED_EXTENSIONS").split(",") if x.strip()}
    since = datetime.now(UTC) - timedelta(days=int(setting("LOOKBACK_DAYS", "7")))
    processed = load_state(state_path)
    downloaded = 0

    url = (f"{GRAPH}/me/mailFolders/{mail_folder_id}/messages?"
           "$top=50&$select=id,subject,receivedDateTime,from,hasAttachments&$orderby=receivedDateTime desc")
    while url:
        page = graph_get(url, headers)
        for message in page["value"]:
            received = datetime.fromisoformat(message["receivedDateTime"].replace("Z", "+00:00"))
            if received < since:
                save_state(state_path, processed)
                print(f"Done. Downloaded {downloaded} attachment(s).")
                return 0
            address = ((message.get("from") or {}).get("emailAddress") or {}).get("address", "").casefold()
            if not message.get("hasAttachments") or (sender and address != sender) or (subject and subject not in (message.get("subject") or "").casefold()):
                continue
            attachments = graph_get(f"{GRAPH}/me/messages/{message['id']}/attachments", headers)["value"]
            for attachment in attachments:
                if not attachment.get("@odata.type", "").endswith("fileAttachment") or attachment.get("isInline"):
                    continue
                name = attachment.get("name") or "attachment"
                if extensions and Path(name).suffix.casefold() not in extensions:
                    continue
                key = f"{message['id']}:{attachment['id']}"
                if key in processed:
                    continue
                destination = unique_destination(folder, name)
                download_attachment(f"{GRAPH}/me/messages/{message['id']}/attachments/{attachment['id']}/$value", headers, destination)
                processed.add(key)
                downloaded += 1
                print(f"Downloaded: {destination}")
        url = page.get("@odata.nextLink")
    save_state(state_path, processed)
    print(f"Done. Downloaded {downloaded} attachment(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
