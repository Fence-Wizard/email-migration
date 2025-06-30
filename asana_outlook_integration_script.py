#!/usr/bin/env python3
"""
Asana Outlook Integration Script

Dependencies:
  pip install msal requests asana
"""
import os
import time
import sys
import msal
import requests
try:
    import asana
except ImportError:
    sys.exit("ERROR: Missing dependency 'asana'. Please install via 'pip install asana'")
from requests.exceptions import HTTPError

# ------------------------------------------------------------
# CONFIGURATION — EDIT THESE VALUES
# ------------------------------------------------------------
TENANT_ID             = "10a53ce9-0d13-47d2-876c-57bfa6433582"
CLIENT_ID             = "4ed4dfba-881a-40e4-9190-ac0e90cb1796"
CLIENT_SECRET         = "RWm8Q~offFmH4d_CuU1YTeiOiyzz6018LBfc7bMX"
SCOPES                = ["https://graph.microsoft.com/.default"]

MAIL_USER             = "TMyers@hurricanefence.com"
MAIL_FOLDER_PATH      = ["Inbox", "2024 Jobs", "Nova", "2411001"]

ASANA_PAT             = "2/1208111632809268/1210537347912656:56277a87abb51ac95aedaf8d58c585ca"
ASANA_WORKSPACE_GID   = "1208111550442047"
ASANA_PROJECT_GID     = "1210440111585066"
ASANA_SECTION_GID     = "1210445509137313"

PROCESSED_IDS_FILE    = "processed_ids.txt"
TEMP_DIR              = "temp_attachments"
SLEEP_INTERVAL        = 0.5  # seconds between operations

LOCATION_FIELD_GID    = "1208148657165841"
JOB_NUMBER_FIELD_GID  = "1210440111585075"
# ------------------------------------------------------------
# END CONFIGURATION
# ------------------------------------------------------------

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" not in result:
        raise Exception(f"Token error: {result.get('error_description')}")
    print("[Graph] Acquired access token")
    return result["access_token"]


def load_processed_ids(path):
    if not os.path.exists(path):
        return set()
    with open(path, "r", encoding="utf-8") as f:
        return set(line.strip() for line in f if line.strip())


def save_processed_id(path, msg_id):
    os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        f.write(msg_id + "\n")


def ensure_temp_dir(path):
    os.makedirs(path, exist_ok=True)


def connect_asana(pat):
    config = asana.Configuration()
    config.access_token = pat
    client = asana.ApiClient(config)
    tasks_api = asana.TasksApi(client)
    attach_api = asana.AttachmentsApi(client)
    sections_api = asana.SectionsApi(client)
    user = asana.UsersApi(client).get_user("me", {})
    print(f"[Asana] Connected as {user['name']} ({user['email']})")
    return tasks_api, attach_api, sections_api


def get_target_folder_id(token, path_list):
    headers = {"Authorization": f"Bearer {token}"}
    folder_id = None
    for part in path_list:
        url = (
            f"{GRAPH_BASE}/users/{MAIL_USER}/mailFolders/{folder_id}/childFolders"
            if folder_id else f"{GRAPH_BASE}/users/{MAIL_USER}/mailFolders"
        )
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        items = resp.json().get("value", [])
        match = next((i for i in items if i.get("displayName") == part), None)
        if not match:
            raise Exception(f"Folder '{part}' not found at {url}")
        folder_id = match.get("id")
    print(f"[Graph] Folder ID '{folder_id}' for path {'/'.join(path_list)}")
    return folder_id


def process_message(msg, tasks_api, attach_api, sections_api, location, job_num, token):
    subj = msg.get("subject", "(No Subject)")
    received = msg.get("receivedDateTime", "")
    sender = (
        msg.get("from", {}).get("emailAddress", {})
           .get("address", "")
    )

    # safely extract body content or fallback to preview
    if isinstance(msg.get("body"), dict):
        body = msg["body"].get("content", "")
    else:
        body = msg.get("bodyPreview", "")

    notes = (
        f"**Location:** {location}\n"
        f"**Job #:** {job_num}\n"
        f"**From:** {sender}\n"
        f"**Received:** {received}\n\n"
        f"{body}"
    )

    task_payload = {
        "name": subj,
        "notes": notes,
        "projects": [ASANA_PROJECT_GID],
        "workspace": ASANA_WORKSPACE_GID,
    }
    task = tasks_api.create_task({"data": task_payload}, {})
    gid = task.get("gid")

    # add to section without extra opts
    sections_api.add_task_for_section(
        ASANA_SECTION_GID,
        {"data": {"task": gid}}
    )

    # update custom fields
    update_payload = {
        "custom_fields": {
            LOCATION_FIELD_GID: location,
            JOB_NUMBER_FIELD_GID: int(job_num)
        }
    }
    tasks_api.update_task(gid, {"data": update_payload}, {})

    # handle attachments
    ensure_temp_dir(TEMP_DIR)
    headers = {"Authorization": f"Bearer {token}"}
    parent_id = msg.get("parentFolderId")
    for att in msg.get("attachments", []):
        if att.get("@odata.type", "").endswith("ItemAttachment"):
            continue
        att_id = att.get("id")
        url = (
            f"{GRAPH_BASE}/users/{MAIL_USER}/mailFolders/"
            f"{parent_id}/messages/{msg['id']}/"
            f"attachments/{att_id}/$value"
        )
        try:
            r = requests.get(url, headers=headers)
            r.raise_for_status()
        except HTTPError:
            if r.status_code == 413:
                print(f"[SKIP] Attachment too large: {att.get('name')}")
                continue
            else:
                raise
        local = os.path.join(TEMP_DIR, att.get("name"))
        with open(local, "wb") as f:
            f.write(r.content)
        with open(local, "rb") as f:
            attach_api.create_attachment_on_task(gid, f, {})
        os.remove(local)


def main():
    token = get_access_token()
    done = load_processed_ids(PROCESSED_IDS_FILE)
    tasks_api, attach_api, sections_api = connect_asana(ASANA_PAT)

    fid = get_target_folder_id(token, MAIL_FOLDER_PATH)
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"{GRAPH_BASE}/users/{MAIL_USER}/mailFolders/{fid}/messages?"
        f"$select=id,subject,receivedDateTime,from,body,parentFolderId&$expand=attachments"
    )
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    msgs = resp.json().get("value", [])

    for msg in msgs:
        mid = msg.get("id")
        if mid in done:
            continue
        try:
            loc = MAIL_FOLDER_PATH[-2]
            job = MAIL_FOLDER_PATH[-1]
            process_message(msg, tasks_api, attach_api, sections_api, loc, job, token)
            save_processed_id(PROCESSED_IDS_FILE, mid)
        except Exception as e:
            print(f"[ERROR] {mid}: {e}")
        time.sleep(SLEEP_INTERVAL)

    print("✅ Dry-run complete.")


if __name__ == "__main__":
    main()
