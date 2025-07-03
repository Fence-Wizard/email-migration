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
import logging
import random
import threading
import traceback
from io import StringIO
from bs4 import BeautifulSoup
import ast
from dotenv import load_dotenv
try:
    import asana
except ImportError:
    sys.exit("ERROR: Missing dependency 'asana'. Please install via 'pip install asana'")
from requests.exceptions import HTTPError

# load environment variables and configure logging
ENV_FILE = os.getenv("ENV_FILE", ".env")
load_dotenv(ENV_FILE)
logging.basicConfig(
    level=logging.DEBUG,
    format='[%(asctime)s] [%(levelname)s] %(name)s: %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# ------------------------------------------------------------
# CONFIGURATION — EDIT THESE VALUES
# ------------------------------------------------------------
TENANT_ID             = os.getenv("TENANT_ID")
CLIENT_ID             = os.getenv("CLIENT_ID")
CLIENT_SECRET         = os.getenv("CLIENT_SECRET")
SCOPES                = ["https://graph.microsoft.com/.default"]

MAIL_USER             = os.getenv("MAIL_USER")
MAIL_FOLDER_PATH      = ast.literal_eval(os.getenv("MAIL_FOLDER_PATH", "[]"))

ASANA_PAT             = os.getenv("ASANA_PAT")
ASANA_WORKSPACE_GID   = os.getenv("ASANA_WORKSPACE_GID")
ASANA_PROJECT_GID     = os.getenv("ASANA_PROJECT_GID")
ASANA_SECTION_GID     = os.getenv("ASANA_SECTION_GID")
BUDGET_SECTION_GID    = os.getenv("BUDGET_SECTION_GID")
QUOTE_SECTION_GID     = os.getenv("QUOTE_SECTION_GID")
ORDER_SECTION_GID     = os.getenv("ORDER_SECTION_GID")

PROCESSED_IDS_FILE    = "processed_ids.txt"
TEMP_DIR              = "temp_attachments"
SLEEP_INTERVAL        = 0.5  # seconds between operations

LOCATION_FIELD_GID    = os.getenv("LOCATION_FIELD_GID")
JOB_NUMBER_FIELD_GID  = os.getenv("JOB_NUMBER_FIELD_GID")
# ------------------------------------------------------------
# END CONFIGURATION
# ------------------------------------------------------------

required_vars = [TENANT_ID, CLIENT_ID, CLIENT_SECRET, MAIL_USER, ASANA_PAT,
                 ASANA_WORKSPACE_GID, ASANA_PROJECT_GID, ASANA_SECTION_GID]
if not all(required_vars):
    logger.error("Missing required environment variables. Check your .env file.")
    sys.exit(1)

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
    logger.info("[Graph] Acquired access token")
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
    logger.info("[Asana] Connected as %s (%s)", user['name'], user['email'])
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
    logger.info("[Graph] Folder ID '%s' for path %s", folder_id, '/'.join(path_list))
    return folder_id


def process_message(msg, tasks_api, attach_api, sections_api, location, job_num, token):
    subj = msg.get("subject", "(No Subject)")
    received = msg.get("receivedDateTime", "")
    sender = (
        msg.get("from", {}).get("emailAddress", {})
           .get("address", "")
    )

    # --- DEBUG: inspect the raw body node ---
    body_node = msg.get("body")
    logger.debug("Body node for message %s: %r", msg.get("id"), body_node)
    if isinstance(body_node, dict):
        if "content" not in body_node:
            logger.error(
                "Message %s: 'body' missing 'content' key. Available keys: %r",
                msg.get("id"),
                list(body_node.keys()),
            )
        body = body_node.get("content", "")
    else:
        preview = msg.get("bodyPreview")
        logger.debug(
            "Using bodyPreview for message %s: %r", msg.get("id"), preview
        )
        body = preview or ""

    # Sanitize HTML bodies to plain text
    if body.lstrip().startswith("<"):
        soup = BeautifulSoup(body, "html.parser")
        clean_body = soup.get_text(separator="\n").strip()
    else:
        clean_body = body.strip()

    notes = (
        f"**Location:** {location}\n"
        f"**Job #:** {job_num}\n"
        f"**From:** {sender}\n"
        f"**Received:** {received}\n\n"
        f"{clean_body}"
    )

    subject_lower = subj.lower()
    if "budget" in subject_lower and BUDGET_SECTION_GID:
        section_gid = BUDGET_SECTION_GID
    elif "quotation" in subject_lower and QUOTE_SECTION_GID:
        section_gid = QUOTE_SECTION_GID
    elif "order confirmation" in subject_lower and ORDER_SECTION_GID:
        section_gid = ORDER_SECTION_GID
    else:
        section_gid = ASANA_SECTION_GID

    task_payload = {
        "name": subj,
        "notes": notes,
        "projects": [ASANA_PROJECT_GID],
        "workspace": ASANA_WORKSPACE_GID,
    }
    task = tasks_api.create_task({"data": task_payload}, {})
    gid = task.get("gid")

    # Add the task to the chosen section
    try:
        sections_api.add_task_for_section(
            section_gid,
            {
                "body": {
                    "data": {
                        "task": gid
                    }
                }
            }
        )
    except Exception as e:
        logger.error(
            "Failed to add task %s to section %s: %s",
            gid,
            section_gid,
            e,
        )

    # update custom fields on the created task (wrap under 'body')
    update_payload = {
        LOCATION_FIELD_GID: location,
        JOB_NUMBER_FIELD_GID: int(job_num)
    }
    tasks_api.update_task(
        gid,
        {
            "body": {
                "data": update_payload
            }
        }
    )

    # handle attachments
    ensure_temp_dir(TEMP_DIR)
    headers = {"Authorization": f"Bearer {token}"}
    parent_id = msg.get("parentFolderId")
    for att in msg.get("attachments", []):
        if att.get("@odata.type", "").endswith("ItemAttachment"):
            continue
        if att.get("size", 0) > 3 * 1024 * 1024:
            logger.warning("[SKIP] Attachment too large: %s", att.get("name"))
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
                logger.warning("[SKIP] Attachment too large: %s", att.get('name'))
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

    # diagnostics: dump all sections for this project
    all_secs = sections_api.get_sections_for_project(
        ASANA_PROJECT_GID,
        {"opt_fields": "gid,name"}
    )
    logger.info(
        "Project sections available: %r",
        [(s["gid"], s["name"]) for s in all_secs]
    )

    fid = get_target_folder_id(token, MAIL_FOLDER_PATH)
    # ask Microsoft Graph to include id, subject, body, and date on each message
    headers = {"Authorization": f"Bearer {token}"}
    params = {
        "$select": "id,subject,body,receivedDateTime,from,parentFolderId",
        "$expand": "attachments",
        "$top": 50,
    }
    url = f"{GRAPH_BASE}/users/{MAIL_USER}/mailFolders/{fid}/messages"
    next_url = url
    while next_url:
        req_params = params if next_url == url else None
        logger.debug(
            "Graph GET %s with params: %s", next_url, req_params)
        resp = requests.get(next_url, headers=headers, params=req_params)
        try:
            resp.raise_for_status()
        except HTTPError as err:
            logger.error("Request failed: %s", err)
            break
        data = resp.json()
        logger.debug(
            "Raw Graph response sample: %r",
            data.get("value", [])[:2]
        )
        msgs = data.get("value", [])
        logger.debug(
            "Graph returned %d messages; keys of first 3: %r",
            len(msgs),
            [list(m.keys()) for m in msgs[:3]]
        )

        for msg in msgs:
            # --- DEBUG: inspect each msg before processing ---
            logger.debug(
                "Inspecting message %s: keys=%r",
                msg.get("id"),
                list(msg.keys())
            )
            if 'body' not in msg:
                logger.warning(
                    "Skipping message %s—no 'body' key. Available keys: %s",
                    msg.get("id"),
                    list(msg.keys())
                )
                continue
            content = msg['body'].get('content')
            if not content:
                logger.warning(
                    "Message %s has empty body.content. Raw body node: %r",
                    msg.get("id"),
                    msg['body']
                )
                continue
            mid = msg.get("id")
            if mid in done:
                continue

            try:
                loc = MAIL_FOLDER_PATH[-2]
                job = MAIL_FOLDER_PATH[-1]
                process_message(msg, tasks_api, attach_api, sections_api, loc, job, token)
                save_processed_id(PROCESSED_IDS_FILE, mid)
            except Exception as e:
                # emit full stack trace and raw message for post-mortem
                logger.exception(
                    "Error processing message %s; raw payload: %r", mid, msg
                )
            time.sleep(SLEEP_INTERVAL)

        next_url = data.get("@odata.nextLink")

    logger.info("✅ Dry-run complete.")


def _matrix_effect(stop_event):
    """Print random binary lines until stop_event is set."""
    width = 80
    while not stop_event.is_set():
        print("".join(random.choice("01") for _ in range(width)))
        time.sleep(0.05)
    if os.name == "nt":
        os.system("cls")
    else:
        os.system("clear")


if __name__ == "__main__":
    tb_buf = StringIO()
    stop_evt = threading.Event()
    mat_thread = threading.Thread(target=_matrix_effect, args=(stop_evt,), daemon=True)
    mat_thread.start()

    try:
        main()
    except Exception:
        tb_buf.write(traceback.format_exc())
    finally:
        stop_evt.set()
        mat_thread.join()

    tb_text = tb_buf.getvalue()
    if tb_text:
        print("\n\n=== Traceback (most recent call last) ===\n")
        sys.stdout.write(tb_text)
    else:
        print("\n\n\u2714 Migration completed without unhandled exceptions.\n")
