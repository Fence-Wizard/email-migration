import os
import json
import sys
import time
import random
import threading
import shutil
import requests
import msal
import pandas as pd
import matplotlib.pyplot as plt
from textblob import TextBlob
from sklearn.decomposition import LatentDirichletAllocation
from sklearn.feature_extraction.text import CountVectorizer
from dotenv import load_dotenv
from io import BytesIO
from datetime import datetime
import openpyxl
import pytesseract
from PIL import Image

try:
    from pdfminer.high_level import extract_text_to_fp
    _HAS_PDFMINER = True
except ImportError:  # pragma: no cover - optional dependency
    _HAS_PDFMINER = False
    extract_text_to_fp = None

try:
    from docx import Document
    _HAS_DOCX = True
except ImportError:  # pragma: no cover - optional dependency
    _HAS_DOCX = False
    Document = None


def _matrix_rain(stop_event):
    """Continuously print green Matrix-style text until stop_event is set."""
    chars = "abcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()"
    width = shutil.get_terminal_size((80, 20)).columns
    while not stop_event.is_set():
        line = "".join(random.choice(chars) for _ in range(width))
        sys.stdout.write("\033[32m" + line + "\033[0m\n")
        sys.stdout.flush()
        time.sleep(0.05)


_stop_matrix = threading.Event()
_matrix_thread = threading.Thread(target=_matrix_rain, args=(_stop_matrix,), daemon=True)
_matrix_thread.start()

# ─── Config & Auth ─────────────────────────────────────────────────────────────────────────
load_dotenv()  # expects .env with CLIENT_ID, TENANT_ID, CLIENT_SECRET

CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID     = os.getenv("TENANT_ID")
USERNAME      = "tmyers@hurricanefence.com"
SCOPES        = ["https://graph.microsoft.com/.default"]
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"

app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)
token = app.acquire_token_for_client(SCOPES)["access_token"]
HEADERS = {"Authorization": f"Bearer {token}"}
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# ─── Locate root folder "2024 Jobs" and map IDs to paths ──────────

def find_root_folder(name: str) -> str:
    """Return folder id by display name or raise RuntimeError."""
    resp = requests.get(f"{GRAPH_BASE}/users/{USERNAME}/mailFolders", headers=HEADERS)
    resp.raise_for_status()
    for f in resp.json().get("value", []):
        if f.get("displayName") == name:
            return f["id"]
    raise RuntimeError(f"Folder named '{name}' not found.")

ROOT_NAME = "2024 Jobs"
root_id = find_root_folder(ROOT_NAME)

folder_paths = {}


def map_paths(path, fid):
    folder_paths[fid] = path
    url = f"{GRAPH_BASE}/users/{USERNAME}/mailFolders/{fid}/childFolders"
    resp = requests.get(url, headers=HEADERS)
    resp.raise_for_status()
    for child in resp.json().get("value", []):
        map_paths(path + [child["displayName"]], child["id"])


map_paths([ROOT_NAME], root_id)

# ─── Fetch messages from the "2024 Jobs" folder ───────────────────


def fetch_inbox_messages():
    # Page through "2024 Jobs" folder messages
    initial_url = f"{GRAPH_BASE}/users/{USERNAME}/mailFolders/{root_id}/messages"
    initial_params = {
        "$select": "id,subject,from,receivedDateTime,body,conversationId,inReplyTo,parentFolderId",
        "$top": 50,
    }
    all_msgs = []
    next_link = initial_url
    first_call = True

    while next_link:
        resp = requests.get(
            next_link,
            headers=HEADERS,
            params=initial_params if first_call else None,
        )
        first_call = False

        resp.raise_for_status()
        data = resp.json()
        for m in data.get("value", []):
            fid = m.get("parentFolderId")
            path = folder_paths.get(fid, [])
            m["year"] = path[1] if len(path) > 1 else ""
            m["location"] = path[2] if len(path) > 2 else ""
            m["job_num"] = path[3] if len(path) > 3 else ""
            m["thread_id"] = m.get("conversationId")
            m["reply_to"] = m.get("inReplyTo")

            att_url = f"{GRAPH_BASE}/users/{USERNAME}/messages/{m['id']}/attachments"
            att_resp = requests.get(att_url, headers=HEADERS)
            att_resp.raise_for_status()
            m["attachments"] = att_resp.json().get("value", [])

            all_msgs.append(m)
        next_link = data.get("@odata.nextLink")

    return all_msgs

# ─── Build DataFrame ─────────────────────────────────────────────────────────────────────────
msgs = fetch_inbox_messages()
df = pd.DataFrame(msgs)
df["sender"] = df["from"].apply(lambda f: f.get("emailAddress", {}).get("address"))
df["receivedDateTime"] = pd.to_datetime(df["receivedDateTime"])
df["body.content"] = df["body"].apply(lambda b: b.get("content") if isinstance(b, dict) else "")
df = df[[
    "id",
    "subject",
    "sender",
    "receivedDateTime",
    "body.content",
    "thread_id",
    "reply_to",
    "year",
    "location",
    "job_num",
    "attachments",
]]

# ─── Extract attachment text ────────────────────────────────────────────────────
def extract_attachment_text(att):
    url = att.get("@odata.mediaReadLink")
    if not url:
        return ""
    resp = requests.get(url, headers=HEADERS)
    resp.raise_for_status()
    buf = BytesIO(resp.content)
    name = att.get("name", "").lower()
    if name.endswith(".pdf") and _HAS_PDFMINER:
        out = BytesIO()
        extract_text_to_fp(buf, out)
        return out.getvalue().decode(errors="ignore")
    if name.endswith(".docx") and _HAS_DOCX:
        doc = Document(buf)
        return "\n".join(p.text for p in doc.paragraphs)
    if name.endswith(".xlsx"):
        wb = openpyxl.load_workbook(buf, data_only=True)
        texts = []
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                texts.append(" ".join(str(c) for c in row if c is not None))
        return "\n".join(texts)
    if any(name.endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"]):
        img = Image.open(buf)
        return pytesseract.image_to_string(img)
    return ""

df["attachment_text"] = df["attachments"].apply(
    lambda atts: "\n---\n".join(extract_attachment_text(a) for a in (atts or []))
)

# ─── Combine for full-text analysis ─────────────────────────────────────────────
df["full_text"] = df["body.content"].fillna("") + "\n\n" + df["attachment_text"].fillna("")

# ─── 1. Basic Descriptives ────────────────────────────────────────────────────────────────────────
print(f"Total messages: {len(df)}")
print("Top 5 senders:\n", df["sender"].value_counts().head())

# ─── 2. Time Series Plot ────────────────────────────────────────────────────────────────────────

ts = df.set_index("receivedDateTime").resample("W").size()
plt.figure()
ts.plot(title="Emails per Week")
plt.tight_layout()
plt.show()

# ─── 3. Sentiment Analysis ────────────────────────────────────────────────────────────────────────

df["sentiment"] = df["full_text"].apply(lambda t: TextBlob(t).sentiment.polarity)
print("Average sentiment:", df["sentiment"].mean())

# ─── 4. Topic Modeling ────────────────────────────────────────────────────────────────────────
# Vectorize the previews
vec = CountVectorizer(max_df=0.9, min_df=5, stop_words="english")
X = vec.fit_transform(df["full_text"].fillna(""))
lda = LatentDirichletAllocation(n_components=5, random_state=0)
lda.fit(X)
terms = vec.get_feature_names_out()
for idx, topic in enumerate(lda.components_):
    top_terms = [terms[i] for i in topic.argsort()[-10:]]
    print(f"Topic {idx+1}: {', '.join(top_terms)}")

# ─── 5. Save to CSV for further analysis ────────────────────────────────────────────────────────────────────────

output_file = "tmyers_inbox_summary.csv"
try:
    df.to_csv(output_file, index=False)
    print(f"Saved summary CSV: {output_file}")
except PermissionError:
    # Fallback to home directory with timestamp
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fallback = os.path.expanduser(f"~/tmyers_inbox_summary_{ts}.csv")
    df.to_csv(fallback, index=False)
    print(f"Permission denied on '{output_file}'. Saved to fallback path: {fallback}")

_stop_matrix.set()
_matrix_thread.join()
