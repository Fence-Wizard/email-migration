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

# ─── Map folder IDs to path elements ──────────────────────────────
root = requests.get(f"{GRAPH_BASE}/users/{USERNAME}/mailFolders/Inbox", headers=HEADERS)
root.raise_for_status()
root_id = root.json()["id"]

folder_paths = {}
def map_paths(path, fid):
    folder_paths[fid] = path
    url = f"{GRAPH_BASE}/users/{USERNAME}/mailFolders/{fid}/childFolders"
    resp = requests.get(url, headers=HEADERS)
    resp.raise_for_status()
    for child in resp.json().get("value", []):
        map_paths(path + [child["displayName"]], child["id"])

map_paths(["Inbox"], root_id)

# ─── Fetch All Inbox Messages ─────────────────────────────────────────────────────────────────────────

def fetch_inbox_messages():
    # Initial endpoint and query parameters
    initial_url = f"{GRAPH_BASE}/users/{USERNAME}/mailFolders/Inbox/messages"
    # only select message core fields here—no expand on attachments
    initial_params = {
        "$select": "id,subject,from,receivedDateTime,bodyPreview,parentFolderId",
        "$top": 50
    }
    all_msgs = []
    next_link = initial_url
    first_call = True

    # Page through results: apply params only on the first call
    while next_link:
        resp = requests.get(
            next_link,
            headers=HEADERS,
            params=initial_params if first_call else None
        )
        first_call = False

        resp.raise_for_status()
        data = resp.json()
        for m in data.get("value", []):
            # annotate folder context
            fid = m.get("parentFolderId")
            path = folder_paths.get(fid, [])
            m["year"]     = path[1] if len(path) > 1 else ""
            m["location"] = path[2] if len(path) > 2 else ""
            m["job_num"]  = path[3] if len(path) > 3 else ""

            # now fetch attachments metadata separately
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
# normalize sender
df["sender"] = df["from"].apply(lambda f: f.get("emailAddress", {}).get("address"))
df["receivedDateTime"] = pd.to_datetime(df["receivedDateTime"])
df = df[["id","subject","sender","receivedDateTime","bodyPreview","year","location","job_num","attachments"]]

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
    return ""

df["attachment_text"] = df["attachments"].apply(
    lambda atts: "\n---\n".join(extract_attachment_text(a) for a in (atts or []))
)

# ─── Combine for full-text analysis ─────────────────────────────────────────────
df["full_text"] = (df["bodyPreview"].fillna("") + "\n" + df["attachment_text"]).fillna("")

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

df.to_csv("tmyers_inbox_summary.csv", index=False)
print("Saved summary CSV: tmyers_inbox_summary.csv")

_stop_matrix.set()
_matrix_thread.join()
