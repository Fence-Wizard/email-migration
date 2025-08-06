import os
import json
import requests
import msal
import pandas as pd
import matplotlib.pyplot as plt
from textblob import TextBlob
from sklearn.decomposition import LatentDirichletAllocation
from sklearn.feature_extraction.text import CountVectorizer
from dotenv import load_dotenv

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

# ─── Fetch All Inbox Messages ─────────────────────────────────────────────────────────────────────────

def fetch_inbox_messages():
    # Initial endpoint and query parameters
    initial_url = f"https://graph.microsoft.com/v1.0/users/{USERNAME}/mailFolders/Inbox/messages"
    initial_params = {
        "$select": "id,subject,from,receivedDateTime,bodyPreview",
        "$top": 50
    }
    all_msgs = []
    next_link = initial_url
    first_call = True

    # Page through results: apply params only on the first call
    while next_link:
        if first_call:
            resp = requests.get(next_link, headers=HEADERS, params=initial_params)
            first_call = False
        else:
            resp = requests.get(next_link, headers=HEADERS)

        resp.raise_for_status()
        data = resp.json()
        all_msgs.extend(data.get("value", []))
        next_link = data.get("@odata.nextLink")

    return all_msgs

# ─── Build DataFrame ─────────────────────────────────────────────────────────────────────────
msgs = fetch_inbox_messages()
df = pd.DataFrame(msgs)
# normalize sender
df["sender"] = df["from"].apply(lambda f: f.get("emailAddress", {}).get("address"))
df["receivedDateTime"] = pd.to_datetime(df["receivedDateTime"])
df = df[["id","subject","sender","receivedDateTime","bodyPreview"]]

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

df["sentiment"] = df["bodyPreview"].apply(lambda t: TextBlob(t).sentiment.polarity)
print("Average sentiment:", df["sentiment"].mean())

# ─── 4. Topic Modeling ────────────────────────────────────────────────────────────────────────
# Vectorize the previews
vec = CountVectorizer(max_df=0.9, min_df=5, stop_words="english")
X = vec.fit_transform(df["bodyPreview"].fillna(""))
lda = LatentDirichletAllocation(n_components=5, random_state=0)
lda.fit(X)
terms = vec.get_feature_names_out()
for idx, topic in enumerate(lda.components_):
    top_terms = [terms[i] for i in topic.argsort()[-10:]]
    print(f"Topic {idx+1}: {', '.join(top_terms)}")

# ─── 5. Save to CSV for further analysis ────────────────────────────────────────────────────────────────────────

df.to_csv("tmyers_inbox_summary.csv", index=False)
print("Saved summary CSV: tmyers_inbox_summary.csv")
