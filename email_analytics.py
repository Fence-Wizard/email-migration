import os
import asyncio
import msal
import pandas as pd
import structlog
from dotenv import load_dotenv

from models import EmailMessage

import httpx
from tenacity import AsyncRetrying, retry_if_exception_type, wait_exponential, stop_after_attempt
from transformers import pipeline
from pymdptoolbox.mdp import ValueIteration


load_dotenv()
structlog.configure(processors=[structlog.processors.JSONRenderer()])
logger = structlog.get_logger()


CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
USERNAME = os.getenv("USERNAME", "tmyers@hurricanefence.com")
SCOPES = ["https://graph.microsoft.com/.default"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)
token = app.acquire_token_for_client(SCOPES)["access_token"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
HEADERS = {"Authorization": f"Bearer {token}"}


sentiment_analyzer = pipeline("sentiment-analysis", truncation=True)


def analyze_sentiment(text: str) -> float:
    if not text:
        return 0.0
    result = sentiment_analyzer(text[:512])[0]
    return result["score"] if result["label"].startswith("POS") else -result["score"]


async def fetch_messages_from_folder(folder_id: str):
    url = f"{GRAPH_BASE}/users/{USERNAME}/mailFolders/{folder_id}/messages"
    params = {
        "$select": "id,subject,from,receivedDateTime,body,conversationId,parentFolderId",
        "$top": 50,
    }
    async with httpx.AsyncClient() as client:
        async for attempt in AsyncRetrying(
            retry=retry_if_exception_type(httpx.HTTPError),
            wait=wait_exponential(min=1, max=10),
            stop=stop_after_attempt(3),
        ):
            with attempt:
                resp = await client.get(url, headers=HEADERS, params=params)
    resp.raise_for_status()
    data = resp.json()
    for raw in data.get("value", []):
        yield EmailMessage.parse_obj(raw)


def analyze_pipeline(df: pd.DataFrame) -> pd.DataFrame:
    df["full_text"] = df["body"].apply(lambda b: b.get("content", ""))
    df["sentiment"] = df["full_text"].apply(analyze_sentiment)
    return df


async def main():
    folder_id = os.getenv("MAIL_FOLDER_ID", "Inbox")
    messages = []
    async for m in fetch_messages_from_folder(folder_id):
        messages.append(m.dict())

    df = pd.DataFrame(messages)
    df = analyze_pipeline(df)

    n_states = 5
    n_actions = 2
    P = [
        [[1.0 / n_states for _ in range(n_states)] for _ in range(n_states)]
        for _ in range(n_actions)
    ]
    R = [[0.0 for _ in range(n_states)] for _ in range(n_actions)]
    vi = ValueIteration(P, R, discount=0.95)
    vi.run()
    policy = vi.policy
    logger.info("Computed optimal policy", policy=policy.tolist())

    print(df.head())


if __name__ == "__main__":
    asyncio.run(main())

