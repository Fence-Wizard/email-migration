#!/usr/bin/env python3
"""
Modernized PhD-level email analytics pipeline:

1. Modular, Type-Annotated Architecture
   • pydantic models for all message schemas.
2. Async I/O & Concurrency
   • httpx.AsyncClient + asyncio for non-blocking Graph calls.
3. Robust Paging & Error-Handling
   • async paginate() for @odata.nextLink.
   • tenacity retries with exponential backoff.
4. Provenance & Reproducibility
   • config.toml, log Git SHA, checksum requirements.
5. Advanced NLP & Analytical Layer
   • transformers embeddings + HDBSCAN clustering.
6. Game-Theory & Process Mining
   • pymdptoolbox MDP stubs for buyer/vendor process modeling.
"""

import asyncio
import os
from dotenv import load_dotenv
load_dotenv()  # make os.getenv() read .env immediately

import toml
import structlog
import msal
from tenacity import retry, stop_after_attempt, wait_exponential
from httpx import AsyncClient, HTTPError
from pydantic import BaseModel
from datetime import datetime
from typing import List
from transformers import AutoTokenizer, AutoModel
import numpy as np
try:
    import hdbscan
except ImportError:
    hdbscan = None
try:
    from pymdptoolbox.mdp import ValueIteration
except ImportError:
    ValueIteration = None

# structured logger
logger = structlog.get_logger()

class EmailMessage(BaseModel):
    id: str
    subject: str
    sender: str
    received: datetime
    body: str

async def async_paginate(client: AsyncClient, url: str, params: dict):
    items = []
    while url:
        resp = await client.get(url, params=params)
        try:
            resp.raise_for_status()
        except HTTPError:
            try:
                body = resp.json()
            except ValueError:
                body = resp.text
            logger.error("HTTP request failed", url=str(resp.url), body=body)
            raise
        data = resp.json()
        items.extend(data.get('value', []))
        url = data.get('@odata.nextLink')
        params = {}
    return items

def acquire_token(graph_cfg: dict) -> str:
    """Obtain an OAuth access token for Microsoft Graph."""
    authority = f"https://login.microsoftonline.com/{graph_cfg['tenant_id']}"
    scopes = ["https://graph.microsoft.com/.default"]
    # Prefer Client Credentials flow if a client_secret is present
    if graph_cfg.get('client_secret'):
        app = msal.ConfidentialClientApplication(
            graph_cfg['client_id'],
            authority=authority,
            client_credential=graph_cfg['client_secret'],
        )
        result = app.acquire_token_for_client(scopes=scopes)
    # Fallback to ROPC flow only if no client_secret
    elif graph_cfg.get('username') and graph_cfg.get('password'):
        app = msal.PublicClientApplication(graph_cfg['client_id'], authority=authority)
        result = app.acquire_token_by_username_password(
            graph_cfg['username'],
            graph_cfg['password'],
            scopes=scopes,
        )
    else:
        raise RuntimeError('No valid authentication credentials found in config.')
    if 'access_token' not in result:
        raise RuntimeError(f"Token acquisition failed: {result.get('error_description')}")
    return result['access_token']

@retry(stop=stop_after_attempt(3), wait=wait_exponential(min=2, max=10))
async def fetch_inbox(config: dict) -> List[EmailMessage]:
    token = acquire_token(config['graph'])
    headers = {"Authorization": f"Bearer {token}"}
    async with AsyncClient(base_url=config['graph']['base_url'], headers=headers) as client:
        raw = await async_paginate(client, '/me/mailFolders/Inbox/messages', {'$top':50})
    return [
        EmailMessage(
            id=m['id'],
            subject=m.get('subject',''),
            sender=m['from']['emailAddress']['address'],
            received=datetime.fromisoformat(m['receivedDateTime']),
            body=m.get('bodyPreview','')
        )
        for m in raw
    ]

def main():
    # load config from config.toml, or fall back to environment
    try:
        cfg = toml.load('config.toml')
    except (FileNotFoundError, toml.decoder.TomlDecodeError):
        logger.warning("config.toml not found or invalid; falling back to .env settings")
        cfg = {
            "graph": {
                # try AZ_*, otherwise fall back to your existing names
                "client_id":     os.getenv("AZ_CLIENT_ID")     or os.getenv("CLIENT_ID"),
                "client_secret": os.getenv("AZ_CLIENT_SECRET") or os.getenv("CLIENT_SECRET"),
                "tenant_id":     os.getenv("AZ_TENANT_ID")     or os.getenv("TENANT_ID"),
                "username":      os.getenv("AZ_USERNAME")      or os.getenv("MAIL_USER"),
                "password":      os.getenv("AZ_PASSWORD")      or os.getenv("MAIL_PASSWORD"),
                "base_url":      os.getenv("AZ_BASE_URL", "https://graph.microsoft.com/v1.0"),
            },
            "analysis": {
                "top_n": int(os.getenv("ANALYSIS_TOP_N", "5")),
            },
            "nlp": {
                "model": os.getenv("NLP_MODEL", "distilbert-base-uncased"),
            },
            "meta": {
                "git_sha": os.getenv("GIT_SHA", ""),
            },
        }

    logger.info("Starting pipeline", git_sha=cfg.get("meta", {}).get("git_sha", ""))

    emails = asyncio.run(fetch_inbox(cfg))

    # NLP embedding
    tok = AutoTokenizer.from_pretrained(cfg['nlp']['model'])
    mdl = AutoModel.from_pretrained(cfg['nlp']['model'])
    texts = [e.body for e in emails]
    enc = tok(texts, return_tensors='pt', padding=True, truncation=True)
    out = mdl(**enc)
    embs = out.last_hidden_state.mean(dim=1).detach().numpy()
    # perform clustering on the embeddings (if available)
    if hdbscan is None:
        logger.warning("hdbscan not available; skipping clustering step")
        clusters = np.zeros(len(embs), dtype=int)
    else:
        clusters = hdbscan.HDBSCAN(min_cluster_size=5).fit_predict(embs)
        logger.info("Clusters discovered", clusters=np.unique(clusters))

    # Game-theory / MDP stub (only if pymdptoolbox is installed)
    if ValueIteration is None:
        logger.warning("pymdptoolbox not installed; skipping MDP analysis")
    else:
        # placeholder P, R to demonstrate usage
        n_states, n_actions = 5, 2
        P = [[[1.0 / n_states for _ in range(n_states)] for _ in range(n_states)] for _ in range(n_actions)]
        R = [[0.0 for _ in range(n_states)] for _ in range(n_actions)]
        vi = ValueIteration(P, R, discount=0.95)
        vi.run()
        logger.info("Computed MDP policy", policy=vi.policy.tolist())

if __name__ == '__main__':
    main()

