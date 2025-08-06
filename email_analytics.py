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
   • environment variables, log Git SHA, checksum requirements.
5. Advanced NLP & Analytical Layer
   • transformers embeddings + HDBSCAN clustering.
6. Game-Theory & Process Mining
   • pymdptoolbox MDP stubs for buyer/vendor process modeling.
"""

import asyncio
import os
import json
import re
from dotenv import load_dotenv

import structlog
import msal
from tenacity import retry, stop_after_attempt, wait_exponential
from httpx import AsyncClient, HTTPError
from pydantic import BaseModel
from datetime import datetime
from typing import List, Optional
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

def _resolve_authority(graph_cfg: dict) -> str:
    """Derive a usable Azure AD authority URL.

    If ``AUTHORITY`` is supplied, that takes precedence. Otherwise we validate
    ``TENANT_ID`` and fall back to the multi-tenant ``common`` endpoint when a
    proper tenant cannot be determined. This mirrors Azure's own guidance and
    prevents hard failures when a placeholder or malformed tenant is supplied.
    """

    if graph_cfg.get('authority'):
        return graph_cfg['authority']

    tenant_id = graph_cfg.get('tenant_id')
    if tenant_id and re.fullmatch(r"[0-9a-fA-F-]{36}", tenant_id):
        return f"https://login.microsoftonline.com/{tenant_id}"

    logger.warning("TENANT_ID missing or invalid; using 'common' authority")
    return "https://login.microsoftonline.com/common"


def acquire_token(graph_cfg: dict) -> str:
    """Obtain an OAuth access token for Microsoft Graph."""
    authority = _resolve_authority(graph_cfg)
    scopes = ["https://graph.microsoft.com/.default"]
    auth_mode = graph_cfg.get('auth_mode', 'app')
    if auth_mode == 'app':
        if not graph_cfg.get('client_secret'):
            raise RuntimeError('client_secret required for app auth_mode')
        app = msal.ConfidentialClientApplication(
            graph_cfg['client_id'],
            authority=authority,
            client_credential=graph_cfg['client_secret'],
        )
        result = app.acquire_token_for_client(scopes=scopes)
    elif auth_mode == 'delegated':
        if not (graph_cfg.get('username') and graph_cfg.get('password')):
            raise RuntimeError('username and password required for delegated auth_mode')
        app = msal.PublicClientApplication(graph_cfg['client_id'], authority=authority)
        result = app.acquire_token_by_username_password(
            graph_cfg['username'],
            graph_cfg['password'],
            scopes=scopes,
        )
    else:
        raise RuntimeError(f"Unsupported auth_mode: {auth_mode}")
    if 'access_token' not in result:
        raise RuntimeError(f"Token acquisition failed: {result.get('error_description')}")
    return result['access_token']

@retry(stop=stop_after_attempt(3), wait=wait_exponential(min=2, max=10))
async def fetch_inbox(config: dict, user_id: Optional[str] = None) -> List[EmailMessage]:
    token = acquire_token(config['graph'])
    headers = {"Authorization": f"Bearer {token}"}
    async with AsyncClient(base_url=config['graph']['base_url'], headers=headers) as client:
        if config['graph'].get('client_secret'):
            if not user_id:
                raise RuntimeError("user_id is required for app-only authentication")
            endpoint = f"/users/{user_id}/mailFolders/Inbox/messages"
        else:
            endpoint = "/me/mailFolders/Inbox/messages"
        raw = await async_paginate(client, endpoint, {'$top':50})
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
    # Load environment variables from a local .env file
    script_dir = os.path.dirname(os.path.realpath(__file__))
    dotenv_file = os.path.join(script_dir, '.env')
    load_dotenv(dotenv_file)

    cfg = {
        'graph': {
            'client_id':     os.getenv('CLIENT_ID'),
            'client_secret': os.getenv('CLIENT_SECRET'),
            'tenant_id':     os.getenv('TENANT_ID'),
            'authority':     os.getenv('AUTHORITY'),
            'base_url':      os.getenv('GRAPH_BASE_URL', 'https://graph.microsoft.com/v1.0'),
            'auth_mode':     os.getenv('AZ_AUTH_MODE', 'app'),
            'username':      os.getenv('AZ_USERNAME'),
            'password':      os.getenv('AZ_PASSWORD'),
            'user_id':       os.getenv('AZ_USER_ID'),
        },
        'mail': {
            'user':        os.getenv('MAIL_USER'),
            'folder_path': json.loads(os.getenv('MAIL_FOLDER_PATH', '[]')),
        },
        'analysis': {'top_n': int(os.getenv('TOP_N', '5'))},
        'nlp':      {'model': os.getenv('NLP_MODEL', 'distilbert-base-uncased')},
        'meta':     {'git_sha': os.getenv('GIT_SHA', '')},
    }

    logger.info("Starting pipeline", git_sha=cfg.get('meta', {}).get('git_sha', ''))

    # If running in app-only mode, default user_id to the configured username
    if cfg['graph'].get('client_secret'):
        user_id = cfg['graph'].get('user_id') or cfg['graph'].get('username')
        if not user_id:
            raise RuntimeError("user_id is required for app-only authentication")
    else:
        user_id = None
    emails = asyncio.run(fetch_inbox(cfg, user_id=user_id))

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

