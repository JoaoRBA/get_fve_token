
import os
import re
import time
from datetime import datetime, timedelta, timezone
from typing import Optional, Dict

import requests
import msal
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()
DEFAULT_REGEX_TOKEN = r"\b(\d{6})\b"

class GraphMailClient:
    def __init__(self, tenant_id: Optional[str] = None, client_id: Optional[str] = None,
                 client_secret: Optional[str] = None, session: Optional[requests.Session] = None,
                 timeout_s: int = 15):
        self.tenant_id = tenant_id or os.environ["GRAPH_TENANT_ID"]
        self.client_id = client_id or os.environ["GRAPH_CLIENT_ID"]
        self.client_secret = client_secret or os.environ["GRAPH_CLIENT_SECRET"]
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.scopes = ["https://graph.microsoft.com/.default"]
        self._msal_app = msal.ConfidentialClientApplication(
            self.client_id, authority=self.authority, client_credential=self.client_secret
        )
        self.session = session or requests.Session()
        self.timeout_s = timeout_s

    def _get_headers(self) -> Dict[str, str]:
        token = self._msal_app.acquire_token_silent(self.scopes, account=None)
        if not token:
            token = self._msal_app.acquire_token_for_client(scopes=self.scopes)
        if "access_token" not in token:
            raise RuntimeError(f"Falha OAuth: {token.get('error_description') or token}")
        return {"Authorization": f"Bearer {token['access_token']}", "ConsistencyLevel": "eventual"}

    @staticmethod
    def _html_to_text(content: str) -> str:
        return BeautifulSoup(content, "html.parser").get_text(" ", strip=True)

    def fetch_token(self, mailbox: str, subject_keyword: str, timeout_seconds: int = 300,
                    regex_token: str = DEFAULT_REGEX_TOKEN, mark_read: bool = True) -> Optional[str]:
        start_utc = datetime.now(timezone.utc)
        end_utc = start_utc + timedelta(seconds=timeout_seconds)

        #safe_subject = subject_keyword.replace("'", "''")
        params = {
            "$select": "id,subject,receivedDateTime,isRead,body",
            "$orderby": "receivedDateTime desc",
            "$top": "25",
            "$count": "true",
            "$filter": (
                f"isRead eq false and "
                f"receivedDateTime ge {start_utc.isoformat()} and "
                f"contains(subject,'{subject_keyword}')"
            ),
        }
        base_url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages"

        while datetime.now(timezone.utc) < end_utc:
            resp = self.session.get(base_url, headers=self._get_headers(), params=params, timeout=self.timeout_s)
            if resp.status_code in (429, 500, 502, 503, 504):
                time.sleep(2); continue
            if resp.status_code not in (200, 403):
                time.sleep(1); continue

            data = resp.json()
            for m in data.get("value", []):
                body = m.get("body") or {}
                content = body.get("content") or ""
                text = self._html_to_text(content) if (body.get("contentType") or "").lower() == "html" else content

                hit = re.search(regex_token, text)
                if not hit:
                    continue

                token_val = hit.group(1)
                if mark_read:
                    self._mark_as_read(mailbox, m["id"])
                return token_val

            time.sleep(1)

        return None

    def _mark_as_read(self, mailbox: str, message_id: str) -> None:
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages/{message_id}"
        r = self.session.patch(
            url,
            headers={**self._get_headers(), "Content-Type": "application/json"},
            json={"isRead": True},
            timeout=self.timeout_s,
        )
        # Ignora falhas de marcação para não quebrar o fluxo
        _ = r.status_code


def fetch_token_simple(mailbox: str, subject_keyword: str, timeout_seconds: int = 300,
                       regex_token: str = DEFAULT_REGEX_TOKEN, mark_read: bool = True,
                       tenant_id: Optional[str] = None, client_id: Optional[str] = None,
                       client_secret: Optional[str] = None) -> Optional[str]:
    client = GraphMailClient(tenant_id=tenant_id, client_id=client_id, client_secret=client_secret)
    return client.fetch_token(
        mailbox=mailbox, subject_keyword=subject_keyword, timeout_seconds=timeout_seconds,
        regex_token=regex_token, mark_read=mark_read
    )
