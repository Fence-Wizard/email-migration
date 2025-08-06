from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any


class EmailAttachment(BaseModel):
    name: str
    contentType: Optional[str] = None
    mediaReadLink: Optional[str] = Field(None, alias='@odata.mediaReadLink')


class EmailMessage(BaseModel):
    id: str
    subject: str
    sender: Dict[str, Any] = Field(alias='from')
    receivedDateTime: str
    conversationId: Optional[str]
    parentFolderId: Optional[str]
    body: Dict[str, Any]
    attachments: List[EmailAttachment] = []

