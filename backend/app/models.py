from typing import Literal

from pydantic import BaseModel


class CommentReplyRequest(BaseModel):
    text: str
    author: str = "You"


class CommentResolveRequest(BaseModel):
    done: bool = True
    author: str = "Unknown"


class AdjudicateRequest(BaseModel):
    author: str = "AI Assistant"


class SetApiKeyRequest(BaseModel):
    # None on any field means "leave unchanged" so the API key, base URL, and
    # model can be saved independently from the settings modal's separate
    # forms; "" clears that field back to its environment-variable default.
    api_key: str | None = None
    base_url: str | None = None
    model: str | None = None


class TextRangeEditRequest(BaseModel):
    start_run_id: str
    start_offset: int
    end_run_id: str
    end_offset: int
    new_text: str
    author: str


class CreateCommentRequest(BaseModel):
    """A new top-level comment: either on a text selection (the run range
    the frontend's selection popup computes) or on a slide / shape directly.
    PowerPoint comments attach to a slide position, so a selection comment
    is stored against the selection's shape and remembers the selected text."""

    text: str
    author: str
    start_run_id: str | None = None
    start_offset: int | None = None
    end_run_id: str | None = None
    end_offset: int | None = None
    slide_id: str | None = None
    shape_id: str | None = None


class ApplyAdjudicationRequest(BaseModel):
    """Apply a (possibly human-edited) adjudication decision -- either one
    the AI just suggested via /suggest-adjudications, or one typed by hand
    -- without re-invoking the AI."""

    action: str  # "edit" | "reply" | "no_change"
    reply: str = ""
    replacement_text: str | None = None
    # Verbatim text currently on the slide that replacement_text replaces;
    # only needed when the comment is attached to a whole shape rather than
    # to a selected span.
    original_text: str | None = None
    reasoning: str = ""
    author: str = "AI Assistant"


class AiReviewRequest(BaseModel):
    author: str = "AI Reviewer"


class SaveDecisionRequest(BaseModel):
    """Formal comment adjudication for the Comment Resolution Matrix export --
    distinct from the pptx-level resolve/reopen toggle, which only affects the
    PowerPoint comment thread's own resolved flag."""

    decision: Literal["accept", "reject", "info_only", "defer"]
    reason: str
    ref: str = ""  # external spec/requirement reference (e.g. "3.2.2"); manual, not derived
    author: str = "Unknown"


class SetDocumentMetaRequest(BaseModel):
    # None means "leave unchanged", matching SetApiKeyRequest's convention.
    doc_ref: str | None = None
    rev: str | None = None


class ChatMessageRequest(BaseModel):
    message: str
    author: str = "Unknown"
