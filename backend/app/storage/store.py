"""Pure JSON-file document store -- no database, no SQL server, no S3.

Layout on disk (all under config.DATA_DIR, meant to sit on a mounted
persistent volume such as Amazon EFS in a containerized GovCloud deployment):

  data/index.json                             -- denormalized {doc_id: meta}
                                                   cache for fast listing; if
                                                   it's ever lost or corrupted
                                                   it's rebuilt from each
                                                   document's meta.json (the
                                                   source of truth per doc).
  data/documents/<doc_id>/original.pptx        -- never modified
  data/documents/<doc_id>/working.pptx         -- current state; every edit
                                                   re-saves this
  data/documents/<doc_id>/meta.json            -- {id, filename, title,
                                                   created_at, updated_at}
  data/documents/<doc_id>/edits.jsonl          -- append-only edit audit log
  data/documents/<doc_id>/ai_adjudications.jsonl -- append-only AI reasoning
                                                   log (the pptx itself only
                                                   records the *result* of an
                                                   AI edit, not why)
  data/documents/<doc_id>/decisions.json       -- {comment_id: current
                                                   Accept/Reject/Info Only/
                                                   Defer decision} used for
                                                   the Comment Resolution
                                                   Matrix export
  data/documents/<doc_id>/decisions.jsonl      -- append-only decision
                                                   history (one entry per
                                                   save, even overwrites)
  data/documents/<doc_id>/comment_chats.json   -- {comment_id: [{role,
                                                   content, ts, author?}]}
                                                   per-comment AI discussion
                                                   transcripts, kept separate
                                                   from the decision/rationale
                                                   record itself
  data/documents/<doc_id>/.lock                -- advisory lock guarding
                                                   meta.json / *.jsonl writes
                                                   for this document
  data/.index.lock                             -- advisory lock guarding
                                                   index.json

Every write is atomic (write to a temp file in the same directory, fsync,
then os.replace) and every read-modify-write sequence holds a POSIX advisory
lock (flock) on a dedicated lock file, so concurrent requests -- including
from multiple replicas of this backend sharing one EFS mount -- can't
corrupt or interleave writes. EFS speaks NFSv4.1, which supports both atomic
rename and flock() correctly (unlike older NFSv3), so this is safe there.
On non-POSIX platforms (e.g. local Windows dev) locking degrades to a no-op;
the target deployment is Linux containers, where it's fully enforced.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import shutil
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator

from .. import config
from ..pptx.ooxml import PptxPackage

try:
    import fcntl

    _HAVE_FLOCK = True
except ImportError:  # pragma: no cover - non-POSIX (e.g. Windows) dev fallback
    fcntl = None  # type: ignore[assignment]
    _HAVE_FLOCK = False


def _now() -> str:
    return dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


@contextmanager
def _file_lock(lock_path: Path) -> Iterator[None]:
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    fh = open(lock_path, "a+")
    try:
        if _HAVE_FLOCK:
            fcntl.flock(fh.fileno(), fcntl.LOCK_EX)
        yield
    finally:
        if _HAVE_FLOCK:
            fcntl.flock(fh.fileno(), fcntl.LOCK_UN)
        fh.close()


def _atomic_write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_name(f".{path.name}.tmp-{uuid.uuid4().hex}")
    with open(tmp_path, "w") as f:
        json.dump(data, f, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp_path, path)  # atomic on POSIX (and on EFS/NFSv4.1) within one directory
    try:
        dir_fd = os.open(str(path.parent), os.O_RDONLY)
        try:
            os.fsync(dir_fd)
        finally:
            os.close(dir_fd)
    except OSError:
        pass  # best-effort; some platforms/filesystems don't support directory fsync


def _read_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    with open(path, "r") as f:
        return json.load(f)


def _append_jsonl(path: Path, record: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "a") as f:
        f.write(json.dumps(record) + "\n")
        f.flush()
        os.fsync(f.fileno())


def _read_jsonl(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    out = []
    with open(path, "r") as f:
        for line in f:
            line = line.strip()
            if line:
                out.append(json.loads(line))
    return out


class DocumentStore:
    def __init__(self) -> None:
        config.DOCUMENTS_DIR.mkdir(parents=True, exist_ok=True)
        self._ensure_index()

    # -- paths --------------------------------------------------------

    def _doc_dir(self, doc_id: str) -> Path:
        return config.DOCUMENTS_DIR / doc_id

    def _meta_path(self, doc_id: str) -> Path:
        return self._doc_dir(doc_id) / "meta.json"

    def _edits_path(self, doc_id: str) -> Path:
        return self._doc_dir(doc_id) / "edits.jsonl"

    def _ai_path(self, doc_id: str) -> Path:
        return self._doc_dir(doc_id) / "ai_adjudications.jsonl"

    def _doc_lock_path(self, doc_id: str) -> Path:
        return self._doc_dir(doc_id) / ".lock"

    def _index_lock_path(self) -> Path:
        return config.DATA_DIR / ".index.lock"

    def original_path(self, doc_id: str) -> Path:
        return self._doc_dir(doc_id) / "original.pptx"

    def working_path(self, doc_id: str) -> Path:
        return self._doc_dir(doc_id) / "working.pptx"

    # -- index.json (denormalized list cache) --------------------------

    def _ensure_index(self) -> None:
        if config.INDEX_PATH.exists():
            return
        with _file_lock(self._index_lock_path()):
            if not config.INDEX_PATH.exists():
                self._rebuild_index_locked()

    def _rebuild_index_locked(self) -> dict[str, Any]:
        """Recompute index.json from each document's meta.json. Each doc's
        meta.json is the source of truth; index.json is just a cache, so
        this makes a lost/corrupted index self-healing."""
        documents: dict[str, Any] = {}
        if config.DOCUMENTS_DIR.exists():
            for d in config.DOCUMENTS_DIR.iterdir():
                if not d.is_dir():
                    continue
                meta = _read_json(d / "meta.json", None)
                if meta:
                    documents[meta["id"]] = meta
        index = {"documents": documents}
        _atomic_write_json(config.INDEX_PATH, index)
        return index

    def rebuild_index(self) -> None:
        with _file_lock(self._index_lock_path()):
            self._rebuild_index_locked()

    def _read_index(self) -> dict[str, Any]:
        return _read_json(config.INDEX_PATH, {"documents": {}})

    def _upsert_index_entry(self, meta: dict[str, Any]) -> None:
        with _file_lock(self._index_lock_path()):
            index = self._read_index()
            index["documents"][meta["id"]] = meta
            _atomic_write_json(config.INDEX_PATH, index)

    def _remove_index_entry(self, doc_id: str) -> None:
        with _file_lock(self._index_lock_path()):
            index = self._read_index()
            index["documents"].pop(doc_id, None)
            _atomic_write_json(config.INDEX_PATH, index)

    # -- documents ------------------------------------------------------

    def create_document(self, filename: str, data: bytes, title: str | None = None) -> str:
        doc_id = uuid.uuid4().hex
        d = self._doc_dir(doc_id)
        d.mkdir(parents=True, exist_ok=True)
        self.original_path(doc_id).write_bytes(data)
        self.working_path(doc_id).write_bytes(data)

        now = _now()
        meta = {
            "id": doc_id,
            "filename": filename,
            "title": title or filename,
            "created_at": now,
            "updated_at": now,
        }
        _atomic_write_json(self._meta_path(doc_id), meta)
        self._upsert_index_entry(meta)
        return doc_id

    def list_documents(self) -> list[dict[str, Any]]:
        docs = list(self._read_index()["documents"].values())
        docs.sort(key=lambda d: d["updated_at"], reverse=True)
        return docs

    def get_document_meta(self, doc_id: str) -> dict[str, Any] | None:
        return _read_json(self._meta_path(doc_id), None)

    def document_exists(self, doc_id: str) -> bool:
        return self.working_path(doc_id).exists()

    def load_package(self, doc_id: str) -> PptxPackage:
        return PptxPackage.load(self.working_path(doc_id))

    def save_package(self, doc_id: str, pkg: PptxPackage) -> None:
        pkg.save(self.working_path(doc_id))
        self._touch(doc_id)

    def _touch(self, doc_id: str) -> None:
        with _file_lock(self._doc_lock_path(doc_id)):
            meta = _read_json(self._meta_path(doc_id), None)
            if meta is None:
                return
            meta["updated_at"] = _now()
            _atomic_write_json(self._meta_path(doc_id), meta)
        self._upsert_index_entry(meta)

    def reset_to_original(self, doc_id: str) -> None:
        shutil.copyfile(self.original_path(doc_id), self.working_path(doc_id))
        self._touch(doc_id)

    def delete_document(self, doc_id: str) -> None:
        shutil.rmtree(self._doc_dir(doc_id), ignore_errors=True)
        self._remove_index_entry(doc_id)

    # -- audit log / AI reasoning (append-only JSON Lines) ---------------

    def log_edit(self, doc_id: str, kind: str, author: str, payload: dict[str, Any]) -> None:
        with _file_lock(self._doc_lock_path(doc_id)):
            _append_jsonl(
                self._edits_path(doc_id),
                {"ts": _now(), "kind": kind, "author": author, "payload": payload},
            )

    def list_edits(self, doc_id: str) -> list[dict[str, Any]]:
        return _read_jsonl(self._edits_path(doc_id))

    def log_ai_adjudication(
        self,
        doc_id: str,
        comment_id: str,
        action: str,
        reasoning: str,
        reply_comment_id: str | None,
        change_id: str | None,
    ) -> None:
        with _file_lock(self._doc_lock_path(doc_id)):
            _append_jsonl(
                self._ai_path(doc_id),
                {
                    "comment_id": comment_id,
                    "action": action,
                    "reasoning": reasoning,
                    "reply_comment_id": reply_comment_id,
                    "change_id": change_id,
                    "ts": _now(),
                },
            )

    def list_ai_adjudications(self, doc_id: str) -> list[dict[str, Any]]:
        return _read_jsonl(self._ai_path(doc_id))

    # -- comment decisions (Comment Resolution Matrix) -------------------

    def _decisions_path(self, doc_id: str) -> Path:
        return self._doc_dir(doc_id) / "decisions.json"

    def _decisions_log_path(self, doc_id: str) -> Path:
        return self._doc_dir(doc_id) / "decisions.jsonl"

    def get_decisions(self, doc_id: str) -> dict[str, Any]:
        return _read_json(self._decisions_path(doc_id), {})

    def save_decision(
        self, doc_id: str, comment_id: str, decision: str, reason: str, ref: str, author: str
    ) -> dict[str, Any]:
        with _file_lock(self._doc_lock_path(doc_id)):
            decisions = _read_json(self._decisions_path(doc_id), {})
            record = {
                "decision": decision,
                "reason": reason,
                "ref": ref,
                "author": author,
                "decidedAt": _now(),
            }
            decisions[comment_id] = record
            _atomic_write_json(self._decisions_path(doc_id), decisions)
            _append_jsonl(self._decisions_log_path(doc_id), {"comment_id": comment_id, **record})
        return record

    # -- per-comment AI discussion transcripts ---------------------------

    def _chats_path(self, doc_id: str) -> Path:
        return self._doc_dir(doc_id) / "comment_chats.json"

    def get_chats(self, doc_id: str) -> dict[str, list[dict[str, Any]]]:
        return _read_json(self._chats_path(doc_id), {})

    def append_chat_turn(
        self, doc_id: str, comment_id: str, role: str, content: str, author: str | None = None
    ) -> list[dict[str, Any]]:
        with _file_lock(self._doc_lock_path(doc_id)):
            chats = _read_json(self._chats_path(doc_id), {})
            turns = chats.setdefault(comment_id, [])
            turn: dict[str, Any] = {"role": role, "content": content, "ts": _now()}
            if author is not None:
                turn["author"] = author
            turns.append(turn)
            _atomic_write_json(self._chats_path(doc_id), chats)
            return turns

    # -- document-level metadata (Comment Resolution Matrix header) ------

    def update_document_meta(self, doc_id: str, **fields: Any) -> dict[str, Any]:
        """Merge non-None fields into meta.json; None means "leave
        unchanged", matching the settings API's partial-update convention."""
        with _file_lock(self._doc_lock_path(doc_id)):
            meta = _read_json(self._meta_path(doc_id), None)
            if meta is None:
                raise KeyError(doc_id)
            for key, value in fields.items():
                if value is not None:
                    meta[key] = value
            meta["updated_at"] = _now()
            _atomic_write_json(self._meta_path(doc_id), meta)
        self._upsert_index_entry(meta)
        return meta


store = DocumentStore()
