import importlib
import json
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

FIXTURE = Path(__file__).parent / "fixtures" / "sample.pptx"


@pytest.fixture()
def store(tmp_path, monkeypatch):
    # DATA_DIR is read once at import time by app.config, and app.storage.store
    # binds `config` (and computes its own module-level `store` singleton) at
    # its own import time -- so isolating one test's data dir means dropping
    # every already-imported `app.*` module and reimporting fresh under the
    # env var, the same pattern test_api.py uses. A stale, collection-time
    # `from app.storage.store import DocumentStore` would silently keep
    # pointing at whatever data dir was active when this module was first
    # imported (the real repo ./data dir, in a full test-suite run) instead
    # of this test's tmp_path.
    monkeypatch.setenv("PPTX_DEV_DATA_DIR", str(tmp_path))
    for mod in list(sys.modules):
        if mod.startswith("app"):
            del sys.modules[mod]
    store_mod = importlib.import_module("app.storage.store")
    return store_mod.DocumentStore()


def _upload(store) -> str:
    return store.create_document("sample.pptx", FIXTURE.read_bytes())


def test_create_writes_meta_json_and_index(store):
    doc_id = _upload(store)

    meta_path = store._meta_path(doc_id)
    assert meta_path.exists()
    meta = json.loads(meta_path.read_text())
    assert meta["id"] == doc_id
    assert meta["filename"] == "sample.pptx"

    from app import config

    index = json.loads(config.INDEX_PATH.read_text())
    assert doc_id in index["documents"]
    assert index["documents"][doc_id]["filename"] == "sample.pptx"


def test_list_documents_sorted_by_updated_at_desc(store):
    a = _upload(store)
    b = _upload(store)
    store._touch(a)  # bump a's updated_at so it sorts first
    docs = store.list_documents()
    assert docs[0]["id"] == a
    assert docs[1]["id"] == b


def test_save_package_touches_meta_and_index(store):
    doc_id = _upload(store)
    before = store.get_document_meta(doc_id)["updated_at"]

    pkg = store.load_package(doc_id)
    store.save_package(doc_id, pkg)

    after_meta = store.get_document_meta(doc_id)["updated_at"]
    after_index = store._read_index()["documents"][doc_id]["updated_at"]
    assert after_meta == after_index
    assert after_meta >= before


def test_reset_to_original_restores_bytes(store):
    doc_id = _upload(store)
    store.working_path(doc_id).write_bytes(b"corrupted")
    store.reset_to_original(doc_id)
    assert store.working_path(doc_id).read_bytes() == FIXTURE.read_bytes()


def test_delete_document_removes_dir_and_index_entry(store):
    doc_id = _upload(store)
    store.delete_document(doc_id)
    assert not store._doc_dir(doc_id).exists()
    assert doc_id not in store._read_index()["documents"]
    assert store.get_document_meta(doc_id) is None


def test_edit_log_is_ordered_jsonl(store):
    doc_id = _upload(store)
    store.log_edit(doc_id, "replace_run_text", "You", {"run_id": "b0.r0"})
    store.log_edit(doc_id, "accept_change", "You", {"change_id": "1"})

    edits = store.list_edits(doc_id)
    assert [e["kind"] for e in edits] == ["replace_run_text", "accept_change"]
    assert edits[0]["payload"] == {"run_id": "b0.r0"}

    raw_lines = store._edits_path(doc_id).read_text().strip().splitlines()
    assert len(raw_lines) == 2
    for line in raw_lines:
        json.loads(line)  # each line is independently valid JSON


def test_ai_adjudication_log(store):
    doc_id = _upload(store)
    store.log_ai_adjudication(doc_id, "0", "edit", "clear reasoning", "1", "2")
    entries = store.list_ai_adjudications(doc_id)
    assert len(entries) == 1
    assert entries[0]["action"] == "edit"
    assert entries[0]["reasoning"] == "clear reasoning"


def test_index_self_heals_when_missing(store):
    from app import config

    doc_id = _upload(store)
    config.INDEX_PATH.unlink()
    assert not config.INDEX_PATH.exists()

    healed = type(store)()  # fresh instance; constructor rebuilds index if missing
    docs = healed.list_documents()
    assert len(docs) == 1
    assert docs[0]["id"] == doc_id


def test_rebuild_index_recovers_from_corruption(store):
    from app import config

    doc_id = _upload(store)
    config.INDEX_PATH.write_text("{not valid json")

    store.rebuild_index()
    docs = store.list_documents()
    assert len(docs) == 1
    assert docs[0]["id"] == doc_id


def test_concurrent_log_edit_no_lost_writes(store):
    doc_id = _upload(store)

    def write(i: int) -> None:
        store.log_edit(doc_id, "stress", "tester", {"i": i})

    with ThreadPoolExecutor(max_workers=16) as ex:
        list(ex.map(write, range(100)))

    edits = store.list_edits(doc_id)
    assert len(edits) == 100
    assert {e["payload"]["i"] for e in edits} == set(range(100))
