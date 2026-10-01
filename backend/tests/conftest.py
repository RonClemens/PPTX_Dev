import shutil
import zipfile
from pathlib import Path

import pytest

FIXTURES = Path(__file__).parent / "fixtures"
SAMPLE = FIXTURES / "sample.pptx"


@pytest.fixture()
def sample_copy(tmp_path) -> Path:
    """A throwaway copy of sample.pptx the test is free to load and save."""
    dst = tmp_path / "sample.pptx"
    shutil.copyfile(SAMPLE, dst)
    return dst


def rewrite_zip(src: Path, dst: Path, replace: dict[str, bytes] | None = None, add: dict[str, bytes] | None = None):
    """Copy a zip, replacing/adding named parts (bytes) -- used to build
    fixture variants (e.g. a deck with modern comments) from sample.pptx."""
    replace = replace or {}
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, replace.get(item.filename, zin.read(item.filename)))
        for name, data in (add or {}).items():
            zout.writestr(name, data)
