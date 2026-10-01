"""Builds sample.pptx: a small 4-slide deck with a title slide, bullets, a
table, a picture, speaker notes, and PowerPoint-style *legacy* comments
(one threaded reply, one slide-level comment) written as raw XML -- the way
PowerPoint itself would, independent of this app's own comment writer.

Dev-only (needs `pip install python-pptx pillow`); the generated file is
committed, so running the app and tests needs neither.

    python backend/tests/fixtures/make_sample.py
"""
from __future__ import annotations

import io
import shutil
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches, Pt

HERE = Path(__file__).parent
REPO = HERE.parents[2]

EMU_PER_UNIT = 914400 / 576  # legacy comment p:pos units are 1/576 inch


def _picture() -> io.BytesIO:
    img = Image.new("RGB", (640, 360), "#EAF1FB")
    d = ImageDraw.Draw(img)
    for i, (x, label) in enumerate([(40, "Sensor"), (250, "Processor"), (460, "Display")]):
        d.rectangle([x, 120, x + 140, 240], fill="#2F5597", outline="#1F3864", width=3)
        d.text((x + 40, 172), label, fill="white")
        if i < 2:
            d.line([x + 140, 180, x + 210, 180], fill="#1F3864", width=4)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)
    return buf


def build_deck(path: Path) -> dict[str, tuple[int, int]]:
    """Build the comment-free deck; returns {name: (x_emu, y_emu)} anchor
    points inside the shapes the comments will point at."""
    prs = Presentation()

    s1 = prs.slides.add_slide(prs.slide_layouts[0])
    s1.shapes.title.text = "Q3 Project Status Review"
    s1.placeholders[1].text = "Prepared by the Program Office"

    s2 = prs.slides.add_slide(prs.slide_layouts[1])
    s2.shapes.title.text = "Key Achievements"
    body = s2.placeholders[1]
    tf = body.text_frame
    tf.text = "Completed preliminary design review ahead of schedule"
    for line in [
        "Requirement ACA-HRS-0190 verified in lab testing",
        "Reduced unit cost by 12% versus the baseline estimate",
        "Delivered the first prototype to the customer",
    ]:
        tf.add_paragraph().text = line
    s2.notes_slide.notes_text_frame.text = (
        "Mention that the cost reduction is relative to the FY23 baseline."
    )

    s3 = prs.slides.add_slide(prs.slide_layouts[5])
    s3.shapes.title.text = "Schedule"
    rows = [
        ("Milestone", "Date", "Status"),
        ("Critical Design Review", "2025-03-14", "Complete"),
        ("Prototype Delivery", "2025-06-02", "On track"),
        ("Qualification Testing", "2025-09-15", "At risk"),
    ]
    gf = s3.shapes.add_table(len(rows), 3, Inches(0.8), Inches(2.0), Inches(8.4), Inches(2.4))
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            gf.table.cell(r, c).text = val

    s4 = prs.slides.add_slide(prs.slide_layouts[5])
    s4.shapes.title.text = "System Architecture"
    s4.shapes.add_picture(_picture(), Inches(1.5), Inches(1.9), width=Inches(7.0))

    prs.save(path)

    def inside(shape):
        return (shape.left + Inches(0.15), shape.top + Inches(0.15))

    return {
        "body": inside(body),
        "table": inside(gf),
        "slide4": (Inches(0.3), Inches(6.6)),
    }


def _legacy_comments_xml(entries: list[tuple[int, int, str, int, int, str, str | None]]) -> bytes:
    """entries: (authorId, idx, dt, x_emu, y_emu, text, parent 'aid-idx' or None)"""
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<p:cmLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
           'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">']
    for aid, idx, dt, x, y, text, parent in entries:
        ext = ""
        if parent:
            paid, pidx = parent.split("-")
            ext = ('<p:extLst><p:ext uri="{C676402C-5697-4E1C-873F-D02D1690AC5C}">'
                   '<p15:threadingInfo xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" '
                   f'timeZoneBias="0"><p15:parentCm authorId="{paid}" idx="{pidx}"/></p15:threadingInfo>'
                   '</p:ext></p:extLst>')
        out.append(
            f'<p:cm authorId="{aid}" dt="{dt}" idx="{idx}">'
            f'<p:pos x="{round(x / EMU_PER_UNIT)}" y="{round(y / EMU_PER_UNIT)}"/>'
            f'<p:text>{text}</p:text>{ext}</p:cm>'
        )
    out.append("</p:cmLst>")
    return "".join(out).encode("utf-8")


def add_comments(src: Path, dst: Path, pts: dict[str, tuple[int, int]]) -> None:
    authors = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:cmAuthorLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cmAuthor id="0" name="Alex Reviewer" initials="AR" lastIdx="2" clrIdx="0"/>'
        '<p:cmAuthor id="1" name="Pat Author" initials="PA" lastIdx="1" clrIdx="1"/>'
        "</p:cmAuthorLst>"
    ).encode("utf-8")
    comments = {
        # slide2.xml -> comment on the bullets; Pat replies
        "ppt/comments/comment1.xml": _legacy_comments_xml([
            (0, 1, "2025-02-03T09:15:00.000", *pts["body"],
             "Can we cite the test report number for the ACA-HRS-0190 verification?", None),
            (1, 1, "2025-02-04T14:02:00.000", *pts["body"],
             "Good catch - it is TR-2025-014. I will add it to the speaker notes.", "0-1"),
        ]),
        "ppt/comments/comment2.xml": _legacy_comments_xml([
            (0, 2, "2025-02-03T09:20:00.000", *pts["table"],
             "Please confirm the Qualification Testing date - it conflicts with the integrated master schedule.", None),
        ]),
        "ppt/comments/comment3.xml": _legacy_comments_xml([
            (0, 3, "2025-02-03T09:31:00.000", *pts["slide4"],
             "The diagram is hard to read when projected; consider a larger version.", None),
        ]),
    }
    slide_for_comment = {"ppt/comments/comment1.xml": "slide2", "ppt/comments/comment2.xml": "slide3",
                         "ppt/comments/comment3.xml": "slide4"}

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                extra = ('<Override PartName="/ppt/commentAuthors.xml" ContentType="application/'
                         'vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/>')
                for part in comments:
                    extra += (f'<Override PartName="/{part}" ContentType="application/'
                              'vnd.openxmlformats-officedocument.presentationml.comments+xml"/>')
                data = data.replace(b"</Types>", extra.encode() + b"</Types>")
            elif item.filename == "ppt/_rels/presentation.xml.rels":
                rel = ('<Relationship Id="rIdAuthors" Type="http://schemas.openxmlformats.org/officeDocument/'
                       '2006/relationships/commentAuthors" Target="commentAuthors.xml"/>')
                data = data.replace(b"</Relationships>", rel.encode() + b"</Relationships>")
            else:
                for part, slide in slide_for_comment.items():
                    if item.filename == f"ppt/slides/_rels/{slide}.xml.rels":
                        rel = ('<Relationship Id="rIdCmt" Type="http://schemas.openxmlformats.org/officeDocument/'
                               f'2006/relationships/comments" Target="../comments/{Path(part).name}"/>')
                        data = data.replace(b"</Relationships>", rel.encode() + b"</Relationships>")
            zout.writestr(item, data)
        zout.writestr("ppt/commentAuthors.xml", authors)
        for part, data in comments.items():
            zout.writestr(part, data)


if __name__ == "__main__":
    tmp = HERE / "_sample_nocomments.pptx"
    pts = build_deck(tmp)
    out = HERE / "sample.pptx"
    add_comments(tmp, out, pts)
    tmp.unlink()
    (REPO / "samples").mkdir(exist_ok=True)
    shutil.copyfile(out, REPO / "samples" / "sample.pptx")
    print("wrote", out, "and", REPO / "samples" / "sample.pptx")
