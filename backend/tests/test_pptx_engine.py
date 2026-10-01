"""Engine-level tests: parse sample.pptx (whose legacy comments were written
as raw XML the way PowerPoint writes them, independent of this app's own
comment writer), then exercise the writer and round-trip fidelity."""
import zipfile

import pytest
from lxml import etree

from app.pptx import PptxPackage, build_document_model, render_markdown, writer
from app.pptx.ooxml import NS, qn


def load(path) -> PptxPackage:
    return PptxPackage.load(path)


def roundtrip(pkg: PptxPackage, tmp_path) -> PptxPackage:
    out = tmp_path / "out.pptx"
    pkg.save(out)
    return PptxPackage.load(out)


def shape(doc, slide: int, name_or_id: str) -> dict:
    return next(s for s in doc["slides"][slide]["shapes"] if s["id"] == name_or_id)


# -- parsing -------------------------------------------------------------------


def test_slides_titles_and_size(sample_copy):
    doc = build_document_model(load(sample_copy))
    assert [s["title"] for s in doc["slides"]] == [
        "Q3 Project Status Review", "Key Achievements", "Schedule", "System Architecture",
    ]
    assert doc["slideSize"]["widthIn"] == 10.0 and doc["slideSize"]["heightIn"] == 7.5
    assert [s["index"] for s in doc["slides"]] == [1, 2, 3, 4]


def test_placeholder_geometry_and_text_style_are_inherited_from_layout_and_master(sample_copy):
    doc = build_document_model(load(sample_copy))
    body = shape(doc, 1, "s1.sh3")
    # The slide's body placeholder carries no xfrm of its own: position comes from the layout/master.
    assert (body["w"], body["h"]) > (0, 0) and body["x"] > 0
    first = body["paragraphs"][0]
    assert first["bullet"] == "•"  # bullet defined only in the master's bodyStyle
    assert first["runs"][0]["size"] == 32.0  # sz defined only in the master's bodyStyle
    title = shape(doc, 1, "s1.sh2")
    assert title["isTitle"] is True
    assert title["paragraphs"][0]["alignment"] == "center"


def test_table_and_picture_and_notes(sample_copy):
    doc = build_document_model(load(sample_copy))
    table = shape(doc, 2, "s2.sh3")
    assert table["type"] == "table" and len(table["rows"]) == 4 and len(table["cols"]) == 3
    assert table["rows"][1]["cells"][0]["paragraphs"][0]["runs"][0]["text"] == "Critical Design Review"
    pic = shape(doc, 3, "s3.sh3")
    assert pic["type"] == "picture" and pic["relId"]
    assert "FY23 baseline" in doc["slides"][1]["notes"]


def test_legacy_comments_threading_and_shape_mapping(sample_copy):
    comments = build_document_model(load(sample_copy))["comments"]
    assert set(comments) == {"0-1", "1-1", "0-2", "0-3"}
    top = comments["0-1"]
    assert top["author"] == "Alex Reviewer" and top["initials"] == "AR"
    assert top["slideId"] == "s1" and top["parentId"] is None
    # PowerPoint stores only a position; it is mapped to the shape containing it.
    assert top["shapeId"] == "s1.sh3"
    assert comments["1-1"]["parentId"] == "0-1" and comments["1-1"]["author"] == "Pat Author"
    assert comments["0-2"]["shapeId"] == "s2.sh3"
    assert comments["0-3"]["shapeId"] is None  # slide-level
    assert all(c["done"] is False and c["kind"] == "comment" for c in comments.values())


def test_markdown_context_marks_comment_anchors(sample_copy):
    md = render_markdown(build_document_model(load(sample_copy)))
    assert "## Slide 2: Key Achievements" in md
    assert "{>>comment:0-1<<}" in md
    assert "{>>comment:0-3 (about the slide as a whole)<<}" in md
    assert "| Milestone | Date | Status |" in md
    assert "Speaker notes" in md and "FY23 baseline" in md


# -- untouched round trip ----------------------------------------------------------


def test_save_without_changes_keeps_every_part_byte_identical(sample_copy, tmp_path):
    pkg = load(sample_copy)
    build_document_model(pkg)  # parsing must not mark anything dirty
    out = tmp_path / "out.pptx"
    pkg.save(out)
    with zipfile.ZipFile(sample_copy) as a, zipfile.ZipFile(out) as b:
        assert sorted(a.namelist()) == sorted(b.namelist())
        for name in a.namelist():
            assert a.read(name) == b.read(name), name


# -- comments ----------------------------------------------------------------------


def test_add_comment_for_selection_stores_quote_and_survives_roundtrip(sample_copy, tmp_path):
    pkg = load(sample_copy)
    cid = writer.add_comment_for_selection(
        pkg, "s1.sh3.p2.r0", 0, "s1.sh3.p2.r0", 24, "Jane Reviewer", "Which baseline?"
    )
    assert cid == "2-1"  # new author -> id 2, first idx
    again = roundtrip(pkg, tmp_path)
    c = build_document_model(again)["comments"][cid]
    assert c["quote"] == "Reduced unit cost by 12%"
    assert c["shapeId"] == "s1.sh3" and c["slideId"] == "s1" and c["author"] == "Jane Reviewer"


def test_comment_on_slide_without_comments_wires_up_part_rel_and_content_type(sample_copy, tmp_path):
    pkg = load(sample_copy)
    writer.add_comment(pkg, slide_id="s0", author="Jane Reviewer", text="Add the date")
    out = tmp_path / "out.pptx"
    pkg.save(out)
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
        assert "ppt/comments/comment4.xml" in names
        assert b"/ppt/comments/comment4.xml" in z.read("[Content_Types].xml")
        assert b"comment4.xml" in z.read("ppt/slides/_rels/slide1.xml.rels")
        authors = etree.fromstring(z.read("ppt/commentAuthors.xml"))
        jane = next(a for a in authors if a.get("name") == "Jane Reviewer")
        assert jane.get("lastIdx") == "1"
    assert build_document_model(PptxPackage.load(out))["comments"]["2-1"]["slideId"] == "s0"


def test_deck_with_no_comments_at_all_gets_authors_part_created(sample_copy, tmp_path):
    # Strip every comment relationship/part to simulate a clean deck.
    from tests.conftest import rewrite_zip

    clean = tmp_path / "clean.pptx"
    with zipfile.ZipFile(sample_copy) as z:
        names = z.namelist()
        ct = z.read("[Content_Types].xml")
        pres_rels = z.read("ppt/_rels/presentation.xml.rels")
        slide_rels = {n: z.read(n) for n in names if n.startswith("ppt/slides/_rels/")}
    import re

    ct = re.sub(rb'<Override PartName="/ppt/comment[^"]*"[^>]*/>', b"", ct)
    pres_rels = re.sub(rb'<Relationship Id="rIdAuthors"[^>]*/>', b"", pres_rels)
    slide_rels = {n: re.sub(rb'<Relationship Id="rIdCmt"[^>]*/>', b"", d) for n, d in slide_rels.items()}
    with zipfile.ZipFile(sample_copy) as zin, zipfile.ZipFile(clean, "w") as zout:
        for item in zin.infolist():
            if item.filename.startswith("ppt/comment"):
                continue
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = ct
            elif item.filename == "ppt/_rels/presentation.xml.rels":
                data = pres_rels
            elif item.filename in slide_rels:
                data = slide_rels[item.filename]
            zout.writestr(item, data)

    pkg = load(clean)
    assert build_document_model(pkg)["comments"] == {}
    cid = writer.add_comment(pkg, slide_id="s1", shape_id="s1.sh3", author="Jane Reviewer", text="Hi")
    again = roundtrip(pkg, tmp_path)
    comments = build_document_model(again)["comments"]
    assert list(comments) == [cid] and comments[cid]["shapeId"] == "s1.sh3"


def test_reply_is_threaded_with_parentcm_and_pointing_at_the_same_spot(sample_copy, tmp_path):
    pkg = load(sample_copy)
    rid = writer.add_comment_reply(pkg, "0-1", "Jane Reviewer", "Thanks!")
    out = tmp_path / "out.pptx"
    pkg.save(out)
    with zipfile.ZipFile(out) as z:
        xml = z.read("ppt/comments/comment1.xml")
    root = etree.fromstring(xml)
    cm = [c for c in root if c.get("authorId") == "2"][0]
    pc = cm.find(f".//{{{NS['p15']}}}parentCm")
    assert (pc.get("authorId"), pc.get("idx")) == ("0", "1")
    doc = build_document_model(PptxPackage.load(out))
    assert doc["comments"][rid]["parentId"] == "0-1"
    assert doc["comments"][rid]["shapeId"] == "s1.sh3"


def test_resolve_and_reopen_legacy_comment(sample_copy, tmp_path):
    pkg = load(sample_copy)
    writer.resolve_comment(pkg, "0-2", True)
    assert build_document_model(roundtrip(pkg, tmp_path))["comments"]["0-2"]["done"] is True
    pkg2 = load(tmp_path / "out.pptx")
    writer.resolve_comment(pkg2, "0-2", False)
    assert build_document_model(roundtrip(pkg2, tmp_path))["comments"]["0-2"]["done"] is False


def test_unknown_comment_and_shape_errors(sample_copy):
    pkg = load(sample_copy)
    with pytest.raises(writer.EditError):
        writer.resolve_comment(pkg, "9-9")
    with pytest.raises(writer.EditError):
        writer.add_comment_reply(pkg, "9-9", "x", "y")
    with pytest.raises(writer.EditError):
        writer.add_comment(pkg, slide_id="s1", shape_id="s1.sh999", author="x", text="y")
    with pytest.raises(writer.EditError):
        writer.add_comment(pkg, slide_id="s99", author="x", text="y")
    with pytest.raises(writer.EditError):
        writer.add_comment(pkg, slide_id="s1", shape_id="s2.sh3", author="x", text="y")


# -- text edits ----------------------------------------------------------------------


def _split_first_bullet_into_three_runs(pkg: PptxPackage) -> None:
    """Make 'Completed preliminary design review ahead of schedule' three
    differently-formatted runs: 'Completed ' / 'preliminary design' (bold) / ' review ahead of schedule'."""
    tree = pkg.slide_tree(1)
    p = [p for p in tree.iter(qn("a:p")) if "Completed preliminary" in "".join(p.itertext())][0]
    r = p.find(qn("a:r"))
    r.find(qn("a:t")).text = "Completed "
    r2 = etree.fromstring(etree.tostring(r))
    r2.find(qn("a:t")).text = "preliminary design"
    rpr = r2.find(qn("a:rPr"))
    if rpr is None:
        rpr = etree.Element(qn("a:rPr"))
        r2.insert(0, rpr)
    rpr.set("b", "1")
    r3 = etree.fromstring(etree.tostring(r))
    r3.find(qn("a:t")).text = " review ahead of schedule"
    r.addnext(r2)
    r2.addnext(r3)
    pkg.mark_dirty(pkg.slide_parts[1])


def test_replace_within_one_run_keeps_formatting_and_leaves_resolved_edit_record(sample_copy):
    pkg = load(sample_copy)
    rec = writer.replace_text_range(
        pkg, "s1.sh3.p1.r0", 12, "s1.sh3.p1.r0", 24, "HRS-0190 requirement", "Jane Reviewer"
    )
    doc = build_document_model(pkg)
    para = shape(doc, 1, "s1.sh3")["paragraphs"][1]
    assert para["runs"][0]["text"] == "Requirement HRS-0190 requirement verified in lab testing"
    note = doc["comments"][rec]
    assert note["kind"] == "edit" and note["done"] is True and note["author"] == "Jane Reviewer"
    assert "ACA-HRS-0190" in note["text"] and "HRS-0190 requirement" in note["text"]
    assert note["shapeId"] == "s1.sh3"


def test_replace_across_runs_first_run_keeps_its_formatting(sample_copy):
    pkg = load(sample_copy)
    _split_first_bullet_into_three_runs(pkg)
    doc = build_document_model(pkg)
    runs = shape(doc, 1, "s1.sh3")["paragraphs"][0]["runs"]
    assert [r["text"] for r in runs] == ["Completed ", "preliminary design", " review ahead of schedule"]
    assert runs[1].get("bold") is True

    # Select "design review" = from run1 offset 12 through run2 offset 7.
    writer.replace_text_range(pkg, runs[1]["id"], 12, runs[2]["id"], 7, "plan", "Jane Reviewer", record=False)
    runs = shape(build_document_model(pkg), 1, "s1.sh3")["paragraphs"][0]["runs"]
    assert [r["text"] for r in runs] == ["Completed ", "preliminary plan", " ahead of schedule"]
    assert runs[1].get("bold") is True  # the first affected run's formatting carries the new text


def test_replace_whole_trailing_run_removes_it(sample_copy):
    pkg = load(sample_copy)
    _split_first_bullet_into_three_runs(pkg)
    runs = shape(build_document_model(pkg), 1, "s1.sh3")["paragraphs"][0]["runs"]
    writer.replace_text_range(pkg, runs[1]["id"], 0, runs[2]["id"], len(runs[2]["text"]), "X", "J", record=False)
    runs = shape(build_document_model(pkg), 1, "s1.sh3")["paragraphs"][0]["runs"]
    assert [r["text"] for r in runs] == ["Completed ", "X"]


def test_reversed_selection_is_normalized(sample_copy):
    pkg = load(sample_copy)
    writer.replace_text_range(pkg, "s1.sh3.p3.r0", 19, "s1.sh3.p3.r0", 14, "FIRST", "J", record=False)
    text = shape(build_document_model(pkg), 1, "s1.sh3")["paragraphs"][3]["runs"][0]["text"]
    assert text == "Delivered the FIRST prototype to the customer"


def test_selection_across_paragraphs_is_rejected(sample_copy):
    pkg = load(sample_copy)
    with pytest.raises(writer.EditError, match="single paragraph"):
        writer.replace_text_range(pkg, "s1.sh3.p0.r0", 0, "s1.sh3.p1.r0", 3, "x", "J")


def test_replace_inside_a_table_cell(sample_copy):
    pkg = load(sample_copy)
    writer.replace_text_range(pkg, "s2.sh3.r3.c2.p0.r0", 0, "s2.sh3.r3.c2.p0.r0", 7, "On track", "J", record=False)
    cell = shape(build_document_model(pkg), 2, "s2.sh3")["rows"][3]["cells"][2]
    assert cell["paragraphs"][0]["runs"][0]["text"] == "On track"


def test_unchanged_text_creates_no_edit_record(sample_copy):
    pkg = load(sample_copy)
    assert writer.replace_text_range(pkg, "s1.sh3.p3.r0", 0, "s1.sh3.p3.r0", 9, "Delivered", "J") is None


def test_replace_commented_span_needs_a_quote_or_explicit_original(sample_copy):
    pkg = load(sample_copy)
    with pytest.raises(writer.EditError, match="specify which text"):
        writer.replace_commented_span(pkg, "0-1", "new", "AI Assistant")
    writer.replace_commented_span(pkg, "0-1", "TR-2025-014 verified", "AI Assistant", original_text="verified in lab testing")
    text = shape(build_document_model(pkg), 1, "s1.sh3")["paragraphs"][1]["runs"][0]["text"]
    assert text == "Requirement ACA-HRS-0190 TR-2025-014 verified"
    with pytest.raises(writer.EditError, match="not found"):
        writer.replace_commented_span(pkg, "0-1", "new", "AI Assistant", original_text="text that is not there")
    with pytest.raises(writer.EditError, match="slide, not to any text"):
        writer.replace_commented_span(pkg, "0-3", "new", "AI Assistant", original_text="x")


def test_replace_commented_span_updates_the_stored_quote(sample_copy):
    pkg = load(sample_copy)
    cid = writer.add_comment_for_selection(
        pkg, "s1.sh3.p2.r0", 0, "s1.sh3.p2.r0", 24, "Jane Reviewer", "Which baseline?"
    )
    writer.replace_commented_span(pkg, cid, "Cut unit cost by 12%", "AI Assistant")
    assert build_document_model(pkg)["comments"][cid]["quote"] == "Cut unit cost by 12%"


# -- modern (Microsoft 365) comments --------------------------------------------------

MODERN_AUTHORS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p188:authorLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main">'
    '<p188:author id="{11111111-1111-1111-1111-111111111111}" name="Morgan Modern" initials="MM" '
    'userId="morgan@example.com" providerId="AD"/></p188:authorLst>'
).encode()

MODERN_COMMENTS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p188:cmLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main">'
    '<p188:cm id="{22222222-2222-2222-2222-222222222222}" authorId="{11111111-1111-1111-1111-111111111111}" '
    'created="2025-03-01T10:00:00.000">'
    '<ac:deMkLst xmlns:ac="http://schemas.microsoft.com/office/drawing/2013/main/command" '
    'xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command">'
    '<pc:docMk/><pc:sldMk cId="1" sldId="256"/><ac:spMk id="2" creationId="{00000000-0000-0000-0000-000000000000}"/>'
    "</ac:deMkLst>"
    '<p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Modern comment on the title</a:t></a:r></a:p></p188:txBody>'
    "</p188:cm></p188:cmLst>"
).encode()


@pytest.fixture()
def modern_deck(sample_copy, tmp_path):
    from tests.conftest import rewrite_zip

    out = tmp_path / "modern.pptx"
    with zipfile.ZipFile(sample_copy) as z:
        ct = z.read("[Content_Types].xml").replace(
            b"</Types>",
            b'<Override PartName="/ppt/authors.xml" ContentType="application/vnd.ms-powerpoint.authors+xml"/>'
            b'<Override PartName="/ppt/comments/modernComment_100_1.xml" '
            b'ContentType="application/vnd.ms-powerpoint.comments+xml"/></Types>',
        )
        pres_rels = z.read("ppt/_rels/presentation.xml.rels").replace(
            b"</Relationships>",
            b'<Relationship Id="rIdAu" Type="http://schemas.microsoft.com/office/2018/10/relationships/authors" '
            b'Target="authors.xml"/></Relationships>',
        )
        slide_rels = z.read("ppt/slides/_rels/slide1.xml.rels").replace(
            b"</Relationships>",
            b'<Relationship Id="rIdMc" Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" '
            b'Target="../comments/modernComment_100_1.xml"/></Relationships>',
        )
    rewrite_zip(
        sample_copy, out,
        replace={
            "[Content_Types].xml": ct,
            "ppt/_rels/presentation.xml.rels": pres_rels,
            "ppt/slides/_rels/slide1.xml.rels": slide_rels,
        },
        add={"ppt/authors.xml": MODERN_AUTHORS, "ppt/comments/modernComment_100_1.xml": MODERN_COMMENTS},
    )
    return out


def test_modern_comments_are_read(modern_deck):
    comments = build_document_model(load(modern_deck))["comments"]
    cid = "{22222222-2222-2222-2222-222222222222}"
    assert cid in comments
    c = comments[cid]
    assert c["author"] == "Morgan Modern" and c["text"] == "Modern comment on the title"
    assert c["slideId"] == "s0" and c["shapeId"] == "s0.sh2" and c["done"] is False


def test_modern_reply_and_resolve_stay_modern(modern_deck, tmp_path):
    pkg = load(modern_deck)
    cid = "{22222222-2222-2222-2222-222222222222}"
    rid = writer.add_comment_reply(pkg, cid, "Jane Reviewer", "Agreed")
    writer.resolve_comment(pkg, cid, True)
    out = tmp_path / "out.pptx"
    pkg.save(out)
    with zipfile.ZipFile(out) as z:
        root = etree.fromstring(z.read("ppt/comments/modernComment_100_1.xml"))
        authors = etree.fromstring(z.read("ppt/authors.xml"))
    cm = root[0]
    assert cm.get("status") == "resolved"
    reply = cm.find(f".//{{{NS['p188']}}}reply")
    assert reply.get("id") == rid and "Agreed" in "".join(reply.itertext())
    assert reply.get("authorId") in [a.get("id") for a in authors if a.get("name") == "Jane Reviewer"]
    comments = build_document_model(PptxPackage.load(out))["comments"]
    assert comments[cid]["done"] is True and comments[rid]["parentId"] == cid


# -- inherited text color ----------------------------------------------------------


def test_table_text_color_is_not_inherited_from_master_defaults(sample_copy):
    """Table text color comes from the (unmodeled) table style in PowerPoint, so
    runs carry a color only when set explicitly -- the renderer then picks a
    readable color for the cell fill (white on the blue header row)."""
    table = shape(build_document_model(load(sample_copy)), 2, "s2.sh3")
    assert all("color" not in r for row in table["rows"] for c in row["cells"] for p in c["paragraphs"] for r in p["runs"])
    assert table["firstRow"] is True and table["headerFill"]


def test_font_ref_color_on_a_styled_shape_becomes_the_default_text_color(sample_copy):
    pkg = load(sample_copy)
    # Add a filled rectangle whose text color comes only from p:style/fontRef (white).
    xml = (
        f'<p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:nvSpPr><p:cNvPr id="90" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="2000000" cy="800000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></p:spPr>'
        '<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
        '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>'
    )
    pkg.slide_tree(0).find(f"{qn('p:cSld')}/{qn('p:spTree')}").append(etree.fromstring(xml))
    box = shape(build_document_model(pkg), 0, "s0.sh90")
    assert box["fill"] == "1F3864"
    assert box["paragraphs"][0]["runs"][0]["color"] == "FFFFFF"  # lt1 -> white, from fontRef


# -- layout decorations, groups, rotation ----------------------------------------------


def test_layout_decorations_are_returned_as_background_shapes_but_placeholders_are_not(sample_copy):
    pkg = load(sample_copy)
    layout_part = pkg.layout_part(pkg.slide_parts[1])
    xml = (
        f'<p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:nvSpPr><p:cNvPr id="77" name="Accent bar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="150000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></p:spPr></p:sp>'
    )
    pkg.tree(layout_part).find(f"{qn('p:cSld')}/{qn('p:spTree')}").append(etree.fromstring(xml))
    slide = build_document_model(pkg)["slides"][1]
    bars = [s for s in slide["backgroundShapes"] if s.get("fill") == "C00000"]
    assert len(bars) == 1 and bars[0]["w"] == 9144000 and bars[0]["id"].startswith("s1.bg")
    # Layout/master placeholders (the empty "click to add title" boxes) are templates, never drawn.
    assert not any(s.get("placeholder") for s in slide["backgroundShapes"])
    # And a slide that opts out of master shapes doesn't get them.
    pkg.slide_tree(1).set("showMasterSp", "0")
    assert build_document_model(pkg)["slides"][1]["backgroundShapes"] == []


def test_group_transform_maps_child_coordinates_to_slide_space_and_rotation_is_kept(sample_copy):
    pkg = load(sample_copy)
    xml = (
        f'<p:grpSp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:nvGrpSpPr><p:cNvPr id="80" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="2000000" cy="1000000"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="500000"/></a:xfrm></p:grpSpPr>'
        '<p:sp><p:nvSpPr><p:cNvPr id="81" name="Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm rot="5400000"><a:off x="100000" y="50000"/><a:ext cx="400000" cy="200000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>In a group</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp>'
    )
    pkg.slide_tree(0).find(f"{qn('p:cSld')}/{qn('p:spTree')}").append(etree.fromstring(xml))
    child = shape(build_document_model(pkg), 0, "s0.sh81")
    # Group is scaled 2x in both axes and offset by (1_000_000, 2_000_000).
    assert (child["x"], child["y"], child["w"], child["h"]) == (1200000, 2100000, 800000, 400000)
    assert child["rot"] == 90.0
    # Text in a group is addressable like any other shape's text.
    pkg2 = load(sample_copy)
    pkg2.slide_tree(0).find(f"{qn('p:cSld')}/{qn('p:spTree')}").append(etree.fromstring(xml))
    writer.replace_text_range(pkg2, "s0.sh81.p0.r0", 0, "s0.sh81.p0.r0", 10, "Edited", "J", record=False)
    assert shape(build_document_model(pkg2), 0, "s0.sh81")["paragraphs"][0]["runs"][0]["text"] == "Edited"
