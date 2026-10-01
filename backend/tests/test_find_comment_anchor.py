"""Unit tests for markdown_render.find_comment_anchor -- the reverse of
comment_locations' derivedRef/derivedPara guess, used when an imported
Comment Resolution Matrix row matches no existing comment (a clean master
deck paired with a separately-maintained review spreadsheet) and a brand-new
comment needs to be attached somewhere based on that row's Ref #/Para # alone.
"""
from app.pptx.markdown_render import find_comment_anchor


def _shape(sid: str, text: str, title: bool = False) -> dict:
    s = {
        "id": sid, "type": "shape",
        "paragraphs": [{"id": f"{sid}.p0", "runs": [{"id": f"{sid}.p0.r0", "type": "text", "text": text}]}],
    }
    if title:
        s["isTitle"] = True
    return s


DOC = {
    "slides": [
        {"id": "s0", "index": 1, "shapes": [_shape("s0.sh2", "Overview", title=True)]},
        {"id": "s1", "index": 2, "shapes": [
            _shape("s1.sh2", "Requirements", title=True),
            _shape("s1.sh3", "ACA-HRS-0190: The system shall operate at 115 VAC."),
            _shape("s1.sh4", "Unrelated text with no requirement id."),
        ]},
        {"id": "s2", "index": 3, "shapes": [_shape("s2.sh2", "Slide without a requirement")]},
    ]
}


def test_anchors_on_shape_containing_the_exact_requirement_id():
    assert find_comment_anchor(DOC, ref=None, para="ACA-HRS-0190") == {
        "slide_id": "s1", "shape_id": "s1.sh3", "quote": "ACA-HRS-0190",
    }


def test_requirement_id_match_is_case_and_whitespace_insensitive():
    anchor = find_comment_anchor(DOC, ref=None, para="  aca-hrs-0190 ")
    assert anchor and anchor["shape_id"] == "s1.sh3"


def test_slide_ref_anchors_on_that_slides_title():
    assert find_comment_anchor(DOC, ref="Slide 2", para=None) == {
        "slide_id": "s1", "shape_id": "s1.sh2", "quote": None,
    }
    assert find_comment_anchor(DOC, ref="2", para=None)["slide_id"] == "s1"
    assert find_comment_anchor(DOC, ref="slide #2", para=None)["slide_id"] == "s1"


def test_slide_ref_without_title_anchors_slide_level():
    doc = {"slides": [{"id": "s0", "index": 1, "shapes": [_shape("s0.sh2", "text")]}]}
    assert find_comment_anchor(doc, ref="1", para=None) == {"slide_id": "s0", "shape_id": None, "quote": None}


def test_para_is_preferred_over_ref_when_both_given():
    anchor = find_comment_anchor(DOC, ref="Slide 1", para="ACA-HRS-0190")
    assert anchor["slide_id"] == "s1" and anchor["shape_id"] == "s1.sh3"


def test_returns_none_when_nothing_matches():
    assert find_comment_anchor(DOC, ref=None, para=None) is None
    assert find_comment_anchor(DOC, ref="Slide 99", para=None) is None
    assert find_comment_anchor(DOC, ref="3.2.2", para=None) is None
    assert find_comment_anchor(DOC, ref=None, para="XYZ-9999") is None
