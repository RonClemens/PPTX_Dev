from app.pptx.anchor import locate_quote_in_paragraph, locate_quote_in_shape, locate_quote_in_slide

PARA = {
    "id": "s1.sh3.p0",
    "type": "paragraph",
    "runs": [
        {"id": "s1.sh3.p0.r0", "type": "text", "text": "Hello "},
        {"id": "s1.sh3.p0.r1", "type": "text", "text": "brave"},
        {"id": "s1.sh3.p0.r2", "type": "text", "text": " world."},
    ],
}


def test_locate_quote_within_single_run():
    assert locate_quote_in_paragraph(PARA, "brave") == ("s1.sh3.p0.r1", 0, "s1.sh3.p0.r1", 5)


def test_locate_quote_spanning_multiple_runs():
    assert locate_quote_in_paragraph(PARA, "lo brave wo") == ("s1.sh3.p0.r0", 3, "s1.sh3.p0.r2", 3)


def test_locate_quote_at_paragraph_start_and_end():
    assert locate_quote_in_paragraph(PARA, "Hello") == ("s1.sh3.p0.r0", 0, "s1.sh3.p0.r0", 5)
    assert locate_quote_in_paragraph(PARA, "world.") == ("s1.sh3.p0.r2", 1, "s1.sh3.p0.r2", 7)


def test_locate_whole_paragraph():
    assert locate_quote_in_paragraph(PARA, "Hello brave world.") == ("s1.sh3.p0.r0", 0, "s1.sh3.p0.r2", 7)


def test_not_found_and_blank():
    assert locate_quote_in_paragraph(PARA, "nope") is None
    assert locate_quote_in_paragraph(PARA, "   ") is None


def test_line_breaks_are_not_part_of_the_anchor_space():
    para = {
        "id": "p",
        "runs": [
            {"id": "p.r0", "type": "text", "text": "one"},
            {"id": "p.r1", "type": "break", "text": "\n"},
            {"id": "p.r2", "type": "text", "text": "two"},
        ],
    }
    assert locate_quote_in_paragraph(para, "two") == ("p.r2", 0, "p.r2", 3)


def test_locate_in_shape_searches_every_paragraph_and_table_cells():
    shape = {"id": "s0.sh2", "type": "shape", "paragraphs": [
        {"id": "a", "runs": [{"id": "a.r0", "type": "text", "text": "first"}]},
        {"id": "b", "runs": [{"id": "b.r0", "type": "text", "text": "second para"}]},
    ]}
    assert locate_quote_in_shape(shape, "second") == ("b.r0", 0, "b.r0", 6)

    table = {"id": "s0.sh3", "type": "table", "rows": [
        {"cells": [{"paragraphs": [{"id": "c0", "runs": [{"id": "c0.r0", "type": "text", "text": "Qtr"}]}]},
                   {"paragraphs": [{"id": "c1", "runs": [{"id": "c1.r0", "type": "text", "text": "Revenue"}]}]}]},
    ]}
    assert locate_quote_in_shape(table, "Revenue") == ("c1.r0", 0, "c1.r0", 7)


def test_locate_in_slide_respects_candidate_shape_ids():
    slide = {"shapes": [
        {"id": "x", "type": "shape", "paragraphs": [{"id": "x.p0", "runs": [{"id": "x.p0.r0", "type": "text", "text": "alpha"}]}]},
        {"id": "y", "type": "shape", "paragraphs": [{"id": "y.p0", "runs": [{"id": "y.p0.r0", "type": "text", "text": "alpha beta"}]}]},
    ]}
    assert locate_quote_in_slide(slide, ["x", "y"], "alpha") == ("x.p0.r0", 0, "x.p0.r0", 5)
    assert locate_quote_in_slide(slide, ["y"], "beta") == ("y.p0.r0", 6, "y.p0.r0", 10)
    assert locate_quote_in_slide(slide, ["x"], "beta") is None
