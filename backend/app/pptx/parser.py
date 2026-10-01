"""Slides (lxml trees) -> JSON-serializable document model.

The model is a read projection used for storage, markdown rendering, AI
context, and the frontend. IDs are *positional paths* recomputed from the
live XML every time we parse:

    slide      s2
    shape      s2.sh7            (7 = the shape's own p:cNvPr id, stable
                                  across edits, unique within the slide)
    paragraph  s2.sh7.p0         (table cell: s2.sh7.r1.c0.p0)
    run        s2.sh7.p0.r3

Edits always re-parse afterwards, so run/paragraph IDs are only ever used
within one request/response round trip, never persisted as if stable.

PowerPoint stores most visual properties by *inheritance* (slide placeholder
-> layout placeholder -> master placeholder -> master text styles ->
presentation defaults), so this module resolves that chain for position,
font size/weight/color, alignment and bullets. The browser then renders
absolutely-positioned shapes with no knowledge of OOXML at all.
"""
from __future__ import annotations

import colorsys
from typing import Any

from lxml import etree

from . import comments as comments_mod
from .ooxml import (
    EMU_PER_INCH,
    NS,
    PRESENTATION_PART,
    REL_NOTES,
    PptxPackage,
    local,
    qn,
)

A = NS["a"]
SHAPE_TAGS = ("sp", "pic", "graphicFrame", "grpSp", "cxnSp")

# Placeholder types whose text we treat as the slide's title.
TITLE_TYPES = ("title", "ctrTitle")


def shape_model_id(slide_idx: int, cnv_id: str | int) -> str:
    return f"s{slide_idx}.sh{cnv_id}"


# -- paragraphs / runs (shared with writer.py) -----------------------------


def iter_para_items(p_el):
    """The ordered run-level children of an a:p that the model exposes as
    runs: a:r (text), a:br (line break), a:fld (field). Shared by parser and
    writer so run indices always agree."""
    for child in p_el:
        if local(child) in ("r", "br", "fld"):
            yield child


def run_text(r_el) -> str:
    if local(r_el) == "br":
        return "\n"
    t = r_el.find(qn("a:t"))
    return (t.text or "") if t is not None else ""


# -- theme / color ---------------------------------------------------------

_DEFAULT_COLORS = {
    "dk1": "000000", "lt1": "FFFFFF", "dk2": "44546A", "lt2": "E7E6E6",
    "accent1": "4472C4", "accent2": "ED7D31", "accent3": "A5A5A5",
    "accent4": "FFC000", "accent5": "5B9BD5", "accent6": "70AD47",
    "hlink": "0563C1", "folHlink": "954F72",
}
_DEFAULT_CLRMAP = {
    "bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2",
    "accent1": "accent1", "accent2": "accent2", "accent3": "accent3",
    "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
    "hlink": "hlink", "folHlink": "folHlink",
}


class SlideCtx:
    """Per-slide resolution context: layout/master trees, theme colors."""

    def __init__(self, pkg: PptxPackage, slide_part: str):
        self.pkg = pkg
        self.slide_part = slide_part
        self.layout_part = pkg.layout_part(slide_part)
        self.master_part = pkg.master_part(self.layout_part)
        self.layout = pkg.tree(self.layout_part) if self.layout_part else None
        self.master = pkg.tree(self.master_part) if self.master_part else None
        self.colors = dict(_DEFAULT_COLORS)
        theme_part = pkg.theme_part(self.master_part)
        if theme_part:
            scheme = pkg.tree(theme_part).find(f".//{{{A}}}clrScheme")
            if scheme is not None:
                for c in scheme:
                    inner = next(iter(c), None)
                    if inner is None:
                        continue
                    val = inner.get("val") if local(inner) == "srgbClr" else inner.get("lastClr")
                    if val:
                        self.colors[local(c)] = val.upper()
        self.clrmap = dict(_DEFAULT_CLRMAP)
        if self.master is not None:
            cm = self.master.find(qn("p:clrMap"))
            if cm is not None:
                self.clrmap.update({k: v for k, v in cm.attrib.items()})
        pres = pkg.tree(PRESENTATION_PART)
        self.default_text_style = pres.find(qn("p:defaultTextStyle"))
        self.master_tx_styles = self.master.find(qn("p:txStyles")) if self.master is not None else None

    def color_from(self, clr_el) -> str | None:
        """a:srgbClr / a:schemeClr / a:sysClr / a:prstClr element -> hex."""
        if clr_el is None:
            return None
        kind = local(clr_el)
        base: str | None = None
        if kind == "srgbClr":
            base = (clr_el.get("val") or "").upper() or None
        elif kind == "sysClr":
            base = (clr_el.get("lastClr") or "").upper() or None
        elif kind == "schemeClr":
            val = clr_el.get("val") or ""
            if val == "phClr":
                return None
            base = self.colors.get(self.clrmap.get(val, val))
        elif kind == "prstClr":
            base = {"black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000",
                    "blue": "0000FF", "yellow": "FFFF00"}.get(clr_el.get("val") or "")
        if not base or len(base) != 6:
            return None
        r, g, b = (int(base[i:i + 2], 16) / 255 for i in (0, 2, 4))
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        for mod in clr_el:
            try:
                v = int(mod.get("val", "0")) / 100000
            except ValueError:
                continue
            m = local(mod)
            if m == "lumMod":
                l = min(1.0, l * v)
            elif m == "lumOff":
                l = min(1.0, l + v)
            elif m == "tint":
                r, g, b = colorsys.hls_to_rgb(h, l, s)
                r, g, b = (c * v + (1 - v) for c in (r, g, b))
                h, l, s = colorsys.rgb_to_hls(r, g, b)
            elif m == "shade":
                l = l * v
        r, g, b = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)
        return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))

    def fill_color(self, holder) -> str | None:
        """Solid fill color of an spPr/rPr/tcPr-like holder, or None."""
        if holder is None:
            return None
        sf = holder.find(qn("a:solidFill"))
        if sf is None:
            return None
        return self.color_from(next(iter(sf), None))


def _ph_of(sp):
    nv = sp.find(qn("p:nvSpPr"))
    if nv is None:
        nv = sp.find(qn("p:nvPicPr"))
    if nv is None:
        return None
    return nv.find(f"{qn('p:nvPr')}/{qn('p:ph')}")


def _norm_ph_type(ph) -> str:
    t = ph.get("type") or "obj"
    if t in TITLE_TYPES:
        return "title"
    if t in ("body", "subTitle", "obj"):
        return "body"
    return t


def _find_matching_ph(tree, ph):
    """The sp in `tree` (a layout or master) whose placeholder `ph` (from
    the level below) inherits from: same idx if both have one, else same
    normalized type."""
    if tree is None or ph is None:
        return None
    candidates = [sp for sp in tree.iter(qn("p:sp")) if _ph_of(sp) is not None]
    idx = ph.get("idx")
    if idx is not None:
        for sp in candidates:
            if _ph_of(sp).get("idx") == idx and _norm_ph_type(_ph_of(sp)) == _norm_ph_type(ph):
                return sp
    want = _norm_ph_type(ph)
    for sp in candidates:
        if _norm_ph_type(_ph_of(sp)) == want and (idx is None or _ph_of(sp).get("idx") in (None, idx, "1")):
            return sp
    for sp in candidates:
        if _norm_ph_type(_ph_of(sp)) == want:
            return sp
    return None


def _xfrm_of(sp):
    spPr = sp.find(qn("p:spPr"))
    if spPr is not None:
        x = spPr.find(qn("a:xfrm"))
        if x is not None and x.find(qn("a:off")) is not None:
            return x
    return None


# -- text style resolution --------------------------------------------------


def _collect_ppr(out: dict[str, Any], ppr, ctx: SlideCtx) -> None:
    """Merge one pPr-shaped element (a paragraph's own pPr, or an
    a:lvlNpPr in some list style) into `out`, never overriding a value an
    earlier (higher-priority) element already set."""
    if ppr is None:
        return
    for k in ("algn", "marL", "indent"):
        if k not in out and ppr.get(k) is not None:
            out[k] = ppr.get(k)
    if "bullet" not in out:
        if ppr.find(qn("a:buNone")) is not None:
            out["bullet"] = None
        elif ppr.find(qn("a:buChar")) is not None:
            out["bullet"] = ppr.find(qn("a:buChar")).get("char") or "•"
        elif ppr.find(qn("a:buAutoNum")) is not None:
            out["bullet"] = "#"
    if "lineSpacing" not in out:
        pct = ppr.find(f"{qn('a:lnSpc')}/{qn('a:spcPct')}")
        if pct is not None and (pct.get("val") or "").isdigit():
            out["lineSpacing"] = int(pct.get("val")) / 100000
    if "spaceBefore" not in out:
        pts = ppr.find(f"{qn('a:spcBef')}/{qn('a:spcPts')}")
        if pts is not None and (pts.get("val") or "").isdigit():
            out["spaceBefore"] = int(pts.get("val")) / 100
    d = ppr.find(qn("a:defRPr"))
    if d is not None:
        _collect_rpr(out, d, ctx, prefix="")


def _collect_rpr(out: dict[str, Any], rpr, ctx: SlideCtx, prefix: str = "") -> None:
    if rpr is None:
        return
    for attr, key in (("sz", "sz"), ("b", "b"), ("i", "i"), ("u", "u")):
        k = prefix + key
        if k not in out and rpr.get(attr) is not None:
            out[k] = rpr.get(attr)
    if prefix + "color" not in out:
        c = ctx.fill_color(rpr)
        if c:
            out[prefix + "color"] = c


def _style_chain(ctx: SlideCtx, sp, ph) -> list:
    """Containers of a:lvlNpPr elements, highest priority first."""
    chain = []
    own = sp.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}") if sp is not None else None
    if own is not None:
        chain.append(own)
    if ph is not None:
        lay_sp = _find_matching_ph(ctx.layout, ph)
        if lay_sp is not None:
            ls = lay_sp.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}")
            if ls is not None:
                chain.append(ls)
            lph = _ph_of(lay_sp)
            mas_sp = _find_matching_ph(ctx.master, lph if lph is not None else ph)
            if mas_sp is not None:
                ms = mas_sp.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}")
                if ms is not None:
                    chain.append(ms)
        t = _norm_ph_type(ph)
        if ctx.master_tx_styles is not None:
            if t == "title":
                node = ctx.master_tx_styles.find(qn("p:titleStyle"))
            elif t in ("dt", "ftr", "sldNum"):
                node = ctx.master_tx_styles.find(qn("p:otherStyle"))
            else:
                node = ctx.master_tx_styles.find(qn("p:bodyStyle"))
            if node is not None:
                chain.append(node)
    elif ctx.master_tx_styles is not None:
        node = ctx.master_tx_styles.find(qn("p:otherStyle"))
        if node is not None:
            chain.append(node)
    if ctx.default_text_style is not None:
        chain.append(ctx.default_text_style)
    return chain


def _font_ref_container(ctx: SlideCtx, sp):
    """A shape's p:style/a:fontRef carries the text color for text in that
    shape (e.g. white on a filled rectangle). Wrap it as a pseudo list-style
    so it slots into the inheritance chain right after the shape's own
    lstStyle and ahead of the master's defaults."""
    clr = sp.find(f"{qn('p:style')}/{qn('a:fontRef')}")
    if clr is None or len(clr) == 0:
        return None
    holder = etree.Element(qn("a:lstStyle"))
    for lvl in range(1, 10):
        lv = etree.SubElement(holder, qn(f"a:lvl{lvl}pPr"))
        d = etree.SubElement(lv, qn("a:defRPr"))
        sf = etree.SubElement(d, qn("a:solidFill"))
        sf.append(etree.fromstring(etree.tostring(clr[0])))
    return holder


def _resolve_bodypr(ctx: SlideCtx, sp, ph) -> dict[str, Any]:
    out: dict[str, Any] = {}
    nodes = []
    if sp is not None:
        nodes.append(sp.find(f"{qn('p:txBody')}/{qn('a:bodyPr')}"))
    if ph is not None:
        lay_sp = _find_matching_ph(ctx.layout, ph)
        if lay_sp is not None:
            nodes.append(lay_sp.find(f"{qn('p:txBody')}/{qn('a:bodyPr')}"))
            mas_sp = _find_matching_ph(ctx.master, _ph_of(lay_sp) if _ph_of(lay_sp) is not None else ph)
            if mas_sp is not None:
                nodes.append(mas_sp.find(f"{qn('p:txBody')}/{qn('a:bodyPr')}"))
    for node in nodes:
        if node is None:
            continue
        for k in ("anchor", "lIns", "tIns", "rIns", "bIns", "wrap"):
            if k not in out and node.get(k) is not None:
                out[k] = node.get(k)
        if "fontScale" not in out:
            na = node.find(qn("a:normAutofit"))
            if na is not None and (na.get("fontScale") or "").isdigit():
                out["fontScale"] = int(na.get("fontScale")) / 100000
    return out


def _parse_paragraph(
    p_el, pid: str, ctx: SlideCtx, chain: list, font_scale: float, inherit_color: bool = True
) -> dict[str, Any]:
    own_ppr = p_el.find(qn("a:pPr"))
    lvl = int(own_ppr.get("lvl", "0")) if own_ppr is not None and (own_ppr.get("lvl") or "0").isdigit() else 0
    lvl = max(0, min(lvl, 8))

    base: dict[str, Any] = {}
    _collect_ppr(base, own_ppr, ctx)
    for container in chain:
        node = container.find(qn(f"a:lvl{lvl + 1}pPr"))
        _collect_ppr(base, node, ctx)
    if not inherit_color:
        # Table text color comes from the table *style* in PowerPoint (white
        # header text on a colored header row, ...), which isn't modeled, so
        # only a color set on the run itself counts; the renderer picks a
        # readable default for the cell's fill.
        base.pop("color", None)

    runs: list[dict[str, Any]] = []
    for ri, r_el in enumerate(iter_para_items(p_el)):
        rid = f"{pid}.r{ri}"
        if local(r_el) == "br":
            runs.append({"id": rid, "type": "break", "text": "\n"})
            continue
        rpr = r_el.find(qn("a:rPr"))
        rp: dict[str, Any] = {}
        _collect_rpr(rp, rpr, ctx)
        merged = {**base, **rp}  # run-level attributes win over inherited defaults
        node: dict[str, Any] = {"id": rid, "type": "text", "text": run_text(r_el)}
        sz = merged.get("sz")
        if sz and str(sz).isdigit():
            node["size"] = round(int(sz) / 100 * font_scale, 2)
        if merged.get("b") in ("1", "true"):
            node["bold"] = True
        if merged.get("i") in ("1", "true"):
            node["italic"] = True
        if merged.get("u") not in (None, "none"):
            node["underline"] = True
        color = merged.get("color")
        if color:
            node["color"] = color
        if local(r_el) == "fld":
            node["field"] = True
        runs.append(node)

    para: dict[str, Any] = {"id": pid, "type": "paragraph", "runs": runs, "level": lvl}
    alg = {"ctr": "center", "r": "right", "just": "justify", "justLow": "justify", "dist": "justify"}.get(
        base.get("algn", "")
    )
    if alg:
        para["alignment"] = alg
    bullet = base.get("bullet")
    if bullet and any(r.get("text", "").strip() for r in runs):
        para["bullet"] = bullet
    for k in ("marL", "indent"):
        if str(base.get(k, "")).lstrip("-").isdigit():
            para[k] = int(base[k])
    if "lineSpacing" in base:
        para["lineSpacing"] = base["lineSpacing"]
    if "spaceBefore" in base:
        para["spaceBefore"] = base["spaceBefore"]
    # Default size for an empty paragraph (so its line height is right).
    sz = base.get("sz")
    if sz and str(sz).isdigit():
        para["defaultSize"] = round(int(sz) / 100 * font_scale, 2)
    return para


def _parse_txbody(
    tx, id_prefix: str, ctx: SlideCtx, chain: list, font_scale: float, inherit_color: bool = True
) -> list[dict[str, Any]]:
    if tx is None:
        return []
    return [
        _parse_paragraph(p, f"{id_prefix}.p{pi}", ctx, chain, font_scale, inherit_color)
        for pi, p in enumerate(tx.findall(qn("a:p")))
    ]


# -- shapes ----------------------------------------------------------------


class Matrix:
    """Group-shape coordinate transform (child space -> slide space)."""

    def __init__(self, sx=1.0, sy=1.0, tx=0.0, ty=0.0):
        self.sx, self.sy, self.tx, self.ty = sx, sy, tx, ty

    def apply(self, x: int, y: int, w: int, h: int) -> tuple[int, int, int, int]:
        return (
            round(self.tx + self.sx * x), round(self.ty + self.sy * y),
            round(self.sx * w), round(self.sy * h),
        )

    def child(self, grp_xfrm) -> "Matrix":
        off, ext = grp_xfrm.find(qn("a:off")), grp_xfrm.find(qn("a:ext"))
        choff, chext = grp_xfrm.find(qn("a:chOff")), grp_xfrm.find(qn("a:chExt"))
        if off is None or ext is None or choff is None or chext is None:
            return self
        ox, oy = int(off.get("x", 0)), int(off.get("y", 0))
        ex, ey = int(ext.get("cx", 0)), int(ext.get("cy", 0))
        cox, coy = int(choff.get("x", 0)), int(choff.get("y", 0))
        cex, cey = int(chext.get("cx", 0)) or 1, int(chext.get("cy", 0)) or 1
        fx, fy = ex / cex, ey / cey
        return Matrix(
            self.sx * fx, self.sy * fy,
            self.tx + self.sx * (ox - cox * fx), self.ty + self.sy * (oy - coy * fy),
        )


def _geometry(xfrm, matrix: Matrix) -> tuple[int, int, int, int] | None:
    if xfrm is None:
        return None
    off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        return None
    return matrix.apply(
        int(off.get("x", 0)), int(off.get("y", 0)), int(ext.get("cx", 0)), int(ext.get("cy", 0))
    )


def _line_of(ctx: SlideCtx, spPr) -> dict[str, Any] | None:
    if spPr is None:
        return None
    ln = spPr.find(qn("a:ln"))
    if ln is None or ln.find(qn("a:noFill")) is not None:
        return None
    color = ctx.fill_color(ln)
    if not color:
        return None
    w = int(ln.get("w", "12700")) if (ln.get("w") or "12700").isdigit() else 12700
    return {"color": color, "w": w}


def _cnv_pr(el):
    for nv in ("p:nvSpPr", "p:nvPicPr", "p:nvGraphicFramePr", "p:nvGrpSpPr", "p:nvCxnSpPr"):
        n = el.find(qn(nv))
        if n is not None:
            return n.find(qn("p:cNvPr"))
    return None


def _parse_table(gf, sid: str, ctx: SlideCtx, geom) -> dict[str, Any] | None:
    tbl = gf.find(f".//{qn('a:tbl')}")
    if tbl is None:
        return None
    grid = [int(g.get("w", "0")) for g in tbl.findall(f"{qn('a:tblGrid')}/{qn('a:gridCol')}")]
    tbl_pr = tbl.find(qn("a:tblPr"))
    chain = _style_chain(ctx, None, None)
    rows = []
    for ri, tr in enumerate(tbl.findall(qn("a:tr"))):
        cells = []
        for ci, tc in enumerate(tr.findall(qn("a:tc"))):
            tc_pr = tc.find(qn("a:tcPr"))
            cell = {
                "paragraphs": _parse_txbody(
                    tc.find(qn("a:txBody")), f"{sid}.r{ri}.c{ci}", ctx, chain, 1.0, inherit_color=False
                ),
            }
            if tc.get("gridSpan"):
                cell["colSpan"] = int(tc.get("gridSpan"))
            if tc.get("rowSpan"):
                cell["rowSpan"] = int(tc.get("rowSpan"))
            if tc.get("hMerge") == "1":
                cell["hMerge"] = True
            if tc.get("vMerge") == "1":
                cell["vMerge"] = True
            fill = ctx.fill_color(tc_pr)
            if fill:
                cell["fill"] = fill
            cells.append(cell)
        rows.append({"h": int(tr.get("h", "0")), "cells": cells})
    x, y, w, h = geom or (0, 0, sum(grid), sum(r["h"] for r in rows))
    return {
        "type": "table", "x": x, "y": y, "w": w, "h": h, "cols": grid, "rows": rows,
        "firstRow": tbl_pr is not None and tbl_pr.get("firstRow") == "1",
        "headerFill": ctx.colors.get("accent1"),
    }


def _walk_tree(
    container, slide_idx: int, ctx: SlideCtx, matrix: Matrix, out: list[dict[str, Any]], *,
    background: bool = False,
) -> None:
    for el in container:
        tag = local(el)
        if tag not in SHAPE_TAGS:
            continue
        cnv = _cnv_pr(el)
        if cnv is None:
            continue
        cid = cnv.get("id", "0")
        sid = shape_model_id(slide_idx, cid)

        if tag == "grpSp":
            xf = el.find(f"{qn('p:grpSpPr')}/{qn('a:xfrm')}")
            _walk_tree(el, slide_idx, ctx, matrix.child(xf) if xf is not None else matrix, out,
                       background=background)
            continue

        ph = _ph_of(el)
        if background and ph is not None:
            continue  # layout/master placeholders are templates, not content

        geom = None
        if tag in ("sp", "pic", "cxnSp"):
            xf = _xfrm_of(el)
            if xf is None and ph is not None:
                lay_sp = _find_matching_ph(ctx.layout, ph)
                xf = _xfrm_of(lay_sp) if lay_sp is not None else None
                if xf is None:
                    mas_sp = _find_matching_ph(ctx.master, _ph_of(lay_sp) if lay_sp is not None and _ph_of(lay_sp) is not None else ph)
                    xf = _xfrm_of(mas_sp) if mas_sp is not None else None
                # Inherited geometry is already in slide space (it's from a
                # layout/master, not the group the slide shape sits in).
                geom = _geometry(xf, Matrix()) if xf is not None else None
            else:
                geom = _geometry(xf, matrix)
        elif tag == "graphicFrame":
            geom = _geometry(el.find(qn("p:xfrm")), matrix)

        base: dict[str, Any] = {
            "id": sid, "name": cnv.get("name") or "", "descr": cnv.get("descr") or "",
        }
        if background:
            base["background"] = True

        if tag == "graphicFrame":
            if background:
                continue
            table = _parse_table(el, sid, ctx, geom)
            if table is not None:
                out.append({**base, **table})
            elif geom:
                out.append({**base, "type": "other", "x": geom[0], "y": geom[1], "w": geom[2], "h": geom[3],
                            "label": "Chart / embedded object (not previewed)"})
            continue

        if geom is None:
            continue
        x, y, w, h = geom
        base.update({"x": x, "y": y, "w": w, "h": h})
        xf_el = _xfrm_of(el)
        if xf_el is not None and (xf_el.get("rot") or "0").lstrip("-").isdigit() and int(xf_el.get("rot", "0")):
            base["rot"] = int(xf_el.get("rot")) / 60000

        if tag == "pic":
            if background:
                continue
            blip = el.find(f".//{qn('a:blip')}")
            base["type"] = "picture"
            base["relId"] = blip.get(qn("r:embed")) if blip is not None else None
            out.append(base)
            continue

        spPr = el.find(qn("p:spPr"))
        if tag == "cxnSp":
            line = _line_of(ctx, spPr)
            out.append({**base, "type": "line", "line": line or {"color": "888888", "w": 12700}})
            continue

        # p:sp
        base["type"] = "shape"
        if ph is not None:
            base["placeholder"] = ph.get("type") or "obj"
        fill = ctx.fill_color(spPr)
        if fill:
            base["fill"] = fill
        line = _line_of(ctx, spPr)
        if line:
            base["line"] = line
        geom_el = spPr.find(qn("a:prstGeom")) if spPr is not None else None
        if geom_el is not None and geom_el.get("prst"):
            base["geom"] = geom_el.get("prst")

        tx = el.find(qn("p:txBody"))
        if tx is not None and not background:
            body = _resolve_bodypr(ctx, el, ph)
            fs = body.get("fontScale", 1.0)
            chain = _style_chain(ctx, el, ph)
            if ph is None:
                fr = _font_ref_container(ctx, el)
                if fr is not None:
                    chain.insert(1 if el.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}") is not None else 0, fr)
            base["paragraphs"] = _parse_txbody(tx, sid, ctx, chain, fs)
            base["anchor"] = {"ctr": "middle", "b": "bottom"}.get(body.get("anchor", "t"), "top")
            base["insets"] = {
                "l": int(body.get("lIns", 91440)), "t": int(body.get("tIns", 45720)),
                "r": int(body.get("rIns", 91440)), "b": int(body.get("bIns", 45720)),
            }
            if body.get("wrap") == "none":
                base["nowrap"] = True
        else:
            base["paragraphs"] = []
        out.append(base)


def _slide_background(ctx: SlideCtx, slide_tree) -> str:
    for tree in (slide_tree, ctx.layout, ctx.master):
        if tree is None:
            continue
        bg = tree.find(f"{qn('p:cSld')}/{qn('p:bg')}")
        if bg is None:
            continue
        pr = bg.find(qn("p:bgPr"))
        c = ctx.fill_color(pr) if pr is not None else None
        if c:
            return c
        ref = bg.find(qn("p:bgRef"))
        if ref is not None:
            c = ctx.color_from(next(iter(ref), None))
            if c:
                return c
    return ctx.colors.get(ctx.clrmap.get("bg1", "lt1"), "FFFFFF")


def _parse_notes(pkg: PptxPackage, slide_part: str) -> str:
    notes_part = pkg.related_part(slide_part, REL_NOTES)
    if not notes_part:
        return ""
    tree = pkg.tree(notes_part)
    paras: list[str] = []
    for sp in tree.iter(qn("p:sp")):
        ph = _ph_of(sp)
        if ph is not None and ph.get("type") == "body":
            for p in sp.iter(qn("a:p")):
                paras.append("".join(run_text(r) for r in iter_para_items(p)))
    return "\n".join(paras).strip()


def parse_slide(pkg: PptxPackage, idx: int) -> dict[str, Any]:
    slide_part = pkg.slide_parts[idx]
    tree = pkg.tree(slide_part)
    ctx = SlideCtx(pkg, slide_part)

    shapes: list[dict[str, Any]] = []
    background_shapes: list[dict[str, Any]] = []
    sp_tree = tree.find(f"{qn('p:cSld')}/{qn('p:spTree')}")
    if sp_tree is not None:
        _walk_tree(sp_tree, idx, ctx, Matrix(), shapes)
    if tree.get("showMasterSp") != "0":
        if ctx.layout is not None:
            lt = ctx.layout.find(f"{qn('p:cSld')}/{qn('p:spTree')}")
            if lt is not None and ctx.layout.get("showMasterSp") != "0":
                master_tree = ctx.master.find(f"{qn('p:cSld')}/{qn('p:spTree')}") if ctx.master is not None else None
                if master_tree is not None:
                    _walk_tree(master_tree, idx, ctx, Matrix(), background_shapes, background=True)
            if lt is not None:
                _walk_tree(lt, idx, ctx, Matrix(), background_shapes, background=True)
    # Background shape ids come from layout/master and can collide with slide
    # shape ids; they're display-only, so namespace them.
    for i, s in enumerate(background_shapes):
        s["id"] = f"s{idx}.bg{i}"

    title = None
    for s in shapes:
        if s.get("placeholder") in TITLE_TYPES:
            title = " ".join(
                "".join(r.get("text", "") for r in p["runs"]) for p in s.get("paragraphs", [])
            ).strip() or None
            if title:
                s["isTitle"] = True
                break

    layout_name = None
    if ctx.layout is not None:
        c = ctx.layout.find(qn("p:cSld"))
        layout_name = c.get("name") if c is not None else None

    return {
        "id": f"s{idx}",
        "index": idx + 1,
        "title": title,
        "layout": layout_name,
        "hidden": tree.get("show") == "0",
        "background": _slide_background(ctx, tree),
        "shapes": shapes,
        "backgroundShapes": background_shapes,
        "notes": _parse_notes(pkg, slide_part),
    }


# -- comments ---------------------------------------------------------------

# Offset (EMU) a new comment pin sits inside its shape's top-left corner.
PIN_INSET = 182880


def _hit_test(shapes: list[dict[str, Any]], x: int, y: int) -> dict[str, Any] | None:
    best = None
    for s in shapes:
        if s["type"] not in ("shape", "table", "picture", "other"):
            continue
        if s["x"] <= x <= s["x"] + s["w"] and s["y"] <= y <= s["y"] + s["h"]:
            if best is None or s["w"] * s["h"] < best["w"] * best["h"]:
                best = s
    return best


def build_comments(pkg: PptxPackage, slides: list[dict[str, Any]]) -> dict[str, Any]:
    recs = comments_mod.read_comments(pkg)
    by_id = {r.id: r for r in recs}
    out: dict[str, Any] = {}
    for r in recs:
        slide = slides[r.slide_idx]
        shape = None
        shape_id = r.shape_id
        if shape_id is None and r.parent_id and by_id.get(r.parent_id):
            shape_id = by_id[r.parent_id].shape_id
        if shape_id is not None:
            mid = shape_model_id(r.slide_idx, shape_id)
            shape = next((s for s in slide["shapes"] if s["id"] == mid), None)
        if shape is None and r.pos_emu is not None:
            shape = _hit_test(slide["shapes"], *r.pos_emu)
        if r.pos_emu is not None and r.shape_id is None:
            pos = {"x": r.pos_emu[0], "y": r.pos_emu[1]}
        elif shape is not None:
            pos = {"x": shape["x"] + PIN_INSET, "y": shape["y"] + PIN_INSET}
        elif r.pos_emu is not None:
            pos = {"x": r.pos_emu[0], "y": r.pos_emu[1]}
        else:
            pos = {"x": PIN_INSET, "y": PIN_INSET}
        out[r.id] = {
            "id": r.id,
            "author": r.author or "Unknown",
            "initials": r.initials,
            "date": r.date,
            "text": r.text,
            "done": r.done,
            "parentId": r.parent_id,
            "slideId": slide["id"],
            "shapeId": shape["id"] if shape else None,
            "quote": r.quote,
            "kind": r.kind,
            "pos": pos,
        }
    return out


def build_document_model(pkg: PptxPackage) -> dict[str, Any]:
    slides = [parse_slide(pkg, i) for i in range(len(pkg.slide_parts))]
    w, h = pkg.slide_size
    return {
        "slides": slides,
        "comments": build_comments(pkg, slides),
        "slideSize": {
            "widthEmu": w, "heightEmu": h,
            "widthIn": round(w / EMU_PER_INCH, 3), "heightIn": round(h / EMU_PER_INCH, 3),
        },
    }
