"""Low-level OOXML helpers shared by the builder, the lint and the template inspector."""

import colorsys
import hashlib
import io
import json
import posixpath
import re
import uuid
import zipfile

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

import themes

SLOTS = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
         "accent4", "accent5", "accent6", "hlink", "folHlink")
DEFAULT_CLRMAP = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2",
                  **{f"accent{i}": f"accent{i}" for i in range(1, 7)},
                  "hlink": "hlink", "folHlink": "folHlink"}
TEMPLATE_CT = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
PRESENTATION_CT = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
NO_STYLE_TABLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"


def open_presentation(path):
    # python-pptx refuses the template content type; swapping it is all
    # PowerPoint itself does when it starts a deck from a .potx.
    with zipfile.ZipFile(path) as z:
        types = z.read("[Content_Types].xml").decode("utf-8")
        if TEMPLATE_CT not in types:
            return Presentation(path)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                data = z.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    data = types.replace(TEMPLATE_CT, PRESENTATION_CT).encode("utf-8")
                out.writestr(item, data)
    buf.seek(0)
    return Presentation(buf)


def remove_all_slides(prs):
    sld_ids = prs.slides._sldIdLst
    for sld_id in list(sld_ids):
        prs.part.drop_rel(sld_id.rId)
        sld_ids.remove(sld_id)


def theme_part(prs):
    return prs.slide_master.part.part_related_by(RT.THEME)


def _slot_hex(elem):
    srgb = elem.find(qn("a:srgbClr"))
    if srgb is not None:
        return srgb.get("val", "000000").upper()
    sys_clr = elem.find(qn("a:sysClr"))
    if sys_clr is not None:
        default = "FFFFFF" if sys_clr.get("val") == "window" else "000000"
        return sys_clr.get("lastClr", default).upper()
    return None


def theme_info(prs):
    root = etree.fromstring(theme_part(prs).blob)
    scheme = root.find(f"{qn('a:themeElements')}/{qn('a:clrScheme')}")
    colors = {}
    if scheme is not None:
        for slot in SLOTS:
            elem = scheme.find(qn(f"a:{slot}"))
            if elem is not None and _slot_hex(elem):
                colors[slot] = _slot_hex(elem)
    clrmap = dict(DEFAULT_CLRMAP)
    cm = prs.slide_master._element.find(qn("p:clrMap"))
    if cm is not None:
        clrmap.update({k: v for k, v in cm.attrib.items() if k in DEFAULT_CLRMAP})
    fonts = {}
    font_scheme = root.find(f"{qn('a:themeElements')}/{qn('a:fontScheme')}")
    for key, tag in (("major", "a:majorFont"), ("minor", "a:minorFont")):
        latin = font_scheme.find(f"{qn(tag)}/{qn('a:latin')}") if font_scheme is not None else None
        fonts[key] = latin.get("typeface") if latin is not None else "Calibri"
    return {"colors": colors, "clrmap": clrmap, "fonts": fonts, "name": root.get("name", "")}


def scheme_hex(info, val):
    if val == "phClr" or not val:
        return None
    slot = info["clrmap"].get(val, val) if val in ("bg1", "tx1", "bg2", "tx2") else val
    return info["colors"].get(slot)


def role_colors(info):
    return {role: scheme_hex(info, val) for role, val in themes.ROLE_SCHEME.items()}


def _apply_mods(hex6, elem):
    r, g, b = (int(hex6[i:i + 2], 16) / 255 for i in (0, 2, 4))
    for mod in elem:
        tag = etree.QName(mod).localname
        v = int(mod.get("val", "100000")) / 100000
        if tag in ("lumMod", "lumOff"):
            h, l, s = colorsys.rgb_to_hls(r, g, b)
            l = l * v if tag == "lumMod" else l + v
            r, g, b = colorsys.hls_to_rgb(h, min(1.0, max(0.0, l)), s)
        elif tag == "tint":
            r, g, b = (c + (1 - c) * (1 - v) for c in (r, g, b))
        elif tag == "shade":
            r, g, b = (c * v for c in (r, g, b))
    return "".join(f"{round(min(1.0, max(0.0, c)) * 255):02X}" for c in (r, g, b))


def color_hex(color_parent, info):
    if color_parent is None:
        return None
    for child in color_parent:
        tag = etree.QName(child).localname
        if tag == "srgbClr":
            base = child.get("val", "").upper()
        elif tag == "schemeClr":
            base = scheme_hex(info, child.get("val"))
        elif tag == "sysClr":
            base = _slot_hex(color_parent)
        elif tag == "prstClr":
            base = {"black": "000000", "white": "FFFFFF"}.get(child.get("val"))
        else:
            continue
        return _apply_mods(base, child) if base and len(base) == 6 else None
    return None


def slide_background_hex(slide, info):
    for owner in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        bg = owner._element.find(f"{qn('p:cSld')}/{qn('p:bg')}")
        if bg is None:
            continue
        bg_pr = bg.find(qn("p:bgPr"))
        if bg_pr is not None:
            fill = bg_pr[0] if len(bg_pr) else None
            if fill is not None and fill.tag == qn("a:solidFill"):
                return color_hex(fill, info)
            return None
        bg_ref = bg.find(qn("p:bgRef"))
        if bg_ref is not None:
            return color_hex(bg_ref, info)
    return scheme_hex(info, "bg1")


def apply_theme(prs, theme):
    c = theme["colors"]
    if theme["dark"]:
        slots = {"dk1": c["bg"], "lt1": c["text"], "dk2": c["rule"], "lt2": c["muted"]}
        # Keeps the slot names truthful (dk is the dark colour) while text
        # still resolves through tx1, so dark and light decks share one code path.
        clrmap = {"bg1": "dk1", "tx1": "lt1", "bg2": "dk2", "tx2": "lt2"}
    else:
        slots = {"dk1": c["text"], "lt1": c["bg"], "dk2": c["muted"], "lt2": c["rule"]}
        clrmap = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
    slots.update({"accent1": c["accent"], "accent2": c["series"], "accent3": c["third"],
                  "accent4": c["muted"], "accent5": c["accent"], "accent6": c["series"],
                  "hlink": c["accent"], "folHlink": c["muted"]})
    part = theme_part(prs)
    root = etree.fromstring(part.blob)
    elements = root.find(qn("a:themeElements"))
    scheme = elements.find(qn("a:clrScheme"))
    for slot, value in slots.items():
        elem = scheme.find(qn(f"a:{slot}"))
        for child in list(elem):
            elem.remove(child)
        etree.SubElement(elem, qn("a:srgbClr")).set("val", value)
    name = f"deck-builder {theme['name']}"
    scheme.set("name", name)
    root.set("name", name)
    font_scheme = elements.find(qn("a:fontScheme"))
    font_scheme.find(f"{qn('a:majorFont')}/{qn('a:latin')}").set("typeface", theme["fonts"]["major"])
    font_scheme.find(f"{qn('a:minorFont')}/{qn('a:latin')}").set("typeface", theme["fonts"]["minor"])
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    cm = prs.slide_master._element.find(qn("p:clrMap"))
    for key, value in clrmap.items():
        cm.set(key, value)


def placeholder_text_style(ph, master):
    """(align, anchor) a layout placeholder gives its text, inherited the way
    PowerPoint does it: the layout placeholder, then the master placeholder,
    then the master's title or body text style. Codes are the planner's:
    align l/ctr/r, anchor t/m/b."""
    base = getattr(ph, "_base_placeholder", None)
    align = anchor = None
    for el in [ph._element] + ([base._element] if base is not None else []):
        body = el.find(qn("p:txBody"))
        if body is None:
            continue
        lvl1 = body.find(f"{qn('a:lstStyle')}/{qn('a:lvl1pPr')}")
        if align is None and lvl1 is not None:
            align = lvl1.get("algn")
        body_pr = body.find(qn("a:bodyPr"))
        if anchor is None and body_pr is not None:
            anchor = body_pr.get("anchor")
    if align is None:
        kind = ph._element.ph.get("type", "obj")
        style = "p:titleStyle" if kind in ("title", "ctrTitle") else "p:bodyStyle"
        lvl1 = master._element.find(f"{qn('p:txStyles')}/{qn(style)}/{qn('a:lvl1pPr')}")
        align = lvl1.get("algn") if lvl1 is not None else None
    return {"ctr": "ctr", "r": "r"}.get(align, "l"), {"ctr": "m", "b": "b"}.get(anchor, "t")


def scale_master(prs, sx, sy):
    # The default template is 4:3; widening only the slide would leave the
    # master's placeholders, and every slide a user adds later, squeezed left.
    xfrm_tags = {qn("a:xfrm"), qn("p:xfrm")}
    owners = [prs.slide_master] + list(prs.slide_layouts)
    for owner in owners:
        for tag, kx, ky in ((qn("a:off"), "x", "y"), (qn("a:ext"), "cx", "cy")):
            for elem in owner._element.iter(tag):
                if elem.getparent().tag not in xfrm_tags:
                    continue
                elem.set(kx, str(int(round(int(elem.get(kx)) * sx))))
                elem.set(ky, str(int(round(int(elem.get(ky)) * sy))))


def set_alt_text(shape, text):
    shape._element.find(".//" + qn("p:cNvPr")).set("descr", text)


def alt_text(shape_element):
    c_nv_pr = shape_element.find(".//" + qn("p:cNvPr"))
    return c_nv_pr.get("descr", "") if c_nv_pr is not None else ""


_PPR_REPLACED = {qn(f"a:{t}") for t in (
    "lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts",
    "buFontTx", "buFont", "buNone", "buAutoNum", "buChar", "buBlip")}


def set_paragraph(paragraph, *, align="l", indent_pt=0.0, hanging_pt=0.0, line_spacing=1.0,
                  space_before_pt=0.0, bullet=None, bullet_font="Arial"):
    p_pr = paragraph._p.get_or_add_pPr()
    p_pr.set("algn", align)
    p_pr.set("marL", str(int(indent_pt * 12700)))
    p_pr.set("indent", str(-int(hanging_pt * 12700)) if bullet else "0")
    for child in list(p_pr):
        if child.tag in _PPR_REPLACED:
            p_pr.remove(child)
    new = []
    ln = etree.Element(qn("a:lnSpc"))
    etree.SubElement(ln, qn("a:spcPct")).set("val", str(int(round(line_spacing * 100000))))
    new.append(ln)
    sb = etree.Element(qn("a:spcBef"))
    etree.SubElement(sb, qn("a:spcPts")).set("val", str(int(round(space_before_pt * 100))))
    new.append(sb)
    if bullet:
        bf = etree.Element(qn("a:buFont"))
        bf.set("typeface", bullet_font)
        bc = etree.Element(qn("a:buChar"))
        bc.set("char", bullet)
        new += [bf, bc]
    else:
        new.append(etree.Element(qn("a:buNone")))
    # Schema order puts these before a:tabLst, a:defRPr and a:extLst.
    for i, elem in enumerate(new):
        p_pr.insert(i, elem)


def add_slide_number(paragraph, number):
    fld = etree.Element(qn("a:fld"))
    fld.set("id", "{" + str(uuid.uuid4()).upper() + "}")
    fld.set("type", "slidenum")
    r_pr = etree.SubElement(fld, qn("a:rPr"))
    r_pr.set("lang", "en-US")
    etree.SubElement(fld, qn("a:t")).text = str(number)
    end = paragraph._p.find(qn("a:endParaRPr"))
    if end is not None:
        end.addprevious(fld)
    else:
        paragraph._p.append(fld)
    return r_pr


def style_run_props(r_pr, *, size_pt, bold=False, italic=False, color=None, font_ref=None):
    r_pr.set("sz", str(int(round(size_pt * 100))))
    r_pr.set("b", "1" if bold else "0")
    r_pr.set("i", "1" if italic else "0")
    for child in list(r_pr):
        if child.tag in (qn("a:solidFill"), qn("a:latin")):
            r_pr.remove(child)
    if color:
        fill = etree.SubElement(r_pr, qn("a:solidFill"))
        etree.SubElement(fill, qn("a:schemeClr")).set("val", color)
    if font_ref:
        etree.SubElement(r_pr, qn("a:latin")).set("typeface", font_ref)


def set_table_style_none(table):
    tbl_pr = table._tbl.tblPr
    style_id = tbl_pr.find(qn("a:tableStyleId"))
    if style_id is None:
        style_id = etree.SubElement(tbl_pr, qn("a:tableStyleId"))
    style_id.text = NO_STYLE_TABLE
    tbl_pr.set("firstRow", "1")
    tbl_pr.set("bandRow", "0")


_TCPR_REPLACED = {qn(f"a:{t}") for t in (
    "lnL", "lnR", "lnT", "lnB", "solidFill", "noFill", "gradFill", "blipFill", "pattFill", "grpFill")}


def _line(tag, border):
    ln = etree.Element(qn(tag))
    if not border:
        ln.set("w", "0")
        etree.SubElement(ln, qn("a:noFill"))
        return ln
    width_pt, scheme = border
    ln.attrib.update({"w": str(int(width_pt * 12700)), "cap": "flat", "cmpd": "sng", "algn": "ctr"})
    fill = etree.SubElement(ln, qn("a:solidFill"))
    etree.SubElement(fill, qn("a:schemeClr")).set("val", scheme)
    etree.SubElement(ln, qn("a:prstDash")).set("val", "solid")
    return ln


def set_cell_borders(cell, bottom=None, top=None):
    tc_pr = cell._tc.get_or_add_tcPr()
    for child in list(tc_pr):
        if child.tag in _TCPR_REPLACED:
            tc_pr.remove(child)
    new = [_line("a:lnL", None), _line("a:lnR", None), _line("a:lnT", top), _line("a:lnB", bottom),
           etree.Element(qn("a:noFill"))]
    for i, elem in enumerate(new):
        tc_pr.insert(i, elem)


def fingerprint(prs):
    # A spec built against a template records this, and the build refuses a
    # template whose layouts or theme changed since, instead of drifting quietly.
    layouts = []
    for layout in prs.slide_layouts:
        phs = []
        for ph in layout.placeholders:
            pf = ph.placeholder_format
            phs.append([pf.idx, str(pf.type), ph.left, ph.top, ph.width, ph.height])
        layouts.append({"name": layout.name, "placeholders": phs})
    info = theme_info(prs)
    canon = {"size": [prs.slide_width, prs.slide_height], "layouts": layouts,
             "colors": info["colors"], "clrmap": info["clrmap"], "fonts": info["fonts"]}
    return hashlib.sha256(json.dumps(canon, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


SLIDE_PART_RE = re.compile(r"^ppt/slides/slide\d+\.xml$")
HEX6_RE = re.compile(r"^[0-9A-Fa-f]{6}$")


def validate_package(path):
    problems = []

    def add(code, part, message):
        problems.append({"severity": "error", "code": code, "part": part, "message": message})

    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        parsed = {}
        for name in sorted(names):
            if name.endswith((".xml", ".rels")):
                try:
                    parsed[name] = etree.fromstring(z.read(name))
                except etree.XMLSyntaxError as e:
                    add("ooxml-parse", name, f"XML does not parse: {e}")
        types = parsed.get("[Content_Types].xml")
        if types is not None:
            ns = "{http://schemas.openxmlformats.org/package/2006/content-types}"
            overrides = {o.get("PartName") for o in types.iter(ns + "Override")}
            defaults = {d.get("Extension", "").lower() for d in types.iter(ns + "Default")}
            for name in sorted(names - {"[Content_Types].xml"}):
                if name.endswith("/"):
                    continue
                ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
                if "/" + name not in overrides and ext not in defaults:
                    add("ooxml-content-type", name, "part has no content type")
        rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        for name, root in parsed.items():
            if not name.endswith(".rels"):
                continue
            source_dir = posixpath.dirname(posixpath.dirname(name))
            for rel in root.iter(rel_ns + "Relationship"):
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target", "")
                resolved = target[1:] if target.startswith("/") else posixpath.normpath(posixpath.join(source_dir, target))
                if resolved not in names:
                    add("ooxml-rel-target", name, f"relationship {rel.get('Id')} points at missing {resolved}")
        for name, root in parsed.items():
            if not SLIDE_PART_RE.match(name):
                continue
            # Theme style references add outlines and shadows that PowerPoint,
            # Keynote and Google Slides each draw differently.
            for sp in root.iter(qn("p:sp")):
                if sp.find(qn("p:style")) is not None:
                    add("ooxml-style", name, "shape carries a p:style theme reference")
                    break
            if next(root.iter(qn("p:cxnSp")), None) is not None:
                add("ooxml-connector", name, "connector shape present")
            for tag in ("a:rPr", "a:defRPr", "a:endParaRPr"):
                for elem in root.iter(qn(tag)):
                    sz = elem.get("sz")
                    if sz is not None and not (sz.isdigit() and 100 <= int(sz) <= 400000):
                        add("ooxml-size", name, f"font size sz={sz} outside 100..400000")
            for elem in root.iter(qn("a:srgbClr")):
                if not HEX6_RE.match(elem.get("val", "")):
                    add("ooxml-color", name, f"colour {elem.get('val')!r} is not 6 hex digits")
            seen = set()
            for elem in root.iter(qn("p:cNvPr")):
                sid = elem.get("id")
                if sid in seen:
                    add("ooxml-dup-id", name, f"shape id {sid} used twice")
                seen.add(sid)
    return problems
