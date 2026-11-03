#!/usr/bin/env python3
"""Make a Hebrew document right-to-left, or check that it is.

    python hebrew_rtl.py fix <file> [<file> ...]     fix in place (Word, PowerPoint, Excel, HTML)
    python hebrew_rtl.py check <file> [<file> ...]   report what is still left-to-right; exit 1 if anything is

Whatever library built the file (python-docx, docx for Node, python-pptx, pptxgenjs, openpyxl, a template),
running `fix` on it afterwards gives the same result:
  Word        every paragraph RTL (w:bidi), Hebrew runs RTL (w:rtl) with complex-script font, size and bold
              matching the Latin ones, tables RTL (w:bidiVisual), sections RTL, the same in the styles so new
              paragraphs typed or added later are RTL too. A "right" alignment on a Hebrew paragraph is
              removed: in an RTL paragraph Word reads left/right as start/end, so "right" puts the text on
              the left; an RTL paragraph with no alignment is right-aligned.
  PowerPoint  every paragraph with Hebrew rtl="1" and right-aligned (centred and justified stay), Hebrew runs
              lang he-IL with a complex-script font, tables rtl="1".
  Excel       every sheet right-to-left (alignment is left alone: in Excel it is visual, and a design may use it).
  HTML        <html lang="he" dir="rtl">.
Only the XML parts are rewritten; images, charts and everything else in the file are copied as they are.
Standard library only, on purpose: it ships inside the plugin and runs on the consultant's computer.
"""
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.dom import minidom

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
HEBREW = re.compile(r"[֐-׿יִ-ﭏ]")

# Child order inside Word elements (ECMA-376); Word rejects a file whose children are out of order.
P_PR = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
        "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
        "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
        "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
        "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
R_PR = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
        "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
        "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
        "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
TBL_PR = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW",
          "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
          "tblCaption", "tblDescription", "tblPrChange"]
SECT_PR = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc",
           "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg",
           "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"]
A_RPR = ["ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "effectLst", "effectDag",
         "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick",
         "hlinkMouseOver", "rtl", "extLst"]
WS = ["sheetPr", "dimension", "sheetViews"]


# ---------- XML helpers (minidom keeps every prefix and namespace declaration as it was) ----------

def kids(el, ns, name=None):
    return [c for c in el.childNodes if c.nodeType == c.ELEMENT_NODE and c.namespaceURI == ns
            and (name is None or c.localName == name)]


def kid(el, ns, name):
    found = kids(el, ns, name)
    return found[0] if found else None


def descendants(el, ns, name):
    return [e for e in el.getElementsByTagNameNS(ns, name)]


def text_of(el, ns, name="t"):
    return "".join(n.data for t in descendants(el, ns, name) for n in t.childNodes if n.nodeType == n.TEXT_NODE)


def qname(parent, name):
    return f"{parent.prefix}:{name}" if parent.prefix else name


def ensure(parent, name, order, ns=None):
    """The child `name` of `parent`, created in its schema position if missing."""
    ns = ns or parent.namespaceURI
    existing = kid(parent, ns, name)
    if existing is not None:
        return existing, False
    new = parent.ownerDocument.createElementNS(ns, qname(parent, name))
    pos = order.index(name) if name in order else len(order)
    before = None
    for c in kids(parent, ns):
        if c.localName in order and order.index(c.localName) > pos:
            before = c
            break
    parent.insertBefore(new, before)
    return new, True


def ensure_first(parent, name):
    """pPr/tblPr/rPr-style property element that must be the first child of `parent`."""
    existing = kid(parent, parent.namespaceURI, name)
    if existing is not None:
        return existing, False
    new = parent.ownerDocument.createElementNS(parent.namespaceURI, qname(parent, name))
    first = next((c for c in parent.childNodes if c.nodeType == c.ELEMENT_NODE), None)
    parent.insertBefore(new, first)
    return new, True


def wattr(el, name):
    return el.getAttributeNS(W, name) if el.hasAttributeNS(W, name) else None


def set_wattr(el, name, value):
    el.setAttributeNS(W, qname(el, name), value)


def is_off(el):
    return el is not None and wattr(el, "val") in ("0", "false", "off")


# ---------- Word ----------

def _w_complex_script(r_pr, found, fix):
    """Hebrew is drawn with the complex-script font, size and bold, which Word keeps apart from the Latin ones."""
    fonts = kid(r_pr, W, "rFonts")
    if fonts is not None and wattr(fonts, "cs") is None and wattr(fonts, "cstheme") is None:
        latin = wattr(fonts, "ascii") or wattr(fonts, "hAnsi")
        if latin:
            found["fonts"] += 1
            if fix:
                set_wattr(fonts, "cs", latin)
    for latin, cs in (("sz", "szCs"), ("b", "bCs"), ("i", "iCs")):
        src = kid(r_pr, W, latin)
        if src is not None and kid(r_pr, W, cs) is None:
            found["sizes"] += 1
            if fix:
                el, _ = ensure(r_pr, cs, R_PR)
                if wattr(src, "val") is not None:
                    set_wattr(el, "val", wattr(src, "val"))


def _w_paragraph(p, found, fix):
    p_pr = kid(p, W, "pPr")
    hebrew = bool(HEBREW.search(text_of(p, W)))
    bidi = kid(p_pr, W, "bidi") if p_pr is not None else None
    if bidi is None or is_off(bidi):
        found["paragraphs"] += 1
        if fix:
            p_pr = p_pr or ensure_first(p, "pPr")[0]
            bidi, _ = ensure(p_pr, "bidi", P_PR)
            if bidi.hasAttributeNS(W, "val"):
                bidi.removeAttributeNS(W, "val")
    jc = kid(p_pr, W, "jc") if p_pr is not None else None
    if hebrew and jc is not None and wattr(jc, "val") in ("right", "end"):
        found["alignment"] += 1
        if fix:
            p_pr.removeChild(jc)
    for r in kids(p, W, "r") + [r for h in kids(p, W, "hyperlink") for r in kids(h, W, "r")]:
        if not HEBREW.search(text_of(r, W)):
            continue
        r_pr = kid(r, W, "rPr")
        rtl = kid(r_pr, W, "rtl") if r_pr is not None else None
        if rtl is None or is_off(rtl):
            found["runs"] += 1
            if fix:
                r_pr = r_pr or ensure_first(r, "rPr")[0]
                rtl, _ = ensure(r_pr, "rtl", R_PR)
                if rtl.hasAttributeNS(W, "val"):
                    rtl.removeAttributeNS(W, "val")
        if r_pr is not None:
            _w_complex_script(r_pr, found, fix)


def _w_part(doc, found, fix):
    root = doc.documentElement
    for p in descendants(root, W, "p"):
        _w_paragraph(p, found, fix)
    for tbl in descendants(root, W, "tbl"):
        tbl_pr = kid(tbl, W, "tblPr")
        flag = kid(tbl_pr, W, "bidiVisual") if tbl_pr is not None else None
        if flag is None or is_off(flag):
            found["tables"] += 1
            if fix:
                tbl_pr = tbl_pr or ensure_first(tbl, "tblPr")[0]
                flag, _ = ensure(tbl_pr, "bidiVisual", TBL_PR)
                if flag.hasAttributeNS(W, "val"):
                    flag.removeAttributeNS(W, "val")
    for sect in descendants(root, W, "sectPr"):
        if kid(sect, W, "bidi") is None:
            found["sections"] += 1
            if fix:
                ensure(sect, "bidi", SECT_PR)


def _w_styles(doc, found, fix):
    root = doc.documentElement
    defaults = kid(root, W, "docDefaults")
    if defaults is not None:
        p_default = kid(defaults, W, "pPrDefault")
        p_pr = kid(p_default, W, "pPr") if p_default is not None else None
        if p_pr is None or kid(p_pr, W, "bidi") is None:
            found["styles"] += 1
            if fix:
                p_default = p_default or ensure(defaults, "pPrDefault", ["rPrDefault", "pPrDefault"])[0]
                p_pr = p_pr or ensure_first(p_default, "pPr")[0]
                ensure(p_pr, "bidi", P_PR)
        r_default = kid(defaults, W, "rPrDefault")
        r_pr = kid(r_default, W, "rPr") if r_default is not None else None
        if r_pr is not None:
            if fix:
                lang, _ = ensure(r_pr, "lang", R_PR)
                if wattr(lang, "bidi") != "he-IL":
                    set_wattr(lang, "bidi", "he-IL")
            _w_complex_script(r_pr, found, fix)
    for style in kids(root, W, "style"):
        if wattr(style, "type") not in ("paragraph", "character", "table"):
            continue
        p_pr = kid(style, W, "pPr")
        if p_pr is not None:
            bidi = kid(p_pr, W, "bidi")
            jc = kid(p_pr, W, "jc")
            if is_off(bidi):
                found["styles"] += 1
                if fix:
                    bidi.removeAttributeNS(W, "val")
            if jc is not None and wattr(jc, "val") in ("right", "end"):
                found["alignment"] += 1
                if fix:
                    p_pr.removeChild(jc)
        r_pr = kid(style, W, "rPr")
        if r_pr is not None:
            _w_complex_script(r_pr, found, fix)


def word(doc, name, found, fix):
    if name == "word/styles.xml":
        _w_styles(doc, found, fix)
    elif re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$", name):
        _w_part(doc, found, fix)


# ---------- PowerPoint ----------

def powerpoint(doc, name, found, fix):
    if not re.match(r"ppt/(slides/slide|notesSlides/notesSlide)\d+\.xml$", name):
        return
    root = doc.documentElement
    for p in descendants(root, A, "p"):
        if not HEBREW.search(text_of(p, A)):
            continue
        p_pr = kid(p, A, "pPr")
        if p_pr is None or p_pr.getAttribute("rtl") not in ("1", "true") or p_pr.getAttribute("algn") in ("", "l"):
            found["paragraphs"] += 1
            if fix:
                p_pr = p_pr or ensure_first(p, "pPr")[0]
                p_pr.setAttribute("rtl", "1")
                if p_pr.getAttribute("algn") in ("", "l"):
                    p_pr.setAttribute("algn", "r")
        for r in kids(p, A, "r"):
            if not HEBREW.search(text_of(r, A)):
                continue
            r_pr = kid(r, A, "rPr")
            if r_pr is None or r_pr.getAttribute("lang") != "he-IL" or (
                    kid(r_pr, A, "latin") is not None and kid(r_pr, A, "cs") is None):
                found["runs"] += 1
                if fix:
                    r_pr = r_pr or ensure_first(r, "rPr")[0]
                    r_pr.setAttribute("lang", "he-IL")
                    latin = kid(r_pr, A, "latin")
                    if latin is not None and kid(r_pr, A, "cs") is None:
                        cs, _ = ensure(r_pr, "cs", A_RPR, A)
                        cs.setAttribute("typeface", latin.getAttribute("typeface"))
    for tbl in descendants(root, A, "tbl"):
        tbl_pr = kid(tbl, A, "tblPr")
        if tbl_pr is None or tbl_pr.getAttribute("rtl") not in ("1", "true"):
            found["tables"] += 1
            if fix:
                tbl_pr = tbl_pr or ensure_first(tbl, "tblPr")[0]
                tbl_pr.setAttribute("rtl", "1")


# ---------- Excel ----------

def excel(doc, name, found, fix):
    root = doc.documentElement
    if re.match(r"xl/worksheets/sheet\d+\.xml$", name):
        views = kid(root, S, "sheetViews")
        if views is None or not kids(views, S, "sheetView"):
            found["sheets"] += 1
            if fix:
                views = views or ensure(root, "sheetViews", WS, S)[0]
                view = doc.createElementNS(S, qname(views, "sheetView"))
                view.setAttribute("workbookViewId", "0")
                view.setAttribute("rightToLeft", "1")
                views.appendChild(view)
            return
        for view in kids(views, S, "sheetView"):
            if view.getAttribute("rightToLeft") not in ("1", "true"):
                found["sheets"] += 1
                if fix:
                    view.setAttribute("rightToLeft", "1")


# ---------- files ----------

HANDLERS = {".docx": word, ".docm": word, ".dotx": word, ".pptx": powerpoint, ".xlsx": excel, ".xlsm": excel}
LABELS = {
    "paragraphs": "פסקאות שלא מימין לשמאל", "runs": "קטעי טקסט בעברית שלא מסומנים RTL",
    "alignment": "יישור שמציב עברית בצד שמאל", "tables": "טבלאות משמאל לימין",
    "sections": "עמודים (sections) משמאל לימין", "styles": "סגנונות משמאל לימין",
    "fonts": "גופן לעברית חסר", "sizes": "גודל או הדגשה שלא חלים על עברית", "sheets": "גיליונות משמאל לימין",
    "html": "דף בלי dir=\"rtl\"",
}
REPORTED = ["paragraphs", "runs", "alignment", "tables", "sections", "styles", "fonts", "sizes", "sheets", "html"]


def office(path, fix):
    handler = HANDLERS[path.suffix.lower()]
    found = dict.fromkeys(LABELS, 0)
    changed = {}
    with zipfile.ZipFile(path) as z:
        for info in z.infolist():
            if not info.filename.endswith(".xml"):
                continue
            raw = z.read(info)
            if b"<" not in raw:
                continue
            doc = minidom.parseString(raw)
            before = sum(found.values())
            handler(doc, info.filename, found, fix)
            if fix and sum(found.values()) != before:
                head = re.match(rb"<\?xml[^>]*\?>\s*", raw)
                decl = head.group(0).decode() if head else '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                changed[info.filename] = (decl + doc.documentElement.toxml()).encode("utf-8")
        if fix and changed:
            with tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix, dir=tempfile.gettempdir()) as tmp:
                tmp_path = Path(tmp.name)
            with zipfile.ZipFile(tmp_path, "w") as out:
                for info in z.infolist():
                    out.writestr(info, changed.get(info.filename) or z.read(info), compress_type=info.compress_type)
    if fix and changed:
        shutil.move(str(tmp_path), str(path))
    return found


def html_file(path, fix):
    found = dict.fromkeys(LABELS, 0)
    text = path.read_text(encoding="utf-8")
    tag = re.search(r"<html\b[^>]*>", text, re.I)
    ok = tag and re.search(r'\bdir\s*=\s*["\']?rtl', tag.group(0), re.I)
    if not ok:
        found["html"] += 1
        if fix:
            if tag:
                new = re.sub(r'\s(dir|lang)\s*=\s*("[^"]*"|\'[^\']*\'|\S+)', "", tag.group(0)[:-1]) + ' lang="he" dir="rtl">'
                text = text[:tag.start()] + new + text[tag.end():]
            else:
                text = '<!doctype html>\n<html lang="he" dir="rtl">\n<meta charset="utf-8">\n' + text + "\n</html>\n"
            path.write_text(text, encoding="utf-8")
    return found


def run(path, fix):
    suffix = path.suffix.lower()
    if suffix in HANDLERS:
        return office(path, fix)
    if suffix in (".html", ".htm"):
        return html_file(path, fix)
    return None


def main():
    for stream in (sys.stdout, sys.stderr):  # a Windows console in cp1252/cp1255 cannot print ✓
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 3 or sys.argv[1] not in ("fix", "check"):
        print(__doc__.strip().splitlines()[2])
        print(__doc__.strip().splitlines()[3])
        sys.exit(2)
    mode, paths = sys.argv[1], [Path(p) for p in sys.argv[2:]]
    bad = False
    for path in paths:
        if not path.exists():
            print(f"✗ {path}: הקובץ לא נמצא")
            bad = True
            continue
        found = run(path, mode == "fix")
        if found is None:
            print(f"– {path.name}: לא מסמך Word, PowerPoint, Excel או HTML. אין מה לבדוק")
            continue
        issues = {k: found[k] for k in REPORTED if found[k]}
        if mode == "fix":
            print(f"✓ {path.name}: " + ("תוקן מימין לשמאל (" + ", ".join(f"{LABELS[k]}: {v}" for k, v in issues.items()) + ")"
                                       if issues else "כבר מימין לשמאל"))
        elif issues:
            bad = True
            print(f"✗ {path.name}: " + ", ".join(f"{LABELS[k]}: {v}" for k, v in issues.items()))
        else:
            print(f"✓ {path.name}: מימין לשמאל")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
