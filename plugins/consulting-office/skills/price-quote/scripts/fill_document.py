#!/usr/bin/env python3
"""Fill the consultant's designed Word forms: only the blanks change, everything else stays byte for byte.

    python fill_document.py quote  <template> <out> --to "<name>" --date 03/11/2026 --hourly 450
                                   [--monthly 12000] [--meetings 4] [--drop hourly|monthly] [--service ייעוץ|גישור]
    python fill_document.py demand <template> <out> --to "<name>" --date 03/11/2026 --period "אוקטובר 2026"
                                   --amount 4500 --vat 18 --account 123456

The amount is before VAT; the total is worked out here. The file is written only if every blank was filled, then
goes through hebrew_rtl (fix and check). Standard library only, like hebrew_rtl.py: it ships inside the plugin and
runs on the consultant's computer.
"""
import argparse
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.dom import minidom

sys.path.insert(0, str(Path(__file__).resolve().parent))
import hebrew_rtl as H  # noqa: E402

W = H.W
BLANK = re.compile(r"_{3,}")


def fail(message):
    print(f"✗ {message}")
    sys.exit(1)


def money(value):
    text = f"{value:,.2f}"
    return text[:-3] if text.endswith(".00") else text


def set_text(t, text):
    t.firstChild.data = text
    t.setAttribute("xml:space", "preserve")


def runs_of(p):
    return [t for t in H.descendants(p, W, "t") if t.firstChild is not None]


def paragraph_with(doc, text):
    for p in H.descendants(doc.documentElement, W, "p"):
        if text in H.text_of(p, W):
            return p
    fail(f'לא נמצא בתבנית: "{text}"')


def run_with(doc, text):
    for t in H.descendants(doc.documentElement, W, "t"):
        if t.firstChild is not None and text in t.firstChild.data:
            return t
    fail(f'לא נמצא בתבנית: "{text}"')


def blank_in(p):
    for t in runs_of(p):
        if BLANK.search(t.firstChild.data):
            return t
    fail(f'אין מקום למלא ב: "{H.text_of(p, W)[:40]}"')


def remove_ancestor(node, tag):
    while node is not None and node.localName != tag:
        node = node.parentNode
    node.parentNode.removeChild(node)


def fill_quote(doc, a):
    head = run_with(doc, "לכבוד:")
    set_text(head, BLANK.sub(lambda m: "\0", head.firstChild.data).replace("\0", a.to, 1).replace("\0", a.date, 1))

    if a.service:
        title = paragraph_with(doc, "הצעת מחיר לייעוץ עסקי")
        parts = runs_of(title)
        set_text(parts[0], f"הצעת מחיר ל{a.service} עסקי")
        for t in parts[1:]:
            remove_ancestor(t, "r")

    if a.drop == "hourly":
        remove_ancestor(run_with(doc, "ייעוץ לפי שעה"), "tr")
        remove_ancestor(run_with(doc, "מועד התשלום לפי שעה"), "p")
    else:
        if a.hourly is None:
            fail("חסר מחיר לשעה (--hourly), או --drop hourly אם המסלול לא מוצע")
        t = blank_in(paragraph_with(doc, "ש״ח לשעה"))
        set_text(t, BLANK.sub(money(a.hourly), t.firstChild.data))

    if a.drop == "monthly":
        remove_ancestor(run_with(doc, "ריטיינר חודשי"), "tr")
        terms = run_with(doc, "במסלול החודשי")
        set_text(terms, re.sub(r"\s*במסלול החודשי[^.]*\.", "", terms.firstChild.data))
    else:
        if a.monthly is not None:
            set_text(run_with(doc, "מחיר חודשי לסיכום מראש"), f"{money(a.monthly)} ש״ח לחודש")
            remove_ancestor(run_with(doc, "לפי מספר הפגישות ומשכן"), "r")
        if a.meetings is not None:
            t = run_with(doc, "פגישות בחודש")
            set_text(t, re.sub(r"\d+", str(a.meetings), t.firstChild.data))
    if a.drop:
        set_text(run_with(doc, "מסלולי התקשרות"), "מסלול התקשרות")
    return f"הצעת מחיר ל{a.to}"


def fill_demand(doc, a):
    if a.amount is None or a.vat is None or not a.period or not a.account:
        fail("בדרישת תשלום חובה: --amount, --vat, --period, --account")
    # "תאריך" + blank + ":" as designed reads "תאריך03/11/2026:" once filled, so the colon moves to the label.
    date_runs = runs_of(paragraph_with(doc, "__________"))
    if date_runs[0].firstChild.data != "תאריך":
        fail('שורת "תאריך" בתבנית השתנתה')
    set_text(date_runs[0], "תאריך: ")
    set_text(date_runs[1], a.date)
    set_text(date_runs[2], "")

    to = paragraph_with(doc, "לכבוד")
    set_text(blank_in(to), a.to)
    period = paragraph_with(doc, "הפעילות בחודש")
    filled = blank_in(period)
    set_text(filled, a.period)
    # The colon after the period is its own run; left unmarked it jumps next to the year's digits.
    colon = runs_of(period)[-1]
    if colon.firstChild.data.strip() == ":":
        r_pr = H.kid(colon.parentNode, W, "rPr") or H.ensure_first(colon.parentNode, "rPr")[0]
        H.ensure(r_pr, "rtl", H.R_PR)

    before = round(a.amount, 2)
    total = round(a.amount * (1 + a.vat / 100), 2)
    for label, value in (("ייעוץ וליווי עסקי שוטף", before), ('סה"כ לפני מע"מ', before)):
        t = run_with(doc, label)
        set_text(t, t.firstChild.data.replace("0 ₪", f"{money(value)} ₪"))
    runs = runs_of(paragraph_with(doc, 'סה"כ לתשלום'))
    set_text(next(t for t in runs if t.firstChild.data.strip() == "0 ₪"), f"{money(total)} ₪")

    t = blank_in(paragraph_with(doc, "מספר חשבון"))
    set_text(t, BLANK.sub(a.account, t.firstChild.data))
    return f"{money(before)} ₪ לפני מע\"מ, {money(total)} ₪ כולל מע\"מ {a.vat:g}%"


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("kind", choices=["quote", "demand"])
    ap.add_argument("template", type=Path)
    ap.add_argument("out", type=Path)
    ap.add_argument("--to", required=True)
    ap.add_argument("--date", required=True)
    ap.add_argument("--overwrite", action="store_true")
    ap.add_argument("--hourly", type=float)
    ap.add_argument("--monthly", type=float)
    ap.add_argument("--meetings", type=int)
    ap.add_argument("--drop", choices=["hourly", "monthly"])
    ap.add_argument("--service", choices=["ייעוץ", "גישור"])
    ap.add_argument("--period")
    ap.add_argument("--amount", type=float)
    ap.add_argument("--vat", type=float)
    ap.add_argument("--account")
    a = ap.parse_args()
    # A company name like בע"מ goes into the file name, and Windows refuses these characters there.
    a.out = a.out.with_name(re.sub(r'[<>:"/\\|?*]', "", a.out.name))

    if a.out.exists() and not a.overwrite:
        fail(f"כבר קיים: {a.out}. שואלים לפני שמחליפים, ואז --overwrite")
    with zipfile.ZipFile(a.template) as z:
        doc = minidom.parseString(z.read("word/document.xml"))
        summary = (fill_quote if a.kind == "quote" else fill_demand)(doc, a)
        left = BLANK.search(H.text_of(doc.documentElement, W))
        if left:
            fail("נשאר בתבנית מקום ריק. הקובץ לא נכתב")
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp_path = Path(tmp.name)
        with zipfile.ZipFile(tmp_path, "w") as out:
            for info in z.infolist():
                data = z.read(info)
                if info.filename == "word/document.xml":
                    head = re.match(rb"<\?xml[^>]*\?>\s*", data)
                    data = (head.group(0) if head else b"") + doc.documentElement.toxml().encode("utf-8")
                out.writestr(info, data, compress_type=info.compress_type)
    H.office(tmp_path, True)
    if sum(H.office(tmp_path, False).values()):
        tmp_path.unlink()
        fail("הקובץ לא עבר את בדיקת הכיוון. לא נכתב")
    a.out.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(tmp_path), str(a.out))
    print(f"✓ {a.out}: {summary}")


if __name__ == "__main__":
    main()
