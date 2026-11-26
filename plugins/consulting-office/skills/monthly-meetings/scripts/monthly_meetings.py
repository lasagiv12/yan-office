#!/usr/bin/env python3
"""Monthly meetings table: every client meeting of one month in one Excel file.

    python monthly_meetings.py build data.json --xlsx "2026-11 פגישות חודשיות.xlsx"   create or rebuild the month
    python monthly_meetings.py read "2026-11 פגישות חודשיות.xlsx" [--json data.json]   the month back as a data file

The assistant collects the meetings (calendar, client files, meeting documents) and writes a data file; this
script only lays them out, so the table looks the same every month. Rebuilding a month that already has a file
keeps what the consultant typed in the "הערות" column (matched by date, start time and client).
Values, not formulas, so the file reads the same in previews and Protected View. Needs only openpyxl.
Standalone on purpose: it ships inside the plugin.

Data file (JSON):
  month: "YYYY-MM", consultant, updated: "YYYY-MM-DD"
  meetings: [{date: "YYYY-MM-DD", start: "HH:MM" or "", end: "HH:MM" or "", client, subject, place,
              participants, status, doc, source, notes}]
    status: one of STATUSES. doc: the meeting document's file name, or "". source: "יומן", "תיק" or "יומן ותיק".
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

MONTHS_HE = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"]
DAYS_HE = ["שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת", "ראשון"]  # date.weekday(): Monday = 0
STATUSES = ["התקיימה", "לא תועדה", "מתוכננת", "בוטלה"]
COLUMNS = [("תאריך", 12), ("יום", 9), ("שעה", 13), ("משך (שעות)", 11), ("לקוח", 20), ("נושא", 34),
           ("מקום", 16), ("משתתפים", 26), ("סטטוס", 12), ("מסמך פגישה", 34), ("מקור", 11), ("הערות", 34)]
WRAP = {"נושא", "משתתפים", "מסמך פגישה", "הערות", "מקום"}

TEAL, INK2 = "0F5E63", "5B6670"
FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor=TEAL)
TOTAL_FILL = PatternFill("solid", fgColor="E6F0EF")
STATUS_FILL = {"בוטלה": "EEEEEE", "לא תועדה": "FFF4CC", "מתוכננת": "E8F0FB"}
THIN = Side(style="thin", color="C9D3D2")


def fail(msg):
    print(f"שגיאה: {msg}", file=sys.stderr)
    sys.exit(1)


def parse_date(s):
    try:
        return dt.date.fromisoformat(str(s)[:10])
    except ValueError:
        fail(f"תאריך לא תקין: {s!r} (צריך YYYY-MM-DD)")


def hours(start, end):
    if not start or not end:
        return None
    try:
        a = dt.datetime.strptime(start, "%H:%M")
        b = dt.datetime.strptime(end, "%H:%M")
    except ValueError:
        fail(f"שעה לא תקינה: {start!r}–{end!r} (צריך HH:MM)")
    minutes = (b - a).total_seconds() / 60
    if minutes <= 0:
        fail(f"שעת הסיום לפני שעת ההתחלה: {start}–{end}")
    return round(minutes / 60, 2)


def key(m):
    return (str(m.get("date", ""))[:10], m.get("start") or "", (m.get("client") or "").strip())


def validate(data):
    month = str(data.get("month", ""))
    try:
        year, mon = (int(x) for x in month.split("-"))
        assert 1 <= mon <= 12
    except (ValueError, AssertionError):
        fail(f"month לא תקין: {month!r} (צריך YYYY-MM)")
    meetings = data.get("meetings") or []
    for m in meetings:
        d = parse_date(m.get("date"))
        if (d.year, d.month) != (year, mon):
            fail(f"פגישה מחוץ לחודש {month}: {m.get('date')} {m.get('client')}")
        if not (m.get("client") or "").strip():
            fail(f"פגישה בלי לקוח: {m.get('date')} {m.get('subject')}")
        if m.get("status") not in STATUSES:
            fail(f"סטטוס לא מוכר {m.get('status')!r} ({m.get('date')} {m.get('client')}). מותר: {', '.join(STATUSES)}")
        hours(m.get("start"), m.get("end"))
    return year, mon, sorted(meetings, key=key)


def read_rows(path):
    """The meetings sheet back as dicts (same keys as the data file)."""
    wb = load_workbook(path)
    if "פגישות" not in wb.sheetnames:
        fail(f"{path}: אין לשונית \"פגישות\"")
    ws = wb["פגישות"]
    head_row = None
    for r in range(1, 8):
        if ws.cell(row=r, column=1).value == "תאריך":
            head_row = r
            break
    if head_row is None:
        fail(f"{path}: לא נמצאה שורת הכותרות")
    names = [ws.cell(row=head_row, column=c).value for c in range(1, len(COLUMNS) + 1)]
    rows = []
    for r in range(head_row + 1, ws.max_row + 1):
        vals = dict(zip(names, (ws.cell(row=r, column=c).value for c in range(1, len(COLUMNS) + 1))))
        if not vals.get("תאריך") or not vals.get("לקוח") or vals.get("תאריך") == "סה״כ":
            continue
        d = vals["תאריך"]
        d = d.date() if isinstance(d, dt.datetime) else d
        if isinstance(d, str):
            d = dt.datetime.strptime(d, "%Y/%m/%d").date()
        start, _, end = (vals.get("שעה") or "").partition("–")
        rows.append({"date": d.isoformat(), "start": start.strip(), "end": end.strip(), "client": vals["לקוח"],
                     "subject": vals.get("נושא") or "", "place": vals.get("מקום") or "",
                     "participants": vals.get("משתתפים") or "", "status": vals.get("סטטוס") or "",
                     "doc": vals.get("מסמך פגישה") or "", "source": vals.get("מקור") or "",
                     "notes": vals.get("הערות") or ""})
    return rows


def style_header(ws, row):
    for c, (name, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=row, column=c, value=name)
        cell.font = Font(bold=True, color="FFFFFF", name=FONT)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = width


def build(data, out):
    year, mon, meetings = validate(data)
    out = Path(out)
    if out.exists():  # keep the consultant's notes
        typed = {key(r): r["notes"] for r in read_rows(out) if r.get("notes")}
        for m in meetings:
            if not m.get("notes") and key(m) in typed:
                m["notes"] = typed[key(m)]
    consultant = data.get("consultant") or ""
    title = f"פגישות {MONTHS_HE[mon - 1]} {year}" + (f" · {consultant}" if consultant else "")
    updated = data.get("updated") or dt.date.today().isoformat()

    wb = Workbook()
    ws = wb.active
    ws.title = "פגישות"
    ws.sheet_view.rightToLeft = True
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=16, name=FONT, color=TEAL)
    held = [m for m in meetings if m["status"] == "התקיימה"]
    held_hours = sum(hours(m.get("start"), m.get("end")) or 0 for m in held)
    ws["A2"] = (f"{len(held)} פגישות התקיימו · {held_hours:g} שעות · "
                f"עודכן {parse_date(updated).strftime('%Y/%m/%d')}")
    ws["A2"].font = Font(italic=True, name=FONT, color=INK2)
    head = 4
    style_header(ws, head)
    r = head
    for r, m in enumerate(meetings, start=head + 1):
        d = parse_date(m["date"])
        h = hours(m.get("start"), m.get("end"))
        when = f"{m['start']}–{m['end']}" if m.get("start") and m.get("end") else (m.get("start") or "")
        values = [d, DAYS_HE[d.weekday()], when, h, m["client"].strip(), m.get("subject") or "",
                  m.get("place") or "", m.get("participants") or "", m["status"], m.get("doc") or "",
                  m.get("source") or "", m.get("notes") or ""]
        fill = STATUS_FILL.get(m["status"])
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name=FONT, color="8A939B" if m["status"] == "בוטלה" else "17233B")
            cell.alignment = Alignment(vertical="top", wrap_text=COLUMNS[c - 1][0] in WRAP,
                                       horizontal="center" if c in (1, 2, 3, 4, 9, 11) else None)
            if fill and c == 9:
                cell.fill = PatternFill("solid", fgColor=fill)
        ws.cell(row=r, column=1).number_format = "yyyy/mm/dd"
        ws.cell(row=r, column=4).number_format = "0.##"
    if meetings:
        t = Table(displayName="tblMonthMeetings", ref=f"A{head}:{get_column_letter(len(COLUMNS))}{r}")
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(t)
    else:
        ws.cell(row=head + 1, column=1, value="אין פגישות בחודש הזה").font = Font(italic=True, name=FONT, color=INK2)
        r = head + 1
    total = r + 1
    ws.cell(row=total, column=1, value="סה״כ")
    ws.cell(row=total, column=3, value=f"{len(held)} התקיימו")
    ws.cell(row=total, column=4, value=held_hours)
    ws.cell(row=total, column=4).number_format = "0.##"
    for c in range(1, len(COLUMNS) + 1):
        cell = ws.cell(row=total, column=c)
        cell.font = Font(bold=True, name=FONT)
        cell.fill = TOTAL_FILL
        cell.border = Border(top=THIN)
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = ws.cell(row=head + 1, column=1)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{head}:{head}"

    # by client
    wc = wb.create_sheet("לפי לקוח")
    wc.sheet_view.rightToLeft = True
    wc["A1"] = f"לפי לקוח · {MONTHS_HE[mon - 1]} {year}"
    wc["A1"].font = Font(bold=True, size=14, name=FONT, color=TEAL)
    cols = [("לקוח", 22), ("התקיימו", 11), ("שעות", 10), ("לא תועדו", 11), ("מתוכננות", 11), ("בוטלו", 10)]
    for c, (name, width) in enumerate(cols, start=1):
        cell = wc.cell(row=3, column=c, value=name)
        cell.font = Font(bold=True, color="FFFFFF", name=FONT)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        wc.column_dimensions[get_column_letter(c)].width = width
    clients = sorted({m["client"].strip() for m in meetings})
    row = 3
    sums = [0, 0.0, 0, 0, 0]
    for row, client in enumerate(clients, start=4):
        mine = [m for m in meetings if m["client"].strip() == client]
        by = {s: [m for m in mine if m["status"] == s] for s in STATUSES}
        vals = [len(by["התקיימה"]), sum(hours(m.get("start"), m.get("end")) or 0 for m in by["התקיימה"]),
                len(by["לא תועדה"]), len(by["מתוכננת"]), len(by["בוטלה"])]
        sums = [a + b for a, b in zip(sums, vals)]
        wc.cell(row=row, column=1, value=client).font = Font(name=FONT)
        for c, v in enumerate(vals, start=2):
            cell = wc.cell(row=row, column=c, value=v)
            cell.font = Font(name=FONT)
            cell.alignment = Alignment(horizontal="center")
            if c == 3:
                cell.number_format = "0.##"
    row += 1
    wc.cell(row=row, column=1, value="סה״כ")
    for c, v in enumerate(sums, start=2):
        wc.cell(row=row, column=c, value=v)
    for c in range(1, len(cols) + 1):
        cell = wc.cell(row=row, column=c)
        cell.font = Font(bold=True, name=FONT)
        cell.fill = TOTAL_FILL
        cell.alignment = Alignment(horizontal="center")
    wc.cell(row=row, column=3).number_format = "0.##"
    wc.cell(row=row + 2, column=1, value="שעות: רק פגישות שהתקיימו, לפי היומן. פגישה בלי שעות ביומן לא נכנסת לסכום.").font = \
        Font(italic=True, name=FONT, color=INK2)

    wb.properties.creator = consultant or "consulting-office"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"✓ {out} · {len(meetings)} שורות, {len(held)} התקיימו, {held_hours:g} שעות")


def main():
    ap = argparse.ArgumentParser(description="טבלת פגישות חודשית")
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build")
    b.add_argument("data")
    b.add_argument("--xlsx", required=True)
    rd = sub.add_parser("read")
    rd.add_argument("xlsx")
    rd.add_argument("--json")
    a = ap.parse_args()
    if a.cmd == "build":
        build(json.loads(Path(a.data).read_text(encoding="utf-8")), a.xlsx)
    else:
        rows = read_rows(a.xlsx)
        text = json.dumps({"meetings": rows}, ensure_ascii=False, indent=2)
        if a.json:
            Path(a.json).write_text(text, encoding="utf-8")
        else:
            print(text)


if __name__ == "__main__":
    main()
