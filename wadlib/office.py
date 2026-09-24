"""Office through its own object model (COM) - the fast, exact path.

The accessibility tree is how a person's eyes see Excel; COM is how Excel sees itself.
Reading 10,000 cells through the grid takes minutes and only covers what is on screen;
through COM it takes a blink, includes formulas and hidden rows, and writes really
stick. wad attaches to the Excel / Word the user already has open (never a hidden one
unless --start), and still reads every write back."""
import json
import os
import time

from . import state
from .registry import Arg, WadError, command

XL_ERR_BASE = -2146828288          # 0x800A0000: Value2 returns errors as base + xlErr code
XL_ERRORS = {2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!", 2029: "#NAME?",
             2036: "#NUM!", 2042: "#N/A", 2043: "#GETTING_DATA", 2045: "#SPILL!",
             2046: "#CONNECT!", 2047: "#BLOCKED!", 2048: "#UNKNOWN!", 2049: "#FIELD!",
             2050: "#CALC!"}
BUSY = {-2147418111, -2146777998, -2147417846}     # call rejected / VBA ignore / retry later


def _com():
    try:
        import pythoncom
        import pywintypes
        import win32com.client
        return pythoncom, pywintypes, win32com.client
    except ImportError:
        raise WadError("MISSING_DEPENDENCY", "Office commands need pywin32",
                       "pip install pywin32") from None


def patient(fn, timeout=10.0):
    """Office refuses COM calls while a cell is being edited or a dialog is open. Retry
    for a while, then say what to do instead of hanging."""
    try:
        import pywintypes
        com_error = pywintypes.com_error
    except ImportError:                  # only reachable with a stand-in Office (tests)
        com_error = ()
    deadline = time.time() + timeout
    while True:
        try:
            return fn()
        except com_error as e:
            if e.hresult not in BUSY or time.time() > deadline:
                if e.hresult in BUSY:
                    raise WadError("OFFICE_BUSY", "Office is busy and refuses automation",
                                   "a cell is in edit mode or a dialog is open: "
                                   "`wad press esc --window <app>` and retry") from None
                detail = e.excinfo[2] if e.excinfo and len(e.excinfo) > 2 else e.strerror
                raise WadError("OFFICE_ERROR", f"Office said: {detail}") from None
            time.sleep(0.3)


def app(progid, start):
    pythoncom, pywintypes, client = _com()
    pythoncom.CoInitialize()
    try:
        return client.GetActiveObject(progid)
    except pywintypes.com_error:
        if not start:
            name = progid.split(".")[0]
            raise WadError("OFFICE_NOT_RUNNING", f"{name} is not running",
                           f"open {name} first (`wad launch {name.lower()}.exe`) or pass --start")
    a = client.Dispatch(progid)
    a.Visible = True
    return a


def cell_out(v):
    if isinstance(v, int) and XL_ERR_BASE + 1900 < v < XL_ERR_BASE + 2200:
        return XL_ERRORS.get(v - XL_ERR_BASE, "#ERROR!")
    if isinstance(v, float) and v.is_integer() and abs(v) < 2 ** 53:
        return int(v)
    return v


def grid_out(v):
    """Value2 is a scalar for one cell, a tuple of row tuples for a range."""
    if not isinstance(v, tuple):
        return [[cell_out(v)]]
    return [[cell_out(c) for c in row] for row in v]


def _sheet(xl, book, sheet):
    if book:
        wb = None
        for i in range(1, xl.Workbooks.Count + 1):
            b = xl.Workbooks(i)
            if book.lower() in (b.Name.lower(), b.FullName.lower()) or book.lower() in b.Name.lower():
                wb = b
                break
        if wb is None:
            raise WadError("NOT_FOUND", f"no open workbook like {book!r}",
                           "see `wad excel-info`")
    else:
        wb = xl.ActiveWorkbook
        if wb is None:
            raise WadError("NOT_FOUND", "Excel has no workbook open", "open or create one first")
    if not sheet:
        return wb, wb.ActiveSheet
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        if ws.Name.lower() == sheet.lower():
            return wb, ws
    raise WadError("NOT_FOUND", f"no sheet {sheet!r} in {wb.Name!r}",
                   "sheets: " + ", ".join(wb.Worksheets(i).Name
                                          for i in range(1, wb.Worksheets.Count + 1)))


def book_args():
    return [Arg("book", help="workbook name (default: the active one)"),
            Arg("sheet", help="sheet name (default: the active one)"),
            Arg("start", bool, help="start Excel if it is not running")]


@command("excel-info", "open workbooks, their sheets, used ranges and the selection",
         Arg("start", bool, help="start Excel if it is not running"),
         group="office", readonly=True)
def cmd_excel_info(args):
    xl = app("Excel.Application", args.start)

    def read():
        books = []
        for i in range(1, xl.Workbooks.Count + 1):
            b = xl.Workbooks(i)
            sheets = []
            for j in range(1, b.Worksheets.Count + 1):
                ws = b.Worksheets(j)
                sheets.append({"name": ws.Name, "used": ws.UsedRange.Address.replace("$", "")})
            books.append({"name": b.Name, "path": b.FullName, "saved": bool(b.Saved),
                          "sheets": sheets})
        active = xl.ActiveWorkbook
        sel = None
        try:
            sel = xl.Selection.Address.replace("$", "")
        except Exception:
            pass
        return {"books": books, "active_book": active.Name if active else None,
                "active_sheet": active.ActiveSheet.Name if active else None, "selection": sel}
    info = patient(read)
    lines = []
    for b in info["books"]:
        mark = "*" if b["name"] == info["active_book"] else " "
        lines.append(f"{mark} {b['name']}{'' if b['saved'] else ' (unsaved)'}")
        for s in b["sheets"]:
            lines.append(f"    {s['name']:<24} used {s['used']}")
    lines.append(f"selection: {info['selection']}")
    return {"ok": True, **info}, "\n".join(lines)


@command("excel-read", "read a range's values (numbers stay numbers, errors as #DIV/0! ...)",
         Arg("range", positional=True, optional=True,
             help="A1:D20, a name, or omit for the used range"),
         *book_args(),
         Arg("formulas", bool, help="return formulas instead of values"),
         Arg("text", bool, help="return what Excel displays (formatted), cell by cell"),
         group="office", readonly=True)
def cmd_excel_read(args):
    xl = app("Excel.Application", args.start)

    def read():
        wb, ws = _sheet(xl, args.book, args.sheet)
        rng = ws.Range(args.range) if args.range else ws.UsedRange
        addr = rng.Address.replace("$", "")
        if args.text:
            rows = [[rng.Cells(r, c).Text for c in range(1, rng.Columns.Count + 1)]
                    for r in range(1, rng.Rows.Count + 1)]
        elif args.formulas:
            rows = grid_out(rng.Formula)
        else:
            rows = grid_out(rng.Value2)
        return wb.Name, ws.Name, addr, rows
    book, sheet, addr, rows = patient(read)
    if len(rows) * (len(rows[0]) if rows else 0) > 5000:
        note = f"\n({len(rows)} rows - use --json or a smaller range for all of it)"
        shown = rows[:50]
    else:
        note, shown = "", rows
    text = f"{book} / {sheet} / {addr}\n" + "\n".join(
        "\t".join("" if v is None else str(v) for v in row) for row in shown) + note
    return {"ok": True, "book": book, "sheet": sheet, "range": addr, "rows": rows}, text


def _parse_values(args):
    if args.values:
        try:
            v = json.loads(args.values)
        except ValueError as e:
            raise WadError("USAGE", f"--values is not JSON: {e}") from None
        if not isinstance(v, list):
            v = [[v]]
        elif v and not isinstance(v[0], list):
            v = [v]
    elif args.tsv is not None:
        v = [line.split("\t") for line in args.tsv.replace("\r\n", "\n").rstrip("\n").split("\n")]
        v = [[_num(c) for c in row] for row in v]
    else:
        raise WadError("USAGE", "give --values '[[1,2],[3,4]]' or --tsv 'a<TAB>b'")
    width = max(len(r) for r in v)
    return [r + [None] * (width - len(r)) for r in v]


def _num(s):
    try:
        f = float(s)
        return int(f) if f.is_integer() and "." not in s else f
    except ValueError:
        return s


def _same(want, have):
    if isinstance(want, str) and want.startswith("="):
        return True                             # a formula: its result is checked for errors
    if want in (None, ""):
        return have in (None, "")
    if isinstance(want, (int, float)) and isinstance(have, (int, float)):
        return abs(want - have) < 1e-9 * max(1, abs(want))
    return str(want) == str(have)


@command("excel-write", "write values or formulas starting at a cell; read back and checked",
         Arg("range", positional=True, help="top-left cell, e.g. B2"),
         Arg("values", help="JSON: 5, [1,2,3] (a row) or [[1,2],[3,4]]; strings starting "
             "with = are formulas"),
         Arg("tsv", help="tab-separated rows instead of JSON"),
         *book_args(), group="office")
def cmd_excel_write(args):
    values = _parse_values(args)
    xl = app("Excel.Application", args.start)
    rows, cols = len(values), len(values[0])

    def write():
        wb, ws = _sheet(xl, args.book, args.sheet)
        rng = ws.Range(args.range).Cells(1, 1).Resize(rows, cols)
        has_formula = any(isinstance(c, str) and c.startswith("=") for r in values for c in r)
        if has_formula:
            for i, row in enumerate(values, 1):
                for j, v in enumerate(row, 1):
                    cell = rng.Cells(i, j)
                    if isinstance(v, str) and v.startswith("="):
                        try:
                            cell.Formula2 = v          # dynamic arrays spill, no implicit @
                        except Exception:
                            cell.Formula = v
                    else:
                        cell.Value2 = v
        else:
            rng.Value2 = tuple(tuple(r) for r in values)
        xl.Calculate()
        return wb.Name, ws.Name, rng.Address.replace("$", ""), grid_out(rng.Value2)
    book, sheet, addr, back = patient(write)
    wrong = [(i, j) for i, row in enumerate(values) for j, v in enumerate(row)
             if not _same(v, back[i][j])]
    errors = [f"{addr.split(':')[0]}+({i},{j})={back[i][j]}" for i, row in enumerate(back)
              for j, v in enumerate(row) if isinstance(v, str) and v.startswith("#")
              and v in XL_ERRORS.values()]
    state.trace("excel-write", {"book": book, "sheet": sheet, "range": addr, "cells": rows * cols})
    if wrong:
        i, j = wrong[0]
        raise WadError("VERIFY_FAILED", f"{len(wrong)} cell(s) did not keep their value, e.g. "
                       f"row {i + 1} col {j + 1}: wanted {values[i][j]!r}, got {back[i][j]!r}",
                       "the sheet may be protected, or data validation rejected it")
    text = f"wrote {rows}x{cols} to {book} / {sheet} / {addr} - read back OK"
    if errors:
        text += "\n  formula errors: " + ", ".join(errors[:10])
    return {"ok": True, "range": addr, "values": back, "formula_errors": errors}, text


@command("excel-run", "run an Excel command on the workbook: save, save-as, calculate, "
         "autofit, select",
         Arg("action", positional=True, choices=["save", "save-as", "calculate", "autofit",
                                                  "select"]),
         Arg("target", help="save-as: file path; autofit/select: a range"),
         *book_args(), group="office")
def cmd_excel_run(args):
    xl = app("Excel.Application", args.start)

    def run():
        wb, ws = _sheet(xl, args.book, args.sheet)
        if args.action == "save":
            wb.Save()
            return f"saved {wb.FullName}"
        if args.action == "save-as":
            if not args.target:
                raise WadError("USAGE", "save-as needs --target C:\\path\\file.xlsx")
            wb.SaveAs(args.target)
            return f"saved as {wb.FullName}"
        if args.action == "calculate":
            xl.CalculateFull()
            return "recalculated"
        rng = ws.Range(args.target) if args.target else ws.UsedRange
        if args.action == "autofit":
            rng.Columns.AutoFit()
            return f"autofit {rng.Address}"
        ws.Activate()
        rng.Select()
        return f"selected {rng.Address}"
    msg = patient(run, timeout=30)
    state.trace("excel-run", {"action": args.action, "target": args.target})
    return {"ok": True, "message": msg}, msg


# ---------------------------------------------------------------------------
# Word
# ---------------------------------------------------------------------------

def _doc(word, name):
    if not name:
        if word.Documents.Count == 0:
            raise WadError("NOT_FOUND", "Word has no document open")
        return word.ActiveDocument
    for i in range(1, word.Documents.Count + 1):
        d = word.Documents(i)
        if name.lower() in d.Name.lower():
            return d
    raise WadError("NOT_FOUND", f"no open document like {name!r}")


@command("word-read", "the text of an open Word document, paragraph by paragraph",
         Arg("doc", help="document name (default: the active one)"),
         Arg("start", bool, help="start Word if it is not running"),
         group="office", readonly=True)
def cmd_word_read(args):
    word = app("Word.Application", args.start)

    def read():
        d = _doc(word, args.doc)
        return d.Name, d.Content.Text
    name, text = patient(read)
    paras = [p for p in text.split("\r")]
    while paras and not paras[-1]:
        paras.pop()
    return {"ok": True, "doc": name, "paragraphs": paras}, f"{name}\n" + "\n".join(paras)


@command("word-write", "add text to an open Word document (end, start, or replace text); "
         "read back and checked",
         Arg("text", positional=True),
         Arg("at", choices=["end", "start", "replace"], default="end"),
         Arg("find", help="with --at replace: the text to replace (all occurrences)"),
         Arg("doc", help="document name (default: the active one)"),
         Arg("start", bool, help="start Word (with a new document) if it is not running"),
         group="office")
def cmd_word_write(args):
    word = app("Word.Application", args.start)

    def write():
        if word.Documents.Count == 0 and args.start:
            word.Documents.Add()
        d = _doc(word, args.doc)
        body = args.text.replace("\r\n", "\n").replace("\n", "\r")
        if args.at == "end":
            d.Content.InsertAfter(body)
        elif args.at == "start":
            d.Content.InsertBefore(body)
        else:
            if not args.find:
                raise WadError("USAGE", "--at replace needs --find TEXT")
            f = d.Content.Find
            f.ClearFormatting()
            f.Replacement.ClearFormatting()
            # Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
            #         MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace)
            f.Execute(args.find, False, False, False, False, False, True, 1, False, body, 2)
        return d.Name, d.Content.Text
    name, text = patient(write)
    want = args.text.replace("\r\n", "\n").replace("\n", "\r").strip()
    if want and want not in text:
        raise WadError("VERIFY_FAILED", "the text is not in the document after writing",
                       "the document may be protected or in read-only view")
    if args.at == "replace" and args.find and args.find in text and args.find not in args.text:
        raise WadError("VERIFY_FAILED", f"{args.find!r} is still in the document")
    state.trace("word-write", {"doc": name, "chars": len(args.text), "at": args.at})
    return {"ok": True, "doc": name}, f"wrote {len(args.text)} chars to {name} ({args.at}) - checked"


# ---------------------------------------------------------------------------
# Outlook - reading and drafting freely; SENDING only when a person allowed it
# ---------------------------------------------------------------------------
FOLDERS = {"inbox": 6, "sent": 5, "drafts": 16, "outbox": 4, "deleted": 3, "junk": 23}
OUTLOOK_LAST = "outlook_last.json"


def _folder(ns, name):
    key = (name or "inbox").lower()
    if key in FOLDERS:
        return ns.GetDefaultFolder(FOLDERS[key])
    folder = ns.GetDefaultFolder(6).Parent            # the mailbox root
    for part in name.replace("\\", "/").split("/"):
        try:
            folder = folder.Folders(part)
        except Exception:
            raise WadError("NOT_FOUND", f"no mail folder {name!r}",
                           "inbox, sent, drafts, outbox, deleted, junk, or Inbox/Sub") from None
    return folder


def _outlook(start):
    o = app("Outlook.Application", start)
    return o, o.GetNamespace("MAPI")


@command("outlook-list", "recent mail in a folder: sender, subject, time, unread (newest first)",
         Arg("folder", default="inbox", help="inbox | sent | drafts | deleted | Inbox/Sub"),
         Arg("limit", int, default=20), Arg("unread", bool, help="only unread"),
         Arg("search", help="only subjects or senders containing this"),
         Arg("start", bool, help="start Outlook if it is not running"),
         group="office", readonly=True)
def cmd_outlook_list(args):
    _, ns = _outlook(args.start)

    def read():
        items = _folder(ns, args.folder).Items
        items.Sort("[ReceivedTime]", True)
        if args.unread:
            items = items.Restrict("[UnRead] = True")
        rows, want = [], (args.search or "").lower()
        for i in range(1, items.Count + 1):
            it = items.Item(i)
            if getattr(it, "Class", 43) != 43:        # 43 = MailItem; skip meeting notices
                continue
            subject, sender = it.Subject or "", getattr(it, "SenderName", "") or ""
            if want and want not in subject.lower() and want not in sender.lower():
                continue
            rows.append({"n": len(rows) + 1, "id": it.EntryID, "subject": subject,
                         "from": sender, "received": str(getattr(it, "ReceivedTime", "")),
                         "unread": bool(getattr(it, "UnRead", False))})
            if len(rows) >= args.limit:
                break
        return rows
    rows = patient(read, timeout=30)
    state.write_json(os.path.join(state.STATE_DIR, OUTLOOK_LAST), [r["id"] for r in rows])
    text = "\n".join(f"{r['n']:>3} {'*' if r['unread'] else ' '} {r['received'][:16]}  "
                     f"{r['from'][:24]:<24} {r['subject'][:70]}" for r in rows) or "(no mail)"
    return {"ok": True, "mail": rows}, text


def _item(ns, which):
    ref = str(which)
    if ref.isdigit() and len(ref) < 6:              # a number from the last outlook-list
        ids = state.read_json(os.path.join(state.STATE_DIR, OUTLOOK_LAST), None) or []
        if not 1 <= int(ref) <= len(ids):
            raise WadError("NOT_FOUND", f"no message #{ref} in the last outlook-list",
                           "run outlook-list first, or pass the message id")
        ref = ids[int(ref) - 1]
    try:
        return ns.GetItemFromID(ref)
    except Exception:
        raise WadError("NOT_FOUND", "no mail item with that id (moved or deleted?)",
                       "run outlook-list again") from None


@command("outlook-read", "one message: headers, text, attachment names",
         Arg("which", positional=True, help="its number in the last outlook-list, or its id"),
         Arg("max-chars", int, default=20000), Arg("start", bool),
         group="office", readonly=True)
def cmd_outlook_read(args):
    _, ns = _outlook(args.start)

    def read():
        it = _item(ns, args.which)
        atts = [it.Attachments.Item(i).FileName for i in range(1, it.Attachments.Count + 1)]
        return {"subject": it.Subject, "from": getattr(it, "SenderName", ""),
                "from_address": getattr(it, "SenderEmailAddress", ""), "to": it.To,
                "cc": it.CC, "received": str(getattr(it, "ReceivedTime", "")),
                "body": (it.Body or "")[:args.max_chars], "attachments": atts}
    m = patient(read, timeout=30)
    text = (f"From: {m['from']} <{m['from_address']}>\nTo: {m['to']}\n"
            + (f"Cc: {m['cc']}\n" if m["cc"] else "") + f"Date: {m['received']}\n"
            f"Subject: {m['subject']}\n" + (f"Attachments: {', '.join(m['attachments'])}\n"
                                             if m["attachments"] else "") + "\n" + m["body"])
    return {"ok": True, **m}, text


@command("outlook-draft", "write a mail and save it in Drafts - never sends; read back",
         Arg("to", required=True, help="addresses, separated by ;"),
         Arg("subject", required=True), Arg("body", default=""), Arg("cc"),
         Arg("attach", help="file paths, separated by ;"),
         Arg("show", bool, help="also open it on screen for the person to review"),
         Arg("start", bool), group="office")
def cmd_outlook_draft(args):
    o, ns = _outlook(args.start)
    files = [f.strip() for f in (args.attach or "").split(";") if f.strip()]
    for f in files:
        if not os.path.exists(f):
            raise WadError("FILE_ERROR", f"attachment not found: {f}")

    def draft():
        m = o.CreateItem(0)                           # olMailItem
        m.To, m.Subject, m.Body = args.to, args.subject, args.body.replace("\\n", "\n")
        if args.cc:
            m.CC = args.cc
        for f in files:
            m.Attachments.Add(os.path.abspath(f))
        m.Save()
        back = ns.GetItemFromID(m.EntryID)
        if back.Subject != args.subject or back.Attachments.Count != len(files):
            raise WadError("VERIFY_FAILED", "the saved draft does not match what was written")
        if args.show:
            m.Display(False)
        return m.EntryID
    entry = patient(draft, timeout=30)
    state.trace("outlook-draft", {"to": args.to, "subject": args.subject, "files": len(files)})
    return ({"ok": True, "id": entry},
            f"draft saved (not sent): {args.subject!r} to {args.to}"
            + (f", {len(files)} attachment(s)" if files else ""))


@command("outlook-send", "SEND a saved draft - only when the policy allows sending",
         Arg("which", positional=True, help="the draft's id (from outlook-draft)"),
         Arg("start", bool), group="office")
def cmd_outlook_send(args):
    from .system import policy
    if not policy().get("allow_send"):
        raise WadError("POLICY_DENIED", "sending mail is off",
                       f'a person can allow it in {state.POLICY_FILE} with "allow_send": true; '
                       "until then leave the draft for them to send")
    _, ns = _outlook(args.start)

    def send():
        m = _item(ns, args.which)
        subject, to = m.Subject, m.To
        m.Send()
        return subject, to
    subject, to = patient(send, timeout=30)
    state.trace("outlook-send", {"to": to, "subject": subject})
    return {"ok": True, "subject": subject, "to": to}, f"SENT {subject!r} to {to}"


# ---------------------------------------------------------------------------
# PowerPoint
# ---------------------------------------------------------------------------
LAYOUTS = {"title-content": 2, "title-only": 11, "blank": 12, "title": 1}


def _pres(pp, name):
    if pp.Presentations.Count == 0:
        raise WadError("NOT_FOUND", "PowerPoint has no presentation open")
    if not name:
        return pp.ActivePresentation
    for i in range(1, pp.Presentations.Count + 1):
        p = pp.Presentations(i)
        if name.lower() in p.Name.lower():
            return p
    raise WadError("NOT_FOUND", f"no open presentation like {name!r}")


def _slide_texts(slide):
    texts = []
    for i in range(1, slide.Shapes.Count + 1):
        sh = slide.Shapes(i)
        if sh.HasTextFrame and sh.TextFrame.HasText:
            texts.append(sh.TextFrame.TextRange.Text.replace("\r", "\n"))
    return texts


def _title(slide):
    try:
        if slide.Shapes.HasTitle:
            return slide.Shapes.Title.TextFrame.TextRange.Text
    except Exception:
        pass
    return ""


def pres_args():
    return [Arg("pres", help="presentation name (default: the active one)"),
            Arg("start", bool, help="start PowerPoint if it is not running")]


@command("ppt-read", "the text of every slide (or one), with slide titles",
         Arg("slide", int, help="only this slide number"), *pres_args(),
         group="office", readonly=True)
def cmd_ppt_read(args):
    pp = app("PowerPoint.Application", args.start)

    def read():
        p = _pres(pp, args.pres)
        numbers = [args.slide] if args.slide else range(1, p.Slides.Count + 1)
        out = []
        for n in numbers:
            if not 1 <= n <= p.Slides.Count:
                raise WadError("NOT_FOUND", f"no slide {n} (there are {p.Slides.Count})")
            sl = p.Slides(n)
            out.append({"slide": n, "title": _title(sl), "texts": _slide_texts(sl)})
        return p.Name, out
    name, slides = patient(read)
    lines = [name]
    for s in slides:
        lines.append(f"--- slide {s['slide']}: {s['title']}")
        lines += [t for t in s["texts"] if t != s["title"]]
    return {"ok": True, "presentation": name, "slides": slides}, "\n".join(lines)


@command("ppt-add-slide", "add a slide with a title and body text; read back",
         Arg("title", required=True), Arg("body", default="", help="lines separated by \\n"),
         Arg("layout", choices=list(LAYOUTS), default="title-content"),
         Arg("at", int, help="position (default: the end)"), *pres_args(), group="office")
def cmd_ppt_add_slide(args):
    pp = app("PowerPoint.Application", args.start)

    def add():
        if pp.Presentations.Count == 0 and args.start:
            pp.Presentations.Add()
        p = _pres(pp, args.pres)
        at = args.at or p.Slides.Count + 1
        sl = p.Slides.Add(at, LAYOUTS[args.layout])
        if sl.Shapes.HasTitle:
            sl.Shapes.Title.TextFrame.TextRange.Text = args.title
        body = args.body.replace("\\n", "\r").replace("\n", "\r")
        if body:
            if sl.Shapes.Placeholders.Count < 2:
                raise WadError("USAGE", f"layout {args.layout!r} has no body placeholder",
                               "use --layout title-content")
            sl.Shapes.Placeholders(2).TextFrame.TextRange.Text = body
        back = p.Slides(at)
        if _title(back) != args.title:
            raise WadError("VERIFY_FAILED", "the new slide does not show the title")
        return p.Name, at, p.Slides.Count
    name, at, total = patient(add)
    state.trace("ppt-add-slide", {"pres": name, "at": at, "title": args.title})
    return ({"ok": True, "slide": at, "slides": total},
            f"added slide {at} {args.title!r} to {name} ({total} slides) - checked")


@command("ppt-replace", "replace text on every slide (shapes and tables); counts and checks",
         Arg("find", required=True), Arg("replace", required=True, help="the new text"),
         *pres_args(), group="office")
def cmd_ppt_replace(args):
    pp = app("PowerPoint.Application", args.start)

    def run():
        p = _pres(pp, args.pres)
        n = 0
        for si in range(1, p.Slides.Count + 1):
            sl = p.Slides(si)
            for i in range(1, sl.Shapes.Count + 1):
                sh = sl.Shapes(i)
                if not (sh.HasTextFrame and sh.TextFrame.HasText):
                    continue
                tr = sh.TextFrame.TextRange
                while tr.Find(args.find) is not None and args.find not in args.replace:
                    tr.Replace(args.find, args.replace)
                    n += 1
        left = sum(t.count(args.find) for si in range(1, p.Slides.Count + 1)
                   for t in _slide_texts(p.Slides(si)))
        return p.Name, n, left
    name, n, left = patient(run)
    if left and args.find not in args.replace:
        raise WadError("VERIFY_FAILED", f"{left} occurrence(s) of {args.find!r} remain")
    state.trace("ppt-replace", {"pres": name, "count": n})
    return {"ok": True, "replaced": n}, f"replaced {n} occurrence(s) in {name} - checked"


@command("ppt-save", "save the presentation (or --to a new file: .pptx, .pdf)",
         Arg("to", help="save as this path instead; .pdf exports a PDF"), *pres_args(),
         group="office")
def cmd_ppt_save(args):
    pp = app("PowerPoint.Application", args.start)
    target = args.to

    def save():
        p = _pres(pp, args.pres)
        if not target:
            p.Save()
            return p.FullName
        full = os.path.abspath(target)
        if full.lower().endswith(".pdf"):
            p.SaveAs(full, 32)                        # ppSaveAsPDF
        else:
            p.SaveAs(full)
        return full
    path = patient(save, timeout=60)
    return {"ok": True, "path": path}, f"saved {path}"
