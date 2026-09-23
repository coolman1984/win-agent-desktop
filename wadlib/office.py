"""Office through its own object model (COM) - the fast, exact path.

The accessibility tree is how a person's eyes see Excel; COM is how Excel sees itself.
Reading 10,000 cells through the grid takes minutes and only covers what is on screen;
through COM it takes a blink, includes formulas and hidden rows, and writes really
stick. wad attaches to the Excel / Word the user already has open (never a hidden one
unless --start), and still reads every write back."""
import json
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
    _, pywintypes, _ = _com()
    deadline = time.time() + timeout
    while True:
        try:
            return fn()
        except pywintypes.com_error as e:
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
