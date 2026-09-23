"""Excel demo: an agent builds a Sales Command Center, hands-free, then proves it.

    python demo_excel.py

What happens, in a brand-new Excel process of its own (your open workbooks are
never touched):
  1. clicks "Blank workbook" on the start screen through its accessibility tree
  2. pastes 200 generated sales orders and turns them into an Excel Table via the
     real Insert > Table dialog - reading the range the dialog proposes first
  3. adds a Revenue column with one structured formula; Excel fills all 200 rows
  4. builds a Dashboard sheet with dynamic-array formulas (SORT/UNIQUE/SUMIFS/
     XLOOKUP/SORTBY), cell styles, data bars, icon sets and a chart - all picked
     from Excel's own galleries by name
  5. reads every result back out of Excel and checks it against an independent
     calculation in Python - so the dashboard is proven, not just pretty

Keyboard input only ever goes to the demo's own Excel window: if another window
comes to the front, the demo pauses until you click the demo Excel window again.
"""
import ctypes
import os
import random
import re
import subprocess
import sys
import time
from collections import defaultdict
from ctypes import wintypes
from datetime import date, timedelta

import uiautomation as auto

EXCEL = r"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "demo_output")
N_ORDERS = 200

u32 = ctypes.windll.user32
k32 = ctypes.windll.kernel32
u32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
u32.LoadKeyboardLayoutW.argtypes = [wintypes.LPCWSTR, wintypes.UINT]
u32.LoadKeyboardLayoutW.restype = ctypes.c_void_p
u32.GetKeyboardLayout.argtypes = [wintypes.DWORD]
u32.GetKeyboardLayout.restype = ctypes.c_void_p
u32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, ctypes.c_void_p]
WM_INPUTLANGCHANGEREQUEST = 0x0050
EN_US = 0x0409

# ---------------------------------------------------------------------------
# Console narration
# ---------------------------------------------------------------------------
os.system("")                                   # enable ANSI colours in conhost
C = {"b": "\033[1m", "g": "\033[92m", "c": "\033[96m", "y": "\033[93m",
     "r": "\033[91m", "d": "\033[2m", "x": "\033[0m"}
T0 = time.time()
STATS = {"uia": 0, "keys": 0, "mouse": 0}


def banner(text):
    print(f"\n{C['b']}{C['c']}{'=' * 74}\n  {text}\n{'=' * 74}{C['x']}")


def stage(n, total, text):
    print(f"\n{C['b']}[{n}/{total}] {text}{C['x']}  {C['d']}t+{time.time() - T0:.1f}s{C['x']}")


def ok(text):
    print(f"   {C['g']}OK{C['x']}  {text}")


def info(text):
    print(f"   {C['d']}..{C['x']}  {text}")


# ---------------------------------------------------------------------------
# The data, and the independent truth to check Excel against
# ---------------------------------------------------------------------------
REGIONS = ["North", "South", "East", "West", "Central"]
PRODUCTS = {"OLED TV 65": 1899.0, "QLED TV 55": 999.0, "Soundbar Pro": 449.0,
            "Monitor 32": 379.0, "Projector 4K": 1299.0, "Smart Hub": 129.0}
CHANNELS = ["Online", "Retail", "Partner"]


def make_orders():
    rnd = random.Random(2026)
    start = date(2026, 1, 1)
    rows = []
    for i in range(N_ORDERS):
        product = rnd.choice(list(PRODUCTS))
        price = round(PRODUCTS[product] * rnd.choice([1.0, 1.0, 0.95, 0.9, 0.85]), 2)
        units = rnd.choice([1, 2, 3, 4, 5, 6, 8, 10, 12, 15, 20, 25])
        rows.append({"Order ID": f"SO-{1001 + i}",
                     "Date": (start + timedelta(days=rnd.randrange(180))).isoformat(),
                     "Region": rnd.choice(REGIONS), "Product": product,
                     "Channel": rnd.choice(CHANNELS), "Units": units, "Unit Price": price})
    return rows


def truth(rows):
    by_region, count, by_product, by_channel = defaultdict(float), defaultdict(int), defaultdict(float), defaultdict(float)
    best = None
    for r in rows:
        rev = r["Units"] * r["Unit Price"]
        by_region[r["Region"]] += rev
        count[r["Region"]] += 1
        by_product[r["Product"]] += rev
        by_channel[r["Channel"]] += rev
        if best is None or rev > best[0]:
            best = (rev, f"{r['Order ID']} - {r['Product']}")
    return {"region": dict(by_region), "count": dict(count), "total": sum(by_region.values()),
            "top_product": max(by_product, key=by_product.get),
            "top_channel": max(by_channel, key=by_channel.get), "biggest": best[1],
            "last_revenue": rows[-1]["Units"] * rows[-1]["Unit Price"]}


# ---------------------------------------------------------------------------
# Driving Excel
# ---------------------------------------------------------------------------
class Excel:
    def __init__(self):
        self.proc = None
        self.w = None
        self.hwnd = None

    # -- windows and focus -------------------------------------------------
    def launch(self):
        self.proc = subprocess.Popen([EXCEL, "/x"])
        deadline = time.time() + 90
        while time.time() < deadline:
            try:
                for top in auto.GetRootControl().GetChildren():
                    if top.ProcessId == self.proc.pid and top.ClassName == "XLMAIN":
                        self.w, self.hwnd = top, top.NativeWindowHandle
                        return
            except Exception:
                pass
            time.sleep(0.4)
        raise RuntimeError("the new Excel window did not appear within 90s")

    def input_language(self):
        thread = u32.GetWindowThreadProcessId(self.hwnd, None)
        return (u32.GetKeyboardLayout(thread) or 0) & 0xFFFF

    def force_english_input(self):
        """Ribbon key tips (Alt, H, O, R ...) only match English letters. With an
        Arabic (or any non-Latin) layout active the same keys type other characters,
        the key tips miss, and the text meant for a ribbon box lands in a cell.
        Only the demo's own Excel window is switched; every other window keeps its
        layout."""
        if self.input_language() == EN_US:
            return True
        hkl = u32.LoadKeyboardLayoutW("00000409", 1)          # KLF_ACTIVATE
        for _ in range(3):
            u32.PostMessageW(self.hwnd, WM_INPUTLANGCHANGEREQUEST, 0, hkl)
            deadline = time.time() + 2
            while time.time() < deadline:
                if self.input_language() == EN_US:
                    return True
                time.sleep(0.1)
        return False

    def bring_front(self):
        h = self.hwnd
        if u32.GetForegroundWindow() == h:
            return True
        if u32.IsIconic(h):
            u32.ShowWindow(h, 9)
        fg = u32.GetForegroundWindow()
        fg_thread = u32.GetWindowThreadProcessId(fg, None)
        me = k32.GetCurrentThreadId()
        u32.AttachThreadInput(me, fg_thread, True)
        try:
            u32.BringWindowToTop(h)
            u32.SetForegroundWindow(h)
        finally:
            u32.AttachThreadInput(me, fg_thread, False)
        time.sleep(0.3)
        return u32.GetForegroundWindow() == h

    def ours_in_front(self):
        """The main window, or one of its menus/dialogs (an open ribbon menu is its own
        foreground popup - pulling the main window forward would close it)."""
        fg = u32.GetForegroundWindow()
        if fg == self.hwnd:
            return True
        pid = wintypes.DWORD()
        u32.GetWindowThreadProcessId(fg, ctypes.byref(pid))
        return self.proc is not None and pid.value == self.proc.pid

    def guard(self):
        """Keys go to whatever window is in front - so never send one unless it is ours."""
        if self.ours_in_front():
            return
        # A closed dialog or a slow start can leave nothing of ours in front; take the
        # demo window back ourselves first, and only wait for a person if that fails.
        for _ in range(3):
            if self.bring_front():
                return
            time.sleep(0.4)
        warned = False
        deadline = time.time() + 300
        while not self.ours_in_front():
            if not warned:
                print(f"   {C['y']}PAUSED{C['x']}  another window is in front - click the demo "
                      f"Excel window ('{self.w.Name}') to continue. Nothing is typed meanwhile.")
                warned = True
            if time.time() > deadline:
                raise RuntimeError("the demo Excel window was not brought back within 5 minutes")
            time.sleep(0.3)
        info("resumed")
        time.sleep(0.3)

    # -- accessibility lookups ------------------------------------------------
    def find(self, root, name, ctype=None, parent=None, depth=20, timeout=5.0):
        deadline = time.time() + timeout
        while True:
            try:
                for c, _d in auto.WalkControl(root, maxDepth=depth):
                    try:
                        if c.Name == name and (ctype is None or c.ControlTypeName == ctype) and (
                                parent is None or c.GetParentControl().Name == parent):
                            return c
                    except Exception:
                        continue
            except Exception:
                pass            # the app rebuilt part of the tree mid-walk; walk again
            if time.time() > deadline:
                raise RuntimeError(f"could not find {ctype or 'control'} {name!r}")
            time.sleep(0.3)

    def invoke(self, ctrl):
        ctrl.GetPattern(auto.PatternId.InvokePattern).Invoke()
        STATS["uia"] += 1

    def expand(self, ctrl):
        ctrl.GetPattern(auto.PatternId.ExpandCollapsePattern).Expand()
        STATS["uia"] += 1

    def ribbon(self):
        return self.w.PaneControl(Name="Ribbon", searchDepth=6)

    def tab(self, name):
        tabs = self.ribbon().TabControl(Name="Ribbon Tabs", searchDepth=6)
        tabs.TabItemControl(Name=name, searchDepth=1).GetPattern(
            auto.PatternId.SelectionItemPattern).Select()
        STATS["uia"] += 1
        time.sleep(0.4)

    def dialog(self, title, timeout=8):
        deadline = time.time() + timeout
        while time.time() < deadline:
            for top in auto.GetRootControl().GetChildren():
                if top.ProcessId == self.proc.pid and top.Name == title:
                    return top
            d = self.w.WindowControl(Name=title, searchDepth=3)
            if d.Exists(0):
                return d
            time.sleep(0.3)
        raise RuntimeError(f"the {title!r} dialog did not open")

    # -- keyboard (guarded) --------------------------------------------------
    def keys(self, spec, wait=0.15):
        self.guard()
        auto.SendKeys(spec, interval=0.02, waitTime=wait)
        STATS["keys"] += 1

    def text(self, s, interval=0.012):
        self.guard()
        auto.SendKeys("".join({"{": "{{}", "}": "{}}"}.get(ch, ch) for ch in s),
                      interval=interval, waitTime=0.1)
        STATS["keys"] += 1

    def cell_mode(self):
        """Excel's own status-bar word: Ready / Enter / Edit / Point."""
        ctrl = getattr(self, "_mode_ctrl", None)
        if ctrl is not None:
            try:
                name = ctrl.Name or ""
                if name.startswith("Cell Mode"):
                    return name.replace("Cell Mode", "").strip()
            except Exception:
                pass
        self._mode_ctrl = None
        try:
            for c, _d in auto.WalkControl(self.w, maxDepth=9):
                if (c.Name or "").startswith("Cell Mode"):
                    self._mode_ctrl = c
                    return c.Name.replace("Cell Mode", "").strip()
        except Exception:
            pass
        return ""

    def wait_ready(self, timeout=20):
        """Poll until Excel reports Ready. Right after a paste, an auto-fit or a
        200-row fill it is still busy, and a copy taken then comes back empty."""
        deadline = time.time() + timeout
        escaped = False
        while True:
            mode = self.cell_mode()
            if mode in ("Ready", ""):
                return
            if time.time() > deadline:
                if escaped:
                    raise RuntimeError(f"Excel stayed in {mode!r} mode instead of Ready")
                self.keys("{Esc}", wait=0.3)
                escaped, deadline = True, time.time() + 5
            time.sleep(0.2)

    def name_box(self):
        nb = getattr(self, "_name_box", None)
        if nb is None or not nb.Exists(0):
            nb = self.w.EditControl(Name="Name Box", searchDepth=6)
            if not nb.Exists(20):
                raise RuntimeError("Excel's Name Box did not appear within 20s")
            self._name_box = nb
        return nb

    def goto(self, ref):
        self.guard()
        self.wait_ready()
        self.name_box().SetFocus()
        self.keys("{Ctrl}a", wait=0.05)
        self.text(ref, interval=0.01)
        self.keys("{Enter}", wait=0.2)

    def put(self, ref, value, interval=0.012):
        self.goto(ref)
        self.text(value, interval=interval)
        self.keys("{Enter}", wait=0.25)

    def paste(self, ref, tsv):
        auto.SetClipboardText(tsv)
        self.goto(ref)
        self.keys("{Ctrl}v", wait=0.8)
        self.keys("{Esc}", wait=0.1)

    def read(self, ref):
        """What Excel DISPLAYS for a range, read straight from the grid's cells.

        Each visible cell is a DataItem whose value is its displayed text (formatted
        formula results included). Not the clipboard: on a managed PC, copying OUT of
        Office can come back empty to other programs even though pasting in works.
        Going to the range first scrolls it into view, so its cells are in the tree."""
        self.goto(ref)
        self.wait_ready()
        first, _, last = ref.upper().partition(":")
        last = last or first
        c1, r1 = re.match(r"([A-Z]+)(\d+)", first).groups()
        c2, r2 = re.match(r"([A-Z]+)(\d+)", last).groups()
        cols = [chr(c) for c in range(ord(c1), ord(c2) + 1)]
        grid = self.w.DataGridControl(AutomationId="Grid", searchDepth=12)
        rows = []
        for r in range(int(r1), int(r2) + 1):
            row = []
            for col in cols:
                # depth 2: a cell inside an Excel Table sits under the table's own element
                cell = grid.DataItemControl(AutomationId=f"{col}{r}", searchDepth=2)
                if not cell.Exists(5):
                    raise RuntimeError(f"cell {col}{r} is not in Excel's accessibility tree")
                vp = cell.GetPattern(auto.PatternId.ValuePattern)
                row.append((vp.Value if vp else "") or "")
            rows.append(row)
        return rows

    def keytips(self, seq):
        """Alt, then one key-tip letter at a time. Sent as one burst ("{Alt}hor") Excel
        loses letters now and then, and whatever follows is typed into a cell instead."""
        self.wait_ready()
        self.guard()
        auto.PressKey(auto.Keys.VK_MENU, waitTime=0.05)
        auto.ReleaseKey(auto.Keys.VK_MENU, waitTime=0.5)
        STATS["keys"] += 1
        for ch in seq:
            self.keys(ch, wait=0.45)

    def rename_sheet(self, name):
        """Home > Format > Rename Sheet, proven by reading the tab names back.

        Done from Z1000, a cell nothing uses, and Z1000 is cleared afterwards: if the
        key tips ever miss, the name lands there instead of on real data."""
        for attempt in range(3):
            self.goto("Z1000")
            self.keytips("hor")
            self.text(name)
            self.keys("{Enter}", wait=0.5)
            self.wait_ready()
            self.goto("Z1000")
            self.keys("{Delete}", wait=0.2)
            if name in self.sheet_tabs():
                return
            info(f"renaming to {name!r} did not take (tabs {self.sheet_tabs()}), retrying")
        raise RuntimeError(f"could not rename the sheet to {name!r}")

    def open_menu(self, name, expect):
        """Expand a ribbon menu and PROVE it is open (its item `expect` is visible)
        before anything is typed - a shortcut sent to a closed menu lands in a cell."""
        m = self.find(self.ribbon(), name)
        for _ in range(3):
            self.expand(m)
            try:
                self.find(m, expect, timeout=2.0)
                return m
            except RuntimeError:
                self.keys("{Esc}", wait=0.3)
        raise RuntimeError(f"the {name!r} menu would not open")

    def gallery(self, ref, menu, accel, submenu, item, group=None):
        """Home > <menu> > <submenu> (via its accelerator) > <item>, by accessible name."""
        self.goto(ref)
        self.tab("Home")
        m = self.open_menu(menu, submenu)
        self.keys(accel, wait=0.6)
        try:
            target = self.find(m, item, ctype="ListItemControl", parent=group, timeout=3.0)
        except RuntimeError:
            self.keys("{Esc}{Esc}{Esc}", wait=0.3)
            raise
        self.invoke(target)
        time.sleep(0.4)

    def style(self, ref, name):
        self.goto(ref)
        self.tab("Home")
        cs = self.find(self.ribbon(), "Cell Styles")
        self.expand(cs)
        time.sleep(0.6)
        self.invoke(self.find(cs, name, ctype="ListItemControl"))
        time.sleep(0.3)

    def chart(self):
        for c, _d in auto.WalkControl(self.w, maxDepth=8):
            if c.ControlTypeName == "ImageControl" and (c.Name or "").startswith("Chart"):
                return c
        raise RuntimeError("no chart found on the sheet")

    def sheet_tabs(self):
        return [c.Name for c, _d in auto.WalkControl(self.w, maxDepth=10)
                if c.ControlTypeName == "TabItemControl" and c.AutomationId == "SheetTab"]


def num(s):
    cleaned = (s or "").replace(",", "").replace("%", "").strip()
    try:
        return float(cleaned)
    except ValueError:
        raise RuntimeError(f"expected a number from Excel but read {s!r}") from None


# ---------------------------------------------------------------------------
# The show
# ---------------------------------------------------------------------------
def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    os.makedirs(OUT, exist_ok=True)
    total = 9
    banner("SALES COMMAND CENTER - built in Excel by an agent, hands-free, then verified")
    print("   Please keep your hands off the mouse and keyboard for about two minutes.")
    print("   The demo only ever types into its OWN Excel window, and pauses if you switch away.")
    time.sleep(3)

    try:
        saved_clip = auto.GetClipboardText()
    except Exception:
        saved_clip = None
    rows = make_orders()
    want = truth(rows)
    xl = Excel()
    try:
        with auto.UIAutomationInitializerInThread():
            # 1 -------------------------------------------------------------
            stage(1, total, "Launch a private Excel process and pick 'Blank workbook' from its start screen")
            xl.launch()
            ok(f"Excel pid {xl.proc.pid}, window {xl.hwnd} - a separate process from any Excel you have open")
            blank = xl.find(xl.w, "Blank workbook", ctype="ListItemControl", depth=16, timeout=30)
            xl.invoke(blank)
            deadline = time.time() + 30
            while not xl.w.Name.startswith("Book") and time.time() < deadline:
                time.sleep(0.3)
            ok(f"clicked 'Blank workbook' through UI Automation -> {xl.w.Name!r}")
            wp = xl.w.GetPattern(auto.PatternId.WindowPattern)
            if wp:
                wp.SetWindowVisualState(auto.WindowVisualState.Maximized)
            xl.bring_front()
            xl.name_box()
            xl.wait_ready()
            if not xl.force_english_input():
                raise RuntimeError(f"could not switch the demo Excel window to English (US) input "
                                   f"(it is 0x{xl.input_language():04x}) - ribbon key tips need it")
            ok("keyboard layout of the demo Excel window set to English (US) - other windows keep theirs")

            # 2 -------------------------------------------------------------
            stage(2, total, f"Load {N_ORDERS} sales orders")
            xl.rename_sheet("Orders")
            header = list(rows[0])
            tsv = "\t".join(header) + "\n" + "\n".join(
                "\t".join(str(r[h]) for h in header) for r in rows) + "\n"
            xl.paste("A1", tsv)
            got_header = xl.read("A1:G1")[0]
            got_last = xl.read(f"A{N_ORDERS + 1}")[0][0]
            if got_header != header or got_last != rows[-1]["Order ID"]:
                raise RuntimeError(f"the paste did not land as expected: header {got_header}, "
                                   f"last order {got_last!r}")
            ok(f"{N_ORDERS} rows x {len(header)} columns pasted and read back "
               f"(header + last order {got_last}); sheet renamed 'Orders'")

            # 3 -------------------------------------------------------------
            stage(3, total, "Insert > Table - and read what the dialog proposes before accepting it")
            xl.goto("A1")
            xl.tab("Insert")
            xl.invoke(xl.find(xl.ribbon(), "Table", ctype="ButtonControl"))
            dlg = xl.dialog("Create Table")
            rng = xl.find(dlg, "Where is the data for your table?", ctype="EditControl")
            proposed = rng.GetPattern(auto.PatternId.ValuePattern).Value
            expected = f"=$A$1:$G${N_ORDERS + 1}"
            info(f"dialog proposes {proposed!r}, the data is {expected!r}")
            if proposed != expected:
                raise RuntimeError(f"Create Table proposed {proposed}, not {expected} - refusing to guess")
            xl.invoke(xl.find(dlg, "OK", ctype="ButtonControl"))
            time.sleep(0.8)
            ok("range confirmed, OK pressed - the orders are now Excel Table 'Table1'")

            # 4 -------------------------------------------------------------
            stage(4, total, "Add a Revenue column: ONE structured formula, Excel fills all rows")
            xl.put("H1", "Revenue")
            xl.put("H2", "=[@Units]*[@[Unit Price]]", interval=0.03)
            xl.goto(f"G2:H{N_ORDERS + 1}")
            xl.keys("{Ctrl}{Shift}1")
            xl.gallery(f"H2:H{N_ORDERS + 1}", "Conditional Formatting", "d", "Data Bars", "Green Data Bar",
                       group="Gradient Fill")
            xl.goto("A:H")
            xl.keytips("hoi")
            last = xl.read(f"H{N_ORDERS + 1}")[0][0]
            info(f"last row H{N_ORDERS + 1} reads {last}; Python says {want['last_revenue']:,.2f}")
            if abs(num(last) - want["last_revenue"]) > 0.005:
                raise RuntimeError("the calculated column did not fill down correctly")
            ok(f"formula auto-filled {N_ORDERS} rows, formatted, green data bars, columns auto-fit")
            xl.keys("{Ctrl}{Home}", wait=0.4)
            xl.w.CaptureToImage(os.path.join(OUT, "excel_orders.png"))

            # 5 -------------------------------------------------------------
            stage(5, total, "Add a Dashboard sheet (the 'Add Sheet' button, by name)")
            xl.invoke(xl.find(xl.w, "Add Sheet", ctype="ButtonControl", depth=12))
            time.sleep(0.6)
            xl.rename_sheet("Dashboard")
            ok(f"sheets now: {xl.sheet_tabs()}")

            # 6 -------------------------------------------------------------
            stage(6, total, "Write the dashboard with dynamic-array formulas")
            # Title formatting goes on while A1 is still empty: if a key tip ever
            # misses, the font name lands in A1 and the title below overwrites it.
            xl.style("A1", "Title")
            xl.goto("A1")
            xl.keytips("hff")
            xl.text("Segoe UI Semibold")
            xl.keys("{Enter}", wait=0.4)
            xl.goto("A1")
            xl.keytips("hfs")
            xl.text("24")
            xl.keys("{Enter}", wait=0.4)
            xl.put("A1", "SALES COMMAND CENTER")
            if xl.read("A1")[0][0] != "SALES COMMAND CENTER":
                raise RuntimeError("the title did not land in A1")
            xl.put("A2", '="Built hands-free by an AI agent  |  "&ROWS(Table1)&" orders  |  "&TEXT(TODAY(),"dd mmm yyyy")')
            for ref, head in zip("ABCDE", ["Region", "Revenue", "Orders", "Avg Order", "Share"]):
                xl.put(f"{ref}4", head, interval=0.008)
            formulas = [
                ("A5", "=SORT(UNIQUE(Table1[Region]))"),
                ("B5", "=SUMIFS(Table1[Revenue],Table1[Region],A5#)"),
                ("C5", "=COUNTIFS(Table1[Region],A5#)"),
                ("D5", "=B5#/C5#"),
                ("E5", "=B5#/SUM(B5#)"),
                ("A10", "TOTAL"),
                ("B10", "=SUM(Table1[Revenue])"),
                ("C10", "=ROWS(Table1)"),
                ("D10", "=B10/C10"),
                ("E10", "=SUM(E5#)"),
                ("G4", "Top product"),
                ("G5", "=INDEX(SORTBY(UNIQUE(Table1[Product]),SUMIFS(Table1[Revenue],Table1[Product],UNIQUE(Table1[Product])),-1),1)"),
                ("G7", "Top channel"),
                ("G8", "=INDEX(SORTBY(UNIQUE(Table1[Channel]),SUMIFS(Table1[Revenue],Table1[Channel],UNIQUE(Table1[Channel])),-1),1)"),
                ("G10", "Biggest single order"),
                ("G11", '=XLOOKUP(MAX(Table1[Revenue]),Table1[Revenue],Table1[Order ID]&" - "&Table1[Product])'),
            ]
            for ref, f in formulas:
                xl.put(ref, f, interval=0.02 if f.startswith("=") else 0.008)
                if f.startswith("="):
                    info(f"{ref:<4} {f}")
            ok("5 spilled regions, totals and 3 KPIs - every number is a live formula")

            # 7 -------------------------------------------------------------
            stage(7, total, "Style it - cell styles, number formats, data bars, icon sets, no gridlines")
            xl.style("A2", "Explanatory Text")
            xl.style("A4:E4", "Blue, Accent1")
            xl.style("A10:E10", "Total")
            for ref in ("G4", "G7", "G10"):
                xl.style(ref, "Heading 4")
            for ref in ("G5", "G8", "G11"):
                xl.style(ref, "Output")
            xl.goto("B5:B10")
            xl.keys("{Ctrl}{Shift}1")
            xl.goto("D5:D10")
            xl.keys("{Ctrl}{Shift}1")
            xl.goto("E5:E10")
            xl.keys("{Ctrl}{Shift}5")
            xl.gallery("B5:B9", "Conditional Formatting", "d", "Data Bars", "Blue Data Bar", group="Gradient Fill")
            xl.gallery("E5:E9", "Conditional Formatting", "i", "Icon Sets", "3 Arrows (Colored)")
            for block in ("A4:E10", "G4:G11"):          # not A1/A2: the title must not widen column A
                xl.goto(block)
                xl.keytips("hoi")
            xl.tab("View")
            g = xl.find(xl.ribbon(), "Gridlines", ctype="CheckBoxControl")
            tp = g.GetPattern(auto.PatternId.TogglePattern)
            if tp and tp.ToggleState == 1:
                tp.Toggle()
                STATS["uia"] += 1
            xl.tab("Home")
            ok("Title / Accent / Total / Output styles, formats, blue data bars, arrows, gridlines off")

            # 8 -------------------------------------------------------------
            stage(8, total, "Insert > Column chart > 'Clustered Column' - picked from the gallery by name")
            xl.goto("A4:B9")
            xl.tab("Insert")
            col = xl.find(xl.ribbon(), "Insert Column or Bar Chart")
            xl.expand(col)
            time.sleep(0.6)
            xl.invoke(xl.find(col, "Clustered Column", ctype="ListItemControl"))
            time.sleep(1.2)
            info("Excel drops a new chart mid-window; cut it and paste at I4 to dock it beside the KPIs")
            xl.keys("{Ctrl}x", wait=0.6)
            xl.goto("I4")
            xl.keys("{Ctrl}v", wait=1.0)
            chart = xl.chart()
            title = xl.find(chart, "Chart Title", depth=4)
            title.GetPattern(auto.PatternId.SelectionItemPattern).Select()
            STATS["uia"] += 1
            time.sleep(0.4)
            xl.text("Revenue by Region")
            xl.keys("{Enter}", wait=0.5)
            xl.keys("{Esc}", wait=0.3)
            xl.tab("Home")
            ok("clustered column chart docked at I4, titled 'Revenue by Region'")

            # 9 -------------------------------------------------------------
            stage(9, total, "TRUST, BUT VERIFY - read every number back from Excel, compare with Python")
            grid = xl.read("A5:E9")
            checks = []
            print(f"\n   {'Region':<9}{'Excel revenue':>16}{'Python revenue':>17}{'Orders':>8}   ")
            for region, rev, orders, _avg, share in grid:
                py_rev = want["region"].get(region, float("nan"))
                good = abs(num(rev) - round(py_rev, 2)) < 0.006 and int(num(orders)) == want["count"].get(region)
                checks.append(good)
                mark = f"{C['g']}MATCH{C['x']}" if good else f"{C['r']}DIFF{C['x']}"
                print(f"   {region:<9}{rev:>16}{py_rev:>17,.2f}{orders:>8}   {mark}  share {share}")
            tot = xl.read("B10:C10")[0]
            good = abs(num(tot[0]) - round(want["total"], 2)) < 0.006 and int(num(tot[1])) == N_ORDERS
            checks.append(good)
            print(f"   {'TOTAL':<9}{tot[0]:>16}{want['total']:>17,.2f}{tot[1]:>8}   "
                  f"{C['g'] + 'MATCH' + C['x'] if good else C['r'] + 'DIFF' + C['x']}")
            print()
            for ref, label, key in (("G5", "Top product", "top_product"),
                                    ("G8", "Top channel", "top_channel"),
                                    ("G11", "Biggest order", "biggest")):
                got = xl.read(ref)[0][0]
                good = got == want[key]
                checks.append(good)
                mark = f"{C['g']}MATCH{C['x']}" if good else f"{C['r']}DIFF{C['x']}"
                print(f"   {label:<14} Excel {got!r:<28} Python {want[key]!r:<28} {mark}")

            print("\n   The chart itself, read through its accessibility tree:")
            chart = xl.chart()
            bars = {}
            for c, _d in auto.WalkControl(chart, maxDepth=6):
                m = re.match(r'Series ".*" Point "(.+)"\s+Value: ([\d,.]+)', c.Name or "")
                if m:
                    bars[m.group(1)] = num(m.group(2))
            for region in sorted(want["region"]):
                got = bars.get(region)
                good = got is not None and abs(got - round(want["region"][region], 2)) < 0.006
                checks.append(good)
                mark = f"{C['g']}MATCH{C['x']}" if good else f"{C['r']}DIFF{C['x']}"
                print(f"   bar {region:<9}{(f'{got:,.2f}' if got is not None else 'missing'):>16}"
                      f"{want['region'][region]:>17,.2f}   {mark}")
            ttl_val = ""
            for c, _d in auto.WalkControl(chart, maxDepth=3):
                if c.ControlTypeName == "EditControl" and c.Name == "Chart Title":
                    lp = c.GetPattern(auto.PatternId.LegacyIAccessiblePattern)
                    ttl_val = (lp.Value if lp else "") or ""
            good = ttl_val == "Revenue by Region"
            checks.append(good)
            print(f"   chart title {ttl_val!r:<24} {C['g'] + 'MATCH' + C['x'] if good else C['r'] + 'DIFF' + C['x']}")
            xl.keys("{Ctrl}{Home}", wait=0.3)
            time.sleep(0.5)
            shot = os.path.join(OUT, "excel_dashboard.png")
            xl.w.CaptureToImage(shot)

        passed = sum(checks)
        elapsed = time.time() - T0
        banner(f"{'ALL' if passed == len(checks) else 'NOT ALL'} {passed}/{len(checks)} CHECKS "
               f"{'PASSED' if passed == len(checks) else 'PASSED - see DIFF lines above'}")
        print(f"   Built and verified in {elapsed:.0f}s:  {STATS['uia']} accessibility actions, "
              f"{STATS['keys']} guarded keystroke bursts, {STATS['mouse']} mouse clicks.")
        print(f"   Screenshots: {os.path.join(OUT, 'excel_orders.png')}")
        print(f"                {shot}")
        print("   The workbook is left open for you to explore - it has not been saved.")
        return 0 if passed == len(checks) else 1
    except Exception as e:
        print(f"\n   {C['r']}STOPPED{C['x']}  {e}")
        print("   The demo Excel window is left as it is so you can see where it stopped.")
        return 2
    finally:
        try:
            if saved_clip is not None:
                auto.SetClipboardText(saved_clip)
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(main())
