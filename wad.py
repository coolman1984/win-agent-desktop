"""wad - Windows Agent Desktop.

The ideas of lahfir/agent-desktop (macOS, Rust) rebuilt for Windows on UI Automation:
observe an app through its accessibility tree instead of pixels, hand out short refs
(@e3) for the elements, and act on refs through UIA patterns (headless) rather than
mouse coordinates. A real mouse click is used only when asked for (--headed).

    python wad.py windows
    python wad.py launch calculator
    python wad.py snapshot --window Calculator -i
    python wad.py click @e12
    python wad.py type @e5 "hello"
    python wad.py get @e3
    python wad.py press ctrl+a --window Notepad
"""
import argparse
import json
import os
import subprocess
import sys
import time
from datetime import datetime

import uiautomation as auto

STATE_DIR = os.path.join(os.environ.get("LOCALAPPDATA", "."), "win-agent-desktop")
SNAPSHOT_FILE = os.path.join(STATE_DIR, "last_snapshot.json")
TRACE_FILE = os.path.join(STATE_DIR, "trace.jsonl")

APPS = {
    "calculator": "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App",
    "notepad": "Microsoft.WindowsNotepad_8wekyb3d8bbwe!App",
    "paint": "Microsoft.Paint_8wekyb3d8bbwe!App",
}

INTERACTIVE_TYPES = {
    "ButtonControl", "EditControl", "CheckBoxControl", "RadioButtonControl",
    "ComboBoxControl", "ListItemControl", "MenuItemControl", "TabItemControl",
    "HyperlinkControl", "SliderControl", "SpinnerControl", "TreeItemControl",
    "SplitButtonControl", "DocumentControl", "DataItemControl",
}

PATTERNS = {
    "invoke": auto.PatternId.InvokePattern,
    "value": auto.PatternId.ValuePattern,
    "toggle": auto.PatternId.TogglePattern,
    "select": auto.PatternId.SelectionItemPattern,
    "expand": auto.PatternId.ExpandCollapsePattern,
    "text": auto.PatternId.TextPattern,
}


class WadError(Exception):
    def __init__(self, code, message, hint=""):
        super().__init__(message)
        self.code, self.message, self.hint = code, message, hint


# ---------------------------------------------------------------------------
# Output and trace
# ---------------------------------------------------------------------------

def emit(args, payload, text):
    if getattr(args, "json", False):
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(text)


def trace(command, detail):
    os.makedirs(STATE_DIR, exist_ok=True)
    with open(TRACE_FILE, "a", encoding="utf-8") as fh:
        fh.write(json.dumps({"t": datetime.now().isoformat(timespec="milliseconds"),
                             "cmd": command, **detail}, ensure_ascii=False) + "\n")


# ---------------------------------------------------------------------------
# Windows
# ---------------------------------------------------------------------------

def top_windows():
    out = []
    for w in auto.GetRootControl().GetChildren():
        if not w.Name or w.IsOffscreen:
            continue
        out.append(w)
    return out


def find_window(title=None, hwnd=None):
    if hwnd:
        w = auto.ControlFromHandle(int(hwnd))
        if not w:
            raise WadError("WINDOW_NOT_FOUND", f"no window with handle {hwnd}",
                           "run `wad.py windows` to list open windows")
        return w
    if not title:
        # Titles drift (Notepad renames an unsaved tab after its first line), so
        # with no window named, the window of the last snapshot is used by handle.
        try:
            snap = load_snapshot()
        except WadError:
            raise WadError("USAGE", "name a window with --window TITLE or --hwnd N") from None
        w = auto.ControlFromHandle(snap["hwnd"])
        if not w:
            raise WadError("WINDOW_GONE", f"the last snapshot's window {snap['window']!r} is closed",
                           "name a window with --window TITLE")
        return w
    low = title.lower()
    wins = top_windows()
    exact = [w for w in wins if w.Name.lower() == low]
    part = [w for w in wins if low in w.Name.lower()]
    hits = exact or part
    if not hits:
        raise WadError("WINDOW_NOT_FOUND", f"no open window titled like {title!r}",
                       "run `wad.py windows`, or `wad.py launch <app>` first")
    if len(hits) > 1 and not exact:
        names = ", ".join(repr(w.Name) for w in hits[:6])
        raise WadError("AMBIGUOUS_WINDOW", f"{len(hits)} windows match {title!r}: {names}",
                       "use a longer title or --hwnd")
    return hits[0]


def wait_window(title, timeout):
    deadline = time.time() + timeout
    while True:
        try:
            return find_window(title)
        except WadError as e:
            if e.code != "WINDOW_NOT_FOUND" or time.time() > deadline:
                raise
        time.sleep(0.3)


# ---------------------------------------------------------------------------
# Elements, refs and snapshots
# ---------------------------------------------------------------------------

def supported(ctrl):
    return [k for k, pid in PATTERNS.items() if ctrl.GetPattern(pid)]


def rect_of(ctrl):
    r = ctrl.BoundingRectangle
    return [r.left, r.top, r.right, r.bottom]


def is_interactive(ctrl, pats):
    return ctrl.ControlTypeName in INTERACTIVE_TYPES or bool(
        set(pats) & {"invoke", "toggle", "select", "expand"})


def value_of(ctrl):
    vp = ctrl.GetPattern(auto.PatternId.ValuePattern)
    if vp:
        try:
            return vp.Value
        except Exception:
            return None
    return None


def describe(ctrl):
    pats = supported(ctrl)
    return {
        "role": ctrl.ControlTypeName.replace("Control", ""),
        "name": ctrl.Name or "",
        "aid": ctrl.AutomationId or "",
        "class": ctrl.ClassName or "",
        "enabled": bool(ctrl.IsEnabled),
        "offscreen": bool(ctrl.IsOffscreen),
        "rect": rect_of(ctrl),
        "actions": pats,
        "value": value_of(ctrl),
        "interactive": is_interactive(ctrl, pats),
    }


def step_key(ctrl, index):
    return [ctrl.ControlTypeName, ctrl.AutomationId or "", ctrl.Name or "", index]


def walk(root, root_path, max_depth, counter, elements, parent_ref=None, depth=0):
    """Depth-first walk that records every element with the path needed to find it again."""
    node = describe(root)
    ref = f"e{counter[0]}"
    counter[0] += 1
    node.update(ref=ref, depth=depth, parent=parent_ref, path=root_path,
                rid=list(root.GetRuntimeId() or []))
    elements.append(node)
    children = []
    node["truncated"] = False
    if max_depth is None or depth < max_depth:
        try:
            children = root.GetChildren()
        except Exception:
            children = []
    else:
        node["truncated"] = bool(root.GetFirstChildControl())
    node["child_count"] = len(children)
    for i, child in enumerate(children):
        # Apps like Excel destroy and rebuild elements while they are being read;
        # one that vanished mid-walk is skipped rather than failing the snapshot.
        mark = len(elements)
        try:
            walk(child, root_path + [step_key(child, i)], max_depth, counter,
                 elements, ref, depth + 1)
        except Exception:
            del elements[mark:]
    return elements


def load_snapshot():
    try:
        with open(SNAPSHOT_FILE, encoding="utf-8") as fh:
            return json.load(fh)
    except FileNotFoundError:
        raise WadError("NO_SNAPSHOT", "no snapshot taken yet",
                       "run `wad.py snapshot --window <title>` first") from None


def lookup_ref(ref):
    snap = load_snapshot()
    ref = ref.lstrip("@")
    if ":" in ref:
        sid, ref = ref.split(":", 1)
        if sid != snap["id"]:
            raise WadError("STALE_SNAPSHOT", f"ref belongs to snapshot {sid}, latest is {snap['id']}",
                           "take a new snapshot and use its refs")
    for el in snap["elements"]:
        if el["ref"] == ref:
            return snap, el
    raise WadError("REF_NOT_FOUND", f"@{ref} is not in snapshot {snap['id']}",
                   "take a new snapshot and use its refs")


def resolve(ref):
    """Re-find a ref's live element by walking its recorded path from the window.

    The runtime id is checked at the end: a match means it is the very element that
    was observed; a mismatch with the same role and name means the app rebuilt that
    element (accepted, reported as re-resolved); anything else is refused."""
    snap, el = lookup_ref(ref)
    ctrl = auto.ControlFromHandle(snap["hwnd"])
    if not ctrl:
        raise WadError("WINDOW_GONE", f"window {snap['window']!r} is closed",
                       "launch the app again and take a new snapshot")
    for ctype, aid, name, index in el["path"]:
        kids = ctrl.GetChildren()
        same = [k for k in kids if k.ControlTypeName == ctype
                and (k.AutomationId or "") == aid and (k.Name or "") == name]
        if len(same) == 1:
            ctrl = same[0]
        elif index < len(kids) and kids[index].ControlTypeName == ctype:
            ctrl = kids[index]
        elif same:
            ctrl = same[0]
        else:
            raise WadError("ELEMENT_GONE", f"@{el['ref']} ({el['role']} {el['name']!r}) is no longer there",
                           "the UI changed - take a new snapshot")
    live = list(ctrl.GetRuntimeId() or [])
    status = "exact" if live == el["rid"] else "re-resolved"
    if status != "exact" and (ctrl.ControlTypeName.replace("Control", "") != el["role"]
                              or (ctrl.Name or "") != el["name"]):
        raise WadError("ELEMENT_CHANGED", f"@{el['ref']} now points at a different element",
                       "take a new snapshot")
    return ctrl, el, status


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

def keep_for_interactive(elements):
    """Interactive elements plus the containers that lead to them."""
    by_ref = {e["ref"]: e for e in elements}
    keep = set()
    for e in elements:
        if e["interactive"] and not e["offscreen"]:
            r = e["ref"]
            while r:
                keep.add(r)
                r = by_ref[r]["parent"]
    return keep


def render(elements, interactive_only):
    keep = keep_for_interactive(elements) if interactive_only else None
    lines = []
    for e in elements:
        if keep is not None and e["ref"] not in keep:
            continue
        label = f'"{e["name"]}"' if e["name"] else ""
        extra = []
        if e["aid"]:
            extra.append(f"id={e['aid']}")
        if e["value"] not in (None, ""):
            v = str(e["value"]).replace("\n", " ")
            extra.append(f"value={v[:40]!r}")
        if e["actions"]:
            extra.append("[" + ",".join(e["actions"]) + "]")
        if not e["enabled"]:
            extra.append("(disabled)")
        if e.get("truncated") and not e["interactive"]:
            extra.append(f"(more inside: --root @{e['ref']})")
        lines.append(f"{'  ' * e['depth']}@{e['ref']} {e['role']} {label} {' '.join(extra)}".rstrip())
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

def cmd_windows(args):
    rows = []
    for w in top_windows():
        rows.append({"hwnd": w.NativeWindowHandle, "title": w.Name, "class": w.ClassName,
                     "pid": w.ProcessId})
    text = "\n".join(f"{r['hwnd']:>10}  pid {r['pid']:<7} {r['title'][:60]!r}  ({r['class']})"
                     for r in rows)
    emit(args, {"ok": True, "windows": rows}, text or "(no windows)")


def cmd_launch(args):
    target = APPS.get(args.app.lower())
    if target:
        subprocess.Popen(["explorer.exe", f"shell:AppsFolder\\{target}"])
    else:
        subprocess.Popen([args.app])
    title = args.title or args.app
    w = wait_window(title, args.timeout)
    trace("launch", {"app": args.app, "hwnd": w.NativeWindowHandle})
    emit(args, {"ok": True, "hwnd": w.NativeWindowHandle, "title": w.Name},
         f"launched: {w.Name!r}  hwnd {w.NativeWindowHandle}")


def cmd_snapshot(args):
    if args.root:
        root, el, _ = resolve(args.root)
        snap_prev = load_snapshot()
        hwnd, wname, base_path = snap_prev["hwnd"], snap_prev["window"], el["path"]
    else:
        root = find_window(args.window, args.hwnd)
        hwnd, wname, base_path = root.NativeWindowHandle, root.Name, []
    started = time.time()
    elements = walk(root, base_path, args.depth, [1], [])
    sid = "s" + datetime.now().strftime("%H%M%S%f")[:9]
    snap = {"id": sid, "hwnd": hwnd, "window": wname, "taken": datetime.now().isoformat(),
            "elements": elements}
    os.makedirs(STATE_DIR, exist_ok=True)
    with open(SNAPSHOT_FILE, "w", encoding="utf-8") as fh:
        json.dump(snap, fh, ensure_ascii=False)
    trace("snapshot", {"window": wname, "snapshot": sid, "elements": len(elements)})
    body = render(elements, args.interactive)
    shown = body.count("\n") + 1 if body else 0
    head = (f"snapshot {sid}  window {wname!r}  {len(elements)} elements "
            f"({shown} shown) in {time.time() - started:.1f}s")
    emit(args, {"ok": True, "snapshot": sid, "window": wname,
                "elements": [e for e in elements
                             if not args.interactive or e["ref"] in keep_for_interactive(elements)]},
         head + "\n" + body)


def cmd_find(args):
    snap = load_snapshot()
    hits = []
    for e in snap["elements"]:
        if args.role and e["role"].lower() != args.role.lower():
            continue
        if args.name and args.name.lower() not in e["name"].lower():
            continue
        if args.aid and args.aid.lower() != e["aid"].lower():
            continue
        hits.append(e)
    text = "\n".join(f"@{e['ref']} {e['role']} {e['name']!r} [{','.join(e['actions'])}]"
                     for e in hits) or "(no match)"
    emit(args, {"ok": True, "matches": hits}, text)


def headed_click(ctrl):
    ctrl.Click(simulateMove=False)


def cmd_click(args):
    ctrl, el, status = resolve(args.ref)
    how = None
    if args.headed:
        headed_click(ctrl)
        how = "mouse"
    else:
        for name, pid, act in (
                ("invoke", auto.PatternId.InvokePattern, lambda p: p.Invoke()),
                ("toggle", auto.PatternId.TogglePattern, lambda p: p.Toggle()),
                ("select", auto.PatternId.SelectionItemPattern, lambda p: p.Select()),
                ("expand", auto.PatternId.ExpandCollapsePattern,
                 lambda p: p.Collapse() if p.ExpandCollapseState == 1 else p.Expand()),
                ("default-action", auto.PatternId.LegacyIAccessiblePattern,
                 lambda p: p.DoDefaultAction())):
            p = ctrl.GetPattern(pid)
            if p:
                act(p)
                how = name
                break
    if not how:
        raise WadError("NOT_CLICKABLE", f"@{el['ref']} supports no click action",
                       "retry with --headed to use a real mouse click")
    trace("click", {"ref": el["ref"], "role": el["role"], "name": el["name"], "via": how})
    emit(args, {"ok": True, "ref": el["ref"], "via": how, "resolved": status},
         f"clicked @{el['ref']} {el['role']} {el['name']!r} via {how} ({status})")


def cmd_type(args):
    ctrl, el, status = resolve(args.ref)
    vp = ctrl.GetPattern(auto.PatternId.ValuePattern)
    if vp and not vp.IsReadOnly and not args.keys:
        vp.SetValue((vp.Value or "") + args.text if args.append else args.text)
        how = "value"
    else:
        ctrl.SetFocus()
        if args.append:
            press_combo([auto.Keys.VK_CONTROL, auto.Keys.VK_END])
        escaped = "".join({"{": "{{}", "}": "{}}"}.get(ch, ch) for ch in args.text)
        auto.SendKeys(escaped, interval=0.01, waitTime=0.1)
        how = "keys"
    trace("type", {"ref": el["ref"], "chars": len(args.text), "via": how})
    emit(args, {"ok": True, "ref": el["ref"], "via": how},
         f"typed {len(args.text)} chars into @{el['ref']} {el['role']} via {how}")


def text_of(ctrl):
    tp = ctrl.GetPattern(auto.PatternId.TextPattern)
    if tp:
        try:
            return tp.DocumentRange.GetText(-1)
        except Exception:
            return None
    return None


def cmd_get(args):
    ctrl, el, status = resolve(args.ref)
    info = describe(ctrl)
    txt = text_of(ctrl)
    if txt is not None:
        info["text"] = txt
    tg = ctrl.GetPattern(auto.PatternId.TogglePattern)
    if tg:
        info["toggle"] = {0: "off", 1: "on", 2: "indeterminate"}.get(tg.ToggleState)
    info["resolved"] = status
    lines = [f"@{el['ref']} {info['role']} {info['name']!r}"]
    for k in ("value", "text", "toggle", "enabled", "rect", "actions"):
        if info.get(k) not in (None, ""):
            lines.append(f"  {k:<8}: {info[k]}")
    emit(args, {"ok": True, **info}, "\n".join(lines))


KEYS = {
    "ctrl": auto.Keys.VK_CONTROL, "control": auto.Keys.VK_CONTROL, "alt": auto.Keys.VK_MENU,
    "shift": auto.Keys.VK_SHIFT, "win": auto.Keys.VK_LWIN, "enter": auto.Keys.VK_RETURN,
    "tab": auto.Keys.VK_TAB, "esc": auto.Keys.VK_ESCAPE, "escape": auto.Keys.VK_ESCAPE,
    "space": auto.Keys.VK_SPACE, "backspace": auto.Keys.VK_BACK, "delete": auto.Keys.VK_DELETE,
    "home": auto.Keys.VK_HOME, "end": auto.Keys.VK_END, "up": auto.Keys.VK_UP,
    "down": auto.Keys.VK_DOWN, "left": auto.Keys.VK_LEFT, "right": auto.Keys.VK_RIGHT,
    "pageup": auto.Keys.VK_PRIOR, "pagedown": auto.Keys.VK_NEXT,
}


def vk(name):
    name = name.lower()
    if name in KEYS:
        return KEYS[name]
    if len(name) == 1 and name.isalnum():
        return ord(name.upper())
    if name.startswith("f") and name[1:].isdigit() and 1 <= int(name[1:]) <= 12:
        return auto.Keys.VK_F1 + int(name[1:]) - 1
    raise WadError("USAGE", f"unknown key {name!r}")


def press_combo(codes):
    for c in codes:
        auto.PressKey(c, waitTime=0.02)
    for c in reversed(codes):
        auto.ReleaseKey(c, waitTime=0.02)


def cmd_press(args):
    codes = [vk(k) for k in args.combo.split("+")]
    w = find_window(args.window, args.hwnd)
    w.SetActive()
    time.sleep(0.2)
    press_combo(codes)
    trace("press", {"combo": args.combo})
    emit(args, {"ok": True, "combo": args.combo}, f"pressed {args.combo}")


def cmd_focus(args):
    if args.ref:
        ctrl, el, _ = resolve(args.ref)
        ctrl.SetFocus()
        what = f"@{el['ref']}"
    else:
        w = find_window(args.window, args.hwnd)
        w.SetActive()
        what = repr(w.Name)
    emit(args, {"ok": True}, f"focused {what}")


def cmd_wait(args):
    deadline = time.time() + args.timeout
    w = wait_window(args.window, args.timeout)
    want = args.name.lower()
    while True:
        for ctrl, _depth in auto.WalkControl(w, maxDepth=args.depth):
            if want in (ctrl.Name or "").lower():
                emit(args, {"ok": True, "name": ctrl.Name},
                     f"found {ctrl.ControlTypeName.replace('Control', '')} {ctrl.Name!r}")
                return
        if time.time() > deadline:
            raise WadError("TIMEOUT", f"nothing named like {args.name!r} within {args.timeout}s")
        time.sleep(0.3)


def cmd_screenshot(args):
    w = find_window(args.window, args.hwnd)
    w.SetActive()
    time.sleep(0.3)
    path = os.path.abspath(args.out)
    w.CaptureToImage(path)
    trace("screenshot", {"window": w.Name, "path": path})
    emit(args, {"ok": True, "path": path}, f"saved {path}")


def cmd_clipboard(args):
    if args.action == "get":
        text = auto.GetClipboardText()
        emit(args, {"ok": True, "text": text}, text)
    elif args.action == "set":
        auto.SetClipboardText(args.text or "")
        emit(args, {"ok": True}, "clipboard set")
    else:
        auto.SetClipboardText("")
        emit(args, {"ok": True}, "clipboard cleared")


def cmd_close(args):
    w = find_window(args.window, args.hwnd)
    wp = w.GetPattern(auto.PatternId.WindowPattern)
    name = w.Name
    if not wp:
        raise WadError("NOT_CLOSABLE", f"{name!r} has no window close action")
    wp.Close()
    trace("close", {"window": name})
    emit(args, {"ok": True}, f"closed {name!r}")


def build_parser():
    p = argparse.ArgumentParser(prog="wad", description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--json", action="store_true", help="structured JSON output")
    sub = p.add_subparsers(dest="cmd", required=True)

    def win_opts(sp, required=False):
        sp.add_argument("--window", "-w", help="window title (substring)")
        sp.add_argument("--hwnd", type=int, help="window handle")

    sp = sub.add_parser("windows", help="list top-level windows")
    sp.set_defaults(fn=cmd_windows)

    sp = sub.add_parser("launch", help="start an app and wait for its window")
    sp.add_argument("app", help="calculator | notepad | paint | path to an exe")
    sp.add_argument("--title", help="window title to wait for (default: the app name)")
    sp.add_argument("--timeout", type=float, default=15)
    sp.set_defaults(fn=cmd_launch)

    sp = sub.add_parser("snapshot", help="accessibility tree of a window, with refs")
    win_opts(sp)
    sp.add_argument("-i", "--interactive", action="store_true",
                    help="only interactive elements and the containers leading to them")
    sp.add_argument("--depth", type=int, help="skeleton: stop at this depth")
    sp.add_argument("--root", help="drill down: snapshot only below this ref")
    sp.set_defaults(fn=cmd_snapshot)

    sp = sub.add_parser("find", help="search the last snapshot")
    sp.add_argument("--role")
    sp.add_argument("--name")
    sp.add_argument("--aid", help="AutomationId")
    sp.set_defaults(fn=cmd_find)

    sp = sub.add_parser("click", help="activate an element (UIA pattern; --headed for mouse)")
    sp.add_argument("ref")
    sp.add_argument("--headed", action="store_true")
    sp.set_defaults(fn=cmd_click)

    sp = sub.add_parser("type", help="set or type text into an element")
    sp.add_argument("ref")
    sp.add_argument("text")
    sp.add_argument("--append", action="store_true")
    sp.add_argument("--keys", action="store_true", help="send keystrokes instead of setting the value")
    sp.set_defaults(fn=cmd_type)

    sp = sub.add_parser("get", help="read an element's current state")
    sp.add_argument("ref")
    sp.set_defaults(fn=cmd_get)

    sp = sub.add_parser("press", help="key combo, e.g. ctrl+a, alt+f4, enter")
    sp.add_argument("combo")
    win_opts(sp)
    sp.set_defaults(fn=cmd_press)

    sp = sub.add_parser("focus", help="focus an element (ref) or a window")
    sp.add_argument("ref", nargs="?")
    win_opts(sp)
    sp.set_defaults(fn=cmd_focus)

    sp = sub.add_parser("wait", help="wait until an element with this name appears")
    sp.add_argument("--window", "-w", required=True)
    sp.add_argument("--name", required=True)
    sp.add_argument("--timeout", type=float, default=10)
    sp.add_argument("--depth", type=int, default=25)
    sp.set_defaults(fn=cmd_wait)

    sp = sub.add_parser("screenshot", help="capture a window to PNG")
    win_opts(sp)
    sp.add_argument("out")
    sp.set_defaults(fn=cmd_screenshot)

    sp = sub.add_parser("clipboard", help="get | set TEXT | clear")
    sp.add_argument("action", choices=["get", "set", "clear"])
    sp.add_argument("text", nargs="?")
    sp.set_defaults(fn=cmd_clipboard)

    sp = sub.add_parser("close", help="close a window through its window pattern")
    win_opts(sp)
    sp.set_defaults(fn=cmd_close)
    return p


def main(argv=None):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    args = build_parser().parse_args(argv)
    try:
        with auto.UIAutomationInitializerInThread():
            args.fn(args)
        return 0
    except WadError as e:
        payload = {"ok": False, "code": e.code, "message": e.message, "hint": e.hint}
        if args.json:
            print(json.dumps(payload, indent=2))
        else:
            print(f"ERROR {e.code}: {e.message}" + (f"\n  hint: {e.hint}" if e.hint else ""))
        return 1


if __name__ == "__main__":
    sys.exit(main())
