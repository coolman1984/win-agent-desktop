"""The core commands: observe, act, mouse, keyboard, windows.

Every command returns (payload, text): the payload is what --json, MCP clients and batch
files get, the text is the compact line a person or an agent reads."""
import glob
import os
import shutil
import subprocess
import time

import uiautomation as auto

from . import inputs, state, uia, verify, win32
from .registry import Arg, WadError, command, target_arg, verify_args, window_args

APPS = {
    "calculator": "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App",
    "notepad": "Microsoft.WindowsNotepad_8wekyb3d8bbwe!App",
    "paint": "Microsoft.Paint_8wekyb3d8bbwe!App",
}

# Windows whose ValuePattern accepts a value, reads it back, and throws it away.
VALUE_LIARS = {"XLMAIN"}


def ok(**payload):
    return {"ok": True, **payload}


def act_line(verb, el, how, changes, verified):
    line = f"{verb} {el['ref']} {el['role']} {el['name']!r} via {how}"
    if changes:
        line += "\n  changed: " + "; ".join(changes[:6])
    elif verified is False:
        line += "\n  (no visible effect yet - check with `get`/`snapshot` before repeating)"
    return line


def where(el, hwnd):
    """Where an action happened, durably: a selector (the one given, or one made from the
    element) and the window title - what `trace --export` needs to replay it."""
    sel = el["ref"] if not uia.is_ref(el["ref"]) else uia.selector_for(el)
    try:
        title = auto.ControlFromHandle(hwnd).Name if hwnd else None
    except Exception:
        title = None
    return {"selector": sel, "window": title}


def preflight(ctrl, el):
    if not ctrl.IsEnabled:
        raise WadError("NOT_ENABLED", f"{el['ref']} {el['role']} {el['name']!r} is disabled",
                       "something must happen first (fill a required field, pick an option)")


def target(args, timeout=0.0):
    return uia.resolve_target(args.target, getattr(args, "window", None),
                              getattr(args, "hwnd", None), timeout)


# ---------------------------------------------------------------------------
# Observe
# ---------------------------------------------------------------------------

@command("windows", "list top-level windows (title, handle, process)", group="observe",
         readonly=True)
def cmd_windows(args):
    rows = []
    for w in uia.top_windows():
        pid = w.ProcessId
        try:
            exe = win32.process_name(pid)
        except Exception:
            exe = ""
        rows.append({"hwnd": w.NativeWindowHandle, "title": w.Name, "class": w.ClassName,
                     "pid": pid, "exe": exe})
    text = "\n".join(f"{r['hwnd']:>10}  pid {r['pid']:<7} {r['exe'][:18]:<18} {r['title'][:60]!r}"
                     for r in rows)
    return ok(windows=rows), text or "(no windows)"


def _start_menu_shortcut(name):
    low = name.lower()
    roots = [os.path.join(os.environ.get("PROGRAMDATA", ""), "Microsoft", "Windows",
                          "Start Menu", "Programs"),
             os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows",
                          "Start Menu", "Programs")]
    found = []
    for root in roots:
        for p in glob.glob(os.path.join(root, "**", "*.lnk"), recursive=True):
            stem = os.path.splitext(os.path.basename(p))[0].lower()
            if stem == low:
                return p
            if low in stem:
                found.append(p)
    return min(found, key=len) if found else None


def _store_app_id(name):
    """AppUserModelID of a Store / packaged app by its Start-menu name."""
    try:
        out = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command",
             "Get-StartApps | ForEach-Object { $_.Name + '|' + $_.AppID }"],
            capture_output=True, text=True, timeout=20).stdout
    except Exception:
        return None
    low, part = name.lower(), None
    for line in out.splitlines():
        n, _, app_id = line.partition("|")
        if n.strip().lower() == low:
            return app_id.strip()
        if part is None and low in n.lower():
            part = app_id.strip()
    return part


@command("launch", "start an app (known name, exe, path, Start-menu or Store app name) and "
         "wait for its NEW window",
         Arg("app", positional=True, help="calculator | notepad | excel.exe | C:\\path\\app.exe | "
             "a Start-menu name"),
         Arg("title", help="window title to wait for (default: the app name)"),
         Arg("timeout", float, default=20.0, help="seconds to wait for the window"),
         group="windows")
def cmd_launch(args):
    before = {w.NativeWindowHandle for w in uia.top_windows()}
    name = args.app
    how = None
    if name.lower() in APPS:
        subprocess.Popen(["explorer.exe", f"shell:AppsFolder\\{APPS[name.lower()]}"])
        how = "store app"
    elif os.path.exists(name) or shutil.which(name):
        subprocess.Popen([shutil.which(name) or name])
        how = "executable"
    else:
        lnk = _start_menu_shortcut(name)
        if lnk:
            os.startfile(lnk)
            how = "start menu"
        else:
            app_id = _store_app_id(name)
            if not app_id:
                raise WadError("APP_NOT_FOUND", f"no executable, shortcut or app named {name!r}",
                               "give the full path to the .exe, or the name as shown in Start")
            subprocess.Popen(["explorer.exe", f"shell:AppsFolder\\{app_id}"])
            how = "store app"
    title = args.title or os.path.splitext(os.path.basename(name))[0]
    w = uia.wait_window(title, args.timeout, exclude=before)
    state.trace("launch", {"app": name, "hwnd": w.NativeWindowHandle, "via": how})
    return (ok(hwnd=w.NativeWindowHandle, title=w.Name, via=how),
            f"launched {w.Name!r}  hwnd {w.NativeWindowHandle}  ({how})")


@command("snapshot", "accessibility tree of a window, with refs for every element",
         *window_args(),
         Arg("interactive", bool, short="-i",
             help="only interactive elements and the containers leading to them"),
         Arg("depth", int, help="skeleton: stop at this depth (then drill in with --root)"),
         Arg("root", help="drill down: snapshot only below this ref"),
         Arg("popup", bool, help="snapshot the app's open menu / dialog instead of its window"),
         Arg("limit", int, default=uia.MAX_ELEMENTS, help="stop after this many elements"),
         group="observe", readonly=True)
def cmd_snapshot(args):
    if args.root:
        root, el, _, hwnd = uia.resolve_ref(args.root)
        prev = state.load_snapshot()
        wname, base_path = prev["window"], el["path"]
        hwnd = prev["hwnd"]
    else:
        if args.popup:
            main = uia.find_window(args.window, args.hwnd)
            root = uia.popup_of(main.NativeWindowHandle)
        else:
            root = uia.find_window(args.window, args.hwnd)
        hwnd, wname, base_path = root.NativeWindowHandle, root.Name, []
        if not hwnd:
            raise WadError("NO_HANDLE", "that surface has no window handle",
                           "snapshot its parent window and drill in with --root")
    started = time.time()
    snap = uia.take_snapshot(root, hwnd, wname, base_path, args.depth, args.limit)
    elements = snap["elements"]
    state.trace("snapshot", {"window": wname, "snapshot": snap["id"], "elements": len(elements)})
    body = uia.render(elements, args.interactive)
    shown = body.count("\n") + 1 if body else 0
    head = (f"snapshot {snap['id']}  window {wname!r}  {len(elements)} elements "
            f"({shown} shown) in {time.time() - started:.1f}s")
    notes = []
    if snap["limited"]:
        notes.append(f"stopped at {args.limit} elements - use --depth, then --root to drill in")
    if len(elements) <= 2:
        notes.append(_thin_tree_hint(hwnd))
    keep = uia.keep_for_interactive(elements) if args.interactive else None
    payload = ok(snapshot=snap["id"], window=wname, hwnd=hwnd, notes=notes,
                 elements=[e for e in elements if keep is None or e["ref"] in keep])
    return payload, "\n".join([head, body] + [f"note: {n}" for n in notes])


def _thin_tree_hint(hwnd):
    """An almost empty tree has three usual causes; say which one applies."""
    try:
        pid = win32.pid_of(hwnd)
        if not win32.is_admin() and win32.process_elevated(pid) is not False:
            return ("this app runs as administrator and wad does not - Windows hides it from "
                    "us (UIPI). Run the agent's terminal as administrator")
    except Exception:
        pass
    return ("almost nothing is exposed - the app draws its own UI (game, canvas, remote "
            "desktop). Use `wad screenshot` + `wad ocr` / `wad click-xy` instead")


@command("find", "search the last snapshot by role / name / id (prints refs and selectors)",
         Arg("role"), Arg("name", help="substring, case-insensitive"),
         Arg("aid", help="AutomationId"),
         Arg("interactive", bool, short="-i", help="only interactive elements"),
         group="observe", readonly=True)
def cmd_find(args):
    snap = state.load_snapshot()
    hits = []
    for e in snap["elements"]:
        if args.role and e["role"].lower() != args.role.lower():
            continue
        if args.name and args.name.lower() not in e["name"].lower():
            continue
        if args.aid and args.aid.lower() != e["aid"].lower():
            continue
        if args.interactive and not e["interactive"]:
            continue
        hits.append({**e, "selector": uia.selector_for(e)})
    text = "\n".join(f"{e['ref']} {e['role']} {e['name']!r} [{','.join(e['actions'])}]"
                     f"   selector: {e['selector']}" for e in hits) or "(no match)"
    return ok(matches=hits), text


@command("get", "read an element's live state: value, text, toggle, expanded, selected...",
         target_arg(), *window_args(), group="observe", readonly=True)
def cmd_get(args):
    ctrl, el, status, _ = target(args)
    info = uia.describe(ctrl)
    txt = uia.text_of(ctrl)
    if txt is not None:
        info["text"] = txt
    info.update(uia.states_of(ctrl))
    info["focused"] = bool(ctrl.HasKeyboardFocus)
    info["resolved"] = status
    info["selector"] = uia.selector_for(info)
    lines = [f"{el['ref']} {info['role']} {info['name']!r}"]
    for k in ("value", "text", "toggle", "expanded", "selected", "focused", "enabled", "rect",
              "actions", "selector"):
        if info.get(k) not in (None, ""):
            v = info[k]
            if k == "text" and len(str(v)) > 2000:
                v = str(v)[:2000] + f"... ({len(info[k])} chars)"
            lines.append(f"  {k:<9}: {v}")
    return ok(ref=el["ref"], **info), "\n".join(lines)


@command("wait", "wait until text / an element / a window appears (or is gone)",
         Arg("window", short="-w", help="window title; with nothing else, wait for the window"),
         Arg("hwnd", int),
         Arg("name", help="text to wait for in any element name or field value"),
         Arg("target", help="a selector to wait for (role=Button name=OK)"),
         Arg("gone", bool, help="wait until it disappears instead"),
         Arg("timeout", float, default=10.0),
         group="observe", readonly=True)
def cmd_wait(args):
    deadline = time.time() + args.timeout
    if not args.name and not args.target:
        if not args.window:
            raise WadError("USAGE", "wait for what? give --name, --target or --window")
        if args.gone:
            while any(args.window.lower() in w.Name.lower() for w in uia.top_windows()):
                if time.time() > deadline:
                    raise WadError("TIMEOUT", f"window {args.window!r} still open")
                time.sleep(0.3)
            return ok(), f"window {args.window!r} is gone"
        w = uia.wait_window(args.window, args.timeout)
        return ok(hwnd=w.NativeWindowHandle, title=w.Name), f"window {w.Name!r} is open"
    while True:
        try:
            w = uia.find_window(args.window, args.hwnd)
        except WadError as e:
            if e.code != "WINDOW_NOT_FOUND" or time.time() > deadline:
                raise
            time.sleep(0.3)
            continue
        if args.target:
            try:
                ctrl = uia.find_selector(w, args.target)
                found = True
            except WadError as e:
                if e.code == "AMBIGUOUS_TARGET":
                    found, ctrl = True, None
                elif e.code == "ELEMENT_NOT_FOUND":
                    found, ctrl = False, None
                else:
                    raise
            what = args.target
        else:
            found, ctrl = verify.text_present(w.NativeWindowHandle, args.name), None
            what = args.name
        if found != args.gone:
            state_word = "gone" if args.gone else "present"
            desc = f" ({uia.role_of(ctrl)} {ctrl.Name!r})" if ctrl is not None else ""
            return ok(what=what, state=state_word), f"{what!r} is {state_word}{desc}"
        if time.time() > deadline:
            raise WadError("TIMEOUT", f"{what!r} {'still there' if args.gone else 'not found'} "
                           f"after {args.timeout}s",
                           "snapshot to see what the app shows instead")
        time.sleep(0.3)


# ---------------------------------------------------------------------------
# Act (accessibility patterns - no mouse, no focus stealing)
# ---------------------------------------------------------------------------

def _semantic_click(ctrl):
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
            return name
    return None


@command("click", "activate an element through accessibility (no mouse); --headed for a "
         "real, guarded mouse click",
         target_arg(), *window_args(),
         Arg("headed", bool, help="real mouse click at the element's center"),
         *verify_args())
def cmd_click(args):
    ctrl, el, status, hwnd = target(args)
    preflight(ctrl, el)
    before = verify.observe(ctrl, hwnd)
    if args.headed:
        how = "mouse:" + inputs.physical_click(ctrl, hwnd)
    else:
        how = _semantic_click(ctrl)
        if not how:
            raise WadError("NOT_CLICKABLE", f"{el['ref']} supports no click action",
                           "retry with --headed to use a real mouse click")
    changes, verified = verify.after_action(args, ctrl, hwnd, before)
    state.trace("click", {"target": el["ref"], "role": el["role"], "name": el["name"],
                          "via": how, "changes": changes})
    return (ok(ref=el["ref"], via=how, resolved=status, changes=changes, verified=verified,
               **where(el, hwnd)),
            act_line("clicked", el, how, changes, verified))


def _mouse_command(name, help, button, count):
    @command(name, help, target_arg(), *window_args(), *verify_args(), group="mouse")
    def fn(args):
        ctrl, el, status, hwnd = target(args)
        preflight(ctrl, el)
        before = verify.observe(ctrl, hwnd)
        landed = inputs.physical_click(ctrl, hwnd, button, count)
        changes, verified = verify.after_action(args, ctrl, hwnd, before)
        state.trace(name, {"target": el["ref"], "name": el["name"], "landed": landed,
                           "changes": changes})
        return (ok(ref=el["ref"], landed=landed, changes=changes, verified=verified,
                   **where(el, hwnd)),
                act_line(name.replace("-", " ") + "ed", el, "mouse", changes, verified))
    return fn


_mouse_command("right-click", "real right click on an element (opens context menus; then "
               "`snapshot --popup`)", "right", 1)
_mouse_command("double-click", "real double click on an element", "left", 2)


@command("hover", "move the real mouse pointer over an element (tooltips, hover menus)",
         target_arg(), *window_args(), group="mouse")
def cmd_hover(args):
    ctrl, el, _, hwnd = target(args)
    inputs.ensure_visible(ctrl)
    x, y = inputs.center(ctrl)
    inputs.guard(hwnd)
    auto.MoveTo(x, y, moveSpeed=1, waitTime=0.3)
    return ok(ref=el["ref"], x=x, y=y), f"pointer over {el['ref']} {el['role']} {el['name']!r}"


@command("drag", "drag with the real mouse from one element to another (or between points)",
         Arg("source", positional=True, optional=True, help="ref or selector to drag from"),
         Arg("dest", positional=True, optional=True, help="ref or selector to drop on"),
         Arg("from-xy", help="X,Y screen pixels instead of a source element"),
         Arg("to-xy", help="X,Y screen pixels instead of a destination element"),
         *window_args(), group="mouse")
def cmd_drag(args):
    def point(ref, xy, label):
        if xy:
            try:
                x, y = (int(float(v)) for v in xy.split(","))
            except ValueError:
                raise WadError("USAGE", f"--{label}-xy takes X,Y") from None
            return x, y, None, None
        if not ref:
            raise WadError("USAGE", f"give a {label} element or --{label}-xy")
        ctrl, el, _, hwnd = uia.resolve_target(ref, args.window, args.hwnd)
        inputs.ensure_visible(ctrl)
        x, y = inputs.center(ctrl)
        return x, y, ctrl, hwnd
    x1, y1, c1, h1 = point(args.source, args.from_xy, "from")
    x2, y2, _, h2 = point(args.dest, args.to_xy, "to")
    inputs.guard(h1 or h2)
    if c1 is not None:
        inputs.hit_check(c1, x1, y1)
    before = verify.observe(c1, h1) if c1 is not None else None
    auto.DragDrop(x1, y1, x2, y2, moveSpeed=1, waitTime=0.3)
    changes = verify.settle(c1, h1, before) if before else []
    state.trace("drag", {"from": [x1, y1], "to": [x2, y2], "changes": changes})
    text = f"dragged ({x1},{y1}) -> ({x2},{y2})"
    if changes:
        text += "\n  changed: " + "; ".join(changes[:6])
    return ok(start=[x1, y1], end=[x2, y2], changes=changes), text


def _read_back(ctrl):
    return uia.content_of(ctrl)


@command("type", "put text into a field: value pattern first, verified, and automatically "
         "retyped as real Unicode keystrokes if the app ignored it",
         target_arg(), Arg("text", positional=True, help="the text (any language)"),
         *window_args(),
         Arg("append", bool, help="add to the end instead of replacing"),
         Arg("keys", bool, help="skip the value pattern and type keystrokes directly"),
         Arg("submit", bool, help="press Enter afterwards"),
         *verify_args())
def cmd_type(args):
    ctrl, el, status, hwnd = target(args)
    preflight(ctrl, el)
    before = verify.observe(ctrl, hwnd)
    old = _read_back(ctrl) or ""
    want = (old + args.text) if args.append else args.text
    window_class = ""
    try:
        window_class = auto.ControlFromHandle(hwnd).ClassName or ""
    except Exception:
        pass
    vp = ctrl.GetPattern(auto.PatternId.ValuePattern)
    use_value = (vp is not None and not args.keys and window_class not in VALUE_LIARS)
    try:
        use_value = use_value and not vp.IsReadOnly
    except Exception:
        use_value = False
    how, notes = None, []
    try:
        secret = bool(ctrl.IsPassword)
    except Exception:
        secret = False
    if use_value:
        try:
            vp.SetValue(want)
            how = "value"
        except Exception as e:
            notes.append(f"value pattern refused ({e.__class__.__name__}); typed instead")
        if how and not args.no_verify:
            have = verify.read_until(lambda: _read_back(ctrl),
                                     lambda v: verify.same_text(want, v), timeout=0.6)
            if not verify.same_text(want, have):
                notes.append("value pattern did not stick; typed instead")
                how = None
    if how is None:
        inputs.type_text(ctrl, hwnd, args.text, replace=not args.append)
        how = "keys"
    if window_class in VALUE_LIARS and not args.keys:
        notes.append("Excel ignores value-pattern writes; typed. For cells prefer `excel-write`")
    if args.submit:
        inputs.guard(hwnd)
        inputs.press_chord([auto.Keys.VK_RETURN])
    verified = None
    if not args.no_verify:
        def landed(v):
            return verify.same_text(want, v) or (
                how == "keys" and str(args.text).strip() in str(v or ""))
        have = verify.read_until(lambda: _read_back(ctrl), landed,
                                 timeout=2.0 + len(args.text) * 0.01)
        if have is not None and not args.submit:
            verified = landed(have)
            if not verified:
                raise WadError("VERIFY_FAILED",
                               f"after typing, {el['ref']} holds {verify._short(have)!r}, "
                               f"not {verify._short(want)!r}",
                               "an input mask, autocomplete or a read-only field changed it - "
                               "`get` it, then retry or type fewer characters")
    changes, _ = verify.after_action(args, ctrl, hwnd, before)
    state.trace("type", {"target": el["ref"], "chars": len(args.text), "via": how,
                         "verified": verified})
    text = f"typed {len(args.text)} chars into {el['ref']} {el['role']} {el['name']!r} via {how}"
    text += "" if verified is None else (" - verified" if verified else "")
    if notes:
        text += "\n  note: " + "; ".join(notes)
    return ok(ref=el["ref"], via=how, verified=verified, notes=notes, changes=changes,
              secret=secret, **where(el, hwnd)), text


@command("clear", "empty a text field (verified)", target_arg(), *window_args())
def cmd_clear(args):
    ctrl, el, _, hwnd = target(args)
    preflight(ctrl, el)
    vp = ctrl.GetPattern(auto.PatternId.ValuePattern)
    if vp and not vp.IsReadOnly:
        vp.SetValue("")
        time.sleep(0.1)
        if verify.same_text("", _read_back(ctrl)):
            return ok(ref=el["ref"], via="value", **where(el, hwnd)), f"cleared {el['ref']} via value"
    inputs.guard(hwnd)
    ctrl.SetFocus()
    inputs.press_chord([auto.Keys.VK_CONTROL, ord("A")])
    inputs.press_chord([auto.Keys.VK_DELETE])
    left = verify.read_until(lambda: _read_back(ctrl), lambda v: v in (None, ""))
    if left not in (None, ""):
        raise WadError("VERIFY_FAILED", f"{el['ref']} still holds {verify._short(left)!r}")
    return ok(ref=el["ref"], via="keys", **where(el, hwnd)), f"cleared {el['ref']} via keys"


def _set_toggle(args, want):
    ctrl, el, _, hwnd = target(args)
    preflight(ctrl, el)
    tp = ctrl.GetPattern(auto.PatternId.TogglePattern)
    if not tp:
        raise WadError("NOT_TOGGLABLE", f"{el['ref']} {el['role']} has no on/off state")
    for _ in range(3):                                  # on -> indeterminate -> off cycles
        if tp.ToggleState == want:
            break
        tp.Toggle()
        time.sleep(0.05)
    if tp.ToggleState != want:
        raise WadError("VERIFY_FAILED", f"{el['ref']} would not switch "
                       f"{'on' if want else 'off'}", "try `click --headed`")
    word = "checked" if want else "unchecked"
    state.trace(word, {"target": el["ref"], "name": el["name"]})
    return (ok(ref=el["ref"], state="on" if want else "off", **where(el, hwnd)),
            f"{word} {el['ref']} {el['name']!r}")


@command("check", "switch a checkbox / toggle ON (no-op if already on; verified)",
         target_arg(), *window_args())
def cmd_check(args):
    return _set_toggle(args, 1)


@command("uncheck", "switch a checkbox / toggle OFF (verified)", target_arg(), *window_args())
def cmd_uncheck(args):
    return _set_toggle(args, 0)


def _set_expanded(args, want):
    ctrl, el, _, hwnd = target(args)
    ec = ctrl.GetPattern(auto.PatternId.ExpandCollapsePattern)
    if not ec:
        raise WadError("NOT_EXPANDABLE", f"{el['ref']} {el['role']} cannot expand/collapse",
                       "for a menu use `click`; for a sub-menu try its accelerator key")
    before = verify.observe(ctrl, hwnd)
    (ec.Expand if want else ec.Collapse)()
    deadline = time.time() + 2
    while time.time() < deadline and (ec.ExpandCollapseState in (1, 2)) != want:
        time.sleep(0.1)
    changes = verify.settle(ctrl, hwnd, before, wait=0.3)
    word = "expanded" if want else "collapsed"
    return (ok(ref=el["ref"], state=word, changes=changes, **where(el, hwnd)),
            act_line(word, el, word, changes, True))


@command("expand", "open a combo box, tree item, menu or ribbon drop-down", target_arg(),
         *window_args())
def cmd_expand(args):
    return _set_expanded(args, True)


@command("collapse", "close a combo box, tree item or drop-down", target_arg(), *window_args())
def cmd_collapse(args):
    return _set_expanded(args, False)


@command("select", "pick an option by its text in a combo box, list, tab strip or tree "
         "(verified)",
         target_arg("the combo box / list (ref or selector)"),
         Arg("option", positional=True, help="the option's visible text (exact, else substring)"),
         *window_args())
def cmd_select(args):
    ctrl, el, _, hwnd = target(args)
    preflight(ctrl, el)
    ec = ctrl.GetPattern(auto.PatternId.ExpandCollapsePattern)
    opened = False
    if ec:
        try:
            if ec.ExpandCollapseState == 0:
                ec.Expand()
                opened = True
                time.sleep(0.3)
        except Exception:
            pass
    want = args.option.lower()
    items = []
    roots = [ctrl]
    try:                              # Win32 combo boxes drop their list in a new window
        pid = ctrl.ProcessId
        roots += [t for t in auto.GetRootControl().GetChildren()
                  if t.ProcessId == pid and t.ClassName in ("ComboLBox", "#32768", "Popup")]
    except Exception:
        pass
    for root in roots:
        for c, _d in auto.WalkControl(root, includeTop=False, maxDepth=6):
            try:
                if c.GetPattern(auto.PatternId.SelectionItemPattern) and c.Name:
                    items.append(c)
            except Exception:
                continue
    exact = [c for c in items if c.Name.lower() == want]
    part = [c for c in items if want in c.Name.lower()]
    pick = (exact or part or [None])[0]
    if pick is None:
        if opened:
            ec.Collapse()
        names = ", ".join(repr(c.Name) for c in items[:15])
        raise WadError("OPTION_NOT_FOUND", f"no option like {args.option!r}",
                       f"options: {names}" if names else "the list exposes no options")
    sip = pick.GetPattern(auto.PatternId.SelectionItemPattern)
    sip.Select()
    time.sleep(0.2)
    if opened:
        try:
            ec.Collapse()
        except Exception:
            pass
    shown = uia.value_of(ctrl)
    ok_sel = bool(sip.IsSelected) if uia.alive(pick) else False
    if not ok_sel and not (shown and shown.lower() == pick.Name.lower()):
        raise WadError("VERIFY_FAILED", f"{pick.Name!r} was not selected",
                       "try `click --headed` on the option")
    state.trace("select", {"target": el["ref"], "option": pick.Name})
    return (ok(ref=el["ref"], option=pick.Name, value=shown, **where(el, hwnd)),
            f"selected {pick.Name!r} in {el['ref']} {el['role']} {el['name']!r}")


@command("focus", "give keyboard focus to an element, or bring a window to the front",
         Arg("target", positional=True, optional=True, help="ref or selector (omit for window)"),
         *window_args(), group="windows")
def cmd_focus(args):
    if args.target:
        ctrl, el, _, hwnd = target(args)
        inputs.guard(hwnd)
        ctrl.SetFocus()
        return ok(ref=el["ref"]), f"focused {el['ref']} {el['role']} {el['name']!r}"
    w = uia.find_window(args.window, args.hwnd)
    if not win32.bring_front(w.NativeWindowHandle):
        w.SetActive()
    return ok(hwnd=w.NativeWindowHandle), f"front: {w.Name!r}"


def _scroll_owner(ctrl):
    cur = ctrl
    for _ in range(12):
        if cur is None:
            break
        sp = cur.GetPattern(auto.PatternId.ScrollPattern)
        if sp:
            return cur, sp
        cur = cur.GetParentControl()
    return None, None


@command("scroll", "scroll a list/document/page (accessibility scroll; --headed uses the "
         "mouse wheel)",
         target_arg("the scrollable element or anything inside it"),
         Arg("direction", choices=["down", "up", "left", "right"], default="down"),
         Arg("amount", int, default=3, help="steps (small increments or wheel notches)"),
         Arg("page", bool, help="scroll by pages instead of small steps"),
         Arg("headed", bool, help="use the real mouse wheel"),
         *window_args(), group="mouse")
def cmd_scroll(args):
    ctrl, el, _, hwnd = target(args)
    owner, sp = (None, None) if args.headed else _scroll_owner(ctrl)
    if sp is not None:
        # ScrollAmount: 0 LargeDecrement, 1 SmallDecrement, 2 NoAmount, 3 LargeIncrement,
        # 4 SmallIncrement.
        inc = 3 if args.page else 4
        dec = 0 if args.page else 1
        vertical = args.direction in ("down", "up")
        amt = inc if args.direction in ("down", "right") else dec
        pos = lambda: (sp.VerticalScrollPercent if vertical else sp.HorizontalScrollPercent)
        start = pos()
        for _ in range(max(1, args.amount)):
            if vertical:
                sp.Scroll(2, amt)
            else:
                sp.Scroll(amt, 2)
            time.sleep(0.05)
        end = pos()
        moved = start != end
        text = (f"scrolled {args.direction} {args.amount} in {uia.role_of(owner)} "
                f"{owner.Name!r}: {start:.0f}% -> {end:.0f}%")
        if not moved:
            text += " (did not move - already at the end?)"
        return ok(via="scroll", start=start, end=end, moved=moved), text
    x, y = inputs.center(ctrl)
    inputs.guard(hwnd)
    inputs.hit_check(ctrl, x, y)
    auto.MoveTo(x, y, moveSpeed=1, waitTime=0.05)
    if args.direction in ("down", "up"):
        fn = auto.WheelDown if args.direction == "down" else auto.WheelUp
        fn(wheelTimes=max(1, args.amount) * (3 if args.page else 1), interval=0.05, waitTime=0.2)
    else:
        key = auto.Keys.VK_RIGHT if args.direction == "right" else auto.Keys.VK_LEFT
        for _ in range(max(1, args.amount)):
            inputs.press_chord([auto.Keys.VK_SHIFT, key])
    return ok(via="wheel"), f"wheel-scrolled {args.direction} {args.amount} at ({x},{y})"


@command("scroll-to", "scroll an element into view", target_arg(), *window_args(), group="mouse")
def cmd_scroll_to(args):
    ctrl, el, _, _ = target(args)
    sip = ctrl.GetPattern(auto.PatternId.ScrollItemPattern)
    if not sip:
        raise WadError("NOT_SCROLLABLE", f"{el['ref']} cannot be scrolled into view",
                       "use `scroll` on its container")
    sip.ScrollIntoView()
    time.sleep(0.2)
    return ok(ref=el["ref"], offscreen=bool(ctrl.IsOffscreen)), f"scrolled {el['ref']} into view"


# ---------------------------------------------------------------------------
# Keyboard
# ---------------------------------------------------------------------------

@command("press", "key chord or sequence to the app, guarded: 'ctrl+s', 'alt+f4', "
         "'alt h o r' (a sequence, e.g. Office key tips)",
         Arg("combo", positional=True, help="chords joined by +, sequence separated by spaces"),
         *window_args(),
         Arg("interval", float, default=0.4, help="pause between chords of a sequence"),
         group="keyboard")
def cmd_press(args):
    chords = inputs.parse_keys(args.combo)
    w = uia.find_window(args.window, args.hwnd)
    inputs.guard(w.NativeWindowHandle)
    for i, codes in enumerate(chords):
        inputs.guard(w.NativeWindowHandle)
        inputs.press_chord(codes)
        if i < len(chords) - 1:
            time.sleep(args.interval)
    state.trace("press", {"combo": args.combo, "window": w.Name})
    return ok(combo=args.combo), f"pressed {args.combo} in {w.Name!r}"


@command("input-lang", "show or switch the keyboard layout of ONE window (en, ar, or a KLID); "
         "Office key tips need en",
         Arg("lang", positional=True, optional=True, help="en | ar | fr | de | 00000409 ..."),
         *window_args(), group="keyboard")
def cmd_input_lang(args):
    w = uia.find_window(args.window, args.hwnd)
    if args.lang:
        if not win32.set_input_language(w.NativeWindowHandle, args.lang):
            raise WadError("LAYOUT_FAILED", f"{w.Name!r} did not switch to {args.lang}",
                           "is that keyboard layout installed? (Settings > Language)")
    lang = win32.input_language(w.NativeWindowHandle)
    return ok(window=w.Name, langid=f"{lang:04x}"), f"{w.Name!r} input language {lang:04x}"


# ---------------------------------------------------------------------------
# Windows
# ---------------------------------------------------------------------------

@command("window", "minimize / maximize / restore / move / resize / front a window",
         Arg("action", positional=True,
             choices=["minimize", "maximize", "restore", "move", "resize", "front"]),
         *window_args(), Arg("x", int), Arg("y", int), Arg("width", int), Arg("height", int),
         group="windows")
def cmd_window(args):
    w = uia.find_window(args.window, args.hwnd)
    if args.action == "front":
        win32.bring_front(w.NativeWindowHandle)
    elif args.action in ("minimize", "maximize", "restore"):
        wp = w.GetPattern(auto.PatternId.WindowPattern)
        if not wp:
            raise WadError("NOT_SUPPORTED", f"{w.Name!r} has no window state control")
        wp.SetWindowVisualState({"restore": 0, "maximize": 1, "minimize": 2}[args.action])
    else:
        tp = w.GetPattern(auto.PatternId.TransformPattern)
        if not tp:
            raise WadError("NOT_SUPPORTED", f"{w.Name!r} cannot be moved or resized")
        if args.action == "move":
            if args.x is None or args.y is None:
                raise WadError("USAGE", "move needs --x and --y")
            tp.Move(args.x, args.y)
        else:
            if not args.width or not args.height:
                raise WadError("USAGE", "resize needs --width and --height")
            tp.Resize(args.width, args.height)
    time.sleep(0.2)
    return ok(window=w.Name, rect=uia.rect_of(w)), f"{args.action}: {w.Name!r} {uia.rect_of(w)}"


@command("close", "close a window politely (its own close action; the app may ask to save)",
         *window_args(), group="windows")
def cmd_close(args):
    w = uia.find_window(args.window, args.hwnd)
    wp = w.GetPattern(auto.PatternId.WindowPattern)
    name = w.Name
    if not wp:
        raise WadError("NOT_CLOSABLE", f"{name!r} has no window close action",
                       "use `press alt+f4 --window ...`")
    wp.Close()
    state.trace("close", {"window": name})
    return ok(window=name), f"closed {name!r}"


@command("clipboard", "read, set or clear the clipboard text",
         Arg("action", positional=True, choices=["get", "set", "clear"]),
         Arg("text", positional=True, optional=True), group="keyboard")
def cmd_clipboard(args):
    if args.action == "get":
        text = auto.GetClipboardText()
        return ok(text=text), text
    auto.SetClipboardText((args.text or "") if args.action == "set" else "")
    return ok(), "clipboard set" if args.action == "set" else "clipboard cleared"
