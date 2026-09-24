"""Seeing an app: windows, the accessibility tree, refs and selectors.

A ref (e12) names one element of the last snapshot; it is cheap and exact, but it only
lives until the next snapshot. A selector (role=Button name=Save) is searched live every
time, so it survives restarts and is what recorded workflows use."""
import difflib
import re
import shlex
import time
from datetime import datetime

import uiautomation as auto

from . import state, win32
from .registry import WadError

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
    "scroll": auto.PatternId.ScrollPattern,
}

REF_RE = re.compile(r"^@?(?:(s\w+):)?(e\d+)$")
MAX_ELEMENTS = 4000


def role_of(ctrl):
    return (ctrl.ControlTypeName or "").replace("Control", "")


def alive(ctrl):
    """False once the element is gone from the app (reading it raises)."""
    try:
        ctrl.ControlTypeName
        _ = ctrl.Name
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Windows
# ---------------------------------------------------------------------------

def top_windows(include_unnamed=False):
    out = []
    for w in auto.GetRootControl().GetChildren():
        try:
            if (not w.Name and not include_unnamed) or w.IsOffscreen:
                continue
        except Exception:
            continue
        out.append(w)
    return out


def window_of(ctrl):
    """The top-level window an element lives in."""
    root = auto.GetRootControl()
    root_rid = list(root.GetRuntimeId() or [])
    cur = ctrl
    while True:
        parent = cur.GetParentControl()
        if parent is None or list(parent.GetRuntimeId() or []) == root_rid:
            return cur
        cur = parent


def responsive(w):
    """Refuse up front to drive a window Windows itself reports as not responding."""
    if w is not None and win32.is_hung(w.NativeWindowHandle):
        raise WadError("APP_HUNG", f"{w.Name!r} is not responding",
                       "wait for it to recover (`wait --window ... --timeout 30`), or ask the "
                       "person - never click into a hung app")
    return w


def find_window(title=None, hwnd=None):
    return responsive(_find_window(title, hwnd))


def _find_window(title=None, hwnd=None):
    if hwnd:
        w = auto.ControlFromHandle(int(hwnd))
        if not w:
            raise WadError("WINDOW_NOT_FOUND", f"no window with handle {hwnd}",
                           "run `wad windows` to list open windows")
        return w
    if not title:
        # Titles drift (Notepad renames an unsaved tab after its first line), so with no
        # window named, the window of the last snapshot is used by handle.
        try:
            snap = state.load_snapshot()
        except WadError:
            raise WadError("USAGE", "name a window with --window TITLE or --hwnd N") from None
        w = auto.ControlFromHandle(snap["hwnd"])
        if not w or not alive(w):
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
                       "run `wad windows`, or `wad launch <app>` first")
    if len(hits) > 1 and not exact:
        names = ", ".join(f"{w.Name!r} (hwnd {w.NativeWindowHandle})" for w in hits[:6])
        raise WadError("AMBIGUOUS_WINDOW", f"{len(hits)} windows match {title!r}: {names}",
                       "use the full title or --hwnd")
    return hits[0]


def new_window(title, exclude=()):
    """The window titled like `title` that is not in `exclude`, or None (one look)."""
    low = title.lower()
    fresh = [w for w in top_windows()
             if low in w.Name.lower() and w.NativeWindowHandle not in exclude]
    exact = [w for w in fresh if w.Name.lower() == low]
    return (exact or fresh or [None])[0]


def wait_window(title, timeout, exclude=()):
    """Wait for a window titled like `title`; handles in `exclude` (windows that were
    already open) are skipped, so a launch never grabs the user's own copy."""
    deadline = time.time() + timeout
    low = title.lower()
    while True:
        fresh = [w for w in top_windows()
                 if low in w.Name.lower() and w.NativeWindowHandle not in exclude]
        if fresh:
            exact = [w for w in fresh if w.Name.lower() == low]
            return (exact or fresh)[0]
        if time.time() > deadline:
            raise WadError("TIMEOUT", f"no new window titled like {title!r} within {timeout}s",
                           "pass --title with the real window title (see `wad windows`)")
        time.sleep(0.3)


def popup_of(hwnd):
    """The topmost menu or dialog belonging to the same app as `hwnd` (context menus and
    many dialogs are separate top-level windows, often without a title)."""
    main = auto.ControlFromHandle(hwnd)
    pid = main.ProcessId if main else None
    same_app = []
    for w in auto.GetRootControl().GetChildren():      # z-order: topmost first
        try:
            if w.NativeWindowHandle == hwnd or w.IsOffscreen:
                continue
            if w.ClassName == "#32768":                  # a menu is always what is open
                return w
            if pid and w.ProcessId == pid:
                same_app.append(w)
        except Exception:
            continue
    # dialogs before tool windows and IME helpers of the same process
    for w in same_app:
        if w.ClassName == "#32770" or w.ControlTypeName == "WindowControl" and w.Name:
            return w
    if same_app:
        return same_app[0]
    try:                  # modern apps host dialogs inside the main window
        for c in main.GetChildren():
            if c.ControlTypeName == "WindowControl" and c.NativeWindowHandle:
                return c
    except Exception:
        pass
    raise WadError("NO_POPUP", "no open menu or dialog of that app",
                   "open it first (right-click, a menu item), then snapshot --popup")


# ---------------------------------------------------------------------------
# Elements
# ---------------------------------------------------------------------------

def supported(ctrl):
    out = []
    for k, pid in PATTERNS.items():
        try:
            if ctrl.GetPattern(pid):
                out.append(k)
        except Exception:
            pass
    return out


def rect_of(ctrl):
    r = ctrl.BoundingRectangle
    return [r.left, r.top, r.right, r.bottom]


def is_interactive(ctrl, pats):
    return ctrl.ControlTypeName in INTERACTIVE_TYPES or bool(
        set(pats) & {"invoke", "toggle", "select", "expand"})


# MSAA state bits, for controls that only speak the old accessibility API
STATE_CHECKED, STATE_EXPANDED, STATE_COLLAPSED = 0x10, 0x200, 0x400


def legacy_of(ctrl):
    try:
        return ctrl.GetPattern(auto.PatternId.LegacyIAccessiblePattern)
    except Exception:
        return None


def value_of(ctrl):
    vp = ctrl.GetPattern(auto.PatternId.ValuePattern)
    if vp:
        try:
            return vp.Value
        except Exception:
            return None
    # A WinForms control given an AccessibleName (also VB6, Delphi, old MFC) answers only
    # through MSAA: no Value pattern, but its accValue is the shown text or selection.
    lp = legacy_of(ctrl)
    if lp:
        try:
            return lp.Value or None
        except Exception:
            return None
    return None


def text_of(ctrl):
    tp = ctrl.GetPattern(auto.PatternId.TextPattern)
    if tp:
        try:
            return tp.DocumentRange.GetText(-1)
        except Exception:
            return None
    return None


def content_of(ctrl):
    """What a field holds: its value, else its document text."""
    v = value_of(ctrl)
    if v not in (None, ""):
        return v
    t = text_of(ctrl)
    return t if t is not None else v


def states_of(ctrl):
    out = {}
    tg = ctrl.GetPattern(auto.PatternId.TogglePattern)
    if tg:
        try:
            out["toggle"] = {0: "off", 1: "on", 2: "indeterminate"}.get(tg.ToggleState)
        except Exception:
            pass
    ec = ctrl.GetPattern(auto.PatternId.ExpandCollapsePattern)
    if ec:
        try:
            out["expanded"] = {0: False, 1: True, 2: True, 3: None}.get(ec.ExpandCollapseState)
        except Exception:
            pass
    si = ctrl.GetPattern(auto.PatternId.SelectionItemPattern)
    if si:
        try:
            out["selected"] = bool(si.IsSelected)
        except Exception:
            pass
    if "expanded" not in out or "toggle" not in out:
        lp = legacy_of(ctrl)
        try:
            st = lp.State if lp else 0
        except Exception:
            st = 0
        if "expanded" not in out and st & (STATE_EXPANDED | STATE_COLLAPSED):
            out["expanded"] = bool(st & STATE_EXPANDED)
        if "toggle" not in out and role_of(ctrl) in ("CheckBox", "RadioButton") and lp:
            out["toggle"] = "on" if st & STATE_CHECKED else "off"
    return out


def describe(ctrl):
    pats = supported(ctrl)
    return {
        "role": role_of(ctrl),
        "name": ctrl.Name or "",
        "aid": ctrl.AutomationId or "",
        "class": ctrl.ClassName or "",
        "enabled": bool(ctrl.IsEnabled),
        "offscreen": bool(ctrl.IsOffscreen),
        "rect": rect_of(ctrl),
        "actions": pats,
        "value": value_of(ctrl) if "value" in pats else None,
        "interactive": is_interactive(ctrl, pats),
    }


def selector_for(el):
    """A durable selector for an element: its AutomationId when it has one (app authors
    keep those stable), else role and name."""
    def q(v):
        return shlex.quote(v) if re.search(r"[\s'\"]", v) or not v else v
    if el.get("aid") and not el["aid"].isdigit():
        return f"role={el['role']} aid={q(el['aid'])}"
    if el.get("name"):
        return f"role={el['role']} name={q(el['name'])}"
    return f"role={el['role']}"


# ---------------------------------------------------------------------------
# Snapshots and refs
# ---------------------------------------------------------------------------

def step_key(ctrl, index):
    return [ctrl.ControlTypeName, ctrl.AutomationId or "", ctrl.Name or "", index]


def walk(root, root_path, max_depth, counter, elements, parent_ref=None, depth=0,
         focused_rid=None, limit=MAX_ELEMENTS):
    """Depth-first walk that records every element with the path needed to find it again."""
    node = describe(root)
    ref = f"e{counter[0]}"
    counter[0] += 1
    rid = list(root.GetRuntimeId() or [])
    node.update(ref=ref, depth=depth, parent=parent_ref, path=root_path, rid=rid)
    if focused_rid and rid == focused_rid:
        node["focused"] = True
    elements.append(node)
    children = []
    node["truncated"] = False
    if (max_depth is None or depth < max_depth) and len(elements) < limit:
        try:
            children = root.GetChildren()
        except Exception:
            children = []
    else:
        try:
            node["truncated"] = bool(root.GetFirstChildControl())
        except Exception:
            pass
    node["child_count"] = len(children)
    for i, child in enumerate(children):
        if len(elements) >= limit:
            node["truncated"] = True
            break
        # Apps like Excel destroy and rebuild elements while they are being read; one
        # that vanished mid-walk is skipped rather than failing the snapshot.
        mark = len(elements)
        try:
            walk(child, root_path + [step_key(child, i)], max_depth, counter, elements, ref,
                 depth + 1, focused_rid, limit)
        except Exception:
            del elements[mark:]
    return elements


def take_snapshot(root, hwnd, wname, base_path, depth, limit):
    try:
        focused_rid = list(auto.GetFocusedControl().GetRuntimeId() or [])
    except Exception:
        focused_rid = None
    elements = walk(root, base_path, depth, [1], [], focused_rid=focused_rid, limit=limit)
    sid = "s" + datetime.now().strftime("%H%M%S%f")[:9]
    snap = {"id": sid, "hwnd": hwnd, "window": wname, "taken": state.now_iso(),
            "elements": elements, "limited": len(elements) >= limit}
    state.save_snapshot(snap)
    return snap


def lookup_ref(ref):
    snap = state.load_snapshot()
    m = REF_RE.match(ref.strip())
    if not m:
        raise WadError("USAGE", f"{ref!r} is not a ref")
    sid, ref = m.groups()
    if sid and sid != snap["id"]:
        raise WadError("STALE_SNAPSHOT", f"ref belongs to snapshot {sid}, latest is {snap['id']}",
                       "take a new snapshot and use its refs")
    for el in snap["elements"]:
        if el["ref"] == ref:
            return snap, el
    raise WadError("REF_NOT_FOUND", f"@{ref} is not in snapshot {snap['id']}",
                   "take a new snapshot and use its refs")


def resolve_ref(ref):
    """Re-find a ref's live element by walking its recorded path from the window.

    The runtime id is checked at the end: a match means it is the very element that was
    observed; a mismatch with the same role and name means the app rebuilt that element
    (accepted, reported as re-resolved); anything else is refused."""
    snap, el = lookup_ref(ref)
    ctrl = auto.ControlFromHandle(snap["hwnd"])
    if not ctrl or not alive(ctrl):
        raise WadError("WINDOW_GONE", f"window {snap['window']!r} is closed",
                       "launch the app again and take a new snapshot")
    responsive(ctrl)
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
    if status != "exact" and (role_of(ctrl) != el["role"] or (ctrl.Name or "") != el["name"]):
        raise WadError("ELEMENT_CHANGED", f"@{el['ref']} now points at a different element",
                       "take a new snapshot")
    return ctrl, dict(el), status, snap["hwnd"]


# ---------------------------------------------------------------------------
# Selectors
# ---------------------------------------------------------------------------

SELECTOR_KEYS = {"role", "name", "aid", "id", "class", "value", "nth", "text"}
TERM_RE = re.compile(r"^(\w+)(~=|=|:)(.*)$", re.S)


def parse_selector(text):
    """`role=Button name="Save as"`, `name~=sav` (contains), `nth=2`, and `>>` to search
    inside the previous match: `name=Ribbon >> role=Button name=Bold`."""
    try:
        words = shlex.split(text, posix=True)
    except ValueError as e:
        raise WadError("BAD_SELECTOR", f"cannot read selector {text!r}: {e}") from None
    steps, cur = [], []
    for w in words + [">>"]:
        if w == ">>":
            if not cur:
                raise WadError("BAD_SELECTOR", f"empty step in selector {text!r}")
            steps.append(cur)
            cur = []
            continue
        m = TERM_RE.match(w)
        if not m or m.group(1).lower() not in SELECTOR_KEYS:
            raise WadError("BAD_SELECTOR", f"bad selector term {w!r}",
                           "terms: role= name= name~= aid= class= value= nth=, steps joined by >>")
        key, op, val = m.group(1).lower(), m.group(2), m.group(3)
        key = {"id": "aid", "text": "name"}.get(key, key)
        if key == "nth":
            if not val.isdigit() or int(val) < 1:
                raise WadError("BAD_SELECTOR", "nth= takes a number from 1")
            cur.append(("nth", "=", int(val)))
        else:
            cur.append((key, "~=" if op == "~=" else "=", val))
    return steps


def _prop(ctrl, key):
    if key == "role":
        return role_of(ctrl)
    if key == "name":
        return ctrl.Name or ""
    if key == "aid":
        return ctrl.AutomationId or ""
    if key == "class":
        return ctrl.ClassName or ""
    if key == "value":
        return value_of(ctrl) or ""
    return ""


def matches(ctrl, terms):
    for key, op, want in terms:
        if key == "nth":
            continue
        have = _prop(ctrl, key).lower()
        want = want.lower()
        if key == "role":
            want = want.replace("control", "")
        if (op == "=" and have != want) or (op == "~=" and want not in have):
            return False
    return True


def search(root, terms, depth=40):
    hits = []
    try:
        for ctrl, _d in auto.WalkControl(root, includeTop=False, maxDepth=depth):
            try:
                if matches(ctrl, terms):
                    hits.append(ctrl)
            except Exception:
                continue
    except Exception:
        pass            # the app rebuilt part of the tree mid-walk; take what was found
    return hits


def find_selector(root, text, timeout=0.0):
    steps = parse_selector(text)
    deadline = time.time() + timeout
    while True:
        try:
            return _find_steps(root, steps, text)
        except WadError as e:
            if e.code != "ELEMENT_NOT_FOUND" or time.time() >= deadline:
                raise
        time.sleep(0.3)


def _find_steps(root, steps, text):
    cur = root
    for terms in steps:
        hits = search(cur, terms)
        nth = next((v for k, _, v in terms if k == "nth"), None)
        if not hits:
            raise WadError("ELEMENT_NOT_FOUND", f"nothing matches {text!r}",
                           "check the window, or look with `wad snapshot -i` / `wad find`")
        if nth is not None:
            if nth > len(hits):
                raise WadError("ELEMENT_NOT_FOUND", f"{text!r}: nth={nth} but only {len(hits)} match")
            cur = hits[nth - 1]
        elif len(hits) > 1:
            listing = "; ".join(f"{role_of(h)} {h.Name!r}" + (f" aid={h.AutomationId}" if h.AutomationId else "")
                                for h in hits[:5])
            raise WadError("AMBIGUOUS_TARGET", f"{len(hits)} elements match {text!r}: {listing}",
                           "add terms (aid=, role=), nth=N, or >> to scope inside a parent")
        else:
            cur = hits[0]
    return cur


def is_ref(target):
    return bool(REF_RE.match(target.strip()))


def resolve_target(target, window=None, hwnd=None, timeout=0.0, heal=False):
    """(control, description, how-it-was-found, window handle) for a ref or a selector."""
    if not target:
        raise WadError("USAGE", "name a target: a ref (e12) or a selector (role=Button name=Save)")
    if is_ref(target):
        return resolve_ref(target)
    win = find_window(window, hwnd)
    how = "selector"
    try:
        ctrl = find_selector(win, target, timeout)
    except WadError as e:
        if e.code != "ELEMENT_NOT_FOUND":
            raise
        # Menus, drop-down lists and many dialogs are separate top-level windows of the
        # same app, so a recorded step may name the main window but live in a popup.
        ctrl = next((c for c in _in_app_popups(win, target)), None)
        if ctrl is None and heal:
            ctrl, target, score = heal_selector(win, target)
            how = f"healed ({score:.0%})"
        elif ctrl is None:
            raise
    el = describe(ctrl)
    el["ref"] = target
    return ctrl, el, how, win.NativeWindowHandle


def _in_app_popups(win, text):
    try:
        pid, main = win.ProcessId, win.NativeWindowHandle
        tops = [t for t in auto.GetRootControl().GetChildren()
                if t.ProcessId == pid and t.NativeWindowHandle != main]
    except Exception:
        return
    for t in tops:
        try:
            yield find_selector(t, text)
        except WadError:
            continue


HEAL_MIN, HEAL_MARGIN = 0.72, 0.08


def heal_selector(root, text):
    """A step recorded against an older version of the app: the button was renamed
    ("Save" -> "Save file"), its AutomationId changed, or it moved. Find the one element
    that is clearly the same thing - same role, and the same AutomationId or a very
    similar name - or give up. Never guess between close candidates."""
    steps = parse_selector(text)
    scope = _find_steps(root, steps[:-1], text) if len(steps) > 1 else root
    terms = {k: v for k, _, v in steps[-1] if k != "nth"}
    want_role = terms.get("role", "").lower().replace("control", "")
    want_name = terms.get("name", "").lower()
    want_aid = terms.get("aid", "").lower()
    scored = []
    try:
        walked = list(auto.WalkControl(scope, includeTop=False, maxDepth=40))
    except Exception:
        walked = []
    for ctrl, _d in walked:
        try:
            role = role_of(ctrl).lower()
            if want_role and role != want_role:
                continue
            name, aid = (ctrl.Name or "").lower(), (ctrl.AutomationId or "").lower()
        except Exception:
            continue
        score = 0.0
        if want_aid and aid == want_aid:
            score = 1.0
        if want_name and name:
            ratio = difflib.SequenceMatcher(None, want_name, name).ratio()
            if want_name in name or name in want_name:
                ratio = max(ratio, 0.8)
            score = max(score, ratio)
        if score:
            scored.append((score, ctrl))
    scored.sort(key=lambda t: -t[0])
    if not scored or scored[0][0] < HEAL_MIN or (
            len(scored) > 1 and scored[0][0] - scored[1][0] < HEAL_MARGIN):
        near = "; ".join(f"{role_of(c)} {c.Name!r} ({s:.0%})" for s, c in scored[:3])
        raise WadError("ELEMENT_NOT_FOUND", f"nothing matches {text!r}, and no single close "
                       "match to heal it with" + (f" (nearest: {near})" if near else ""),
                       "the app changed: snapshot it and fix this step's selector")
    best = scored[0][1]
    el = describe(best)
    new = selector_for(el)
    if len(steps) > 1:
        new = text.rsplit(">>", 1)[0].strip() + " >> " + new
    return best, new, scored[0][0]


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
            while r and r not in keep:
                keep.add(r)
                r = by_ref[r]["parent"] if r in by_ref else None
    return keep


def render(elements, interactive_only):
    keep = keep_for_interactive(elements) if interactive_only else None
    lines = []
    base = elements[0]["depth"] if elements else 0
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
        if e.get("focused"):
            extra.append("*focused*")
        if e.get("truncated") and not e["interactive"]:
            extra.append(f"(more inside: --root {e['ref']})")
        lines.append(f"{'  ' * (e['depth'] - base)}{e['ref']} {e['role']} {label} "
                     f"{' '.join(extra)}".rstrip())
    return "\n".join(lines)
