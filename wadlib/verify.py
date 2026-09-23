"""Proving an action worked.

Apps lie: Excel's ValuePattern reports success and even reads the new value back while
never storing it, and an Invoke returns fine on a button that did nothing. So every
action is followed by a look at what actually changed - the element, the focus, the
app's windows - and an optional --expect that must become true."""
import time

import uiautomation as auto

from . import uia
from .registry import WadError


def observe(ctrl, hwnd):
    """A cheap fingerprint of the element and its app, to diff before/after an action."""
    sig = {"alive": uia.alive(ctrl)}
    if sig["alive"]:
        try:
            sig["name"] = ctrl.Name or ""
            sig["enabled"] = bool(ctrl.IsEnabled)
            sig["content"] = uia.content_of(ctrl)
            sig.update(uia.states_of(ctrl))
        except Exception:
            sig["alive"] = False
    try:
        f = auto.GetFocusedControl()
        sig["focus"] = (tuple(f.GetRuntimeId() or []), uia.role_of(f), f.Name or "")
    except Exception:
        sig["focus"] = None
    w = auto.ControlFromHandle(hwnd) if hwnd else None
    sig["window"] = (w.Name or "") if w and uia.alive(w) else None
    pid = w.ProcessId if w and sig["window"] is not None else None
    wins = {}
    try:
        for t in auto.GetRootControl().GetChildren():
            try:
                if pid and t.ProcessId == pid and not t.IsOffscreen:
                    wins[t.NativeWindowHandle] = t.Name or f"({uia.role_of(t)})"
            except Exception:
                continue
    except Exception:
        pass
    sig["windows"] = wins
    return sig


def diff(before, after):
    """Human-readable changes, most telling first."""
    out = []
    if before.get("alive") and not after.get("alive"):
        out.append("the element went away")
    for h, name in after["windows"].items():
        if h not in before["windows"]:
            out.append(f"window opened: {name!r}")
    for h, name in before["windows"].items():
        if h not in after["windows"]:
            out.append(f"window closed: {name!r}")
    if before.get("window") != after.get("window") and after.get("window") is not None:
        out.append(f"title: {before.get('window')!r} -> {after.get('window')!r}")
    if after.get("alive"):
        for k in ("content", "toggle", "expanded", "selected", "name", "enabled"):
            if k in after and before.get(k) != after.get(k):
                b, a = before.get(k), after.get(k)
                if k == "content":
                    b, a = _short(b), _short(a)
                out.append(f"{k}: {b!r} -> {a!r}")
    if before.get("focus") != after.get("focus") and after.get("focus"):
        _, role, name = after["focus"]
        out.append(f"focus -> {role} {name!r}")
    return out


def _short(v):
    if v is None:
        return None
    v = str(v).replace("\r", " ").replace("\n", " ")
    return v if len(v) <= 60 else v[:57] + "..."


def settle(ctrl, hwnd, before, wait=0.8):
    """Poll briefly for the UI to react (it is asynchronous) and return what changed."""
    deadline = time.time() + wait
    while True:
        changes = diff(before, observe(ctrl, hwnd))
        if changes or time.time() >= deadline:
            return changes
        time.sleep(0.1)


def text_present(hwnd, text, depth=30):
    """Is `text` in a name or value anywhere in this app's windows?"""
    want = text.lower()
    main = auto.ControlFromHandle(hwnd)
    if not main or not uia.alive(main):
        return False
    roots = [main]
    try:
        pid = main.ProcessId
        roots += [t for t in auto.GetRootControl().GetChildren()
                  if t.NativeWindowHandle != hwnd and t.ProcessId == pid]
    except Exception:
        pass
    for root in roots:
        if want in (root.Name or "").lower():
            return True
        try:
            for c, _d in auto.WalkControl(root, maxDepth=depth):
                try:
                    if want in (c.Name or "").lower():
                        return True
                    if c.ControlTypeName in ("EditControl", "DocumentControl", "TextControl"):
                        if want in str(uia.value_of(c) or "").lower():
                            return True
                except Exception:
                    continue
        except Exception:
            continue
    return False


def expect(hwnd, text, gone=False, timeout=5.0):
    deadline = time.time() + timeout
    while True:
        present = text_present(hwnd, text)
        if present != gone:
            return
        if time.time() >= deadline:
            what = "still there" if gone else "did not appear"
            raise WadError("VERIFY_FAILED", f"{text!r} {what} within {timeout}s",
                           "the action ran but did not have the expected effect - snapshot "
                           "to see the real state before retrying")
        time.sleep(0.25)


def after_action(args, ctrl, hwnd, before):
    """Shared post-condition step. Returns (changes, verified) for the payload."""
    if getattr(args, "no_verify", False):
        return [], None
    changes = settle(ctrl, hwnd, before)
    if getattr(args, "expect", None):
        expect(hwnd, args.expect, timeout=args.timeout)
    if getattr(args, "expect_gone", None):
        expect(hwnd, args.expect_gone, gone=True, timeout=args.timeout)
    return changes, bool(changes) or bool(getattr(args, "expect", None)
                                          or getattr(args, "expect_gone", None))


def same_text(want, have):
    norm = lambda s: str(s or "").replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")
    return norm(want) == norm(have)
