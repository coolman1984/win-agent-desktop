"""Real keyboard and mouse input - always guarded.

Keystrokes and physical clicks go to whatever is in front / under the pointer, not to
the element we mean. So before any of them: the target app must be in front, and for a
click, the point must really land on the target (or at least inside the same app)."""
import time

import uiautomation as auto

from . import uia, win32
from .registry import WadError

KEYS = {
    "ctrl": auto.Keys.VK_CONTROL, "control": auto.Keys.VK_CONTROL, "alt": auto.Keys.VK_MENU,
    "shift": auto.Keys.VK_SHIFT, "win": auto.Keys.VK_LWIN, "enter": auto.Keys.VK_RETURN,
    "return": auto.Keys.VK_RETURN, "tab": auto.Keys.VK_TAB, "esc": auto.Keys.VK_ESCAPE,
    "escape": auto.Keys.VK_ESCAPE, "space": auto.Keys.VK_SPACE, "backspace": auto.Keys.VK_BACK,
    "delete": auto.Keys.VK_DELETE, "del": auto.Keys.VK_DELETE, "insert": auto.Keys.VK_INSERT,
    "home": auto.Keys.VK_HOME, "end": auto.Keys.VK_END, "up": auto.Keys.VK_UP,
    "down": auto.Keys.VK_DOWN, "left": auto.Keys.VK_LEFT, "right": auto.Keys.VK_RIGHT,
    "pageup": auto.Keys.VK_PRIOR, "pagedown": auto.Keys.VK_NEXT, "apps": auto.Keys.VK_APPS,
    "menu": auto.Keys.VK_APPS, "capslock": auto.Keys.VK_CAPITAL,
    "printscreen": auto.Keys.VK_SNAPSHOT,
}


def vk(name):
    name = name.strip().lower()
    if name in KEYS:
        return KEYS[name]
    if len(name) == 1 and name.isascii() and name.isalnum():
        return ord(name.upper())
    if name.startswith("f") and name[1:].isdigit() and 1 <= int(name[1:]) <= 24:
        return auto.Keys.VK_F1 + int(name[1:]) - 1
    raise WadError("USAGE", f"unknown key {name!r}",
                   "keys: ctrl alt shift win enter tab esc space backspace delete home end "
                   "up down left right pageup pagedown f1-f24 apps, letters and digits")


def parse_keys(spec):
    """'ctrl+s' is one chord; 'alt h o r' is a sequence of chords (Office key tips)."""
    return [[vk(k) for k in chord.split("+")] for chord in spec.split()]


def press_chord(codes):
    for c in codes:
        auto.PressKey(c, waitTime=0.02)
    for c in reversed(codes):
        auto.ReleaseKey(c, waitTime=0.02)


def guard(hwnd):
    """Never send input unless the target app is in front. Bring it forward ourselves
    first; refuse (and send nothing) if another window keeps the foreground."""
    if not hwnd or win32.in_front(hwnd):
        return
    for _ in range(3):
        if win32.bring_front(hwnd):
            return
        time.sleep(0.3)
    if win32.in_front(hwnd):
        return
    raise WadError("FOCUS_LOST", "the target window could not be brought to the front; "
                   "nothing was sent",
                   "another window (or a UAC/elevated window) holds the foreground - close it "
                   "or `wad focus --window <title>` and retry")


def center(ctrl, el=None):
    l, t, r, b = uia.rect_of(ctrl)
    if r <= l or b <= t:
        raise WadError("NOT_VISIBLE", f"{uia.role_of(ctrl)} {ctrl.Name!r} has no area on screen",
                       "scroll it into view (`wad scroll-to`) or open its parent first")
    return (l + r) // 2, (t + b) // 2


def hit_check(ctrl, x, y):
    """'target' when (x, y) lands on the element or inside it, 'same-app' when it lands
    elsewhere in the same app; anything else is refused - that click would hit another
    window that covers ours."""
    try:
        hit = auto.ControlFromPoint(x, y)
    except Exception:
        hit = None
    if hit is None:
        return "unknown"
    want = list(ctrl.GetRuntimeId() or [])
    cur, steps = hit, 0
    while cur is not None and steps < 40:
        if list(cur.GetRuntimeId() or []) == want:
            return "target"
        cur = cur.GetParentControl()
        steps += 1
    try:
        if hit.ProcessId == ctrl.ProcessId:
            return "same-app"
    except Exception:
        pass
    raise WadError("OCCLUDED", f"the point ({x}, {y}) is covered by {uia.role_of(hit)} "
                   f"{hit.Name!r} of another app; nothing was clicked",
                   "bring the target window to the front or move what covers it")


def ensure_visible(ctrl):
    if ctrl.IsOffscreen:
        sip = ctrl.GetPattern(auto.PatternId.ScrollItemPattern)
        if sip:
            try:
                sip.ScrollIntoView()
                time.sleep(0.2)
            except Exception:
                pass


def click_at(x, y, button="left", count=1):
    fn = {"left": auto.Click, "right": auto.RightClick, "middle": auto.MiddleClick}[button]
    for i in range(count):
        fn(x, y, waitTime=0.05 if i < count - 1 else 0.2)


def physical_click(ctrl, hwnd, button="left", count=1):
    ensure_visible(ctrl)
    guard(hwnd)
    x, y = center(ctrl)
    landed = hit_check(ctrl, x, y)
    click_at(x, y, button, count)
    return landed


def type_text(ctrl, hwnd, text, replace):
    """Focus, optionally select all, then type as Unicode (layout-proof)."""
    guard(hwnd)
    try:
        ctrl.SetFocus()
    except Exception:
        pass
    time.sleep(0.05)
    if replace:
        press_chord([auto.Keys.VK_CONTROL, ord("A")])
        time.sleep(0.05)
    else:
        press_chord([auto.Keys.VK_CONTROL, auto.Keys.VK_END])
    win32.type_unicode(text)
    time.sleep(0.1)
