"""Learning from the person: record what they do, as a replayable batch file.

A person does the task once, by hand. wad watches through low-level mouse and keyboard
hooks and turns what happened into the same steps an agent would write: clicks become
`click` on a durable selector, a filled field becomes one `type` with its final value
(read from the field, not reconstructed from keystrokes - so autocomplete, paste and
backspace all come out right), a changed drop-down becomes `select`, and shortcuts
become `press`. `wad batch` replays the file, healing steps if the app changed.

The hooks only queue raw events; a worker thread does the slow accessibility lookups,
because Windows silently removes a low-level hook that answers too slowly."""
import ctypes
import os
import queue
import threading
import time

import uiautomation as auto

from . import state, uia, win32
from .registry import Arg, WadError, command

STOP_FILE = os.path.join(state.STATE_DIR, "record.stop")
RECORDINGS = os.path.join(state.STATE_DIR, "recordings")
EDITABLE = {"Edit", "Document"}
MODIFIERS = {0x10: "shift", 0x11: "ctrl", 0x12: "alt", 0xA0: "shift", 0xA1: "shift",
             0xA2: "ctrl", 0xA3: "ctrl", 0xA4: "alt", 0xA5: "alt", 0x5B: "win", 0x5C: "win"}
NAMED_KEYS = {0x0D: "enter", 0x09: "tab", 0x1B: "esc", 0x21: "pageup", 0x22: "pagedown",
              0x23: "end", 0x24: "home", 0x25: "left", 0x26: "up", 0x27: "right",
              0x28: "down", 0x2E: "delete", 0x2D: "insert", 0x5D: "apps"}
# keys that only edit text: the field's final value already says what they did
TEXT_KEYS = {0x08, 0x2E, 0x25, 0x27, 0x23, 0x24}
STOP_VK = 0x7B                                   # F12, with Ctrl+Shift


def key_name(vk):
    if vk in NAMED_KEYS:
        return NAMED_KEYS[vk]
    if 0x70 <= vk <= 0x87:
        return f"f{vk - 0x6F}"
    if 0x30 <= vk <= 0x39 or 0x41 <= vk <= 0x5A:
        return chr(vk).lower()
    return None


class Recorder:
    """Turns raw input events into steps. Knows nothing about hooks, so it is tested with
    plain calls."""

    def __init__(self, window=None):
        self.window = (window or "").lower()
        self.steps = []
        self.field = None               # (ctrl, selector, window, role, secret, value)
        self.last_click = None

    # -- helpers ------------------------------------------------------------------
    def _top_title(self, ctrl):
        top = uia.window_of(ctrl)
        title = top.Name or ""
        if not title:                   # a menu or drop-down: name the app's main window
            try:
                pid = top.ProcessId
                title = next((w.Name for w in uia.top_windows() if w.ProcessId == pid), "")
            except Exception:
                title = ""
        return top, title

    def _selector(self, ctrl, top):
        """A durable selector, made unique inside its window with nth= if it has to be."""
        el = uia.describe(ctrl)
        sel = uia.selector_for(el)
        try:
            hits = uia.search(top, uia.parse_selector(sel)[0])
            rid = list(ctrl.GetRuntimeId() or [])
            if len(hits) > 1:
                idx = next((i for i, h in enumerate(hits)
                            if list(h.GetRuntimeId() or []) == rid), None)
                if idx is not None:
                    sel += f" nth={idx + 1}"
        except Exception:
            pass
        return sel, el

    def _wanted(self, title):
        return not self.window or self.window in title.lower()

    # -- events -------------------------------------------------------------------
    def on_click(self, x, y, button="left", when=None):
        self.flush_field()
        try:
            ctrl = auto.ControlFromPoint(x, y)
        except Exception:
            ctrl = None
        if ctrl is None:
            return
        top, title = self._top_title(ctrl)
        if not self._wanted(title) or top.ClassName in ("Shell_TrayWnd", "Progman"):
            return
        # an option picked in a drop-down list: the combo's new value is recorded instead
        if top.ClassName == "ComboLBox" or self._inside_combo(ctrl):
            return
        sel, el = self._selector(ctrl, top)
        when = when or time.time()
        if el["role"] in EDITABLE and button == "left":
            step = {"cmd": "focus", "target": sel, "window": title}
        else:
            step = {"cmd": {"left": "click", "right": "right-click"}.get(button, "click"),
                    "target": sel, "window": title}
        last = self.last_click
        if (button == "left" and last and last[0] == sel and when - last[1] < 0.5
                and self.steps and self.steps[-1].get("target") == sel):
            self.steps[-1] = {"cmd": "double-click", "target": sel, "window": title}
            self.last_click = None
            return
        self.last_click = (sel, when)
        if step["cmd"] == "focus" and self.steps and self.steps[-1] == step:
            return
        self.steps.append(step)

    def _inside_combo(self, ctrl):
        cur = ctrl.GetParentControl()
        for _ in range(3):
            if cur is None:
                return False
            if cur.ControlTypeName == "ComboBoxControl" and ctrl.ControlTypeName == "ListItemControl":
                return True
            cur = cur.GetParentControl()
        return False

    def on_key(self, vk, mods=()):
        name = key_name(vk)
        if name is None or vk in MODIFIERS:
            return
        mods = [m for m in ("ctrl", "alt", "shift", "win") if m in mods]
        chord = [m for m in mods if m != "shift"]
        if not chord and (len(name) == 1 or vk in TEXT_KEYS):
            return                                        # typing: the field value tells
        if not chord and name in ("up", "down") and self.field and self.field[3] == "ComboBox":
            return                                        # walking a drop-down: select tells
        if (chord == ["ctrl"] and name in ("a", "c", "v", "x", "z", "y") and self.field
                and self.field[3] in EDITABLE):
            return                                        # editing inside a field: its value tells
        self.flush_field()
        title = self._front_title()
        if not self._wanted(title):
            return
        self.steps.append({"cmd": "press", "combo": "+".join(mods + [name]), "window": title})

    def _front_title(self):
        if win32.IS_WINDOWS:
            fg = win32.foreground()
            w = auto.ControlFromHandle(fg) if fg else None
            if w is not None and uia.alive(w) and w.Name:
                return w.Name
        if self.field:
            return self.field[2]
        return next((s["window"] for s in reversed(self.steps) if s.get("window")), "")

    def poll_focus(self):
        try:
            f = auto.GetFocusedControl()
        except Exception:
            return
        if f is None:
            return
        rid = list(f.GetRuntimeId() or [])
        if self.field and list(self.field[0].GetRuntimeId() or []) == rid:
            return
        self.flush_field()
        role = uia.role_of(f)
        if role not in EDITABLE | {"ComboBox"}:
            self.field = None
            return
        top, title = self._top_title(f)
        if not self._wanted(title):
            self.field = None
            return
        sel, _ = self._selector(f, top)
        try:
            secret = bool(f.IsPassword)
        except Exception:
            secret = False
        self.field = (f, sel, title, role, secret, uia.content_of(f))

    def flush_field(self):
        """If the focused field's value changed since it got focus, record its final value."""
        if not self.field:
            return
        ctrl, sel, title, role, secret, was = self.field
        now = uia.content_of(ctrl) if uia.alive(ctrl) else was
        if now == was or now is None:
            return
        self.field = (ctrl, sel, title, role, secret, now)
        if role == "ComboBox":
            step = {"cmd": "select", "target": sel, "option": str(now), "window": title}
        else:
            step = {"cmd": "type", "target": sel, "window": title,
                    "text": "${ENV:WAD_SECRET}" if secret else str(now)}
        # a focus click on the field just before is implied by typing into it
        if self.steps and self.steps[-1].get("cmd") == "focus" and self.steps[-1]["target"] == sel:
            self.steps.pop()
        self.steps.append(step)

    def finish(self):
        self.flush_field()
        return list(self.steps)


# ---------------------------------------------------------------------------
# Low-level hooks (Windows)
# ---------------------------------------------------------------------------

def _hooks(events, stop, ignore_injected):
    from ctypes import wintypes
    user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32

    class POINT(ctypes.Structure):
        _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

    class MSLL(ctypes.Structure):
        _fields_ = [("pt", POINT), ("mouseData", wintypes.DWORD), ("flags", wintypes.DWORD),
                    ("time", wintypes.DWORD), ("extra", ctypes.c_size_t)]

    class KBDLL(ctypes.Structure):
        _fields_ = [("vkCode", wintypes.DWORD), ("scanCode", wintypes.DWORD),
                    ("flags", wintypes.DWORD), ("time", wintypes.DWORD),
                    ("extra", ctypes.c_size_t)]

    LRESULT = ctypes.c_ssize_t
    PROC = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)
    user32.CallNextHookEx.argtypes = [wintypes.HHOOK, ctypes.c_int, wintypes.WPARAM,
                                      wintypes.LPARAM]
    user32.CallNextHookEx.restype = LRESULT
    user32.SetWindowsHookExW.argtypes = [ctypes.c_int, PROC, wintypes.HINSTANCE, wintypes.DWORD]
    user32.SetWindowsHookExW.restype = wintypes.HHOOK
    user32.UnhookWindowsHookEx.argtypes = [wintypes.HHOOK]
    # 64-bit handles: without these, ctypes cuts them to 32 bits and the hook is refused
    kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
    kernel32.GetModuleHandleW.restype = wintypes.HMODULE
    user32.MsgWaitForMultipleObjects.argtypes = [wintypes.DWORD, ctypes.c_void_p, wintypes.BOOL,
                                                 wintypes.DWORD, wintypes.DWORD]
    user32.PeekMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT,
                                    wintypes.UINT, wintypes.UINT]
    held = set()

    def on_mouse(code, wparam, lparam):
        if code == 0 and wparam in (0x201, 0x204):          # left / right button down
            m = ctypes.cast(lparam, ctypes.POINTER(MSLL)).contents
            if not (ignore_injected and m.flags & 0x1):
                events.put(("click", m.pt.x, m.pt.y, "left" if wparam == 0x201 else "right",
                            time.time()))
        return user32.CallNextHookEx(None, code, wparam, lparam)

    def on_key(code, wparam, lparam):
        if code == 0:
            k = ctypes.cast(lparam, ctypes.POINTER(KBDLL)).contents
            vk = k.vkCode
            down = wparam in (0x100, 0x104)
            if vk in MODIFIERS:
                (held.add if down else held.discard)(MODIFIERS[vk])
            elif down and not (ignore_injected and k.flags & 0x10):
                if vk == STOP_VK and {"ctrl", "shift"} <= held:
                    stop.set()
                else:
                    events.put(("key", vk, frozenset(held)))
        return user32.CallNextHookEx(None, code, wparam, lparam)

    mouse_proc, key_proc = PROC(on_mouse), PROC(on_key)      # keep references alive
    hm = user32.SetWindowsHookExW(14, mouse_proc, kernel32.GetModuleHandleW(None), 0)
    hk = user32.SetWindowsHookExW(13, key_proc, kernel32.GetModuleHandleW(None), 0)
    if not hm or not hk:
        err = ctypes.get_last_error() or kernel32.GetLastError()
        raise WadError("NOT_SUPPORTED", f"Windows refused the input hooks (error {err})",
                       "another security tool may block them; run from a normal desktop")
    msg = wintypes.MSG()
    try:
        while not stop.is_set():
            # QS_ALLINPUT: wake for anything, at least every 100 ms to check `stop`
            user32.MsgWaitForMultipleObjects(0, None, False, 100, 0x04FF)
            while user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1):
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
    finally:
        user32.UnhookWindowsHookEx(hm)
        user32.UnhookWindowsHookEx(hk)


def _worker(rec, events, stop):
    with auto.UIAutomationInitializerInThread():
        while not (stop.is_set() and events.empty()):
            try:
                ev = events.get(timeout=0.15)
            except queue.Empty:
                rec.poll_focus()
                continue
            try:
                if ev[0] == "click":
                    rec.on_click(ev[1], ev[2], ev[3], ev[4])
                    time.sleep(0.05)
                    rec.poll_focus()
                else:
                    rec.on_key(ev[1], ev[2])
                    rec.poll_focus()
            except Exception:
                continue            # an element that vanished mid-lookup is just skipped
        rec.finish()


@command("record", "watch a PERSON do a task and save it as a replayable batch file "
         "(stop with Ctrl+Shift+F12, `record-stop`, or --seconds)",
         Arg("out", positional=True, optional=True, help="the batch file to write"),
         Arg("window", short="-w", help="only record in windows whose title contains this"),
         Arg("seconds", float, help="stop by itself after this long"),
         Arg("ignore-injected", bool, help="skip synthetic input (other tools, remote "
             "control) - only real hands"),
         group="workflow", mcp=False, long=True)
def cmd_record(args):
    if not win32.IS_WINDOWS:
        raise WadError("NOT_SUPPORTED", "recording needs a Windows desktop")
    out = os.path.abspath(args.out or os.path.join(
        RECORDINGS, time.strftime("rec-%Y%m%d-%H%M%S.json")))
    try:
        os.remove(STOP_FILE)
    except OSError:
        pass
    rec, events, stop = Recorder(args.window), queue.Queue(), threading.Event()
    worker = threading.Thread(target=_worker, args=(rec, events, stop), daemon=True)
    worker.start()
    deadline = time.time() + args.seconds if args.seconds else None

    def watch_stop():
        while not stop.is_set():
            if os.path.exists(STOP_FILE) or (deadline and time.time() >= deadline):
                stop.set()
            time.sleep(0.2)
    threading.Thread(target=watch_stop, daemon=True).start()
    _hooks(events, stop, args.ignore_injected)
    worker.join(15)
    steps = rec.finish()
    state.write_json(out, {"steps": steps, "recorded": state.now_iso()})
    try:
        os.remove(STOP_FILE)
    except OSError:
        pass
    lines = [f"{i:>3} {s['cmd']:<12} {s.get('target') or s.get('combo', '')}"
             + (f"  {s['text']!r}" if "text" in s else "") + (f"  {s['option']!r}" if "option" in s else "")
             for i, s in enumerate(steps, 1)]
    return ({"ok": True, "file": out, "steps": steps},
            "\n".join(lines + [f"recorded {len(steps)} steps to {out}", f"replay: wad batch \"{out}\""]))


@command("record-stop", "stop a running `record` (from another terminal or agent)",
         group="workflow")
def cmd_record_stop(args):
    os.makedirs(os.path.dirname(STOP_FILE), exist_ok=True)
    with open(STOP_FILE, "w") as fh:
        fh.write(state.now_iso())
    return {"ok": True}, "asked the recorder to stop"
