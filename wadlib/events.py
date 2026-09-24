"""Being told instead of looking: UI Automation events.

Polling means asking "anything new?" every few hundred milliseconds - slow to notice and
it misses what comes and goes in between (a toast, a dialog that flashes and closes).
UIA can call us instead when a window opens or closes, a menu opens or the focus moves.

Handlers run on UIA's own threads, so the watcher lives on its own multithreaded-COM
thread with its own IUIAutomation object (interfaces must not cross apartments), and
the handlers only put events on a queue. When events cannot be registered (no comtypes,
a locked-down PC, or the test fake) the same Watcher falls back to fast polling."""
import queue
import threading
import time

import uiautomation as auto

from . import uia
from .registry import Arg, command

WINDOW_OPENED, WINDOW_CLOSED, MENU_OPENED = 20016, 20017, 20003
KINDS = {WINDOW_OPENED: "window opened", WINDOW_CLOSED: "window closed",
         MENU_OPENED: "menu opened"}
TREE_SUBTREE = 7
NAME_PROP, PID_PROP, HWND_PROP, TYPE_PROP = 30005, 30002, 30020, 30003


class Watcher:
    """Collects events into a queue until stopped: {"t", "kind", "name", "pid", "hwnd"}."""

    def __init__(self, focus=True, pid=None):
        self.events, self.focus, self.pid = queue.Queue(), focus, pid
        self.mode, self.error = None, None
        self._stop = threading.Event()
        self._ready = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True, name="wad-events")

    def __enter__(self):
        self._thread.start()
        self._ready.wait(5)
        return self

    def __exit__(self, *exc):
        self._stop.set()
        self._thread.join(3)
        return False

    def get(self, timeout):
        try:
            return self.events.get(timeout=timeout)
        except queue.Empty:
            return None

    def _put(self, kind, name, pid, hwnd):
        if self.pid and pid and pid != self.pid:
            return
        self.events.put({"t": time.time(), "kind": kind, "name": name or "", "pid": pid,
                         "hwnd": hwnd})

    # -- real UIA events ----------------------------------------------------------
    def _run(self):
        try:
            self._run_events()
        except Exception as e:
            self.error = f"{e.__class__.__name__}: {e}"
            self._run_polling()

    def _run_events(self):
        import comtypes
        import comtypes.client
        comtypes.CoInitializeEx(comtypes.COINIT_MULTITHREADED)
        try:
            core = comtypes.client.GetModule("UIAutomationCore.dll")
            iuia = comtypes.client.CreateObject(core.CUIAutomation, interface=core.IUIAutomation)
            cache = iuia.CreateCacheRequest()
            for prop in (NAME_PROP, PID_PROP, HWND_PROP):
                cache.AddProperty(prop)
            watcher = self

            def cached(sender):
                vals = []
                for getter in ("CachedName", "CachedProcessId", "CachedNativeWindowHandle"):
                    try:
                        vals.append(getattr(sender, getter))
                    except Exception:
                        vals.append(None)
                return vals

            class Handler(comtypes.COMObject):
                _com_interfaces_ = [core.IUIAutomationEventHandler]

                def HandleAutomationEvent(self, sender, event_id):
                    name, pid, hwnd = cached(sender)
                    watcher._put(KINDS.get(event_id, str(event_id)), name, pid, hwnd)
                    return 0

            class FocusHandler(comtypes.COMObject):
                _com_interfaces_ = [core.IUIAutomationFocusChangedEventHandler]

                def HandleFocusChangedEvent(self, sender):
                    name, pid, hwnd = cached(sender)
                    watcher._put("focus", name, pid, hwnd)
                    return 0

            root = iuia.GetRootElement()
            handler = Handler()
            for ev in KINDS:
                iuia.AddAutomationEventHandler(ev, root, TREE_SUBTREE, cache, handler)
            focus_handler = None
            if self.focus:
                focus_handler = FocusHandler()
                iuia.AddFocusChangedEventHandler(cache, focus_handler)
            self.mode = "events"
            self._ready.set()
            while not self._stop.wait(0.1):
                pass
            iuia.RemoveAllEventHandlers()
        finally:
            comtypes.CoUninitialize()

    # -- fallback: fast polling with the same output ------------------------------------
    def _run_polling(self):
        self.mode = "polling"
        self._ready.set()
        with auto.UIAutomationInitializerInThread():
            def windows():
                out = {}
                for w in auto.GetRootControl().GetChildren():
                    try:
                        out[w.NativeWindowHandle] = (w.Name, w.ProcessId, w.ClassName)
                    except Exception:
                        continue
                return out

            def focus():
                try:
                    f = auto.GetFocusedControl()
                    return tuple(f.GetRuntimeId() or []), f.Name, f.ProcessId
                except Exception:
                    return None
            seen, fcur = windows(), focus()
            while not self._stop.wait(0.15):
                now = windows()
                for h, (name, pid, cls) in now.items():
                    if h not in seen:
                        self._put("menu opened" if cls == "#32768" else "window opened",
                                  name, pid, h)
                for h, (name, pid, _cls) in seen.items():
                    if h not in now:
                        self._put("window closed", name, pid, h)
                seen = now
                if self.focus:
                    f = focus()
                    if f and f != fcur:
                        self._put("focus", f[1], f[2], None)
                    fcur = f or fcur


@command("watch", "listen for a while and report what the desktop did: windows and menus "
         "opening/closing, focus moves (UIA events)",
         Arg("seconds", float, default=10.0, help="how long to listen"),
         Arg("window", short="-w", help="only events from this window's app"),
         Arg("no-focus", bool, help="leave out focus changes"),
         Arg("until", help="stop early when an event's name contains this text"),
         group="observe", readonly=True)
def cmd_watch(args):
    pid = uia.find_window(args.window).ProcessId if args.window else None
    got = []
    started = time.time()
    with Watcher(focus=not args.no_focus, pid=pid) as w:
        deadline = started + args.seconds
        while time.time() < deadline:
            ev = w.get(min(0.2, max(0.01, deadline - time.time())))
            if ev is None:
                continue
            got.append(ev)
            if args.until and args.until.lower() in ev["name"].lower():
                break
        mode, why = w.mode, w.error
    lines = [f"+{e['t'] - started:5.2f}s  {e['kind']:<14} {e['name']!r}" for e in got]
    head = f"{len(got)} events in {time.time() - started:.1f}s ({mode})"
    return ({"ok": True, "events": got, "mode": mode, **({"fallback_reason": why} if why else {})},
            "\n".join([head] + lines) if lines else head)


def wait_for(cond, timeout, pid=None):
    """Re-check `cond` whenever something happens on the desktop (or every 0.5 s),
    instead of on a fixed poll: a window that opens is noticed the moment it does."""
    deadline = time.time() + timeout
    with Watcher(focus=False, pid=pid) as w:
        while True:
            value = cond()
            if value:
                return value
            left = deadline - time.time()
            if left <= 0:
                return None
            w.get(min(0.5, left))
