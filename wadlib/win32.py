"""The few raw Win32 calls UI Automation does not cover: who is in front, bringing a
window forward for real, keyboard layouts, DPI and elevation.

Everything is loaded lazily so the rest of wad imports (and its tests run) anywhere."""
import ctypes
import os
import sys
import time

IS_WINDOWS = sys.platform == "win32"
WM_INPUTLANGCHANGEREQUEST = 0x0050
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
TOKEN_QUERY = 0x0008
TOKEN_ELEVATION = 20
LAYOUTS = {"en": "00000409", "ar": "00000401", "fr": "0000040C", "de": "00000407"}

_api = None


def api():
    global _api
    if _api is None:
        from ctypes import wintypes
        u32, k32, adv = ctypes.windll.user32, ctypes.windll.kernel32, ctypes.windll.advapi32
        u32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
        u32.GetForegroundWindow.restype = wintypes.HWND
        u32.LoadKeyboardLayoutW.argtypes = [wintypes.LPCWSTR, wintypes.UINT]
        u32.LoadKeyboardLayoutW.restype = ctypes.c_void_p
        u32.GetKeyboardLayout.argtypes = [wintypes.DWORD]
        u32.GetKeyboardLayout.restype = ctypes.c_void_p
        u32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, ctypes.c_void_p]
        k32.OpenProcess.restype = wintypes.HANDLE
        k32.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD,
                                                   wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
        adv.OpenProcessToken.argtypes = [wintypes.HANDLE, wintypes.DWORD,
                                         ctypes.POINTER(wintypes.HANDLE)]
        _api = (u32, k32, adv, wintypes)
    return _api


def foreground():
    return api()[0].GetForegroundWindow() or 0


def desktop_info():
    """Identify the desktop this process sees and the desktop receiving user input.

    A process on WinSta0\\exebox-* can run UIA successfully while seeing no windows on
    WinSta0\\Default. A named pipe or ordinary file cannot change that desktop binding.
    Return unknown fields when Windows denies the query instead of claiming access.
    """
    if not IS_WINDOWS:
        return {"station": None, "desktop": None, "input_desktop": None,
                "accessible": None}
    from ctypes import wintypes

    u32, k32, _, _ = api()
    u32.GetProcessWindowStation.restype = wintypes.HANDLE
    u32.GetThreadDesktop.argtypes = [wintypes.DWORD]
    u32.GetThreadDesktop.restype = wintypes.HANDLE
    u32.OpenInputDesktop.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    u32.OpenInputDesktop.restype = wintypes.HANDLE
    u32.CloseDesktop.argtypes = [wintypes.HANDLE]
    u32.GetUserObjectInformationW.argtypes = [wintypes.HANDLE, ctypes.c_int,
                                              ctypes.c_void_p, wintypes.DWORD,
                                              ctypes.POINTER(wintypes.DWORD)]

    def name(handle):
        if not handle:
            return None
        size = wintypes.DWORD()
        u32.GetUserObjectInformationW(handle, 2, None, 0, ctypes.byref(size))  # UOI_NAME
        if not size.value:
            return None
        buf = ctypes.create_unicode_buffer((size.value + 1) // ctypes.sizeof(ctypes.c_wchar))
        return (buf.value if u32.GetUserObjectInformationW(
            handle, 2, buf, ctypes.sizeof(buf), ctypes.byref(size)) else None)

    station = name(u32.GetProcessWindowStation())
    desktop = name(u32.GetThreadDesktop(k32.GetCurrentThreadId()))
    input_handle = u32.OpenInputDesktop(0, False, 0x0001)  # DESKTOP_READOBJECTS
    try:
        input_desktop = name(input_handle)
    finally:
        if input_handle:
            u32.CloseDesktop(input_handle)
    accessible = (desktop.lower() == input_desktop.lower()
                  if desktop and input_desktop else None)
    return {"station": station, "desktop": desktop,
            "input_desktop": input_desktop, "accessible": accessible}


def pid_of(hwnd):
    u32, _, _, wintypes = api()
    pid = wintypes.DWORD()
    u32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def in_front(hwnd, pid=None):
    """The window itself, or any window of its process (an open menu or a dialog is its
    own foreground popup; pulling the main window forward would close it)."""
    fg = foreground()
    if fg == hwnd:
        return True
    return bool(fg) and pid_of(fg) == (pid or pid_of(hwnd))


def bring_front(hwnd):
    """SetForegroundWindow alone is refused unless the caller owns the foreground, so
    attach to the foreground thread's input first - the standard, documented dance."""
    u32, k32, _, _ = api()
    if u32.GetForegroundWindow() == hwnd:
        return True
    if u32.IsIconic(hwnd):
        u32.ShowWindow(hwnd, 9)                       # SW_RESTORE
    fg = u32.GetForegroundWindow()
    fg_thread = u32.GetWindowThreadProcessId(fg, None)
    me = k32.GetCurrentThreadId()
    u32.AttachThreadInput(me, fg_thread, True)
    try:
        u32.BringWindowToTop(hwnd)
        u32.SetForegroundWindow(hwnd)
    finally:
        u32.AttachThreadInput(me, fg_thread, False)
    time.sleep(0.25)
    return u32.GetForegroundWindow() == hwnd


def is_hung(hwnd):
    """Windows' own verdict: the window has not pumped messages for about 5 seconds.
    Every UIA call into such an app blocks, so ask before touching it."""
    if not IS_WINDOWS or not hwnd:
        return False
    try:
        return bool(api()[0].IsHungAppWindow(hwnd))
    except Exception:
        return False


def input_language(hwnd):
    u32 = api()[0]
    thread = u32.GetWindowThreadProcessId(hwnd, None)
    return (u32.GetKeyboardLayout(thread) or 0) & 0xFFFF


def set_input_language(hwnd, lang):
    """Switch only this window's keyboard layout (other windows keep theirs)."""
    klid = LAYOUTS.get(lang, lang)
    want = int(klid, 16) & 0xFFFF
    if input_language(hwnd) == want:
        return True
    u32 = api()[0]
    hkl = u32.LoadKeyboardLayoutW(klid, 1)            # KLF_ACTIVATE
    for _ in range(3):
        u32.PostMessageW(hwnd, WM_INPUTLANGCHANGEREQUEST, 0, hkl)
        deadline = time.time() + 2
        while time.time() < deadline:
            if input_language(hwnd) == want:
                return True
            time.sleep(0.1)
    return False


def set_dpi_aware():
    """Physical pixels everywhere: UIA rectangles, screenshots and mouse coordinates must
    agree, and on a 150% display they only do when the process is per-monitor aware."""
    if not IS_WINDOWS:
        return
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            pass


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def process_elevated(pid):
    """True / False, or None when the token can not be read - which, from a normal
    process, itself usually means the target runs elevated."""
    _, k32, adv, wintypes = api()
    h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h:
        return None
    try:
        tok = wintypes.HANDLE()
        if not adv.OpenProcessToken(h, TOKEN_QUERY, ctypes.byref(tok)):
            return None
        try:
            elev, size = wintypes.DWORD(), wintypes.DWORD()
            if not adv.GetTokenInformation(tok, TOKEN_ELEVATION, ctypes.byref(elev),
                                           ctypes.sizeof(elev), ctypes.byref(size)):
                return None
            return bool(elev.value)
        finally:
            k32.CloseHandle(tok)
    finally:
        k32.CloseHandle(h)


def process_name(pid):
    _, k32, _, wintypes = api()
    h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(1024)
        size = wintypes.DWORD(len(buf))
        if k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
            return os.path.basename(buf.value)
        return ""
    finally:
        k32.CloseHandle(h)


# ---------------------------------------------------------------------------
# Layout-proof typing
# ---------------------------------------------------------------------------
INPUT_KEYBOARD, KEYEVENTF_KEYUP, KEYEVENTF_UNICODE = 1, 0x0002, 0x0004
VK_RETURN, VK_TAB = 0x0D, 0x09
_structs = None


def _input_structs():
    global _structs
    if _structs is None:
        from ctypes import wintypes

        class KEYBDINPUT(ctypes.Structure):
            _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                        ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                        ("dwExtraInfo", ctypes.c_size_t)]

        class MOUSEINPUT(ctypes.Structure):
            _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG),
                        ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD),
                        ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.c_size_t)]

        class _U(ctypes.Union):
            _fields_ = [("ki", KEYBDINPUT), ("mi", MOUSEINPUT)]

        class INPUT(ctypes.Structure):
            _fields_ = [("type", wintypes.DWORD), ("u", _U)]

        _structs = (INPUT, KEYBDINPUT)
    return _structs


def _key_events(text):
    """(vk, scan, flags) pairs: printable characters as UTF-16 units with
    KEYEVENTF_UNICODE, newlines and tabs as real Enter/Tab keys."""
    events = []
    for ch in text.replace("\r\n", "\n").replace("\r", "\n"):
        if ch == "\n" or ch == "\t":
            vk = VK_RETURN if ch == "\n" else VK_TAB
            events += [(vk, 0, 0), (vk, 0, KEYEVENTF_KEYUP)]
            continue
        units = ch.encode("utf-16-le")
        for i in range(0, len(units), 2):
            unit = int.from_bytes(units[i:i + 2], "little")
            events += [(0, unit, KEYEVENTF_UNICODE),
                       (0, unit, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP)]
    return events


def type_unicode(text, interval=0.004):
    """Type text as Unicode characters, not as keys. SendKeys-style typing maps each
    character to a key of OUR keyboard layout, and with an Arabic layout active in the
    target the same keys produce other letters; Unicode input arrives as written in
    any layout, any language."""
    INPUT, KEYBDINPUT = _input_structs()
    send = ctypes.windll.user32.SendInput
    events = _key_events(text)
    for i in range(0, len(events), 2):
        pair = (INPUT * 2)()
        for j, (vk, scan, flags) in enumerate(events[i:i + 2]):
            pair[j].type = INPUT_KEYBOARD
            pair[j].u.ki = KEYBDINPUT(vk, scan, flags, 0, 0)
        send(2, pair, ctypes.sizeof(INPUT))
        if interval:
            time.sleep(interval)


def virtual_screen_origin():
    """Top-left of the desktop spanning all monitors (negative with a monitor on the left)."""
    m = ctypes.windll.user32.GetSystemMetrics
    return m(76), m(77)                                  # SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN
