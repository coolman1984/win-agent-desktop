"""Every test runs against tests/fake_uia.py instead of the real UI Automation, with a
private state directory, and with the raw Win32 input calls replaced by recorders."""
import io
import json
import os
import sys
import tempfile
from contextlib import redirect_stdout

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)
os.environ["WAD_STATE_DIR"] = tempfile.mkdtemp(prefix="wad-test-")
os.environ.pop("WAD_SESSION", None)
os.environ.pop("WAD_ALLOW_SYSTEM", None)

import fake_uia  # noqa: E402

sys.modules["uiautomation"] = fake_uia

from wadlib import cli, state, verify, win32  # noqa: E402


@pytest.fixture
def app(monkeypatch):
    parts = fake_uia.build()
    for f in (state.SNAPSHOT_FILE, state.SHOT_FILE, state.TRACE_FILE, state.POLICY_FILE):
        if os.path.exists(f):
            os.remove(f)
    typed = []

    def type_unicode(text, interval=0):
        typed.append(text)
        target = fake_uia.FOCUS[0]
        vp = target.GetPattern(fake_uia.PatternId.ValuePattern) if target else None
        if vp is not None:
            vp._value = vp._value + text      # real keystrokes always land, even in Excel

    def press(code, waitTime=0):
        fake_uia.LOG.append(("press", (code,)))
        target = fake_uia.FOCUS[0]
        if target is not None and hasattr(target, "on_key"):
            target.on_key(code)
        vp = target.GetPattern(fake_uia.PatternId.ValuePattern) if target else None
        if vp is not None and code == ord("A") and any(
                e == ("press", (fake_uia.Keys.VK_CONTROL,)) for e in fake_uia.LOG[-2:]):
            vp._value = ""                     # ctrl+a then typing replaces everything

    monkeypatch.setattr(fake_uia, "PressKey", press)
    monkeypatch.setattr(win32, "type_unicode", type_unicode)
    monkeypatch.setattr(win32, "in_front", lambda hwnd, pid=None: True)
    monkeypatch.setattr(win32, "bring_front", lambda hwnd: True)
    monkeypatch.setattr(win32, "set_dpi_aware", lambda: None)
    monkeypatch.setattr(win32, "process_name", lambda pid: "notepad.exe")
    monkeypatch.setattr(win32, "pid_of", lambda hwnd: 100)
    monkeypatch.setattr(win32, "is_admin", lambda: False)
    monkeypatch.setattr(win32, "process_elevated", lambda pid: False)
    parts["typed"] = typed
    return parts


def wad(*argv):
    """Run wad like a shell would and return its JSON payload."""
    buf = io.StringIO()
    with redirect_stdout(buf):
        code = cli.main(["--json", *argv])
    payload = json.loads(buf.getvalue())
    payload["_exit"] = code
    return payload


@pytest.fixture
def run():
    return wad
