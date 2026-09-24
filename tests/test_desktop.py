"""Desktop access is checked before wad can launch an invisible app."""

from wadlib import commands, win32


def test_desktop_check_reports_interactive_desktop(app, run, monkeypatch):
    monkeypatch.setattr(win32, "desktop_info", lambda: {
        "station": "WinSta0", "desktop": "Default", "input_desktop": "Default",
        "accessible": True})
    out = run("desktop-check")
    assert out["ok"] and out["accessible"] is True
    assert out["visible_windows"] >= 1


def test_isolated_desktop_is_diagnosed_before_launch(app, run, monkeypatch):
    monkeypatch.setattr(win32, "desktop_info", lambda: {
        "station": "WinSta0", "desktop": "CodexSandboxDesktop-test",
        "input_desktop": "Default", "accessible": False})
    monkeypatch.setattr(commands, "spawn", lambda cmd: (_ for _ in ()).throw(
        AssertionError("launch must not spawn from the wrong desktop")))
    check = run("desktop-check")
    launch = run("launch", "notepad.exe")
    assert check["code"] == launch["code"] == "DESKTOP_UNAVAILABLE"
    assert "interactive desktop" in check["hint"]
