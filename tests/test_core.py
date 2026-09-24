"""Seeing and acting: snapshots, refs, selectors, verification, typing, safety guards."""
import json

import fake_uia
import pytest

from wadlib import inputs, uia, win32
from wadlib.registry import WadError


def ref_of(snap, role, name):
    return next(e["ref"] for e in snap["elements"] if e["role"] == role and e["name"] == name)


def test_snapshot_lists_elements_with_refs_and_selectors(app, run):
    snap = run("snapshot", "--window", "Notepad", "-i")
    assert snap["ok"] and snap["window"] == "Untitled - Notepad"
    save = next(e for e in snap["elements"] if e["name"] == "Save")
    assert save["ref"].startswith("e") and "invoke" in save["actions"]
    found = run("find", "--name", "save")
    assert found["matches"][0]["selector"] == "role=Button aid=SaveButton"


def test_ambiguous_window_title_is_refused(app, run):
    fake_uia.ROOT.add(fake_uia.Control("Window", "Notepad Help", hwnd=1003))
    out = run("snapshot", "--window", "notepad")
    assert out["code"] == "AMBIGUOUS_WINDOW" and out["_exit"] == 1


def test_click_by_ref_reports_what_changed(app, run):
    snap = run("snapshot", "--window", "Notepad")
    out = run("click", ref_of(snap, "Button", "Save"))
    assert out["ok"] and out["via"] == "invoke"
    assert any("window opened: 'Save As'" in c for c in out["changes"])
    assert out["selector"] == "role=Button aid=SaveButton"


def test_click_by_selector_and_expect(app, run):
    run("snapshot", "--window", "Notepad")
    out = run("click", "role=Button name=Save", "--expect", "Cancel")
    assert out["ok"] and out["verified"] is True


def test_expect_that_never_happens_fails(app, run):
    run("snapshot", "--window", "Notepad")
    out = run("click", "name=Nothing", "--expect", "Saved", "--timeout", "0.3")
    assert out["code"] == "VERIFY_FAILED"


def test_click_without_effect_is_flagged_not_hidden(app, run):
    run("snapshot", "--window", "Notepad")
    out = run("click", "name=Nothing")
    assert out["ok"] and out["verified"] is False and out["changes"] == []


def test_disabled_element_is_refused_before_acting(app, run):
    run("snapshot", "--window", "Notepad")
    assert run("click", "name=Dead")["code"] == "NOT_ENABLED"


def test_ambiguous_selector_lists_candidates_and_nth_picks(app, run):
    run("snapshot", "--window", "Notepad")
    out = run("click", "role=Button name=OK")
    assert out["code"] == "AMBIGUOUS_TARGET" and "2 elements" in out["message"]
    assert run("click", "role=Button name=OK nth=2")["ok"]
    assert run("click", "name=Right >> role=Button name=OK")["ok"]


def test_selector_parsing():
    steps = uia.parse_selector('role=Button name="Save as" >> name~=ok nth=2')
    assert steps == [[("role", "=", "Button"), ("name", "=", "Save as")],
                     [("name", "~=", "ok"), ("nth", "=", 2)]]
    for bad in ("colour=red", "role=Button >>", "nth=0"):
        with pytest.raises(WadError):
            uia.parse_selector(bad)
    assert uia.is_ref("e12") and uia.is_ref("@s1234:e3") and not uia.is_ref("name=e12")


def test_stale_ref_after_ui_change(app, run):
    snap = run("snapshot", "--window", "Notepad")
    ref = ref_of(snap, "Button", "Nothing")
    app["noop"].remove()
    assert run("click", ref)["code"] in ("ELEMENT_GONE", "ELEMENT_CHANGED")


def test_type_uses_value_pattern_and_verifies(app, run):
    run("snapshot", "--window", "Notepad")
    out = run("type", "name='Text editor'", "hello")
    assert out["ok"] and out["via"] == "value" and out["verified"] is True
    assert app["doc"].GetPattern(fake_uia.PatternId.ValuePattern).Value == "hello"
    out = run("type", "name='Text editor'", " world", "--append")
    assert app["doc"].GetPattern(fake_uia.PatternId.ValuePattern).Value == "hello world"


def test_type_falls_back_to_keys_when_value_pattern_lies(app, run):
    """The Excel lesson: SetValue 'succeeds' and nothing is stored."""
    run("snapshot", "--window", "Notepad")
    out = run("type", "name=Cell", "42")
    assert out["ok"] and out["via"] == "keys" and out["verified"] is True
    assert any("did not stick" in n for n in out["notes"])
    assert app["typed"] == ["42"]


def test_arabic_text_goes_through_unicode_typing(app, run):
    run("snapshot", "--window", "Notepad")
    out = run("type", "name=Cell", "مرحبا")
    assert out["ok"] and app["typed"] == ["مرحبا"]


def test_password_is_never_recorded(app, run):
    run("snapshot", "--window", "Notepad")
    run("type", "name=Password", "hunter2")
    trace = run("trace", "--last", "50")
    steps = [e["step"] for e in trace["entries"] if e["cmd"] == "step"]
    assert steps[-1]["text"] == "${ENV:WAD_SECRET}"
    assert "hunter2" not in json.dumps(trace["entries"])


def test_check_uncheck_are_idempotent_and_verified(app, run):
    run("snapshot", "--window", "Notepad")
    assert run("check", "name='Word wrap'")["state"] == "on"
    assert run("check", "name='Word wrap'")["state"] == "on"
    assert run("uncheck", "name='Word wrap'")["state"] == "off"


def test_select_option_in_combo(app, run):
    run("snapshot", "--window", "Notepad")
    out = run("select", "role=ComboBox name=Font", "consolas")
    assert out["ok"] and out["option"] == "Consolas" and out["value"] == "Consolas"
    bad = run("select", "role=ComboBox name=Font", "Comic Sans")
    assert bad["code"] == "OPTION_NOT_FOUND" and "Courier New" in bad["hint"]


def test_scroll_reports_movement(app, run):
    run("snapshot", "--window", "Notepad")
    out = run("scroll", "name='row 1'", "--amount", "2")
    assert out["ok"] and out["moved"] and out["end"] == 20


def test_get_reads_live_state(app, run):
    run("snapshot", "--window", "Notepad")
    run("check", "name='Word wrap'")
    out = run("get", "name='Word wrap'")
    assert out["toggle"] == "on" and out["selector"] == "role=CheckBox name='Word wrap'"


def test_wait_for_text_and_gone(app, run):
    run("snapshot", "--window", "Notepad")
    run("click", "name=Save")
    assert run("wait", "--window", "Notepad", "--name", "Cancel", "--timeout", "1")["ok"]
    app["dialog"].remove()
    assert run("wait", "--window", "Notepad", "--name", "Cancel", "--gone",
               "--timeout", "1")["ok"]
    assert run("wait", "--window", "Notepad", "--name", "Nope",
               "--timeout", "0.3")["code"] == "TIMEOUT"


def test_popup_snapshot_finds_the_dialog(app, run):
    run("snapshot", "--window", "Notepad")
    run("click", "name=Save")
    out = run("snapshot", "--window", "Untitled - Notepad", "--popup")
    assert out["window"] == "Save As"
    assert run("click", "name=Cancel")["ok"]


def test_input_is_refused_when_app_cannot_be_brought_front(app, run, monkeypatch):
    monkeypatch.setattr(win32, "in_front", lambda hwnd, pid=None: False)
    monkeypatch.setattr(win32, "bring_front", lambda hwnd: False)
    monkeypatch.setattr(inputs.time, "sleep", lambda s: None)
    run("snapshot", "--window", "Notepad")
    out = run("press", "ctrl+s")
    assert out["code"] == "FOCUS_LOST"
    assert not [e for e in fake_uia.LOG if e[0] == "press"]


def test_headed_click_refuses_a_covered_point(app, run, monkeypatch):
    monkeypatch.setattr(fake_uia, "ControlFromPoint", lambda x, y: app["other"])
    run("snapshot", "--window", "Notepad")
    out = run("click", "name=Save", "--headed")
    assert out["code"] == "OCCLUDED"
    assert not [e for e in fake_uia.LOG if e[0] == "click"]


def test_headed_click_on_target_goes_through(app, run, monkeypatch):
    monkeypatch.setattr(fake_uia, "ControlFromPoint", lambda x, y: app["save"])
    run("snapshot", "--window", "Notepad")
    out = run("click", "name=Save", "--headed", "--no-verify")
    assert out["ok"] and out["via"] == "mouse:target"


def test_key_parsing():
    assert inputs.parse_keys("ctrl+shift+s") == [[fake_uia.Keys.VK_CONTROL,
                                                  fake_uia.Keys.VK_SHIFT, ord("S")]]
    assert len(inputs.parse_keys("alt h o r")) == 4
    with pytest.raises(WadError):
        inputs.parse_keys("ctrl+banana")


def test_unicode_key_events_cover_surrogates_and_newlines():
    ev = win32._key_events("a\n😀")
    assert ev[0] == (0, ord("a"), win32.KEYEVENTF_UNICODE)
    assert ev[2] == (win32.VK_RETURN, 0, 0)
    assert len(ev) == 2 + 2 + 4               # the emoji is two UTF-16 units


def test_errors_are_json_with_code_and_hint(app, run):
    out = run("click", "e999")
    assert out["ok"] is False and out["code"] == "NO_SNAPSHOT" and out["hint"]


def test_internal_errors_do_not_crash(app, run, monkeypatch):
    def boom(*a):
        raise RuntimeError("COM went away")
    monkeypatch.setattr(uia, "top_windows", boom)
    out = run("windows")
    assert out["code"] == "INTERNAL" and "COM went away" in out["message"]


def test_typing_verification_waits_for_a_slow_app(app, run, monkeypatch):
    """Real Notepad drew the last Arabic letters ~100 ms after SendInput returned; the
    first read must not be taken as the answer."""
    vp = app["liar"].GetPattern(fake_uia.PatternId.ValuePattern)
    pending, reads = [], [0]

    def slow_type(text, interval=0):
        vp._value += text[:3]
        pending.append(text[3:])

    real = type(vp).Value

    class Lagging(type(vp)):
        @property
        def Value(self):
            reads[0] += 1
            if pending and reads[0] > 3:
                self._value += pending.pop()
            return real.fget(self)

    vp.__class__ = Lagging
    monkeypatch.setattr(win32, "type_unicode", slow_type)
    run("snapshot", "--window", "Notepad")
    out = run("type", "name=Cell", "مرحبا بكم", "--keys")
    assert out["ok"] and out["verified"] is True


def test_select_in_an_msaa_only_combo_walks_with_the_keyboard(app, run):
    """Real WinForms (a ComboBox with an AccessibleName) exposes no options and no
    ExpandCollapse - found by the Windows CI run. wad reads the value while pressing Down."""
    run("snapshot", "--window", "Notepad")
    out = run("select", "role=ComboBox name=Color", "blue")
    assert out["ok"] and out["via"] == "keys" and out["option"] == "Blue"
    assert app["legacy"].Value == "Blue"
    bad = run("select", "role=ComboBox name=Color", "Purple")
    assert bad["code"] == "OPTION_NOT_FOUND" and "'Green'" in bad["hint"]


def test_expand_msaa_combo_with_keys_and_read_legacy_state(app, run):
    run("snapshot", "--window", "Notepad")
    assert run("get", "role=ComboBox name=Color")["expanded"] is False
    out = run("expand", "role=ComboBox name=Color")
    assert out["ok"] and out["via"] == "keys"
    assert run("get", "role=ComboBox name=Color")["expanded"] is True
    assert run("collapse", "role=ComboBox name=Color")["ok"]
    assert app["legacy"].State == 0x400


def test_watchdog_turns_a_frozen_app_into_app_hung(app, run, monkeypatch):
    """A UIA call into an app that stopped pumping messages never returns."""
    import threading
    from wadlib import cli
    from wadlib.registry import COMMANDS
    release = threading.Event()

    def frozen(args):
        release.wait(10)
        return {"ok": True}, "late"
    monkeypatch.setattr(cli, "WATCHDOG", 0.3)
    monkeypatch.setattr(COMMANDS["windows"], "fn", frozen)
    from wadlib.registry import COMMANDS as C
    payload, text = cli.execute_guarded("windows", C["windows"].namespace({}))
    release.set()
    assert payload["code"] == "APP_HUNG" and payload["hung"]


def test_window_reported_hung_by_windows_is_refused(app, run, monkeypatch):
    monkeypatch.setattr(win32, "is_hung", lambda hwnd: hwnd == 1001)
    out = run("snapshot", "--window", "Notepad")
    assert out["code"] == "APP_HUNG"
