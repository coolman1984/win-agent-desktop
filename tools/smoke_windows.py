"""End-to-end check on a REAL Windows desktop (CI runs it on windows-latest).

    python tools/smoke_windows.py

Drives the real Notepad through the real command line - launch, snapshot, selectors,
Unicode typing, verification, screenshot, OCR, the MCP server, trace export and replay,
and closing without saving. Prints one line per check and exits non-zero on failure."""
import json
import os
import subprocess
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WAD = [sys.executable, os.path.join(ROOT, "wad.py")]
OUT = os.path.join(ROOT, "demo_output")
results = []


def wad(*argv, ok=True, timeout=90):
    p = subprocess.run(WAD + ["--json", *argv], capture_output=True, text=True,
                       encoding="utf-8", timeout=timeout)
    try:
        payload = json.loads(p.stdout)
    except ValueError:
        raise AssertionError(f"wad {' '.join(argv)} printed non-JSON:\n{p.stdout}\n{p.stderr}")
    if ok and not payload.get("ok", True):
        raise AssertionError(f"wad {' '.join(argv)} failed: {payload}")
    return payload


def check(name, fn):
    started = time.time()
    try:
        detail = fn() or ""
        results.append((True, name))
        print(f"PASS {name} ({time.time() - started:.1f}s) {detail}", flush=True)
    except Exception as e:
        results.append((False, name))
        print(f"FAIL {name}: {e}", flush=True)


ctx = {}


def launch():
    out = wad("launch", "notepad.exe", "--title", "Notepad", "--timeout", "30")
    ctx["title"] = out["title"]
    return out["title"]


def snapshot():
    out = wad("snapshot", "--window", ctx["title"], "-i")
    roles = {e["role"] for e in out["elements"]}
    assert {"Document", "Edit"} & roles, f"no text area in {sorted(roles)}"
    ctx["editor"] = 'name~="text editor"'
    return f"{len(out['elements'])} elements"


def type_unicode():
    text = "hello wad - مرحبا"
    out = wad("type", ctx["editor"], text, "--keys", "--window", ctx["title"])
    assert out["verified"] is True, out
    got = wad("get", ctx["editor"], "--window", ctx["title"])
    content = got.get("value") or got.get("text") or ""
    assert text in content, f"editor holds {content!r}"
    return out["via"]


def type_value():
    out = wad("type", ctx["editor"], "HELLO OCR 12345", "--window", ctx["title"])
    assert out["verified"] is True, out
    return out["via"]


def screenshot():
    os.makedirs(OUT, exist_ok=True)
    out = wad("screenshot", os.path.join(OUT, "smoke-notepad.png"), "--window", ctx["title"],
              "--marks")
    assert os.path.getsize(out["path"]) > 1000
    return f"{out['size']}"


def ocr():
    out = wad("ocr", "--window", ctx["title"], "--find", "12345", "--lang", "en")
    assert out["matches"], "OCR did not find 12345 on screen"
    return f"at {out['matches'][0]['center']}"


def mcp():
    req = "\n".join(json.dumps(m) for m in [
        {"jsonrpc": "2.0", "id": 1, "method": "initialize",
         "params": {"protocolVersion": "2025-06-18"}},
        {"jsonrpc": "2.0", "method": "notifications/initialized"},
        {"jsonrpc": "2.0", "id": 2, "method": "tools/list"},
        {"jsonrpc": "2.0", "id": 3, "method": "tools/call",
         "params": {"name": "get", "arguments": {"target": ctx["editor"],
                                                 "window": ctx["title"]}}}]) + "\n"
    p = subprocess.run(WAD + ["mcp"], input=req, capture_output=True, text=True,
                       encoding="utf-8", timeout=90)
    msgs = [json.loads(line) for line in p.stdout.splitlines()]
    assert [m["id"] for m in msgs] == [1, 2, 3], p.stdout + p.stderr
    assert "HELLO OCR" in msgs[2]["result"]["content"][0]["text"]
    return f"{len(msgs[1]['result']['tools'])} tools"


def replay():
    flow = os.path.join(tempfile.mkdtemp(), "flow.json")
    wad("trace", "--last", "40", "--export", flow)
    with open(flow, encoding="utf-8") as fh:
        steps = json.load(fh)["steps"]
    steps = [s for s in steps if s["cmd"] == "type"][-1:]
    assert steps, "no type step recorded"
    with open(flow, "w", encoding="utf-8") as fh:
        json.dump(steps, fh)
    out = wad("batch", flow)
    return out["summary"]


def window_ops():
    wad("window", "maximize", "--window", ctx["title"])
    wad("window", "restore", "--window", ctx["title"])
    out = wad("window", "move", "--window", ctx["title"], "--x", "40", "--y", "40")
    assert out["rect"][:2] == [40, 40], out["rect"]
    return f"{out['rect']}"


def headed_click():
    out = wad("click", ctx["editor"], "--headed", "--window", ctx["title"], "--no-verify")
    assert out["via"].startswith("mouse:"), out
    return out["via"]


def context_menu():
    wad("right-click", ctx["editor"], "--window", ctx["title"], "--no-verify")
    time.sleep(0.8)
    pop = wad("snapshot", "--window", ctx["title"], "--popup")
    names = {e["name"] for e in pop["elements"]}
    wad("press", "esc", "--window", ctx["title"])
    assert any("select all" in n.lower() or "undo" in n.lower() for n in names), sorted(names)
    return f"{len(names)} names in the menu"


def press_and_clear():
    wad("press", "ctrl+end", "--window", ctx["title"])
    out = wad("clear", ctx["editor"], "--window", ctx["title"])
    got = wad("get", ctx["editor"], "--window", ctx["title"])
    assert not (got.get("value") or got.get("text") or "").strip(), got
    wad("type", ctx["editor"], "line one\nline two\nline three", "--window", ctx["title"])
    return out["via"]


def dialog_controls():
    """The Replace dialog (classic Notepad): typing into a dialog, a checkbox, closing."""
    wad("press", "ctrl+h", "--window", ctx["title"])
    time.sleep(1.0)
    pop = wad("snapshot", "--window", ctx["title"], "--popup", ok=False)
    if not pop.get("ok") or not any(e["role"] == "CheckBox" for e in pop["elements"]):
        wad("press", "esc", "--window", ctx["title"])
        return "no Replace dialog in this Notepad (skipped)"
    box = next(e["name"] for e in pop["elements"] if e["role"] == "CheckBox")
    sel = f'role=CheckBox name="{box}"'
    assert wad("check", sel)["state"] == "on"
    assert wad("check", sel)["state"] == "on"
    assert wad("uncheck", sel)["state"] == "off"
    edit = next(e for e in pop["elements"] if e["role"] == "Edit")
    wad("type", edit["ref"], "two")
    got = wad("get", edit["ref"])
    assert got.get("value") == "two", got
    wad("click", "role=Button name=Cancel", "--expect-gone", "Match case", "--timeout", "5",
        ok=False)
    wad("snapshot", "--window", ctx["title"])
    return f"checkbox {box!r} + edit ok"


def wait_and_find():
    wad("snapshot", "--window", ctx["title"], "-i")
    out = wad("find", "--role", "MenuItem")
    wad("wait", "--window", ctx["title"], "--target", ctx["editor"], "--timeout", "5")
    return f"{len(out['matches'])} menu items"


def close_without_saving():
    wad("close", "--window", ctx["title"])
    time.sleep(1.5)
    pop = wad("snapshot", "--window", ctx["title"], "--popup", ok=False)
    if not pop.get("ok"):             # newer Notepad asks inside its own window
        wad("snapshot", "--window", ctx["title"])
    wad("click", "name~=\"don't save\"", "--no-verify")
    wad("wait", "--window", ctx["title"], "--gone", "--timeout", "15")


# --- a real Windows Forms app: the controls Notepad lacks --------------------------
APP = "wad test app"


def forms_launch():
    """Launched while `wad watch` listens: the window opening must arrive as a UIA event."""
    ctx["watch"] = subprocess.Popen(WAD + ["--json", "watch", "--seconds", "45", "--until", APP,
                                           "--no-focus"], stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE, text=True, encoding="utf-8")
    time.sleep(2.5)                        # let the watcher register its handlers
    ps1 = os.path.join(ROOT, "tools", "test_app.ps1")
    ctx["forms"] = subprocess.Popen(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                                     "-STA", "-File", ps1])
    out = wad("wait", "--window", APP, "--timeout", "40")
    wad("snapshot", "--window", APP, "-i")
    return out["title"]


def forms_watch():
    stdout, _ = ctx["watch"].communicate(timeout=60)
    got = json.loads(stdout)
    assert got["mode"] == "events", got
    assert any(e["kind"] == "window opened" and APP in e["name"] for e in got["events"]), got
    return f"{len(got['events'])} events, mode {got['mode']}"


def forms_type():
    out = wad("type", "aid=nameBox", "Ahmed أحمد", "--window", APP)
    assert out["verified"] is True, out
    return out["via"]


def forms_check():
    assert wad("check", "role=CheckBox name=Subscribe", "--window", APP)["state"] == "on"
    assert wad("uncheck", "role=CheckBox name=Subscribe", "--window", APP)["state"] == "off"
    assert wad("check", "role=CheckBox name=Subscribe", "--window", APP)["state"] == "on"


def forms_select():
    out = wad("select", "role=ComboBox name=Color", "blue", "--window", APP)
    assert out["option"] == "Blue", out
    return f"value {out['value']!r}"


def forms_expand():
    wad("expand", "role=ComboBox name=Color", "--window", APP)
    got = wad("get", "role=ComboBox name=Color", "--window", APP)
    wad("collapse", "role=ComboBox name=Color", "--window", APP)
    assert got.get("expanded") is True, got
    return "expanded then collapsed"


def forms_scroll():
    out = wad("scroll", "role=List name=Numbers", "--amount", "10", "--window", APP)
    assert out["moved"], out
    return f"{out.get('start')} -> {out.get('end')}"


def forms_disabled():
    out = wad("click", "role=Button name=Disabled", "--window", APP, ok=False)
    assert out["code"] == "NOT_ENABLED", out


def forms_submit():
    out = wad("click", "role=Button name=Submit", "--window", APP,
              "--expect", "Submitted: Ahmed أحمد / Blue / True")
    return "; ".join(out["changes"][:2])


def forms_hover_and_ocr_click():
    wad("hover", "role=Button name=Submit", "--window", APP)
    wad("type", "aid=nameBox", "Second", "--window", APP)
    seen = wad("ocr", "--window", APP, "--find", "Submit", "--lang", "en")["matches"]
    out = wad("click-text", "Submit", "--window", APP, "--expect", "Submitted: Second",
              "--lang", "en", ok=False)
    assert out.get("ok"), f"{out} - OCR saw {seen}"
    return f"clicked {out['text']!r} at ({out['screen_x']}, {out['screen_y']})"


def forms_record():
    """A 'person' (real, injected input) fills the form while `wad record` watches; the
    recording must replay on its own."""
    flow = os.path.join(tempfile.mkdtemp(), "recorded.json")
    rec = subprocess.Popen(WAD + ["--json", "record", flow, "--window", APP, "--seconds", "40"],
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
                           encoding="utf-8")
    time.sleep(3)                          # hooks installed
    wad("click", "aid=nameBox", "--headed", "--window", APP, "--no-verify")
    wad("type", "aid=nameBox", "Recorded", "--keys", "--window", APP)
    wad("click", "role=Button name=Submit", "--headed", "--window", APP, "--no-verify")
    time.sleep(1.5)
    wad("record-stop")
    raw, err = rec.communicate(timeout=60)
    out = json.loads(raw) if raw.strip().startswith("{") else {"raw": raw, "stderr": err}
    assert "steps" in out, out
    steps = out["steps"]
    typed = [st for st in steps if st["cmd"] == "type" and st.get("text") == "Recorded"]
    clicked = [st for st in steps if st["cmd"] == "click" and "Submit" in st["target"]]
    assert typed and clicked, steps
    wad("type", "aid=nameBox", "something else", "--window", APP)
    replay = wad("batch", flow)
    wad("wait", "--window", APP, "--name", "Submitted: Recorded", "--timeout", "10")
    return f"{len(steps)} steps recorded; replay {replay['summary']}"


def forms_report():
    wad("report", "clear")
    wad("report", "on")
    try:
        wad("click", "role=CheckBox name=Subscribe", "--window", APP)
        wad("type", "aid=nameBox", "Report", "--window", APP)
    finally:
        wad("report", "off")
    out = wad("report", "build", "--out", os.path.join(OUT, "smoke-report.html"))
    size = os.path.getsize(out["file"])
    assert out["steps"] == 2 and size > 5000, (out, size)
    return f"{out['steps']} steps, {size // 1024} KB"


def forms_detect_quick():
    """No vision server on the runner: detect must fall back to OCR words, clickable."""
    out = wad("detect", "--window", APP, "--marks")
    assert out["mode"] == "ocr", out
    ref = next(e["ref"] for e in out["elements"] if e["content"] == "Submit")
    wad("type", "aid=nameBox", "Eye", "--window", APP)
    wad("click-mark", ref, "--expect", "Submitted: Eye")
    return f"{len(out['elements'])} words, clicked {ref}"


def forms_close():
    wad("close", "--window", APP)
    wad("wait", "--window", APP, "--gone", "--timeout", "15")
    ctx["forms"].wait(timeout=15)


# --- the browser, from the inside --------------------------------------------------
def browser(name):
    def run():
        page = "file:///" + os.path.join(ROOT, "tools", "test_page.html").replace("\\", "/")
        port = "9444" if name == "edge" else "9445"
        wad("browser-launch", page, "--browser", name, "--port", port)
        try:
            snap = wad("browser-snapshot")
            assert any(e["name"] == "Email" for e in snap["elements"]), snap["elements"]
            wad("browser-type", "#email", "ahmed@example.com")
            wad("browser-click", "text=I agree")
            wad("browser-click", "text=Send", "--expect", "Hello ahmed@example.com / true")
            assert wad("browser-click", "text=Locked", ok=False)["code"] == "NOT_ENABLED"
            wad("browser-screenshot", os.path.join(OUT, f"smoke-{name}.png"))
        finally:
            wad("browser-close", ok=False)
        return f"{len(snap['elements'])} elements"
    return run


def main():
    for name, fn in [("doctor", lambda: wad("doctor", ok=False)["checks"][0]["detail"]),
                     ("launch notepad", launch), ("snapshot", snapshot),
                     ("type unicode keys (verified)", type_unicode),
                     ("type via value (verified)", type_value),
                     ("screenshot --marks", screenshot), ("ocr finds text", ocr),
                     ("mcp server over stdio", mcp), ("trace export + batch replay", replay),
                     ("window maximize/restore/move", window_ops),
                     ("headed click lands on target", headed_click),
                     ("right-click + popup snapshot", context_menu),
                     ("press + clear (verified)", press_and_clear),
                     ("dialog: check/uncheck/type", dialog_controls),
                     ("find + wait --target", wait_and_find),
                     ("close without saving", close_without_saving),
                     ("forms: launch test app", forms_launch),
                     ("watch: window opened arrives as an event", forms_watch),
                     ("forms: type (Unicode, verified)", forms_type),
                     ("forms: check / uncheck", forms_check),
                     ("forms: select in combo box", forms_select),
                     ("forms: expand / collapse", forms_expand),
                     ("forms: scroll a list", forms_scroll),
                     ("forms: disabled button refused", forms_disabled),
                     ("forms: submit --expect result", forms_submit),
                     ("forms: hover + OCR click-text", forms_hover_and_ocr_click),
                     ("record a person, replay it", forms_record),
                     ("visual step report (on, build, off)", forms_report),
                     ("smart eye quick mode (OCR) + click-mark", forms_detect_quick),
                     ("forms: close", forms_close),
                     ("browser: Edge via DevTools", browser("edge")),
                     ("browser: Chrome via DevTools", browser("chrome"))]:
        check(name, fn)
        if name == "forms: launch test app" and not results[-1][0]:
            break
    failed = [n for good, n in results if not good]
    print(f"\n{len(results) - len(failed)}/{len(results)} passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
