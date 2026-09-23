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


def close_without_saving():
    wad("close", "--window", ctx["title"])
    time.sleep(1.5)
    pop = wad("snapshot", "--window", ctx["title"], "--popup", ok=False)
    if not pop.get("ok"):             # newer Notepad asks inside its own window
        wad("snapshot", "--window", ctx["title"])
    wad("click", "name~=\"don't save\"", "--no-verify")
    wad("wait", "--window", ctx["title"], "--gone", "--timeout", "15")


def main():
    for name, fn in [("doctor", lambda: wad("doctor", ok=False)["checks"][0]["detail"]),
                     ("launch notepad", launch), ("snapshot", snapshot),
                     ("type unicode keys (verified)", type_unicode),
                     ("type via value (verified)", type_value),
                     ("screenshot --marks", screenshot), ("ocr finds text", ocr),
                     ("mcp server over stdio", mcp), ("trace export + batch replay", replay),
                     ("close without saving", close_without_saving)]:
        check(name, fn)
        if name == "launch notepad" and not results[-1][0]:
            break
    failed = [n for good, n in results if not good]
    print(f"\n{len(results) - len(failed)}/{len(results)} passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
