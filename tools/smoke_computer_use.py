"""Exercise wad's MCP computer-control path on a real interactive Windows desktop.

Starts only the repository's small WinForms test app, types a harmless value through
MCP, submits it, reads the result, and closes that test app. No user document is used.

    python tools/smoke_computer_use.py
"""
import json
import argparse
import os
import subprocess
import sys
import time


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WAD = [sys.executable, os.path.join(ROOT, "wad.py")]
APP = os.path.join(ROOT, "tools", "test_app.ps1")
TITLE = "wad test app"


def exchange(method, params=None):
    messages = [
        {"jsonrpc": "2.0", "id": 1, "method": "initialize",
         "params": {"protocolVersion": "2025-06-18"}},
        {"jsonrpc": "2.0", "id": 2, "method": method, "params": params or {}},
    ]
    request = "\n".join(json.dumps(msg) for msg in messages) + "\n"
    proc = subprocess.run(WAD + ["mcp"], input=request, capture_output=True,
                          text=True, encoding="utf-8", timeout=45)
    replies = [json.loads(line) for line in proc.stdout.splitlines()]
    if proc.returncode or [reply.get("id") for reply in replies] != [1, 2]:
        raise AssertionError(f"MCP protocol failure: {proc.stdout!r} {proc.stderr!r}")
    return replies[1]["result"]


def tool(name, **arguments):
    result = exchange("tools/call", {"name": name, "arguments": {**arguments, "json": True}})
    body = json.loads(result["content"][0]["text"])
    if result.get("isError") or not body.get("ok", True):
        raise AssertionError(f"{name} failed: {body}")
    return body


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--ocr", action="store_true", help="also verify Windows OCR on the test app")
    args = parser.parse_args()
    listed = {entry["name"] for entry in exchange("tools/list")["tools"]}
    assert {"desktop-check", "windows", "inspect", "snapshot", "type", "get", "click", "close"} <= listed
    desktop = tool("desktop-check")
    assert desktop["accessible"] is True, desktop

    before = {w["hwnd"] for w in tool("windows")["windows"]}
    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    app = subprocess.Popen(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                            "-STA", "-File", APP], stdin=subprocess.DEVNULL,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           creationflags=flags)
    hwnd = None
    try:
        deadline = time.time() + 20
        while time.time() < deadline:
            matches = [w for w in tool("windows")["windows"]
                       if w["hwnd"] not in before and w["title"] == TITLE]
            if len(matches) == 1:
                hwnd = matches[0]["hwnd"]
                break
            time.sleep(0.3)
        assert hwnd is not None, "test app did not open a new window"

        snap = tool("snapshot", hwnd=hwnd, interactive=True)
        assert any(e["aid"] == "nameBox" for e in snap["elements"]), snap
        typed = tool("type", target="aid=nameBox", text="wad MCP live test", hwnd=hwnd)
        assert typed["verified"] is True, typed
        read = tool("get", target="aid=nameBox", hwnd=hwnd)
        assert read.get("value") == "wad MCP live test", read
        clicked = tool("click", target="role=Button name=Submit", hwnd=hwnd,
                       expect="Submitted: wad MCP live test")
        assert clicked["verified"] is True, clicked
        inspected = tool("inspect", hwnd=hwnd, visual=args.ocr)
        assert inspected["accessible_count"] > 0, inspected
        assert any(e["aid"] == "nameBox" and e.get("value") == "wad MCP live test"
                   for e in inspected["elements"]), "inspect did not read the test field"
        if args.ocr:
            assert inspected["coverage"]["ocr"] == "available", inspected["warnings"]
            assert any("Submitted:" in line["text"] and "wad" in line["text"]
                       for line in inspected["ocr_lines"]), "OCR missed the result label"
        print(f"PASS: MCP connected, saw {desktop['visible_windows']} windows, typed and "
              "read back text, clicked Submit, inspected the result" +
              (" with OCR" if args.ocr else "") + ", and verified it")
    finally:
        if hwnd is not None:
            try:
                tool("close", hwnd=hwnd)
            except Exception as exc:
                print(f"Could not close test window through MCP: {exc}", file=sys.stderr)
        try:
            app.wait(timeout=5)
        except subprocess.TimeoutExpired:
            app.terminate()  # only the test process started above
            app.wait(timeout=5)


if __name__ == "__main__":
    main()
