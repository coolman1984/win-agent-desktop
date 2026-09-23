"""Live demo of wad: drives Calculator and Notepad through UI Automation only.

    python demo.py            # watch it, with pauses between steps
    python demo.py --fast     # no pauses

Everything is found by name in the accessibility tree, the way an agent would -
no coordinates, no hard-coded refs. Your own Notepad tabs are never typed into:
the demo opens its own new tab and closes it with "Don't save".
"""
import argparse
import json
import os
import sys
import time

import uiautomation as auto

import wad

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "demo_output")
PAUSE = 1.2


def say(title):
    print(f"\n{'=' * 70}\n  {title}\n{'=' * 70}")
    time.sleep(PAUSE)


def run(*argv):
    shown = " ".join(a if " " not in a else f'"{a}"' for a in argv)
    print(f"\n> wad {shown}")
    code = wad.main(list(argv))
    time.sleep(PAUSE / 2)
    return code


def quiet_snapshot(window=None):
    """Refresh the stored snapshot without printing the whole tree."""
    argv = ["snapshot", "-i"] + (["--window", window] if window else [])
    with open(os.devnull, "w", encoding="utf-8") as sink:
        old, sys.stdout = sys.stdout, sink
        try:
            wad.main(argv)
        finally:
            sys.stdout = old


def ref_for(role, name, exact=True):
    """Look a control up by role and name in the latest snapshot and return its ref."""
    snap = wad.load_snapshot()
    for e in snap["elements"]:
        if e["role"].lower() == role.lower() and (
                e["name"] == name if exact else name.lower() in e["name"].lower()):
            return e["ref"]
    raise SystemExit(f"demo: no {role} named {name!r} in the snapshot")


def window_exists(fragment):
    return any(fragment.lower() in w.Name.lower() for w in wad.top_windows())


def calculator_part():
    say("PART 1 - Calculator: 125 x 8 - 250 = ? using only accessibility refs")
    had_calc = window_exists("Calculator")
    run("launch", "calculator")
    run("snapshot", "--window", "Calculator", "-i", "--depth", "3")
    print("  (skeleton first - cheap overview; now the full interactive tree)")
    quiet_snapshot("Calculator")

    print(f"  full snapshot stored: {len(wad.load_snapshot()['elements'])} elements with refs")

    run("click", ref_for("Button", "Clear"))
    for name in ("One", "Two", "Five", "Multiply by", "Eight", "Minus", "Two", "Five", "Zero",
                 "Equals"):
        run("click", ref_for("Button", name))
    display = ref_for("Text", "Display is", exact=False)
    run("get", display)
    os.makedirs(OUT, exist_ok=True)
    run("screenshot", "--window", "Calculator", os.path.join(OUT, "calculator.png"))
    if not had_calc:
        run("close", "--window", "Calculator")


def notepad_part():
    say("PART 2 - Notepad: new tab, type, read back, menu, close without saving")
    had_notepad = window_exists("Notepad")
    run("launch", "notepad", "--title", "Notepad")
    run("snapshot", "--window", "Notepad", "-i")
    run("click", ref_for("Button", "Add New Tab"))
    time.sleep(0.8)
    run("snapshot", "-i", "--depth", "4")
    quiet_snapshot()

    editor = ref_for("Document", "Text editor")
    run("type", editor, "Hello from wad - Windows Agent Desktop.")
    run("type", editor, " This part is real keystrokes {with braces}.", "--keys", "--append")
    run("get", editor)
    os.makedirs(OUT, exist_ok=True)
    run("screenshot", os.path.join(OUT, "notepad.png"))

    say("Open the File menu and close the demo tab through it")
    run("snapshot", "-i")
    run("click", ref_for("MenuItem", "File"))
    time.sleep(0.6)
    quiet_snapshot()

    run("find", "--role", "MenuItem")
    run("click", ref_for("MenuItem", "Close tab"))
    time.sleep(0.7)
    quiet_snapshot()

    run("find", "--role", "Button", "--name", "save")
    run("click", ref_for("Button", "Don't save"))
    time.sleep(0.6)
    run("snapshot", "-i", "--depth", "6")
    if not had_notepad:
        run("close")


def main():
    global PAUSE
    ap = argparse.ArgumentParser()
    ap.add_argument("--fast", action="store_true")
    args = ap.parse_args()
    if args.fast:
        PAUSE = 0
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    trace_start = os.path.getsize(wad.TRACE_FILE) if os.path.exists(wad.TRACE_FILE) else 0
    with auto.UIAutomationInitializerInThread():
        say("wad - Windows Agent Desktop (ideas from lahfir/agent-desktop, on UI Automation)")
        run("windows")
        calculator_part()
        notepad_part()
    say("Trace of this demo (every action is logged as JSON)")
    with open(wad.TRACE_FILE, encoding="utf-8") as fh:
        fh.seek(trace_start)
        for line in fh:
            e = json.loads(line)
            what = {k: v for k, v in e.items() if k not in ("t", "cmd")}
            print(f"  {e['t'][11:23]}  {e['cmd']:<10} {what}")
    print(f"\n  screenshots: {OUT}\n  done.")


if __name__ == "__main__":
    main()
