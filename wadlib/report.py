"""The visual step report: a picture of the app after every action, as one HTML page.

OFF by default - screenshots cost time and disk, and they capture whatever is on screen.
A person turns it on (`wad report on`, or WAD_REPORT=1 for one terminal), runs the task,
then `wad report build` writes a single self-contained page: every step with its command,
its result and what the app looked like right after it."""
import base64
import html
import json
import os
import shutil

from . import state, uia
from .registry import Arg, WadError, command

SKIP = {"report", "trace", "guide", "doctor", "mcp", "batch", "record", "watch"}
MAX_WIDTH = 1000


def report_dir():
    return os.path.join(state.STATE_DIR, "report", state.SESSION or "default")


def enabled():
    if os.environ.get("WAD_REPORT") == "1":
        return True
    if os.environ.get("WAD_REPORT") == "0":
        return False
    cfg = state.read_json(state.CONFIG_FILE, None) or {}
    return bool(cfg.get("report"))


def _window_rect(payload):
    """The window the step acted on: the one the payload names, else the last snapshot's."""
    hwnd = payload.get("hwnd")
    if not hwnd and payload.get("window"):
        try:
            hwnd = uia.find_window(payload["window"]).NativeWindowHandle
        except WadError:
            hwnd = None
    if not hwnd:
        snap = state.read_json(state.SNAPSHOT_FILE, None)
        hwnd = snap and snap.get("hwnd")
    if not hwnd:
        return None
    import uiautomation as auto
    w = auto.ControlFromHandle(hwnd)
    if not w or not uia.alive(w):
        return None
    return uia.rect_of(w)


def capture_step(name, ns, payload, text):
    """Called after every command when the report is on; never fails the command."""
    if name in SKIP or not enabled():
        return
    try:
        from . import registry
        if registry.COMMANDS[name].readonly:
            return
        folder = report_dir()
        os.makedirs(folder, exist_ok=True)
        n = len([f for f in os.listdir(folder) if f.endswith(".json")]) + 1
        image = None
        rect = _window_rect(payload)
        if rect and rect[2] > rect[0] and rect[3] > rect[1]:
            try:
                from PIL import ImageGrab
                img = ImageGrab.grab(bbox=tuple(rect), all_screens=True)
                if img.width > MAX_WIDTH:
                    img = img.resize((MAX_WIDTH, round(img.height * MAX_WIDTH / img.width)))
                image = f"{n:04d}.jpg"
                img.convert("RGB").save(os.path.join(folder, image), quality=70)
            except Exception:
                image = None
        args = {k: v for k, v in vars(ns).items()
                if k not in ("fn", "cmd", "json") and not k.startswith("_")
                and v not in (None, False)}
        if payload.get("secret") and "text" in args:
            args["text"] = "********"
        state.write_json(os.path.join(folder, f"{n:04d}.json"), {
            "n": n, "t": state.now_iso(), "cmd": name, "args": args,
            "ok": bool(payload.get("ok", True)), "text": text, "image": image})
    except Exception:
        pass


def _steps():
    folder = report_dir()
    if not os.path.isdir(folder):
        return []
    out = []
    for f in sorted(os.listdir(folder)):
        if f.endswith(".json"):
            try:
                with open(os.path.join(folder, f), encoding="utf-8") as fh:
                    out.append(json.load(fh))
            except (OSError, ValueError):
                continue
    return out


PAGE = """<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>wad step report</title><style>
:root{--bg:#f6f7f9;--card:#fff;--ink:#1c1f24;--dim:#5d6470;--ok:#1d7f4e;--bad:#b3261e;--line:#dde1e7}
@media (prefers-color-scheme:dark){:root{--bg:#15171b;--card:#1e2126;--ink:#e8eaed;--dim:#9aa0a8;--ok:#5ccf92;--bad:#ff8a80;--line:#30343b}}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 system-ui,Segoe UI,sans-serif}
main{max-width:1060px;margin:0 auto;padding:24px 16px}h1{font-size:22px;margin:0 0 4px}
.sum{color:var(--dim);margin-bottom:20px}.step{background:var(--card);border:1px solid var(--line);
border-left:5px solid var(--ok);border-radius:10px;padding:14px 16px;margin:14px 0}
.step.bad{border-left-color:var(--bad)}.head{display:flex;gap:10px;align-items:baseline;flex-wrap:wrap}
.n{font-weight:700}.cmd{font-family:Consolas,monospace;font-size:14px;overflow-wrap:anywhere}
.t{color:var(--dim);font-size:13px;margin-left:auto}pre{white-space:pre-wrap;overflow-wrap:anywhere;
margin:8px 0;font:13px/1.45 Consolas,monospace;color:var(--dim)}img{max-width:100%;border:1px solid var(--line);
border-radius:6px;margin-top:6px}</style></head><body><main>
<h1>wad step report</h1><div class="sum">{summary}</div>{steps}</main></body></html>"""


def build(out):
    steps = _steps()
    folder = report_dir()
    cards = []
    for s in steps:
        arg_text = " ".join(f"--{k.replace('_', '-')} {v}" if v is not True else f"--{k}"
                            for k, v in s["args"].items())
        img = ""
        if s.get("image"):
            try:
                with open(os.path.join(folder, s["image"]), "rb") as fh:
                    data = base64.b64encode(fh.read()).decode("ascii")
                img = f'<img alt="the app after step {s["n"]}" src="data:image/jpeg;base64,{data}">'
            except OSError:
                pass
        cards.append(
            f'<section class="step{"" if s["ok"] else " bad"}"><div class="head">'
            f'<span class="n">{s["n"]}</span><span class="cmd">{html.escape(s["cmd"])} '
            f'{html.escape(arg_text)}</span><span class="t">{html.escape(s["t"][11:19])}</span>'
            f'</div><pre>{html.escape(s["text"])}</pre>{img}</section>')
    bad = sum(1 for s in steps if not s["ok"])
    summary = (f"{len(steps)} steps, {bad} failed - "
               f"{html.escape(steps[0]['t'][:19]) if steps else ''}")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(PAGE.replace("{summary}", summary).replace("{steps}", "\n".join(cards)))
    return len(steps)


@command("report", "the visual step report (OFF by default): on | off | status | build | clear",
         Arg("action", positional=True, choices=["on", "off", "status", "build", "clear"]),
         Arg("out", help="build: the HTML file to write (default: in the report folder)"),
         group="workflow")
def cmd_report(args):
    cfg = state.read_json(state.CONFIG_FILE, None) or {}
    if args.action in ("on", "off"):
        cfg["report"] = args.action == "on"
        state.write_json(state.CONFIG_FILE, cfg)
        word = "ON: every action will be photographed" if cfg["report"] else "OFF"
        return {"ok": True, "report": cfg["report"]}, f"step report {word}"
    if args.action == "clear":
        shutil.rmtree(report_dir(), ignore_errors=True)
        return {"ok": True}, "step report cleared"
    if args.action == "build":
        out = os.path.abspath(args.out or os.path.join(report_dir(), "report.html"))
        os.makedirs(os.path.dirname(out), exist_ok=True)
        n = build(out)
        if not n:
            raise WadError("NOT_FOUND", "no recorded steps to report",
                           "turn it on with `wad report on`, run the task, then build")
        return {"ok": True, "file": out, "steps": n}, f"wrote {out} ({n} steps)"
    n = len(_steps())
    on = enabled()
    return ({"ok": True, "report": on, "steps": n, "folder": report_dir()},
            f"step report {'ON' if on else 'OFF'} - {n} steps recorded in {report_dir()}")

