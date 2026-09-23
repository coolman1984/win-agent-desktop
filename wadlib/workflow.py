"""Running many steps, replaying what worked, and checking the setup.

The agent-heavy way to automate is: explore with snapshots and refs, then keep what
worked. Every successful action is recorded with a durable selector, `trace --export`
turns a session into a batch file, and `batch` replays it deterministically - no model
needed until something changes."""
import json
import os
import platform
import re
import sys
import time

from . import __version__, state, win32
from .registry import COMMANDS, Arg, WadError, command

GUIDE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                     ".claude", "skills", "wad-desktop", "SKILL.md")
ENV_RE = re.compile(r"\$\{ENV:([A-Za-z_][A-Za-z0-9_]*)\}")


def _expand(v):
    if isinstance(v, str):
        def sub(m):
            if m.group(1) not in os.environ:
                raise WadError("USAGE", f"batch needs the environment variable {m.group(1)}",
                               "set it before running (secrets are never stored in batch files)")
            return os.environ[m.group(1)]
        return ENV_RE.sub(sub, v)
    return v


def load_steps(path):
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, ValueError) as e:
        raise WadError("USAGE", f"cannot read batch file: {e}") from None
    steps = data.get("steps") if isinstance(data, dict) else data
    if not isinstance(steps, list):
        raise WadError("USAGE", 'a batch file is a JSON list of steps (or {"steps": [...]})')
    return steps


@command("batch", "run a JSON list of steps in one go (a recorded or hand-written workflow)",
         Arg("file", positional=True, help='JSON: [{"cmd": "click", "target": "role=Button '
             'name=OK"}, {"sleep": 0.5}, ...]'),
         Arg("keep-going", bool, help="continue after a failed step"),
         group="workflow")
def cmd_batch(args):
    from .cli import execute
    steps = load_steps(args.file)
    results, failed = [], 0
    for i, step in enumerate(steps, 1):
        if "sleep" in step:
            time.sleep(float(step["sleep"]))
            continue
        name = step.get("cmd")
        if name not in COMMANDS or name in ("batch", "mcp"):
            raise WadError("USAGE", f"step {i}: unknown command {name!r}")
        values = {k: _expand(v) for k, v in step.items()
                  if k not in ("cmd", "retry", "optional", "note")}
        tries = 1 + int(step.get("retry", 0))
        for attempt in range(tries):
            ns = COMMANDS[name].namespace(values)
            payload, text = execute(name, ns)
            if payload.get("ok", True) or attempt == tries - 1:
                break
            time.sleep(0.8)
        good = payload.get("ok", True)
        results.append({"step": i, "cmd": name, "ok": good, "text": text,
                        **({} if good else {"code": payload.get("code")})})
        if not good and not step.get("optional"):
            failed += 1
            if not args.keep_going:
                break
    lines = [f"{'ok ' if r['ok'] else 'ERR'} {r['step']:>3} {r['cmd']:<12} "
             + (r["text"].splitlines() or [""])[0] for r in results]
    done = sum(r["ok"] for r in results)
    summary = f"{done}/{len(results)} steps ok" + (f", {failed} failed" if failed else "")
    payload = {"ok": failed == 0, "results": results, "summary": summary}
    if failed:
        payload.update(code="BATCH_FAILED", message=summary,
                       hint="snapshot the app at the failing step; fix that step's selector")
    return payload, "\n".join(lines + [summary])


REPLAY_SKIP = {"fn", "cmd", "json"}


def replay_record(name, ns, payload):
    """Called after every successful action: the step, with a durable selector in place of
    a snapshot ref, and secrets replaced by an environment-variable placeholder."""
    cmd = COMMANDS[name]
    if cmd.readonly or cmd.group in ("workflow",) or name in ("mcp",):
        return
    defaults = {a.dest: a.default for a in cmd.args}
    values = {k: v for k, v in vars(ns).items()
              if k not in REPLAY_SKIP and v not in (None, False) and v != defaults.get(k)}
    if payload.get("selector") and "target" in values:
        values["target"] = payload["selector"]
        if payload.get("window") and not values.get("window") and not values.get("hwnd"):
            values["window"] = payload["window"]
    values.pop("hwnd", None)
    if payload.get("secret") and "text" in values:
        values["text"] = "${ENV:WAD_SECRET}"
    state.trace("step", {"step": {"cmd": name, **values}})


@command("trace", "the action log: last steps, or --export a replayable batch file",
         Arg("last", int, default=20, help="how many entries"),
         Arg("export", help="write the recorded steps of the last --last entries to this "
             "batch file"),
         group="workflow", readonly=True)
def cmd_trace(args):
    entries = state.read_trace(args.last)
    if args.export:
        steps = [e["step"] for e in entries if e.get("cmd") == "step"]
        with open(args.export, "w", encoding="utf-8") as fh:
            json.dump({"steps": steps}, fh, ensure_ascii=False, indent=2)
        return ({"ok": True, "file": os.path.abspath(args.export), "steps": len(steps)},
                f"exported {len(steps)} steps to {args.export}")
    lines = []
    for e in entries:
        what = {k: v for k, v in e.items() if k not in ("t", "cmd", "session") and v is not None}
        lines.append(f"{e['t'][11:23]}  {e['cmd']:<12} {json.dumps(what, ensure_ascii=False)[:160]}")
    return {"ok": True, "entries": entries}, "\n".join(lines) or "(empty trace)"


@command("doctor", "check this PC: Windows, Python, packages, elevation, DPI, keyboard, policy",
         group="workflow", readonly=True)
def cmd_doctor(args):
    checks = []

    def add(name, good, detail):
        checks.append({"check": name, "status": good, "detail": detail})

    add("windows", "ok" if win32.IS_WINDOWS else "fail",
        platform.platform() if win32.IS_WINDOWS else "wad drives Windows apps; this is not Windows")
    add("python", "ok" if sys.version_info >= (3, 9) else "warn", sys.version.split()[0])
    for mod, need, why in (("uiautomation", "fail", "core"), ("win32com.client", "warn",
                           "excel-* / word-* (pip install pywin32)"),
                           ("PIL", "warn", "screen capture, --marks, --region (pip install Pillow)"),
                           ("winrt.windows.media.ocr", "warn",
                            "ocr / click-text (pip install -r requirements-ocr.txt)")):
        try:
            m = __import__(mod, fromlist=["_"])
            add(mod, "ok", getattr(m, "VERSION", getattr(m, "__version__", "installed")))
        except Exception:
            add(mod, need, f"missing - {why}")
    if win32.IS_WINDOWS:
        admin = win32.is_admin()
        add("elevation", "ok" if admin else "info",
            "running as administrator: can drive every app" if admin else
            "not administrator: apps running as administrator are invisible to wad (UIPI)")
        try:
            fg = win32.foreground()
            add("keyboard", "info", f"foreground window input language "
                f"{win32.input_language(fg):04x} (0409 = English US; key tips need English)")
        except Exception as e:
            add("keyboard", "warn", str(e))
    try:
        os.makedirs(state.STATE_DIR, exist_ok=True)
        add("state dir", "ok", state.STATE_DIR)
    except OSError as e:
        add("state dir", "fail", f"{state.STATE_DIR}: {e}")
    from .system import policy
    p = policy()
    add("system policy", "info", "shell/file/process commands ON (" +
        "; ".join(p["write_roots"]) + ")" if p.get("system") else
        "shell/file/process commands OFF (default)")
    icons = {"ok": "ok  ", "warn": "WARN", "fail": "FAIL", "info": "info"}
    text = f"wad {__version__}\n" + "\n".join(f"  {icons[c['status']]} {c['check']:<24} {c['detail']}"
                                              for c in checks)
    return {"ok": not any(c["status"] == "fail" for c in checks), "checks": checks}, text


@command("guide", "print the agent playbook (how to drive Windows apps well with wad)",
         group="workflow", readonly=True)
def cmd_guide(args):
    try:
        with open(GUIDE, encoding="utf-8") as fh:
            text = fh.read()
    except OSError:
        raise WadError("NOT_FOUND", f"guide not found at {GUIDE}") from None
    text = re.sub(r"\A---.*?---\s*", "", text, flags=re.S)       # drop skill front matter
    return {"ok": True, "guide": text}, text


@command("mcp", "serve every command as MCP tools over stdio (for Claude, Cursor, ...)",
         group="workflow", mcp=False)
def cmd_mcp(args):
    raise WadError("USAGE", "mcp is started from the command line: wad mcp")
