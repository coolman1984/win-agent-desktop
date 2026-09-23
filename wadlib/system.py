"""Powers beyond the screen - shell, files, processes - OFF until a person turns them on.

Clicking through an app is limited by what the app allows; a shell is not. So these
commands obey a policy file the agent cannot change through wad itself:

    %LOCALAPPDATA%\\win-agent-desktop\\policy.json
    {"system": true,                          # allow shell / file / process commands
     "write_roots": ["C:\\\\Users\\\\me\\\\Documents\\\\wad"],   # the only places writes may go
     "allow_kill": false}                     # process-kill

(or WAD_ALLOW_SYSTEM=1 for one terminal). Some commands are refused even then: this is a
seat belt against an agent's mistakes, not a sandbox against a hostile one - run agents
you do not trust in a separate Windows account or VM."""
import os
import re
import subprocess

from . import state
from .registry import Arg, WadError, command

DENY = [
    (r"\bformat-volume\b|\bformat(\.com)?\s+[a-z]:", "formatting a drive"),
    (r"\b(diskpart|bcdedit|vssadmin|wbadmin|cipher\s+/w)\b", "disk / boot / backup tools"),
    (r"\bremove-item\b.*-recurse.*\b[a-z]:\\?\s*($|[\"'])", "deleting a whole drive"),
    (r"\b(rm|del|rd|rmdir|remove-item)\b.*\b(c:\\windows|c:\\program files|system32)",
     "deleting system files"),
    (r"\breg(\.exe)?\s+(delete|add)\s+hklm", "changing machine-wide registry"),
    (r"\b(remove-itemproperty|set-itemproperty|new-itemproperty)\b.*hklm:", "changing HKLM"),
    (r"\b(shutdown|restart-computer|stop-computer)\b", "shutting the PC down"),
    (r"\b(set-mppreference|add-mppreference)\b", "changing Defender"),
    (r"\bnet\s+user\b.*(/add|/delete)|\b(new|remove)-localuser\b", "creating/deleting users"),
    (r"\b(invoke-webrequest|iwr|curl|wget|invoke-restmethod|irm)\b.*\|\s*(iex|invoke-expression)",
     "running code downloaded from the internet"),
    (r"\bset-executionpolicy\b", "changing the script execution policy"),
    (r"\bclear-(eventlog|recyclebin)\b|\bwevtutil\s+cl\b", "wiping logs / recycle bin"),
    (r"win-agent-desktop[\\/]+(policy\.json|trace\.jsonl)", "changing wad's own policy or trace"),
]
CRITICAL = {"csrss.exe", "winlogon.exe", "lsass.exe", "services.exe", "smss.exe", "wininit.exe",
            "svchost.exe", "system", "dwm.exe", "explorer.exe", "fontdrvhost.exe"}
MAX_OUT = 20000


def policy():
    p = state.read_json(state.POLICY_FILE, None) or {}
    if os.environ.get("WAD_ALLOW_SYSTEM") == "1":
        p = {**p, "system": True}
    roots = p.get("write_roots") or [os.path.join(os.path.expanduser("~"), "Documents", "wad")]
    p["write_roots"] = [os.path.normcase(os.path.abspath(os.path.expandvars(r))) for r in roots]
    return p


def require_system():
    p = policy()
    if not p.get("system"):
        raise WadError("POLICY_DENIED", "system commands are off",
                       f"a person can enable them in {state.POLICY_FILE} "
                       '({"system": true}) or with WAD_ALLOW_SYSTEM=1')
    return p


def denied(cmdline):
    low = cmdline.lower()
    for pattern, why in DENY:
        if re.search(pattern, low):
            return why
    return None


def inside(path, roots):
    full = os.path.normcase(os.path.realpath(os.path.abspath(path)))
    return any(full == r or full.startswith(r.rstrip("\\/") + os.sep) for r in roots)


@command("shell", "run a PowerShell command and return its output (policy-gated; dangerous "
         "commands always refused)",
         Arg("command", positional=True), Arg("timeout", float, default=60.0),
         Arg("cwd", help="working directory"),
         group="system")
def cmd_shell(args):
    require_system()
    why = denied(args.command)
    if why:
        state.trace("shell-refused", {"command": args.command, "why": why})
        raise WadError("POLICY_DENIED", f"refused: {why}",
                       "this is never allowed through wad; a person must do it by hand")
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command",
                            args.command], capture_output=True, text=True,
                           timeout=args.timeout, cwd=args.cwd, encoding="utf-8",
                           errors="replace")
    except subprocess.TimeoutExpired:
        raise WadError("TIMEOUT", f"the command ran over {args.timeout}s and was stopped") from None
    except FileNotFoundError:
        raise WadError("NOT_SUPPORTED", "PowerShell is not available here") from None
    out, err = r.stdout[-MAX_OUT:], r.stderr[-MAX_OUT:]
    state.trace("shell", {"command": args.command, "exit": r.returncode})
    text = out + (("\n[stderr]\n" + err) if err.strip() else "") + f"\n[exit {r.returncode}]"
    return {"ok": r.returncode == 0, "exit": r.returncode, "stdout": out, "stderr": err}, text


@command("file-read", "read a text file (policy-gated)", Arg("path", positional=True),
         Arg("max-chars", int, default=100000), group="system", readonly=True)
def cmd_file_read(args):
    require_system()
    try:
        with open(args.path, encoding="utf-8", errors="replace") as fh:
            text = fh.read(args.max_chars + 1)
    except OSError as e:
        raise WadError("FILE_ERROR", str(e)) from None
    cut = len(text) > args.max_chars
    text = text[:args.max_chars]
    return {"ok": True, "path": args.path, "text": text, "truncated": cut}, text + (
        "\n[truncated]" if cut else "")


@command("file-write", "write a text file inside the policy's write_roots (read back)",
         Arg("path", positional=True), Arg("text", positional=True),
         Arg("append", bool), group="system")
def cmd_file_write(args):
    p = require_system()
    if not inside(args.path, p["write_roots"]) or inside(args.path, [os.path.normcase(
            os.path.abspath(state.STATE_DIR))]):
        raise WadError("POLICY_DENIED", f"{args.path} is outside the allowed folders",
                       "allowed: " + "; ".join(p["write_roots"]))
    os.makedirs(os.path.dirname(os.path.abspath(args.path)), exist_ok=True)
    with open(args.path, "a" if args.append else "w", encoding="utf-8") as fh:
        fh.write(args.text)
    with open(args.path, encoding="utf-8") as fh:
        back = fh.read()
    if not back.endswith(args.text):
        raise WadError("VERIFY_FAILED", "the file does not end with what was written")
    state.trace("file-write", {"path": args.path, "chars": len(args.text)})
    return {"ok": True, "path": os.path.abspath(args.path)}, f"wrote {len(args.text)} chars to {args.path}"


@command("file-list", "list a folder (policy-gated)", Arg("path", positional=True),
         Arg("pattern", help="only names containing this"), group="system", readonly=True)
def cmd_file_list(args):
    require_system()
    try:
        names = sorted(os.listdir(args.path))
    except OSError as e:
        raise WadError("FILE_ERROR", str(e)) from None
    rows = []
    for n in names:
        if args.pattern and args.pattern.lower() not in n.lower():
            continue
        full = os.path.join(args.path, n)
        is_dir = os.path.isdir(full)
        rows.append({"name": n, "dir": is_dir, "size": None if is_dir else os.path.getsize(full)})
    text = "\n".join(f"{'<dir>' if r['dir'] else r['size']:>12}  {r['name']}" for r in rows)
    return {"ok": True, "entries": rows}, text or "(empty)"


@command("process-list", "running processes (policy-gated)", Arg("name", help="filter"),
         group="system", readonly=True)
def cmd_process_list(args):
    require_system()
    r = subprocess.run(["tasklist", "/fo", "csv", "/nh"], capture_output=True, text=True,
                       errors="replace")
    rows = []
    for line in r.stdout.splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) >= 5 and (not args.name or args.name.lower() in parts[0].lower()):
            rows.append({"name": parts[0], "pid": int(parts[1]), "memory": parts[4]})
    return {"ok": True, "processes": rows}, "\n".join(
        f"{r['pid']:>7}  {r['name']:<32} {r['memory']}" for r in rows) or "(none)"


@command("process-kill", "end a process by pid or exe name (needs allow_kill; system "
         "processes always refused)",
         Arg("which", positional=True, help="pid or exe name"), group="system")
def cmd_process_kill(args):
    p = require_system()
    if not p.get("allow_kill"):
        raise WadError("POLICY_DENIED", "process-kill is off",
                       f'enable it in {state.POLICY_FILE} with "allow_kill": true')
    name = args.which.lower()
    if not name.isdigit() and not name.endswith(".exe"):
        name += ".exe"
    if name in CRITICAL:
        raise WadError("POLICY_DENIED", f"{name} is part of Windows and is never killed")
    if name.isdigit():
        r = subprocess.run(["tasklist", "/fi", f"PID eq {name}", "/fo", "csv", "/nh"],
                           capture_output=True, text=True, errors="replace")
        exe = r.stdout.split('","')[0].strip('"').lower() if r.stdout.strip() else ""
        if exe in CRITICAL:
            raise WadError("POLICY_DENIED", f"pid {name} is {exe}, part of Windows")
        cmd = ["taskkill", "/pid", name, "/f"]
    else:
        cmd = ["taskkill", "/im", name, "/f"]
    r = subprocess.run(cmd, capture_output=True, text=True, errors="replace")
    state.trace("process-kill", {"which": args.which, "exit": r.returncode})
    if r.returncode != 0:
        raise WadError("KILL_FAILED", (r.stderr or r.stdout).strip())
    return {"ok": True}, (r.stdout or "").strip()
