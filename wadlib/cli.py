"""Entry point: load every command module, run one command, print or return its result."""
import json
import sys
import time

import uiautomation as auto

from . import __version__, state, win32
from . import commands, office, system, vision, workflow  # noqa: F401  (register commands)
from .registry import COMMANDS, WadError, build_parser

DESCRIPTION = f"""wad {__version__} - Windows Agent Desktop: drive any Windows app like a person
would, through its accessibility tree; verified actions, guarded input, MCP server.

  wad snapshot --window Notepad -i        see the app, with refs (e12)
  wad click e12 --expect "Saved"          act, and prove it worked
  wad type "role=Edit name=Search" hi     selectors work anywhere a ref does
  wad guide                               the full playbook for agents
"""


def execute(name, ns):
    """Run one command; always returns (payload, text) - errors become payloads too."""
    started = time.time()
    try:
        payload, text = COMMANDS[name].fn(ns)
        if payload.get("ok", True):
            workflow.replay_record(name, ns, payload)
        return payload, text
    except WadError as e:
        payload = {"ok": False, "code": e.code, "message": e.message, "hint": e.hint}
        text = f"ERROR {e.code}: {e.message}" + (f"\n  hint: {e.hint}" if e.hint else "")
        state.trace("error", {"command": name, "code": e.code, "message": e.message,
                              "ms": int((time.time() - started) * 1000)})
        return payload, text
    except Exception as e:                   # an app misbehaving must not kill an agent run
        payload = {"ok": False, "code": "INTERNAL", "message": f"{e.__class__.__name__}: {e}",
                   "hint": "the UI probably changed mid-action - snapshot and retry; if it "
                           "repeats, it is a wad bug (see the trace)"}
        state.trace("error", {"command": name, "code": "INTERNAL", "message": str(e)})
        return payload, f"ERROR INTERNAL: {payload['message']}\n  hint: {payload['hint']}"


def main(argv=None):
    for stream in (sys.stdout, sys.stdin):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    win32.set_dpi_aware()
    args = build_parser(DESCRIPTION).parse_args(argv)
    if args.cmd == "mcp":
        from . import mcp
        return mcp.serve()
    with auto.UIAutomationInitializerInThread():
        payload, text = execute(args.cmd, args)
    if getattr(args, "json", False):
        print(json.dumps(payload, ensure_ascii=False, indent=2, default=str))
    else:
        print(text)
    return 0 if payload.get("ok", True) else 1


if __name__ == "__main__":
    sys.exit(main())
