"""wad as an MCP server (Model Context Protocol, stdio transport).

Any MCP client - Claude Desktop, Claude Code, Cursor, VS Code - gets every wad command as
a tool, generated from the same declarations as the command line. No SDK dependency: the
stdio transport is newline-delimited JSON-RPC 2.0, small enough to implement exactly."""
import base64
import json
import sys
import traceback

import uiautomation as auto

from . import __version__
from .registry import COMMANDS, WadError, input_schema

PROTOCOLS = ["2025-06-18", "2025-03-26", "2024-11-05"]
INSTRUCTIONS = """wad drives Windows desktop apps through their accessibility tree.
Loop: `snapshot` (window, -i) -> act on a ref (e12) or a selector (role=Button name=Save)
-> read the reported `changed:` lines -> snapshot again when the UI changed.
Rules: prefer click/type/select/check (accessibility, no mouse) over headed or xy clicks;
add `expect` when you know what success looks like; never repeat an action whose effect
you did not check; for Excel/Word/Outlook/PowerPoint use excel-*/word-*/outlook-*/ppt-*;
for web pages use browser-* (browser-launch first); when the tree is empty (games,
canvases, remote desktop) use screenshot + ocr + click-text/click-xy, or detect +
click-mark. Never send mail or submit payments unless the person asked for exactly that.
Call the `guide` tool once for the full playbook."""


def tools():
    out = []
    for c in COMMANDS.values():
        if not c.mcp:
            continue
        schema = input_schema(c)
        schema["properties"]["json"] = {"type": "boolean",
                                        "description": "return the full JSON payload instead of text"}
        tool = {"name": c.name, "description": c.help, "inputSchema": schema,
                "annotations": {"readOnlyHint": c.readonly,
                                "destructiveHint": c.group == "system" and not c.readonly,
                                "openWorldHint": False}}
        out.append(tool)
    return out


def call(name, arguments):
    from .cli import execute_guarded
    c = COMMANDS.get(name)
    if c is None or not c.mcp:
        return {"content": [{"type": "text", "text": f"unknown tool {name!r}"}], "isError": True}
    try:
        ns = c.namespace(arguments or {})
    except WadError as e:
        return {"content": [{"type": "text", "text": f"ERROR {e.code}: {e.message}"
                             + (f"\n  hint: {e.hint}" if e.hint else "")}], "isError": True}
    payload, text = execute_guarded(name, ns)
    body = json.dumps(payload, ensure_ascii=False, default=str) if ns.json else text
    content = [{"type": "text", "text": body}]
    if c.image and payload.get("ok") and payload.get("path"):
        try:
            with open(payload["path"], "rb") as fh:
                content.append({"type": "image", "mimeType": "image/png",
                                "data": base64.b64encode(fh.read()).decode("ascii")})
        except OSError:
            pass
    return {"content": content, "isError": not payload.get("ok", True)}


def handle(msg):
    """One JSON-RPC message in, the response (or None for a notification) out."""
    method, mid = msg.get("method"), msg.get("id")
    if mid is None:
        return None                          # notifications (initialized, cancelled...)
    try:
        if method == "initialize":
            asked = (msg.get("params") or {}).get("protocolVersion")
            result = {"protocolVersion": asked if asked in PROTOCOLS else PROTOCOLS[0],
                      "capabilities": {"tools": {"listChanged": False}},
                      "serverInfo": {"name": "wad", "version": __version__},
                      "instructions": INSTRUCTIONS}
        elif method == "ping":
            result = {}
        elif method == "tools/list":
            result = {"tools": tools()}
        elif method == "tools/call":
            p = msg.get("params") or {}
            result = call(p.get("name"), p.get("arguments"))
        elif method in ("resources/list", "prompts/list"):
            result = {method.split("/")[0]: []}
        else:
            return {"jsonrpc": "2.0", "id": mid,
                    "error": {"code": -32601, "message": f"method not found: {method}"}}
    except Exception as e:
        traceback.print_exc(file=sys.stderr)
        return {"jsonrpc": "2.0", "id": mid, "error": {"code": -32603, "message": str(e)}}
    return {"jsonrpc": "2.0", "id": mid, "result": result}


def serve(stdin=None, stdout=None):
    """Read requests line by line. Commands never print (they return), and anything a
    library prints anyway goes to stderr, so stdout carries nothing but protocol."""
    stdin = stdin or sys.stdin
    out = stdout or sys.stdout
    real_stdout, sys.stdout = sys.stdout, sys.stderr
    try:
        with auto.UIAutomationInitializerInThread():
            for line in stdin:
                line = line.strip()
                if not line:
                    continue
                try:
                    msg = json.loads(line)
                except ValueError:
                    resp = {"jsonrpc": "2.0", "id": None,
                            "error": {"code": -32700, "message": "parse error"}}
                else:
                    resp = ([r for r in (handle(m) for m in msg) if r] if isinstance(msg, list)
                            else handle(msg))
                if resp:
                    out.write(json.dumps(resp, ensure_ascii=False, default=str) + "\n")
                    out.flush()
    finally:
        sys.stdout = real_stdout
    return 0
