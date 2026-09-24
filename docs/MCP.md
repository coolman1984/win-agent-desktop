# wad as an MCP server

`python wad.py mcp` serves every wad command as an MCP tool over stdio (JSON-RPC 2.0,
protocol 2025-06-18, older revisions accepted). The tools are generated from the same
declarations as the command line, so names and arguments match `docs/COMMANDS.md`
exactly: `inspect`, `snapshot`, `click`, `type`, `excel-write`, ...

wad's MCP server is its project-owned computer-control connection. It does not depend on
the separate Computer Use runtime or its native pipe. Run `wad desktop-check` in the
same environment first; if it reports `DESKTOP_UNAVAILABLE`, launch the MCP server from
the signed-in user's interactive desktop session. See [COMPUTER_USE.md](COMPUTER_USE.md).

What a client gets:

- **instructions** at connect time: the core loop and rules, so the model starts right.
- **tool annotations**: `readOnlyHint` on observe tools (`inspect`, `snapshot`, `get`, `find`...),
  `destructiveHint` on system tools - clients can auto-approve reads and ask for writes.
- **text results** by default (compact: refs, `changed:` lines, errors with a hint);
  pass `"json": true` in any call for the full payload.
- **images**: `screenshot` returns the PNG with the result, so a vision model can look.
- `isError: true` on every failure, with the error code and hint in the text.

## Claude Desktop

`%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "wad": {
      "command": "python",
      "args": ["C:\\path\\to\\win-agent-desktop\\wad.py", "mcp"]
    }
  }
}
```

## Claude Code

```powershell
claude mcp add wad -- python C:\path\to\win-agent-desktop\wad.py mcp
```

Claude Code also picks up the skill in `.claude/skills/wad-desktop/` when started in this
repository; elsewhere, copy that folder to `%USERPROFILE%\.claude\skills\`.

## Cursor / VS Code / others

Any client that launches stdio servers: command `python`, args
`["C:\\path\\to\\wad.py", "mcp"]`. After `pip install -e .` the command is simply `wad mcp`.

## Several agents at once

Set `WAD_SESSION` differently for each server process (`"env": {"WAD_SESSION": "a"}`):
each gets its own snapshot and screenshot state, so refs never cross between agents. The
trace file is shared and every entry carries its session.

## Notes

- The server runs in the user's desktop session - it sees and drives what the user sees.
  Run it as the same user (and elevated only if the apps you drive are elevated).
- System tools (`shell`, `file-*`, `process-*`) are listed but refuse to run until a
  person enables them in the policy file (see [SAFETY.md](SAFETY.md)).
