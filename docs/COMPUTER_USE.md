# Control Windows apps when Computer Use is unavailable

wad has its own way for an assistant to see and control Windows apps. Its command line
and MCP connection expose the same actions: list windows, inspect an app, click, type,
take screenshots, and work with Office. They do not require the separate Computer Use
plugin.

## Step 1: Check desktop access

Open a normal terminal while signed in to Windows, go to this project folder, and run:

```powershell
python wad.py desktop-check
python wad.py windows
python wad.py inspect --all
```

A working check looks like `desktop ready: WinSta0\Default (11 visible windows)`; the
number of windows will vary. `windows` should then list the apps wad can see. These
checks only read information; they do not click or type. To look inside one app, run
`python wad.py inspect --window "APP TITLE"`; add `--visual` to capture its window and
read visible text with Windows OCR. OCR needs `pip install -r requirements-ocr.txt`.
See [INSPECTION.md](INSPECTION.md) for examples and coverage limits.

If you see `DESKTOP_UNAVAILABLE`, the process is on an isolated execution desktop. An
assistant or MCP client started on that desktop will inherit the same problem. Run the
wad connection from the signed-in user's interactive desktop session and repeat the
check. If the screen is locked, unlock it first. If an individual app runs as
administrator, wad needs the same elevation to see that app.

`launch` also checks desktop access. It will explain the problem before opening an app
that wad would be unable to see.

## Step 2: Connect the assistant to wad

Configure an assistant that supports MCP to start `python wad.py mcp` in the interactive
Windows session. One common MCP configuration shape is:

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

Replace the example path with the actual path to `wad.py`; keep double backslashes in
the JSON configuration. The assistant can then call the `desktop-check`, `windows`,
`inspect`, `snapshot`, `click`, `type`, and other wad tools. The MCP process must run in the
interactive desktop session too;
starting it as a child of an isolated desktop does not move it to the user's desktop.
See [MCP.md](MCP.md) for client examples and [SAFETY.md](SAFETY.md) for input guards.

## What the two errors mean

| What you see | Meaning | Next step |
|---|---|---|
| `Computer Use native pipe is unavailable` | The separate Computer Use plugin cannot connect to its own Windows helper. | Repair or restart that plugin's runtime, or use wad's MCP connection above. |
| `DESKTOP_UNAVAILABLE` from wad | wad is running on a different Windows desktop from the one receiving your input. | Start wad or its MCP connection in the interactive desktop session. |
| `windows` shows no apps but `desktop-check` is ready | No visible top-level windows were found in that session, or the target app is hidden/elevated. | Open the app there, check `windows` again, and use `doctor` for elevation information. |

The Computer Use pipe is a live Windows named-pipe endpoint, not a missing project file.
Creating a file or a similarly named pipe in this repository will not start its helper.
wad MCP supplies computer control through this project's own implementation; it does
not claim to repair or implement the separate plugin's private connection.

## Repeat the live check

Developers can run `python tools/smoke_computer_use.py --ocr` from the interactive desktop
after installing OCR, or omit `--ocr` for the accessibility and MCP check.
It opens only wad's bundled test app, exercises the MCP connection, reads back the
result, and closes that test app. It does not open or edit personal documents.
