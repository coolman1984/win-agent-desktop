# wad - Windows Agent Desktop

Desktop automation for AI agents on Windows, built on UI Automation. The ideas come
from [lahfir/agent-desktop](https://github.com/lahfir/agent-desktop) (macOS): read an
app through its accessibility tree instead of pixels, give every element a short ref
(`e12`), and act on refs through accessibility patterns rather than mouse coordinates.

```powershell
pip install -r requirements.txt
python wad.py windows
python wad.py launch calculator
python wad.py snapshot --window Calculator -i      # interactive elements, with refs
python wad.py click e72                            # Invoke pattern, no mouse
python wad.py get e16                              # read state back
python wad.py type e3 "hello" --append             # ValuePattern, or --keys for keystrokes
python wad.py snapshot --depth 3                   # skeleton first ...
python wad.py snapshot --root e13                  # ... then drill in
```

In PowerShell write refs without the `@` (`e72`): PowerShell reads `@e72` as splatting.

Commands: `windows`, `launch`, `snapshot`, `find`, `click` (`--headed` for a real
mouse click), `type`, `get`, `press`, `focus`, `wait`, `screenshot`, `clipboard`,
`close`. `--json` gives structured output; errors carry a code and a hint. Every
action is appended to `%LOCALAPPDATA%\win-agent-desktop\trace.jsonl`.

## Demos

- `Run_Demo.bat` - Calculator (125 x 8 - 250 via Invoke only) and Notepad (a new tab,
  typing, the File menu, "Don't save"). Existing Notepad tabs are never typed into.
- `Run_Excel_Demo.bat` - builds a sales dashboard in its own Excel process: 200 orders,
  Insert > Table through the real dialog, a structured calculated column,
  dynamic-array formulas, cell styles, data bars, icon sets and a chart picked from
  Excel's galleries by name - then reads every number and every chart bar back and
  checks it against an independent Python calculation. Keep hands off for ~2 minutes.

## Things learned the hard way

- Setting an Excel cell through its ValuePattern reports success and reads back the
  new value, but Excel never stores it. Enter data the way a person does, and verify
  by reading it back.
- Office only opens ribbon drop-downs while it is the active window, and an open menu
  is its own foreground popup. A keystroke guard must accept the app's own popups, or
  it closes the menu it is about to use.
- Sub-menus (Data Bars, Icon Sets) do not open through ExpandCollapse; their
  accelerator key does, after which every gallery item is reachable by name.
- Keystrokes go to whatever window is in front, so they are only ever sent after
  checking the foreground window belongs to the demo's own process.
- Window titles drift (Notepad renames an unsaved tab after its first line), so
  windows are tracked by handle once found.

## Status

The Excel demo has passed end to end (9/9 checks). The polished version with the
docked chart and chart verification has not yet completed a clean run.
