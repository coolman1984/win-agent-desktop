# wad - Windows Agent Desktop

Desktop automation for AI agents on Windows. wad reads an app through its accessibility
tree instead of pixels, gives every element a short ref (`e12`) or a durable selector
(`role=Button name=Save`), acts through accessibility patterns rather than mouse
coordinates - and **proves every action worked**. Ideas from
[lahfir/agent-desktop](https://github.com/lahfir/agent-desktop) (macOS), rebuilt for
Windows and taken further ([how it compares](docs/COMPARISON.md)).

```powershell
pip install -r requirements.txt                    # + requirements-ocr.txt for OCR
python wad.py doctor                               # is this PC ready?
python wad.py launch notepad.exe --title Notepad
python wad.py snapshot --window Notepad -i         # interactive elements, with refs
python wad.py type 'name~="text editor"' "مرحبا - hello"   # any language, verified
python wad.py click e14 --expect "Save As"         # act, and prove the effect
python wad.py mcp                                  # every command as an MCP tool
```

## What it does

- **See**: `snapshot` (skeleton with `--depth`, drill in with `--root`, menus/dialogs with
  `--popup`), `find`, `get`, `wait`, `windows`.
- **Act without the mouse**: `click`, `type`, `clear`, `select`, `check`/`uncheck`,
  `expand`/`collapse`, `scroll`, `scroll-to`, `focus` - through UIA patterns, so they work
  on covered windows and never steal the pointer.
- **Real input, guarded**: `click --headed`, `right-click`, `double-click`, `hover`,
  `drag`, `press` (chords and sequences), `input-lang`. Nothing is sent unless the target
  app is in front, and a click point covered by another app is refused.
- **Proof**: every action reports what changed (`window opened: 'Save As'`,
  `content: '' -> 'hi'`); `type`/`check`/`select` read back; `--expect TEXT` fails unless
  the app shows it. Excel's lying value pattern is detected and retyped.
- **Vision fallback** for apps without a tree: `screenshot` (with `--marks`), Windows'
  built-in `ocr`, `click-text`, `click-xy` in screenshot pixels.
- **Office through COM**: `excel-info`, `excel-read`, `excel-write` (read back),
  `excel-run`, `word-read`, `word-write`.
- **System power, off by default**: `shell`, `file-read/-write/-list`, `process-list/-kill`
  behind a policy file and a never-allowed list ([SAFETY.md](docs/SAFETY.md)).
- **Record once, replay forever**: every action is logged with a selector;
  `trace --export flow.json` + `batch flow.json` repeat it without a model.
- **MCP server** built in (`wad mcp`), generated from the same command declarations as
  the CLI ([MCP.md](docs/MCP.md)).

All 47 commands: [docs/COMMANDS.md](docs/COMMANDS.md). In PowerShell write refs without
`@` and wrap selectors in single quotes.

## For agents

- The playbook: [.claude/skills/wad-desktop/SKILL.md](.claude/skills/wad-desktop/SKILL.md)
  (Claude Code loads it automatically here; `python wad.py guide` prints it anywhere;
  MCP clients get the short version at connect time).
- Error codes and what to do: [references/errors.md](.claude/skills/wad-desktop/references/errors.md).
- How real apps behave: [references/recipes.md](.claude/skills/wad-desktop/references/recipes.md).

## Demos

- `Run_Demo.bat` - Calculator (125 x 8 - 250 via Invoke only) and Notepad (a new tab,
  typing, the File menu, "Don't save"). Existing Notepad tabs are never typed into.
- `Run_Excel_Demo.bat` - builds a sales dashboard in its own Excel process: 200 orders,
  Insert > Table through the real dialog, a structured calculated column,
  dynamic-array formulas, cell styles, data bars, icon sets and a chart picked from
  Excel's galleries by name - then reads every number and every chart bar back and
  checks it against an independent Python calculation. Keep hands off for ~2 minutes.
  Last run: 15/15 checks, 0 mouse clicks, about 2.5 minutes.

## Things learned the hard way

- Setting an Excel cell through its ValuePattern reports success and reads back the
  new value, but Excel never stores it. Enter data the way a person does (or through
  COM), and verify by reading it back.
- Keystrokes are queued, not delivered: right after typing, the app may still be drawing
  the last characters. Verification must poll until the text settles.
- Office only opens ribbon drop-downs while it is the active window, and an open menu
  is its own foreground popup. A keystroke guard must accept the app's own popups, or
  it closes the menu it is about to use.
- Sub-menus (Data Bars, Icon Sets) do not open through ExpandCollapse; their
  accelerator key does, after which every gallery item is reachable by name.
- Keystrokes go to whatever window is in front, so they are only ever sent after
  checking the foreground window belongs to the target's own process.
- Window titles drift (Notepad renames an unsaved tab after its first line), so
  windows are tracked by handle once found.
- Ribbon key tips (Alt, H, O, R) only match English letters. With an Arabic keyboard
  layout active they miss, and the text meant for the ribbon lands in a cell. Switch only
  that window's layout (`input-lang en`); type text as Unicode so it never depends on
  the layout.
- Copying OUT of Office can come back empty to other programs on a managed PC even
  though pasting in works. Read cells from the grid or COM instead.
- An app running as administrator is invisible to a non-elevated automation process
  (UIPI); its tree comes back almost empty. wad says so instead of guessing.

## Development

```powershell
pip install -r requirements-dev.txt
python -m pytest                     # the logic, against a fake UI Automation (any OS)
python tools/smoke_windows.py        # the real thing, on a Windows desktop
python tools/gen_docs.py             # after changing a command
```

CI runs the unit tests on Linux and Windows and drives real Notepad on `windows-latest`
on every push. Layout and design: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).
Contributor notes: [CLAUDE.md](CLAUDE.md).
