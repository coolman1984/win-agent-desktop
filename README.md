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
python wad.py desktop-check                        # can this process see the user desktop?
python wad.py windows                              # list visible app windows
python wad.py inspect --all                        # check what is detectable in open apps
python wad.py inspect --window Notepad --visual    # accessible controls + screenshot + OCR
python wad.py snapshot --window Notepad -i         # inspect an open app
python wad.py mcp                                  # every command as an MCP tool
```

Before typing a test in Notepad, create and verify a blank tab; Windows 11 may restore
someone's existing tabs. See the [Notepad recipe](.claude/skills/wad-desktop/references/recipes.md).

## What it does

- **See**: `inspect` inventories visible windows or combines one window's accessible
  controls with optional screenshot and OCR; `snapshot` (skeleton with `--depth`, drill
  in with `--root`, menus/dialogs with `--popup`), `find`, `get`, `wait`, `windows`.
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
  `excel-run`, `word-read`, `word-write`, `outlook-list/-read/-draft` (sending only when a
  person allows it), `ppt-read`, `ppt-add-slide`, `ppt-replace`, `ppt-save`.
- **Browsers from the inside**: `browser-launch` (Edge/Chrome, own profile, local-only
  DevTools port), `browser-snapshot` (refs `b12`), `browser-click` (trusted clicks),
  `browser-type` (React-safe, read back), `browser-text`, `-open`, `-tabs`, `-wait`,
  `-eval`, `-screenshot`, `-close`.
- **Smart eye** for icons without words: `detect` asks an OmniParser vision server what is
  on screen (refs `v7`), `click-mark v7` clicks it ([VISION.md](docs/VISION.md)).
- **Learns from a person**: `record job.json` watches them do the task (mouse and keyboard
  hooks) and writes a replayable file; Ctrl+Shift+F12 stops.
- **Heals itself**: a replayed step whose button was renamed finds the one clear match and
  says so; it never guesses between close candidates.
- **Survives frozen apps**: a hung app comes back as `APP_HUNG` instead of freezing the
  agent with it (`WAD_WATCHDOG` seconds).
- **Hears the desktop**: `watch` reports windows, menus and focus as UIA events; `wait`
  and `launch` wake the moment a window appears.
- **Step report, off by default**: `report on` photographs every action, `report build`
  writes one HTML page of the whole run.
- **System power, off by default**: `shell`, `file-read/-write/-list`, `process-list/-kill`
  behind a policy file and a never-allowed list ([SAFETY.md](docs/SAFETY.md)).
- **Record once, replay forever**: every action is logged with a selector;
  `trace --export flow.json` + `batch flow.json` repeat it without a model.
- **MCP server** built in (`wad mcp`), generated from the same command declarations as
  the CLI ([MCP.md](docs/MCP.md)).
- **Desktop access check**: `desktop-check` detects when an isolated execution desktop
  cannot see the person's apps; `launch` refuses early in that case. wad's own MCP server
  remains a computer-control path when a separate Computer Use runtime is unavailable
  ([COMPUTER_USE.md](docs/COMPUTER_USE.md)).

All commands: [docs/COMMANDS.md](docs/COMMANDS.md). In PowerShell write refs without
`@` and wrap selectors in single quotes.

## If the assistant cannot see your apps

Run `python wad.py desktop-check` in the same Windows session as the assistant. If it
says `desktop ready`, run `python wad.py windows` to see the available app windows. If it
says `DESKTOP_UNAVAILABLE`, wad is running on a separate execution desktop. Start the
assistant's wad MCP connection from your signed-in desktop session, as described in
[COMPUTER_USE.md](docs/COMPUTER_USE.md). This project cannot repair the separate
Computer Use plugin's missing native pipe; wad MCP provides its own connection.

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

- A newly opened Windows 11 Notepad window can restore old tabs and show a missing-file
  dialog. Create and verify a blank `Untitled` tab before typing test or draft text.
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
- A WinForms drop-down with an accessible name speaks only the old MSAA API: no expand
  action, no option elements. Picking an option means walking it with the keyboard and
  reading the shown value after each step - which is what `select` falls back to.
- OCR finds "Submit" inside "Submitted:" too; a whole-word match must win, or the click
  lands on the label instead of the button.
- A child process that inherits our stdout keeps the pipe open after wad exits (the
  caller waits forever), and under `wad mcp` its output would corrupt the protocol. Every
  app wad starts is fully detached.
- On company PCs `HTTP_PROXY` is often set, and even a request to `127.0.0.1` goes to the
  proxy and hangs. Local DevTools traffic always bypasses proxies.
- Windows silently removes a low-level input hook that answers too slowly, so the recorder
  only queues raw events in the hook and does the accessibility lookups on another thread.
- ctypes assumes every Windows function returns a 32-bit int. A module handle on 64-bit
  Windows does not fit, and SetWindowsHookEx quietly refused the cut-down value. Every
  handle-returning call needs its types declared.
- A recorded "type" is the field's final value, read from the field - reconstructing text
  from keystrokes gets paste, autocomplete, dead keys and backspace wrong.
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
