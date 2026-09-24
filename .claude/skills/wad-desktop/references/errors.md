# wad error codes - what happened and what to do next

Every failure is `{"ok": false, "code": ..., "message": ..., "hint": ...}` (exit code 1).
Nothing half-happens silently: when an input guard refuses, nothing was sent.

## Finding things

| Code | Meaning | Do this |
|---|---|---|
| `NO_SNAPSHOT` | No snapshot taken yet in this session | `snapshot --window "<title>"` |
| `STALE_SNAPSHOT` | The ref (`s123:e4`) is from an older snapshot | Use refs from the latest snapshot |
| `REF_NOT_FOUND` | No such ref in the latest snapshot | Re-read the snapshot; refs are renumbered each time |
| `ELEMENT_GONE` | The element's place in the tree is gone | The UI changed - snapshot again |
| `ELEMENT_CHANGED` | The ref now points at a different element | Snapshot again; never retry the old ref |
| `ELEMENT_NOT_FOUND` | Nothing matches the selector | Check the window; `find --name ...` on a fresh snapshot |
| `AMBIGUOUS_TARGET` | The selector matches several elements (listed) | Add `aid=`, `role=`, `nth=N`, or scope with `>>` |
| `BAD_SELECTOR` | The selector cannot be parsed | Terms: `role= name= name~= aid= class= value= nth=`, steps joined by ` >> ` |
| `WINDOW_NOT_FOUND` | No window titled like that | `windows`; the title may have changed (unsaved docs rename themselves) |
| `AMBIGUOUS_WINDOW` | Several windows match the title part | Use the full title or `--hwnd` |
| `WINDOW_GONE` | The snapshot's window was closed | Relaunch or pick another window |
| `NO_POPUP` | `--popup` found no menu/dialog of that app | Open it first; some apps draw menus inside the main window - snapshot that |
| `NO_HANDLE` | The surface has no window handle | Snapshot its parent window and drill in with `--root` |
| `APP_NOT_FOUND` | `launch` found no exe, shortcut or Store app | Give the full path to the .exe, or the Start menu name |
| `OPTION_NOT_FOUND` | `select` found no such option (the hint lists them) | Use one of the listed names |
| `TEXT_NOT_FOUND` | OCR did not find the text | `ocr` to see what it reads; try `--lang`, a shorter text, or `screenshot --region` to zoom |

## Acting

| Code | Meaning | Do this |
|---|---|---|
| `NOT_ENABLED` | The element is disabled | Something must happen first (a required field, a selection) |
| `NOT_CLICKABLE` | No accessibility action on it | `click --headed` (real, guarded mouse) |
| `NOT_VISIBLE` | It has no area on screen | `scroll-to`, open its parent, or maximize the window |
| `NOT_TOGGLABLE` / `NOT_EXPANDABLE` / `NOT_SCROLLABLE` | Wrong command for this element | `get` shows its `actions`; use one of those |
| `NOT_CLOSABLE` | The window has no close action | `press alt+f4 --window ...` |
| `NOT_SUPPORTED` | The window cannot do that (move/resize/state), or no PowerShell | Use keyboard shortcuts instead |
| `VERIFY_FAILED` | The action ran but the expected effect did not happen | Look (`get`, `snapshot`) before retrying; the app may have shown an error, a mask changed the text, or the click hit nothing |
| `FOCUS_LOST` | The app could not be brought to the front; nothing was typed/clicked | Something else holds the foreground (a UAC prompt, an elevated window, a fullscreen app). Resolve that - do not retry in a loop |
| `OCCLUDED` | Another app's window covers the click point; nothing was clicked | Bring the target forward (`focus --window`), move/close what covers it |
| `LAYOUT_FAILED` | `input-lang` could not switch the window's keyboard layout | Is that layout installed? Settings > Time & language > Language |
| `APP_HUNG` | The app is not responding (Windows says so, or it did not answer within the watchdog time); wad stopped waiting | Do not retry in a loop: check `windows`, `wait` for it to recover, or ask the person. Slow but healthy apps: raise `WAD_WATCHDOG` (seconds, default 60) |
| `TIMEOUT` | Waited and it did not happen | Snapshot to see what the app shows instead (an error dialog?) |

## Office

| Code | Meaning | Do this |
|---|---|---|
| `OFFICE_NOT_RUNNING` | Excel/Word is not open | `launch excel.exe`, or pass `--start` |
| `OFFICE_BUSY` | Office refuses automation: a cell is in edit mode or a dialog is open | `press esc --window "<app>"`, close the dialog, retry |
| `OFFICE_ERROR` | Office rejected the call (bad range, protected sheet...) | Read the message - it is Office's own |
| `NOT_FOUND` | No such workbook / sheet / document | `excel-info` lists them |

## Vision, system, setup

| Code | Meaning | Do this |
|---|---|---|
| `NO_SCREENSHOT` | `click-xy` without a screenshot to map from | `screenshot` first, or `--screen` for absolute pixels |
| `OCR_LANGUAGE` | No OCR engine for that language | Install the language's OCR pack, or omit `--lang` |
| `MISSING_DEPENDENCY` | An optional package is missing (the hint names it) | `pip install ...` as the hint says |
| `POLICY_DENIED` | A system command is off, outside the allowed folders, or never allowed | Ask the person; never try to work around it |
| `FILE_ERROR` | The file or folder cannot be read | Check the path |
| `KILL_FAILED` | taskkill failed | Read the message (access denied = elevated process) |
| `BATCH_FAILED` | A batch step failed (see `results`) | Snapshot at that step, fix its selector |
| `USAGE` | Wrong or missing arguments | The message names the argument |
| `INTERNAL` | Something unexpected (usually the UI changed mid-action) | Snapshot and retry once; if it repeats, report it with `trace` |
