# Architecture

```
wad.py                  entry point (python wad.py ...), keeps `import wad` working
wadlib/
  registry.py           Arg / Command / @command: ONE declaration per command
  cli.py                argparse front end, execute() (errors -> payloads), main()
  mcp.py                MCP server over stdio, tools generated from the registry
  state.py              state dir, last snapshot / screenshot, trace, sessions
  uia.py                windows, tree walk, snapshots, refs, selectors, rendering
  verify.py             before/after fingerprints, changes, --expect
  inputs.py             keys, guarded real mouse, foreground and occlusion checks
  win32.py              raw Win32: foreground, layouts, DPI, elevation, Unicode SendInput
  commands.py           observe / act / mouse / keyboard / windows commands
  vision.py             screenshots, marks, coordinate mapping, Windows OCR
  office.py             Excel and Word through COM
  system.py             policy-gated shell / files / processes
  workflow.py           batch, replay records, trace, doctor, guide
tests/                  pytest against tests/fake_uia.py (runs on any OS)
tools/gen_docs.py       docs/COMMANDS.md from the registry (a test keeps it fresh)
tools/smoke_windows.py  end-to-end on a real Windows desktop (CI: windows-latest)
.claude/skills/wad-desktop/   the agent playbook (also served by `wad guide`)
```

## One declaration, four front ends

A command is a function plus its `Arg`s, registered with `@command`. From that single
declaration wad builds the command-line parser, the MCP tool with its JSON schema and
annotations, batch-step validation (`Command.namespace`) and the command reference. They
cannot drift apart - adding a command is one decorated function.

Commands never print. They return `(payload, text)`: the payload for `--json`, MCP
`json: true` and batch; the text for people and for token-frugal agents. Failures raise
`WadError(code, message, hint)`; `cli.execute` turns every error - including unexpected
exceptions from a misbehaving app - into a payload, logs it, and never crashes the caller.

## Seeing: snapshots and refs

`uia.walk` does a depth-first walk of the UIA tree, recording for every element its role,
name, AutomationId, class, rectangle, supported patterns, value, runtime id, and its
**path** from the window: a list of `(control type, AutomationId, name, index)` steps.
The snapshot is saved to disk, so refs work across separate processes.

Resolving a ref walks the path again from the window: a unique `(type, aid, name)` match
per step, else the recorded index. At the end the runtime id is compared: equal means
the very element observed; different but same role and name means the app rebuilt it
(accepted, reported `re-resolved`); anything else is refused. Walks stop at `--limit`
elements and skip elements that vanish mid-walk.

Selectors (`uia.parse_selector` / `find_selector`) search the live tree instead and fail
on ambiguity. Both paths return the same `(control, description, how, window handle)`.

## Acting: patterns first, input second

`click` tries Invoke, Toggle, SelectionItem, ExpandCollapse, then LegacyIAccessible's
default action - no pointer, no focus stealing, works on covered windows. Real input
(`--headed`, `right-click`, `drag`, `press`, typing fallback) always goes through
`inputs.guard` (target app in front, or nothing is sent) and, for clicks, `hit_check`
(the point must land on the element or at least in the same app).

Typing uses `SendInput` with `KEYEVENTF_UNICODE`: characters are sent as UTF-16 code
units, not as keys of the current layout, so any language arrives intact no matter which
keyboard layout the target window has.

## Proving: verification

`verify.observe` fingerprints the element (alive, name, enabled, content, toggle,
expanded, selected), the focused element, the window title and the app's top-level
windows. After an action wad polls briefly for a difference and reports it as
`changed:` lines. `--expect` / `--expect-gone` search the app's windows for text.
Commands with a known intended result read it back: `type` compares the field's content
and falls back from the value pattern to keystrokes when the app ignored the value
(Excel); `check`, `select`, `clear`, `excel-write`, `word-write`, `file-write` likewise.

## Vision

`vision.capture` grabs a window, region or the whole virtual screen (Pillow), optionally
draws the snapshot's refs on it, downsizes wide images, and stores origin and scale.
`click-xy` maps a point in the last screenshot back to screen pixels, so an agent can use
exactly what it saw. OCR uses Windows' own `Windows.Media.Ocr` through the WinRT
projection - local, fast, supports every language pack installed on the PC.

## Office

Excel and Word are attached through `GetActiveObject` (the user's running instance; a new
visible one only with `--start`). Calls retry while Office rejects them (cell edit mode,
open dialog) and then fail `OFFICE_BUSY` with the fix. `Value2` keeps numbers as numbers
and turns error values into their names. Writes use `Formula2` for formulas (dynamic
arrays spill instead of getting an implicit `@`) and are always read back.

## Recording and replay

After every successful non-read-only command, `workflow.replay_record` logs the step
with the durable selector (the one given, or `selector_for` the element), the window
title, non-default arguments only, and password text replaced by `${ENV:WAD_SECRET}`.
`trace --export` collects those steps; `batch` runs them with per-step `retry`,
`optional` and `sleep`.

## Testing

`tests/fake_uia.py` implements the slice of the `uiautomation` API wad uses, with an app
that has the behaviours that matter (a dialog-opening button, a lying value pattern,
ambiguous names, a disabled button, a password field, a scrollable list). The test
fixture swaps the raw Win32 input for recorders. `tools/smoke_windows.py` then drives the
real Notepad on a real Windows desktop in CI: launch, snapshot, Unicode typing, value
typing, screenshot with marks, OCR, the MCP server, replay, window operations, a headed
click, a context menu, a dialog's checkbox and edit, and closing without saving.
