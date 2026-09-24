# Working on wad

wad is Windows desktop automation for agents. Code in `wadlib/`, entry `wad.py`, design
in [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md). The agent-facing playbook is
`.claude/skills/wad-desktop/` - keep it true when behaviour changes.

## Rules of the codebase

- **One declaration per command.** Add or change a command with `@command(...)` and `Arg`s
  in the right module; the CLI, MCP tool, batch validation and docs follow. Then run
  `python tools/gen_docs.py` (a test fails if `docs/COMMANDS.md` is stale).
- **Commands return `(payload, text)` and never print.** MCP relies on stdout carrying only
  protocol. Failures raise `WadError(CODE, message, hint)`; the hint says what to do next.
  A new code must be added to `.claude/skills/wad-desktop/references/errors.md` (a test
  checks).
- **Every action proves itself**: observe before, `verify.after_action` after, read back
  when the intended result is known. Never report success from "the call returned".
- **Real input only through `inputs`**: `guard(hwnd)` before keys, `hit_check` before a
  physical click. No direct `SendKeys`/`Click` in commands.
- **Text is typed with `win32.type_unicode`** (layout-proof), not SendKeys.
- **Anything app-specific learned the hard way** goes in a comment where it matters, in the
  README's "learned the hard way", and in `references/recipes.md` if agents need it.
- Win32 calls live in `win32.py`, loaded lazily; everything else must import on Linux.
- **Check desktop access before launching apps.** `desktop-check` distinguishes the
  signed-in user's desktop from an isolated execution desktop. Keep the CLI and MCP
  recovery steps in `docs/COMPUTER_USE.md` current when this behavior changes.
- Comments explain *why* (the app behaviour that forced the code), not what.

## Testing

- `python -m pytest` - runs anywhere against `tests/fake_uia.py`. Model a new app
  behaviour in the fake (e.g. a lying pattern, a lagging field) before relying on it.
- `python tools/smoke_windows.py` - real Notepad on a real desktop; CI runs it on
  `windows-latest` on every push. When it fails, it is a real bug until proven otherwise:
  read the log, reproduce in the fake, fix, add a test.
- Excel/Word COM paths cannot run in CI (no Office on runners); test value handling with
  unit tests and the real thing with `Run_Excel_Demo.bat` on a PC with Office.
