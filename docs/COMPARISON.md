# How wad compares (research, September 2026)

What the most-used desktop-agent tools do, what wad took from each, and where it goes
further. Sources are linked; star counts and features change - check before relying on
them.

| | wad | [Windows-MCP](https://github.com/CursorTouch/Windows-MCP) | [Terminator](https://github.com/mediar-ai/terminator) | [agent-desktop](https://github.com/lahfir/agent-desktop) | [UFO2](https://microsoft.github.io/UFO/ufo2/overview/) |
|---|---|---|---|---|---|
| Platform | Windows | Windows | Windows | macOS (Windows planned) | Windows |
| Sees through | UIA tree; screenshots + Windows OCR as fallback | UIA tree, optional screenshot | UIA tree, pixels, DOM | AX tree | UIA + OmniParser vision model |
| Acts through | UIA patterns; guarded real input | mostly coordinates | selectors, patterns | AX actions; headed opt-in | UIA + app APIs |
| Element identity | refs (path + runtime id) **and** selectors | element ids per state capture | selectors | refs (snapshot-qualified) | control labels |
| Verifies effects | every action: observed changes, read-back, `--expect` | no | no (per its docs) | postcondition contradictions -> ACTION_FAILED | via LLM |
| Refuses unsafe input | foreground guard + occlusion hit-test | no | background mode | interaction lease | - |
| Layout-proof typing | Unicode SendInput | - | - | - | - |
| Office | COM read/write with read-back | via PowerShell | - | - | app APIs (Excel/Word COM) |
| Record / replay | trace -> batch with selectors, secrets as env vars | - | records human workflows -> YAML | trace viewer | - |
| MCP | built in, no SDK, generated from the command registry | yes | yes | FFI + CLI | - |
| System access | off by default, policy file, hard deny-list | PowerShell, registry, files | - | - | - |
| Agent playbook | skill + `guide` + MCP instructions | README | CLAUDE.md | bundled skill | - |

## What we took

- **agent-desktop**: refs from a snapshot, skeleton-then-drill (`--depth`, `--root`),
  headless-first with `--headed` opt-in, error codes with suggestions, a bundled skill,
  failing on stale or ambiguous targets instead of guessing.
- **Windows-MCP**: a screenshot mode alongside the tree, scroll/drag/hover, window
  management, the MCP surface, token-aware images.
- **Terminator**: durable selectors, deterministic replay of workflows so the model is
  only needed when something changes.
- **UFO2**: hybrid perception (tree first, vision when the tree is thin) and going through
  an app's own API (Office COM) when it has one.

## Where wad goes further

1. **Proof after every action.** Others return "done" when the call returned. wad reports
   what visibly changed, reads values back, and `--expect` turns "I think it worked" into
   a checked fact. It also auto-recovers the one lie we hit in practice (Excel's value
   pattern) by retyping.
2. **Input that cannot go astray.** Keystrokes need the target in front; real clicks are
   hit-tested against the element. A foreign window in the way means nothing is sent.
3. **Any language, any layout.** Text goes in as Unicode, so Arabic text into an app with
   an English layout (or the reverse) arrives as written; `input-lang` fixes the one case
   that needs a layout (Office key tips).
4. **Explore with AI, replay without it.** Every successful step is recorded with a
   selector; `trace --export` + `batch` replays it deterministically.
5. **Safe by default for system power.** Shell/file/process tools exist but are off until
   a person opts in, with write roots and never-allowed commands.
6. **Tested on a real desktop in CI.** Every push drives real Notepad on Windows.

## Closed since the first comparison

- **Vision model** for tree-less apps: `detect` / `click-mark` through an OmniParser server.
- **Browser DOM mode**: `browser-*` through the DevTools protocol.
- **Recording a human** (Terminator's idea): `record` with low-level hooks, final field
  values instead of keystrokes.
- **Events** (UFO/agent-desktop use them): `watch`, event-driven `wait` / `launch`.
- **Outlook and PowerPoint** through COM, with sending gated by policy.
- Also beyond the others: self-healing replay that refuses to guess, a watchdog for hung
  apps, and an opt-in visual step report.

## Known gaps (honest list)

- The vision model is not bundled: someone has to run the OmniParser server (it wants a
  GPU). Without it, OCR and the calling model's own eyes remain.
- `browser-*` drives wad's own browser profile; a person's already-open browser is only
  reachable through the accessibility tree.
- Excel, Word, Outlook and PowerPoint COM paths are tested against stand-ins, not real
  Office, because CI runners have no Office.
- The recorder records the windows a person uses; it does not record drag-and-drop or
  mouse-wheel scrolling yet.
