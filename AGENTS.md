# For AI agents

**Using wad to operate Windows apps?** Read
[.claude/skills/wad-desktop/SKILL.md](.claude/skills/wad-desktop/SKILL.md) first (or run
`python wad.py guide`). The short version:

1. `python wad.py doctor`, then `python wad.py desktop-check`, then `python wad.py inspect --all`.
   If desktop access fails, use the interactive-session setup in
   [docs/COMPUTER_USE.md](docs/COMPUTER_USE.md) before trying to launch an app.
2. `snapshot --window "<title>" -i` -> act on a ref or selector -> read the `changed:`
   lines -> snapshot again when the UI changed.
3. Use `--expect "<text>"` when you know what success looks like; never repeat an action
   whose effect you did not check.
4. Excel/Word data: `excel-*` / `word-*`. No tree: `screenshot`, `ocr`, `click-text`.
5. Never type into the user's own documents; stop before send/delete/overwrite unless
   the task clearly asked for it.

**Changing wad itself?** Read [CLAUDE.md](CLAUDE.md).
