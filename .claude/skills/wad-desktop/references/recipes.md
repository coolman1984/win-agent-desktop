# Recipes - how real apps behave, and the moves that work

Learned by driving them; each note is something that broke once.

## Any app: open your own, find by name, close cleanly

```
python wad.py launch notepad.exe --title Notepad        # waits for a NEW window only
python wad.py snapshot --window "Untitled - Notepad" -i
python wad.py type 'name~="text editor"' "Report draft"
python wad.py press ctrl+s --window "Untitled - Notepad"
python wad.py snapshot --window "Untitled - Notepad" --popup      # the Save As dialog
```

`launch` ignores windows that were already open, so it never grabs the user's copy.
Titles drift ("Untitled - Notepad" becomes "Report draft - Notepad" after saving); after
the first snapshot, commands without `--window` use that window's handle, which does not
drift.

## Save As / Open dialogs (common file dialogs)

- The file name box: `role=Edit name="File name:"` (Win32) - a `ComboBox` wraps it; use the
  Edit. Type the FULL path there, then `press enter` - faster and surer than browsing.
- "Confirm Save As / already exists" appears as its own dialog: `snapshot --popup`, and
  only answer "Yes" if replacing was part of the task.

## Notepad (Windows 11)

- The editor is `role=Document name="Text editor"`; each tab is a `TabItem`.
  "Add New Tab" opens yours - never type into tabs you did not create.
- Closing an unsaved tab asks inside the window (not a popup): after `close`, snapshot
  the window and click `name~="don't save"`.

## Calculator

- Every key is a named button: `click "name=Seven"`, `click "name=Plus"`, `click "name=Equals"`.
- Read the result: `get "aid=CalculatorResults"` (its name is "Display is 42").

## Office: ribbon, menus, galleries

- Ribbon tabs: `click "role=TabItem name=Insert"`. Buttons: `click "name=Table"`.
- Drop-downs: `expand "name=Conditional Formatting"`, then find the item by name.
  Sub-menus (Data Bars, Icon Sets) do NOT open through expand - press their accelerator
  key after the menu is open, then everything is reachable by name.
- Office only opens ribbon drop-downs while it is the active window, and an open menu is
  its own foreground popup. wad's guard accepts the app's own popups, so it will not close
  the menu it is about to use.
- Key tips: `input-lang en --window "<book> - Excel"` first, then `press "alt h o r"`
  (one key at a time; bursts lose letters).
- Dialogs (Insert Table, Format Cells) are windows of the same process:
  `snapshot --window Excel --popup`.

## Excel

- **Data goes through COM**: `excel-write B2 --values '[["Region","Sales"],["East",120]]'`
  writes, recalculates and reads every cell back; strings starting with `=` are formulas
  (dynamic arrays spill). `excel-read A1:D20` returns typed values and error names
  (`#DIV/0!`, `#SPILL!`). `excel-read --text` gives what Excel displays (formatted).
- Never trust `type` into a cell through the grid: Excel's value pattern accepts and
  forgets. (wad already types keystrokes there, but COM is the right tool.)
- `OFFICE_BUSY` = a cell is in edit mode or a dialog is open: `press esc` then retry.
- Formatting, charts, conditional formats: the ribbon (above), then verify by reading.
- Reading the grid itself (when COM is unavailable): cells are `DataItem`s whose
  AutomationId is the address (`B7`) and whose value is the displayed text - cells inside
  an Excel Table sit one level deeper. Only visible cells are in the tree: go to the range
  first (Name Box: `type "role=Edit name='Name Box'" B7 --submit`).

## Word

- `word-read` for the text; `word-write "text" --at end` / `--at replace --find old`.
- Formatting through the ribbon, like Excel.

## Outlook / mail and anything that SENDS

- Composing is fine; pressing Send is irreversible. Unless the task explicitly said to
  send, stop at a filled draft and ask.

## Browsers (Edge, Chrome)

- Prefer `browser-*` (DevTools): `browser-launch`, `browser-snapshot`, `browser-click`,
  `browser-type`. Names come from labels (`aria-label`, `<label for>`, placeholder), so
  `text=Email` finds the field labelled Email.
- wad's browser has its own profile: the person's saved logins are NOT there. If a task
  needs their account, let them sign in in that window first.
- The person's own, already-open browser can only be driven through the accessibility
  tree (below) - do not close or reuse their tabs.
- A click that did not change the page says so ("the page did not visibly change"): look
  (`browser-text`, `browser-snapshot`) before clicking again.

## Electron apps (Teams, VS Code, Slack) and web content through the tree

- Web content is in the tree (Chromium exposes it): links are `Hyperlink`, fields are
  `Edit`, buttons `Button`. Snapshots of big pages are huge - use `--depth` and `--root`,
  or `find` / selectors with `name~=`.
- Chromium builds the web tree lazily; if a page shows only a few elements, wait a second
  and snapshot again, or `click` inside the page once.

## Outlook

- `outlook-list` numbers the mails; `outlook-read 2` reads the 2nd. Drafts are safe:
  `outlook-draft --to ... --subject ... --body "Line 1\nLine 2" --attach "C:\f.pdf"`.
- Sending is irreversible and off: `outlook-send` works only after the person sets
  `"allow_send": true` in the policy. Otherwise leave the draft and tell them.

## Old-style apps (WinForms, VB6, Delphi, MFC)

- Controls that only speak the old accessibility API (MSAA) show up with few patterns:
  a WinForms drop-down given an accessible name has no expand action and no option
  elements. wad handles it: `select` walks the list with Home/Down while reading the
  shown value, `expand`/`collapse` use Alt+Down/Alt+Up, and `get` reads the old
  checked/expanded state. The AutomationId of a WinForms control is its `Name` in code
  (`aid=nameBox`) - usually the most stable selector.

## Win32 menus and context menus

- Menu bar: `click "role=MenuItem name=File"` opens it; the menu is a popup window:
  `snapshot --window "<app>" --popup` then click the item.
- Context menus: `right-click eN`, then `snapshot --popup` (class `#32768`).

## Settings, installers, admin tools

- Many admin tools (Device Manager, regedit, some installers) run elevated. If wad sees
  only 1-2 elements, doctor/snapshot will say it is UIPI: the terminal must run as
  administrator. A UAC prompt runs on the secure desktop - no automation can click it;
  ask the person.

## Games, remote desktops, canvases, video

- No accessibility tree: `screenshot`, then `ocr` / `click-text` / `click-xy`.
- Remote desktop / Citrix windows are a picture of another PC: automate inside that PC
  (run wad there) when you can; OCR + clicks otherwise.
