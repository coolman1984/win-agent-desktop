# Selectors

Anywhere a command takes a target, it accepts either a **ref** from the last snapshot
(`e12`, or snapshot-qualified `s1234567:e12`) or a **selector**, searched live in the
window (`--window`/`--hwnd`, else the last snapshot's window).

| Term | Matches |
|---|---|
| `role=Button` | control type (Button, Edit, Document, MenuItem, ListItem, CheckBox, ComboBox, TabItem, TreeItem, Hyperlink, DataItem, Window, Pane...) |
| `name=Save` | the accessible name, exact, case-insensitive |
| `name~=sav` | the name contains the text |
| `aid=SaveButton` (or `id=`) | AutomationId - set by the app's developers; the most stable identity when present |
| `class=Edit` | the window class name |
| `value=42` / `value~=4` | the element's current value |
| `nth=2` | the 2nd match (in tree order) |
| `A >> B` | find A, then search for B only inside it |

Every term of a step must match (AND). Any term takes `=`, `:` or `~=`
(`role:Button` works too). Quote values with spaces: `name="Save as"` or
`name='Save as'`. In PowerShell, wrap the whole selector in single quotes:
`python wad.py click 'role=Button name="Save as"'`.

## Behaviour

- **Zero matches** -> `ELEMENT_NOT_FOUND`. **Several** -> `AMBIGUOUS_TARGET`, listing up to
  five with role, name and aid. wad never picks one for you; add a term, `nth=`, or `>>`.
- Searches go 40 levels deep and skip elements that vanish mid-search (Excel rebuilds its
  tree constantly).
- `wait --target "<selector>"` waits for one to appear (or `--gone`).

## Where selectors come from

- `find` prints one for every hit; `get` prints the selector of any element.
- wad writes selectors as `role=X aid=Y` when an AutomationId exists (and is not a bare
  number, which some frameworks regenerate), else `role=X name=Y`.
- Every successful action is logged with its selector; `trace --export` turns them into
  a batch file, so a flow explored with refs replays with selectors.

## Choosing good ones

1. `aid=` when the app sets it (most WinUI/WPF/UWP apps, Office's grid cells: `aid=B7`).
2. `role=` + `name=` for buttons, menu items, tabs - names are what users see and rarely
   change within a version. Localised apps change names with the UI language.
3. Scope with `>>` when the same name appears in several places:
   `name="Page Setup" >> role=Button name=OK`.
4. Avoid `nth=` in recorded flows - positions change more than names.
