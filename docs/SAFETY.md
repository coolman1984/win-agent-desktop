# Safety model

An agent driving a real desktop can do real damage: type a password into a chat window,
click "Delete" in the wrong app, close someone's unsaved work. wad is built so that the
common mistakes fail loudly instead of happening.

## What wad guarantees

| Risk | Guard |
|---|---|
| Keystrokes land in whatever window is in front | Before every keystroke/real click, the target app must be in front (its own menus and dialogs count). wad brings it forward itself; if another window keeps the foreground, it sends **nothing** and fails `FOCUS_LOST`. |
| A real click hits a window that covers the target | The point is hit-tested first: it must land on the element, or at least inside the same app. Otherwise nothing is clicked (`OCCLUDED`). |
| A ref from an old snapshot hits a different element | Refs re-resolve by path and runtime id; a changed element fails `ELEMENT_CHANGED`/`ELEMENT_GONE`. |
| A selector matches the wrong one of several | Several matches fail `AMBIGUOUS_TARGET` - wad never guesses. |
| "Success" that did not happen | Every action reports observed changes; `type`/`check`/`select`/`clear`/`excel-write`/`word-write` read back and fail `VERIFY_FAILED`; `--expect` makes any action prove its effect. |
| Acting on a disabled control | Refused before acting (`NOT_ENABLED`). |
| Launch grabbing the user's own window | `launch` only accepts windows that did not exist before it ran. |
| Secrets in logs | Text typed into password fields is never logged; recorded flows use `${ENV:WAD_SECRET}`. |
| A different keyboard layout garbling text | Text is sent as Unicode characters, not keys, so it arrives as written in any layout. |

## Mail, browser, recording, pictures

- **Mail is drafted, not sent.** `outlook-draft` only saves to Drafts. `outlook-send` refuses
  until a person sets `"allow_send": true` in the policy file.
- **The browser is wad's own.** `browser-launch` uses a separate profile under the state
  folder (not the person's cookies and logins) and binds DevTools to 127.0.0.1. While it
  is open, any program of the same Windows user can reach that port - `browser-close`
  when done. `browser-eval` runs JavaScript in that page only.
- **Recording captures what a person does** in the windows it watches (`--window` limits
  it): clicks and final field values. Password fields are written as
  `${ENV:WAD_SECRET}`, never their text. Recording only runs while `record` is running
  and stops on Ctrl+Shift+F12.
- **The step report is off by default.** When a person turns it on, it saves pictures of
  the app after every action in the state folder; `report clear` deletes them. Password
  text is masked in the report.
- **The smart eye sends screenshots** to the OmniParser server you configure
  (`WAD_OMNIPARSER_URL`). Keep it on your own machine or one you trust.
- **Hung apps are left alone**: wad refuses to act on a window Windows reports as not
  responding, and gives up on any call that does not return within the watchdog time.

## System commands (shell, files, processes)

Off by default. A person enables them in
`%LOCALAPPDATA%\win-agent-desktop\policy.json`:

```json
{
  "system": true,
  "write_roots": ["C:\\Users\\me\\Documents\\wad-output"],
  "allow_kill": false
}
```

or with `WAD_ALLOW_SYSTEM=1` for one terminal. Even when enabled:

- `file-write` only writes inside `write_roots` (default `Documents\wad`), never into
  wad's own state directory.
- `process-kill` needs `allow_kill`, and never touches Windows' own processes.
- `shell` refuses, always: formatting drives, disk/boot/backup tools, deleting system
  folders or whole drives, HKLM registry changes, shutdown, Defender changes, creating
  or deleting users, piping downloads into `iex`, changing the execution policy, wiping
  logs, and editing wad's policy or trace.

**This is a seat belt, not a sandbox.** A deny-list catches an agent's mistakes; it cannot
stop a determined adversary (there are endless ways to phrase a command). For untrusted
agents or untrusted input (web pages, emails the agent reads), run the agent in a
separate Windows account, a VM, or Windows Sandbox, and keep system commands off.

## Prompt injection

Text the agent reads from apps (web pages, documents, emails) can contain instructions.
wad cannot tell those apart from data - the agent must. The skill tells agents to treat
app content as data and to stop before irreversible actions (send, delete, pay, overwrite)
the task did not clearly ask for.

## The trace

Every action, refusal and error is appended to
`%LOCALAPPDATA%\win-agent-desktop\trace.jsonl` with a timestamp and session - `wad trace`
shows it. It is the record of what an agent did on the machine.
