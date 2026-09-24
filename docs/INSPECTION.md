# Check what wad can see in Windows apps

`inspect` is a read-only check of open Windows apps. It works through wad's command line
and MCP connection. It reports what each detection method actually returned, so an
assistant can see when an app hides its contents.

## Try it on your PC

From a terminal running in your signed-in Windows desktop:

```powershell
python wad.py doctor
python wad.py inspect --all
python wad.py inspect --window "APP TITLE"
python wad.py inspect --window "APP TITLE" --visual
```

Copy `APP TITLE` from the `--all` result. `--all` takes a shallow look at every visible
window to keep the check quick. Inspect one window for more detail. Use `--depth 6`
if the app puts controls deeper in its accessibility tree, or raise `--limit 200` if
the result says `limited: true`. Some apps expose only a frame through accessibility;
`--visual` can still find text drawn inside it.

`--visual` saves a screenshot of that window and tries local Windows OCR. Install its
optional packages once with `python -m pip install -r requirements-ocr.txt`. OCR may
misread small text, and a screenshot can include overlapping windows. Check the
`coverage` and `warnings` fields in the result.

For buttons and icons with no accessible label, a separately running OmniParser server
can be used with:

```powershell
python wad.py inspect --window "APP TITLE" --model
```

`--model` sends the screenshot to the server configured by `WAD_OMNIPARSER_URL`
(default `127.0.0.1:8000`). Use `--url` to select another server. See
[VISION.md](VISION.md) for server setup. The model is optional and is not bundled with
wad. `--model` is never called during a plain inspection.

## What the results mean

- `elements` are controls and text an app exposes through Windows accessibility.
  `interactive: true` marks controls wad may be able to operate. The `rect` gives their
  position on screen. Password values are omitted.
- `ocr_lines` are text seen in pixels, with positions within the saved screenshot.
- `visual_elements` are model detections of visual controls, when requested and available.
- `limited: true` means the scan stopped at the requested depth or element limit.
  It does not mean the rest of the window is empty.
- `coverage` says which methods ran, were unavailable, or were not requested.

No Windows tool can guarantee it will detect *everything* inside every app. Windows can
block access to elevated apps, secure prompts, locked sessions, and content an app does
not expose. OCR can miss or misread text, and icons may need a vision model or a person
to view the screenshot. Use `desktop-check` and `doctor` when the window itself cannot
be seen.

For a live check using only the project's disposable test app, run
`python tools/smoke_computer_use.py --ocr`. It opens the test app, uses wad through MCP,
checks accessible content and OCR, and closes that test app.
