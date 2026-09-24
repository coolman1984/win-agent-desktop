# Seeing pixels: screenshots, OCR and the smart eye

Most apps are driven through their accessibility tree. When an app shows almost nothing
there (games, canvases, remote desktops, video, some custom-drawn apps), wad has three
ways to work from the picture instead. Try them in this order:

| Tool | Needs | Good at | Speed |
|---|---|---|---|
| `screenshot` | Pillow | the calling model looks at the PNG itself (MCP returns the image) | instant |
| `ocr`, `click-text` | `requirements-ocr.txt` (Windows' built-in OCR) | anything with words on it | ~1 s |
| `detect`, `click-mark` | an OmniParser server | icons and buttons **without** words | 1-20 s |

Every position is in the last screenshot's pixels, and wad maps it back to the screen
(window position, scaling, several monitors). Real clicks from all of these go through
the same guards as the rest of wad: the app must be in front, and the point must be
inside it.

## OCR

```
python wad.py ocr --window "Game"                    # every line with its center
python wad.py ocr --window "Game" --find "Start"     # just where "Start" is
python wad.py click-text "Start" --window "Game" --expect "Level 1"
```

Windows OCR is made for documents; wad reads a 2x enlarged copy so that small UI text
(a 9 pt button caption) is found too. A whole-word match wins over a word that only
contains the text ("Submit" beats "Submitted:"). `--lang ar` reads Arabic if that
language's OCR pack is installed (Settings > Time & language > Language).

## The smart eye (OmniParser)

[OmniParser](https://github.com/microsoft/OmniParser) is Microsoft's open vision model
that finds interactive elements - icons, buttons, fields - in a screenshot, with a short
description of each. wad does not ship the model (it is several GB and wants a GPU); it
talks to an OmniParser server you run, on this PC or another one.

Run the server (once, on a machine with an NVIDIA GPU; CPU works but takes ~10-20 s per
screenshot):

```bash
git clone https://github.com/microsoft/OmniParser && cd OmniParser
pip install -r requirements.txt
# download the V2 weights into weights/ as the OmniParser README describes, then:
cd omnitool/omniparserserver
python -m omniparserserver --device cuda --port 8000
```

Then in wad:

```
set WAD_OMNIPARSER_URL=http://127.0.0.1:8000          (or pass --url each time)
python wad.py detect --window "Game" -i --marks       # v1, v2 ... with centers
python wad.py click-mark v7 --expect "Paused"
```

`detect` returns each element's ref (`v1`...), kind (icon / text), whether the model
thinks it is clickable, its box and center; `--marks` saves a copy of the picture with
the refs drawn on it. `click-mark` clicks the center of a ref from the last `detect`.

The screenshot goes to the server you configured. Point `WAD_OMNIPARSER_URL` only at a
machine you trust with what is on your screen.
