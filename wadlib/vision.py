"""The fallback eyes: screenshots, OCR and clicks by position.

For what the accessibility tree does not show - games, canvases, remote desktops, apps
that draw their own controls. The agent looks at the picture (and at OCR text with
boxes), then clicks by position. Positions are always in the LAST SCREENSHOT's pixels,
so an agent can use exactly what it sees; wad maps them back to the screen itself."""
import asyncio
import os
import time

import uiautomation as auto

from . import inputs, state, uia, verify, win32
from .registry import Arg, WadError, command, window_args

MAX_WIDTH = 1600


def _pil():
    try:
        from PIL import Image, ImageDraw, ImageGrab
        return Image, ImageDraw, ImageGrab
    except ImportError:
        return None


def capture(window=None, hwnd=None, screen=False, region=None, out=None, marks=False,
            max_width=MAX_WIDTH):
    """Save a PNG and remember its geometry: origin on screen and scale."""
    os.makedirs(state.SHOTS_DIR, exist_ok=True)
    path = os.path.abspath(out or os.path.join(state.SHOTS_DIR,
                                               f"shot-{time.strftime('%H%M%S')}.png"))
    pil = _pil()
    w = None
    if screen:
        if pil is None:
            raise WadError("MISSING_DEPENDENCY", "full-screen capture needs Pillow",
                           "pip install Pillow")
        img = pil[2].grab(all_screens=True)
        # all_screens images start at the virtual screen's top-left, which can be negative
        origin = list(win32.virtual_screen_origin())
        box = None
        hwnd_used, title = None, "(screen)"
    else:
        w = uia.find_window(window, hwnd)
        try:
            win32.bring_front(w.NativeWindowHandle)
        except Exception:
            w.SetActive()
        time.sleep(0.3)
        box = l, t, r, b = uia.rect_of(w)
        origin = [l, t]
        hwnd_used, title = w.NativeWindowHandle, w.Name
        if pil is not None:
            img = pil[2].grab(bbox=(l, t, r, b), all_screens=True)
        else:
            w.CaptureToImage(path)
            img = None
    if region:
        try:
            rx, ry, rw, rh = (int(float(v)) for v in region.split(","))
        except ValueError:
            raise WadError("USAGE", "--region takes X,Y,WIDTH,HEIGHT (window pixels)") from None
        if img is None:
            raise WadError("MISSING_DEPENDENCY", "--region needs Pillow", "pip install Pillow")
        img = img.crop((rx, ry, rx + rw, ry + rh))
        origin = [origin[0] + rx, origin[1] + ry]
    scale = 1.0
    if img is not None:
        if marks:
            _draw_marks(img, origin, pil[1])
        if max_width and img.width > max_width:
            scale = max_width / img.width
            img = img.resize((max_width, round(img.height * scale)))
        img.save(path)
        size = [img.width, img.height]
    else:
        size = [box[2] - box[0], box[3] - box[1]]
    shot = {"path": path, "origin": origin, "scale": scale, "size": size,
            "hwnd": hwnd_used, "window": title, "taken": state.now_iso()}
    state.write_json(state.SHOT_FILE, shot)
    return shot


def _draw_marks(img, origin, ImageDraw):
    """Draw each interactive element of the last snapshot with its ref ("set of marks"),
    so the agent can pick a ref by looking at the picture."""
    try:
        snap = state.load_snapshot()
    except WadError:
        return
    d = ImageDraw.Draw(img)
    for e in snap["elements"]:
        if not e["interactive"] or e["offscreen"]:
            continue
        l, t, r, b = e["rect"]
        l, t, r, b = l - origin[0], t - origin[1], r - origin[0], b - origin[1]
        if r <= 0 or b <= 0 or l >= img.width or t >= img.height or r <= l or b <= t:
            continue
        d.rectangle([l, t, r, b], outline=(255, 0, 180), width=2)
        label = e["ref"]
        tw = 7 * len(label) + 4
        d.rectangle([l, max(0, t - 14), l + tw, max(14, t)], fill=(255, 0, 180))
        d.text((l + 2, max(0, t - 13)), label, fill=(255, 255, 255))


def to_screen(x, y, screen=False):
    """Map a point from the last screenshot's pixels to real screen pixels."""
    if screen:
        return int(x), int(y)
    shot = state.read_json(state.SHOT_FILE, ("NO_SCREENSHOT", "no screenshot taken yet",
                                             "run `wad screenshot` first, or pass --screen"))
    s = shot.get("scale") or 1.0
    return int(shot["origin"][0] + x / s), int(shot["origin"][1] + y / s)


@command("screenshot", "capture a window (or the screen) to PNG; coordinates in it can be "
         "used directly by click-xy",
         Arg("out", positional=True, optional=True, help="PNG path (default: state dir)"),
         *window_args(),
         Arg("screen", bool, help="the whole screen instead of a window"),
         Arg("region", help="X,Y,WIDTH,HEIGHT inside the window, to zoom in"),
         Arg("marks", bool, help="draw the last snapshot's refs on the picture (needs Pillow)"),
         Arg("max-width", int, default=MAX_WIDTH, help="downscale wider images (0 = never)"),
         group="vision", readonly=True, image=True)
def cmd_screenshot(args):
    shot = capture(args.window, args.hwnd, args.screen, args.region, args.out, args.marks,
                   args.max_width)
    state.trace("screenshot", {"window": shot["window"], "path": shot["path"]})
    note = f"  (scaled x{shot['scale']:.2f})" if shot["scale"] != 1.0 else ""
    return ({"ok": True, **shot},
            f"saved {shot['path']}  {shot['size'][0]}x{shot['size'][1]}{note} of {shot['window']!r}")


@command("click-xy", "real mouse click at a point of the last screenshot (or --screen pixels)",
         Arg("x", float, positional=True), Arg("y", float, positional=True),
         Arg("button", choices=["left", "right", "middle"], default="left"),
         Arg("double", bool, help="double click"),
         Arg("screen", bool, help="x,y are absolute screen pixels"),
         Arg("expect", help="fail unless this text appears afterwards"),
         Arg("timeout", float, default=5.0),
         group="vision")
def cmd_click_xy(args):
    x, y = to_screen(args.x, args.y, args.screen)
    hwnd = None
    if not args.screen:
        shot = state.read_json(state.SHOT_FILE, None) or {}
        hwnd = shot.get("hwnd")
    if hwnd:
        inputs.guard(hwnd)
        _check_point_in_app(hwnd, x, y)
    inputs.click_at(x, y, args.button, 2 if args.double else 1)
    if args.expect and hwnd:
        verify.expect(hwnd, args.expect, timeout=args.timeout)
    state.trace("click-xy", {"x": x, "y": y, "button": args.button})
    return {"ok": True, "screen_x": x, "screen_y": y}, f"clicked ({x}, {y}) on screen"


def _check_point_in_app(hwnd, x, y):
    main = auto.ControlFromHandle(hwnd)
    if main is None:
        return
    inputs.hit_check(main, x, y)


# ---------------------------------------------------------------------------
# OCR - Windows' own engine (Windows.Media.Ocr), no cloud, no extra model
# ---------------------------------------------------------------------------

def _ocr_modules():
    try:
        from winrt.windows.globalization import Language
        from winrt.windows.graphics.imaging import BitmapDecoder
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.storage import FileAccessMode, StorageFile
        return Language, BitmapDecoder, OcrEngine, FileAccessMode, StorageFile
    except ImportError:
        raise WadError("MISSING_DEPENDENCY", "OCR needs the Windows Runtime packages",
                       "pip install -r requirements-ocr.txt") from None


async def _ocr_file(path, lang):
    Language, BitmapDecoder, OcrEngine, FileAccessMode, StorageFile = _ocr_modules()
    f = await StorageFile.get_file_from_path_async(path)
    stream = await f.open_async(FileAccessMode.READ)
    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()
    if lang:
        engine = OcrEngine.try_create_from_language(Language(lang))
    else:
        engine = OcrEngine.try_create_from_user_profile_languages()
    if engine is None:
        raise WadError("OCR_LANGUAGE", f"no OCR engine for {lang or 'the user languages'}",
                       "install the language's OCR pack: Settings > Time & language > Language")
    result = await engine.recognize_async(bitmap)
    lines = []
    for line in result.lines:
        words = [{"text": w.text, "box": [w.bounding_rect.x, w.bounding_rect.y,
                                          w.bounding_rect.width, w.bounding_rect.height]}
                 for w in line.words]
        if not words:
            continue
        xs = [b["box"][0] for b in words]
        ys = [b["box"][1] for b in words]
        x2 = [b["box"][0] + b["box"][2] for b in words]
        y2 = [b["box"][1] + b["box"][3] for b in words]
        lines.append({"text": line.text, "words": words,
                      "box": [min(xs), min(ys), max(x2) - min(xs), max(y2) - min(ys)]})
    return lines


OCR_UPSCALE = 2
OCR_MAX_SIDE = 9000                     # Windows OCR refuses images above ~10000 px


def ocr(shot, lang=None):
    """Windows OCR is tuned for documents; UI text (a 9 pt button caption) is too small
    for it and simply disappears. Reading a 2x enlarged copy finds it; the boxes are
    scaled back, so positions stay in screenshot pixels."""
    path, factor = shot["path"], 1
    pil = _pil()
    if pil is not None:
        img = pil[0].open(path)
        if max(img.size) * OCR_UPSCALE <= OCR_MAX_SIDE:
            factor = OCR_UPSCALE
            path = os.path.splitext(shot["path"])[0] + "-ocr.png"
            img.resize((img.width * factor, img.height * factor),
                       pil[0].LANCZOS).save(path)
    lines = asyncio.run(_ocr_file(path, lang))
    for ln in lines:                    # back to screenshot pixels, plus centers
        for w in ln["words"]:
            w["box"] = [round(v / factor) for v in w["box"]]
        ln["box"] = [round(v / factor) for v in ln["box"]]
        x, y, w, h = ln["box"]
        ln["center"] = [round(x + w / 2), round(y + h / 2)]
    return lines


def find_text(lines, text, nth=1):
    """Matches in reading order; each narrowed to the fewest words that contain the text,
    so "Edit" in the line "File Edit View" is clicked on "Edit", not mid-line."""
    want = text.lower()
    hits = []
    for ln in lines:
        if want not in ln["text"].lower():
            continue
        words, span = ln["words"], None
        for size in range(1, len(words) + 1):
            for i in range(len(words) - size + 1):
                joined = " ".join(w["text"] for w in words[i:i + size])
                if want in joined.lower():
                    span = (joined, words[i:i + size])
                    break
            if span:
                break
        if span:
            boxes = [w["box"] for w in span[1]]
            x0, y0 = min(b[0] for b in boxes), min(b[1] for b in boxes)
            x1, y1 = max(b[0] + b[2] for b in boxes), max(b[1] + b[3] for b in boxes)
            hits.append({"text": span[0], "line": ln["text"],
                         "center": [round((x0 + x1) / 2), round((y0 + y1) / 2)]})
        else:                       # the text spans OCR's word split oddly: use the line
            hits.append({"text": ln["text"], "line": ln["text"], "center": ln["center"]})
    # a whole-word match beats a word that merely contains it ("Submit" vs "Submitted:")
    hits.sort(key=lambda h: 0 if h["text"].lower().strip(" :.,") == want else 1)
    if not hits:
        return None, hits
    return (hits[nth - 1] if nth <= len(hits) else None), hits


@command("ocr", "read the text of a window/screen from pixels (Windows OCR), with positions "
         "usable by click-xy",
         *window_args(), Arg("screen", bool), Arg("region", help="X,Y,WIDTH,HEIGHT in the window"),
         Arg("lang", help="OCR language tag, e.g. en, ar (default: the user's languages)"),
         Arg("find", help="only report where this text is"),
         group="vision", readonly=True)
def cmd_ocr(args):
    shot = capture(args.window, args.hwnd, args.screen, args.region, max_width=0)
    lines = ocr(shot, args.lang)
    if args.find:
        _, hits = find_text(lines, args.find)
        text = "\n".join(f"{h['center'][0]},{h['center'][1]}  {h['text']!r}  (line {h['line']!r})"
                         for h in hits) or f"(no {args.find!r} on screen)"
        return {"ok": True, "matches": hits, "screenshot": shot["path"]}, text
    text = "\n".join(f"{ln['center'][0]:>5},{ln['center'][1]:<5} {ln['text']}" for ln in lines)
    return ({"ok": True, "lines": lines, "screenshot": shot["path"]},
            text or "(no text recognised)")


@command("click-text", "find text on screen by OCR and click it (for apps with no "
         "accessibility tree)",
         Arg("text", positional=True), *window_args(),
         Arg("nth", int, default=1, help="which match, from the top-left"),
         Arg("button", choices=["left", "right"], default="left"),
         Arg("double", bool),
         Arg("lang", help="OCR language tag"),
         Arg("expect", help="fail unless this text appears afterwards"),
         Arg("timeout", float, default=5.0),
         group="vision")
def cmd_click_text(args):
    shot = capture(args.window, args.hwnd, max_width=0)
    lines = ocr(shot, args.lang)
    hit, hits = find_text(lines, args.text, args.nth)
    if hit is None:
        raise WadError("TEXT_NOT_FOUND", f"OCR found no {args.text!r}"
                       + (f" #{args.nth} (only {len(hits)})" if hits else ""),
                       "check `wad ocr` output; try --lang, or a shorter part of the text")
    x, y = to_screen(*hit["center"])
    inputs.guard(shot["hwnd"])
    _check_point_in_app(shot["hwnd"], x, y)
    inputs.click_at(x, y, args.button, 2 if args.double else 1)
    if args.expect:
        verify.expect(shot["hwnd"], args.expect, timeout=args.timeout)
    state.trace("click-text", {"text": args.text, "x": x, "y": y})
    return ({"ok": True, "text": hit["text"], "screen_x": x, "screen_y": y},
            f"clicked {hit['text']!r} at ({x}, {y})")


# ---------------------------------------------------------------------------
# Smart eye - a vision model that finds buttons and icons in pixels (OmniParser)
# ---------------------------------------------------------------------------
DETECT_URL = os.environ.get("WAD_OMNIPARSER_URL", "http://127.0.0.1:8000")
DETECT_FILE = os.path.join(state.STATE_DIR, "last_detect.json")


def _post_json(url, body, timeout):
    import json
    import urllib.parse
    import urllib.request
    host = urllib.parse.urlparse(url).hostname or ""
    handlers = [urllib.request.ProxyHandler({})] if host in ("127.0.0.1", "localhost") else []
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.build_opener(*handlers).open(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def detect(shot, url=DETECT_URL, timeout=120):
    """Send the screenshot to an OmniParser server; return elements in screenshot pixels."""
    import base64
    with open(shot["path"], "rb") as fh:
        b64 = base64.b64encode(fh.read()).decode("ascii")
    try:
        r = _post_json(url.rstrip("/") + "/parse/", {"base64_image": b64}, timeout)
    except Exception as e:
        raise WadError("DETECTOR_UNAVAILABLE", f"no OmniParser server at {url} ({e.__class__.__name__})",
                       "start one (see docs/VISION.md) or set WAD_OMNIPARSER_URL; without it use "
                       "ocr / click-text, or look at the screenshot yourself") from None
    w, h = shot["size"]
    out = []
    for item in r.get("parsed_content_list") or []:
        box = item.get("bbox") or [0, 0, 0, 0]
        if max(box) <= 1.0:                     # OmniParser answers in fractions of the image
            box = [box[0] * w, box[1] * h, box[2] * w, box[3] * h]
        x1, y1, x2, y2 = (round(v) for v in box)
        out.append({"ref": f"v{len(out) + 1}", "type": item.get("type", ""),
                    "content": (item.get("content") or "").strip(),
                    "interactive": bool(item.get("interactivity", True)),
                    "box": [x1, y1, x2 - x1, y2 - y1], "center": [(x1 + x2) // 2, (y1 + y2) // 2]})
    return out


def ocr_elements(shot):
    """detect's fallback: each OCR word as an element, in the same shape as the model's."""
    shot = dict(shot, size=shot["size"])
    out = []
    for ln in ocr(shot):
        for w in ln["words"]:
            x, y, bw, bh = w["box"]
            out.append({"ref": f"v{len(out) + 1}", "type": "text", "content": w["text"],
                        "interactive": True, "box": [x, y, bw, bh],
                        "center": [x + bw // 2, y + bh // 2]})
    return out


@command("detect", "find buttons, icons and text in the PIXELS with a vision model (OmniParser "
         "server) - for apps with no accessibility tree; refs v1.. work with click-mark",
         *window_args(), Arg("screen", bool), Arg("region", help="X,Y,WIDTH,HEIGHT in the window"),
         Arg("interactive", bool, short="-i", help="only elements the model thinks are clickable"),
         Arg("url", help="OmniParser server (default WAD_OMNIPARSER_URL or 127.0.0.1:8000)"),
         Arg("marks", bool, help="also save a picture with the v-refs drawn on it"),
         group="vision", readonly=True, image=True)
def cmd_detect(args):
    shot = capture(args.window, args.hwnd, args.screen, args.region)
    mode = "model"
    try:
        found = detect(shot, args.url or DETECT_URL)
    except WadError as e:
        if e.code != "DETECTOR_UNAVAILABLE" or args.url:
            raise
        # No vision server: still give every piece of text on screen a v-ref, from
        # Windows' own OCR - words are most of what gets clicked. Icons need the model.
        try:
            found = ocr_elements(shot)
        except WadError:
            raise e from None
        mode = "ocr"
    if args.interactive:
        found = [e for e in found if e["interactive"]]
    state.write_json(DETECT_FILE, {"shot": shot, "elements": found})
    path = shot["path"]
    pil = _pil()
    if args.marks and pil is not None:
        img = pil[0].open(path).convert("RGB")
        d = pil[1].Draw(img)
        for e in found:
            x, y, w, h = e["box"]
            d.rectangle([x, y, x + w, y + h], outline=(0, 160, 255), width=2)
            d.rectangle([x, max(0, y - 14), x + 7 * len(e["ref"]) + 4, max(14, y)], fill=(0, 160, 255))
            d.text((x + 2, max(0, y - 13)), e["ref"], fill=(255, 255, 255))
        path = os.path.splitext(path)[0] + "-marks.png"
        img.save(path)
    lines = [f"{e['ref']:<5} {e['type']:<5} {'*' if e['interactive'] else ' '} "
             f"{e['center'][0]:>5},{e['center'][1]:<5} {e['content'][:60]!r}" for e in found]
    head = f"{len(found)} elements (* = clickable)"
    if mode == "ocr":
        head += " - text only (no vision server; icons without words need one, docs/VISION.md)"
    return ({"ok": True, "elements": found, "path": path, "mode": mode},
            "\n".join([head] + lines))


@command("click-mark", "real mouse click on an element found by `detect` (v12), guarded",
         Arg("ref", positional=True, help="v-ref from the last detect"),
         Arg("button", choices=["left", "right"], default="left"), Arg("double", bool),
         Arg("expect", help="fail unless this text appears afterwards"),
         Arg("timeout", float, default=5.0), group="vision")
def cmd_click_mark(args):
    last = state.read_json(DETECT_FILE, ("NO_SNAPSHOT", "no detect result yet",
                                         "run `wad detect --window ...` first"))
    el = next((e for e in last["elements"] if e["ref"] == args.ref.lstrip("@")), None)
    if el is None:
        raise WadError("REF_NOT_FOUND", f"{args.ref} is not in the last detect result")
    shot = last["shot"]
    s = shot.get("scale") or 1.0
    x, y = int(shot["origin"][0] + el["center"][0] / s), int(shot["origin"][1] + el["center"][1] / s)
    if shot.get("hwnd"):
        inputs.guard(shot["hwnd"])
        _check_point_in_app(shot["hwnd"], x, y)
    inputs.click_at(x, y, args.button, 2 if args.double else 1)
    if args.expect and shot.get("hwnd"):
        verify.expect(shot["hwnd"], args.expect, timeout=args.timeout)
    state.trace("click-mark", {"ref": el["ref"], "content": el["content"], "x": x, "y": y})
    return ({"ok": True, "screen_x": x, "screen_y": y, "content": el["content"]},
            f"clicked {el['ref']} {el['content']!r} at ({x}, {y})")
