"""Read-only inventory of visible windows and their accessible contents."""

from . import uia, vision
from .registry import Arg, WadError, command, window_args


def _accessible(window, depth, limit, detailed):
    elements = uia.walk(window, [], depth, [1], [], limit=limit)
    visible = [e for e in elements if not e["offscreen"]]
    result = {"hwnd": window.NativeWindowHandle, "title": window.Name,
              "accessible_count": len(visible), "limited": len(elements) >= limit or
              any(e.get("truncated") for e in elements),
              "interactive_count": sum(bool(e["interactive"]) for e in visible)}
    if detailed:
        rows = []
        for e in visible:
            row = {k: e[k] for k in ("role", "name", "aid", "rect", "interactive",
                                      "enabled", "depth")}
            # Password controls must never be read back, even in an inspection report.
            value = e.get("value")
            if value:
                row["value"] = str(value)[:300]
            rows.append(row)
        result["elements"] = rows
    else:
        result["sample"] = [e["name"][:80] for e in visible
                            if e["name"] and e["depth"] > 0][:8]
    return result


@command("inspect", "check what wad can detect inside one window, or summarize all visible "
         "windows; optionally read pixels with OCR or a vision model",
         *window_args(), Arg("all", bool, help="summarize every visible window"),
         Arg("depth", int, help="tree depth (default 1 for --all, 4 for one window)"),
         Arg("limit", int, help="maximum elements per window (default 40 for --all, 200 for one)"),
         Arg("visual", bool, help="save a screenshot and try local Windows OCR"),
         Arg("model", bool, help="also send the screenshot to the configured OmniParser server"),
         Arg("url", help="OmniParser server URL (only with --model)"),
         group="observe", readonly=True)
def cmd_inspect(args):
    depth = args.depth if args.depth is not None else (1 if args.all else 4)
    limit = args.limit if args.limit is not None else (40 if args.all else 200)
    if depth < 0 or not 1 <= limit <= uia.MAX_ELEMENTS:
        raise WadError("USAGE", "depth must be at least 0 and limit must be 1 to 4000")
    if args.all and (args.window or args.hwnd or args.visual or args.model or args.url):
        raise WadError("USAGE", "--all cannot be combined with a target or visual options",
                       "inspect one window with --window to use --visual or --model")
    if args.url and not args.model:
        raise WadError("USAGE", "--url requires --model")
    if args.all:
        windows = []
        for window in uia.top_windows():
            try:
                uia.responsive(window)
                windows.append(_accessible(window, depth, limit, False))
            except Exception as exc:
                windows.append({"hwnd": window.NativeWindowHandle, "title": window.Name,
                                "error": str(exc)})
        payload = {"ok": True, "windows": windows, "count": len(windows),
                   "method": "Windows accessibility", "depth": depth, "limit": limit}
        lines = [f"{w['hwnd']:>10}  {w['title'][:55]!r}: " +
                 (w.get("error") or f"{w['accessible_count']} elements, "
                  f"{w['interactive_count']} interactive") for w in windows]
        return payload, "\n".join(lines) or "(no visible windows)"

    window = uia.find_window(args.window, args.hwnd)
    result = _accessible(window, depth, limit, True)
    result["coverage"] = {"accessibility": "available", "screenshot": "not requested",
                          "ocr": "not requested", "model": "not requested"}
    warnings = []
    if args.visual or args.model:
        try:
            shot = vision.capture(hwnd=window.NativeWindowHandle, max_width=0)
            result["screenshot"] = shot["path"]
            result["coverage"]["screenshot"] = "available"
        except WadError as exc:
            result["coverage"]["screenshot"] = "unavailable"
            warnings.append(f"screenshot: {exc.message}; {exc.hint}")
            shot = None
        if shot and args.visual:
            try:
                result["ocr_lines"] = vision.ocr(shot)
                result["coverage"]["ocr"] = "available"
            except WadError as exc:
                result["coverage"]["ocr"] = "unavailable"
                warnings.append(f"OCR: {exc.message}; {exc.hint}")
        if shot and args.model:
            try:
                result["visual_elements"] = vision.detect(shot, args.url or vision.DETECT_URL)
                result["coverage"]["model"] = "available"
            except WadError as exc:
                result["coverage"]["model"] = "unavailable"
                warnings.append(f"model: {exc.message}; {exc.hint}")
    result["warnings"] = warnings
    lines = [f"{result['title']!r}: {result['accessible_count']} accessible elements, "
             f"{result['interactive_count']} interactive"]
    lines += [f"  {'  ' * e['depth']}{e['role']} {e['name'][:80]!r}" +
              (f" = {e['value'][:80]!r}" if "value" in e else "")
              for e in result["elements"] if e["name"] or e.get("value")]
    if "screenshot" in result:
        lines.append(f"screenshot: {result['screenshot']}")
    if "ocr_lines" in result:
        lines.append(f"OCR: {len(result['ocr_lines'])} lines")
    if "visual_elements" in result:
        lines.append(f"model: {len(result['visual_elements'])} elements")
    lines += [f"warning: {w}" for w in warnings]
    return {"ok": True, **result}, "\n".join(lines)
