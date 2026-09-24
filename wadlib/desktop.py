"""Check whether wad is attached to the desktop the person is using."""

from . import uia, win32
from .registry import WadError, command


HINT = ("run wad or `wad mcp` in the signed-in user's interactive desktop session; "
        "an isolated desktop cannot be repaired by creating a file or named pipe")


def require_interactive():
    info = win32.desktop_info()
    if info["accessible"] is False:
        raise WadError("DESKTOP_UNAVAILABLE",
                       f"wad is on desktop {info['desktop']!r}, while user input goes to "
                       f"{info['input_desktop']!r}", HINT)
    return info


@command("desktop-check", "check whether wad can see the user's interactive Windows "
         "desktop (useful when windows is empty or a Computer Use helper is unavailable)",
         group="observe", readonly=True)
def cmd_desktop_check(args):
    info = win32.desktop_info()
    if info["accessible"] is False:
        raise WadError("DESKTOP_UNAVAILABLE",
                       f"wad is on desktop {info['desktop']!r}, while user input goes to "
                       f"{info['input_desktop']!r}", HINT)
    visible = len(uia.top_windows()) if info["accessible"] is True else None
    text = (f"desktop ready: {info['station']}\\{info['desktop']} "
            f"({visible} visible windows)" if info["accessible"] is True else
            "desktop access could not be determined on this platform")
    return {"ok": True, **info, "visible_windows": visible}, text
