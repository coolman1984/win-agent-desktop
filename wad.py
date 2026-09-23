"""wad - Windows Agent Desktop. Run `python wad.py --help`, or `python wad.py guide`.

The code lives in the wadlib package; this file keeps `python wad.py ...` and
`import wad` (used by the demos) working."""
import sys

from wadlib.cli import main
from wadlib.state import SNAPSHOT_FILE, STATE_DIR, TRACE_FILE, load_snapshot
from wadlib.uia import find_window, top_windows

__all__ = ["main", "load_snapshot", "top_windows", "find_window", "STATE_DIR",
           "SNAPSHOT_FILE", "TRACE_FILE"]

if __name__ == "__main__":
    sys.exit(main())
