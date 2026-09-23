"""Where wad keeps what must survive between commands: the last snapshot (so refs work
across separate processes), the last screenshot's geometry, the trace and the policy.

WAD_STATE_DIR moves the whole directory; WAD_SESSION gives an agent its own snapshot and
screenshot files, so two agents on one PC never resolve each other's refs."""
import json
import os
import re
from datetime import datetime

from .registry import WadError

STATE_DIR = os.environ.get("WAD_STATE_DIR") or os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "win-agent-desktop")
_session = re.sub(r"[^A-Za-z0-9_-]", "", os.environ.get("WAD_SESSION", ""))
_suffix = f"-{_session}" if _session else ""
SNAPSHOT_FILE = os.path.join(STATE_DIR, f"last_snapshot{_suffix}.json")
SHOT_FILE = os.path.join(STATE_DIR, f"last_screenshot{_suffix}.json")
TRACE_FILE = os.path.join(STATE_DIR, "trace.jsonl")
POLICY_FILE = os.path.join(STATE_DIR, "policy.json")
SHOTS_DIR = os.path.join(STATE_DIR, "shots")


def now_iso():
    return datetime.now().isoformat(timespec="milliseconds")


def trace(command, detail):
    try:
        os.makedirs(STATE_DIR, exist_ok=True)
        with open(TRACE_FILE, "a", encoding="utf-8") as fh:
            fh.write(json.dumps({"t": now_iso(), "cmd": command, "session": _session or None,
                                 **detail}, ensure_ascii=False, default=str) + "\n")
    except OSError:
        pass            # a full disk must not turn a successful action into a failure


def read_json(path, missing):
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except FileNotFoundError:
        if missing is None:
            return None
        raise WadError(*missing) from None


def write_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False)
    os.replace(tmp, path)


def load_snapshot():
    return read_json(SNAPSHOT_FILE, ("NO_SNAPSHOT", "no snapshot taken yet",
                                     "run `wad snapshot --window <title>` first"))


def save_snapshot(snap):
    write_json(SNAPSHOT_FILE, snap)


def read_trace(last):
    try:
        with open(TRACE_FILE, encoding="utf-8") as fh:
            lines = fh.readlines()
    except FileNotFoundError:
        return []
    out = []
    for line in lines[-last:]:
        try:
            out.append(json.loads(line))
        except ValueError:
            continue
    return out
