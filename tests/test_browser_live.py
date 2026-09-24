"""browser-* against a REAL Chromium (headless), through the real command line.

Runs wherever a Chromium-based browser is found (CI's ubuntu runner has Chrome; set
WAD_TEST_CHROME to point at one); skipped otherwise."""
import json
import os
import shutil
import subprocess
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CANDIDATES = [os.environ.get("WAD_TEST_CHROME", ""), "/opt/pw-browsers/chromium-1194/chrome-linux/chrome",
              shutil.which("google-chrome") or "", shutil.which("chromium") or "",
              shutil.which("chromium-browser") or ""]
CHROME = next((c for c in CANDIDATES if c and os.path.exists(c)), None)
pytestmark = pytest.mark.skipif(CHROME is None, reason="no Chromium-based browser here")
BOOT = ("import sys; sys.path[:0]=[%r, %r]; import fake_uia; "
        "sys.modules['uiautomation']=fake_uia; "
        "from wadlib.cli import main; sys.exit(main(sys.argv[1:]))"
        % (ROOT, os.path.join(ROOT, "tests")))


@pytest.fixture(scope="module")
def wad(tmp_path_factory):
    env = {**os.environ, "WAD_STATE_DIR": str(tmp_path_factory.mktemp("wadb")),
           "WAD_BROWSER_NO_SANDBOX": "1"}

    def run(*argv):
        p = subprocess.run([sys.executable, "-c", BOOT, "--json", *argv], capture_output=True,
                           text=True, env=env, timeout=90)
        return json.loads(p.stdout)
    url = "file://" + os.path.join(ROOT, "tools", "test_page.html").replace("\\", "/")
    out = run("browser-launch", url, "--exe", CHROME, "--headless", "--port", "9377")
    assert out["ok"], out
    yield run
    run("browser-close")


def test_snapshot_names_fields_by_their_labels(wad):
    names = [(e["ref"], e["tag"], e["name"]) for e in wad("browser-snapshot")["elements"]]
    assert ("b1", "input", "Email") in names and ("b2", "input", "I agree") in names


def test_type_click_expect_and_read_back(wad):
    assert wad("browser-type", "#email", "ahmed@example.com")["value"] == "ahmed@example.com"
    assert wad("browser-click", "text=I agree")["changed"] is True
    out = wad("browser-click", "text=Send", "--expect", "Hello ahmed@example.com / true")
    assert out["ok"], out
    assert wad("browser-text", "#out")["text"] == "Hello ahmed@example.com / true"


def test_refusals_are_explicit(wad):
    assert wad("browser-click", "text=Locked")["code"] == "NOT_ENABLED"
    assert wad("browser-click", "button")["code"] == "AMBIGUOUS_TARGET"
    assert wad("browser-click", "button", "--nth", "1")["ok"]
    assert wad("browser-click", "#nope")["code"] == "ELEMENT_NOT_FOUND"
    assert wad("browser-type", "text=Send", "x")["code"] == "NOT_EDITABLE"
    assert wad("browser-eval", "(() => { throw new Error('boom') })()")["code"] == "BROWSER_ERROR"


def test_tabs_open_and_switch(wad):
    wad("browser-open", "about:blank", "--new-tab")
    tabs = wad("browser-tabs")["tabs"]
    assert len(tabs) == 2 and next(t for t in tabs if t["current"])["url"] == "about:blank"
    first = next(t["n"] for t in tabs if "test_page" in t["url"])
    wad("browser-tabs", "--use", str(first))
    assert wad("browser-wait", "Second section")["ok"]
