"""The browser from the inside: Chrome / Edge through the DevTools protocol (CDP).

A web page seen through the accessibility tree is thousands of nodes and every step is a
tree walk. Through CDP wad asks the page itself: a snapshot of what can be clicked is one
JavaScript call, clicks are real (trusted) mouse events at the element, typing is real
text input that React/Angular forms accept, and every action is read back.

wad starts its own browser with its own profile (never the person's everyday profile,
their cookies and logins) and a debugging port bound to 127.0.0.1 only. Any program of
the same user can reach that port while the browser is open - close it when done."""
import base64
import json
import os
import shutil
import time
import urllib.request

from . import state
from .commands import spawn
from .registry import Arg, WadError, command

BROWSER_FILE = os.path.join(state.STATE_DIR, "browser.json")
EXES = {
    "edge": [r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe",
             r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"],
    "chrome": [r"%ProgramFiles%\Google\Chrome\Application\chrome.exe",
               r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe",
               r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"],
}

# Page-side helper, sent with each call: find ONE element for a target, or say why not.
FIND_JS = r"""
(function (target, nth) {
  const vis = e => { const r = e.getBoundingClientRect(), s = getComputedStyle(e);
    return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none'; };
  const label = e => {
    const box = /^(checkbox|radio)$/i.test(e.type || '');
    const tied = e.labels && e.labels.length ? e.labels[0].innerText : '';
    return (e.getAttribute('aria-label') || tied || (e.tagName === 'INPUT' ? '' : e.innerText) ||
      e.getAttribute('placeholder') || e.getAttribute('title') || e.getAttribute('alt') ||
      (box ? '' : e.value) || '').trim().replace(/\s+/g, ' '); };
  let hits;
  if (/^b\d+$/.test(target)) {
    hits = [...document.querySelectorAll('[data-wad-ref="' + target + '"]')];
  } else if (target.startsWith('text=')) {
    const want = target.slice(5).trim().toLowerCase();
    const all = [...document.querySelectorAll('a,button,input,select,textarea,label,summary,' +
      '[role],[onclick],[tabindex],h1,h2,h3,h4,li,td,th,span,div,p')].filter(vis);
    hits = all.filter(e => label(e).toLowerCase() === want);
    if (!hits.length) hits = all.filter(e => label(e).toLowerCase().includes(want));
    // the innermost element that carries the text, not every ancestor around it
    hits = hits.filter(e => !hits.some(o => o !== e && e.contains(o)));
  } else {
    try { hits = [...document.querySelectorAll(target)]; }
    catch (err) { return {error: 'BAD_SELECTOR', message: String(err)}; }
    const shown = hits.filter(vis); if (shown.length) hits = shown;
  }
  if (!hits.length) return {count: 0};
  if (hits.length > 1 && !nth) return {count: hits.length,
    list: hits.slice(0, 5).map(e => e.tagName.toLowerCase() + ' ' + JSON.stringify(label(e).slice(0, 40)))};
  const el = hits[(nth || 1) - 1];
  if (!el) return {count: hits.length};
  document.querySelectorAll('[data-wad-target]').forEach(e => e.removeAttribute('data-wad-target'));
  el.setAttribute('data-wad-target', '1');
  el.scrollIntoView({block: 'center', inline: 'center'});
  const r = el.getBoundingClientRect();
  return {count: hits.length, tag: el.tagName.toLowerCase(), name: label(el).slice(0, 80),
          x: r.left + r.width / 2, y: r.top + r.height / 2, w: r.width, h: r.height,
          disabled: !!el.disabled, editable: el.isContentEditable || /^(input|textarea|select)$/i.test(el.tagName),
          value: ('value' in el) ? String(el.value) : el.innerText};
})
"""

SNAPSHOT_JS = r"""
(function (limit) {
  const vis = e => { const r = e.getBoundingClientRect(), s = getComputedStyle(e);
    return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none'; };
  const label = e => {
    const box = /^(checkbox|radio)$/i.test(e.type || '');
    const tied = e.labels && e.labels.length ? e.labels[0].innerText : '';
    return (e.getAttribute('aria-label') || tied || (e.tagName === 'INPUT' ? '' : e.innerText) ||
      e.getAttribute('placeholder') || e.getAttribute('title') || e.getAttribute('alt') ||
      (box ? '' : e.value) || '').trim().replace(/\s+/g, ' '); };
  document.querySelectorAll('[data-wad-ref]').forEach(e => e.removeAttribute('data-wad-ref'));
  const els = [...document.querySelectorAll('a[href],button,input:not([type=hidden]),select,' +
    'textarea,summary,[role=button],[role=link],[role=tab],[role=menuitem],[role=checkbox],' +
    '[role=option],[onclick],[contenteditable=true]')].filter(vis);
  const out = [];
  for (const e of els.slice(0, limit)) {
    const ref = 'b' + (out.length + 1);
    e.setAttribute('data-wad-ref', ref);
    out.push({ref, tag: e.tagName.toLowerCase(), role: e.getAttribute('role') || '',
      type: e.getAttribute('type') || '', name: label(e).slice(0, 80),
      value: ('value' in e && e.tagName !== 'BUTTON') ? String(e.value).slice(0, 80) : null,
      checked: ('checked' in e) ? e.checked : null, disabled: !!e.disabled,
      href: e.getAttribute('href') || null});
  }
  return {title: document.title, url: location.href, total: els.length, elements: out};
})
"""


# what a click can change: the address, the text, and the state of every form field
PAGE_STATE_JS = ("location.href + '|' + (document.body ? document.body.innerText : '') + '|' + "
                 "[...document.querySelectorAll('input,select,textarea')].map(e => "
                 "/^(checkbox|radio)$/.test(e.type) ? e.checked : e.value).join('\\u0001')")


def _ws():
    try:
        import websocket
        return websocket
    except ImportError:
        raise WadError("MISSING_DEPENDENCY", "browser commands need websocket-client",
                       "pip install websocket-client") from None


# Never through a proxy: on company PCs HTTP_PROXY is often set, and a request for
# 127.0.0.1 sent to the proxy hangs or reaches the wrong machine.
_LOCAL = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def _http(port, path, method="GET"):
    req = urllib.request.Request(f"http://127.0.0.1:{port}{path}", method=method)
    with _LOCAL.open(req, timeout=5) as r:
        body = r.read().decode("utf-8")
    try:
        return json.loads(body or "null")
    except ValueError:
        return body                              # /json/activate answers in plain text


def _config():
    cfg = state.read_json(BROWSER_FILE, ("NO_BROWSER", "no browser started by wad",
                                         "run `wad browser-launch` first"))
    try:
        _http(cfg["port"], "/json/version")
    except Exception:
        raise WadError("NO_BROWSER", "the wad browser is not running (closed?)",
                       "run `wad browser-launch` again") from None
    return cfg


def _pages(port):
    return [t for t in _http(port, "/json/list") if t.get("type") == "page"]


class CDP:
    """One DevTools connection to one tab: send a command, wait for its answer."""

    def __init__(self, ws_url):
        self.ws = _ws().create_connection(ws_url, timeout=30, suppress_origin=True,
                                          http_no_proxy=["127.0.0.1", "localhost"])
        self.n = 0

    def call(self, method, **params):
        self.n += 1
        self.ws.send(json.dumps({"id": self.n, "method": method, "params": params}))
        while True:
            msg = json.loads(self.ws.recv())
            if msg.get("id") == self.n:
                if "error" in msg:
                    raise WadError("BROWSER_ERROR", f"{method}: {msg['error'].get('message')}")
                return msg.get("result", {})

    def eval(self, expr, await_promise=False):
        r = self.call("Runtime.evaluate", expression=expr, returnByValue=True,
                      awaitPromise=await_promise, userGesture=True)
        if "exceptionDetails" in r:
            d = r["exceptionDetails"]
            text = (d.get("exception") or {}).get("description") or d.get("text")
            raise WadError("BROWSER_ERROR", f"the page threw: {text}")
        return r.get("result", {}).get("value")

    def close(self):
        try:
            self.ws.close()
        except Exception:
            pass


def tab():
    """A connection to the current tab (the one wad last used, else the first)."""
    cfg = _config()
    pages = _pages(cfg["port"])
    if not pages:
        raise WadError("NO_BROWSER", "the wad browser has no open tab",
                       "`wad browser-open <url> --new-tab`")
    cur = next((p for p in pages if p["id"] == cfg.get("tab")), pages[0])
    return CDP(cur["webSocketDebuggerUrl"]), cur


def _find(cdp, target, nth=None):
    r = cdp.eval(f"{FIND_JS}({json.dumps(target)}, {json.dumps(nth)})")
    if r.get("error"):
        raise WadError("BAD_SELECTOR", f"{target!r}: {r['message']}",
                       "a CSS selector, text=Visible text, or a ref (b12) from browser-snapshot")
    if not r.get("count"):
        raise WadError("ELEMENT_NOT_FOUND", f"nothing on the page matches {target!r}",
                       "browser-snapshot to see what is there; the page may still be loading")
    if "x" not in r:
        raise WadError("AMBIGUOUS_TARGET", f"{r['count']} elements match {target!r}: "
                       + "; ".join(r.get("list", [])), "add --nth N or use a more exact selector")
    return r


def _wait_text(cdp, text, gone=False, timeout=5.0):
    deadline = time.time() + timeout
    probe = f"document.body ? document.body.innerText.includes({json.dumps(text)}) : false"
    while True:
        try:
            present = bool(cdp.eval(probe))
        except WadError:
            present = False                      # mid-navigation: no document yet
        if present != gone:
            return True
        if time.time() >= deadline:
            return False
        time.sleep(0.2)


def _settle(cdp, timeout=10.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if cdp.eval("document.readyState") == "complete":
                return True
        except WadError:
            pass
        time.sleep(0.2)
    return False


def _find_exe(name):
    for p in EXES[name]:
        full = os.path.expandvars(p)
        if os.path.exists(full):
            return full
    return shutil.which("msedge" if name == "edge" else "chrome")


def browser_args():
    return [Arg("nth", int, help="which match, when several do")]


@command("browser-launch", "start Edge or Chrome with its own profile and a local-only "
         "DevTools port, for the browser-* commands",
         Arg("url", positional=True, optional=True, help="page to open"),
         Arg("browser", choices=["edge", "chrome"], default="edge"),
         Arg("port", int, default=9222), Arg("headless", bool, help="no visible window"),
         Arg("exe", help="path to another Chromium-based browser (Brave, Chromium, ...)"),
         group="browser")
def cmd_browser_launch(args):
    try:
        _http(args.port, "/json/version")
        running = True
    except Exception:
        running = False
    if not running:
        exe = args.exe or _find_exe(args.browser)
        if not exe or not os.path.exists(exe) and not shutil.which(exe):
            raise WadError("APP_NOT_FOUND", f"{args.browser} is not installed",
                           "try --browser " + ("chrome" if args.browser == "edge" else "edge"))
        profile = os.path.join(state.STATE_DIR, f"browser-profile-{args.browser}")
        cmd = [exe, f"--remote-debugging-port={args.port}", "--remote-debugging-address=127.0.0.1",
               f"--user-data-dir={profile}", "--no-first-run", "--no-default-browser-check",
               "--disable-features=msEdgeSidebarV2,msHubApps"]
        if os.environ.get("WAD_BROWSER_NO_SANDBOX"):     # containers running as root only
            cmd.append("--no-sandbox")
        if args.headless:
            cmd.append("--headless=new")
        cmd.append(args.url or "about:blank")
        spawn(cmd)
        deadline = time.time() + 30
        while True:
            try:
                _http(args.port, "/json/version")
                break
            except Exception:
                if time.time() > deadline:
                    raise WadError("TIMEOUT", "the browser did not open its DevTools port",
                                   "is another browser using the port? try --port 9333") from None
                time.sleep(0.3)
    state.write_json(BROWSER_FILE, {"port": args.port, "browser": args.browser})
    cdp, page = tab()
    try:
        if running and args.url:
            cdp.call("Page.navigate", url=args.url)
        _settle(cdp)
        title, url = cdp.eval("document.title"), cdp.eval("location.href")
    finally:
        cdp.close()
    state.write_json(BROWSER_FILE, {"port": args.port, "browser": args.browser, "tab": page["id"]})
    state.trace("browser-launch", {"browser": args.browser, "url": url})
    return ({"ok": True, "title": title, "url": url, "port": args.port},
            f"{args.browser} {'already running' if running else 'started'} on 127.0.0.1:{args.port}"
            f" - {title!r} {url}")


@command("browser-tabs", "list the wad browser's tabs; --use N makes tab N the current one",
         Arg("use", int, help="1-based tab number to make current"),
         group="browser", readonly=True)
def cmd_browser_tabs(args):
    cfg = _config()
    pages = _pages(cfg["port"])
    if args.use:
        if not 1 <= args.use <= len(pages):
            raise WadError("NOT_FOUND", f"there are {len(pages)} tabs")
        cfg["tab"] = pages[args.use - 1]["id"]
        _http(cfg["port"], f"/json/activate/{cfg['tab']}")
        state.write_json(BROWSER_FILE, cfg)
    rows = [{"n": i, "title": p.get("title"), "url": p.get("url"), "current": p["id"] == cfg.get("tab")}
            for i, p in enumerate(pages, 1)]
    return {"ok": True, "tabs": rows}, "\n".join(
        f"{'*' if r['current'] else ' '} {r['n']} {r['title']!r} {r['url']}" for r in rows)


@command("browser-open", "go to a URL in the current tab (or --new-tab) and wait for it to load",
         Arg("url", positional=True), Arg("new-tab", bool),
         Arg("expect", help="fail unless this text appears on the page"),
         Arg("timeout", float, default=15.0), group="browser")
def cmd_browser_open(args):
    cfg = _config()
    if args.new_tab:
        page = _http(cfg["port"], f"/json/new?{urllib.request.quote(args.url, safe='')}", "PUT")
        cfg["tab"] = page["id"]
        state.write_json(BROWSER_FILE, cfg)
    cdp, _ = tab()
    try:
        if not args.new_tab:
            cdp.call("Page.navigate", url=args.url)
            time.sleep(0.3)
        loaded = _settle(cdp, args.timeout)
        if args.expect and not _wait_text(cdp, args.expect, timeout=args.timeout):
            raise WadError("VERIFY_FAILED", f"{args.expect!r} did not appear on {args.url}")
        title, url = cdp.eval("document.title"), cdp.eval("location.href")
    finally:
        cdp.close()
    state.trace("browser-open", {"url": url})
    return ({"ok": True, "title": title, "url": url, "loaded": loaded},
            f"opened {url} - {title!r}" + ("" if loaded else " (still loading)"))


@command("browser-snapshot", "the page's clickable and fillable elements, with refs (b12)",
         Arg("limit", int, default=300), group="browser", readonly=True)
def cmd_browser_snapshot(args):
    cdp, _ = tab()
    try:
        r = cdp.eval(f"{SNAPSHOT_JS}({args.limit})")
    finally:
        cdp.close()
    lines = [f"page {r['title']!r} {r['url']}  ({r['total']} interactive"
             + (f", first {args.limit} shown)" if r["total"] > args.limit else ")")]
    for e in r["elements"]:
        kind = e["role"] or (f"{e['tag']}[{e['type']}]" if e["type"] else e["tag"])
        extra = []
        if e["value"] not in (None, ""):
            extra.append(f"value={e['value']!r}")
        if e["checked"] is not None and e["type"] in ("checkbox", "radio"):
            extra.append("checked" if e["checked"] else "unchecked")
        if e["disabled"]:
            extra.append("(disabled)")
        if e["href"] and e["tag"] == "a":
            extra.append(f"-> {e['href'][:60]}")
        lines.append(f"{e['ref']} {kind} {e['name']!r} {' '.join(extra)}".rstrip())
    return {"ok": True, **r}, "\n".join(lines)


def _mouse_click(cdp, x, y, count=1):
    for kind in ("mouseMoved", "mousePressed", "mouseReleased"):
        params = {"type": kind, "x": x, "y": y}
        if kind != "mouseMoved":
            params.update(button="left", clickCount=count)
        cdp.call("Input.dispatchMouseEvent", **params)


@command("browser-click", "click an element of the page with a real (trusted) mouse event",
         Arg("target", positional=True, help="b12 | text=Sign in | a CSS selector"),
         *browser_args(),
         Arg("expect", help="fail unless this text appears on the page afterwards"),
         Arg("timeout", float, default=5.0), group="browser")
def cmd_browser_click(args):
    cdp, _ = tab()
    try:
        el = _find(cdp, args.target, args.nth)
        if el["disabled"]:
            raise WadError("NOT_ENABLED", f"{el['tag']} {el['name']!r} is disabled")
        before = cdp.eval(PAGE_STATE_JS)
        _mouse_click(cdp, el["x"], el["y"])
        time.sleep(0.3)
        _settle(cdp, 10)
        if args.expect and not _wait_text(cdp, args.expect, timeout=args.timeout):
            raise WadError("VERIFY_FAILED", f"{args.expect!r} did not appear after the click",
                           "browser-snapshot / browser-text to see what the page shows")
        after = cdp.eval(PAGE_STATE_JS)
        url = cdp.eval("location.href")
    finally:
        cdp.close()
    changed = before != after
    state.trace("browser-click", {"target": args.target, "name": el["name"], "changed": changed})
    return ({"ok": True, "tag": el["tag"], "name": el["name"], "changed": changed, "url": url},
            f"clicked {el['tag']} {el['name']!r}" + ("" if changed or args.expect else
                                                   "\n  (the page did not visibly change)"))


@command("browser-type", "type into a page field like a person (works with React/Angular "
         "forms); read back",
         Arg("target", positional=True, help="b12 | text=Email | a CSS selector"),
         Arg("text", positional=True), *browser_args(),
         Arg("append", bool, help="add to what is there"),
         Arg("submit", bool, help="press Enter afterwards"), group="browser")
def cmd_browser_type(args):
    cdp, _ = tab()
    try:
        el = _find(cdp, args.target, args.nth)
        if not el["editable"]:
            raise WadError("NOT_EDITABLE", f"{el['tag']} {el['name']!r} is not a text field",
                           "target the input itself (browser-snapshot shows input/textarea)")
        cdp.eval("(() => { const e = document.querySelector('[data-wad-target]'); e.focus();"
                 + ("" if args.append else
                    " if (e.select) e.select(); else document.execCommand('selectAll');")
                 + " })()")
        if not args.append:
            cdp.call("Input.dispatchKeyEvent", type="keyDown", key="Backspace", code="Backspace",
                     windowsVirtualKeyCode=8)
            cdp.call("Input.dispatchKeyEvent", type="keyUp", key="Backspace", code="Backspace",
                     windowsVirtualKeyCode=8)
        cdp.call("Input.insertText", text=args.text)
        have = cdp.eval("(() => { const e = document.querySelector('[data-wad-target]');"
                        " return ('value' in e) ? String(e.value) : e.innerText; })()")
        want = (el["value"] or "") + args.text if args.append else args.text
        if have != want and args.text not in (have or ""):
            raise WadError("VERIFY_FAILED", f"the field holds {have!r}, not {want!r}",
                           "the page may reformat or reject input; check with browser-snapshot")
        if args.submit:
            for kind in ("keyDown", "keyUp"):
                cdp.call("Input.dispatchKeyEvent", type=kind, key="Enter", code="Enter",
                         windowsVirtualKeyCode=13, **({"text": "\r"} if kind == "keyDown" else {}))
            time.sleep(0.3)
            _settle(cdp, 10)
    finally:
        cdp.close()
    state.trace("browser-type", {"target": args.target, "chars": len(args.text)})
    return ({"ok": True, "value": have},
            f"typed {len(args.text)} chars into {el['tag']} {el['name']!r} - verified")


@command("browser-text", "the visible text of the page (or of one element)",
         Arg("target", positional=True, optional=True, help="b12 | text=... | CSS (default: page)"),
         Arg("max-chars", int, default=20000), group="browser", readonly=True)
def cmd_browser_text(args):
    cdp, _ = tab()
    try:
        if args.target:
            _find(cdp, args.target)
            text = cdp.eval("document.querySelector('[data-wad-target]').innerText")
        else:
            text = cdp.eval("document.body ? document.body.innerText : ''")
    finally:
        cdp.close()
    text = text or ""
    cut = len(text) > args.max_chars
    return ({"ok": True, "text": text[:args.max_chars], "truncated": cut},
            text[:args.max_chars] + ("\n[truncated]" if cut else ""))


@command("browser-wait", "wait until text is on the page (or --gone)",
         Arg("text", positional=True), Arg("gone", bool), Arg("timeout", float, default=10.0),
         group="browser", readonly=True)
def cmd_browser_wait(args):
    cdp, _ = tab()
    try:
        ok = _wait_text(cdp, args.text, args.gone, args.timeout)
    finally:
        cdp.close()
    if not ok:
        raise WadError("TIMEOUT", f"{args.text!r} {'still there' if args.gone else 'not on the page'}"
                       f" after {args.timeout}s")
    return {"ok": True}, f"{args.text!r} is {'gone' if args.gone else 'on the page'}"


@command("browser-eval", "run JavaScript in the page and return its (JSON) result",
         Arg("script", positional=True, help="an expression; wrap statements in (() => {...})()"),
         group="browser")
def cmd_browser_eval(args):
    cdp, _ = tab()
    try:
        value = cdp.eval(args.script, await_promise=True)
    finally:
        cdp.close()
    state.trace("browser-eval", {"chars": len(args.script)})
    return {"ok": True, "value": value}, json.dumps(value, ensure_ascii=False, indent=1)


@command("browser-screenshot", "capture the page as PNG (the viewport)",
         Arg("out", positional=True, optional=True), group="browser", readonly=True, image=True)
def cmd_browser_screenshot(args):
    cdp, _ = tab()
    try:
        data = cdp.call("Page.captureScreenshot", format="png")["data"]
    finally:
        cdp.close()
    os.makedirs(state.SHOTS_DIR, exist_ok=True)
    path = os.path.abspath(args.out or os.path.join(state.SHOTS_DIR,
                                                    f"page-{time.strftime('%H%M%S')}.png"))
    with open(path, "wb") as fh:
        fh.write(base64.b64decode(data))
    return {"ok": True, "path": path}, f"saved {path}"


@command("browser-close", "close the wad browser (and its DevTools port)", group="browser")
def cmd_browser_close(args):
    cfg = _config()
    cdp = CDP(_http(cfg["port"], "/json/version")["webSocketDebuggerUrl"])
    try:
        cdp.call("Browser.close")
    except Exception:
        pass
    finally:
        cdp.close()
    try:
        os.remove(BROWSER_FILE)
    except OSError:
        pass
    return {"ok": True}, "browser closed"
