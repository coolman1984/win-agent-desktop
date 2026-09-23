"""A small, faithful-enough stand-in for the `uiautomation` package, so wad's logic
(selectors, refs, verification, the MCP protocol, batch replay...) is tested on any OS.

It models one fake app with the behaviours that matter: a button that opens a dialog, a
field whose value pattern lies (like Excel's), a checkbox, a combo box, a disabled
button, two buttons with the same name, a password box and a scrollable list."""
import itertools
import types

_ids = itertools.count(1)


class PatternId:
    InvokePattern = 10000
    SelectionPattern = 10001
    ValuePattern = 10002
    RangeValuePattern = 10003
    ScrollPattern = 10004
    ExpandCollapsePattern = 10005
    WindowPattern = 10009
    SelectionItemPattern = 10010
    ScrollItemPattern = 10017
    LegacyIAccessiblePattern = 10018
    TextPattern = 10014
    TogglePattern = 10015
    TransformPattern = 10016


class Keys:
    pass


for _i, _n in enumerate(["VK_CONTROL", "VK_MENU", "VK_SHIFT", "VK_LWIN", "VK_RETURN", "VK_TAB",
                         "VK_ESCAPE", "VK_SPACE", "VK_BACK", "VK_DELETE", "VK_INSERT", "VK_HOME",
                         "VK_END", "VK_UP", "VK_DOWN", "VK_LEFT", "VK_RIGHT", "VK_PRIOR",
                         "VK_NEXT", "VK_APPS", "VK_CAPITAL", "VK_SNAPSHOT"]):
    setattr(Keys, _n, 0x100 + _i)
Keys.VK_F1 = 0x70


class Rect:
    def __init__(self, l, t, r, b):
        self.left, self.top, self.right, self.bottom = l, t, r, b


class Gone(Exception):
    pass


class Control:
    def __init__(self, role, name="", aid="", cls="", children=(), patterns=None, enabled=True,
                 offscreen=False, rect=(0, 0, 100, 20), hwnd=0, pid=100, password=False):
        self._role, self._name, self.AutomationId, self.ClassName = role, name, aid, cls
        self.patterns = patterns or {}
        self._enabled, self.IsOffscreen = enabled, offscreen
        self.BoundingRectangle = Rect(*rect)
        self.NativeWindowHandle, self.ProcessId, self.IsPassword = hwnd, pid, password
        self.rid = [42, next(_ids)]
        self.parent = None
        self.children = []
        self.dead = False
        for c in children:
            self.add(c)

    def add(self, c):
        c.parent = self
        self.children.append(c)
        return c

    def remove(self):
        self.dead = True
        if self.parent:
            self.parent.children.remove(self)

    def _check(self):
        if self.dead:
            raise Gone("element not available")

    @property
    def ControlTypeName(self):
        self._check()
        return self._role + "Control"

    @property
    def Name(self):
        self._check()
        return self._name

    @Name.setter
    def Name(self, v):
        self._name = v

    @property
    def IsEnabled(self):
        return self._enabled

    @property
    def HasKeyboardFocus(self):
        return FOCUS[0] is self

    def GetChildren(self):
        self._check()
        return list(self.children)

    def GetFirstChildControl(self):
        return self.children[0] if self.children else None

    def GetParentControl(self):
        return self.parent

    def GetPattern(self, pid):
        return self.patterns.get(pid)

    def GetRuntimeId(self):
        return self.rid

    def SetFocus(self):
        FOCUS[0] = self
        return True

    def SetActive(self):
        return True

    def CaptureToImage(self, path):
        open(path, "wb").close()


# --- patterns ---------------------------------------------------------------
class Invoke:
    def __init__(self, fn=lambda: None):
        self.fn = fn

    def Invoke(self):
        self.fn()


class Value:
    def __init__(self, value="", readonly=False, lies=False):
        self._value, self.IsReadOnly, self.lies = value, readonly, lies

    @property
    def Value(self):
        return self._value

    def SetValue(self, v):
        if not self.lies:
            self._value = v


class Toggle:
    def __init__(self, state=0):
        self.ToggleState = state

    def Toggle(self):
        self.ToggleState = 0 if self.ToggleState else 1


class ExpandCollapse:
    def __init__(self):
        self.ExpandCollapseState = 0

    def Expand(self):
        self.ExpandCollapseState = 1

    def Collapse(self):
        self.ExpandCollapseState = 0


class SelectionItem:
    def __init__(self, on_select=None):
        self.IsSelected = False
        self.on_select = on_select

    def Select(self):
        self.IsSelected = True
        if self.on_select:
            self.on_select()


class Scroll:
    def __init__(self):
        self.VerticalScrollPercent = 0.0
        self.HorizontalScrollPercent = 0.0

    def Scroll(self, h, v):
        step = {4: 10, 3: 40, 1: -10, 0: -40, 2: 0}
        self.VerticalScrollPercent = min(100, max(0, self.VerticalScrollPercent + step[v]))
        self.HorizontalScrollPercent = min(100, max(0, self.HorizontalScrollPercent + step[h]))


class Window:
    def __init__(self, ctrl):
        self.ctrl = ctrl
        self.state = 0

    def Close(self):
        self.ctrl.remove()

    def SetWindowVisualState(self, s):
        self.state = s


# --- the fake desktop --------------------------------------------------------
ROOT = Control("Pane", "Desktop 1")
FOCUS = [None]
LOG = []


def build():
    """A fresh desktop with the fake app on it; returns the named parts."""
    global ROOT
    ROOT = Control("Pane", "Desktop 1")
    FOCUS[0] = None
    LOG.clear()
    parts = {}

    main = Control("Window", "Untitled - Notepad", cls="Notepad", hwnd=1001, pid=100,
                   rect=(0, 0, 800, 600))
    main.patterns[PatternId.WindowPattern] = Window(main)
    doc_value = Value("")
    doc = main.add(Control("Document", "Text editor", aid="RichEditD2DPT",
                           patterns={PatternId.ValuePattern: doc_value}, rect=(0, 50, 800, 500)))

    def open_dialog():
        dlg = Control("Window", "Save As", cls="#32770", hwnd=1002, pid=100)
        dlg.patterns[PatternId.WindowPattern] = Window(dlg)
        cancel = dlg.add(Control("Button", "Cancel", aid="2"))
        cancel.patterns[PatternId.InvokePattern] = Invoke(dlg.remove)
        ROOT.add(dlg)
        parts["dialog"] = dlg

    save = main.add(Control("Button", "Save", aid="SaveButton",
                            patterns={PatternId.InvokePattern: Invoke(open_dialog)}))
    noop = main.add(Control("Button", "Nothing", patterns={PatternId.InvokePattern: Invoke()}))
    wrap = main.add(Control("CheckBox", "Word wrap", patterns={PatternId.TogglePattern: Toggle()}))
    combo_value = Value("Arial", readonly=True)
    combo = main.add(Control("ComboBox", "Font", patterns={
        PatternId.ExpandCollapsePattern: ExpandCollapse(), PatternId.ValuePattern: combo_value}))
    for font in ("Arial", "Consolas", "Courier New"):
        def pick(font=font):
            combo_value._value = font
        combo.add(Control("ListItem", font, patterns={
            PatternId.SelectionItemPattern: SelectionItem(pick)}))
    dead = main.add(Control("Button", "Dead", enabled=False,
                            patterns={PatternId.InvokePattern: Invoke()}))
    left = main.add(Control("Pane", "Left"))
    right = main.add(Control("Pane", "Right"))
    ok1 = left.add(Control("Button", "OK", patterns={PatternId.InvokePattern: Invoke()}))
    ok2 = right.add(Control("Button", "OK", patterns={PatternId.InvokePattern: Invoke()}))
    liar = main.add(Control("Edit", "Cell", patterns={PatternId.ValuePattern: Value("", lies=True)}))
    pw = main.add(Control("Edit", "Password", password=True,
                          patterns={PatternId.ValuePattern: Value("")}))
    lst = main.add(Control("List", "Items", patterns={PatternId.ScrollPattern: Scroll()}))
    lst.add(Control("ListItem", "row 1", patterns={PatternId.SelectionItemPattern: SelectionItem()}))
    ROOT.add(main)
    other = Control("Window", "Other App", hwnd=2001, pid=200)
    ROOT.add(other)
    parts.update(main=main, doc=doc, save=save, noop=noop, wrap=wrap, combo=combo, dead=dead,
                 ok1=ok1, ok2=ok2, liar=liar, pw=pw, list=lst, other=other)
    return parts


def _all(node):
    yield node
    for c in node.children:
        yield from _all(c)


# --- module-level API --------------------------------------------------------
def GetRootControl():
    return ROOT


def ControlFromHandle(h):
    for c in _all(ROOT):
        if c.NativeWindowHandle == h and not c.dead:
            return c
    return None


def WalkControl(root, includeTop=False, maxDepth=0xFFFFFFFF):
    def rec(node, depth):
        if depth > maxDepth:
            return
        if depth > 0 or includeTop:
            yield node, depth
        for c in list(node.children):
            yield from rec(c, depth + 1)
    yield from rec(root, 0)


def GetFocusedControl():
    return FOCUS[0] or ROOT


def ControlFromPoint(x, y):
    return None


class UIAutomationInitializerInThread:
    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


_clip = [""]


def GetClipboardText():
    return _clip[0]


def SetClipboardText(t):
    _clip[0] = t


def _log(name):
    return lambda *a, **k: LOG.append((name, a))


PressKey = _log("press")
ReleaseKey = _log("release")
SendKeys = _log("sendkeys")
Click = _log("click")
RightClick = _log("rightclick")
MiddleClick = _log("middleclick")
MoveTo = _log("moveto")
DragDrop = _log("dragdrop")
WheelDown = _log("wheeldown")
WheelUp = _log("wheelup")
