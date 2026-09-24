"""Stand-ins for the slices of Outlook's and PowerPoint's object models wad uses, so the
Office commands are tested without Office (CI runners have none)."""
import itertools

_ids = itertools.count(1000)


class Coll:
    """A 1-based COM collection: Count, Item(i), and calling it like coll(i)."""

    def __init__(self, items=()):
        self.items = list(items)

    @property
    def Count(self):
        return len(self.items)

    def Item(self, i):
        return self.items[i - 1]

    __call__ = Item


# --- Outlook -----------------------------------------------------------------------
class Attachment:
    def __init__(self, path):
        self.FileName = path.replace("\\", "/").rsplit("/", 1)[-1]


class Attachments(Coll):
    def Add(self, path):
        self.items.append(Attachment(path))


class Mail:
    Class = 43

    def __init__(self, store, subject="", sender="", body="", received="2026-09-24 09:00",
                 unread=False):
        self.store, self.Subject, self.SenderName, self.Body = store, subject, sender, body
        self.SenderEmailAddress = sender.lower().replace(" ", ".") + "@example.com"
        self.To, self.CC, self.ReceivedTime, self.UnRead = "me@example.com", "", received, unread
        self.Attachments = Attachments()
        self.EntryID = f"ID{next(_ids)}"
        self.sent = self.displayed = False
        store[self.EntryID] = self

    def Save(self):
        self.store[self.EntryID] = self

    def Send(self):
        self.sent = True

    def Display(self, modal=False):
        self.displayed = True


class Items(Coll):
    def Sort(self, key, descending):
        self.items.sort(key=lambda m: m.ReceivedTime, reverse=descending)

    def Restrict(self, flt):
        assert flt == "[UnRead] = True"
        return Items([m for m in self.items if m.UnRead])


class Folder:
    def __init__(self, mails):
        self.mails = mails

    @property
    def Items(self):
        return Items(self.mails)


class Outlook:
    def __init__(self):
        self.store = {}
        self.inbox = [Mail(self.store, "Invoice March", "Sara Ali", "Please pay.", "2026-09-22 10:00"),
                      Mail(self.store, "Lunch?", "Omar", "12:30?", "2026-09-23 11:00", unread=True),
                      Mail(self.store, "Report ready", "Sara Ali", "See attached.", "2026-09-24 08:00")]
        self.inbox[2].Attachments.Add("C:/files/report.xlsx")

    def GetNamespace(self, name):
        return self

    def GetDefaultFolder(self, n):
        return Folder(self.inbox if n == 6 else [])

    def GetItemFromID(self, entry):
        return self.store[entry]

    def CreateItem(self, kind):
        return Mail(self.store)


# --- PowerPoint ----------------------------------------------------------------------
class TextRange:
    def __init__(self, text=""):
        self.Text = text

    def Find(self, what):
        return self if what in self.Text else None

    def Replace(self, what, new):
        if what not in self.Text:
            return None
        self.Text = self.Text.replace(what, new, 1)
        return self


class Shape:
    def __init__(self, text=None):
        self.HasTextFrame = text is not None
        self.TextFrame = type("TF", (), {})()
        self.TextFrame.TextRange = TextRange(text or "")

    @property
    def _has(self):
        return bool(self.TextFrame.TextRange.Text)


class Shapes(Coll):
    def __init__(self, items, has_title=True):
        super().__init__(items)
        self.HasTitle = has_title
        for s in items:
            s.TextFrame.HasText = True

    @property
    def Title(self):
        return self.items[0]

    @property
    def Placeholders(self):
        return Coll(self.items)


class Slide:
    def __init__(self, title, *body, layout=2):
        shapes = [Shape(title)] + ([Shape("\r".join(body))] if layout != 11 else [])
        self.Shapes = Shapes(shapes, has_title=layout != 12)


class Slides(Coll):
    def Add(self, at, layout):
        s = Slide("", layout=layout)
        self.items.insert(at - 1, s)
        return s


class Presentation:
    def __init__(self, name):
        self.Name, self.FullName = name, "C:/decks/" + name
        self.Slides = Slides([Slide("Q3 results", "Revenue up 12%", "Costs flat"),
                              Slide("Next steps", "Hire 2 engineers")])
        self.saved_as = None

    def Save(self):
        self.saved_as = self.FullName

    def SaveAs(self, path, fmt=None):
        self.saved_as, self.FullName = path, path


class PowerPoint:
    def __init__(self):
        self.Presentations = Coll([Presentation("Board.pptx")])

    @property
    def ActivePresentation(self):
        return self.Presentations.items[0]
