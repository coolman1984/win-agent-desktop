"""One declaration per command feeds the command line, the MCP server, batch files and
the generated command reference - so the four can never disagree."""
import argparse


class WadError(Exception):
    """A failure the agent can act on: a stable code, what happened, what to do next."""

    def __init__(self, code, message, hint=""):
        super().__init__(message)
        self.code, self.message, self.hint = code, message, hint


class Arg:
    """One parameter. `type` is str, int, float or bool (a bool is a --flag)."""

    def __init__(self, name, type=str, help="", positional=False, short=None, choices=None,
                 default=None, required=False, optional=False):
        self.name, self.type, self.help = name, type, help
        self.positional, self.short, self.choices = positional, short, choices
        self.default, self.required, self.optional = default, required, optional

    @property
    def dest(self):
        return self.name.replace("-", "_")

    @property
    def is_required(self):
        return self.required or (self.positional and not self.optional)


class Command:
    def __init__(self, name, fn, help, args, group, mcp, readonly, image, long):
        self.name, self.fn, self.help, self.args = name, fn, help, args
        self.group, self.mcp, self.readonly, self.image = group, mcp, readonly, image
        self.long = long            # runs for as long as it was asked to: no watchdog

    def namespace(self, values):
        """argparse-style namespace from a dict (MCP arguments, a batch step)."""
        known = {a.dest for a in self.args} | {"json"}
        unknown = set(k.replace("-", "_") for k in values) - known
        if unknown:
            raise WadError("USAGE", f"{self.name}: unknown argument(s) {', '.join(sorted(unknown))}",
                           f"arguments: {', '.join(a.name for a in self.args) or '(none)'}")
        ns = argparse.Namespace(cmd=self.name, json=bool(values.get("json")))
        given = {k.replace("-", "_"): v for k, v in values.items()}
        for a in self.args:
            v = given.get(a.dest, a.default if a.type is not bool else bool(a.default))
            if v is None and a.is_required:
                raise WadError("USAGE", f"{self.name}: {a.name} is required", a.help)
            if v is not None and a.type in (int, float) and not isinstance(v, bool):
                try:
                    v = a.type(v)
                except (TypeError, ValueError):
                    raise WadError("USAGE", f"{self.name}: {a.name} must be a number") from None
            if a.choices and v is not None and v not in a.choices:
                raise WadError("USAGE", f"{self.name}: {a.name} must be one of {a.choices}")
            setattr(ns, a.dest, v)
        return ns


COMMANDS = {}
GROUPS = ["observe", "act", "mouse", "keyboard", "windows", "vision", "browser", "office",
          "system", "workflow"]


def command(name, help, *args, group="act", mcp=True, readonly=False, image=False,
            long=False):
    def deco(fn):
        COMMANDS[name] = Command(name, fn, help, list(args), group, mcp, readonly, image, long)
        return fn
    return deco


# Parameter sets shared by many commands --------------------------------------
def window_args():
    return [Arg("window", short="-w", help="window title (exact, else a unique substring)"),
            Arg("hwnd", int, help="window handle (from `windows`)")]


def target_arg(help="a ref from the last snapshot (e12) or a selector (role=Button name=Save)"):
    return Arg("target", positional=True, help=help)


def verify_args():
    return [Arg("expect", help="fail unless this text appears in the app afterwards"),
            Arg("expect-gone", help="fail unless this text disappears from the app afterwards"),
            Arg("timeout", float, default=5.0, help="seconds to wait for --expect/--expect-gone"),
            Arg("no-verify", bool, help="skip the after-action check (faster, blind)")]


# Front ends ---------------------------------------------------------------------
def build_parser(description=""):
    p = argparse.ArgumentParser(prog="wad", description=description,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--json", action="store_true", help="structured JSON output")
    sub = p.add_subparsers(dest="cmd", required=True, metavar="COMMAND")
    for c in COMMANDS.values():
        sp = sub.add_parser(c.name, help=c.help, description=c.help)
        sp.add_argument("--json", action="store_true", default=argparse.SUPPRESS,
                        help="structured JSON output")
        for a in c.args:
            kw = {"help": a.help}
            if a.type is bool:
                kw["action"] = "store_true"
            else:
                kw["type"] = a.type
                if a.choices:
                    kw["choices"] = a.choices
                if a.default is not None:
                    kw["default"] = a.default
            if a.positional:
                if a.optional:
                    kw["nargs"] = "?"
                sp.add_argument(a.dest, **kw)
            else:
                flags = [f"--{a.name}"] + ([a.short] if a.short else [])
                if a.required:
                    kw["required"] = True
                sp.add_argument(*flags, dest=a.dest, **kw)
        sp.set_defaults(fn=c.fn)
    return p


JSON_TYPES = {str: "string", int: "integer", float: "number", bool: "boolean"}


def input_schema(c):
    props, required = {}, []
    for a in c.args:
        prop = {"type": JSON_TYPES[a.type], "description": a.help}
        if a.choices:
            prop["enum"] = list(a.choices)
        if a.default is not None and a.type is not bool:
            prop["default"] = a.default
        props[a.dest] = prop
        if a.is_required:
            required.append(a.dest)
    schema = {"type": "object", "properties": props}
    if required:
        schema["required"] = required
    return schema
