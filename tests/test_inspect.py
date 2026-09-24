"""The combined inspection command is safe and reports partial visual coverage."""

from wadlib import vision
from wadlib.registry import WadError


def test_inspect_all_and_target(app, run):
    all_windows = run("inspect", "--all")
    assert all_windows["ok"] and all_windows["count"] >= 1
    assert all("accessible_count" in w or "error" in w for w in all_windows["windows"])

    one = run("inspect", "--window", "Untitled - Notepad")
    assert one["ok"]
    assert one["coverage"]["accessibility"] == "available"
    assert one["interactive_count"] >= 1
    assert any(e["name"] == "Save" for e in one["elements"])


def test_inspect_visual_reports_missing_ocr_without_losing_tree(app, run, monkeypatch):
    monkeypatch.setattr(vision, "capture", lambda **kwargs: {"path": "test.png"})

    def missing(_shot):
        raise WadError("MISSING_DEPENDENCY", "OCR unavailable", "install OCR package")

    monkeypatch.setattr(vision, "ocr", missing)
    result = run("inspect", "--window", "Untitled - Notepad", "--visual")
    assert result["ok"]
    assert result["coverage"]["ocr"] == "unavailable"
    assert result["screenshot"] == "test.png"
    assert result["accessible_count"] > 0
    assert "install OCR package" in result["warnings"][0]


def test_inspect_model_is_opt_in(app, run, monkeypatch):
    calls = []
    monkeypatch.setattr(vision, "capture", lambda **kwargs: {"path": "test.png"})
    monkeypatch.setattr(vision, "detect", lambda shot, url: calls.append(url) or
                        [{"type": "icon", "content": "Save"}])
    regular = run("inspect", "--window", "Untitled - Notepad")
    assert regular["ok"] and not calls
    result = run("inspect", "--window", "Untitled - Notepad", "--model", "--url",
                 "http://127.0.0.1:8000")
    assert result["ok"] and calls == ["http://127.0.0.1:8000"]
    assert result["visual_elements"][0]["content"] == "Save"


def test_inspect_rejects_visual_all(app, run):
    result = run("inspect", "--all", "--visual")
    assert not result["ok"] and result["code"] == "USAGE"


def test_inspect_never_reads_password_value(app, run):
    from fake_uia import PatternId

    password = next(c for c in app["main"].children if c.IsPassword)
    password.GetPattern(PatternId.ValuePattern)._value = "private test secret"
    result = run("inspect", "--window", "Untitled - Notepad")
    assert result["ok"]
    assert "private test secret" not in str(result)
