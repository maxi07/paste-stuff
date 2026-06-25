"""Tests for Paste Stuff.

These are designed to run in CI on a Windows runner, because ``main`` imports
Windows-only packages (``keyboard``, ``pywin32``, ``win32com``) at module level.
They exercise the pure, side-effect-free logic (config loading/validation,
text preview, command channel helpers) without touching the clipboard, the
keyboard, the registry or any window.
"""

import json

import pytest

import main


# --- _preview ---------------------------------------------------------------

def test_preview_short_text_unchanged():
    assert main._preview("hello") == "hello"


def test_preview_truncates_long_text():
    result = main._preview("x" * 100, length=40)
    assert result == "x" * 40 + "..."
    assert len(result) == 43


def test_preview_replaces_newlines_with_spaces():
    assert main._preview("line1\nline2") == "line1 line2"


# --- make_paste_callback ----------------------------------------------------

def test_make_paste_callback_returns_callable():
    callback = main.make_paste_callback("some text", "ctrl+shift+1")
    assert callable(callback)


# --- _startup_command -------------------------------------------------------

def test_startup_command_contains_quoted_script():
    command = main._startup_command()
    assert main.SCRIPT in command
    assert command.startswith('"')


# --- load_config ------------------------------------------------------------

def _write_config(tmp_path, content):
    path = tmp_path / "config.json"
    path.write_text(content, encoding="utf-8")
    return str(path)


def test_load_config_returns_mapping(tmp_path, monkeypatch):
    path = _write_config(
        tmp_path, json.dumps({"shortcuts": {"ctrl+shift+1": "hi"}}))
    monkeypatch.setattr(main, "CONFIG_PATH", path)
    assert main.load_config() == {"ctrl+shift+1": "hi"}


def test_load_config_coerces_values_to_str(tmp_path, monkeypatch):
    path = _write_config(
        tmp_path, json.dumps({"shortcuts": {"ctrl+shift+1": 123}}))
    monkeypatch.setattr(main, "CONFIG_PATH", path)
    assert main.load_config() == {"ctrl+shift+1": "123"}


def test_load_config_missing_file_exits(tmp_path, monkeypatch):
    monkeypatch.setattr(main, "CONFIG_PATH", str(tmp_path / "absent.json"))
    with pytest.raises(SystemExit) as excinfo:
        main.load_config()
    assert excinfo.value.code == 1


def test_load_config_invalid_json_exits(tmp_path, monkeypatch):
    path = _write_config(tmp_path, "{ this is not valid json ")
    monkeypatch.setattr(main, "CONFIG_PATH", path)
    with pytest.raises(SystemExit) as excinfo:
        main.load_config()
    assert excinfo.value.code == 2


def test_load_config_non_object_exits(tmp_path, monkeypatch):
    path = _write_config(tmp_path, json.dumps([1, 2, 3]))
    monkeypatch.setattr(main, "CONFIG_PATH", path)
    with pytest.raises(SystemExit) as excinfo:
        main.load_config()
    assert excinfo.value.code == 3


def test_load_config_shortcuts_wrong_type_exits(tmp_path, monkeypatch):
    path = _write_config(tmp_path, json.dumps({"shortcuts": "nope"}))
    monkeypatch.setattr(main, "CONFIG_PATH", path)
    with pytest.raises(SystemExit) as excinfo:
        main.load_config()
    assert excinfo.value.code == 3


def test_load_config_empty_shortcuts_exits(tmp_path, monkeypatch):
    path = _write_config(tmp_path, json.dumps({"shortcuts": {}}))
    monkeypatch.setattr(main, "CONFIG_PATH", path)
    with pytest.raises(SystemExit) as excinfo:
        main.load_config()
    assert excinfo.value.code == 3


def test_shipped_config_is_valid():
    """The config.json shipped in the repo must load and be non-empty."""
    shortcuts = main.load_config()
    assert isinstance(shortcuts, dict)
    assert len(shortcuts) >= 1
    assert all(isinstance(k, str) and isinstance(v, str)
               for k, v in shortcuts.items())


def test_load_config_at_limit_keeps_all(tmp_path, monkeypatch):
    shortcuts = {f"ctrl+shift+{i}": f"text {i}"
                 for i in range(main.MAX_SHORTCUTS)}
    path = _write_config(tmp_path, json.dumps({"shortcuts": shortcuts}))
    monkeypatch.setattr(main, "CONFIG_PATH", path)
    assert main.load_config() == shortcuts


def test_load_config_truncates_above_limit(tmp_path, monkeypatch):
    shortcuts = {f"key{i:02d}": f"text {i}"
                 for i in range(main.MAX_SHORTCUTS + 5)}
    path = _write_config(tmp_path, json.dumps({"shortcuts": shortcuts}))
    monkeypatch.setattr(main, "CONFIG_PATH", path)
    result = main.load_config()
    assert len(result) == main.MAX_SHORTCUTS


def test_load_config_truncation_keeps_first_in_order(tmp_path, monkeypatch):
    shortcuts = {f"key{i:02d}": f"text {i}"
                 for i in range(main.MAX_SHORTCUTS + 3)}
    path = _write_config(tmp_path, json.dumps({"shortcuts": shortcuts}))
    monkeypatch.setattr(main, "CONFIG_PATH", path)
    result = main.load_config()
    expected = dict(list(shortcuts.items())[:main.MAX_SHORTCUTS])
    assert result == expected


# --- _normalize_hotkey ------------------------------------------------------

def test_normalize_hotkey_lowercases_and_sorts():
    assert main._normalize_hotkey("Shift+Ctrl") == "ctrl+shift"


def test_normalize_hotkey_is_order_independent():
    assert (main._normalize_hotkey("ctrl+alt+delete")
            == main._normalize_hotkey("delete+ctrl+alt"))


def test_normalize_hotkey_applies_aliases():
    assert main._normalize_hotkey("win+l") == main._normalize_hotkey("windows+l")
    assert main._normalize_hotkey("control+esc") == "ctrl+escape"


def test_normalize_hotkey_strips_whitespace_and_blanks():
    assert main._normalize_hotkey(" ctrl +  shift ") == "ctrl+shift"


# --- reserved hotkeys -------------------------------------------------------

def test_reserved_hotkeys_contains_known_combos():
    assert main._normalize_hotkey("ctrl+alt+delete") in main._RESERVED_HOTKEYS
    assert main._normalize_hotkey("alt+f4") in main._RESERVED_HOTKEYS
    assert main._normalize_hotkey("win+l") in main._RESERVED_HOTKEYS


def test_reserved_hotkey_matches_regardless_of_spelling():
    # Win+L written as "windows+l" with reordered parts must still match.
    assert main._normalize_hotkey("l+windows") in main._RESERVED_HOTKEYS


def test_ordinary_hotkey_is_not_reserved():
    assert main._normalize_hotkey("ctrl+shift+1") not in main._RESERVED_HOTKEYS


# --- command channel --------------------------------------------------------

def test_handle_command_unknown_does_not_raise():
    # An unrecognised command should only be logged, never crash the server.
    main.handle_command("DEFINITELY_NOT_A_COMMAND")


def test_send_command_returns_false_when_no_server():
    # Nothing is listening on the loopback port during the test run.
    assert main.send_command("RELOAD", timeout=0.5) is False


# --- run_action("about") ----------------------------------------------------

def test_about_action_opens_repo_url(monkeypatch):
    import webbrowser

    opened = []
    monkeypatch.setattr(webbrowser, "open", opened.append)
    main.run_action("about", None)
    assert opened == [main.REPO_URL]


def test_about_action_swallows_browser_errors(monkeypatch):
    import webbrowser

    def boom(_url):
        raise RuntimeError("no browser")

    monkeypatch.setattr(webbrowser, "open", boom)
    monkeypatch.setattr(main, "notify_error", lambda *a, **k: None)
    # A failing browser launch must be handled, never propagate.
    main.run_action("about", None)


# --- _app_version -----------------------------------------------------------

def test_app_version_returns_cached_value(monkeypatch):
    monkeypatch.setattr(main, "_version_cache", "v9.9.9")
    # A cached value short-circuits before any file or git lookup.
    assert main._app_version() == "v9.9.9"


def test_app_version_reads_version_txt_in_base_dir(tmp_path, monkeypatch):
    (tmp_path / "version.txt").write_text("v1.2.3\n", encoding="utf-8")
    monkeypatch.setattr(main, "_version_cache", None)
    monkeypatch.setattr(main, "BASE_DIR", str(tmp_path))
    monkeypatch.setattr(main, "BUNDLE_DIR", str(tmp_path))
    assert main._app_version() == "v1.2.3"


def test_app_version_reads_version_txt_in_bundle_dir(tmp_path, monkeypatch):
    base = tmp_path / "base"
    bundle = tmp_path / "bundle"
    base.mkdir()
    bundle.mkdir()
    (bundle / "version.txt").write_text("v4.5.6", encoding="utf-8")
    monkeypatch.setattr(main, "_version_cache", None)
    monkeypatch.setattr(main, "BASE_DIR", str(base))
    monkeypatch.setattr(main, "BUNDLE_DIR", str(bundle))
    assert main._app_version() == "v4.5.6"


def test_app_version_strips_utf8_bom(tmp_path, monkeypatch):
    (tmp_path / "version.txt").write_text("v7.0.0", encoding="utf-8-sig")
    monkeypatch.setattr(main, "_version_cache", None)
    monkeypatch.setattr(main, "BASE_DIR", str(tmp_path))
    monkeypatch.setattr(main, "BUNDLE_DIR", str(tmp_path))
    assert main._app_version() == "v7.0.0"


def test_app_version_ignores_empty_version_txt_and_uses_git(
        tmp_path, monkeypatch):
    import subprocess

    (tmp_path / "version.txt").write_text("   \n", encoding="utf-8")
    monkeypatch.setattr(main, "_version_cache", None)
    monkeypatch.setattr(main, "BASE_DIR", str(tmp_path))
    monkeypatch.setattr(main, "BUNDLE_DIR", str(tmp_path))

    class _Result:
        returncode = 0
        stdout = "v2.0.0\n"

    monkeypatch.setattr(subprocess, "run", lambda *a, **k: _Result())
    assert main._app_version() == "v2.0.0"


def test_app_version_falls_back_to_git_tag(tmp_path, monkeypatch):
    import subprocess

    monkeypatch.setattr(main, "_version_cache", None)
    monkeypatch.setattr(main, "BASE_DIR", str(tmp_path))
    monkeypatch.setattr(main, "BUNDLE_DIR", str(tmp_path))

    class _Result:
        returncode = 0
        stdout = "v3.1.4\n"

    monkeypatch.setattr(subprocess, "run", lambda *a, **k: _Result())
    assert main._app_version() == "v3.1.4"


def test_app_version_git_nonzero_returns_dev(tmp_path, monkeypatch):
    import subprocess

    monkeypatch.setattr(main, "_version_cache", None)
    monkeypatch.setattr(main, "BASE_DIR", str(tmp_path))
    monkeypatch.setattr(main, "BUNDLE_DIR", str(tmp_path))

    class _Result:
        returncode = 128
        stdout = ""

    monkeypatch.setattr(subprocess, "run", lambda *a, **k: _Result())
    assert main._app_version() == "dev"


def test_app_version_git_missing_returns_dev(tmp_path, monkeypatch):
    import subprocess

    monkeypatch.setattr(main, "_version_cache", None)
    monkeypatch.setattr(main, "BASE_DIR", str(tmp_path))
    monkeypatch.setattr(main, "BUNDLE_DIR", str(tmp_path))

    def _missing(*_a, **_k):
        raise FileNotFoundError("git not installed")

    monkeypatch.setattr(subprocess, "run", _missing)
    assert main._app_version() == "dev"


def test_app_version_git_timeout_returns_dev(tmp_path, monkeypatch):
    import subprocess

    monkeypatch.setattr(main, "_version_cache", None)
    monkeypatch.setattr(main, "BASE_DIR", str(tmp_path))
    monkeypatch.setattr(main, "BUNDLE_DIR", str(tmp_path))

    def _timeout(*_a, **_k):
        raise subprocess.TimeoutExpired(cmd="git", timeout=2)

    monkeypatch.setattr(subprocess, "run", _timeout)
    assert main._app_version() == "dev"
