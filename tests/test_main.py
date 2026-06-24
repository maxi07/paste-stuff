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


# --- command channel --------------------------------------------------------

def test_handle_command_unknown_does_not_raise():
    # An unrecognised command should only be logged, never crash the server.
    main.handle_command("DEFINITELY_NOT_A_COMMAND")


def test_send_command_returns_false_when_no_server():
    # Nothing is listening on the loopback port during the test run.
    assert main.send_command("RELOAD", timeout=0.5) is False
