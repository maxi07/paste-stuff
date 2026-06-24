"""
Paste Stuff - a tiny background app that pastes configured text snippets
into ANY Windows application.

It has NO window UI. Instead it lives as an icon in the Windows taskbar:

- Right-click the taskbar icon to get a menu (a Windows "Jump List", just like
  the Microsoft 365 app) listing every configured snippet plus the app actions
  (Edit config / Reload config / Run at startup / Quit).
- Or trigger any snippet from anywhere with its global keyboard shortcut.

Everything is configured through config.json (hotkey -> text).
When started from a console, every action is logged to that console.

Run modes:
    main.py                         -> the resident background app
    main.py --action paste --key K  -> tell the running app to paste snippet K
    main.py --action reload         -> reload config
    main.py --action autostart      -> toggle run-at-startup
    main.py --action edit           -> open config.json
    main.py --action quit           -> quit the running app
(The --action variants are what the taskbar Jump List entries launch.)
"""

import argparse
import ctypes
import json
import logging
import os
import socket
import sys
import threading
import time
import tkinter as tk

# Third-party dependencies (see requirements.txt). If any is missing the app
# cannot run, so show a clear Windows alert and exit instead of crashing
# silently -- important because the app normally runs without a console
# (pythonw.exe), where an uncaught ImportError would leave no trace at all.
try:
    import keyboard
    import pyperclip
    import pythoncom
    import win32api
    import win32con
    import win32gui
    import win32process
    from win32com.propsys import propsys, pscon
    from win32com.shell import shell
except ImportError as exc:
    _msg = (
        "Paste Stuff can't start because a required Python package is "
        f"missing:\n\n    {exc}\n\n"
        "Install the dependencies and try again:\n\n"
        "    pip install -r requirements.txt"
    )
    try:
        ctypes.windll.user32.MessageBoxW(
            0, _msg, "Paste Stuff \u2013 Missing dependency",
            0x10 | 0x10000 | 0x40000)  # ICONERROR | SETFOREGROUND | TOPMOST
    except Exception:
        pass
    print(_msg, file=sys.stderr)
    sys.exit(1)

APP_NAME = "Paste Stuff"
APP_ID = "MaxKrause.PasteStuff"  # AppUserModelID that owns the taskbar button.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
ICON_PATH = os.path.join(BASE_DIR, "icon.ico")
SCRIPT = os.path.join(BASE_DIR, "main.py")

_PYTHONW = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
LAUNCHER_EXE = _PYTHONW if os.path.exists(_PYTHONW) else sys.executable

# Local loopback channel the Jump List helpers use to talk to the running app.
HOST = "127.0.0.1"
PORT = 50573

# Registry key used for the optional "run at startup" feature.
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"

# Window classes we never treat as "the app the user was working in".
_SKIP_CLASSES = {
    "Shell_TrayWnd", "Shell_SecondaryTrayWnd", "WorkerW", "Progman",
    "NotifyIconOverflowWindow", "Windows.UI.Core.CoreWindow",
}

_hotkey_handles = []      # keyboard hooks we registered, so we can remove them.
_last_active_hwnd = None  # last foreground window that wasn't ours.
_root = None              # the (hidden) Tk window.

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(APP_NAME)


# --- User-facing alerts -----------------------------------------------------

# This app usually runs without a console (pythonw.exe), so log messages are
# invisible to the user. Windows MessageBox alerts are therefore how we surface
# problems. The flags below configure the icon and bring the box to the front.
_MB_OK = 0x00000000
_MB_ICONERROR = 0x00000010
_MB_ICONWARNING = 0x00000030
_MB_ICONINFORMATION = 0x00000040
_MB_SETFOREGROUND = 0x00010000
_MB_TOPMOST = 0x00040000


def _show_message_box(message, title, flags):
    try:
        ctypes.windll.user32.MessageBoxW(
            0, str(message), str(title),
            flags | _MB_SETFOREGROUND | _MB_TOPMOST)
    except Exception as exc:  # never let an alert failure crash the app.
        log.error("Could not display alert '%s': %s", title, exc)


def _notify(message, icon, title, block):
    """Show a Windows alert. Runs off-thread unless ``block`` is set."""
    if block:
        _show_message_box(message, title, _MB_OK | icon)
    else:
        threading.Thread(
            target=_show_message_box, args=(message, title, _MB_OK | icon),
            daemon=True).start()


def notify_error(message, title=None, block=False):
    """Pop up an error alert (the message is expected to be logged already)."""
    _notify(message, _MB_ICONERROR, title or f"{APP_NAME} \u2013 Error", block)


def notify_warning(message, title=None, block=False):
    """Pop up a warning alert (the message is expected to be logged already)."""
    _notify(message, _MB_ICONWARNING, title or f"{APP_NAME} \u2013 Warning", block)


# --- Config -----------------------------------------------------------------

# Signature of the last config problem we alerted about, so repeated reloads of
# a still-broken file don't spam the user with identical pop-ups.
_last_config_error = None


def load_config(notify=False):
    """Read config.json and return the {hotkey: text} mapping.

    A broken config never crashes the app: the problem is logged and (when
    ``notify`` is set) shown in a Windows alert, and an empty mapping is
    returned so the rest of the program keeps running.
    """
    global _last_config_error
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            raise ValueError("the top-level value must be a JSON object")
        shortcuts = data.get("shortcuts", {})
        if not isinstance(shortcuts, dict):
            raise ValueError("'shortcuts' must be an object of hotkey -> text")
        if len(shortcuts) == 0:
            raise ValueError("the 'shortcuts' object must not be empty")
        _last_config_error = None
        return {str(k): str(v) for k, v in shortcuts.items()}
    except FileNotFoundError:
        log.warning("config.json not found.")
        if notify and _last_config_error != "missing":
            notify_warning(
                f"config.json was not found at:\n{CONFIG_PATH}\n\n"
                "Create the file with at least one shortcut, then start "
                "Paste Stuff again.", block=True)
        _last_config_error = "missing"
        sys.exit(1)
    except json.JSONDecodeError as exc:
        log.error("Could not parse config.json: %s", exc)
        signature = f"json:{exc.lineno}:{exc.colno}:{exc.msg}"
        if notify and _last_config_error != signature:
            notify_error(
                "config.json is not valid JSON and could not be loaded:\n\n"
                f"{exc.msg} (line {exc.lineno}, column {exc.colno}).\n\n"
                "Fix the file (e.g. a missing comma, bracket or quote) and "
                "start Paste Stuff again.", block=True)
        _last_config_error = signature
        sys.exit(2)
    except ValueError as exc:
        log.error("config.json is misconfigured: %s", exc)
        signature = f"value:{exc}"
        if notify and _last_config_error != signature:
            notify_error(
                f"config.json is misconfigured:\n\n{exc}.\n\n"
                "Expected a structure like:\n"
                '{\n  "shortcuts": {\n    "ctrl+shift+1": "your text"\n  }\n}\n\n'
                "Fix the file and start Paste Stuff again.", block=True)
        _last_config_error = signature
        sys.exit(3)
    except OSError as exc:
        log.error("Could not read config.json: %s", exc)
        signature = f"os:{exc}"
        if notify and _last_config_error != signature:
            notify_error(
                f"config.json could not be read:\n\n{exc}\n\n"
                "Check the file's permissions, then start Paste Stuff again.",
                block=True)
        _last_config_error = signature
        sys.exit(4)


def _preview(text, length=40):
    preview = text.replace("\n", " ")
    return preview[:length] + "..." if len(preview) > length else preview


# --- Pasting ----------------------------------------------------------------

def _send_paste():
    """Release any held modifiers and send Ctrl+V to the focused window."""
    time.sleep(0.05)
    for mod in ("ctrl", "shift", "alt", "windows"):
        keyboard.release(mod)
    keyboard.send("ctrl+v")


def paste_text(text, hotkey=None):
    """Paste triggered by a global hotkey (target app is already focused)."""
    log.info("Hotkey '%s' triggered -> pasting: \"%s\"", hotkey, _preview(text))
    try:
        pyperclip.copy(text)
        _send_paste()
        log.info("Paste sent for hotkey '%s'.", hotkey)
    except Exception as exc:
        log.error("Failed to paste snippet for hotkey '%s': %s", hotkey, exc)
        notify_error(
            f"Could not paste the snippet for hotkey '{hotkey}':\n\n{exc}\n\n"
            "The clipboard or the target window may be unavailable. "
            "Try again.")


# Native control classes that handle WM_PASTE, so we can paste directly into
# them without simulating any keystroke.
_EDIT_CLASSES = ("edit", "richedit", "richedit20a", "richedit20w",
                 "richedit50w", "richedit60w", "richedit60a")


def _focused_control(hwnd):
    """Return the control with keyboard focus inside the given window."""
    target_tid, _ = win32process.GetWindowThreadProcessId(hwnd)
    cur_tid = win32api.GetCurrentThreadId()
    attached = False
    focus = 0
    try:
        if target_tid and target_tid != cur_tid:
            win32process.AttachThreadInput(cur_tid, target_tid, True)
            attached = True
        focus = win32gui.GetFocus()
    except Exception:
        focus = 0
    finally:
        if attached:
            try:
                win32process.AttachThreadInput(cur_tid, target_tid, False)
            except Exception:
                pass
    return focus or hwnd


def _force_foreground(hwnd):
    """Reliably bring a window to the foreground (beating the focus lock)."""
    if win32gui.GetForegroundWindow() == hwnd:
        return
    cur_tid = win32api.GetCurrentThreadId()
    fg = win32gui.GetForegroundWindow()
    fg_tid = win32process.GetWindowThreadProcessId(fg)[0] if fg else 0
    target_tid = win32process.GetWindowThreadProcessId(hwnd)[0]
    attached = []
    for tid in {fg_tid, target_tid}:
        if tid and tid != cur_tid:
            try:
                win32process.AttachThreadInput(cur_tid, tid, True)
                attached.append(tid)
            except Exception:
                pass
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.BringWindowToTop(hwnd)
        win32gui.SetForegroundWindow(hwnd)
    except Exception as exc:
        log.warning("Could not restore previous window focus: %s", exc)
    finally:
        for tid in attached:
            try:
                win32process.AttachThreadInput(cur_tid, tid, False)
            except Exception:
                pass


def paste_from_taskbar(text, label):
    """Paste a snippet chosen from the taskbar menu into the previous window.

    For classic edit controls we paste directly via WM_PASTE (no keystroke).
    Everything else (browsers, Electron, Office, UWP, ...) gets a reliable
    focus restore plus Ctrl+V, which works universally.
    """
    log.info("Menu item '%s' selected -> pasting: \"%s\"", label, _preview(text))
    try:
        pyperclip.copy(text)
        hwnd = _last_active_hwnd
        if not hwnd:
            log.warning(
                "No previous window remembered; pasting into current focus.")
            _send_paste()
            return

        control = _focused_control(hwnd)
        cls = ""
        try:
            cls = win32gui.GetClassName(control)
        except Exception:
            pass

        if cls.lower() in _EDIT_CLASSES:
            try:
                win32gui.PostMessage(control, win32con.WM_PASTE, 0, 0)
                log.info("Pasted directly via WM_PASTE into '%s' (hwnd=%s).",
                         cls, control)
                return
            except Exception as exc:
                log.warning("Direct WM_PASTE failed (%s); using Ctrl+V.", exc)

        _force_foreground(hwnd)
        log.info(
            "Restored focus to previous window (hwnd=%s); pasting via Ctrl+V.",
            hwnd)
        time.sleep(0.12)
        _send_paste()
        log.info("Paste sent for menu item '%s'.", label)
    except Exception as exc:
        log.error("Failed to paste menu item '%s': %s", label, exc)
        notify_error(
            f"Could not paste '{label}':\n\n{exc}\n\n"
            "The clipboard or the target window may be unavailable. "
            "Try again.")


def make_paste_callback(text, hotkey):
    """Build a hotkey callback that pastes on its own thread (non-blocking)."""
    def _callback():
        threading.Thread(
            target=paste_text, args=(text, hotkey), daemon=True
        ).start()
    return _callback


# --- Hotkeys ----------------------------------------------------------------

def clear_hotkeys():
    """Remove every hotkey we previously registered."""
    while _hotkey_handles:
        handle = _hotkey_handles.pop()
        try:
            keyboard.remove_hotkey(handle)
        except (KeyError, ValueError):
            pass


def register_hotkeys(notify=False):
    """Clear existing hotkeys and (re)register them from the current config.

    Each shortcut is registered independently: one invalid entry is logged,
    optionally reported to the user, and skipped, while every other (valid)
    shortcut keeps working.
    """
    clear_hotkeys()
    shortcuts = load_config(notify=notify)
    registered = 0
    invalid = []
    for hotkey, text in shortcuts.items():
        try:
            handle = keyboard.add_hotkey(hotkey, make_paste_callback(text, hotkey))
            _hotkey_handles.append(handle)
            registered += 1
            log.info("Registered hotkey '%s'.", hotkey)
        except Exception as exc:
            log.error("Invalid hotkey '%s': %s", hotkey, exc)
            invalid.append((hotkey, str(exc)))
    log.info("Loaded %d shortcut(s).", registered)
    if invalid and notify:
        lines = "\n".join(f'  \u2022 "{hk}"  ({err})' for hk, err in invalid)
        notify_warning(
            f"{len(invalid)} shortcut(s) in config.json could not be "
            f"registered and were skipped:\n\n{lines}\n\n"
            f"The other {registered} shortcut(s) work normally. Fix the hotkey "
            'syntax (e.g. "ctrl+shift+1") and reload the config.')
    return registered


# --- Icon -------------------------------------------------------------------

# icon.ico is a static asset shipped with the app; the taskbar button and the
# Jump List entries point at it via ICON_PATH.


# --- Run at Windows startup -------------------------------------------------

def _startup_command():
    return f'"{LAUNCHER_EXE}" "{SCRIPT}"'


def is_autostart_enabled():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
            winreg.QueryValueEx(key, APP_NAME)
        return True
    except (FileNotFoundError, OSError):
        return False


def set_autostart(enable):
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE
        ) as key:
            if enable:
                winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, _startup_command())
            else:
                try:
                    winreg.DeleteValue(key, APP_NAME)
                except FileNotFoundError:
                    pass
        log.info("Run at startup %s.", "enabled" if enable else "disabled")
    except OSError as exc:
        log.error("Could not update autostart setting: %s", exc)
        notify_error(
            f"Could not {'enable' if enable else 'disable'} run at startup:\n\n"
            f"{exc}\n\n"
            "The Windows registry could not be updated.")


# --- Taskbar Jump List ------------------------------------------------------

def _task_link(title, extra_args):
    """Create a Jump List task that re-launches this script with given args."""
    link = pythoncom.CoCreateInstance(
        shell.CLSID_ShellLink, None,
        pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink)
    link.SetPath(LAUNCHER_EXE)
    link.SetArguments(f'"{SCRIPT}" {extra_args}')
    link.SetWorkingDirectory(BASE_DIR)
    link.SetIconLocation(ICON_PATH, 0)
    link.SetDescription(title[:250])
    store = link.QueryInterface(propsys.IID_IPropertyStore)
    store.SetValue(pscon.PKEY_Title,
                   propsys.PROPVARIANTType(title, pythoncom.VT_LPWSTR))
    store.Commit()
    return link


def _separator_link():
    link = pythoncom.CoCreateInstance(
        shell.CLSID_ShellLink, None,
        pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink)
    store = link.QueryInterface(propsys.IID_IPropertyStore)
    store.SetValue(pscon.PKEY_AppUserModel_IsDestListSeparator,
                   propsys.PROPVARIANTType(True, pythoncom.VT_BOOL))
    store.Commit()
    return link


def build_jump_list(shortcuts):
    """Build the right-click taskbar menu from the current config."""
    try:
        cdl = pythoncom.CoCreateInstance(
            shell.CLSID_DestinationList, None,
            pythoncom.CLSCTX_INPROC_SERVER, shell.IID_ICustomDestinationList)
        cdl.SetAppID(APP_ID)
        cdl.BeginList()

        col = pythoncom.CoCreateInstance(
            shell.CLSID_EnumerableObjectCollection, None,
            pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IObjectCollection)

        for hotkey, text in shortcuts.items():
            title = f"{hotkey}    {_preview(text, 28)}"
            col.AddObject(_task_link(title, f'--action paste --key "{hotkey}"'))
        if shortcuts:
            col.AddObject(_separator_link())

        col.AddObject(_task_link("Edit config", "--action edit"))
        col.AddObject(_task_link("Reload config", "--action reload"))
        startup_label = ("Disable run at startup" if is_autostart_enabled()
                         else "Enable run at startup")
        col.AddObject(_task_link(startup_label, "--action autostart"))

        cdl.AddUserTasks(col.QueryInterface(shell.IID_IObjectArray))
        cdl.CommitList()
        log.info("Taskbar menu updated with %d snippet(s).", len(shortcuts))
    except Exception as exc:
        log.error("Could not build taskbar menu: %s", exc)
        notify_warning(
            f"The taskbar right-click menu could not be built:\n\n{exc}\n\n"
            "Keyboard shortcuts still work. Try reloading the config.")


# --- Foreground window tracking ---------------------------------------------

def start_foreground_tracker():
    """Continuously remember the last foreground window that isn't ours."""
    def _poll():
        global _last_active_hwnd
        while True:
            try:
                hwnd = win32gui.GetForegroundWindow()
                if hwnd:
                    title = win32gui.GetWindowText(hwnd)
                    cls = win32gui.GetClassName(hwnd)
                    if title and title != APP_NAME and cls not in _SKIP_CLASSES:
                        _last_active_hwnd = hwnd
            except Exception:
                pass
            time.sleep(0.2)

    threading.Thread(target=_poll, daemon=True).start()


# --- Command channel (Jump List helper -> running app) ----------------------

def reload_and_rebuild():
    register_hotkeys(notify=True)
    build_jump_list(load_config())


def handle_command(line):
    """Dispatch a command received from a Jump List helper process."""
    parts = line.strip().split("\t")
    cmd = parts[0].upper()
    if cmd == "PASTE":
        key = parts[1] if len(parts) > 1 else ""
        text = load_config().get(key)
        if text is None:
            log.warning("No snippet configured for '%s'.", key)
            notify_warning(
                f"No snippet is configured for '{key}'.\n\n"
                "It may have been renamed or removed. Reload the config after "
                "editing config.json.")
            return
        threading.Thread(
            target=paste_from_taskbar, args=(text, key), daemon=True).start()
    elif cmd == "RELOAD":
        log.info("Reload config requested.")
        reload_and_rebuild()
    elif cmd == "AUTOSTART":
        set_autostart(not is_autostart_enabled())
        build_jump_list(load_config())
    elif cmd == "QUIT":
        log.info("Quit requested.")
        if _root is not None:
            _root.after(0, _quit)
    else:
        log.warning("Unknown command: %s", line.strip())


def start_command_server():
    """Bind the loopback port. Returns False if another instance owns it."""
    srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        srv.bind((HOST, PORT))
    except OSError:
        return False
    srv.listen(5)

    def _serve():
        pythoncom.CoInitialize()
        while True:
            try:
                conn, _ = srv.accept()
            except OSError:
                break
            with conn:
                data = conn.recv(4096).decode("utf-8", "replace")
            if data:
                handle_command(data)

    threading.Thread(target=_serve, daemon=True).start()
    return True


def send_command(line, timeout=2):
    """Send a single command to the running app. Returns True on success."""
    try:
        with socket.create_connection((HOST, PORT), timeout=timeout) as s:
            s.sendall((line + "\n").encode("utf-8"))
        return True
    except OSError:
        return False


# --- Helper (--action) mode -------------------------------------------------

def run_action(action, key):
    """Executed by the short-lived process a Jump List entry launches."""
    if action == "edit":
        try:
            os.startfile(CONFIG_PATH)
        except OSError as exc:
            log.error("Could not open config.json: %s", exc)
            notify_error(
                f"Could not open config.json at:\n{CONFIG_PATH}\n\n{exc}")
        return
    if action == "paste":
        if send_command(f"PASTE\t{key}"):
            return
        # Fallback: app not running -> best-effort local paste.
        text = load_config(notify=True).get(key)
        if text:
            try:
                time.sleep(0.3)
                pyperclip.copy(text)
                _send_paste()
            except Exception as exc:
                log.error("Fallback paste for '%s' failed: %s", key, exc)
                notify_error(f"Could not paste the snippet for '{key}':\n\n{exc}")
        else:
            notify_warning(f"No snippet is configured for '{key}'.")
        return
    command = {"reload": "RELOAD", "autostart": "AUTOSTART", "quit": "QUIT"}.get(action)
    if command:
        send_command(command)


# --- Resident app -----------------------------------------------------------

def _quit():
    log.info("Shutting down. Removing hotkeys.")
    clear_hotkeys()
    if _root is not None:
        _root.destroy()


def run_resident():
    global _root

    # Own the taskbar button under our AppUserModelID (needed for the menu).
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception as exc:
        log.warning("Could not set AppUserModelID: %s", exc)

    pythoncom.CoInitialize()

    if not start_command_server():
        log.info("%s is already running. Exiting.", APP_NAME)
        return

    log.info("%s starting up.", APP_NAME)
    register_hotkeys(notify=True)
    start_foreground_tracker()

    # Minimal, invisible window: it exists only to give us a taskbar button
    # (and therefore a right-click Jump List). It stays minimized.
    _root = tk.Tk()
    _root.title(APP_NAME)
    _root.geometry("1x1+0+0")
    try:
        _root.iconbitmap(ICON_PATH)
    except Exception:
        pass
    _root.protocol("WM_DELETE_WINDOW", _quit)

    def _keep_hidden(_event=None):
        if _root.state() == "normal":
            _root.iconify()

    _root.bind("<Map>", _keep_hidden)
    _root.after(0, _root.iconify)

    build_jump_list(load_config())
    log.info("Taskbar icon ready. Right-click it for the menu.")
    _root.mainloop()
    log.info("%s stopped.", APP_NAME)


# --- Global error handling ---------------------------------------------------

def _install_global_error_handlers():
    """Make sure any unforeseen error is logged and shown, never a silent crash."""
    def _hook(exc_type, exc_value, exc_tb):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_tb)
            return
        log.error("Unhandled error", exc_info=(exc_type, exc_value, exc_tb))
        notify_error(
            f"{APP_NAME} hit an unexpected error and may not work correctly:\n\n"
            f"{exc_type.__name__}: {exc_value}",
            block=True)

    sys.excepthook = _hook

    def _thread_hook(args):
        if issubclass(args.exc_type, KeyboardInterrupt):
            return
        name = args.thread.name if args.thread else "?"
        log.error("Unhandled error in thread %s", name,
                  exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
        notify_error(
            f"{APP_NAME} hit an unexpected error in a background task:\n\n"
            f"{args.exc_type.__name__}: {args.exc_value}")

    threading.excepthook = _thread_hook


def main():
    _install_global_error_handlers()
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--action", choices=[
        "paste", "reload", "autostart", "edit", "quit"])
    parser.add_argument("--key", default="")
    args, _ = parser.parse_known_args()

    if args.action:
        run_action(args.action, args.key)
    else:
        run_resident()


if __name__ == "__main__":
    main()
