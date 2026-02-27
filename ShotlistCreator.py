# Natalia Raz
# ShotlistCreator v2.1.14 for DaVinci Resolve Studio

import os
import json
import platform
import subprocess
import time
import webbrowser
import sys
import urllib.request

# Windows-only dependencies
try:
    import psutil
    import win32gui
    import win32process
    import win32con
except ImportError:
    # On macOS, these won't import, so we ignore them
    pass

def _bootstrap_resolve_scripting():
    """Make DaVinci Resolve scripting module discoverable on macOS/Windows."""
    resolve_script_api = os.environ.get("RESOLVE_SCRIPT_API")
    if resolve_script_api:
        modules_dir = os.path.join(resolve_script_api, "Modules")
        if modules_dir not in sys.path:
            sys.path.append(modules_dir)

    system = platform.system()
    if system == "Darwin":
        default_api = "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting"
        default_lib = "/Applications/DaVinci Resolve/DaVinci Resolve.app/Contents/Libraries/Fusion/fusionscript.so"
        modules_dir = os.path.join(default_api, "Modules")
        if os.path.isdir(modules_dir) and modules_dir not in sys.path:
            sys.path.append(modules_dir)
        os.environ.setdefault("RESOLVE_SCRIPT_API", default_api)
        os.environ.setdefault("RESOLVE_SCRIPT_LIB", default_lib)
    elif system == "Windows":
        program_data = os.environ.get("PROGRAMDATA", r"C:\ProgramData")
        default_api = os.path.join(program_data, "Blackmagic Design", "DaVinci Resolve", "Support", "Developer", "Scripting")
        modules_dir = os.path.join(default_api, "Modules")
        if os.path.isdir(modules_dir) and modules_dir not in sys.path:
            sys.path.append(modules_dir)
        default_lib = r"C:\Program Files\Blackmagic Design\DaVinci Resolve\fusionscript.dll"
        os.environ.setdefault("RESOLVE_SCRIPT_API", default_api)
        os.environ.setdefault("RESOLVE_SCRIPT_LIB", default_lib)


_bootstrap_resolve_scripting()


def _show_startup_error(message):
    print(message)
    system = platform.system()
    if system == "Windows":
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, message, "ShotlistCreator startup error", 0x10)
        except Exception:
            pass
    elif system == "Darwin":
        try:
            safe_message = message.replace('"', '\\"')
            subprocess.run(
                ["osascript", "-e", f'display dialog "{safe_message}" buttons {{"OK"}} default button "OK"'],
                check=False,
            )
        except Exception:
            pass


def _is_macos_accessibility_trusted():
    if platform.system() != "Darwin":
        return True
    trusted_values = []
    try:
        from ApplicationServices import AXIsProcessTrustedWithOptions, kAXTrustedCheckOptionPrompt
        trusted_values.append(bool(AXIsProcessTrustedWithOptions({kAXTrustedCheckOptionPrompt: False})))
    except Exception:
        pass
    try:
        from Quartz import AXIsProcessTrusted
        trusted_values.append(bool(AXIsProcessTrusted()))
    except Exception:
        pass
    if not trusted_values:
        return True
    return any(trusted_values)


def _request_macos_accessibility_permission(prompt=False):
    if platform.system() != "Darwin":
        return True
    if prompt:
        try:
            from ApplicationServices import AXIsProcessTrustedWithOptions, kAXTrustedCheckOptionPrompt
            AXIsProcessTrustedWithOptions({kAXTrustedCheckOptionPrompt: True})
        except Exception:
            pass
    return _is_macos_accessibility_trusted()


try:
    import DaVinciResolveScript as dvr_script
except Exception as exc:
    _show_startup_error(
        "Could not load DaVinci Resolve scripting API.\n\n"
        "Please install DaVinci Resolve Studio and launch it once, then try again."
    )
    raise RuntimeError("DaVinciResolveScript import failed.") from exc
import xlsxwriter
from PySide6 import QtWidgets, QtCore, QtGui
from PIL import Image
from pynput.keyboard import Controller


# -----------------------------------------------------------------------------
# 1) DaVinci Resolve focusing logic, cross-platform
# -----------------------------------------------------------------------------

def get_resolve_main_window_handle_windows():
    """
    Find DaVinci Resolve's main window on Windows by enumerating processes named 'Resolve.exe'.
    Returns the first top-level visible window handle, or None if not found.
    """
    resolve_pid = None
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and proc.info['name'].lower() == "resolve.exe":
            resolve_pid = proc.info['pid']
            break

    if not resolve_pid:
        return None

    def enum_windows_callback(hwnd, hwnd_list):
        if win32gui.IsWindowVisible(hwnd):
            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
            if window_pid == resolve_pid:
                hwnd_list.append(hwnd)
        return True

    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)
    return windows[0] if windows else None


def focus_on_resolve_windows():
    """
    Restore & focus the main Resolve window on Windows.
    """
    hwnd = get_resolve_main_window_handle_windows()
    if hwnd:
        # Restore if minimized
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.3)
        # Foreground
        win32gui.SetForegroundWindow(hwnd)
    else:
        print("Could not find a visible DaVinci Resolve window.")


def focus_on_timeline():
    """
    Cross-platform function to ensure DaVinci Resolve is frontmost
    before sending keyboard events.
    """
    system = platform.system()
    if system == "Windows":
        try:
            focus_on_resolve_windows()
        except Exception as e:
            print("Failed to focus DaVinci Resolve on Windows:", e)
    elif system == "Darwin":
        # macOS
        subprocess.run(["osascript", "-e", 'tell application "DaVinci Resolve" to activate'])
    else:
        # Linux or others - do nothing or adapt as needed
        pass


# -----------------------------------------------------------------------------
# 2) Standard I/O routines for saving Excel, subfolders, etc.
# -----------------------------------------------------------------------------

resolve = dvr_script.scriptapp("Resolve")
keyboard = Controller()

APP_NAME = "ShotlistCreator"
APP_VERSION = "2.1.14"
__version__ = APP_VERSION
RELEASE_FLAG = True
APP_TITLE = f"{APP_NAME} v{APP_VERSION}" if RELEASE_FLAG else f"{APP_NAME} v{APP_VERSION} (dev)"
README_URL = "https://github.com/natlrazfx/Shotlist-Creator#how-it-works"
SETUP_VIDEO_URL = "https://youtu.be/lGYmBYw0BuA"
SETUP_THUMBNAIL_URL = "https://img.youtube.com/vi/lGYmBYw0BuA/maxresdefault.jpg"
SETUP_LOCAL_IMAGE = os.path.join("assets", "next_marker_bind.png")
SUPPORT_URL = "https://aescripts.com/shotlist-creator-for-davinci-resolve/"
THUMBNAIL_FIELD = "Still/Thumbnail"
TIMELINE_PREFIX = "Timeline:"


def _get_config_path():
    if platform.system() == "Windows":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
        cfg_dir = os.path.join(base, APP_NAME)
    else:
        cfg_dir = os.path.join(os.path.expanduser("~"), ".config", APP_NAME)
    os.makedirs(cfg_dir, exist_ok=True)
    return os.path.join(cfg_dir, "settings.json")


def _load_settings():
    cfg_path = _get_config_path()
    if not os.path.exists(cfg_path):
        return {}
    try:
        with open(cfg_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {}


def _save_settings(data):
    try:
        with open(_get_config_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


def _resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def _current_app_path():
    exe_path = os.path.abspath(sys.executable)
    marker = ".app/"
    if marker in exe_path:
        return exe_path.split(marker, 1)[0] + ".app"
    return exe_path


def _load_setup_thumbnail():
    local_image_path = _resource_path(SETUP_LOCAL_IMAGE)
    if os.path.exists(local_image_path):
        pixmap = QtGui.QPixmap(local_image_path)
        if not pixmap.isNull():
            return pixmap

    try:
        with urllib.request.urlopen(SETUP_THUMBNAIL_URL, timeout=3) as resp:
            img_data = resp.read()
        pixmap = QtGui.QPixmap()
        if pixmap.loadFromData(img_data):
            return pixmap
    except Exception:
        pass
    return None


def _open_macos_accessibility_settings():
    if platform.system() != "Darwin":
        return
    subprocess.Popen(
        ["open", "x-apple.systempreferences:com.apple.preference.security?Privacy_Accessibility"]
    )
    # Bring System Settings to front so macOS permission prompts are visible.
    subprocess.run(
        ["osascript", "-e", 'tell application "System Settings" to activate'],
        check=False,
    )


def _open_macos_applications_folder():
    if platform.system() != "Darwin":
        return
    subprocess.Popen(["open", "/Applications"])


def _ensure_macos_accessibility_permission():
    if platform.system() != "Darwin":
        return True
    if _is_macos_accessibility_trusted():
        return True

    while True:
        app_path = _current_app_path()
        app_hint = ""
        if app_path.startswith("/Volumes/"):
            app_hint = (
                "\n\nCurrent app is running from a mounted DMG volume:\n"
                f"{app_path}\n"
                "Drag ShotlistCreator.app to /Applications and run it from there."
            )

        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle(APP_TITLE)
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setText("Accessibility permission is required.")
        msg.setInformativeText(
            "ShotlistCreator uses keyboard control to jump between markers.\n\n"
            "Open System Settings -> Privacy & Security -> Accessibility,\n"
            "enable ShotlistCreator, then click Recheck.\n\n"
            f"Current app path:\n{app_path}"
            f"{app_hint}"
        )
        open_btn = msg.addButton("Open Settings", QtWidgets.QMessageBox.ActionRole)
        open_apps_btn = msg.addButton("Open Applications Folder", QtWidgets.QMessageBox.ActionRole)
        recheck_btn = msg.addButton("Recheck", QtWidgets.QMessageBox.AcceptRole)
        exit_btn = msg.addButton("Exit", QtWidgets.QMessageBox.RejectRole)
        msg.setDefaultButton(recheck_btn)
        msg.exec()

        if msg.clickedButton() == open_btn:
            _request_macos_accessibility_permission(prompt=True)
            _open_macos_accessibility_settings()
        elif msg.clickedButton() == open_apps_btn:
            _open_macos_applications_folder()
        elif msg.clickedButton() == recheck_btn:
            time.sleep(0.2)
            if _is_macos_accessibility_trusted():
                return True
            QtWidgets.QMessageBox.information(
                None,
                APP_TITLE,
                "Permission is still not active.\n\n"
                f"Current app path:\n{app_path}\n\n"
                "If you just enabled it, quit and relaunch ShotlistCreator, then click Recheck.",
            )
        else:
            return False


def _show_bind_setup_dialog(force=False):
    settings = _load_settings()
    if settings.get("hide_bind_setup_dialog") and not force:
        return

    dialog = QtWidgets.QDialog()
    dialog.setWindowTitle(APP_TITLE)
    dialog.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
    dialog.resize(980, 760)

    layout = QtWidgets.QVBoxLayout(dialog)
    title = QtWidgets.QLabel("One-time setup required: bind Next Marker to keyboard key 0.")
    title.setWordWrap(True)
    title.setStyleSheet("font-size: 30px; font-weight: 700;")
    layout.addWidget(title)

    thumbnail = _load_setup_thumbnail()
    if thumbnail is not None:
        image_label = QtWidgets.QLabel()
        image_label.setPixmap(thumbnail.scaled(920, 420, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
        image_label.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(image_label)

    body = QtWidgets.QLabel(
        "DaVinci Resolve -> Keyboard Customization -> Playback -> Next Marker -> set key 0.\n\n"
        "ShotlistCreator uses this key to move through timeline markers."
    )
    body.setWordWrap(True)
    body.setStyleSheet("font-size: 24px;")
    layout.addWidget(body)

    hide_checkbox = QtWidgets.QCheckBox("Don't show this again")
    hide_checkbox.setChecked(bool(settings.get("hide_bind_setup_dialog")))
    layout.addWidget(hide_checkbox)

    buttons = QtWidgets.QHBoxLayout()
    read_btn = QtWidgets.QPushButton("Read Instructions")
    watch_btn = QtWidgets.QPushButton("Watch Tutorial")
    continue_btn = QtWidgets.QPushButton("Continue")
    continue_btn.setDefault(True)
    read_btn.clicked.connect(lambda: webbrowser.open(README_URL))
    watch_btn.clicked.connect(lambda: webbrowser.open(SETUP_VIDEO_URL))
    continue_btn.clicked.connect(dialog.accept)
    buttons.addWidget(read_btn)
    buttons.addWidget(watch_btn)
    buttons.addWidget(continue_btn)
    layout.addLayout(buttons)

    dialog.exec()

    settings["hide_bind_setup_dialog"] = bool(hide_checkbox.isChecked())
    _save_settings(settings)


def _safe_timeline_item_call(timeline_item, method_name, *args):
    method = getattr(timeline_item, method_name, None)
    if not callable(method):
        return ""
    try:
        return method(*args)
    except TypeError:
        try:
            return method()
        except Exception:
            return ""
    except Exception:
        return ""


def _collect_timeline_item_metadata(timeline_item):
    if not timeline_item:
        return {}

    meta = {}

    # Collect all generic timeline item properties exposed by Resolve.
    try:
        props = timeline_item.GetProperty()
        if isinstance(props, dict):
            for key, value in props.items():
                meta[f"{TIMELINE_PREFIX} {key}"] = value
    except Exception:
        pass

    # Add commonly requested timeline item values from dedicated API calls.
    meta["Record In"] = _safe_timeline_item_call(timeline_item, "GetStart", False)
    meta["Record Out"] = _safe_timeline_item_call(timeline_item, "GetEnd", False)
    meta["Record Duration"] = _safe_timeline_item_call(timeline_item, "GetDuration", False)
    meta["Source In"] = _safe_timeline_item_call(timeline_item, "GetSourceStartFrame")
    meta["Source Out"] = _safe_timeline_item_call(timeline_item, "GetSourceEndFrame")
    meta["Source Start Time"] = _safe_timeline_item_call(timeline_item, "GetSourceStartTime")
    meta["Source End Time"] = _safe_timeline_item_call(timeline_item, "GetSourceEndTime")

    track_info = _safe_timeline_item_call(timeline_item, "GetTrackTypeAndIndex")
    if isinstance(track_info, (list, tuple)) and len(track_info) == 2:
        meta["Track Type"] = track_info[0]
        meta["Track Index"] = track_info[1]

    return meta


def get_save_file_name(project_name):
    app = QtWidgets.QApplication.instance()
    if not app:
        app = QtWidgets.QApplication([])
    options = QtWidgets.QFileDialog.Options()
    default_filename = f"{project_name}_shotlist_v001.xlsx" if project_name else ""
    file_name, _ = QtWidgets.QFileDialog.getSaveFileName(
        None,
        "Save As",
        default_filename,
        "Excel Files (*.xlsx);;All Files (*)",
        options=options,
    )
    return file_name

def ask_replace_or_rename(file_or_folder):
    msgBox = QtWidgets.QMessageBox()
    msgBox.setIcon(QtWidgets.QMessageBox.Question)
    msgBox.setText(f"'{file_or_folder}' already exists. What would you like to do?")
    msgBox.setWindowTitle("File/Folder Exists")
    replace_button = msgBox.addButton("Replace", QtWidgets.QMessageBox.AcceptRole)
    rename_button = msgBox.addButton("Rename", QtWidgets.QMessageBox.NoRole)
    cancel_button = msgBox.addButton("Cancel", QtWidgets.QMessageBox.RejectRole)
    msgBox.setDefaultButton(replace_button)

    msgBox.exec()

    if msgBox.clickedButton() == replace_button:
        return "replace"
    elif msgBox.clickedButton() == rename_button:
        return "rename"
    else:
        return "cancel"

def ask_create_subfolder(output_path, file_name):
    subfolder_name = os.path.splitext(file_name)[0]
    subfolder_path = os.path.join(output_path, subfolder_name)

    while True:
        if os.path.exists(subfolder_path):
            action = ask_replace_or_rename(subfolder_name)
            if action == "replace":
                for root, dirs, files in os.walk(subfolder_path, topdown=False):
                    for f in files:
                        os.remove(os.path.join(root, f))
                    for d in dirs:
                        os.rmdir(os.path.join(root, d))
                break
            elif action == "rename":
                app = QtWidgets.QApplication.instance()
                if not app:
                    app = QtWidgets.QApplication([])
                new_name, ok = QtWidgets.QInputDialog.getText(
                    None,
                    "Rename",
                    "Enter new name for the folder and file:",
                    text=subfolder_name,
                )
                if ok and new_name:
                    subfolder_path = os.path.join(output_path, new_name)
                    file_name = f"{new_name}.xlsx"
                    subfolder_name = new_name
                else:
                    return None, None
            else:
                return None, None
        else:
            os.makedirs(subfolder_path)
            break

    return subfolder_path, file_name

def get_color_format(workbook, color_name):
    color_map = {
        "Rose": "#FF007F",
        "Pink": "#FFC0CB",
        "Lavender": "#E6E6FA",
        "Cyan": "#00FFFF",
        "Fuchsia": "#FF00FF",
        "Mint": "#98FF98",
        "Sand": "#C2B280",
        "Yellow": "#FFFF00",
        "Green": "#00FF00",
        "Blue": "#0000FF",
        "Purple": "#800080",
        "Red": "#FF0000",
        "Cocoa": "#D2691E",
        "Sky": "#87CEEB",
        "Lemon": "#FFF44F",
        "Cream": "#FFFDD0",
    }
    hex_color = color_map.get(color_name, "#FFFFFF")
    return workbook.add_format({"bg_color": hex_color, "valign": "vcenter", "align": "center"})

def open_folder_in_explorer(output_path):
    system = platform.system()
    if system == "Windows":
        subprocess.Popen(["explorer", os.path.normpath(output_path)])
    elif system == "Darwin":
        subprocess.Popen(["open", output_path])
    else:
        print("Unsupported OS for auto-opening folder.")


# -----------------------------------------------------------------------------
# 3) Collecting all metadata keys from the entire timeline
# -----------------------------------------------------------------------------

def gather_all_metadata_keys_from_timeline(timeline):
    """
    Loop over all video tracks in the current timeline,
    gather the union of all clip properties from each MediaPoolItem,
    and return them as a list of keys (with standard fields at front).
    """
    standard_fields = [
        THUMBNAIL_FIELD,
        "Frame",
        "Timecode",
        "Name",
        "Note",
        "Duration",
        "Color",
        "Record In",
        "Record Out",
        "Source In",
        "Source Out",
        "Record Duration",
        "Source Start Time",
        "Source End Time",
        "Track Type",
        "Track Index",
    ]

    # We'll store all discovered keys in a set
    discovered_keys = set()

    # For each video track, get all timeline items
    track_count = timeline.GetTrackCount("video")
    for track_idx in range(1, track_count + 1):
        timeline_items = timeline.GetItemListInTrack("video", track_idx)
        for ti in timeline_items:
            discovered_keys.update(_collect_timeline_item_metadata(ti).keys())
            mp_item = ti.GetMediaPoolItem()
            if mp_item:
                props = mp_item.GetClipProperty()
                discovered_keys.update(props.keys())

    # Exclude standard fields from discovered so we don't duplicate
    discovered_keys.difference_update(standard_fields)

    # Sort them (alphabetically, for instance)
    discovered_sorted = sorted(discovered_keys)

    # The final list has standard fields at front
    all_fields = standard_fields + discovered_sorted
    return all_fields


# -----------------------------------------------------------------------------
# 4) Export Markers to Excel
# -----------------------------------------------------------------------------

def export_markers(timeline, output_path, timecodes, excel_filename, metadata_list, selected_fields, image_size):
    markers = timeline.GetMarkers()
    workbook = xlsxwriter.Workbook(os.path.join(output_path, excel_filename))
    worksheet = workbook.add_worksheet()

    text_format = workbook.add_format({"valign": "vcenter", "align": "left"})
    max_size = image_size

    headers = selected_fields
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header, text_format)

    row = 1

    currentProject = resolve.GetProjectManager().GetCurrentProject()
    gallery = currentProject.GetGallery()
    currentStillAlbum = gallery.GetCurrentStillAlbum()
    stills = currentStillAlbum.GetStills()

    # Write each marker's data
    for idx, (frame, marker) in enumerate(markers.items()):
        col = 0
        data_row = {
            "Frame": frame,
            "Name": marker["name"],
            "Note": marker["note"],
            "Duration": marker["duration"],
            "Color": marker["color"],
            "Timecode": timecodes[idx]
        }
        data_row.update(metadata_list[idx])

        for field in selected_fields:
            if field == THUMBNAIL_FIELD:
                worksheet.write(row, col, "", text_format)
            elif field == "Color":
                worksheet.write(row, col, "", get_color_format(workbook, data_row.get(field, "")))
            else:
                worksheet.write(row, col, data_row.get(field, ""), text_format)
            col += 1
        row += 1

    image_col_index = headers.index(THUMBNAIL_FIELD) if THUMBNAIL_FIELD in headers else None
    row = 1

    # Export stills
    for i, still in enumerate(stills):
        suffix = f"{i + 1:03}"
        tmp_name = f"tmp_{suffix}"
        currentStillAlbum.ExportStills([still], output_path, tmp_name, "png")

    files = sorted(os.listdir(output_path))
    for file in files:
        if file.startswith("tmp") and file.endswith(".png"):
            file_number = int(file.split("_")[1])
            new_name = f"thumb{file_number:03d}.png"

            while os.path.exists(os.path.join(output_path, new_name)):
                action = ask_replace_or_rename(new_name)
                if action == "replace":
                    os.remove(os.path.join(output_path, new_name))
                elif action == "rename":
                    new_name, ok = QtWidgets.QInputDialog.getText(
                        None, "Rename", "Enter new name for the image:", text=new_name
                    )
                    if not ok or not new_name:
                        return
                else:
                    return

            os.rename(os.path.join(output_path, file), os.path.join(output_path, new_name))

            image_file_path = os.path.normpath(os.path.join(output_path, new_name))
            print("Exported image:", image_file_path)

            image = Image.open(image_file_path)
            width, height = image.size
            if width > height:
                new_width = int(max_size)
                new_height = int((max_size / width) * height)
            else:
                new_height = int(max_size)
                new_width = int((max_size / height) * width)

            resized_image = image.resize((new_width, new_height))
            resized_image.save(image_file_path)

            if image_col_index is not None:
                worksheet.insert_image(row, image_col_index, image_file_path, {"x_scale": 1, "y_scale": 1, "object_position": 1})
                worksheet.set_column(image_col_index, image_col_index, new_width / 6)
                worksheet.set_row(row, new_height / 1.33)

            row += 1

    worksheet.autofit()
    workbook.close()


# -----------------------------------------------------------------------------
# 5) Dark theme
# -----------------------------------------------------------------------------

def set_dark_theme(app):
    app.setStyle('Fusion')
    dark_palette = QtGui.QPalette()

    dark_color = QtGui.QColor(45, 45, 45)
    disabled_color = QtGui.QColor(127, 127, 127)

    dark_palette.setColor(QtGui.QPalette.Window, dark_color)
    dark_palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.Base, QtGui.QColor(18, 18, 18))
    dark_palette.setColor(QtGui.QPalette.AlternateBase, dark_color)
    dark_palette.setColor(QtGui.QPalette.ToolTipBase, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.ToolTipText, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.Text, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Text, disabled_color)
    dark_palette.setColor(QtGui.QPalette.Button, dark_color)
    dark_palette.setColor(QtGui.QPalette.ButtonText, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, disabled_color)
    dark_palette.setColor(QtGui.QPalette.BrightText, QtCore.Qt.red)
    dark_palette.setColor(QtGui.QPalette.Link, QtGui.QColor(42, 130, 218))
    dark_palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(42, 130, 218))
    dark_palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Highlight, QtGui.QColor(80, 80, 80))
    dark_palette.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.white)
    dark_palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.HighlightedText, disabled_color)

    app.setPalette(dark_palette)


# -----------------------------------------------------------------------------
# 6) The main reordering/presets UI
# -----------------------------------------------------------------------------

class UserInputDialog(QtWidgets.QDialog):
    def __init__(self, all_fields, parent=None):
        super(UserInputDialog, self).__init__(parent)

        # Keep window on top
        self.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        # Large default size
        self.resize(1200, 800)
        self.setWindowTitle(f"{APP_TITLE} Options")

        self.search_results = []
        self.search_index = 0
        self.field_role = QtCore.Qt.UserRole

        layout = QtWidgets.QVBoxLayout(self)

        instructions = QtWidgets.QLabel("Please set the options below:")
        layout.addWidget(instructions)

        # Timecode
        timecode_label = QtWidgets.QLabel("Enter custom timecode (default is 01:00:00:00):")
        layout.addWidget(timecode_label)
        self.timecode_input = QtWidgets.QLineEdit("01:00:00:00")
        layout.addWidget(self.timecode_input)

        # Delete stills
        self.delete_stills_checkbox = QtWidgets.QCheckBox("Delete all stills from the gallery album")
        layout.addWidget(self.delete_stills_checkbox)

        # Search
        search_label = QtWidgets.QLabel("Search in metadata fields:")
        layout.addWidget(search_label)

        search_layout = QtWidgets.QHBoxLayout()
        self.search_field = QtWidgets.QLineEdit()
        self.search_field.setPlaceholderText("Type something (e.g. 'audio', 'tape', etc.)")
        self.search_field.textChanged.connect(self.search_in_list)
        self.find_next_button = QtWidgets.QPushButton("Find Next")
        self.find_next_button.clicked.connect(self.find_next_match)

        search_layout.addWidget(self.search_field)
        search_layout.addWidget(self.find_next_button)
        layout.addLayout(search_layout)

        # Preset load/save
        preset_buttons_layout = QtWidgets.QHBoxLayout()
        load_preset_button = QtWidgets.QPushButton("Load Preset")
        save_preset_button = QtWidgets.QPushButton("Save Preset")
        setup_guide_button = QtWidgets.QPushButton("Show Setup Guide")


        load_preset_button.clicked.connect(self.on_load_preset_clicked)
        save_preset_button.clicked.connect(self.on_save_preset_clicked)
        setup_guide_button.clicked.connect(self.on_show_setup_guide_clicked)

        # Support button
        donate_button = QtWidgets.QPushButton("Support")
        donate_button.setStyleSheet("""
                    QPushButton {
                        background-color: #8A2BE2; /* Purple */
                        color: white;
                        font-weight: bold;
                    }
                    QPushButton:hover {
                        background-color: #9E47FF; /* Slightly lighter on hover */
                    }
                """)
        donate_button.clicked.connect(lambda: webbrowser.open(SUPPORT_URL))

        preset_buttons_layout.addWidget(load_preset_button)
        preset_buttons_layout.addWidget(save_preset_button)
        preset_buttons_layout.addWidget(setup_guide_button)
        preset_buttons_layout.addWidget(donate_button)
        layout.addLayout(preset_buttons_layout)

        info_layout = QtWidgets.QHBoxLayout()

        # We can use HTML for clickable links
        # setOpenExternalLinks(True) allows user to click the links.
        self.info_label = QtWidgets.QLabel(
            '<span style="font-size:10px;">'
            '<a href="https://www.linkedin.com/in/natalia-raz-0b8329120/">Natalia Raz</a> &nbsp;|&nbsp; '
            '<a href="https://github.com/natlrazfx">GitHub</a> &nbsp;|&nbsp; '
            '<a href="https://vimeo.com/552106671">Vimeo</a>'
            '</span>'
        )
        self.info_label.setOpenExternalLinks(True)

        info_layout.addStretch(1)  # pushes label to the right if you like
        info_layout.addWidget(self.info_label)
        # info_layout.addStretch(1)  # or comment out if you don't want right alignment

        layout.addLayout(info_layout)

        # Metadata label
        metadata_label = QtWidgets.QLabel("Select and reorder the metadata fields:")
        layout.addWidget(metadata_label)

        # QListWidget (reorderable + checkable)
        self.list_widget = QtWidgets.QListWidget()
        self.list_widget.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.list_widget.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        self.list_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        layout.addWidget(self.list_widget, stretch=1)

        # Default fields that are checked
        self.default_selected_fields = [
            THUMBNAIL_FIELD, "Frame", "Timecode", "Name", "Note", "Duration", "Color",
            "Record In", "Record Out",
            "Source In", "Source Out", "Record Duration",
            "Track Type", "Track Index",
            "Clip Name", "FPS", "File Path", "Video Codec",
            "Resolution", "Start TC", "End TC"
        ]
        self.all_fields = list(all_fields)
        self._rebuild_field_list(self.all_fields, set(self.default_selected_fields))

        # Select All / Deselect All
        button_layout = QtWidgets.QHBoxLayout()
        select_all_button = QtWidgets.QPushButton("Select All")
        deselect_all_button = QtWidgets.QPushButton("Deselect All")
        select_all_button.clicked.connect(self.select_all_items)
        deselect_all_button.clicked.connect(self.deselect_all_items)
        button_layout.addWidget(select_all_button)
        button_layout.addWidget(deselect_all_button)
        layout.addLayout(button_layout)

        # Image size
        size_label = QtWidgets.QLabel("Choose the size for the still images:")
        layout.addWidget(size_label)

        size_layout = QtWidgets.QHBoxLayout()
        self.size_combo = QtWidgets.QComboBox()
        self.size_combo.addItems(["SMALL", "LARGE", "CUSTOM"])
        size_layout.addWidget(self.size_combo)

        self.custom_size_input = QtWidgets.QDoubleSpinBox()
        self.custom_size_input.setRange(0.1, 10)
        self.custom_size_input.setSingleStep(0.1)
        self.custom_size_input.setValue(1)
        self.custom_size_input.setVisible(False)
        size_layout.addWidget(self.custom_size_input)
        layout.addLayout(size_layout)

        def on_size_change():
            self.custom_size_input.setVisible(self.size_combo.currentText() == "CUSTOM")

        self.size_combo.currentIndexChanged.connect(on_size_change)

        # OK / Cancel
        ok_cancel_layout = QtWidgets.QHBoxLayout()
        ok_button = QtWidgets.QPushButton("OK")
        cancel_button = QtWidgets.QPushButton("Cancel")
        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        ok_cancel_layout.addWidget(ok_button)
        ok_cancel_layout.addWidget(cancel_button)
        layout.addLayout(ok_cancel_layout)

    def _is_timeline_field(self, field_name):
        return field_name.startswith(f"{TIMELINE_PREFIX} ") or field_name in {
            "Record In",
            "Record Out",
            "Record Duration",
            "Source In",
            "Source Out",
            "Source Start Time",
            "Source End Time",
            "Track Type",
            "Track Index",
        }

    def _add_separator_item(self, label):
        item = QtWidgets.QListWidgetItem(f"────────  {label}  ────────")
        item.setFlags(QtCore.Qt.NoItemFlags)
        item.setForeground(QtGui.QColor(140, 140, 140))
        item.setData(self.field_role, "separator")
        self.list_widget.addItem(item)

    def _add_field_item(self, field_name, checked_fields):
        item = QtWidgets.QListWidgetItem(field_name)
        item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
        item.setData(self.field_role, "field")
        if field_name in checked_fields:
            item.setCheckState(QtCore.Qt.Checked)
        else:
            item.setCheckState(QtCore.Qt.Unchecked)
        self.list_widget.addItem(item)

    def _rebuild_field_list(self, field_order, checked_fields):
        self.list_widget.clear()

        standard_fields = [f for f in field_order if f in self.default_selected_fields and not self._is_timeline_field(f)]
        timeline_fields = [f for f in field_order if self._is_timeline_field(f)]
        clip_fields = [f for f in field_order if f not in standard_fields and f not in timeline_fields]

        if standard_fields:
            self._add_separator_item("Standard Fields")
            for field in standard_fields:
                self._add_field_item(field, checked_fields)
        if timeline_fields:
            self._add_separator_item("Timeline Fields")
            for field in timeline_fields:
                self._add_field_item(field, checked_fields)
        if clip_fields:
            self._add_separator_item("Clip Metadata")
            for field in clip_fields:
                self._add_field_item(field, checked_fields)

    # ----------------------------------------------------------------
    # Preset: one preset per JSON
    # ----------------------------------------------------------------
    def on_save_preset_clicked(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save Preset",
            "",
            "JSON Files (*.json);;All Files (*)"
        )
        if not path:
            return
        fields_in_order = []
        checked_list = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.data(self.field_role) != "field":
                continue
            fields_in_order.append(item.text())
            if item.checkState() == QtCore.Qt.Checked:
                checked_list.append(item.text())

        data = {"order": fields_in_order, "checked": checked_list}
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", f"Failed to save preset:\n{e}")

    def on_load_preset_clicked(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Load Preset",
            "",
            "JSON Files (*.json);;All Files (*)"
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", f"Failed to read preset:\n{e}")
            return

        fields_order = [f for f in data.get("order", []) if f in self.all_fields]
        for f in self.all_fields:
            if f not in fields_order:
                fields_order.append(f)
        checked_fields = {f for f in data.get("checked", []) if f in self.all_fields}
        self._rebuild_field_list(fields_order, checked_fields)

    def on_show_setup_guide_clicked(self):
        _show_bind_setup_dialog(force=True)

    # ----------------------------------------------------------------
    # Searching
    # ----------------------------------------------------------------
    def search_in_list(self):
        query = self.search_field.text().strip().lower()
        self.search_results.clear()
        self.search_index = 0

        if not query:
            self.list_widget.clearSelection()
            return

        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.data(self.field_role) != "field":
                continue
            txt = item.text().lower()
            if query in txt:
                self.search_results.append(i)

        if self.search_results:
            idx = self.search_results[0]
            item = self.list_widget.item(idx)
            self.list_widget.setCurrentItem(item)
            self.list_widget.scrollToItem(item)
        else:
            QtWidgets.QMessageBox.information(self, "Not Found", f"No fields match '{self.search_field.text()}'")

    def find_next_match(self):
        if not self.search_results:
            return
        self.search_index += 1
        if self.search_index >= len(self.search_results):
            self.search_index = 0
        idx = self.search_results[self.search_index]
        item = self.list_widget.item(idx)
        self.list_widget.setCurrentItem(item)
        self.list_widget.scrollToItem(item)

    # ----------------------------------------------------------------
    # Select All / Deselect All
    # ----------------------------------------------------------------
    def select_all_items(self):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.data(self.field_role) == "field":
                item.setCheckState(QtCore.Qt.Checked)

    def deselect_all_items(self):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.data(self.field_role) == "field":
                item.setCheckState(QtCore.Qt.Unchecked)

    # ----------------------------------------------------------------
    # Return final selections
    # ----------------------------------------------------------------
    def get_values(self):
        selected_fields = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.data(self.field_role) != "field":
                continue
            if item.checkState() == QtCore.Qt.Checked:
                selected_fields.append(item.text())

        size_text = self.size_combo.currentText()
        if size_text == "CUSTOM":
            multiplier = self.custom_size_input.value()
            image_size = 260 * multiplier
        elif size_text == "LARGE":
            image_size = 520
        else:
            image_size = 260

        timecode = self.timecode_input.text() or "01:00:00:00"
        delete_stills = self.delete_stills_checkbox.isChecked()
        return selected_fields, image_size, timecode, delete_stills


# -----------------------------------------------------------------------------
# 7) Main script logic
# -----------------------------------------------------------------------------

if __name__ == "__main__":
    # Create Qt app
    app = QtWidgets.QApplication.instance()
    if not app:
        app = QtWidgets.QApplication([])
    set_dark_theme(app)
    app_icon_path = _resource_path("icon.png")
    if os.path.exists(app_icon_path):
        app.setWindowIcon(QtGui.QIcon(app_icon_path))
    if not _ensure_macos_accessibility_permission():
        sys.exit(1)

    if resolve is None:
        QtWidgets.QMessageBox.critical(
            None,
            APP_TITLE,
            "DaVinci Resolve Studio is not available.\n\nOpen Resolve Studio and run ShotlistCreator again.",
        )
        sys.exit(1)

    projectManager = resolve.GetProjectManager()
    if projectManager is None:
        QtWidgets.QMessageBox.critical(
            None,
            APP_TITLE,
            "Could not access Resolve project manager.\n\nPlease restart DaVinci Resolve Studio and try again.",
        )
        sys.exit(1)

    _show_bind_setup_dialog()

    currentProject = projectManager.GetCurrentProject()
    if currentProject is None:
        QtWidgets.QMessageBox.warning(
            None,
            APP_TITLE,
            "No project is currently open.\n\nOpen a project in DaVinci Resolve Studio and run again.",
        )
        sys.exit(1)

    currentTimeline = currentProject.GetCurrentTimeline()
    if currentTimeline is None:
        QtWidgets.QMessageBox.warning(
            None,
            APP_TITLE,
            "No timeline is currently active.\n\nOpen a timeline and run ShotlistCreator again.",
        )
        sys.exit(1)

    while True:
        currentProject = projectManager.GetCurrentProject()
        if currentProject is None:
            QtWidgets.QMessageBox.warning(
                None,
                APP_TITLE,
                "No project is currently open.\n\nOpen a project in DaVinci Resolve Studio and run again.",
            )
            sys.exit(1)

        currentTimeline = currentProject.GetCurrentTimeline()
        if currentTimeline is None:
            QtWidgets.QMessageBox.warning(
                None,
                APP_TITLE,
                "No timeline is currently active.\n\nOpen a timeline and run ShotlistCreator again.",
            )
            sys.exit(1)

        project_name = currentProject.GetName()
        all_fields = gather_all_metadata_keys_from_timeline(currentTimeline)

        dialog = UserInputDialog(all_fields)
        if dialog.exec() != QtWidgets.QDialog.Accepted:
            print("Operation cancelled.")
            break

        selected_fields, image_size, timecode_to_set, delete_stills = dialog.get_values()
        if not selected_fields:
            QtWidgets.QMessageBox.information(
                None,
                APP_TITLE,
                "No metadata fields selected. Please choose at least one field.",
            )
            continue

        # Re-read active project/timeline after dialog, so we always use user's current selection.
        currentProject = projectManager.GetCurrentProject()
        if currentProject is None:
            QtWidgets.QMessageBox.warning(
                None,
                APP_TITLE,
                "No project is currently open.\n\nOpen a project in DaVinci Resolve Studio and run again.",
            )
            continue
        currentTimeline = currentProject.GetCurrentTimeline()
        if currentTimeline is None:
            QtWidgets.QMessageBox.warning(
                None,
                APP_TITLE,
                "No timeline is currently active.\n\nOpen a timeline and run ShotlistCreator again.",
            )
            continue

        markers = currentTimeline.GetMarkers()
        if not markers:
            QtWidgets.QMessageBox.information(
                None,
                APP_TITLE,
                "No markers found on the current timeline.\n\n"
                "Please add markers or switch to a timeline with markers,\n"
                "then press OK to return to options.",
            )
            continue

        # Set timecode
        currentTimeline.SetCurrentTimecode(timecode_to_set)

        # Delete stills if requested
        if delete_stills:
            resolve.OpenPage("color")
            gallery = currentProject.GetGallery()
            currentStillAlbum = gallery.GetCurrentStillAlbum()
            stills = currentStillAlbum.GetStills()
            if stills:
                success = currentStillAlbum.DeleteStills(stills)
                if success:
                    print("All stills have been successfully deleted.")
                else:
                    print("Failed to delete stills.")
            else:
                print("No stills found in the album.")

        # Prepare to collect marker-based data
        timecodes = []
        metadata_list = []

        # Focus the timeline cross-platform
        focus_on_timeline()
        keyboard = Controller()

        # For each marker
        for i, (frame_id, marker) in enumerate(markers.items()):
            numMarkersToEnd = len(markers) - (i + 1)
            print("Number of markers until the end of the timeline:", numMarkersToEnd)

            # Press "0" to jump to next marker
            keyboard.press("0")
            keyboard.release("0")
            time.sleep(0.2)

            # Get the new timecode
            currentTimecode = currentTimeline.GetCurrentTimecode()
            timecodes.append(currentTimecode)

            # Grab still
            currentTimeline.GrabStill()

            # Also gather clip metadata at this marker
            clip_metadata = {}
            current_clip = currentTimeline.GetCurrentVideoItem()
            if current_clip:
                clip_metadata["Clip Name"] = current_clip.GetName()
                clip_metadata.update(_collect_timeline_item_metadata(current_clip))
                mp_item = current_clip.GetMediaPoolItem()
                if mp_item:
                    props = mp_item.GetClipProperty()
                    for k, v in props.items():
                        clip_metadata[k] = v
            else:
                clip_metadata["Clip Name"] = "N/A"
                clip_metadata["Record In"] = ""
                clip_metadata["Record Out"] = ""
                clip_metadata["Source In"] = ""
                clip_metadata["Source Out"] = ""
                clip_metadata["Record Duration"] = ""
                clip_metadata["Source Start Time"] = ""
                clip_metadata["Source End Time"] = ""
                clip_metadata["Track Type"] = ""
                clip_metadata["Track Index"] = ""

            metadata_list.append(clip_metadata)

            if numMarkersToEnd == 0:
                break

        # Ask user for output path
        full_path = get_save_file_name(project_name)
        if not full_path:
            print("No output folder and filename selected.")
            continue

        output_path, excel_filename = os.path.split(full_path)
        if not excel_filename.endswith(".xlsx"):
            excel_filename += ".xlsx"

        # Create subfolder if needed
        output_path, excel_filename = ask_create_subfolder(output_path, excel_filename)
        if output_path:
            export_markers(
                currentTimeline, output_path, timecodes, excel_filename,
                metadata_list, selected_fields, image_size
            )
            print("DONE")
            open_folder_in_explorer(output_path)
            break
