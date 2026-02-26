# Natalia Raz
# ShotlistCreator for DaVinci Resolve Studio

import os
import json
import platform
import subprocess
import time
import webbrowser

# Windows-only dependencies
try:
    import psutil
    import win32gui
    import win32process
    import win32con
except ImportError:
    # On macOS, these won't import, so we ignore them
    pass

import DaVinciResolveScript as dvr_script
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
APP_VERSION = "2.1.2"
__version__ = APP_VERSION
RELEASE_FLAG = True
APP_TITLE = APP_NAME if RELEASE_FLAG else f"{APP_NAME} v{APP_VERSION} (dev)"


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
    standard_fields = ["Frame", "Timecode", "Name", "Note", "Duration", "Color"]

    # We'll store all discovered keys in a set
    discovered_keys = set()

    # For each video track, get all timeline items
    track_count = timeline.GetTrackCount("video")
    for track_idx in range(1, track_count + 1):
        timeline_items = timeline.GetItemListInTrack("video", track_idx)
        for ti in timeline_items:
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

    headers = selected_fields + ["Still"]
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
            if field == "Color":
                worksheet.write(row, col, "", get_color_format(workbook, data_row.get(field, "")))
            else:
                worksheet.write(row, col, data_row.get(field, ""), text_format)
            col += 1
        row += 1

    image_col_index = len(headers) - 1
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


        load_preset_button.clicked.connect(self.on_load_preset_clicked)
        save_preset_button.clicked.connect(self.on_save_preset_clicked)

        # Donate button
        donate_button = QtWidgets.QPushButton("Donate")
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
        donate_button.clicked.connect(lambda: webbrowser.open("https://www.airtm.me/natalia4xk3sygi"))

        preset_buttons_layout.addWidget(load_preset_button)
        preset_buttons_layout.addWidget(save_preset_button)
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
        default_selected_fields = [
            "Frame", "Timecode", "Name", "Note", "Duration", "Color",
            "Clip Name", "FPS", "File Path", "Video Codec",
            "Resolution", "Start TC", "End TC"
        ]

        # Populate the list
        for key in all_fields:
            item = QtWidgets.QListWidgetItem(key)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
            if key in default_selected_fields:
                item.setCheckState(QtCore.Qt.Checked)
            else:
                item.setCheckState(QtCore.Qt.Unchecked)
            self.list_widget.addItem(item)

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

        fields_order = data.get("order", [])
        checked_fields = set(data.get("checked", []))

        self.list_widget.clear()
        for field in fields_order:
            item = QtWidgets.QListWidgetItem(field)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
            if field in checked_fields:
                item.setCheckState(QtCore.Qt.Checked)
            else:
                item.setCheckState(QtCore.Qt.Unchecked)
            self.list_widget.addItem(item)

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
            txt = self.list_widget.item(i).text().lower()
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
            self.list_widget.item(i).setCheckState(QtCore.Qt.Checked)

    def deselect_all_items(self):
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(QtCore.Qt.Unchecked)

    # ----------------------------------------------------------------
    # Return final selections
    # ----------------------------------------------------------------
    def get_values(self):
        selected_fields = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
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

    projectManager = resolve.GetProjectManager()
    currentProject = projectManager.GetCurrentProject()
    currentTimeline = currentProject.GetCurrentTimeline()
    project_name = currentProject.GetName()

    # Gather all possible metadata keys from the entire timeline
    all_fields = gather_all_metadata_keys_from_timeline(currentTimeline)

    # Show the dialog
    dialog = UserInputDialog(all_fields)
    if dialog.exec() == QtWidgets.QDialog.Accepted:
        selected_fields, image_size, timecode_to_set, delete_stills = dialog.get_values()

        if not selected_fields:
            print("No metadata fields selected.")
        else:
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

            markers = currentTimeline.GetMarkers()
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
                    mp_item = current_clip.GetMediaPoolItem()
                    if mp_item:
                        props = mp_item.GetClipProperty()
                        for k, v in props.items():
                            clip_metadata[k] = v
                else:
                    clip_metadata["Clip Name"] = "N/A"

                metadata_list.append(clip_metadata)

                if numMarkersToEnd == 0:
                    break

            # Ask user for output path
            full_path = get_save_file_name(project_name)
            if not full_path:
                print("No output folder and filename selected.")
            else:
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
    else:
        print("Operation cancelled.")
