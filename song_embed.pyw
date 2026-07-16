# -*- coding: utf-8 -*-
"""
Song Embed Tool — Automate embedding song PowerPoint slides into a master
church service PowerPoint using sections as placement markers.
"""

import subprocess
import sys
import importlib.util
import os
import platform
import re
import ctypes
import urllib.request
import urllib.error
import json
import webbrowser

# --- Configuration ---
APP_NAME = "SongEmbed"
ORG_NAME = "ChurchMedia"
APP_VERSION = "2.6.0"
WINDOW_WIDTH = 1000
WINDOW_HEIGHT = 780
DEFAULT_FOLDER_TEXT = "No folder selected — click 📂"
ALLOWED_EXTENSIONS = ('.ppt', '.pptx')
# Matches any section starting with "Song" (case-insensitive)
SONG_SECTION_RE = re.compile(r'^Song.*', re.IGNORECASE)


# --- Dependency Check and Installation ---
def check_and_install_dependencies():
    dependencies = {'PyQt6': 'PyQt6'}
    if platform.system() == "Windows":
        dependencies['pywin32'] = 'win32com'

    missing = []
    for package, module in dependencies.items():
        try:
            importlib.import_module(module)
        except ImportError:
            missing.append(package)

    if not missing:
        return

    message = (f"The following required libraries are missing: {', '.join(missing)}.\n\n"
               f"Would you like to install them now?")
    title = "Missing Dependencies"

    install_allowed = False
    if platform.system() == "Windows":
        res = ctypes.windll.user32.MessageBoxW(0, message, title, 4 | 0x30)
        if res == 6:  # IDYES
            install_allowed = True
    else:
        if sys.stdin and sys.stdin.isatty():
            reply = input(f"{message} (y/n): ").lower().strip()
            if reply == 'y':
                install_allowed = True

    if not install_allowed:
        sys.exit(1)

    python_exe = sys.executable or "python"
    for package in missing:
        print(f"Installing {package}...")
        try:
            subprocess.check_call([python_exe, '-m', 'pip', 'install', package])
        except Exception as e:
            if platform.system() == "Windows":
                ctypes.windll.user32.MessageBoxW(
                    0, f"Failed to install {package}:\n{e}", "Installation Error", 0x10)
            else:
                print(f"Failed to install {package}: {e}")
            sys.exit(1)

    if platform.system() == "Windows":
        ctypes.windll.user32.MessageBoxW(
            0, "Dependencies installed successfully! The application will now start.",
            "Success", 0x40)


check_and_install_dependencies()


# --- Console Hiding ---
def hide_console():
    if platform.system() == "Windows":
        try:
            console_window = ctypes.windll.kernel32.GetConsoleWindow()
            if console_window:
                ctypes.windll.user32.ShowWindow(console_window, 0)
        except Exception:
            pass


hide_console()


# --- Main Application Imports ---
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit,
    QListWidget, QLabel, QFileDialog, QMessageBox, QListWidgetItem,
    QSpacerItem, QSizePolicy, QStyledItemDelegate, QStyle, QCheckBox,
    QComboBox, QGroupBox, QScrollArea, QInputDialog, QSplitter,
    QTreeWidget, QTreeWidgetItem, QDialog, QTextBrowser, QProgressDialog
)
from PyQt6.QtCore import Qt, QSize, QSettings, QRect, pyqtSignal, QEvent, QThread, QTimer
from PyQt6.QtGui import QColor, QIcon, QPainter

# Conditionally import win32com only on Windows
HAS_WIN32COM = False
if platform.system() == "Windows":
    try:
        import win32com.client
        import pythoncom
        HAS_WIN32COM = True
        print("pywin32 found. PowerPoint automation enabled.")
    except ImportError:
        print("WARNING: pywin32 not found. PowerPoint automation disabled.")
    except Exception as e:
        print(f"WARNING: Error importing pywin32 ({e}). PowerPoint automation disabled.")


# ─────────────────────────────────────────────────────────────────────────────
# Dark Theme Stylesheet
# ─────────────────────────────────────────────────────────────────────────────
DARK_STYLESHEET = """
/* ── Base ── */
QWidget {
    background-color: #121214;
    color: #e4e4e7;
    font-family: 'Segoe UI', 'SF Pro Display', -apple-system, sans-serif;
    font-size: 10pt;
}

/* ── Labels ── */
QLabel {
    color: #a1a1aa;
    background-color: transparent;
}
QLabel#sectionLabel {
    color: #71717a;
    font-weight: 600;
    font-size: 8pt;
}
QLabel#folderLabel {
    color: #a1a1aa;
    font-style: italic;
    font-size: 9pt;
}
QLabel#statusLabel {
    color: #71717a;
    font-size: 9pt;
    padding: 2px 0;
}
QLabel#statusError {
    color: #f87171;
    font-size: 9pt;
    padding: 2px 0;
}
QLabel#statusSuccess {
    color: #34d399;
    font-size: 9pt;
    font-weight: 500;
    padding: 2px 0;
}
QLabel#versionLabel {
    color: #52525b;
    font-size: 9pt;
    padding: 2px 0;
}
QLabel#slotAssigned {
    color: #818cf8;
    font-weight: 500;
}
QLabel#slotEmpty {
    color: #94a3b8;
    font-style: italic;
    font-weight: 600;
}
QLabel#slotFilled {
    color: #a1a1aa;
}
QLabel#headerLabel {
    color: #a1a1aa;
    font-weight: 600;
    font-size: 8pt;
    text-transform: uppercase;
}
QLabel#slotName {
    color: #f4f4f5;
    font-weight: 600;
}
QLabel#slotVerses {
    color: #a1a1aa;
    font-size: 9pt;
}

/* ── Line Edit ── */
QLineEdit {
    background-color: #1a1a1e;
    color: #f4f4f5;
    border: 1px solid #2d2d30;
    border-radius: 6px;
    padding: 8px 12px;
    selection-background-color: #666aba;
}
QLineEdit:focus {
    border: 1px solid #666aba;
}

/* ── List Widget ── */
QListWidget {
    background-color: #1a1a1e;
    color: #e4e4e7;
    border: 1px solid #2d2d30;
    border-radius: 6px;
    padding: 6px;
    outline: 0;
}
QListWidget::item {
    border-radius: 4px;
    padding: 4px 8px;
    margin: 1px 0;
}
QListWidget::item:hover {
    background-color: #27272a;
    color: #ffffff;
}
QListWidget::item:selected {
    background-color: #53569c;
    color: #ffffff;
}

/* ── Buttons ── */
QPushButton {
    background-color: #27272a;
    color: #f4f4f5;
    border: 1px solid #3f3f46;
    border-radius: 6px;
    padding: 8px 16px;
    font-weight: 500;
    min-width: 36px;
}
QPushButton:hover {
    background-color: #3f3f46;
    border-color: #52525b;
}
QPushButton:pressed {
    background-color: #18181b;
}
QPushButton:disabled {
    background-color: #18181b;
    color: #71717a;
    border-color: #27272a;
}
QPushButton#embedButton {
    background-color: #059669;
    color: #ffffff;
    border: 1px solid #047857;
    font-weight: 600;
    min-width: 100px;
    padding: 12px 24px;
    font-size: 11pt;
}
QPushButton#embedButton:hover {
    background-color: #047857;
}
QPushButton#embedButton:pressed {
    background-color: #065f46;
}
QPushButton#embedButton:disabled {
    background-color: #064e3b;
    color: #065f46;
    border-color: #064e3b;
}
QPushButton#clearButton {
    background-color: transparent;
    border: 1px solid #3f3f46;
    color: #71717a;
    padding: 4px 10px;
    font-size: 8pt;
    min-width: 50px;
}
QPushButton#clearButton:hover {
    background-color: #27272a;
    color: #f87171;
    border-color: #f87171;
}

/* ── ComboBox ── */
QComboBox {
    background-color: #1a1a1e;
    color: #f4f4f5;
    border: 1px solid #2d2d30;
    border-radius: 6px;
    padding: 8px 12px;
    min-height: 20px;
}
QComboBox:hover {
    border-color: #52525b;
}
QComboBox:focus {
    border: 1px solid #666aba;
}
QComboBox::drop-down {
    border: none;
    padding-right: 10px;
}
QComboBox::down-arrow {
    image: url("SONGEMBED_ARROW_PATH");
    width: 12px;
    height: 12px;
}
QComboBox QAbstractItemView {
    background-color: #1a1a1e;
    color: #f4f4f5;
    border: 1px solid #3f3f46;
    selection-background-color: #666aba;
    selection-color: #ffffff;
    outline: 0;
    padding: 4px;
}

/* ── GroupBox ── */
QGroupBox {
    border: 1px solid #2d2d30;
    border-radius: 8px;
    margin-top: 12px;
    padding: 8px 12px;
    padding-top: 24px;
    font-weight: 500;
    color: #a1a1aa;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: #a1a1aa;
}

/* ── CheckBox ── */
QCheckBox {
    color: #a1a1aa;
    spacing: 6px;
    background-color: transparent;
}
QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border: 1px solid #3f3f46;
    border-radius: 4px;
    background-color: #1a1a1e;
}
QCheckBox::indicator:hover {
    border-color: #52525b;
}
QCheckBox::indicator:checked {
    background-color: #666aba;
    border-color: #53569c;
}

/* ── MessageBox ── */
QMessageBox {
    background-color: #121214;
}
QMessageBox QLabel {
    color: #e4e4e7;
}
QMessageBox QPushButton {
    min-width: 70px;
}

/* ── Sections Tree ── */
QTreeWidget#sectionsTree {
    background-color: #1a1a1e;
    border: 1px solid #2d2d30;
    border-radius: 6px;
    outline: none;
    padding: 2px;
}
QTreeWidget#sectionsTree::item {
    padding: 6px;
    border-radius: 4px;
    margin: 1px 0;
}
QTreeWidget#sectionsTree::item:hover {
    background-color: #27272a;
}
QTreeWidget#sectionsTree::item:selected {
    background-color: #53569c;
    color: #ffffff;
}

QHeaderView::section {
    background-color: #1a1a1e;
    color: #a1a1aa;
    font-weight: 600;
    font-size: 8pt;
    text-transform: uppercase;
    padding: 4px;
    border: none;
    border-bottom: 1px solid #2d2d30;
}
QWidget#clickableRow:hover {
    background-color: #27272a;
}

/* ── Verse Selector ── */
QWidget#verseSelectorPanel {
    background-color: #1a1a1e;
    border: 1px solid #2d2d30;
    border-radius: 6px;
    padding: 8px;
}
QPushButton#scanVersesButton, QPushButton#previewButton {
    background-color: #27272a;
    color: #a1a1aa;
    border: 1px solid #3f3f46;
    border-radius: 6px;
    padding: 6px 14px;
    font-size: 9pt;
}
QPushButton#scanVersesButton:hover, QPushButton#previewButton:hover {
    background-color: #3f3f46;
    color: #f4f4f5;
}
QPushButton#verseButton {
    background-color: #27272a;
    color: #a1a1aa;
    border: 1px solid #3f3f46;
    border-radius: 6px;
    padding: 4px 8px;
    font-weight: 500;
}
QPushButton#verseButton:hover {
    background-color: #3f3f46;
    color: #f4f4f5;
}
QPushButton#verseButton:checked {
    background-color: #666aba;
    color: #ffffff;
    border-color: #53569c;
}
QPushButton#firstLastButton, QPushButton#firstSecondLastButton, QPushButton#allVersesButton {
    background-color: transparent;
    color: #818cf8;
    border: 1px solid #53569c;
    border-radius: 6px;
    padding: 4px 8px;
    font-size: 9pt;
}
QPushButton#firstLastButton:hover, QPushButton#firstSecondLastButton:hover, QPushButton#allVersesButton:hover {
    background-color: #3730a3;
    color: #ffffff;
}
QPushButton#firstLastButton:checked, QPushButton#firstSecondLastButton:checked, QPushButton#allVersesButton:checked {
    background-color: #53569c;
    color: #ffffff;
}
QLabel#verseInfoLabel {
    color: #71717a;
    font-size: 9pt;
    font-style: italic;
    background-color: transparent;
}
"""





# ─────────────────────────────────────────────────────────────────────────────
# Update Checker Classes
# ─────────────────────────────────────────────────────────────────────────────
class UpdateCheckThread(QThread):
    result_ready = pyqtSignal(bool, str, str, str, str)  # success, version, notes, html_url, download_url

    def run(self):
        try:
            req = urllib.request.Request(
                "https://api.github.com/repos/MrKuenning/PPT-Song-Embed/releases/latest",
                headers={'User-Agent': f'SongEmbed/{APP_VERSION}'}
            )
            with urllib.request.urlopen(req, timeout=5) as response:
                data = json.loads(response.read().decode())
                
                tag_name = data.get('tag_name', '')
                if tag_name.startswith('v'):
                    tag_name = tag_name[1:]
                
                # Compare version numbers (e.g., '2.4.1' vs '2.4.0')
                is_newer = False
                if tag_name:
                    try:
                        latest_parts = [int(p) for p in tag_name.split('.')]
                        current_parts = [int(p) for p in APP_VERSION.split('.')]
                        is_newer = latest_parts > current_parts
                    except Exception:
                        # Fallback for non-standard version strings
                        is_newer = tag_name > APP_VERSION and tag_name != APP_VERSION

                if is_newer:
                    body = data.get('body', 'No release notes provided.')
                    html_url = data.get('html_url', '')
                    download_url = ""
                    assets = data.get('assets', [])
                    for asset in assets:
                        if asset.get('name', '').lower().endswith('.exe'):
                            download_url = asset.get('browser_download_url', '')
                            break
                    
                    self.result_ready.emit(True, tag_name, body, html_url, download_url)
                else:
                    self.result_ready.emit(True, "", "", "", "") # Up to date
        except Exception as e:
            self.result_ready.emit(False, str(e), "", "", "")

class DownloadThread(QThread):
    progress = pyqtSignal(int)
    finished_download = pyqtSignal(bool, str) # success, file_path_or_error

    def __init__(self, url, target_path):
        super().__init__()
        self.url = url
        self.target_path = target_path

    def run(self):
        try:
            req = urllib.request.Request(self.url, headers={'User-Agent': f'SongEmbed/{APP_VERSION}'})
            with urllib.request.urlopen(req, timeout=10) as response, open(self.target_path, 'wb') as out_file:
                total_length = response.info().get('Content-Length')
                if total_length is not None:
                    total_length = int(total_length)
                
                downloaded = 0
                chunk_size = 16384
                
                while True:
                    buffer = response.read(chunk_size)
                    if not buffer:
                        break
                    out_file.write(buffer)
                    downloaded += len(buffer)
                    if total_length:
                        percent = int((downloaded / total_length) * 100)
                        self.progress.emit(percent)
                        
            self.finished_download.emit(True, self.target_path)
        except Exception as e:
            self.finished_download.emit(False, str(e))

class UpdateDialog(QDialog):
    def __init__(self, version, notes, html_url, download_url, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Update Available")
        self.setMinimumSize(500, 400)
        self.html_url = html_url
        self.download_url = download_url
        self.version = version
        
        layout = QVBoxLayout(self)
        
        lbl = QLabel(f"A new version of Song Embed (v{version}) is available!")
        lbl.setStyleSheet("font-size: 11pt; font-weight: bold; color: #e4e4e7;")
        layout.addWidget(lbl)
        
        lbl_notes = QLabel("Release Notes:")
        lbl_notes.setStyleSheet("color: #a1a1aa;")
        layout.addWidget(lbl_notes)
        
        self.text_browser = QTextBrowser()
        self.text_browser.setMarkdown(notes)
        self.text_browser.setStyleSheet("background-color: #1a1a1e; color: #e4e4e7; border: 1px solid #2d2d30;")
        layout.addWidget(self.text_browser)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        self.skip_btn = QPushButton("Remind Me Later")
        self.skip_btn.clicked.connect(self.reject)
        
        self.web_btn = QPushButton("Open Webpage")
        self.web_btn.clicked.connect(self.open_webpage)
        
        btn_layout.addWidget(self.skip_btn)
        btn_layout.addWidget(self.web_btn)
        
        if self.download_url:
            self.dl_btn = QPushButton("Direct Download .EXE")
            self.dl_btn.setStyleSheet("background-color: #059669; color: #ffffff; font-weight: bold;")
            self.dl_btn.clicked.connect(self.download_exe)
            btn_layout.addWidget(self.dl_btn)
        else:
            self.web_btn.setStyleSheet("background-color: #059669; color: #ffffff; font-weight: bold;")

        layout.addLayout(btn_layout)
        
    def open_webpage(self):
        webbrowser.open(self.html_url)
        self.accept()

    def download_exe(self):
        base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        target_path = os.path.join(base_dir, f"SongEmbed_v{self.version}.exe")
        
        self.skip_btn.setEnabled(False)
        self.web_btn.setEnabled(False)
        self.dl_btn.setEnabled(False)
        self.dl_btn.setText("Downloading...")
        
        self.progress_dialog = QProgressDialog("Downloading update...", "Cancel", 0, 100, self)
        self.progress_dialog.setWindowTitle("Downloading")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setMinimumDuration(0)
        self.progress_dialog.canceled.connect(self.cancel_download)
        
        self.download_thread = DownloadThread(self.download_url, target_path)
        self.download_thread.progress.connect(self.progress_dialog.setValue)
        self.download_thread.finished_download.connect(self.on_download_finished)
        self.download_thread.start()

    def cancel_download(self):
        if hasattr(self, 'download_thread'):
            self.download_thread.terminate()
            self.download_thread.wait()
        self.skip_btn.setEnabled(True)
        self.web_btn.setEnabled(True)
        if hasattr(self, 'dl_btn'):
            self.dl_btn.setEnabled(True)
            self.dl_btn.setText("Direct Download .EXE")

    def on_download_finished(self, success, result):
        if not self.progress_dialog.wasCanceled():
            self.progress_dialog.close()
        
        if success:
            QMessageBox.information(self, "Download Complete", f"Update downloaded successfully to:\n{result}\n\nPlease close this application and run the new version.")
            self.accept()
        else:
            QMessageBox.warning(self, "Download Failed", f"Failed to download the update:\n{result}")
            self.cancel_download()

# ─────────────────────────────────────────────────────────────────────────────
# Custom delegate to display folder information alongside filename
# ─────────────────────────────────────────────────────────────────────────────
class FileItemDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        super().paint(painter, option, index)
        folder_display = index.data(Qt.ItemDataRole.UserRole + 1)
        if not folder_display:
            return

        painter.save()
        font = painter.font()
        font.setPointSize(font.pointSize() - 1)
        painter.setFont(font)

        if option.state & QStyle.StateFlag.State_Selected:
            painter.setPen(QColor("#c7d2fe"))
        else:
            painter.setPen(QColor("#71717a"))

        fm = painter.fontMetrics()
        fw = fm.horizontalAdvance(folder_display)
        text_rect = option.rect
        folder_x = text_rect.right() - fw - 10

        filename = index.data(Qt.ItemDataRole.DisplayRole)
        fname_w = fm.horizontalAdvance(filename)
        min_x = text_rect.left() + fname_w + 15

        if folder_x > min_x:
            draw_rect = QRect(folder_x, text_rect.top(), fw, text_rect.height())
            painter.drawText(draw_rect,
                             Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
                             folder_display)
        painter.restore()


# ─────────────────────────────────────────────────────────────────────────────
# Main Application
# ─────────────────────────────────────────────────────────────────────────────
class SongEmbedApp(QWidget):
    def __init__(self):
        super().__init__()
        self.settings = QSettings(ORG_NAME, APP_NAME)
        self.current_folder = ""
        self.all_files = []           # [(filename, full_path), ...]
        self.ppt_app = None           # PowerPoint COM object
        self.master_full_name = ""    # FullName of selected master PPT
        self.song_sections = []       # [(section_index, section_name), ...]
        self.slot_assignments = {}    # {section_name: (song_filename, song_path)}
        self.slot_widgets = []        # [(section_name, section_idx, name_lbl, song_lbl, clear_btn)]

        # Verse selection state
        self.parsed_song_path = ""           # Path of the song that was last scanned
        self.parsed_song_structure = []      # [{verse, slides, chorus_slides}, ...]
        self.verse_buttons = []              # List of checkable QPushButtons

        # Initialise COM
        if HAS_WIN32COM:
            try:
                pythoncom.CoInitialize()
                print("COM initialised.")
            except Exception as e:
                print(f"Failed to initialise COM: {e}")

        self.initUI()
        self.apply_dark_theme()
        self.load_initial_folder()
        self.refresh_open_ppts()
        self.setAcceptDrops(True)
        
        # Start update check if enabled
        if self.check_updates_cb.isChecked():
            QTimer.singleShot(3000, self.check_for_updates_auto)

    # ── UI Setup ─────────────────────────────────────────────────────────────

    def initUI(self):
        self.setWindowTitle(f"Song Embed v{APP_VERSION}")

        # Restore geometry
        geometry = self.settings.value("geometry")
        if geometry:
            try:
                self.restoreGeometry(geometry)
            except TypeError:
                if isinstance(geometry, QSize):
                    self.resize(geometry)
                else:
                    self.setGeometry(100, 100, WINDOW_WIDTH, WINDOW_HEIGHT)
        else:
            self.setGeometry(100, 100, WINDOW_WIDTH, WINDOW_HEIGHT)

        # Icon (reuse if present)
        base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_dir, 'icon.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # ── Main layout ──
        main = QVBoxLayout(self)
        main.setContentsMargins(15, 15, 15, 15)
        main.setSpacing(10)

        # ── Split layout ──
        self.split_layout = QSplitter(Qt.Orientation.Horizontal)
        main.addWidget(self.split_layout, 1)

        # ── Left Column ──
        self.left_col_widget = QWidget()
        left_col = QVBoxLayout(self.left_col_widget)
        left_col.setContentsMargins(0, 0, 10, 0)
        left_col.setSpacing(10)
        self.split_layout.addWidget(self.left_col_widget)

        # ── Right Column ──
        right_col_widget = QWidget()
        right_col = QVBoxLayout(right_col_widget)
        right_col.setContentsMargins(10, 0, 0, 0)
        right_col.setSpacing(10)
        self.split_layout.addWidget(right_col_widget)
        
        # Make left side static on window resize, right side dynamic
        self.split_layout.setStretchFactor(0, 0)
        self.split_layout.setStretchFactor(1, 1)
        self.split_layout.setCollapsible(0, False)
        
        # Set initial layout widths (approx 60% left, 40% right)
        self.split_layout.setSizes([600, 400])
        splitter_state = self.settings.value("splitter_state")
        if splitter_state:
            self.split_layout.restoreState(splitter_state)

        # ── Master PPT row ──
        lbl = QLabel("MASTER POWERPOINT")
        lbl.setObjectName("sectionLabel")
        left_col.addWidget(lbl)

        master_row = QHBoxLayout()
        master_row.setSpacing(8)

        self.master_combo = QComboBox()
        self.master_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.master_combo.addItem("— Select an open PowerPoint —")
        self.master_combo.currentIndexChanged.connect(self.on_master_selected)

        self.refresh_btn = QPushButton("🔄")
        self.refresh_btn.setToolTip("Refresh open PowerPoint files")
        self.refresh_btn.setFixedWidth(30)
        self.refresh_btn.clicked.connect(self.refresh_open_ppts)

        self.scan_master_verses_btn = QPushButton("Show Verses")
        self.scan_master_verses_btn.setToolTip("Scan master presentation sections for embedded verses")
        self.scan_master_verses_btn.clicked.connect(self.scan_master_verses)

        master_row.addWidget(self.master_combo, 2)
        master_row.addWidget(self.scan_master_verses_btn, 1)
        master_row.addWidget(self.refresh_btn)
        left_col.addLayout(master_row)

        # ── Sections panel ──
        self.slots_group = QGroupBox("Sections")
        slots_group_layout = QVBoxLayout()
        self.slots_group.setLayout(slots_group_layout)
        
        control_row = QHBoxLayout()
        self.add_section_btn = QPushButton("➕")
        self.add_section_btn.clicked.connect(self.add_section)
        self.remove_section_btn = QPushButton("➖")
        self.remove_section_btn.clicked.connect(self.remove_section)
        self.empty_section_btn = QPushButton("Empty")
        self.empty_section_btn.clicked.connect(self.empty_selected_section)
        self.move_up_btn = QPushButton("⬆ Up")
        self.move_up_btn.clicked.connect(self.move_section_up)
        self.move_down_btn = QPushButton("⬇ Down")
        self.move_down_btn.clicked.connect(self.move_section_down)
        control_row.addWidget(self.add_section_btn)
        control_row.addWidget(self.remove_section_btn)
        control_row.addWidget(self.empty_section_btn)
        control_row.addWidget(self.move_up_btn)
        control_row.addWidget(self.move_down_btn)
        slots_group_layout.addLayout(control_row)

        self.sections_tree = QTreeWidget()
        self.sections_tree.setObjectName("sectionsTree")
        self.sections_tree.setHeaderLabels(["SECTION", "TITLE / STATUS", "VERSES"])
        self.sections_tree.setIndentation(0)
        self.sections_tree.setAllColumnsShowFocus(True)
        self.sections_tree.setSelectionMode(QTreeWidget.SelectionMode.SingleSelection)
        self.sections_tree.itemClicked.connect(self._on_tree_item_clicked)
        
        # Set default column widths
        self.sections_tree.setColumnWidth(0, 150)
        self.sections_tree.setColumnWidth(1, 250)
        self.sections_tree.setColumnWidth(2, 80)
        
        tree_header_state = self.settings.value("tree_header_state")
        if tree_header_state:
            self.sections_tree.header().restoreState(tree_header_state)
            
        slots_group_layout.addWidget(self.sections_tree)
        left_col.addWidget(self.slots_group, 1)

        # ── Song Library ──
        lib_header = QHBoxLayout()
        lbl2 = QLabel("SONG LIBRARY")
        lbl2.setObjectName("sectionLabel")
        lib_header.addWidget(lbl2)
        lib_header.addStretch()

        self.keep_on_top_cb = QCheckBox("Keep on top")
        keep_val = self.settings.value("keep_on_top", True)
        if isinstance(keep_val, str):
            self.keep_on_top_cb.setChecked(keep_val.lower() == 'true')
        else:
            self.keep_on_top_cb.setChecked(bool(keep_val))
        self.keep_on_top_cb.stateChanged.connect(self.toggle_stay_on_top)
        lib_header.addWidget(self.keep_on_top_cb)

        self.settings_btn = QPushButton("⚙")
        self.settings_btn.setToolTip("Settings")
        self.settings_btn.setFixedWidth(30)
        self.settings_btn.clicked.connect(self.open_settings)
        lib_header.addWidget(self.settings_btn)

        right_col.addLayout(lib_header)

        # ── Settings Dialog Setup ──
        self.settings_dialog = QDialog(self)
        self.settings_dialog.setWindowTitle("Settings")
        self.settings_dialog.setMinimumWidth(450)
        settings_layout = QVBoxLayout(self.settings_dialog)

        folder_group = QGroupBox("Song Library")
        folder_layout = QVBoxLayout(folder_group)
        self.folder_label = QLabel(DEFAULT_FOLDER_TEXT)
        self.folder_label.setObjectName("folderLabel")
        self.folder_label.setWordWrap(True)

        folder_btn_layout = QHBoxLayout()
        self.folder_btn = QPushButton("📂 Browse Folder...")
        self.folder_btn.setToolTip("Select folder containing song PPTs")
        self.folder_btn.clicked.connect(self.select_folder)

        folder_btn_layout.addWidget(self.folder_btn)
        folder_btn_layout.addStretch()

        folder_layout.addWidget(self.folder_label)
        folder_layout.addLayout(folder_btn_layout)
        settings_layout.addWidget(folder_group)

        options_group = QGroupBox("General Options")
        settings_opts_layout = QVBoxLayout(options_group)

        self.replace_cb = QCheckBox("Replace Content")
        replace_val = self.settings.value("replace_existing", True)
        self.replace_cb.setChecked(str(replace_val).lower() == 'true' if isinstance(replace_val, str) else bool(replace_val))
        self.replace_cb.stateChanged.connect(lambda: self.settings.setValue("replace_existing", self.replace_cb.isChecked()))

        self.insert_blank_cb = QCheckBox("Add Blank Slide")
        blank_val = self.settings.value("insert_blank", True)
        self.insert_blank_cb.setChecked(str(blank_val).lower() == 'true' if isinstance(blank_val, str) else bool(blank_val))
        self.insert_blank_cb.stateChanged.connect(lambda: self.settings.setValue("insert_blank", self.insert_blank_cb.isChecked()))

        self.check_updates_cb = QCheckBox("Check updates on startup")
        updates_val = self.settings.value("check_updates", True)
        if isinstance(updates_val, str):
            self.check_updates_cb.setChecked(updates_val.lower() == 'true')
        else:
            self.check_updates_cb.setChecked(bool(updates_val))
        self.check_updates_cb.stateChanged.connect(lambda: self.settings.setValue("check_updates", self.check_updates_cb.isChecked()))

        settings_opts_layout.addWidget(self.replace_cb)
        settings_opts_layout.addWidget(self.insert_blank_cb)
        settings_opts_layout.addWidget(self.check_updates_cb)
        settings_layout.addWidget(options_group)

        close_btn_layout = QHBoxLayout()
        
        self.settings_version_label = QLabel(f"v{APP_VERSION}")
        self.settings_version_label.setObjectName("versionLabel")
        self.settings_version_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        
        self.settings_check_updates_btn = QPushButton("Check for Updates")
        self.settings_check_updates_btn.setObjectName("clearButton")
        self.settings_check_updates_btn.setToolTip("Check GitHub for a newer version")
        self.settings_check_updates_btn.clicked.connect(self.check_for_updates_manual)

        close_btn_layout.addWidget(self.settings_version_label)
        close_btn_layout.addWidget(self.settings_check_updates_btn)
        
        close_btn_layout.addStretch()
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.settings_dialog.accept)
        close_btn_layout.addWidget(close_btn)
        settings_layout.addLayout(close_btn_layout)

        search_row = QHBoxLayout()
        search_row.setSpacing(8)

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("🔍  Search songs by name…")
        self.search_box.textChanged.connect(self.filter_files)
        self.search_box.returnPressed.connect(self.select_first_match)
        
        self.inject_one_off_btn = QPushButton("Browse File")
        self.inject_one_off_btn.setToolTip("Select or drag & drop a PPT file anywhere here")
        self.inject_one_off_btn.clicked.connect(self.inject_one_off)
        
        search_row.addWidget(self.search_box)
        search_row.addWidget(self.inject_one_off_btn)
        right_col.addLayout(search_row)

        self.file_list = QListWidget()
        self.file_list.setAlternatingRowColors(True)
        self.file_list.itemSelectionChanged.connect(self.update_embed_state)
        self.file_list.itemSelectionChanged.connect(self._on_song_selection_changed)
        self.file_list.itemDoubleClicked.connect(self.on_double_click_embed)
        self.file_list.setItemDelegate(FileItemDelegate())
        right_col.addWidget(self.file_list, 1)

        # ── Verse Selector Panel ──
        self.verse_panel = QWidget()
        self.verse_panel.setObjectName("verseSelectorPanel")
        verse_panel_layout = QVBoxLayout(self.verse_panel)
        verse_panel_layout.setContentsMargins(8, 6, 8, 6)
        verse_panel_layout.setSpacing(6)

        verse_top_row = QHBoxLayout()
        verse_top_row.setSpacing(8)

        self.scan_verses_btn = QPushButton("🔍 Scan For Verses")
        self.scan_verses_btn.setObjectName("scanVersesButton")
        self.scan_verses_btn.setToolTip("Scan the selected song to detect verses and choruses")
        self.scan_verses_btn.setEnabled(False)
        self.scan_verses_btn.clicked.connect(self.scan_song_verses)

        self.preview_btn = QPushButton("👁 Preview Song")
        self.preview_btn.setObjectName("previewButton")
        self.preview_btn.setEnabled(False)
        self.preview_btn.clicked.connect(self.preview_selected)

        self.first_last_btn = QPushButton("1st && Last")
        self.first_last_btn.setObjectName("firstLastButton")
        self.first_last_btn.setToolTip("Auto-embed first and last verses only")
        self.first_last_btn.setCheckable(True)
        self.first_last_btn.clicked.connect(self.select_first_and_last)
        
        self.first_second_last_btn = QPushButton("1st, 2nd && Last")
        self.first_second_last_btn.setObjectName("firstSecondLastButton")
        self.first_second_last_btn.setToolTip("Auto-embed first two and last verses only")
        self.first_second_last_btn.setCheckable(True)
        self.first_second_last_btn.clicked.connect(self.select_first_second_and_last)
        
        self.all_verses_btn = QPushButton("All")
        self.all_verses_btn.setObjectName("allVersesButton")
        self.all_verses_btn.setToolTip("Auto-embed all verses")
        self.all_verses_btn.setCheckable(True)
        self.all_verses_btn.clicked.connect(self.select_all_verses)
        
        # Restore state
        fl_val = self.settings.value("first_last_toggle", False, type=bool)
        fsl_val = self.settings.value("first_second_last_toggle", False, type=bool)
        all_val = self.settings.value("all_verses_toggle", True, type=bool)
        
        # Ensure only one is checked on startup
        if all_val:
            fl_val = False
            fsl_val = False
        elif fl_val:
            fsl_val = False
            all_val = False
        elif fsl_val:
            fl_val = False
            all_val = False
        else:
            all_val = True
            
        self.first_last_btn.setChecked(fl_val)
        self.first_second_last_btn.setChecked(fsl_val)
        self.all_verses_btn.setChecked(all_val)

        verse_top_row.addWidget(self.scan_verses_btn)
        verse_top_row.addWidget(self.preview_btn)
        verse_top_row.addWidget(self.all_verses_btn)
        verse_top_row.addWidget(self.first_last_btn)
        verse_top_row.addWidget(self.first_second_last_btn)
        verse_top_row.addStretch()
        verse_panel_layout.addLayout(verse_top_row)

        self.verse_info_label = QLabel("")
        self.verse_info_label.setObjectName("verseInfoLabel")
        self.verse_info_label.hide()
        verse_panel_layout.addWidget(self.verse_info_label)

        # Row for dynamically generated verse buttons + shortcuts
        self.verse_buttons_row = QHBoxLayout()
        self.verse_buttons_row.setSpacing(6)
        verse_panel_layout.addLayout(self.verse_buttons_row)

        self.dynamic_verse_layout = QHBoxLayout()
        self.dynamic_verse_layout.setSpacing(6)
        self.verse_buttons_row.addLayout(self.dynamic_verse_layout)

        self.verse_buttons_row.addStretch()



        right_col.addWidget(self.verse_panel)

        # ── Action row ──
        action_row = QHBoxLayout()
        action_row.setSpacing(8)

        target_lbl = QLabel("Target:")
        self.target_combo = QComboBox()
        self.target_combo.setMinimumWidth(110)
        self.target_combo.addItem("— none —")
        self.target_combo.currentIndexChanged.connect(self._on_target_changed)

        self.embed_btn = QPushButton("📥 Embed")
        self.embed_btn.setObjectName("embedButton")
        self.embed_btn.setEnabled(False)
        self.embed_btn.clicked.connect(self.embed_selected)

        action_row.addWidget(self.embed_btn)
        action_row.addSpacerItem(
            QSpacerItem(20, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))
        action_row.addWidget(target_lbl)
        action_row.addWidget(self.target_combo)
        right_col.addLayout(action_row)

        # ── Options row ──
        options_row = QHBoxLayout()
        options_row.setSpacing(12)

        self.require_confirm_cb = QCheckBox("Confirmation")
        confirm_val = self.settings.value("require_confirm", True)
        self.require_confirm_cb.setChecked(str(confirm_val).lower() == 'true' if isinstance(confirm_val, str) else bool(confirm_val))
        self.require_confirm_cb.stateChanged.connect(lambda: self.settings.setValue("require_confirm", self.require_confirm_cb.isChecked()))

        self.auto_append_cb = QCheckBox("Auto New Section")
        append_val = self.settings.value("auto_append", False)
        self.auto_append_cb.setChecked(str(append_val).lower() == 'true' if isinstance(append_val, str) else bool(append_val))
        self.auto_append_cb.stateChanged.connect(lambda: self.settings.setValue("auto_append", self.auto_append_cb.isChecked()))
        self.auto_append_cb.stateChanged.connect(self.update_embed_state)

        options_row.addWidget(self.require_confirm_cb)
        options_row.addWidget(self.auto_append_cb)
        options_row.addStretch()
        right_col.addLayout(options_row)


        # ── Status bar & Version Info ──
        bottom_layout = QHBoxLayout()
        bottom_layout.setContentsMargins(0, 5, 0, 0)

        self.version_label = QLabel(f"v{APP_VERSION}")
        self.version_label.setObjectName("versionLabel")
        self.version_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        bottom_layout.addWidget(self.version_label)

        self.check_updates_btn = QPushButton("Check for Updates")
        self.check_updates_btn.setObjectName("clearButton")
        self.check_updates_btn.setToolTip("Check GitHub for a newer version")
        self.check_updates_btn.clicked.connect(self.check_for_updates_manual)
        bottom_layout.addWidget(self.check_updates_btn)

        bottom_layout.addSpacing(15)

        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        bottom_layout.addWidget(self.status_label)

        bottom_layout.addStretch()

        self.toggle_left_btn = QPushButton("▶ Hide Master Panel")
        self.toggle_left_btn.setObjectName("clearButton")
        self.toggle_left_btn.clicked.connect(self.toggle_left_panel)
        bottom_layout.addWidget(self.toggle_left_btn)

        main.addLayout(bottom_layout)

        # ── Window flags ──
        if self.keep_on_top_cb.isChecked():
            self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)

        self.search_box.installEventFilter(self)
        self.file_list.installEventFilter(self)
        self.search_box.setFocus()

    def eventFilter(self, source, event):
        if event.type() == QEvent.Type.KeyPress:
            if source is self.search_box:
                if event.key() == Qt.Key.Key_Down:
                    if self.file_list.count() > 0:
                        self.file_list.setFocus()
                        if not self.file_list.selectedItems():
                            self.file_list.setCurrentRow(0)
                        return True
            elif source is self.file_list:
                if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                    if self.embed_btn.isEnabled():
                        self.embed_selected()
                        return True
        return super().eventFilter(source, event)

    def apply_dark_theme(self):
        import tempfile
        import base64
        arrow_base64 = b"iVBORw0KGgoAAAANSUhEUgAAAAwAAAAMCAYAAABWdVznAAAAOUlEQVR4nGNgGPqAEcZYuHDVf0KK4+PDGJmQOYQUMzAwMDBhE8SlGEMDMQBDA7othJwKB8QEAm0AAJ0jDXM9WGTXAAAAAElFTkSuQmCC"
        arrow_path = os.path.join(tempfile.gettempdir(), "songembed_down_arrow.png")
        try:
            with open(arrow_path, "wb") as f:
                f.write(base64.b64decode(arrow_base64))
        except Exception:
            pass
            
        css = DARK_STYLESHEET.replace("SONGEMBED_ARROW_PATH", arrow_path.replace("\\", "/"))
        self.setStyleSheet(css)
        QApplication.instance().setStyleSheet(css)

    def open_settings(self):
        self.settings_dialog.exec()

    # ── Stay‐on‐top ─────────────────────────────────────────────────────────

    def toggle_stay_on_top(self):
        flags = self.windowFlags()
        if self.keep_on_top_cb.isChecked():
            flags |= Qt.WindowType.WindowStaysOnTopHint
        else:
            flags &= ~Qt.WindowType.WindowStaysOnTopHint
        self.setWindowFlags(flags)
        self.show()
        self.settings.setValue("keep_on_top", self.keep_on_top_cb.isChecked())

    # ── Status helpers ───────────────────────────────────────────────────────

    def set_status(self, text, error=False, success=False):
        self.status_label.setText(text)
        if error:
            self.status_label.setObjectName("statusError")
        elif success:
            self.status_label.setObjectName("statusSuccess")
        else:
            self.status_label.setObjectName("statusLabel")
        # Force stylesheet refresh on objectName change
        self.status_label.setStyleSheet(self.status_label.styleSheet())
        print(f"[STATUS] {text}")

    # ── COM Helpers ──────────────────────────────────────────────────────────

    def _get_ppt_app(self):
        """Get or create the PowerPoint Application COM object."""
        if not HAS_WIN32COM:
            return None

        if self.ppt_app:
            try:
                _ = self.ppt_app.Presentations.Count
                return self.ppt_app
            except Exception:
                self.ppt_app = None

        try:
            app = win32com.client.GetActiveObject("PowerPoint.Application")
            print("Attached to existing PowerPoint instance.")
        except Exception:
            try:
                app = win32com.client.Dispatch("PowerPoint.Application")
                print("Created new PowerPoint instance.")
            except Exception as e:
                print(f"Error creating PowerPoint instance: {e}")
                return None

        self.ppt_app = app
        return app

    def _get_master_pres(self):
        """Get the COM object for the currently selected master presentation."""
        if not self.master_full_name:
            return None

        ppt_app = self._get_ppt_app()
        if not ppt_app:
            return None

        try:
            for i in range(1, ppt_app.Presentations.Count + 1):
                pres = ppt_app.Presentations.Item(i)
                if pres.FullName == self.master_full_name:
                    return pres
        except Exception as e:
            print(f"Error finding master presentation: {e}")

        return None

    # ── Master PPT Selection ─────────────────────────────────────────────────

    def refresh_open_ppts(self):
        """Enumerate open PowerPoint presentations and populate the combo."""
        previous_full_name = self.master_full_name
        self.master_combo.blockSignals(True)
        self.master_combo.clear()
        self.master_combo.addItem("— Select an open PowerPoint —")

        ppt_app = self._get_ppt_app()
        if not ppt_app:
            self.master_combo.blockSignals(False)
            if not HAS_WIN32COM:
                self.set_status("pywin32 not available — PowerPoint automation disabled", error=True)
            else:
                self.set_status("No PowerPoint instance running")
            return

        try:
            count = ppt_app.Presentations.Count
            reselect_index = 0
            for i in range(1, count + 1):
                pres = ppt_app.Presentations.Item(i)
                name = pres.Name
                full_name = pres.FullName
                self.master_combo.addItem(name, full_name)
                if full_name == previous_full_name:
                    reselect_index = i  # combo index (1-based since placeholder at 0)

            if reselect_index > 0:
                self.master_combo.setCurrentIndex(reselect_index)
            elif count == 1:
                # Auto-select if only one PPT is open
                self.master_combo.setCurrentIndex(1)

            self.set_status(f"Found {count} open presentation{'s' if count != 1 else ''}")
        except Exception as e:
            self.set_status(f"Error enumerating presentations: {e}", error=True)

        self.master_combo.blockSignals(False)
        # Trigger section read for whatever is now selected
        self.on_master_selected(self.master_combo.currentIndex())

    def on_master_selected(self, index):
        """Handle master PPT combo selection change."""
        if index <= 0:
            self.master_full_name = ""
            self.song_sections = []
            self.slot_assignments.clear()
            self.rebuild_slots_display()
            self.update_target_combo()
            self.update_embed_state()
            return

        full_name = self.master_combo.itemData(index)
        if full_name:
            self.master_full_name = full_name
            print(f"Master selected: {full_name}")
            self.read_sections()
            self.rebuild_slots_display()
            self.update_target_combo()
            self.update_embed_state()

    # ── Section Reading ──────────────────────────────────────────────────────

    def read_sections(self):
        """Read all sections from master PPT."""
        self.song_sections = []

        pres = self._get_master_pres()
        if not pres:
            self.set_status("Could not access master presentation", error=True)
            return

        try:
            sp = pres.SectionProperties
            total = sp.Count
            print(f"Master has {total} sections")
            for i in range(1, total + 1):
                name = sp.Name(i)
                slide_count = sp.SlidesCount(i)
                self.song_sections.append((i, name.strip()))
                print(f"  Section {i}: '{name}' ({slide_count} slides)")

            if not self.song_sections:
                self.set_status(
                    "No sections found in master presentation.",
                    error=True)
            else:
                self.set_status(
                    f"Found {len(self.song_sections)} section"
                    f"{'s' if len(self.song_sections) != 1 else ''}",
                    success=True)

        except Exception as e:
            self.set_status(f"Error reading sections: {e}", error=True)

    # ── Slots Display ────────────────────────────────────────────────────────

    def _get_section_title(self, section_idx):
        pres = self._get_master_pres()
        if not pres:
            return ""
        try:
            sp = pres.SectionProperties
            slide_count = sp.SlidesCount(section_idx)
            if slide_count == 0:
                return ""
            
            # Grab title from the second slide if possible, else the first
            target_slide_idx = sp.FirstSlide(section_idx)
            if slide_count > 1:
                target_slide_idx += 1
                
            slide = pres.Slides(target_slide_idx)
            for i in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(i)
                if shape.HasTextFrame:
                    if shape.TextFrame.HasText:
                        text = shape.TextFrame.TextRange.Text.strip()
                        if text:
                            return text.split('\n')[0]
        except Exception:
            pass
        return ""

    def scan_master_verses(self):
        """Scan the master presentation to populate verses for song sections."""
        pres = self._get_master_pres()
        if not pres:
            self.set_status("No master presentation selected", error=True)
            return
            
        self.set_status("Scanning master presentation for verses...", success=True)
        QApplication.processEvents()
        
        scanned = 0
        for widget_info in self.slot_widgets:
            section_name = widget_info[0]
            section_idx = widget_info[1]
            verses_lbl = widget_info[4]
            
            if verses_lbl is not None and "song" in section_name.lower():
                verses_text = self._get_section_verses(section_idx)
                verses_lbl.setText(verses_text)
                scanned += 1
                
        self.set_status(f"Finished scanning {scanned} song section{'s' if scanned != 1 else ''} for verses", success=True)

    def _get_section_verses(self, section_idx):
        """Parse slides in a specific master section to determine which verses are present."""
        pres = self._get_master_pres()
        if not pres:
            return ""
        try:
            sp = pres.SectionProperties
            slide_count = sp.SlidesCount(section_idx)
            if slide_count == 0:
                return ""
            
            first_idx = sp.FirstSlide(section_idx)
            verses = set()
            
            last_tag_type = "other"
            last_tag_value = ""
            
            for i in range(first_idx, first_idx + slide_count):
                slide = pres.Slides(i)
                title_text = ""
                try:
                    if slide.Shapes.HasTitle:
                        title_text = slide.Shapes.Title.TextFrame.TextRange.Text.strip()
                        
                    if not title_text:
                        for j in range(1, slide.Shapes.Count + 1):
                            shape = slide.Shapes(j)
                            if shape.HasTextFrame and shape.TextFrame.HasText:
                                text = shape.TextFrame.TextRange.Text.strip()
                                if text:
                                    title_text = text
                                    break
                except Exception:
                    pass

                tag_type = "other"
                tag_value = ""

                if title_text:
                    prefix = title_text.split(" - ")[0].strip().lower() if " - " in title_text else title_text.split()[0].strip().lower()
                    if prefix.isdigit():
                        tag_type = "verse"
                        tag_value = int(prefix)
                    elif prefix == "c":
                        tag_type = "chorus"
                        tag_value = 0

                if tag_type == "other" and last_tag_type != "other":
                    tag_type = last_tag_type
                    tag_value = last_tag_value
                    
                if tag_type == "verse":
                    verses.add(tag_value)
                    
                last_tag_type = tag_type
                last_tag_value = tag_value
                
            if not verses:
                return ""
                
            sorted_verses = sorted(list(verses))
            return ", ".join(str(v) for v in sorted_verses)
            
        except Exception as e:
            print(f"Error getting section verses: {e}")
            return ""

    def _clear_layout(self, layout):
        """Recursively remove all items from a layout."""
        while layout.count():
            item = layout.takeAt(0)
            w = item.widget()
            if w:
                w.setParent(None)
                w.deleteLater()
            sub = item.layout()
            if sub:
                self._clear_layout(sub)

    def rebuild_slots_display(self):
        """Rebuild the section slots panel."""
        self.sections_tree.clear()
        self.slot_widgets = []

        if not self.song_sections:
            item = QTreeWidgetItem()
            if self.master_full_name:
                item.setText(0, "No sections found in master presentation.")
            else:
                item.setText(0, "Select a master PPT to see sections")
            item.setFirstColumnSpanned(True)
            self.sections_tree.addTopLevelItem(item)
            return

        pres = self._get_master_pres()

        for section_idx, section_name in self.song_sections:
            item = QTreeWidgetItem()
            item.setData(0, Qt.ItemDataRole.UserRole, section_idx)
            self.sections_tree.addTopLevelItem(item)

            display_name = f"🎵 {section_name}:" if "song" in section_name.lower() else f"{section_name}:"
            name_lbl = QLabel(display_name)
            name_lbl.setObjectName("slotName")

            # Determine display text
            assignment = self.slot_assignments.get(section_name)
            if assignment:
                song_lbl = QLabel(f"✓ {assignment[0]}")
                song_lbl.setObjectName("slotAssigned")
            else:
                slide_count = 0
                title_text = ""
                if pres:
                    try:
                        slide_count = pres.SectionProperties.SlidesCount(section_idx)
                        title_text = self._get_section_title(section_idx)
                    except Exception:
                        pass
                if slide_count > 0:
                    display_text = f"({slide_count} slide{'s' if slide_count != 1 else ''})"
                    if title_text:
                        # Truncate title if too long
                        if len(title_text) > 30:
                            title_text = title_text[:27] + "..."
                        display_text = f"{title_text}  {display_text}"
                    song_lbl = QLabel(display_text)
                    song_lbl.setObjectName("slotFilled")
                else:
                    song_lbl = QLabel("— Empty —")
                    song_lbl.setObjectName("slotEmpty")

            verses_text = ""
            verses_lbl = QLabel(verses_text)
            verses_lbl.setObjectName("slotVerses")
            verses_lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

            self.sections_tree.setItemWidget(item, 0, name_lbl)
            self.sections_tree.setItemWidget(item, 1, song_lbl)
            self.sections_tree.setItemWidget(item, 2, verses_lbl)

            self.slot_widgets.append((section_name, section_idx, name_lbl, song_lbl, verses_lbl, item))

    def update_target_combo(self):
        """Populate the target slot dropdown from song sections."""
        current_data = self.target_combo.currentData()
        
        self.target_combo.blockSignals(True)
        self.target_combo.clear()

        if not self.song_sections:
            self.target_combo.addItem("— none —")
        else:
            for section_idx, section_name in self.song_sections:
                self.target_combo.addItem(section_name, (section_idx, section_name))
                
            if current_data:
                idx = self.target_combo.findText(current_data[1])
                if idx >= 0:
                    self.target_combo.setCurrentIndex(idx)

        self.target_combo.blockSignals(False)
        self._highlight_selected_slot()

    def _on_tree_item_clicked(self, item, column):
        section_idx = item.data(0, Qt.ItemDataRole.UserRole)
        if section_idx is not None:
            self.select_slot_by_index(section_idx)

    def select_slot_by_index(self, section_idx):
        for i in range(self.target_combo.count()):
            data = self.target_combo.itemData(i)
            if data and data[0] == section_idx:
                self.target_combo.setCurrentIndex(i)
                break

    def _highlight_selected_slot(self):
        target_data = self.target_combo.currentData()
        target_idx = target_data[0] if target_data else -1

        self.sections_tree.blockSignals(True)
        for widget_info in self.slot_widgets:
            idx = widget_info[1]
            name_lbl = widget_info[2]
            song_lbl = widget_info[3]
            verses_lbl = widget_info[4]
            item = widget_info[5]
            
            if idx == target_idx:
                item.setSelected(True)
                name_lbl.setStyleSheet("color: #ffffff; font-weight: 600;")
                song_lbl.setStyleSheet("color: #ffffff;")
                verses_lbl.setStyleSheet("color: #ffffff; font-size: 9pt;")
            else:
                item.setSelected(False)
                name_lbl.setStyleSheet("")
                song_lbl.setStyleSheet("")
                verses_lbl.setStyleSheet("")
        self.sections_tree.blockSignals(False)

    def _on_target_changed(self, index):
        self._highlight_selected_slot()
        self.update_embed_state()

    # ── Clear Slot ───────────────────────────────────────────────────────────

    def clear_slot(self, section_name, section_idx):
        """Remove all slides from a song section in the master PPT."""
        reply = QMessageBox.question(
            self,
            "Confirm Empty",
            f"Are you sure you want to empty the section '{section_name}'?\nThis will delete all slides in this section.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.No:
            return

        pres = self._get_master_pres()
        if not pres:
            self.set_status("Master presentation not accessible", error=True)
            return

        try:
            sp = pres.SectionProperties

            # Re-find the section (indices may have shifted)
            actual_idx = self._find_section_index(sp, section_name)
            if actual_idx is None:
                self.set_status(f"Section '{section_name}' not found", error=True)
                return

            slide_count = sp.SlidesCount(actual_idx)
            if slide_count == 0:
                self.set_status(f"Section '{section_name}' is already empty")
                self.slot_assignments.pop(section_name, None)
                self.rebuild_slots_display()
                return

            first_slide = sp.FirstSlide(actual_idx)

            # Delete slides backwards
            for i in range(slide_count - 1, -1, -1):
                pres.Slides(first_slide + i).Delete()

            self.slot_assignments.pop(section_name, None)
            self.set_status(f"Cleared {slide_count} slide{'s' if slide_count != 1 else ''} from {section_name}", success=True)

            # Re-read sections (indices shift after deletion) and rebuild display
            self.read_sections()
            self.rebuild_slots_display()
            self.update_target_combo()

        except Exception as e:
            self.set_status(f"Error clearing slot: {e}", error=True)

    def _find_section_index(self, sp, section_name):
        """Find the current index of a section by name (sections can shift)."""
        for i in range(1, sp.Count + 1):
            if sp.Name(i).strip().lower() == section_name.strip().lower():
                return i
        return None

    # ── Song Library (Folder / Search / List) ────────────────────────────────

    def load_initial_folder(self):
        saved = self.settings.value("last_folder", "")
        if saved and os.path.isdir(saved):
            print(f"Loading previous folder: {saved}")
            self.current_folder = saved
            self.folder_label.setText(f"Folder: {self.current_folder}")
            self.scan_folder()
        else:
            self.folder_label.setText(DEFAULT_FOLDER_TEXT)

    def select_folder(self):
        start_dir = self.current_folder or os.path.expanduser("~")
        folder = QFileDialog.getExistingDirectory(
            self, "Select Folder Containing Song PPTs", start_dir)
        if folder:
            self.current_folder = folder
            self.folder_label.setText(f"Folder: {self.current_folder}")
            self.scan_folder()
            self.settings.setValue("last_folder", self.current_folder)
        elif not self.current_folder:
            self.folder_label.setText(DEFAULT_FOLDER_TEXT)

    def scan_folder(self):
        self.all_files = []
        self.file_list.clear()

        if not self.current_folder or not os.path.isdir(self.current_folder):
            return

        print(f"Scanning: {self.current_folder}")
        try:
            for root, _, files in os.walk(self.current_folder):
                for fn in files:
                    if fn.lower().endswith(ALLOWED_EXTENSIONS) and not fn.startswith('~'):
                        self.all_files.append((fn, os.path.join(root, fn)))
            self.all_files.sort(key=lambda x: x[0].lower())
            print(f"Found {len(self.all_files)} song file(s).")
            self.filter_files()
        except Exception as e:
            QMessageBox.critical(self, "Scan Error", f"Error scanning folder:\n{e}")

        self.search_box.setEnabled(bool(self.all_files))
        self.file_list.setEnabled(bool(self.all_files))
        self.update_embed_state()

        if not self.all_files:
            self.file_list.addItem("No PPT/PPTX files found in this folder.")

    def filter_files(self):
        search = self.search_box.text().lower().strip()
        self.file_list.clear()

        if not self.all_files and self.current_folder:
            self.file_list.addItem("No PPT/PPTX files found.")
            self.file_list.setEnabled(False)
            self.update_embed_state()
            return
        elif not self.current_folder:
            return
        else:
            self.file_list.setEnabled(True)

        words = search.split()
        match_count = 0

        for filename, full_path in self.all_files:
            fn_lower = filename.lower()
            if not words or all(w in fn_lower for w in words):
                # Folder display text
                folder_path = os.path.dirname(full_path)
                if folder_path.startswith(self.current_folder):
                    rel = folder_path[len(self.current_folder):].lstrip(os.sep)
                    folder_display = f" - [{rel}]" if rel else " - [Root]"
                else:
                    folder_display = f" - [{os.path.basename(folder_path)}]"

                item = QListWidgetItem(filename)
                item.setData(Qt.ItemDataRole.UserRole, full_path)
                item.setData(Qt.ItemDataRole.UserRole + 1, folder_display)
                item.setToolTip(full_path)
                self.file_list.addItem(item)
                match_count += 1

        if match_count == 0 and self.all_files:
            no_match = QListWidgetItem(f"No matches for '{search}'")
            no_match.setForeground(QColor("#71717a"))
            self.file_list.addItem(no_match)

        self.update_embed_state()

    def select_first_match(self):
        """Select the first item in the list (triggered by Enter in search)."""
        if self.file_list.count() > 0:
            first = self.file_list.item(0)
            if first.data(Qt.ItemDataRole.UserRole):
                self.file_list.setCurrentItem(first)

    def toggle_left_panel(self):
        """Toggle the visibility of the left master panel and shrink/expand the window."""
        handle_width = self.split_layout.handleWidth()
        
        if self.left_col_widget.isHidden():
            old_pos = self.pos()
            old_width = self.width()
            
            self.left_col_widget.show()
            
            added_width = getattr(self, 'saved_left_width', 600)
            grow_amount = added_width + handle_width
            
            # Temporarily remove minimum width to ensure seamless resize
            self.setMinimumWidth(100)
            
            # Expand window and move left to keep right edge perfectly anchored
            self.move(old_pos.x() - grow_amount, old_pos.y())
            self.resize(old_width + grow_amount, self.height())
            
            self.setMinimumWidth(0) # Restore default layout constraint
            
            sizes = self.split_layout.sizes()
            sizes[0] = added_width
            self.split_layout.setSizes(sizes)
            
            self.toggle_left_btn.setText("▶ Hide Master Panel")
        else:
            # Save current width of left panel before hiding
            sizes = self.split_layout.sizes()
            if sizes[0] > 0:
                self.saved_left_width = sizes[0]
            else:
                self.saved_left_width = 600
                
            old_pos = self.pos()
            old_width = self.width()
            
            self.left_col_widget.hide()
            
            # Include the splitter handle width in the total amount we are removing
            shrink_amount = self.saved_left_width + handle_width
            target_width = old_width - shrink_amount
            
            # Temporarily override layout minimum width to allow the window to shrink down
            # exactly to the width of the song library panel.
            self.setMinimumWidth(100)
            
            # Shrink window width and move window right by exact amount
            self.move(old_pos.x() + shrink_amount, old_pos.y())
            self.resize(target_width, self.height())
            
            self.setMinimumWidth(0) # Restore default layout constraint
            
            self.toggle_left_btn.setText("◀ Show Master Panel")

    # ── Embed State ──────────────────────────────────────────────────────────

    def update_embed_state(self):
        """Enable/disable the Embed button based on current state."""
        selected = self.file_list.selectedItems()
        has_song = (len(selected) == 1 and
                    selected[0].data(Qt.ItemDataRole.UserRole) is not None)
        has_master = bool(self.master_full_name)
        has_target = (self.target_combo.currentIndex() >= 0 and
                      self.target_combo.currentData() is not None)
                      
        auto_append = self.auto_append_cb.isChecked()

        self.preview_btn.setEnabled(has_song)
        self.scan_verses_btn.setEnabled(has_song)
        
        can_embed = has_song and has_master and (has_target or auto_append)
        self.embed_btn.setEnabled(can_embed)
        
        # Update text and styling of embed button based on auto section mode
        if auto_append:
            if has_target:
                target_name = self.target_combo.currentText()
                self.embed_btn.setText(f'Add New Section After "{target_name}"')
            else:
                self.embed_btn.setText('Add New Section')
                
            self.embed_btn.setStyleSheet("""
                QPushButton#embedButton {
                    background-color: #2563eb;
                    color: #ffffff;
                    border: 1px solid #2d4373;
                    font-weight: 600;
                    min-width: 100px;
                    padding: 12px 24px;
                    font-size: 11pt;
                }
                QPushButton#embedButton:hover {
                    background-color: #2d4373;
                }
                QPushButton#embedButton:pressed {
                    background-color: #233352;
                }
                QPushButton#embedButton:disabled {
                    background-color: #1e3a8a;
                    color: #233352;
                    border-color: #1e3a8a;
                }
            """)
        else:
            if has_target:
                target_name = self.target_combo.currentText()
                self.embed_btn.setText(f'Embed in Section "{target_name}"')
            else:
                self.embed_btn.setText("📥 Embed")
            self.embed_btn.setStyleSheet("")

    # ── Preview Logic ────────────────────────────────────────────────────────

    def preview_selected(self):
        """Open the selected song in PowerPoint for preview."""
        selected = self.file_list.selectedItems()
        if not selected:
            return

        song_path = selected[0].data(Qt.ItemDataRole.UserRole)
        if not song_path or not os.path.exists(song_path):
            self.set_status("Song file not found", error=True)
            return

        self.set_status(f"Opening preview...")
        QApplication.processEvents()
        
        try:
            if HAS_WIN32COM:
                ppt_app = self._get_ppt_app()
                if ppt_app:
                    normalised = os.path.abspath(os.path.normpath(song_path))
                    ppt_app.Presentations.Open(normalised, ReadOnly=True, WithWindow=True)
                    self.set_status("Preview opened", success=True)
                    return

            # Fallback to os default
            os.startfile(song_path)
            self.set_status("Preview opened", success=True)
        except Exception as e:
            self.set_status(f"Error opening preview: {e}", error=True)

    # ── Verse Selection ──────────────────────────────────────────────────────

    def parse_song_structure(self, song_path):
        """
        Open a song PPT and parse its slides to detect verse/chorus structure.

        Returns a list of dicts:
        [
            {"verse": 1, "slides": [1, 2], "chorus_slides": [3, 4]},
            {"verse": 2, "slides": [5, 6], "chorus_slides": [7, 8]},
            ...
        ]

        Slide numbering is 1-based (PowerPoint convention).
        """
        if not HAS_WIN32COM:
            return []

        ppt_app = self._get_ppt_app()
        if not ppt_app:
            return []

        try:
            normalised = os.path.abspath(os.path.normpath(song_path))
            song_pres = ppt_app.Presentations.Open(normalised, ReadOnly=True, WithWindow=False)
        except Exception as e:
            print(f"Error opening song for scanning: {e}")
            return []

        try:
            slide_count = song_pres.Slides.Count
            if slide_count == 0:
                return []

            # Read the title prefix of every slide
            slide_tags = []  # [(slide_num, tag_type, tag_value)]
            last_tag_type = "other"
            last_tag_value = ""
            
            for i in range(1, slide_count + 1):
                slide = song_pres.Slides(i)
                title_text = ""
                try:
                    if slide.Shapes.HasTitle:
                        title_text = slide.Shapes.Title.TextFrame.TextRange.Text.strip()
                        
                    if not title_text:
                        for j in range(1, slide.Shapes.Count + 1):
                            shape = slide.Shapes(j)
                            if shape.HasTextFrame and shape.TextFrame.HasText:
                                text = shape.TextFrame.TextRange.Text.strip()
                                if text:
                                    title_text = text
                                    break
                except Exception:
                    pass

                tag_type = "other"
                tag_value = ""

                if title_text:
                    prefix = title_text.split(" - ")[0].strip().lower() if " - " in title_text else title_text.split()[0].strip().lower()
                    if prefix.isdigit():
                        tag_type = "verse"
                        tag_value = int(prefix)
                    elif prefix == "c":
                        tag_type = "chorus"
                        tag_value = 0

                # Inherit the previous tag if this slide doesn't explicitly start a new section
                if tag_type == "other" and last_tag_type != "other":
                    tag_type = last_tag_type
                    tag_value = last_tag_value

                slide_tags.append((i, tag_type, tag_value))
                last_tag_type = tag_type
                last_tag_value = tag_value

            # Group into segments: consecutive slides with the same tag
            segments = []  # [{"type": "verse"/"chorus"/"other", "value": int, "slides": [1,2,...]}]
            for slide_num, tag_type, tag_value in slide_tags:
                if segments and segments[-1]["type"] == tag_type and segments[-1]["value"] == tag_value:
                    segments[-1]["slides"].append(slide_num)
                else:
                    segments.append({"type": tag_type, "value": tag_value, "slides": [slide_num]})

            # Build verse structures: each verse segment followed by its chorus
            verses = []
            i = 0
            while i < len(segments):
                seg = segments[i]
                if seg["type"] == "verse":
                    verse_entry = {
                        "verse": seg["value"],
                        "slides": list(seg["slides"]),
                        "chorus_slides": []
                    }
                    # Check if next segment is a chorus
                    if i + 1 < len(segments) and segments[i + 1]["type"] == "chorus":
                        verse_entry["chorus_slides"] = list(segments[i + 1]["slides"])
                        i += 2
                    else:
                        i += 1
                    verses.append(verse_entry)
                else:
                    i += 1

            return verses

        except Exception as e:
            print(f"Error parsing song structure: {e}")
            return []
        finally:
            try:
                song_pres.Close()
            except Exception:
                pass

    def scan_song_verses(self):
        """Scan the currently selected song for verse/chorus structure."""
        selected = self.file_list.selectedItems()
        if not selected:
            return

        song_path = selected[0].data(Qt.ItemDataRole.UserRole)
        if not song_path or not os.path.exists(song_path):
            self.set_status("Song file not found", error=True)
            return

        self.set_status("Scanning song for verses…")
        QApplication.processEvents()

        structure = self.parse_song_structure(song_path)

        self.parsed_song_path = song_path
        self.parsed_song_structure = structure

        self._rebuild_verse_buttons()

        if structure:
            count = len(structure)
            self.set_status(f"Scanned: {count} verse{'s' if count != 1 else ''} detected", success=True)
        else:
            self.set_status("No verse structure detected in this song")
        self._update_verse_info_label()

    def _rebuild_verse_buttons(self):
        """Create or destroy verse toggle buttons based on parsed structure."""
        # Clear existing verse buttons
        self._clear_verse_buttons()

        if not self.parsed_song_structure:
            self.first_last_btn.setVisible(False)
            self.first_second_last_btn.setVisible(False)
            self.all_verses_btn.setVisible(False)
            return

        self.first_last_btn.setVisible(True)
        self.first_second_last_btn.setVisible(True)
        self.all_verses_btn.setVisible(True)

        for entry in self.parsed_song_structure:
            verse_num = entry["verse"]
            total_slides = len(entry["slides"]) + len(entry["chorus_slides"])
            btn = QPushButton(f"Verse {verse_num}")
            btn.setObjectName("verseButton")
            btn.setCheckable(True)
            btn.setChecked(True)
            btn.setToolTip(f"Verse {verse_num}: {total_slides} slide{'s' if total_slides != 1 else ''}")
            btn.toggled.connect(self._update_verse_info_label)
            self.verse_buttons.append(btn)
            self.dynamic_verse_layout.addWidget(btn)

        # Apply the toggle state if it's currently checked
        if self.first_last_btn.isChecked():
            self.select_first_and_last(True)
        elif self.first_second_last_btn.isChecked():
            self.select_first_second_and_last(True)
        elif self.all_verses_btn.isChecked():
            self.select_all_verses(True)

    def _clear_verse_buttons(self):
        """Remove all verse toggle buttons from the layout."""
        for btn in self.verse_buttons:
            self.dynamic_verse_layout.removeWidget(btn)
            btn.setParent(None)
            btn.deleteLater()
        self.verse_buttons = []

    def _reset_verse_panel(self):
        """Reset the verse selection panel to its initial empty state."""
        self.parsed_song_path = ""
        self.parsed_song_structure = []
        self._clear_verse_buttons()
        self.verse_info_label.setText("")
        self.verse_info_label.hide()

    def _update_verse_info_label(self):
        """Update the label text based on current selection."""
        if not self.parsed_song_structure:
            self.verse_info_label.setText("No verse structure detected — all slides will be embedded")
            self.verse_info_label.show()
            return
        
        all_checked = all(btn.isChecked() for btn in self.verse_buttons)
        
        if all_checked:
            total_slides = sum(len(entry["slides"]) + len(entry["chorus_slides"]) for entry in self.parsed_song_structure)
            count = len(self.parsed_song_structure)
            has_chorus = any(len(entry["chorus_slides"]) > 0 for entry in self.parsed_song_structure)
            chorus_text = " - Chorus Detected" if has_chorus else ""
            self.verse_info_label.setText(f"{total_slides} Slides - {count} Verse{'s' if count != 1 else ''}{chorus_text}")
        else:
            slide_indices = self._get_selected_slide_indices()
            num_slides = len(slide_indices) if slide_indices else 0
            self.verse_info_label.setText(f"{num_slides} Slide{'s' if num_slides != 1 else ''} selected")
            
        self.verse_info_label.show()

    def select_first_and_last(self, checked):
        """Toggle verse buttons to select first and last verse if checked, else all verses."""
        if checked:
            self.first_second_last_btn.blockSignals(True)
            self.first_second_last_btn.setChecked(False)
            self.first_second_last_btn.blockSignals(False)
            self.settings.setValue("first_second_last_toggle", False)
            
            self.all_verses_btn.blockSignals(True)
            self.all_verses_btn.setChecked(False)
            self.all_verses_btn.blockSignals(False)
            self.settings.setValue("all_verses_toggle", False)
        elif not self.first_second_last_btn.isChecked():
            # If unchecked and the other isn't checked, default to All
            self.all_verses_btn.blockSignals(True)
            self.all_verses_btn.setChecked(True)
            self.all_verses_btn.blockSignals(False)
            self.settings.setValue("all_verses_toggle", True)
            
        self.settings.setValue("first_last_toggle", checked)

        if not self.verse_buttons:
            return
            
        if checked:
            if len(self.verse_buttons) < 2:
                for btn in self.verse_buttons:
                    btn.setChecked(True)
                return
                
            for i, btn in enumerate(self.verse_buttons):
                btn.setChecked(i == 0 or i == len(self.verse_buttons) - 1)
        else:
            for btn in self.verse_buttons:
                btn.setChecked(True)

    def select_first_second_and_last(self, checked):
        """Toggle verse buttons to select first, second, and last verse if checked, else all verses."""
        if checked:
            self.first_last_btn.blockSignals(True)
            self.first_last_btn.setChecked(False)
            self.first_last_btn.blockSignals(False)
            self.settings.setValue("first_last_toggle", False)
            
            self.all_verses_btn.blockSignals(True)
            self.all_verses_btn.setChecked(False)
            self.all_verses_btn.blockSignals(False)
            self.settings.setValue("all_verses_toggle", False)
        elif not self.first_last_btn.isChecked():
            # If unchecked and the other isn't checked, default to All
            self.all_verses_btn.blockSignals(True)
            self.all_verses_btn.setChecked(True)
            self.all_verses_btn.blockSignals(False)
            self.settings.setValue("all_verses_toggle", True)
            
        self.settings.setValue("first_second_last_toggle", checked)

        if not self.verse_buttons:
            return
            
        if checked:
            if len(self.verse_buttons) < 3:
                for btn in self.verse_buttons:
                    btn.setChecked(True)
                return
                
            for i, btn in enumerate(self.verse_buttons):
                btn.setChecked(i == 0 or i == 1 or i == len(self.verse_buttons) - 1)
        else:
            for btn in self.verse_buttons:
                btn.setChecked(True)

    def select_all_verses(self, checked):
        """Toggle verse buttons to select all verses."""
        if checked:
            self.first_last_btn.blockSignals(True)
            self.first_last_btn.setChecked(False)
            self.first_last_btn.blockSignals(False)
            self.settings.setValue("first_last_toggle", False)
            
            self.first_second_last_btn.blockSignals(True)
            self.first_second_last_btn.setChecked(False)
            self.first_second_last_btn.blockSignals(False)
            self.settings.setValue("first_second_last_toggle", False)
            
            self.settings.setValue("all_verses_toggle", True)
        else:
            # Prevent unchecking "All" if no others are checked
            if not self.first_last_btn.isChecked() and not self.first_second_last_btn.isChecked():
                self.all_verses_btn.blockSignals(True)
                self.all_verses_btn.setChecked(True)
                self.all_verses_btn.blockSignals(False)
                return

        if not self.verse_buttons:
            return
            
        if checked:
            for btn in self.verse_buttons:
                btn.setChecked(True)

    def _on_song_selection_changed(self):
        """Reset the verse panel when the user selects a different song."""
        selected = self.file_list.selectedItems()
        if not selected:
            self._reset_verse_panel()
            return

        song_path = selected[0].data(Qt.ItemDataRole.UserRole)
        if song_path != self.parsed_song_path:
            self._reset_verse_panel()

    def _get_selected_slide_indices(self):
        """
        Build a sorted list of 1-based slide indices from the checked verse buttons.
        Returns None if no verse scan has been performed (meaning embed all slides).
        """
        if not self.parsed_song_structure or not self.verse_buttons:
            return None

        slide_indices = []
        for i, btn in enumerate(self.verse_buttons):
            if btn.isChecked():
                entry = self.parsed_song_structure[i]
                slide_indices.extend(entry["slides"])
                slide_indices.extend(entry["chorus_slides"])

        slide_indices.sort()
        return slide_indices if slide_indices else None

    # ── Embed Logic ──────────────────────────────────────────────────────────

    def on_double_click_embed(self, item):
        """Embed song on double-click if conditions are met."""
        if item.data(Qt.ItemDataRole.UserRole) and self.embed_btn.isEnabled():
            self.embed_selected()

    def _next_song_section_name(self, sp):
        """Return the next available 'Song N' name based on existing sections."""
        max_num = 0
        for i in range(1, sp.Count + 1):
            name = sp.Name(i).strip()
            m = re.match(r'^Song\s*#?(\d+)$', name, re.IGNORECASE)
            if m:
                max_num = max(max_num, int(m.group(1)))
        return f"Song {max_num + 1}"

    def embed_selected(self):
        """Embed the selected song into the chosen target section."""
        selected = self.file_list.selectedItems()
        if not selected:
            return

        song_path = selected[0].data(Qt.ItemDataRole.UserRole)
        song_filename = selected[0].text()
        if not song_path:
            return

        auto_append = self.auto_append_cb.isChecked()
        target_data = self.target_combo.currentData()
        
        if not auto_append and not target_data:
            self.set_status("No target section selected", error=True)
            return

        section_idx = -1
        section_name = ""
        if target_data:
            section_idx, section_name = target_data

        if not os.path.exists(song_path):
            self.set_status(f"Song file not found: {song_path}", error=True)
            self.scan_folder()
            return

        # ── Determine Verse Selection ──
        # If a specific toggle is ON but the song hasn't been scanned, scan it now
        if (self.first_last_btn.isChecked() or self.first_second_last_btn.isChecked() or self.all_verses_btn.isChecked()) and song_path != self.parsed_song_path:
            self.scan_song_verses()

        slide_indices = self._get_selected_slide_indices()
        if slide_indices is not None and song_path == self.parsed_song_path:
            selected_verses = [
                self.parsed_song_structure[i]["verse"]
                for i, btn in enumerate(self.verse_buttons) if btn.isChecked()
            ]
            verse_text = ", ".join(str(v) for v in selected_verses)
            verse_info = f"\nVerses: {verse_text} ({len(slide_indices)} slides)"
        else:
            slide_indices = None
            verse_info = ""

        # ── Auto Append mode ──
        if auto_append:
            pres = self._get_master_pres()
            if not pres:
                self.set_status("Master presentation not accessible", error=True)
                return
            sp = pres.SectionProperties
            new_section_name = self._next_song_section_name(sp)

            if self.require_confirm_cb.isChecked():
                reply = QMessageBox.question(
                    self,
                    "Confirm Append",
                    f"Create new section '{new_section_name}' and embed '{song_filename}'?"
                    f"{verse_info}",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.No:
                    return

            try:
                # Insert the new section right after the currently selected section, or at the end if none selected
                if section_idx >= 0:
                    insert_at = section_idx + 1
                else:
                    insert_at = sp.Count + 1
                sp.AddSection(insert_at, new_section_name)
                print(f"Auto-appended section '{new_section_name}' at index {insert_at}")
            except Exception as e:
                self.set_status(f"Failed to create section: {e}", error=True)
                return

            self.set_status(f"Embedding '{song_filename}' into {new_section_name}…")
            QApplication.processEvents()

            insert_blank = self.insert_blank_cb.isChecked()
            success, inserted = self.do_embed(
                song_path, song_filename, new_section_name,
                replace=False, insert_blank=insert_blank, slide_indices=slide_indices,
                is_auto_append=True
            )

            if success:
                self.slot_assignments[new_section_name] = (song_filename, song_path)
                self.set_status(
                    f"✓ Appended '{song_filename}' → {new_section_name}  ({inserted} slides)",
                    success=True)

                self.read_sections()
                self.rebuild_slots_display()
                self.update_target_combo()

                # Select the newly created section in the target combo
                for i in range(self.target_combo.count()):
                    data = self.target_combo.itemData(i)
                    if data and data[1] == new_section_name:
                        self.target_combo.setCurrentIndex(i)
                        break

                self.search_box.clear()
                self.search_box.setFocus()
                self.update_embed_state()
            return

        # ── Normal embed mode ──
        is_replace = self.replace_cb.isChecked()
        mode_text = "Replace existing slides" if is_replace else "Append to existing slides"

        if self.require_confirm_cb.isChecked():
            reply = QMessageBox.question(
                self,
                "Confirm Embed",
                f"Are you sure you want to embed '{song_filename}' into '{section_name}'?\n\nMode: {mode_text}{verse_info}",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.No:
                return

        self.set_status(f"Embedding '{song_filename}' into {section_name}…")
        QApplication.processEvents()

        insert_blank = self.insert_blank_cb.isChecked()
        success, inserted = self.do_embed(song_path, song_filename, section_name, is_replace, insert_blank, slide_indices=slide_indices)

        if success:
            self.slot_assignments[section_name] = (song_filename, song_path)
            self.set_status(
                f"✓ Embedded '{song_filename}' → {section_name}  ({inserted} slides)",
                success=True)

            # Re-read sections and rebuild UI
            self.read_sections()
            self.rebuild_slots_display()
            self.update_target_combo()


            # Clear search for next song
            self.search_box.clear()
            self.search_box.setFocus()

            self.update_embed_state()

    def do_embed(self, song_path, song_filename, section_name, replace, insert_blank, keep_source_formatting=False, slide_indices=None, is_auto_append=False):
        """
        Core embed logic: insert slides from song_path into the named section
        of the master presentation.

        If slide_indices is provided (a sorted list of 1-based slide numbers),
        only those slides from the song file will be kept. All slides are inserted
        first, then unwanted slides are removed.

        Returns (success: bool, inserted_count: int).
        """
        pres = self._get_master_pres()
        if not pres:
            self.set_status("Master presentation not accessible", error=True)
            return False, 0

        try:
            sp = pres.SectionProperties

            # Find section by name (indices may shift between operations)
            target_idx = self._find_section_index(sp, section_name)
            if target_idx is None:
                self.set_status(f"Section '{section_name}' not found in master", error=True)
                return False, 0

            existing_count = sp.SlidesCount(target_idx)

            # Normalise path for COM
            normalised = os.path.abspath(os.path.normpath(song_path))

            total_before = pres.Slides.Count

            # Determine insertion index
            if existing_count > 0:
                if replace:
                    # Insert AFTER the first slide of the target section to ensure 
                    # the new slides are placed inside the target section.
                    first_slide_idx = sp.FirstSlide(target_idx)
                    insert_idx = first_slide_idx
                else:
                    # Append: Insert AFTER the last slide of the target section.
                    first_slide_idx = sp.FirstSlide(target_idx)
                    insert_idx = first_slide_idx + existing_count - 1
            else:
                # Target section is empty. Insert after the previous sections.
                insert_idx = 0
                for i in range(1, target_idx):
                    insert_idx += sp.SlidesCount(i)

            # ── Insert Front Blank Slide if Empty & Auto Append ──
            front_blank_added = False
            if is_auto_append and total_before == 0 and insert_blank:
                new_slide = pres.Slides.Add(1, 12)  # 12 is ppLayoutBlank
                new_slide.FollowMasterBackground = False
                new_slide.Background.Fill.Solid()
                new_slide.Background.Fill.ForeColor.RGB = 0  # Black
                insert_idx += 1
                total_before += 1
                front_blank_added = True

            # Insert slides
            blank_added = False
            try:
                insert_from = insert_idx

                ppt_app = self._get_ppt_app()
                if keep_source_formatting and ppt_app and pres.Windows.Count > 0:
                    song_pres = ppt_app.Presentations.Open(normalised, ReadOnly=False, WithWindow=False)
                    try:
                        sc = song_pres.Slides.Count
                        if sc > 0:
                            song_pres.Slides.Range().Copy()
                            pres.Windows(1).Activate()
                            if insert_from > 0:
                                pres.Slides(insert_from).Select()
                                ppt_app.CommandBars.ExecuteMso("PasteSourceFormatting")
                            else:
                                dummy = pres.Slides.Add(1, 12)
                                dummy.Select()
                                ppt_app.CommandBars.ExecuteMso("PasteSourceFormatting")
                                dummy.Delete()
                                
                            import time
                            start_wait = time.time()
                            while time.time() - start_wait < 5.0:
                                if pres.Slides.Count > total_before:
                                    break
                                time.sleep(0.1)
                    finally:
                        song_pres.Close()
                else:
                    pres.Slides.InsertFromFile(normalised, insert_from, 1, 999)
            except Exception:
                # Fallback: open song to count slides, then insert with exact range
                ppt_app = self._get_ppt_app()
                if not ppt_app:
                    raise
                song_pres = ppt_app.Presentations.Open(normalised, ReadOnly=True)
                sc = song_pres.Slides.Count
                song_pres.Close()
                if sc == 0:
                    self.set_status("Song file contains no slides", error=True)
                    return False, 0
                
                insert_from = insert_idx
                pres.Slides.InsertFromFile(normalised, insert_from, 1, sc)

            current_count = pres.Slides.Count
            song_slides_added = current_count - total_before

            if song_slides_added == 0:
                self.set_status("Song file contained no slides", error=True)
                return False, 0

            # Add blank slide at the end of the newly inserted song slides
            if insert_blank:
                blank_slide_idx = insert_from + song_slides_added + 1
                new_slide = pres.Slides.Add(blank_slide_idx, 12)  # 12 is ppLayoutBlank
                new_slide.FollowMasterBackground = False
                new_slide.Background.Fill.Solid()
                new_slide.Background.Fill.ForeColor.RGB = 0  # Black
                blank_added = True

            total_after = pres.Slides.Count
            inserted_count = total_after - total_before

            # ── Selective verse embedding: remove unwanted slides ──
            # If slide_indices is provided, delete any newly inserted slides
            # whose original song slide number is NOT in the selection.
            if slide_indices is not None and inserted_count > 0:
                # The newly inserted song slides start at (insert_from + 1) in the
                # master deck. The blank slide (if any) was inserted at the end
                # and is NOT a song slide — it should be preserved.
                song_slides_start = insert_from + 1
                # Total song slides inserted (excluding the blank slide)
                song_slides_count = inserted_count - (1 if blank_added else 0)

                # Build set of original 1-based song slide numbers to keep
                keep_set = set(slide_indices)

                # Delete in reverse order to preserve indices
                slides_to_delete = []
                for offset in range(song_slides_count):
                    original_song_num = offset + 1  # 1-based position in the source song
                    if original_song_num not in keep_set:
                        master_slide_idx = song_slides_start + offset
                        slides_to_delete.append(master_slide_idx)

                for master_idx in reversed(slides_to_delete):
                    pres.Slides(master_idx).Delete()

                # Recalculate inserted count after deletions
                total_after = pres.Slides.Count
                inserted_count = total_after - total_before

                if inserted_count == 0 or (inserted_count == 1 and blank_added):
                    self.set_status("No slides remain after verse filtering", error=True)
                    # Clean up the blank slide if nothing else was inserted
                    if blank_added and inserted_count == 1:
                        pres.Slides(insert_from + 1).Delete()
                    if front_blank_added:
                        pres.Slides(1).Delete()
                    return False, 0

            # Delete old slides and fix up section assignment
            if existing_count > 0:
                if replace:
                    # Delete the original first slide
                    pres.Slides(first_slide_idx).Delete()
                    # The remaining old slides are now shifted down.
                    # E.g. we inserted 3 slides after idx 10. The remaining old slides are at idx 10 + 3 = 13.
                    for _ in range(existing_count - 1):
                        pres.Slides(first_slide_idx + inserted_count).Delete()
                else:
                    # Appending: no slides are deleted. PowerPoint natively associates the new slides
                    # with the section since they were inserted right after the section's last slide.
                    pass
            else:
                # If the section was completely empty, slides were inserted into the previous section.
                # Move them to the target section in reverse order to preserve their sequence.
                # Unless we added a front blank slide to an empty presentation, in which case they are already in the correct section natively.
                if not front_blank_added:
                    for i in range(inserted_count, 0, -1):
                        pres.Slides(insert_idx + i).MoveToSectionStart(target_idx)

            final_inserted = inserted_count + (1 if front_blank_added else 0)
            print(f"Embedded {final_inserted} slides from '{song_filename}' into '{section_name}'. "
                  f"Removed {existing_count} old slide(s).")
            return True, final_inserted

        except Exception as e:
            self.set_status(f"Embed failed: {e}", error=True)
            print(f"ERROR during embed: {e}")
            return False, 0

    # ── Section Management ───────────────────────────────────────────────────

    def add_section(self):
        pres = self._get_master_pres()
        if not pres:
            self.set_status("Could not access master presentation", error=True)
            return
        
        name, ok = QInputDialog.getText(self, "Add Section", "Section Name:")
        if not ok or not name.strip():
            return
            
        target_data = self.target_combo.currentData()
        target_idx = target_data[0] if target_data else pres.SectionProperties.Count
        
        try:
            insert_index = min(pres.SectionProperties.Count + 1, target_idx + 1) if target_data else pres.SectionProperties.Count + 1
            pres.SectionProperties.AddSection(insert_index, name.strip())
            self.set_status(f"Added section '{name}'", success=True)
            self.read_sections()
            self.rebuild_slots_display()
            self.update_target_combo()
        except Exception as e:
            self.set_status(f"Error adding section: {e}", error=True)

    def remove_section(self):
        target_data = self.target_combo.currentData()
        if not target_data:
            self.set_status("No section selected to remove", error=True)
            return
        section_idx, section_name = target_data
        
        reply = QMessageBox.question(
            self,
            "Confirm Remove",
            f"Are you sure you want to remove section '{section_name}' AND ALL ITS SLIDES?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.No:
            return
            
        pres = self._get_master_pres()
        if not pres:
            return
            
        try:
            sp = pres.SectionProperties
            actual_idx = self._find_section_index(sp, section_name)
            if actual_idx:
                sp.Delete(actual_idx, True) # True means delete slides
                self.set_status(f"Removed section '{section_name}'", success=True)
                self.read_sections()
                self.rebuild_slots_display()
                self.update_target_combo()
        except Exception as e:
            self.set_status(f"Error removing section: {e}", error=True)

    def empty_selected_section(self):
        target_data = self.target_combo.currentData()
        if not target_data:
            self.set_status("No section selected to empty", error=True)
            return
        section_idx, section_name = target_data
        self.clear_slot(section_name, section_idx)

    def move_section_up(self):
        self._move_section(-1)

    def move_section_down(self):
        self._move_section(1)
        
    def _move_section(self, direction):
        target_data = self.target_combo.currentData()
        if not target_data:
            self.set_status("No section selected to move", error=True)
            return
        section_idx, section_name = target_data
        pres = self._get_master_pres()
        if not pres:
            return
            
        try:
            sp = pres.SectionProperties
            actual_idx = self._find_section_index(sp, section_name)
            if not actual_idx:
                return
            
            new_idx = actual_idx + direction
            if new_idx < 1 or new_idx > sp.Count:
                return # Can't move further
                
            sp.Move(actual_idx, new_idx)
            self.read_sections()
            self.rebuild_slots_display()
            self.update_target_combo()
            self.select_slot_by_index(new_idx)
            
        except Exception as e:
            self.set_status(f"Error moving section: {e}", error=True)

    def handle_one_off_drop(self, file_path):
        target_data = self.target_combo.currentData()
        if not target_data:
            self.set_status("No target section selected", error=True)
            return
            
        section_idx, section_name = target_data
        filename = os.path.basename(file_path)
        self._process_one_off_injection(file_path, filename, section_name)

    def inject_one_off(self):
        target_data = self.target_combo.currentData()
        if not target_data:
            self.set_status("No target section selected", error=True)
            return

        section_idx, section_name = target_data
        
        start_dir = self.current_folder or os.path.expanduser("~")
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select PowerPoint File", start_dir, "PowerPoint Files (*.ppt *.pptx)"
        )
        if not file_path:
            return
            
        filename = os.path.basename(file_path)
        self._process_one_off_injection(file_path, filename, section_name)
        
    def _process_one_off_injection(self, file_path, filename, section_name):
        if self.require_confirm_cb.isChecked():
            reply = QMessageBox.question(
                self,
                "Confirm Inject",
                f"Are you sure you want to CLEAR section '{section_name}' and inject '{filename}' into it?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return
                
        self.set_status(f"Injecting '{filename}' into {section_name}…")
        QApplication.processEvents()
        
        insert_blank = self.insert_blank_cb.isChecked()
        success, inserted = self.do_embed(file_path, filename, section_name, replace=True, insert_blank=insert_blank, keep_source_formatting=True)
        
        if success:
            self.slot_assignments[section_name] = (filename, file_path)
            self.set_status(
                f"✓ Injected '{filename}' → {section_name}  ({inserted} slides)",
                success=True)

            self.read_sections()
            self.rebuild_slots_display()
            self.update_target_combo()

    # ── Window Events ────────────────────────────────────────────────────────
    
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    path = url.toLocalFile().lower()
                    if path.endswith('.ppt') or path.endswith('.pptx'):
                        event.acceptProposedAction()
                        return
        event.ignore()
        
    def dropEvent(self, event):
        for url in event.mimeData().urls():
            if url.isLocalFile():
                path = url.toLocalFile()
                if path.lower().endswith('.ppt') or path.lower().endswith('.pptx'):
                    self.handle_one_off_drop(path)
                    event.acceptProposedAction()
                    return
        event.ignore()

    def closeEvent(self, event):
        print("Saving settings and closing…")
        self.settings.setValue("geometry", self.saveGeometry())
        if hasattr(self, 'split_layout'):
            self.settings.setValue("splitter_state", self.split_layout.saveState())
        if hasattr(self, 'sections_tree'):
            self.settings.setValue("tree_header_state", self.sections_tree.header().saveState())
        if hasattr(self, 'first_last_btn'):
            self.settings.setValue("first_last_toggle", self.first_last_btn.isChecked())
        if hasattr(self, 'first_second_last_btn'):
            self.settings.setValue("first_second_last_toggle", self.first_second_last_btn.isChecked())
        if hasattr(self, 'all_verses_btn'):
            self.settings.setValue("all_verses_toggle", self.all_verses_btn.isChecked())

        if HAS_WIN32COM:
            try:
                pythoncom.CoUninitialize()
                print("COM uninitialised.")
            except Exception as e:
                print(f"Error uninitialising COM: {e}")

        event.accept()


# ─────────────────────────────────────────────────────────────────────────────
# Entry Point
    # ── Update Checking ──────────────────────────────────────────────────────

    def check_for_updates_auto(self):
        """Check for updates in the background automatically on startup."""
        self.do_update_check(manual=False)

    def check_for_updates_manual(self):
        """Check for updates triggered by user button press."""
        self.check_updates_btn.setText("Checking...")
        self.check_updates_btn.setEnabled(False)
        self.do_update_check(manual=True)

    def do_update_check(self, manual):
        self.update_thread = UpdateCheckThread()
        self.update_thread.result_ready.connect(lambda success, ver, notes, html_url, dl_url: self.on_update_result(success, ver, notes, html_url, dl_url, manual))
        self.update_thread.start()

    def on_update_result(self, success, version, notes, html_url, download_url, manual):
        if manual:
            self.check_updates_btn.setText("Check for Updates")
            self.check_updates_btn.setEnabled(True)
            
        if success:
            if version:
                dialog = UpdateDialog(version, notes, html_url, download_url, self)
                dialog.exec()
            elif manual:
                QMessageBox.information(self, "Up to Date", f"You are running the latest version of Song Embed (v{APP_VERSION}).")
        else:
            if manual:
                QMessageBox.warning(self, "Update Check Failed", f"Could not check for updates:\n{version}")

# ─────────────────────────────────────────────────────────────────────────────
def main():
    if platform.system() == "Windows":
        try:
            # Prevents Windows from grouping this app under python's generic icon on the taskbar
            myappid = 'churchmedia.songembed.app.v1'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

    app = QApplication(sys.argv)
    app.setStyleSheet(DARK_STYLESHEET)
    window = SongEmbedApp()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    if sys.stdin and sys.stdin.isatty():
        print(f"Starting {APP_NAME}…")
    main()
