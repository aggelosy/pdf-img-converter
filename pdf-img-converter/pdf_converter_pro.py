import os
import sys
import re
import json
import shutil
from pathlib import Path
from datetime import datetime

from pdf2image import convert_from_path
import fitz  # PyMuPDF for page count and preview

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QListWidget, QPushButton, QLabel,
    QFileDialog, QProgressBar, QVBoxLayout, QHBoxLayout, QWidget,
    QComboBox, QSpinBox, QCheckBox, QGroupBox, QSplitter,
    QListWidgetItem, QTabWidget, QTextEdit, QScrollArea, QFrame,
    QSlider, QLineEdit, QMessageBox, QMenu, QAction, QSizePolicy,
    QGridLayout, QAbstractItemView, QToolTip
)
from PyQt5.QtCore import (
    QThread, pyqtSignal, Qt, QMimeData, QTimer, QSize, QUrl, QPoint
)
from PyQt5.QtGui import (
    QPixmap, QImage, QColor, QPalette, QFont, QFontDatabase,
    QIcon, QDrag, QDragEnterEvent, QDropEvent, QCursor, QPainter,
    QLinearGradient, QBrush
)


# ─── Dark Theme Stylesheet ────────────────────────────────────────────────────

DARK_STYLESHEET = """
/* ── Industrial / Utilitarian Theme ── */
QMainWindow, QWidget {
    background-color: #1c1c1c;
    color: #c8c8c0;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 12px;
}

QGroupBox {
    border: 1px solid #3a3a3a;
    border-radius: 0px;
    margin-top: 14px;
    padding-top: 6px;
    font-weight: 700;
    color: #f0a500;
    font-size: 10px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 8px;
    padding: 0 6px;
    background: #1c1c1c;
}

QPushButton {
    background-color: #252525;
    color: #c8c8c0;
    border: 1px solid #3a3a3a;
    border-radius: 0px;
    padding: 7px 14px;
    font-family: 'Consolas', 'Courier New', monospace;
    font-weight: 700;
    font-size: 11px;
    letter-spacing: 0.06em;
    min-height: 28px;
}
QPushButton:hover {
    background-color: #2e2e2e;
    border-color: #f0a500;
    color: #f0a500;
}
QPushButton:pressed {
    background-color: #1a1a1a;
    border-color: #c07800;
}
QPushButton:disabled {
    color: #444;
    border-color: #2a2a2a;
    background-color: #1e1e1e;
}

QPushButton#primaryBtn {
    background-color: #f0a500;
    color: #111;
    border: none;
    font-weight: 700;
    font-size: 12px;
    letter-spacing: 0.1em;
}
QPushButton#primaryBtn:hover {
    background-color: #ffc840;
    color: #000;
}
QPushButton#primaryBtn:pressed {
    background-color: #c07800;
}
QPushButton#primaryBtn:disabled {
    background-color: #2a2a2a;
    color: #555;
}

QPushButton#dangerBtn {
    background-color: #2a1010;
    color: #e05050;
    border-color: #5a2020;
    letter-spacing: 0.06em;
}
QPushButton#dangerBtn:hover {
    background-color: #3a1515;
    border-color: #e05050;
    color: #ff7070;
}

QListWidget {
    background-color: #161616;
    border: 1px solid #3a3a3a;
    border-radius: 0px;
    padding: 2px;
    outline: none;
    font-family: 'Consolas', 'Courier New', monospace;
}
QListWidget::item {
    padding: 7px 10px;
    border-bottom: 1px solid #222;
    color: #b0b0a8;
}
QListWidget::item:selected {
    background-color: #2a2000;
    color: #f0a500;
    border-left: 2px solid #f0a500;
}
QListWidget::item:hover:!selected {
    background-color: #1e1e1e;
    color: #d0d0c8;
}

QComboBox {
    background-color: #252525;
    border: 1px solid #3a3a3a;
    border-radius: 0px;
    padding: 5px 10px;
    color: #c8c8c0;
    font-family: 'Consolas', monospace;
    min-height: 26px;
}
QComboBox:hover { border-color: #f0a500; }
QComboBox::drop-down { border: none; width: 22px; }
QComboBox::down-arrow {
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 5px solid #888;
    margin-right: 6px;
}
QComboBox QAbstractItemView {
    background-color: #252525;
    border: 1px solid #3a3a3a;
    selection-background-color: #2a2000;
    selection-color: #f0a500;
    color: #c8c8c0;
    font-family: 'Consolas', monospace;
}

QSpinBox {
    background-color: #252525;
    border: 1px solid #3a3a3a;
    border-radius: 0px;
    padding: 5px 8px;
    color: #c8c8c0;
    font-family: 'Consolas', monospace;
    min-height: 26px;
}
QSpinBox:hover { border-color: #f0a500; }
QSpinBox::up-button, QSpinBox::down-button {
    background-color: #2e2e2e;
    border: none;
    width: 18px;
}
QSpinBox::up-button:hover, QSpinBox::down-button:hover {
    background-color: #3a3a3a;
}

QLineEdit {
    background-color: #252525;
    border: 1px solid #3a3a3a;
    border-radius: 0px;
    padding: 5px 8px;
    color: #c8c8c0;
    font-family: 'Consolas', monospace;
    min-height: 26px;
}
QLineEdit:focus { border-color: #f0a500; }

QProgressBar {
    background-color: #252525;
    border: 1px solid #3a3a3a;
    border-radius: 0px;
    height: 10px;
    text-align: center;
    color: #888;
    font-size: 9px;
    font-family: 'Consolas', monospace;
}
QProgressBar::chunk {
    background-color: #f0a500;
    border-radius: 0px;
}

QTabWidget::pane {
    border: 1px solid #3a3a3a;
    border-radius: 0px;
    background-color: #161616;
}
QTabBar::tab {
    background-color: #1c1c1c;
    color: #555;
    padding: 7px 18px;
    border: 1px solid #2a2a2a;
    border-bottom: none;
    font-family: 'Consolas', monospace;
    font-size: 10px;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    font-weight: 700;
}
QTabBar::tab:selected {
    color: #f0a500;
    background-color: #161616;
    border-top: 2px solid #f0a500;
    border-color: #3a3a3a;
    border-bottom: none;
}
QTabBar::tab:hover:!selected { color: #c8c8c0; background-color: #222; }

QTextEdit {
    background-color: #111;
    border: 1px solid #3a3a3a;
    border-radius: 0px;
    color: #90a090;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 11px;
    padding: 6px;
}

QCheckBox {
    spacing: 8px;
    color: #c8c8c0;
    font-family: 'Consolas', monospace;
}
QCheckBox::indicator {
    width: 14px;
    height: 14px;
    border: 1px solid #3a3a3a;
    border-radius: 0px;
    background-color: #252525;
}
QCheckBox::indicator:checked {
    background-color: #f0a500;
    border-color: #f0a500;
}

QScrollBar:vertical {
    background: #1c1c1c;
    width: 8px;
    border-left: 1px solid #2a2a2a;
}
QScrollBar::handle:vertical {
    background: #3a3a3a;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover { background: #f0a500; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }

QToolTip {
    background-color: #252525;
    color: #f0a500;
    border: 1px solid #f0a500;
    border-radius: 0px;
    padding: 5px 8px;
    font-family: 'Consolas', monospace;
}

QLabel#headerLabel {
    font-size: 18px;
    font-weight: 700;
    color: #f0a500;
    font-family: 'Consolas', monospace;
    letter-spacing: 0.06em;
}
QLabel#subLabel {
    font-size: 10px;
    color: #555;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    font-family: 'Consolas', monospace;
}
QLabel#statusLabel {
    color: #707068;
    font-size: 11px;
    font-family: 'Consolas', monospace;
}
QLabel#statValue {
    font-size: 20px;
    font-weight: 700;
    color: #f0a500;
    font-family: 'Consolas', monospace;
}
QLabel#statDesc {
    font-size: 9px;
    color: #444;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    font-family: 'Consolas', monospace;
}

QFrame#previewFrame {
    background-color: #111;
    border: 1px solid #3a3a3a;
    border-radius: 0px;
}

QMenu {
    background: #252525;
    border: 1px solid #3a3a3a;
    border-radius: 0px;
    font-family: 'Consolas', monospace;
}
QMenu::item { padding: 7px 22px; color: #c8c8c0; }
QMenu::item:selected { background: #2a2000; color: #f0a500; }
"""


# ─── Drop Zone Widget ─────────────────────────────────────────────────────────

class DropZoneList(QListWidget):
    """Drag-and-drop aware list widget for PDF files"""
    files_dropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(u.toLocalFile().lower().endswith('.pdf') for u in urls):
                event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            paths = [
                u.toLocalFile() for u in event.mimeData().urls()
                if u.toLocalFile().lower().endswith('.pdf')
            ]
            if paths:
                self.files_dropped.emit(paths)
        else:
            super().dropEvent(event)


# ─── Image Preview Widget ──────────────────────────────────────────────────────

class ImagePreviewWidget(QLabel):
    """Displays a preview thumbnail of a PDF page"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.setObjectName("previewFrame")
        self.setMinimumSize(200, 260)
        self.setStyleSheet("""
            QLabel#previewFrame {
                background-color: #111;
                border: 1px solid #3a3a3a;
                border-radius: 0px;
                color: #555;
                font-size: 11px;
                font-family: 'Consolas', monospace;
            }
        """)
        self.setText("Drop a PDF to\nsee preview")

    def set_preview(self, pdf_path: str, page_num: int = 0):
        try:
            doc = fitz.open(pdf_path)
            page = doc[page_num]
            mat = fitz.Matrix(0.5, 0.5)
            pix = page.get_pixmap(matrix=mat)
            img = QImage(pix.samples, pix.width, pix.height,
                         pix.stride, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(img)
            pixmap = pixmap.scaled(
                self.width() - 20, self.height() - 20,
                Qt.KeepAspectRatio, Qt.SmoothTransformation
            )
            self.setPixmap(pixmap)
            doc.close()
        except Exception as e:
            self.setText(f"Preview\nunavailable\n\n{str(e)[:40]}")


# ─── Conversion Worker ────────────────────────────────────────────────────────

class ConversionWorker(QThread):
    progress_file   = pyqtSignal(int)        # files done
    progress_page   = pyqtSignal(int, int)   # page done, total pages
    status          = pyqtSignal(str, str)   # message, level (info/ok/error)
    file_done       = pyqtSignal(str, int, float)  # path, pages, seconds
    finished        = pyqtSignal(int, int)   # total converted, total errors

    def __init__(self, pdf_paths, settings):
        super().__init__()
        self.pdf_paths = pdf_paths
        self.settings  = settings
        self._abort    = False

    def abort(self):
        self._abort = True

    def run(self):
        converted, errors = 0, 0
        fmt        = self.settings['format'].upper()
        dpi        = self.settings['dpi']
        quality    = self.settings['quality']
        page_range = self.settings.get('page_range')  # (first, last) or None
        threads    = self.settings['threads']
        prefix     = self.settings['prefix']
        custom_out = self.settings.get('output_folder')
        grayscale  = self.settings.get('grayscale', False)

        for i, pdf_path in enumerate(self.pdf_paths):
            if self._abort:
                self.status.emit("⚠ Conversion aborted by user.", "error")
                break

            pdf_path_obj = Path(pdf_path)
            pdf_name     = pdf_path_obj.stem
            out_name     = f"{prefix}{pdf_name}" if prefix else pdf_name

            if custom_out:
                output_folder = Path(custom_out) / out_name
            else:
                output_folder = pdf_path_obj.parent / out_name
            output_folder.mkdir(parents=True, exist_ok=True)

            self.status.emit(f"Converting: {pdf_name}…", "info")
            t_start = datetime.now()

            try:
                # Convert to PIL images in memory — no output_folder/output_file
                # so pdf2image never holds file handles open. We save ourselves
                # with deterministic names, which also avoids the WinError 32
                # file-locked rename failure on UNC/network paths.
                conv_kwargs = dict(dpi=dpi, thread_count=threads)
                if grayscale:
                    conv_kwargs['grayscale'] = True
                if page_range:
                    conv_kwargs['first_page'] = page_range[0]
                    conv_kwargs['last_page']  = page_range[1]

                pages = convert_from_path(pdf_path, **conv_kwargs)

                # Determine save parameters per format
                fmt_upper = fmt.upper()
                ext_map   = {
                    'PNG': 'png', 'JPEG': 'jpg', 'TIFF': 'tiff',
                    'WEBP': 'webp', 'BMP': 'bmp'
                }
                ext = ext_map.get(fmt_upper, fmt.lower())

                save_kwargs = {}
                pil_fmt     = fmt_upper
                if fmt_upper == 'JPEG':
                    save_kwargs = {'quality': quality, 'optimize': True}
                elif fmt_upper == 'PNG':
                    save_kwargs = {'optimize': True}
                elif fmt_upper == 'TIFF':
                    save_kwargs = {'compression': 'tiff_lzw'}
                elif fmt_upper == 'WEBP':
                    save_kwargs = {'quality': quality, 'method': 6}

                # Save each page with a clean, zero-padded filename
                for idx, img in enumerate(pages, start=1):
                    dest = output_folder / f"{out_name}-{idx:04d}.{ext}"
                    img.save(str(dest), format=pil_fmt, **save_kwargs)
                    img.close()   # release memory immediately

                elapsed = (datetime.now() - t_start).total_seconds()
                self.status.emit(
                    f"✓ {pdf_name}  ·  {len(pages)} pages  ·  {elapsed:.1f}s", "ok"
                )
                self.file_done.emit(str(output_folder), len(pages), elapsed)
                converted += 1

            except Exception as e:
                self.status.emit(f"✗ {pdf_name}: {str(e)}", "error")
                errors += 1

            self.progress_file.emit(i + 1)

        self.finished.emit(converted, errors)


# ─── Stats Card ───────────────────────────────────────────────────────────────

class StatCard(QWidget):
    def __init__(self, label: str, value: str = "0", parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setSpacing(2)
        layout.setContentsMargins(16, 12, 16, 12)

        self.value_lbl = QLabel(value)
        self.value_lbl.setObjectName("statValue")

        desc_lbl = QLabel(label)
        desc_lbl.setObjectName("statDesc")

        layout.addWidget(self.value_lbl)
        layout.addWidget(desc_lbl)

        self.setStyleSheet("""
            StatCard {
                background-color: #161616;
                border: 1px solid #333;
                border-radius: 0px;
            }
        """)

    def set_value(self, val: str):
        self.value_lbl.setText(val)


# ─── Main Window ──────────────────────────────────────────────────────────────

class PDFConverterPro(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self._stats = {"files": 0, "pages": 0, "seconds": 0.0}
        self._output_folder = ""
        self._setups()
        self._build_ui()
        self._apply_theme()

    # ── Setup ──────────────────────────────────────────────────────────────

    def _setups(self):
        self.setWindowTitle("PDF-TO-IMAGE  //  PRO")
        self.setMinimumSize(1020, 680)
        self.resize(1100, 720)

    def _apply_theme(self):
        self.setStyleSheet(DARK_STYLESHEET)

    # ── UI Build ───────────────────────────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Left sidebar
        sidebar = self._build_sidebar()
        root.addWidget(sidebar)

        # Main content
        main = self._build_main()
        root.addWidget(main, 1)

    def _build_sidebar(self) -> QWidget:
        sidebar = QWidget()
        sidebar.setFixedWidth(280)
        sidebar.setStyleSheet("background-color: #141414; border-right: 1px solid #2e2e2e;")
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(20, 24, 20, 20)
        layout.setSpacing(16)

        # Header
        title = QLabel("PDF//IMG")
        title.setObjectName("headerLabel")
        sub   = QLabel("// BATCH CONVERTER v2")
        sub.setObjectName("subLabel")
        layout.addWidget(title)
        layout.addWidget(sub)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("background: #2e2e2e; border: none; max-height: 1px; margin: 4px 0;")
        layout.addWidget(sep)

        # Preview
        self.preview = ImagePreviewWidget()
        layout.addWidget(self.preview)

        # Stats
        stats_grid = QWidget()
        sg = QGridLayout(stats_grid)
        sg.setSpacing(8)
        sg.setContentsMargins(0, 0, 0, 0)

        self.stat_files = StatCard("FILES")
        self.stat_pages = StatCard("PAGES")
        self.stat_time  = StatCard("TIME (s)", "0.0")
        self.stat_queue = StatCard("IN QUEUE")

        sg.addWidget(self.stat_files, 0, 0)
        sg.addWidget(self.stat_pages, 0, 1)
        sg.addWidget(self.stat_time,  1, 0)
        sg.addWidget(self.stat_queue, 1, 1)
        layout.addWidget(stats_grid)

        layout.addStretch()

        # Output folder
        out_grp = QGroupBox("Output Folder")
        out_lay = QVBoxLayout(out_grp)
        out_lay.setSpacing(6)

        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("Same folder as PDF (default)")
        self.output_edit.setReadOnly(True)
        self.output_edit.textChanged.connect(
            lambda t: setattr(self, '_output_folder', t)
        )

        browse_btn = QPushButton("Browse…")
        browse_btn.clicked.connect(self._browse_output)
        browse_btn.setFixedHeight(28)

        out_lay.addWidget(self.output_edit)
        out_lay.addWidget(browse_btn)
        layout.addWidget(out_grp)

        return sidebar

    def _build_main(self) -> QWidget:
        main = QWidget()
        layout = QVBoxLayout(main)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(16)

        # Top: controls row
        top_row = self._build_controls_row()
        layout.addWidget(top_row)

        # Splitter: file list | log+settings
        splitter = QSplitter(Qt.Horizontal)
        splitter.setStyleSheet("QSplitter::handle { background: #1e2433; width: 1px; }")

        left = self._build_file_panel()
        right = self._build_right_panel()
        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        layout.addWidget(splitter, 1)

        # Bottom: progress + convert
        bottom = self._build_bottom_row()
        layout.addWidget(bottom)

        return main

    def _build_controls_row(self) -> QWidget:
        w = QWidget()
        row = QHBoxLayout(w)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(12)

        # Format
        fmt_grp = QGroupBox("Format")
        fl = QHBoxLayout(fmt_grp)
        fl.setContentsMargins(10, 6, 10, 8)
        self.fmt_combo = QComboBox()
        self.fmt_combo.addItems(["PNG", "JPEG", "TIFF", "WebP", "BMP"])
        self.fmt_combo.setFixedWidth(100)
        self.fmt_combo.currentTextChanged.connect(self._on_format_change)
        fl.addWidget(self.fmt_combo)
        row.addWidget(fmt_grp)

        # DPI
        dpi_grp = QGroupBox("DPI")
        dl = QHBoxLayout(dpi_grp)
        dl.setContentsMargins(10, 6, 10, 8)
        self.dpi_combo = QComboBox()
        self.dpi_combo.addItems(["72", "96", "150", "200", "300", "400", "600"])
        self.dpi_combo.setCurrentText("300")
        self.dpi_combo.setFixedWidth(80)
        dl.addWidget(self.dpi_combo)
        row.addWidget(dpi_grp)

        # Quality (JPEG)
        self.quality_grp = QGroupBox("JPEG Quality")
        ql = QHBoxLayout(self.quality_grp)
        ql.setContentsMargins(10, 6, 10, 8)
        self.quality_spin = QSpinBox()
        self.quality_spin.setRange(10, 100)
        self.quality_spin.setValue(90)
        self.quality_spin.setSuffix("%")
        self.quality_spin.setFixedWidth(70)
        ql.addWidget(self.quality_spin)
        self.quality_grp.setEnabled(False)
        row.addWidget(self.quality_grp)

        # Threads
        th_grp = QGroupBox("Threads")
        tl = QHBoxLayout(th_grp)
        tl.setContentsMargins(10, 6, 10, 8)
        self.thread_spin = QSpinBox()
        self.thread_spin.setRange(1, 16)
        self.thread_spin.setValue(6)
        self.thread_spin.setFixedWidth(60)
        tl.addWidget(self.thread_spin)
        row.addWidget(th_grp)

        # Page range
        pr_grp = QGroupBox("Page Range")
        prl = QHBoxLayout(pr_grp)
        prl.setContentsMargins(10, 6, 10, 8)
        prl.setSpacing(6)
        self.range_check = QCheckBox("Enable")
        self.range_check.stateChanged.connect(self._toggle_range)
        self.page_from = QSpinBox()
        self.page_from.setRange(1, 9999)
        self.page_from.setValue(1)
        self.page_from.setFixedWidth(60)
        self.page_from.setEnabled(False)
        lbl_to = QLabel("–")
        lbl_to.setAlignment(Qt.AlignCenter)
        self.page_to = QSpinBox()
        self.page_to.setRange(1, 9999)
        self.page_to.setValue(999)
        self.page_to.setFixedWidth(60)
        self.page_to.setEnabled(False)
        prl.addWidget(self.range_check)
        prl.addWidget(self.page_from)
        prl.addWidget(lbl_to)
        prl.addWidget(self.page_to)
        row.addWidget(pr_grp)

        # Options
        opt_grp = QGroupBox("Options")
        ol = QHBoxLayout(opt_grp)
        ol.setContentsMargins(10, 6, 10, 8)
        ol.setSpacing(12)
        self.grayscale_check = QCheckBox("Grayscale")
        self.open_check      = QCheckBox("Open when done")
        ol.addWidget(self.grayscale_check)
        ol.addWidget(self.open_check)
        row.addWidget(opt_grp)

        # Prefix
        pfx_grp = QGroupBox("Filename Prefix")
        pfl = QHBoxLayout(pfx_grp)
        pfl.setContentsMargins(10, 6, 10, 8)
        self.prefix_edit = QLineEdit()
        self.prefix_edit.setPlaceholderText("(none)")
        self.prefix_edit.setFixedWidth(100)
        pfl.addWidget(self.prefix_edit)
        row.addWidget(pfx_grp)

        row.addStretch()
        return w

    def _build_file_panel(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(0, 0, 8, 0)
        layout.setSpacing(8)

        # Header
        hdr = QHBoxLayout()
        lbl = QLabel("PDF Queue")
        lbl.setStyleSheet("font-weight: 700; color: #f0a500; font-size: 10px; letter-spacing: 0.12em; font-family: 'Consolas', monospace;")
        add_btn = QPushButton("+ Add Files")
        add_btn.clicked.connect(self._load_pdfs)
        add_btn.setFixedHeight(28)
        clr_btn = QPushButton("Clear All")
        clr_btn.setObjectName("dangerBtn")
        clr_btn.clicked.connect(self._clear_list)
        clr_btn.setFixedHeight(28)
        hdr.addWidget(lbl)
        hdr.addStretch()
        hdr.addWidget(add_btn)
        hdr.addWidget(clr_btn)
        layout.addLayout(hdr)

        # Drop zone hint
        hint = QLabel("  Drag & drop PDF files here")
        hint.setStyleSheet("color: #3a3a3a; font-size: 10px; font-style: italic; font-family: 'Consolas', monospace;")
        layout.addWidget(hint)

        # File list
        self.pdf_list = DropZoneList()
        self.pdf_list.files_dropped.connect(self._add_files)
        self.pdf_list.itemSelectionChanged.connect(self._on_file_selected)
        self.pdf_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.pdf_list.customContextMenuRequested.connect(self._show_context_menu)
        layout.addWidget(self.pdf_list, 1)

        return w

    def _build_right_panel(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(8, 0, 0, 0)
        layout.setSpacing(0)

        tabs = QTabWidget()
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Conversion log will appear here…")
        tabs.addTab(self.log_box, "Log")

        # Results tab
        self.result_list = QListWidget()
        self.result_list.setStyleSheet("""
            QListWidget::item { padding: 8px; border-bottom: 1px solid #222; font-family: 'Consolas', monospace; font-size: 11px; }
        """)
        tabs.addTab(self.result_list, "Results")

        layout.addWidget(tabs, 1)
        return w

    def _build_bottom_row(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # Progress bar
        prog_row = QHBoxLayout()
        prog_lbl = QLabel("Progress")
        prog_lbl.setObjectName("statusLabel")
        prog_lbl.setFixedWidth(60)
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_lbl = QLabel("0 / 0")
        self.progress_lbl.setObjectName("statusLabel")
        self.progress_lbl.setFixedWidth(60)
        self.progress_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        prog_row.addWidget(prog_lbl)
        prog_row.addWidget(self.progress_bar, 1)
        prog_row.addWidget(self.progress_lbl)
        layout.addLayout(prog_row)

        # Buttons
        btn_row = QHBoxLayout()
        self.status_lbl = QLabel("Ready — load PDFs to begin")
        self.status_lbl.setObjectName("statusLabel")

        self.abort_btn = QPushButton("✕ Stop")
        self.abort_btn.setObjectName("dangerBtn")
        self.abort_btn.clicked.connect(self._abort_conversion)
        self.abort_btn.setFixedWidth(90)
        self.abort_btn.setVisible(False)

        self.convert_btn = QPushButton("▶  CONVERT ALL")
        self.convert_btn.setObjectName("primaryBtn")
        self.convert_btn.clicked.connect(self._convert)
        self.convert_btn.setFixedWidth(160)
        self.convert_btn.setFixedHeight(42)

        btn_row.addWidget(self.status_lbl)
        btn_row.addStretch()
        btn_row.addWidget(self.abort_btn)
        btn_row.addWidget(self.convert_btn)
        layout.addLayout(btn_row)

        return w

    # ── Slots ──────────────────────────────────────────────────────────────

    def _on_format_change(self, fmt: str):
        self.quality_grp.setEnabled(fmt == "JPEG")

    def _toggle_range(self, state: int):
        on = bool(state)
        self.page_from.setEnabled(on)
        self.page_to.setEnabled(on)

    def _browse_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.output_edit.setText(folder)
            self._output_folder = folder

    def _load_pdfs(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select PDF Files", "", "PDF Files (*.pdf)",
            options=QFileDialog.ReadOnly
        )
        self._add_files(files)

    def _add_files(self, paths: list):
        existing = {self.pdf_list.item(i).text()
                    for i in range(self.pdf_list.count())}
        added = 0
        for p in paths:
            if p not in existing and p.lower().endswith('.pdf'):
                item = QListWidgetItem(f"  {Path(p).name}")
                item.setData(Qt.UserRole, p)
                item.setToolTip(p)
                self.pdf_list.addItem(item)
                existing.add(p)
                added += 1
        if added:
            self._update_queue_stat()
            self._log(f"Added {added} file(s) to queue.", "info")

    def _clear_list(self):
        self.pdf_list.clear()
        self._update_queue_stat()

    def _on_file_selected(self):
        sel = self.pdf_list.selectedItems()
        if sel:
            path = sel[0].data(Qt.UserRole)
            if path:
                self.preview.set_preview(path)
                # Update page range max
                try:
                    doc = fitz.open(path)
                    n = len(doc)
                    doc.close()
                    self.page_to.setValue(n)
                    self.page_to.setMaximum(n)
                    self.page_from.setMaximum(n)
                except Exception:
                    pass

    def _show_context_menu(self, pos: QPoint):
        item = self.pdf_list.itemAt(pos)
        if not item:
            return
        menu = QMenu(self)
        remove_act = QAction("Remove from queue", self)
        remove_act.triggered.connect(lambda: self.pdf_list.takeItem(
            self.pdf_list.row(item)
        ))
        reveal_act = QAction("Reveal in explorer", self)
        path = item.data(Qt.UserRole)
        reveal_act.triggered.connect(lambda: self._reveal(path))
        menu.addAction(remove_act)
        menu.addAction(reveal_act)
        menu.exec_(self.pdf_list.mapToGlobal(pos))

    def _reveal(self, path: str):
        import subprocess, platform
        p = Path(path).parent
        system = platform.system()
        if system == "Windows":
            os.startfile(str(p))
        elif system == "Darwin":
            subprocess.Popen(["open", str(p)])
        else:
            subprocess.Popen(["xdg-open", str(p)])

    def _update_queue_stat(self):
        n = self.pdf_list.count()
        self.stat_queue.set_value(str(n))

    def _log(self, msg: str, level: str = "info"):
        colors = {"info": "#707068", "ok": "#80c060", "error": "#e05050"}
        color  = colors.get(level, "#94a3b8")
        ts     = datetime.now().strftime("%H:%M:%S")
        self.log_box.append(
            f'<span style="color:#334155">[{ts}]</span> '
            f'<span style="color:{color}">{msg}</span>'
        )

    # ── Conversion ────────────────────────────────────────────────────────

    def _collect_settings(self) -> dict:
        settings = {
            'format':        self.fmt_combo.currentText(),
            'dpi':           int(self.dpi_combo.currentText()),
            'quality':       self.quality_spin.value(),
            'threads':       self.thread_spin.value(),
            'grayscale':     self.grayscale_check.isChecked(),
            'prefix':        self.prefix_edit.text().strip(),
            'output_folder': self._output_folder or None,
        }
        if self.range_check.isChecked():
            settings['page_range'] = (self.page_from.value(), self.page_to.value())
        else:
            settings['page_range'] = None
        return settings

    def _convert(self):
        n = self.pdf_list.count()
        if n == 0:
            self.status_lbl.setText("No PDFs in queue.")
            return

        pdf_paths = [self.pdf_list.item(i).data(Qt.UserRole) for i in range(n)]
        settings  = self._collect_settings()

        self.convert_btn.setEnabled(False)
        self.abort_btn.setVisible(True)
        self.progress_bar.setRange(0, n)
        self.progress_bar.setValue(0)
        self.progress_lbl.setText(f"0 / {n}")
        self.result_list.clear()

        self.worker = ConversionWorker(pdf_paths, settings)
        self.worker.progress_file.connect(self._on_progress)
        self.worker.status.connect(self._on_status)
        self.worker.file_done.connect(self._on_file_done)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

        self._log(f"Starting batch: {n} file(s)  ·  {settings['format']}  ·  {settings['dpi']} DPI", "info")

    def _abort_conversion(self):
        if self.worker:
            self.worker.abort()
            self.abort_btn.setEnabled(False)

    def _on_progress(self, done: int):
        total = self.pdf_list.count()
        self.progress_bar.setValue(done)
        self.progress_lbl.setText(f"{done} / {total}")

    def _on_status(self, msg: str, level: str):
        self.status_lbl.setText(msg)
        self._log(msg, level)

    def _on_file_done(self, folder: str, pages: int, secs: float):
        self._stats['files'] += 1
        self._stats['pages'] += pages
        self._stats['seconds'] += secs
        self.stat_files.set_value(str(self._stats['files']))
        self.stat_pages.set_value(str(self._stats['pages']))
        self.stat_time.set_value(f"{self._stats['seconds']:.1f}")

        item = QListWidgetItem(
            f"✓  {Path(folder).name}  ·  {pages} pages  ·  {secs:.1f}s"
        )
        item.setForeground(QColor("#80c060"))
        item.setData(Qt.UserRole, folder)
        self.result_list.addItem(item)
        self.result_list.scrollToBottom()

        if self.open_check.isChecked():
            self._reveal(folder + "/placeholder")

    def _on_finished(self, converted: int, errors: int):
        self.convert_btn.setEnabled(True)
        self.abort_btn.setVisible(False)
        self.abort_btn.setEnabled(True)

        if errors == 0:
            msg = f"✓ All done — {converted} file(s) converted successfully"
            self._log(msg, "ok")
            self.status_lbl.setText(msg)
        else:
            msg = f"Done — {converted} converted, {errors} error(s)"
            self._log(msg, "error")
            self.status_lbl.setText(msg)

        self.worker = None

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.abort()
            self.worker.wait(2000)
        event.accept()


# ─── Entry Point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Smooth font rendering
    font = QFont("Segoe UI", 10)
    app.setFont(font)

    window = PDFConverterPro()
    window.show()
    sys.exit(app.exec_())