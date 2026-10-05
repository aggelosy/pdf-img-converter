"""Responsive desktop PDF-to-image converter.

Requirements:
    pip install -r requirements.txt

pdf2image also requires Poppler. On Windows, select the folder containing
pdftoppm.exe / pdftocairo.exe in the application's Poppler field if it is
not on PATH.
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

try:
    import ttkbootstrap as ttkb
    from ttkbootstrap.constants import BOTH, END, LEFT, RIGHT, X, Y
except ImportError:
    ttkb = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = TkinterDnD = None

ENGINE_IMPORT_ERROR = None
ENGINE_MISSING_MODULE = None
try:
    from pdf2image import convert_from_path, pdfinfo_from_path
except ImportError as exc:
    ENGINE_IMPORT_ERROR = str(exc)
    ENGINE_MISSING_MODULE = exc.name
    convert_from_path = pdfinfo_from_path = None


def engine_dependency_message():
    """Preserve the real import failure and identify the Python used to launch."""
    package = {"pdf2image": "pdf2image", "PIL": "Pillow"}.get(
        (ENGINE_MISSING_MODULE or "").split(".")[0]
    )
    heading = (f"Cannot import {package} in the Python running this app."
               if package else "The PDF conversion engine could not load.")
    command = f'"{sys.executable}" -m pip install --upgrade pdf2image Pillow'
    return (heading + "\n\nActual error: " + str(ENGINE_IMPORT_ERROR)
            + "\n\nPython: " + sys.executable
            + "\n\nIf the packages are installed in another Python, launch the app with that Python."
            + "\nOtherwise, run this in Command Prompt:\n" + command)



from contextlib import contextmanager
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import time
import uuid


INFO_TIMEOUT = 30
RENDER_TIMEOUT = 120
MAX_PAGE_PIXELS = 80_000_000
DEFAULTS = {"output": "", "poppler": "", "dpi": 200, "format": "PNG",
            "pdftocairo": False, "separate_folders": True}


class Cancelled(Exception):
    pass


def check_cancel(cancel):
    if cancel.is_set():
        raise Cancelled()


def parse_page_range(value: str, page_count: int) -> list[int]:
    if page_count < 1:
        raise ValueError("PDF has no pages")
    value = value.strip().lower()
    if not value or value == "all":
        return list(range(1, page_count + 1))
    pages = set()
    for part in value.split(","):
        match = re.fullmatch(r"\s*(\d+)\s*(?:-\s*(\d+)\s*)?", part)
        if not match:
            raise ValueError(f"Invalid page range: {part!r}; use All or 1-3, 5")
        start = int(match[1])
        end = int(match[2] or start)
        if not 1 <= start <= end <= page_count:
            raise ValueError(f"Pages must be between 1 and {page_count}, with range start <= end")
        # Bounds are checked before expanding potentially hostile ranges.
        pages.update(range(start, end + 1))
    return sorted(pages)


def config_path():
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", str(Path.home() / ".config")))
    return base / "PdfToImageConverter" / "settings.json"


def sanitized_config(raw):
    result = dict(DEFAULTS)
    if not isinstance(raw, dict):
        return result
    for key in ("output", "poppler"):
        if isinstance(raw.get(key), str):
            result[key] = raw[key]
    dpi = raw.get("dpi")
    if type(dpi) is int and 72 <= dpi <= 600:
        result["dpi"] = dpi
    if raw.get("format") in ("PNG", "JPEG"):
        result["format"] = raw["format"]
    for key in ("pdftocairo", "separate_folders"):
        if type(raw.get(key)) is bool:
            result[key] = raw[key]
    return result


def load_config(path=None):
    try:
        return sanitized_config(json.loads(Path(path or config_path()).read_text(encoding="utf-8")))
    except (OSError, ValueError):
        return dict(DEFAULTS)


def save_config(data, path=None):
    path = Path(path or config_path())
    temp = path.with_name(f".settings-{uuid.uuid4().hex}.json")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        temp.write_text(json.dumps(sanitized_config(data), indent=2), encoding="utf-8")
        os.replace(temp, path)
    except OSError:
        pass
    finally:
        try:
            temp.unlink(missing_ok=True)
        except OSError:
            pass


def source_directory(source, output=None):
    source = Path(source).resolve()
    identity = os.path.normcase(str(source)).encode("utf-8")
    digest = hashlib.sha256(identity).hexdigest()[:12]
    stem = re.sub(r"[^\w .()-]", "_", source.stem).strip(" .")[:70] or "PDF"
    return Path(output or source.parent).resolve() / f"{stem} - {digest}"


def validate_poppler(folder, use_cairo):
    for name in ("pdfinfo", "pdftocairo" if use_cairo else "pdftoppm"):
        executable = name + (".exe" if os.name == "nt" else "")
        found = (Path(folder) / executable).is_file() if folder else shutil.which(executable)
        if not found:
            raise ValueError(f"Cannot find {executable}. Select the Poppler bin folder containing pdfinfo and the selected renderer.")


def publish_image(rendered, target, overwrite):
    """Publish a completed image; never truncate an existing file during a write."""
    rendered, target = Path(rendered), Path(target)
    if overwrite:
        os.replace(rendered, target)
        return target
    counter = 1
    while True:
        candidate = target if counter == 1 else target.with_name(f"{target.stem} ({counter}){target.suffix}")
        try:
            if os.name == "nt":
                # Windows rename fails if the target exists, unlike POSIX rename.
                os.rename(rendered, candidate)
            else:
                os.link(rendered, candidate)
                rendered.unlink()
            return candidate
        except FileExistsError:
            counter += 1


@contextmanager
def scratch_folder(parent):
    # Isolated scratch files on the destination volume allow atomic publication.
    path = Path(parent).resolve() / (".pdf-render-" + uuid.uuid4().hex)
    path.mkdir()
    try:
        yield path
    finally:
        # Poppler writes flat page files here; do not recurse into arbitrary paths.
        for child in path.iterdir():
            child.unlink()
        path.rmdir()


def validate_page_size(info, dpi):
    """Reject oversized pages when pdfinfo exposes their dimensions in points."""
    dimensions = re.search(r"([\d.]+)\s*x\s*([\d.]+)\s*pts", str(info.get("Page size", "")))
    if dimensions:
        pixels = float(dimensions[1]) * float(dimensions[2]) * (dpi / 72) ** 2
        if pixels > MAX_PAGE_PIXELS:
            raise ValueError(f"Page would exceed {MAX_PAGE_PIXELS:,} pixels at {dpi} DPI. Lower the DPI.")


def convert_files(files, output, poppler, dpi, file_format, page_spec, overwrite,
                  use_pdftocairo, events, cancel, separate_folders=True):
    """Emit file/progress events and exactly one done/cancelled/error event."""
    completed, total = 0, 0
    failed, directories, jobs = [], [], []
    active = None
    poppler_path = str(poppler) if poppler else None
    try:
        if type(dpi) is not int or not 72 <= dpi <= 600 or file_format not in {"png", "jpeg"}:
            raise ValueError("Invalid DPI or output format")
        validate_poppler(poppler, use_pdftocairo)
        for source in files:
            check_cancel(cancel)
            source = Path(source)
            events.put(("file", str(source), "—", "Reading PDF...", "muted"))
            try:
                info = pdfinfo_from_path(str(source), poppler_path=poppler_path, timeout=INFO_TIMEOUT)
                check_cancel(cancel)
                count = int(info["Pages"])
                pages = parse_page_range(page_spec, count)
                jobs.append((source, pages, count))
                events.put(("file", str(source), str(count), "Queued", ""))
            except Cancelled:
                raise
            except Exception as exc:
                failed.append(f"{source}: {exc}")
                events.put(("file", str(source), "—", f"Error: {exc}", "err"))
        check_cancel(cancel)
        total = sum(len(pages) for _, pages, _ in jobs)
        if not jobs:
            raise ValueError("No files could be read.\n\n" + "\n".join(failed))
        started = time.monotonic()
        for source, pages, count in jobs:
            check_cancel(cancel)
            active = (source, count)
            saved_for_file = 0
            try:
                source_folder = source_directory(source, output)
                destination = source_folder if separate_folders else source_folder.parent
                # The source identifier keeps same-name PDFs distinct in shared output.
                filename_prefix = "" if separate_folders else f"{source_folder.name} - "
                destination.mkdir(parents=True, exist_ok=True)
                digits = max(3, len(str(count)))
                events.put(("file", str(source), str(count), "Converting...", ""))
                with scratch_folder(destination) as scratch:
                    for page in pages:
                        check_cancel(cancel)
                        # Read each page's own dimensions, including mixed-size PDFs.
                        info = pdfinfo_from_path(str(source), poppler_path=poppler_path,
                                                 first_page=page, last_page=page, timeout=INFO_TIMEOUT)
                        check_cancel(cancel)
                        # pdfinfo may use "Page N size" when a range is specified.
                        page_info = dict(info)
                        for key, value in info.items():
                            if re.fullmatch(r"Page\s+\d+\s+size", key):
                                page_info["Page size"] = value
                        validate_page_size(page_info, dpi)
                        rendered = convert_from_path(
                            str(source), dpi=dpi, first_page=page, last_page=page,
                            fmt=file_format, output_folder=str(scratch), paths_only=True,
                            output_file="page", poppler_path=poppler_path,
                            use_pdftocairo=use_pdftocairo, thread_count=1,
                            jpegopt={"quality": 92} if file_format == "jpeg" else None,
                            timeout=RENDER_TIMEOUT,
                        )
                        check_cancel(cancel)
                        if len(rendered) != 1:
                            raise RuntimeError(f"Expected one image for page {page}, got {len(rendered)}")
                        rendered_path = Path(rendered[0]).resolve()
                        if rendered_path.parent != scratch or not rendered_path.is_file() or rendered_path.stat().st_size == 0:
                            raise RuntimeError("Renderer did not return a valid output file")
                        suffix = ".png" if file_format == "png" else ".jpg"
                        target = destination / f"{filename_prefix}page {page:0{digits}d}{suffix}"
                        publish_image(rendered_path, target, overwrite)
                        if str(destination) not in directories:
                            directories.append(str(destination))
                        completed += 1
                        saved_for_file += 1
                        elapsed = time.monotonic() - started
                        remaining = (total - completed) * elapsed / completed
                        events.put(("progress", completed, total, source.name, page, f"~{int(remaining)}s remaining"))
                events.put(("file", str(source), str(count), f"Completed ({saved_for_file} images)", "ok"))
            except Cancelled:
                raise
            except Exception as exc:
                failed.append(f"{source} ({saved_for_file} images saved): {exc}")
                events.put(("file", str(source), str(count), f"Error after {saved_for_file} images: {exc}", "err"))
            active = None
        check_cancel(cancel)
        events.put(("done", completed, total, failed, directories))
    except Cancelled:
        if active:
            events.put(("file", str(active[0]), str(active[1]), "Cancelled", "muted"))
        events.put(("cancelled", completed, directories))
    except Exception as exc:
        events.put(("error", str(exc)))


if ENGINE_IMPORT_ERROR is not None:
    convert_files = None

APP_TITLE = "PDF to Image Converter 2.1"



# A muted, Notion-like palette: near-white background, soft gray captions,
# hairline dividers, and near-black text rather than pure black.
INK = "#37352F"
MUTED = "#9B9A97"
CAPTION = "#787774"
DIVIDER = "#E9E9E7"
CANVAS = "#FFFFFF"
SURFACE = "#F7F7F5"
FONT_FAMILY = "Segoe UI"














class PdfConverterApp(ttkb.Window if ttkb else tk.Tk, TkinterDnD.DnDWrapper if TkinterDnD else object):
    def __init__(self) -> None:
        if ttkb:
            super().__init__(themename="flatly")
        else:
            super().__init__()
        self.title(APP_TITLE)
        self.geometry("1000x900")
        self.minsize(960, 880)
        self.configure(background=CANVAS)

        self.config_data = load_config()
        self.running = False
        self.closing = False
        self.output_directories = []
        self.last_report = "No conversion has run yet."
        self.dnd_ready = False
        if TkinterDnD is not None:
            try:
                # Load tkdnd into ttkbootstrap's existing root; keep one Tk interpreter.
                self.TkdndVersion = TkinterDnD._require(self)
                self.dnd_ready = True
            except (RuntimeError, tk.TclError):
                pass

        self.pdf_files: list[Path] = []
        self.events: queue.Queue[tuple] = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None

        self.output_var = tk.StringVar(value=self.config_data.get("output", ""))
        self.poppler_var = tk.StringVar(value=self.config_data.get("poppler", ""))
        self.dpi_var = tk.IntVar(value=self.config_data.get("dpi", 200))
        self.format_var = tk.StringVar(value=self.config_data.get("format", "PNG"))
        self.pages_var = tk.StringVar(value="All")
        self.overwrite_var = tk.BooleanVar(value=False)
        self.separate_folders_var = tk.BooleanVar(value=self.config_data.get("separate_folders", True))
        self.hq_var = tk.BooleanVar(value=self.config_data.get("pdftocairo", False))
        self.status_var = tk.StringVar(value="Select one or more PDF files to begin.")
        self.progress_var = tk.DoubleVar(value=0)

        if ttkb is None:
            self.after(150, self._show_missing_ttkbootstrap)
            return

        self._apply_theme()
        self._build_ui()
        self.poll_id = self.after(100, self._poll_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        if convert_files is None:
            self.convert_button.configure(state="disabled")
            self.after(150, self._show_missing_package)

    # ------------------------------------------------------------------ #
    # Theming and layout
    # ------------------------------------------------------------------ #

    def _apply_theme(self) -> None:
        style = self.style  # ttkbootstrap.Window exposes the active Style
        style.configure(".", font=(FONT_FAMILY, 10), foreground=INK)
        style.configure("TFrame", background=CANVAS)
        style.configure("Surface.TFrame", background=SURFACE)
        style.configure("TLabel", background=CANVAS, foreground=INK)
        style.configure(
            "Caption.TLabel",
            background=CANVAS,
            foreground=CAPTION,
            font=(FONT_FAMILY, 9, "bold"),
        )
        style.configure(
            "Hint.TLabel",
            background=CANVAS,
            foreground=MUTED,
            font=(FONT_FAMILY, 9),
        )
        style.configure(
            "Title.TLabel",
            background=CANVAS,
            foreground=INK,
            font=(FONT_FAMILY, 20, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background=CANVAS,
            foreground=MUTED,
            font=(FONT_FAMILY, 10),
        )
        style.configure("TCheckbutton", background=CANVAS, foreground=INK)
        style.configure("Treeview", rowheight=26, font=(FONT_FAMILY, 10), borderwidth=0)
        style.configure("Treeview.Heading", font=(FONT_FAMILY, 9, "bold"))
        style.configure("TSeparator", background=DIVIDER)

    def _section(self, parent, title: str):
        """A Notion-style section: small caps caption, hairline rule, content below."""
        wrapper = ttkb.Frame(parent, style="TFrame")
        wrapper.pack(fill=X, pady=(22, 8))
        ttkb.Label(wrapper, text=title.upper(), style="Caption.TLabel").pack(anchor="w")
        ttkb.Separator(wrapper, orient="horizontal").pack(fill=X, pady=(6, 0))
        content = ttkb.Frame(parent, style="TFrame")
        content.pack(fill=X)
        return content

    def _build_ui(self) -> None:
        outer = ttkb.Frame(self, padding=(28, 24, 28, 20), style="TFrame")
        outer.pack(fill=BOTH, expand=True)

        header = ttkb.Frame(outer, style="TFrame")
        header.pack(fill=X)
        ttkb.Label(header, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
        ttkb.Label(
            header,
            text="Convert PDF pages to PNG or JPEG images, in bulk.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        # -- Files -------------------------------------------------------
        files_content = self._section(outer, "PDF files")

        tree_frame = ttkb.Frame(files_content, style="TFrame")
        tree_frame.pack(fill=BOTH, expand=True)

        self.file_tree = ttkb.Treeview(
            tree_frame,
            columns=("pages", "status"),
            show="tree headings",
            selectmode="extended",
            height=7,
            bootstyle="light",
        )
        self.file_tree.heading("#0", text="File")
        self.file_tree.heading("pages", text="Pages")
        self.file_tree.heading("status", text="Status")
        self.file_tree.column("#0", width=500, minwidth=250)
        self.file_tree.column("pages", width=70, anchor="center", stretch=False)
        self.file_tree.column("status", width=190)
        self.file_tree.pack(side=LEFT, fill=BOTH, expand=True)
        self.file_tree.tag_configure("ok", foreground="#2E7D32")
        self.file_tree.tag_configure("err", foreground="#C0392B")
        self.file_tree.tag_configure("muted", foreground=MUTED)
        self.file_tree.bind("<Delete>", lambda _e: self._remove_selected())


        scrollbar = ttkb.Scrollbar(tree_frame, orient="vertical", command=self.file_tree.yview, bootstyle="round")
        scrollbar.pack(side=RIGHT, fill=Y, padx=(6, 0))
        self.file_tree.configure(yscrollcommand=scrollbar.set)

        file_buttons = ttkb.Frame(files_content, style="TFrame")
        file_buttons.pack(fill=X, pady=(10, 0))
        self.add_button = ttkb.Button(
            file_buttons, text="Add PDFs", command=self._add_pdfs, bootstyle="secondary-outline"
        )
        self.add_button.pack(side=LEFT)
        self.remove_button = ttkb.Button(
            file_buttons, text="Remove selected", command=self._remove_selected, bootstyle="link"
        )
        self.remove_button.pack(side=LEFT, padx=(10, 0))
        self.clear_button = ttkb.Button(file_buttons, text="Clear", command=self._clear_files, bootstyle="link")
        self.clear_button.pack(side=LEFT, padx=(6, 0))
        ttkb.Label(file_buttons, text="Delete key removes selected files", style="Hint.TLabel").pack(
            side=LEFT, padx=(12, 0)
        )

        self.drop_label = ttkb.Label(files_content,
            text="Drop PDF files here or onto the list" if self.dnd_ready else "Drag-and-drop unavailable; use Add PDFs.",
            style="Hint.TLabel", padding=(0, 8))
        self.drop_label.pack(fill=X)
        if self.dnd_ready:
            for widget in (self, self.file_tree, self.drop_label):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)

        # -- Output --------------------------------------------------------
        output_content = self._section(outer, "Output")

        row1 = ttkb.Frame(output_content, style="TFrame")
        row1.pack(fill=X)
        ttkb.Label(row1, text="Folder", width=10).pack(side=LEFT)
        self.output_entry = ttkb.Entry(row1, textvariable=self.output_var)
        self.output_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 8))
        self.output_button = ttkb.Button(
            row1, text="Browse", command=self._select_output, bootstyle="secondary-outline"
        )
        self.output_button.pack(side=LEFT)
        self.separate_folders_check = ttkb.Checkbutton(
            output_content,
            text="Extract each PDF into its own folder",
            variable=self.separate_folders_var,
            bootstyle="secondary",
        )
        self.separate_folders_check.pack(anchor="w", pady=(6, 0))
        ttkb.Label(
            output_content,
            text="Leave blank to save beside each PDF. Matching names get unique folder or image names.",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(2, 12))

        row2 = ttkb.Frame(output_content, style="TFrame")
        row2.pack(fill=X)
        ttkb.Label(row2, text="Format").pack(side=LEFT)
        self.format_combo = ttkb.Combobox(
            row2, textvariable=self.format_var, values=("PNG", "JPEG"), state="readonly", width=7
        )
        self.format_combo.pack(side=LEFT, padx=(8, 20))
        ttkb.Label(row2, text="DPI").pack(side=LEFT)
        self.dpi_spin = ttkb.Spinbox(row2, from_=72, to=600, textvariable=self.dpi_var, width=6)
        self.dpi_spin.pack(side=LEFT, padx=(8, 20))
        ttkb.Label(row2, text="Pages").pack(side=LEFT)
        self.pages_entry = ttkb.Entry(row2, textvariable=self.pages_var, width=14)
        self.pages_entry.pack(side=LEFT, padx=(8, 6))
        ttkb.Label(row2, text="(All or 1-3, 5)", style="Hint.TLabel").pack(side=LEFT)
        self.overwrite_check = ttkb.Checkbutton(
            row2, text="Overwrite existing", variable=self.overwrite_var, bootstyle="round-toggle"
        )
        self.overwrite_check.pack(side=RIGHT)

        row3 = ttkb.Frame(output_content, style="TFrame")
        row3.pack(fill=X, pady=(12, 0))
        self.hq_check = ttkb.Checkbutton(
            row3,
            text="Use pdftocairo backend (often faster, smoother anti-aliasing)",
            variable=self.hq_var,
            bootstyle="round-toggle",
        )
        self.hq_check.pack(side=LEFT)

        # -- Advanced --------------------------------------------------------
        advanced_content = self._section(outer, "Advanced")
        row4 = ttkb.Frame(advanced_content, style="TFrame")
        row4.pack(fill=X)
        ttkb.Label(row4, text="Poppler bin", width=10).pack(side=LEFT)
        self.poppler_entry = ttkb.Entry(row4, textvariable=self.poppler_var)
        self.poppler_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 8))
        self.poppler_button = ttkb.Button(
            row4, text="Browse", command=self._select_poppler, bootstyle="secondary-outline"
        )
        self.poppler_button.pack(side=LEFT)
        ttkb.Label(
            advanced_content,
            text="Optional if Poppler is already available on PATH.",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        # -- Progress and actions ---------------------------------------
        progress_area = ttkb.Frame(outer, style="TFrame")
        progress_area.pack(fill=X, pady=(24, 0))
        self.progress = ttkb.Progressbar(
            progress_area, variable=self.progress_var, maximum=100, bootstyle="dark-striped"
        )
        self.progress.pack(fill=X)
        ttkb.Label(progress_area, textvariable=self.status_var, style="Hint.TLabel", wraplength=820).pack(
            anchor="w", pady=(8, 0)
        )

        actions = ttkb.Frame(outer, style="TFrame")
        actions.pack(fill=X, pady=(16, 0))
        self.open_button = ttkb.Button(
            actions, text="Open output folder", command=self._open_output, bootstyle="link"
        )
        self.open_button.pack(side=LEFT)
        ttkb.Button(actions, text="View report", command=self._show_report, bootstyle="link").pack(side=LEFT)
        self.cancel_button = ttkb.Button(
            actions, text="Cancel", command=self._cancel, bootstyle="secondary-outline", state="disabled"
        )
        self.cancel_button.pack(side=RIGHT, padx=(0, 8))
        self.convert_button = ttkb.Button(
            actions, text="Convert", command=self._start_conversion, bootstyle="dark"
        )
        self.convert_button.pack(side=RIGHT)

    # ------------------------------------------------------------------ #
    # File list management
    # ------------------------------------------------------------------ #

    def _add_pdfs(self) -> None:
        if self.running:
            return
        names = filedialog.askopenfilenames(title="Select PDF files", filetypes=[("PDF files", "*.pdf")])
        self._add_paths(names)

    def _add_paths(self, names):
        if self.running:
            return
        existing = {os.path.normcase(str(p)) for p in self.pdf_files}
        ignored = 0
        for name in names:
            try:
                path = Path(name).resolve()
                key = os.path.normcase(str(path))
                if key in existing or not path.is_file() or path.suffix.lower() != ".pdf":
                    ignored += 1
                    continue
                self.pdf_files.append(path)
                existing.add(key)
                self.file_tree.insert("", END, iid=str(path), text=path.name, values=("—", "Ready"))
            except OSError:
                ignored += 1
        self.status_var.set(f"{len(self.pdf_files)} PDF file(s) ready. {ignored} duplicate/unsupported items ignored.")

    def _on_drop(self, event):
        if self.running:
            return "none"
        try:
            self._add_paths(self.tk.splitlist(event.data))
            return "copy"
        except tk.TclError as exc:
            messagebox.showerror(APP_TITLE, f"Cannot read dropped paths: {exc}")
            return "none"

    def _remove_selected(self) -> None:
        if self.running:
            return
        selected = set(self.file_tree.selection())
        if not selected:
            return
        self.pdf_files = [path for path in self.pdf_files if str(path) not in selected]
        for item in selected:
            self.file_tree.delete(item)
        self.status_var.set(f"{len(self.pdf_files)} PDF file(s) ready.")

    def _clear_files(self) -> None:
        if self.running:
            return
        self.pdf_files.clear()
        self.file_tree.delete(*self.file_tree.get_children())
        self.progress_var.set(0)
        self.status_var.set("Select one or more PDF files to begin.")

    def _select_output(self) -> None:
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_var.set(folder)

    def _select_poppler(self) -> None:
        folder = filedialog.askdirectory(title="Select the Poppler bin folder")
        if folder:
            self.poppler_var.set(folder)

    # ------------------------------------------------------------------ #
    # Conversion
    # ------------------------------------------------------------------ #

    def _validate_settings(self):
        if convert_files is None:
            self._show_missing_package()
            return None
        if self.format_var.get() not in ("PNG", "JPEG"):
            messagebox.showwarning(APP_TITLE, "Choose PNG or JPEG.")
            return None
        if not self.pdf_files:
            messagebox.showwarning(APP_TITLE, "Please add at least one PDF file.")
            return None
        try:
            dpi = int(self.dpi_var.get())
            if not 72 <= dpi <= 600:
                raise ValueError
        except (ValueError, tk.TclError):
            messagebox.showwarning(APP_TITLE, "DPI must be a whole number from 72 to 600.")
            return None

        output = Path(self.output_var.get().strip()) if self.output_var.get().strip() else None
        if output:
            try:
                output.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                messagebox.showerror(APP_TITLE, f"Cannot create the output folder:\n{exc}")
                return None

        poppler = Path(self.poppler_var.get().strip()) if self.poppler_var.get().strip() else None
        if poppler and not poppler.is_dir():
            messagebox.showwarning(APP_TITLE, "The selected Poppler folder does not exist.")
            return None
        return dpi, output, poppler, self.hq_var.get()

    def _start_conversion(self) -> None:
        if self.running:
            return
        settings = self._validate_settings()
        if settings is None:
            return
        dpi, output, poppler, use_pdftocairo = settings
        self._save_current_settings()
        self.cancel_event.clear()
        self.output_directories = []
        self.last_report = "Conversion is running."
        for item in self.file_tree.get_children():
            self.file_tree.set(item, "status", "Ready")
            self.file_tree.item(item, tags=())
        self.progress_var.set(0)
        self._set_running(True)
        files = list(self.pdf_files)
        file_format = self.format_var.get().lower()
        page_spec = self.pages_var.get()
        overwrite = self.overwrite_var.get()
        separate_folders = self.separate_folders_var.get()
        self.worker = threading.Thread(
            target=self._convert_worker,
            args=(files, output, poppler, dpi, file_format, page_spec, overwrite,
                  use_pdftocairo, separate_folders),
            daemon=True,
        )
        self.worker.start()

    def _save_current_settings(self) -> None:
        try:
            dpi = self.dpi_var.get()
        except (tk.TclError, ValueError):
            dpi = 200
        save_config({"output": self.output_var.get().strip(),
                     "poppler": self.poppler_var.get().strip(), "dpi": dpi,
                     "format": self.format_var.get(), "pdftocairo": self.hq_var.get(),
                     "separate_folders": self.separate_folders_var.get()})

    def _convert_worker(self, files, output, poppler, dpi, file_format,
                        page_spec, overwrite, use_pdftocairo, separate_folders) -> None:
        # No Tk calls occur in this worker.
        convert_files(files, output, poppler, dpi, file_format, page_spec,
                      overwrite, use_pdftocairo, self.events, self.cancel_event,
                      separate_folders=separate_folders)

    def _poll_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "file":
                    _, item, pages, status, tag = event
                    if self.file_tree.exists(item):
                        self.file_tree.set(item, "pages", pages)
                        self.file_tree.set(item, "status", status)
                        self.file_tree.item(item, tags=(tag,) if tag else ())
                elif kind == "progress":
                    _, completed, total, name, page, eta = event
                    self.progress_var.set(completed * 100 / max(total, 1))
                    if not self.cancel_event.is_set():
                        self.status_var.set(f"Converting {name}, page {page} — {completed} of {total} ({eta})")
                elif kind in ("done", "cancelled", "error"):
                    self._set_running(False)
                    if kind == "done":
                        _, count, total, failed, self.output_directories = event
                        self.progress_var.set(count * 100 / max(total, 1))
                        summary = f"Created {count} of {total} queued image(s). {len(failed)} file(s) failed."
                        self.last_report = summary + "\n\n" + "\n".join(failed) + "\n\nOutput folders:\n" + "\n".join(self.output_directories)
                        self.status_var.set(summary)
                        if not self.closing:
                            if failed:
                                messagebox.showwarning(APP_TITLE, summary + "\nUse View report for details.")
                            else:
                                messagebox.showinfo(APP_TITLE, summary)
                    elif kind == "cancelled":
                        _, count, self.output_directories = event
                        self.last_report = f"Cancelled. {count} completed images were kept.\n\n" + "\n".join(self.output_directories)
                        self.status_var.set(f"Cancelled. {count} completed images were kept.")
                        for item in self.file_tree.get_children():
                            if self.file_tree.set(item, "status") in ("Queued", "Ready", "Reading PDF...", "Converting..."):
                                self.file_tree.set(item, "status", "Cancelled")
                    else:
                        self.last_report = event[1]
                        self.status_var.set("Conversion failed. See View report.")
                        if not self.closing:
                            messagebox.showerror(APP_TITLE, self._friendly_error(event[1]))
                    if self.closing:
                        self._finish_close()
                        return
        except queue.Empty:
            pass
        self.poll_id = self.after(100, self._poll_events)

    def _show_report(self):
        window = tk.Toplevel(self)
        window.title("Conversion report")
        window.geometry("850x450")
        text = tk.Text(window, wrap="word", padx=12, pady=12)
        scroll = ttkb.Scrollbar(window, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side=RIGHT, fill=Y)
        text.pack(fill=BOTH, expand=True)
        text.insert("1.0", self.last_report)
        text.configure(state="disabled")

    def _set_running(self, running: bool) -> None:
        self.running = running
        normal_or_disabled = "disabled" if running else "normal"
        for widget in (
            self.add_button,
            self.remove_button,
            self.clear_button,
            self.output_entry,
            self.output_button,
            self.separate_folders_check,
            self.dpi_spin,
            self.pages_entry,
            self.overwrite_check,
            self.hq_check,
            self.poppler_entry,
            self.poppler_button,
            self.convert_button,
        ):
            widget.configure(state=normal_or_disabled)
        self.format_combo.configure(state="disabled" if running else "readonly")
        self.cancel_button.configure(state="normal" if running else "disabled")
        if running:
            self.status_var.set("Reading PDF information...")

    def _cancel(self) -> None:
        self.cancel_event.set()
        self.cancel_button.configure(state="disabled")
        self.status_var.set("Cancelling after the current Poppler call; renderer timeout is 120 seconds…")

    def _open_output(self) -> None:
        raw_output = self.output_var.get().strip()
        folder = Path(self.output_directories[-1]) if self.output_directories else (Path(raw_output) if raw_output else None)
        if folder is None or not folder.exists():
            messagebox.showwarning(APP_TITLE, "No output folder is available yet.")
            return
        try:
            if sys.platform == "win32":
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"Could not open the folder:\n{exc}")

    @staticmethod
    def _friendly_error(error: str) -> str:
        lowered = error.lower()
        if "poppler" in lowered or "page count" in lowered or "pdfinfo" in lowered:
            return (
                "Poppler could not read the PDF. Install Poppler and select its bin folder "
                f"under Advanced.\n\nTechnical details: {error}"
            )
        return f"The conversion could not be completed.\n\n{error}"

    def _show_missing_package(self) -> None:
        messagebox.showerror(APP_TITLE, engine_dependency_message())

    def _show_missing_ttkbootstrap(self) -> None:
        messagebox.showerror(
            APP_TITLE,
            "The ttkbootstrap package is not installed.\n\nRun:\n"
            "python -m pip install ttkbootstrap",
        )
        self.destroy()

    def _on_close(self) -> None:
        if self.closing:
            return
        if self.running:
            if not messagebox.askyesno(APP_TITLE, "Cancel conversion and close after cleanup?"):
                return
            self.closing = True
            self._cancel()
            return
        self._finish_close()

    def _finish_close(self):
        self._save_current_settings()
        if hasattr(self, "poll_id"):
            self.after_cancel(self.poll_id)
        self.destroy()


if __name__ == "__main__":
    PdfConverterApp().mainloop()
