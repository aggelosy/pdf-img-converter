# Changelog

## 3.0.0 — 2026-10-05

### Added

- A Tkinter and ttkbootstrap desktop interface with background conversion, progress, cancellation, and a conversion report.
- An **Extract each PDF into its own folder** checkbox, enabled by default and remembered between launches.
- Shared-folder output when the checkbox is disabled, with source-specific image names.
- Sanitized PDF names and source-path identifiers to distinguish same-name PDFs from different folders.
- Saved preferences for output folder, Poppler folder, DPI, format, renderer, and output layout.
- An overwrite option; existing images are preserved with numbered copies when overwrite is disabled.
- Per-page size checks when dimensions are available, plus timeouts for PDF information and rendering calls.
- Repository documentation, output examples, and a `.gitignore` for local files and generated output.

### Changed

- Replaced the PyQt5 interface and PyMuPDF preview dependency with Tkinter, ttkbootstrap, and optional tkinterdnd2 integration.
- Renamed the launch script from `pdf_converter_pro.py` to `pdf_to_image_converter.py`.
- Limited output formats to PNG and JPEG.
- Changed default resolution from 300 DPI to 200 DPI, with whole-number input from 72 to 600.
- Changed folder and image names to include source-path identifiers. Selected pages now retain their original PDF numbering.
- Fixed JPEG quality at 92 and Poppler rendering at one thread.
- Enlarged the window to accommodate the output-layout control, which is disabled during conversion.

### Removed

- TIFF, WebP, and BMP output.
- First-page preview, grayscale mode, custom filename prefixes, adjustable image-quality controls, and configurable thread count.
- Automatic opening of the destination after conversion; **Open output folder** is available instead.

Changes are documented against the previous public tag, `v2.0.0`.
