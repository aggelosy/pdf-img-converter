# PDF to Image Converter

[![Version 3.0.0](https://img.shields.io/badge/Version-3.0.0-6F42C1)](https://github.com/aggelosy/pdf-img-converter/releases)
![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![PNG and JPEG output](https://img.shields.io/badge/Output-PNG%20%7C%20JPEG-2EA44F)

A desktop application for converting PDF pages into images in bulk. Choose the pages, resolution, image format, and output layout from a simple interface.

All PDF processing runs locally using Poppler. No document uploads are required.

[Quick start](#quick-start) · [Usage](#usage) · [Output layouts](#output-layouts) · [Troubleshooting](#troubleshooting) · [Changelog](CHANGELOG.md)

## Features

- **Batch conversion** with a file picker and drag and drop when available.
- **PNG and JPEG export** at 72–600 DPI, with a default of 200 DPI.
- **Page selection** for all pages or ranges such as `1-3, 5`.
- **Flexible output** with a separate folder for each PDF or images in a shared folder.
- **Matching filename handling** through identifiers based on each PDF's source path.
- **Progress, cancellation, and reports**, plus saved output preferences and overwrite control.

## Requirements

- Python 3.10 or later with Tkinter and a desktop display.
- Python dependencies listed in [requirements.txt](requirements.txt).
- Poppler, installed separately from the Python dependencies.

Windows is the tested platform. Expand the macOS and Linux instructions below for other desktop environments.

## Upgrading from v2.0.0

Version 3.0.0 replaces the PyQt5 interface with Tkinter and ttkbootstrap. Install the updated dependencies and launch `pdf_to_image_converter.py` instead of `pdf_converter_pro.py`.

- Supported output formats are now **PNG and JPEG**. TIFF, WebP, and BMP export are no longer available.
- The default resolution is **200 DPI**. Any whole-number DPI from 72 to 600 can be selected.
- Output names now include a source-path identifier, and selected pages retain their original PDF page numbers.
- The previous preview, grayscale, custom filename prefix, image-quality sliders, and thread-count controls are not available in this interface. JPEG quality is fixed at 92, and rendering uses one Poppler thread.
- Use **Open output folder** to open a destination manually.

See [CHANGELOG.md](CHANGELOG.md) for the release changes.

## Quick start

Download the source archive from [Releases](https://github.com/aggelosy/pdf-img-converter/releases), or clone the repository, then open a terminal in the extracted or cloned project folder.

Install Poppler using the Windows build linked in the [pdf2image installation guide](https://pdf2image.readthedocs.io/en/latest/installation.html). Extract it and locate the executable folder, usually `Library\bin`.

The app requires `pdfinfo.exe` and the selected renderer: `pdftoppm.exe` by default, or `pdftocairo.exe` when that backend is enabled.

Create a Python environment, install dependencies, and launch the app:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe pdf_to_image_converter.py
```

In **Advanced → Poppler bin**, select the Poppler executable folder. Leave this field blank if the required executables are already on your `PATH`.

<details>
<summary>macOS and Linux setup</summary>

Ensure your Python installation includes Tkinter. These platforms have not been tested.

Install Poppler with the command appropriate to your system, as documented in the [pdf2image installation guide](https://pdf2image.readthedocs.io/en/latest/installation.html):

```bash
# macOS with Homebrew
brew install poppler

# Ubuntu / Debian
sudo apt-get install poppler-utils
```

Create a Python environment, install dependencies, and launch:

```bash
python3 -m venv .venv
./.venv/bin/python -m pip install -r requirements.txt
./.venv/bin/python pdf_to_image_converter.py
```

Tkinter support must be available through your Python distribution; it is not installed by `requirements.txt`.

</details>

## Usage

1. Click **Add PDFs**, or drag PDFs onto the file list when available.
2. Select an output folder. Leave it blank to save beside each source PDF.
3. Choose whether to **Extract each PDF into its own folder**.
4. Set **Format**, **DPI**, and **Pages**. Use `All` or a range such as `1-3, 5`.
5. Optionally enable **Overwrite existing** or **Use pdftocairo backend**.
6. Click **Convert**. Use **View report** for results or **Open output folder** to open the most recent destination.

Page selections apply to every PDF. A selection outside a PDF's page count fails that file while other readable PDFs can continue.

**Cancel** stops the batch after the current Poppler call finishes or times out. Images already saved are kept.

## Output layouts

The **Extract each PDF into its own folder** checkbox is enabled by default and remembered between launches.

### Separate folders: checkbox enabled

Two PDFs named `report.pdf` from different source folders produce separate destinations:

```text
output/
├── report - 1a2b3c4d5e6f/
│   ├── page 001.png
│   └── page 002.png
└── report - 7f8e9d0c1b2a/
    ├── page 001.png
    └── page 002.png
```

### Shared folder: checkbox disabled

Images are saved directly into the output folder with a PDF-specific prefix:

```text
output/
├── report - 1a2b3c4d5e6f - page 001.png
├── report - 1a2b3c4d5e6f - page 002.png
├── report - 7f8e9d0c1b2a - page 001.png
└── report - 7f8e9d0c1b2a - page 002.png
```

The identifiers shown are examples. Every PDF receives a sanitized name and a 12-character identifier derived from its resolved source path. Moving or renaming a PDF changes its identifier. JPEG output uses `.jpg`; page numbers retain their original PDF numbering.

With no output folder selected, separate folders or individual images are created beside each PDF according to the checkbox setting.

### Existing images

- **Overwrite disabled:** preserve existing images and add numbered copies, such as `page 001 (2).png`.
- **Overwrite enabled:** replace matching destination images after rendering completes.

Other images in the destination remain in place, including pages from earlier conversions.

<details>
<summary>Saved settings</summary>

The output folder, Poppler folder, DPI, format, renderer, and folder checkbox are saved when conversion starts and when the app closes. The file list, page selection, and overwrite choice reset between launches.

Settings locations:

- **Windows:** `%LOCALAPPDATA%\PdfToImageConverter\settings.json`
- **macOS / Linux:** `$XDG_CONFIG_HOME/PdfToImageConverter/settings.json`, falling back to `~/.config/PdfToImageConverter/settings.json`.

To restore defaults, close the app and remove its `settings.json` file.

</details>

## Troubleshooting

- **Poppler not found:** select the folder containing `pdfinfo` and the selected renderer under **Advanced → Poppler bin**. Reopen the terminal after changing `PATH`.
- **Missing Python package:** run the dependency installation command with the same Python environment used to launch the app.
- **Drag and drop unavailable:** use **Add PDFs**.
- **PDF conversion failed:** check **View report**, the page selection, and output-folder permissions. The interface does not provide password entry for protected PDFs.
- **Page too large:** lower the DPI. When page dimensions are available, the app rejects pages exceeding 80 million pixels.
- **Cancellation takes time:** PDF information calls time out after 30 seconds; page rendering calls time out after 120 seconds.

When reporting a problem, include your operating system, Python version, selected renderer, and the error from **View report**.

