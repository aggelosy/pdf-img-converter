# PDF to Image Converter 3.0.0

Version 3.0.0 introduces a rewritten desktop interface and a choice of output layouts. Convert PDFs into individual folders or save their images in a shared destination, with source-specific names that keep matching PDF filenames separate.

## Highlights

- New **Tkinter and ttkbootstrap interface** with progress, cancellation, and conversion reports.
- **Extract each PDF into its own folder** is enabled by default and remembered between launches.
- Disable the checkbox to write images directly into the selected output folder. Leave the output field blank to save beside each source PDF.
- Source-path identifiers distinguish PDFs with matching names in either layout, including when overwrite is enabled.
- Saved output preferences, existing-image overwrite control, page-size checks, and bounded Poppler calls.

## Upgrading from v2.0.0

This is a major interface and output change:

- Launch **`pdf_to_image_converter.py`** instead of `pdf_converter_pro.py`, and reinstall dependencies from the updated `requirements.txt`.
- Export formats are **PNG and JPEG**. TIFF, WebP, and BMP export have been removed.
- The default DPI is now **200**, with any whole number from 72 to 600 accepted.
- Output folder and image names have changed. Selected pages now use their original PDF page numbers.
- Preview, grayscale, custom filename-prefix, quality-slider, and thread-count controls have been removed. JPEG quality is fixed at 92; rendering uses one Poppler thread.
- Open destinations manually with **Open output folder**.

## Download and run

Download and extract **`pdf-img-converter-v3.0.0-source.zip`**, then open a terminal inside its project folder:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe pdf_to_image_converter.py
```

The source package requires Python 3.10 or later with Tkinter. Install Poppler separately and select its executable folder under **Advanced → Poppler bin**, or make the required executables available on `PATH`.

The release assets include **`SHA256SUMS.txt`** with the source archive's checksum. Windows has been tested; macOS and Linux have not been tested.

[Full changelog](https://github.com/aggelosy/pdf-img-converter/blob/v3.0.0/CHANGELOG.md) · [Changes since v2.0.0](https://github.com/aggelosy/pdf-img-converter/compare/v2.0.0...v3.0.0)
