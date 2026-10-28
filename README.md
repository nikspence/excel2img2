# excel2img2

`excel2img2` is a successor to [glexey's excel2img](https://github.com/glexey/excel2img).
Its main change is `pil_img()`: capture an Excel worksheet or range and return a
`PIL.Image.Image` directly, without writing and reading an intermediate image file.
The original `export_img()` helper remains available for saving images to disk.

## Requirements

- Windows with Microsoft Excel desktop installed and configured.
- Python 3.9 or newer.
- An interactive Windows desktop session with clipboard access. Capture uses
  Excel COM automation and the Windows clipboard, so clipboard contents are replaced.
- Pillow and pywin32 (installed automatically by pip on Windows).

Excel itself is not installed by pip. This library uses Excel to render workbooks;
it does not provide a renderer for Linux, macOS, or environments without Excel.

## Installation

From this directory, which contains `pyproject.toml`:

```powershell
python -m pip install .
```

For development, use an editable installation:

```powershell
python -m pip install -e .
```

These commands install the local source package; publication to PyPI is separate.

## Usage

```python
import excel2img2

# Return an image of the used cells on the first sheet.
image = excel2img2.pil_img("workbook.xlsx")
print(image.size, image.mode)

# Capture a specific worksheet range.
image = excel2img2.pil_img("workbook.xlsx", page="Sheet1", _range="B2:F20")

# Work with the Pillow image directly.
thumbnail = image.copy()
thumbnail.thumbnail((800, 800))

# Saving is optional and controlled by the caller.
image.save("capture.png")
```

To encode an image in memory for another API:

```python
from io import BytesIO
from excel2img2 import pil_img

image = pil_img("workbook.xlsx", page="Sheet1", _range="A1:D12")
buffer = BytesIO()
image.save(buffer, format="PNG")
png_bytes = buffer.getvalue()
```

### `pil_img(fn_excel, page=None, _range=None)`

Returns a loaded `PIL.Image.Image` that remains usable after the workbook closes.

| Argument | Meaning |
| --- | --- |
| `fn_excel` | Path to a workbook Excel can open, such as `.xlsx` or `.xls`. |
| `page` | Sheet name or one-based sheet index. Defaults to the first sheet when no range is supplied. |
| `_range` | Excel range or named range. If omitted, captures the sheet's used cells. |

Examples:

```python
from excel2img2 import pil_img

sheet = pil_img("workbook.xlsx", page="Sales")
cells = pil_img("workbook.xlsx", page="Sales", _range="A1:E15")
qualified = pil_img("workbook.xlsx", _range="'Sales Data'!A1:E15")
named = pil_img("workbook.xlsx", _range="ReportArea")
chart = pil_img("workbook.xlsx", page="Chart1")  # Standalone chart sheet.
```

A sheet-qualified `_range` determines the sheet even if `page` is also supplied.
Unqualified ranges and named ranges are resolved by Excel; qualify them or pass
`page` when a particular sheet is required.

The workbook is opened read-only in a separate Excel instance and closed after
capture. Missing workbooks, unavailable Excel, invalid sheets/ranges, and clipboard
capture failures raise exceptions.

### `export_img(fn_excel, fn_image, page=None, _range=None)`

The file-writing helper supports PNG, BMP, and GIF:

```python
from excel2img2 import export_img

export_img("workbook.xlsx", "capture.png", page="Sheet1", _range="A1:D12")
```

For migration, replace imports of `excel2img` with `excel2img2` and use
`image = excel2img2.pil_img(workbook, page, cell_range)` wherever an in-memory
image is needed. Unlike `export_img()`, `pil_img()` takes no output filename.

## License and attribution

Licensed under the Apache License, Version 2.0; see [LICENSE](LICENSE).
Derived from `excel2img` by Alexey Gaydyukov (glexey), copyright 2016.
Modified in 2026 by Nikolaus Spence to provide the in-memory Pillow image API
and package it as `excel2img2`.
