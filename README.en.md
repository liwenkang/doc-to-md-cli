# doc_to_md_cli Usage Guide

`doc_to_md_cli.py` converts `.doc` / `.docx` files to Markdown using the locally installed Microsoft Word.
It also exports embedded images as PNG files and wires them into the generated Markdown.

## Features

- Convert a single `.doc` / `.docx` file or an entire directory to `.md`
- Automatically detect headings, normal paragraphs, and tables, and generate corresponding Markdown structures
- Export inline images as PNG into a dedicated `_images` directory and reference them in the generated Markdown
- In batch mode, reuse a single Word process and handle common COM / RPC errors with retry and fallback logic
- Provide quiet and verbose modes for use in CI or local debugging

## Requirements

- Operating System: Windows
- Microsoft Word installed (COM automation enabled)
- Python: 3.8+ (this repo currently uses Python 3.14)
- Python and system dependencies:
  - `pywin32` (provides `win32com.client` / `pywintypes`)
  - ImageMagick (CLI tool `magick`, used to convert EMF to PNG when needed)

### Installing Dependencies

Install `pywin32` in your current Python environment:

```bash
pip install pywin32
```

Install ImageMagick (if not already installed):

- Recommended: download the installer from https://imagemagick.org
- During installation, enable "Add application directory to your system path" or ensure the `magick` command is available in your shell.

## Basic Usage

Run the script from the repo root (using `python` or a full interpreter path):

### Single File Conversion

```bash
python scripts/doc_to_md_cli.py "docs/spec-template.docx"
```

By default, a `.md` file with the same base name will be created next to the input:

- Input: `docs/spec-template.docx`
- Output: `docs/spec-template.md`

You can specify the output path with `-o/--output`:

```bash
python scripts/doc_to_md_cli.py "docs/spec-template.docx" -o "docs/spec-template-converted.md"
```

### Directory Batch Conversion

Convert all `.doc` / `.docx` files under a directory:

```bash
python scripts/doc_to_md_cli.py docs
```

Recursively scan subdirectories:

```bash
python scripts/doc_to_md_cli.py docs -r
```

The script will:

- Generate a `.md` file with the same base name for each Word file in place
- Export images to a corresponding `<markdown file name>_images/` directory
- Print batch statistics (number of successful / failed files)

## Command Line Options

```text
python scripts/doc_to_md_cli.py INPUT [options]
```

- `INPUT` (required):
  - A single `.doc` / `.docx` file path, or
  - A directory path (batch mode)

### Common Options

- `-o, --output PATH`

  - Only valid in single-file mode
  - Specify the output Markdown file path. If omitted, the script uses the same base name as the input and only changes the extension to `.md`.

- `--no-optimize`

  - By default, the script performs a few lightweight cleanups on the generated Markdown:
    - Remove common control characters
    - Merge excessive blank lines
    - Fix headings like `##Heading` to `## Heading`
  - With this flag, those cleanups are skipped and the raw output is preserved.

- `--quiet`

  - Quiet mode:
    - Only print final statistics and error messages
    - Do not print detailed per-file processing logs

- `--verbose`
  - Verbose mode:
    - Print source/target paths for each file
    - Print a success message for each conversion

> Note: `--quiet` and `--verbose` cannot be used at the same time.

- `-r, --recursive`
  - When `INPUT` is a directory, recursively scan subdirectories for `.doc` / `.docx` files in batch mode

## Output Structure

Given an input file `docs/spec-template.docx`, typical outputs are:

- Markdown file:
  - `docs/spec-template.md`
- Image directory:
  - `docs/spec-template_images/`
  - Contains exported PNG images: `img_1.png`, `img_2.png`, ...

An example image reference in Markdown looks like:

```markdown
![image1](spec-template_images/img_1.png)
```

During conversion:

- The script first tries to use Word's `SaveAsPicture` to export directly to PNG.
- If the current Word version does not support it or an error occurs, it exports EMF and calls `magick` to convert EMF to PNG.
- After a successful conversion, the intermediate `.emf` file is deleted so that only PNGs remain.

## Error Handling & Retry Strategy

Because Word COM automation can be fragile under heavy or long-running batch operations, the script builds in several safeguards:

- Paragraph / table level fallback:
  - If processing a single paragraph or an entire table fails, that paragraph/table is skipped, while the rest of the document continues to be processed.
- Per-file retry:
  - In single-file mode, if conversion fails once, the script automatically retries one more time.
  - In directory batch mode, each file is attempted up to two times.
- Word process restart:
  - In batch mode, if typical COM errors such as "RPC server unavailable", "call failed", or "invalid pointer" are detected, the script will:
    - Quit the current Word process,
    - Start a fresh Word instance,
    - Then retry the current file once.

These mechanisms significantly improve stability in large-scale batch conversions.

## Tips & Caveats

If you want to extend this tool (for example, custom output directory, disabling image export, supporting more Word object types, etc.), you can build on top of `scripts/doc_to_md_cli.py`.

## FAQ

**Q1: I get an error saying `win32com.client` or `pywintypes` is missing. What should I do?**
Make sure `pywin32` is installed in the Python environment you are using to run the script:

```bash
pip install pywin32
```

Also ensure that the interpreter you use to run the script is the same one where `pywin32` was installed (for example, the same virtual environment).

**Q2: The script says `magick` is not found, or images are not exported as PNG. Why?**
When Word cannot natively export PNG via `SaveAsPicture`, the script falls back to ImageMagick's `magick` command to convert EMF to PNG:

- Verify that ImageMagick is installed correctly and `magick -version` works in your shell.
- If `magick` is not available, the script will fall back to using EMF files, so image links in Markdown may have a `.emf` extension.
- For consistent PNG output, install and properly configure ImageMagick.

**Q3: Can I use this script on macOS or Linux?**
Currently no. The script relies on the Microsoft Word COM automation API, which is only available on Windows with Word installed.

**Q4: Why are some pictures or shapes missing in the generated Markdown?**
The current implementation mainly processes `InlineShapes` in Word documents (inline images). Some floating shapes/objects (for example, with wrapping modes other than inline) may not be detected and exported. In such cases:

- Try changing the picture layout in Word to "In line with text" and re-run the conversion; or
- Extend the script to support additional Word object types.

**Q5: Is it safe to run multiple batch conversions in parallel?**
This is not recommended. Word COM is fragile under multi-process/multi-threaded concurrent access and is prone to contention, RPC errors, and random crashes. Instead:

- Run a single batch job per machine at a time; and
- If you need parallelism, distribute the workload across multiple machines, each running one process.

**Q6: Why doesn't the Markdown layout exactly match the original Word document?**
The goal of this tool is "high-fidelity but not pixel-perfect" Markdown:

- Headings, paragraphs, tables, and images are preserved as much as reasonably possible.
- Complex layouts (multiple columns, special styles, some SmartArt/objects, etc.) are not reproduced 1:1.
- You may still want to do a light manual cleanup on the generated Markdown if you need fully polished output.
