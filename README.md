# NotebookLM Unified Workflow

> Unifies, cleans and converts NotebookLM presentations in a single step.

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## What does it do?

This script automates the entire NotebookLM presentation cleaning process:

1. **Merges** multiple PPTX files into one
2. **Converts** to lossless PDF
3. **Removes** NotebookLM watermark
4. **Converts** back to clean PPTX

<img width="1408" height="768" alt="imagen" src="https://github.com/user-attachments/assets/46cc7e47-5185-4bf3-81ff-0c4b8f34a8c7" />


## Why this project?

NotebookLM generates presentations with a maximum of 15 slides per file. For longer documents, you end up with multiple PPTXs that need to be:
- Merged into one
- Converted to PDF
- Cleaned from watermark
- Converted back to PPTX

**This script does ALL of that with a single command.**

## Installation

```bash
# Clone the repository
git clone https://github.com/Milor123/notebooklm-unified.git
cd notebooklm-unified

# Install dependencies (or the script will auto-install them)
pip install -r requirements.txt
```

## Usage

```bash
# Process all PPTXs in current directory
python workflow.py

# Specify input folder
python workflow.py --input "C:\Downloads\presentations"

# Custom output filename
python workflow.py --output "my_clean_presentation.pptx"

# Debug mode (shows watermark coordinates)
python workflow.py -d
```

### Options

| Option | Description |
|--------|-------------|
| `-i, --input` | Input folder or PPTX file |
| `-o, --output` | Output filename (default: presentacion_limpia.pptx) |
| `-v, --verbose` | Verbose mode |
| `-d, --debug` | Debug: shows watermark coordinates |

## Example

```bash
# Let's say you have:
# part1.pptx (15 slides)
# part2.pptx (15 slides)
# part3.pptx (10 slides)

python workflow.py --output "complete_presentation.pptx"

# Result: complete_presentation.pptx (40 slides, no watermark)
```

## Requirements

- Python 3.8+
- Windows, macOS or Linux

Dependencies are automatically installed the first time you run the script.

## How it works

```
PPTXs (multiple)
    → Merge PPTXs (extract images, maintain order)
    → PDF lossless (img2pdf)
    → Remove watermark (column-by-column sampling algorithm)
    → Clean PPTX (PyMuPDF + python-pptx)
```

## Credits

- **Watermark Remover**: Algorithm based on work by [neosun100/notebooklm-watermark-remover](https://github.com/neosun100/notebooklm-watermark-remover)
- **Libraries**: [python-pptx](https://python-pptx.readthedocs.io/), [img2pdf](https://github.com/ocs/img2pdf), [PyMuPDF](https://pymupdf.readthedocs.io/)

## License

MIT License - see [LICENSE](LICENSE)

---

⭐️ If this script was useful to you, consider giving the project a star!
