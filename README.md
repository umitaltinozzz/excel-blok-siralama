# Excel Block Sorter

**Sorts total blocks and their detail rows in Excel reports — a local Windows tool**

[![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![openpyxl](https://img.shields.io/badge/openpyxl-Excel-217346?style=for-the-badge)](https://openpyxl.readthedocs.io/)
[![PyInstaller](https://img.shields.io/badge/PyInstaller-EXE-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://pyinstaller.org/)
[![Windows](https://img.shields.io/badge/Windows-Desktop-0078D6?style=for-the-badge&logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![License: MIT](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](./LICENSE)
![Status](https://img.shields.io/badge/Status-Completed-blue?style=for-the-badge)

---

## Overview

A local Windows tool for finance and reporting spreadsheets. It detects total rows and the detail rows beneath them, then performs a two-level sort: total blocks from largest to smallest, and detail rows within each block from largest to smallest — without breaking the block structure.

## Example

```
Input                          Output
367.349.746 PERSONEL...        367.349.746 PERSONEL...
117.173.808 DISARIDAN...       124.231.547 AMORTISMAN...
 60.077.177 NAKLIYE...         117.173.808 DISARIDAN...
124.231.547 AMORTISMAN...       60.077.177 NAKLIYE...
```

## Project Status

Completed utility. A prebuilt `Excel_Blok_Siralayici.exe` is included for users without Python.

## Features

- Detects total rows and their detail rows automatically
- Two-level descending sort that keeps each block intact
- Simple GUI and a `.bat` launcher for non-technical users
- Diff tool to compare the original and sorted workbook
- Unit tests for the sorter and diff logic

## Tech Stack

| Layer | Technology |
|---|---|
| Language | Python |
| Excel I/O | openpyxl |
| GUI | Tkinter |
| Packaging | PyInstaller (Windows .exe) |
| Testing | pytest |

## Getting Started

Double-click `Excel_Blok_Siralama.bat` (or the `.exe`) and pick a workbook. From source:

```bash
pip install -r requirements.txt
python excel_block_sorter_gui.py        # GUI
python sort_excel_blocks.py input.xlsx  # CLI
pytest
```

## Project Structure

```
excel-blok-siralama/
├── Excel_Blok_Siralama.bat
├── Excel_Blok_Siralayici.exe
├── README.md
├── excel_block_sorter.py
├── excel_block_sorter_gui.py
├── excel_diff.py
├── requirements.txt
├── sort_excel_blocks.py
├── test_excel_block_sorter.py
├── test_excel_diff.py
```

## License

[MIT License](./LICENSE)
