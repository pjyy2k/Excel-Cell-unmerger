# Excel Cell Unmerger

A lightweight desktop utility that unmerges all merged cells in an Excel file and fills each cell with the original value — making your data ready for sorting, filtering, and analysis.

> 엑셀 파일의 모든 병합 셀을 해제하고 원래 값으로 채워 넣는 도구입니다.

---

## Features

- Supports `.xlsx`, `.xls`, and `.xlsm` formats
- Processes all sheets in the workbook at once
- Fills every unmerged cell with the value from the original merged region
- Saves the result as `Unmerged.xlsx` in the same folder as the source file
- Progress bars per sheet so you can track long-running files

---

## Usage

### Option 1 — Download the EXE (Windows only, no Python required)

1. Download the latest release from the [Releases page](https://github.com/pjyy2k/Excel-Cell-unmerger/releases)
2. Run `Excel-cells-unmerge.exe`
3. Select your Excel file in the dialog
4. Find `Unmerged.xlsx` in the same folder as your original file

### Option 2 — Run from source (Python)

**Requirements**

| Package | Purpose |
|---------|---------|
| [xlwings](https://www.xlwings.org/) | Excel automation |
| [tqdm](https://github.com/tqdm/tqdm) | Progress bars |
| tkinter | File dialogs and message boxes (included with Python) |

**Install dependencies**

```bash
pip install xlwings tqdm
```

**Run**

```bash
python xlwings_unmerge.py
```

---

## Build from Source (PyInstaller)

Requires [PyInstaller](https://pyinstaller.org/):

```bash
pip install pyinstaller
pyinstaller xlwings_unmerge.py --splash "./splash.png" -F -w
```

Or use the included spec file:

```bash
pyinstaller xlwings_unmerge.spec
```

The executable will be placed in the `dist/` folder.

> **Note:** Before building, uncomment the `pyi_splash` lines at the top of `xlwings_unmerge.py`.

---

## How It Works

For each sheet, the tool iterates over every cell in the used range. When a merged cell is found, it captures the value, unmerges the region, then writes that value back to every cell that was part of the merge.

```
Before                     After
┌───────────────┐          ┌───────┬───────┬───────┐
│               │          │  ABC  │  ABC  │  ABC  │
│      ABC      │   →      ├───────┼───────┼───────┤
│               │          │  ABC  │  ABC  │  ABC  │
└───────────────┘          └───────┴───────┴───────┘
```

---

## License

MIT
