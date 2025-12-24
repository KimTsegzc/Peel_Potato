Peel Potato — FastBI for Excel

1. Purpose

Peel Potato is a small, fast desktop helper for quickly generating common Excel charts from worksheet data. It is intended as a lightweight floating tool you run alongside Excel to build and modify charts (Line, Bar, Column, Pie, Area, Scatter, Radar) using simple range or column inputs.

2. Functions

- Floating PyQt6 GUI that discovers the active Excel instance and open workbooks/sheets.
- Create charts embedded in the selected worksheet by specifying:
  - Dim (category / X): column letter or explicit range (e.g. A or A2:A13)
  - Values: one or multiple columns/ranges (e.g. B, C or B:C or B2:B13)
  - Optional Labels column/range for point labels
  - Multi-series mode (clustered/stacked/100% stacked, pie/doughnut variants)
- "Create Chart" to insert a chart and remember it.
- "Change Chart" to modify the most recently created chart in-place (replace series / update chart type).

3. Packages and version constraints

Recommended Python environment:
- Python 3.8 — 3.11 (tested on 3.9)

Python packages (minimum recommended):
- xlwings >= 0.27 (used to communicate with Excel via COM)
- PyQt6 >= 6.2 (GUI)
- pywin32 (provides COM constants used by Excel integration — often installed with xlwings on Windows)

Example minimal requirements.txt lines:

xlwings>=0.27
PyQt6>=6.2
pywin32>=305

4. Architecture and build methods

Architecture overview
- peel_potato.py: single-file PyQt6 application.
  - UI layer: PyQt6 floating window (controls, inputs, buttons).
  - Excel bridge: xlwings (wraps pywin32 COM objects) to read ranges and create ChartObjects.
  - Chart logic: simple mapping from chosen chart type + multi-series mode to Excel ChartType constants; series are added explicitly so the category axis (Dim) is respected.

Runtime behavior
- The app connects to the active Excel instance (xlwings.apps.active). It reads sheet ranges using xlwings.Range objects and uses the sheet.api (COM) to create ChartObjects and Series.
- The tool tries to be permissive about input syntax: column letters (B or B:C), comma lists (B,C), or explicit ranges (B2:B13). When a plain column letter is used, the app infers the data rows from Excel's UsedRange and excludes the header row.

How to run (development / quick test)
1. Create and activate a virtualenv (recommended):
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
2. Install packages:
   pip install xlwings PyQt6 pywin32
3. Run the app with the venv Python:
   .\.venv\Scripts\python.exe peel_potato.py

(If you already have a venv in this project the executable path may be e.g.
D:/Projs/Peel Potato/.venv/Scripts/python.exe peel_potato.py)

How to build a single exe (Windows)
- Install PyInstaller in the environment: pip install pyinstaller
- Build a one-file executable (example):
  pyinstaller --onefile --windowed peel_potato.py

Notes when packaging:
- The app uses xlwings/COM which requires Excel installed on target machine.
- PyInstaller builds for Windows must be run on Windows.
- You may need to include additional runtime files/data (icons, manifests) depending on your build.

Limitations and known issues
- The prototype targets Windows (COM Automation). It will not work on macOS the same way (macOS uses a different xlwings backend).
- Histogram support and advanced chart customizations are intentionally removed in the prototype; adding histograms or statistical binning requires extra logic (or using Excel analysis tool).
- Very large ranges: the code uses UsedRange and may conservatively select long spans for column-letter inputs; explicit ranges (B2:B1000) are recommended for performance.
- Chart formatting is minimal — the prototype focuses on fast creation and in-place modification.

Next steps / ideas
- Add a small examples panel with one-click presets.
- Add persistent last-used values and a small history menu.
- Add an "Export PNG" button to save the current chart as an image.
- Package using PyInstaller and create a small installer (Inno Setup / similar).

Contact
- This is a local prototype script. For enhancements, provide sample workbooks or reproducible inputs and I will adjust the behavior.

-- End of readme --
