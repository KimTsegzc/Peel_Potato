# Peel Potato V3.1

**Date:** December 20, 2025  
**Author:** XIE Xin, Guangzhou  
**Development Tools:** Claude Sonnet 4.5

---

## Overview

Peel Potato is a FastBI tool for Excel, providing quick data visualization and analysis capabilities.

## Files Included

- `peel_potato.py` - Main application entry point with GUI
- `peel_potato_engine.py` - Core engine for data processing
- `peel_potato_logic.py` - Business logic implementation
- `peel_potato_prettify.py` - UI styling and formatting
- `create_sample_data.py` - Sample data generator for testing

## Requirements

- Python 3.12+
- xlwings
- PyQt6
- pywin32
- pandas
- openpyxl

## Installation

1. Create a virtual environment:
   ```bash
   python -m venv .venv
   ```

2. Activate the virtual environment:
   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```

3. Install dependencies:
   ```bash
   pip install xlwings PyQt6 pywin32 pandas openpyxl
   ```

## Usage

Run the main application:
```bash
python peel_potato.py
```

## Version History

### V3.1 (2025-12-20)
- Environment setup and dependency management
- Sample data generation utility added
- PyQt6 compatibility improvements

---

**© 2025 XIE Xin. All rights reserved.**
