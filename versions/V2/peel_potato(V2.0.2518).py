import sys
import time
import xlwings as xw
import os
from PyQt6 import QtWidgets, QtCore, QtGui
import peel_potato_prettify

try:
    from win32com.client import constants as xlconst
except Exception:
    xlconst = None  


class PeelPotato(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Peel Potato — FastBI for Excel")
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowType.WindowStaysOnTopHint)
        self.setFixedWidth(420)
        
        # Set base font: 12pt Microsoft YaHei UI
        base_font = QtGui.QFont("Microsoft YaHei UI", 12)
        self.setFont(base_font)

        # set window icon from project icon (prefer 60x60 variants) if available
        try:
            icon_dir = os.path.dirname(__file__)
            candidates = [
                os.path.join(icon_dir, 'Icon_app.ico'),
                os.path.join(icon_dir, 'Icon_high_res.ico'),
            ]
            for icon_path in candidates:
                if os.path.exists(icon_path):
                    try:
                        # try to load and scale to 60x60 for crispness
                        pix = QtGui.QPixmap(icon_path)
                        if not pix.isNull():
                            pix = pix.scaled(60, 60, QtCore.Qt.AspectRatioMode.KeepAspectRatio, QtCore.Qt.TransformationMode.SmoothTransformation)
                            self.setWindowIcon(QtGui.QIcon(pix))
                        else:
                            # fallback: let QIcon try to load (may support SVG)
                            self.setWindowIcon(QtGui.QIcon(icon_path))
                    except Exception:
                        try:
                            self.setWindowIcon(QtGui.QIcon(icon_path))
                        except Exception:
                            pass
                    break
        except Exception:
            pass

        layout = QtWidgets.QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)  # Set consistent margins
        layout.setSpacing(8)  # Reduce spacing between elements

        form = QtWidgets.QFormLayout()
        
        # Active workbook/sheet notification (read-only) - 10pt font
        self.active_label = QtWidgets.QLabel("(no Excel detected)")
        active_font = QtGui.QFont("Microsoft YaHei UI", 10)
        self.active_label.setFont(active_font)
        form.addRow("Active at:", self.active_label)

        # Loading / status indicator - 10pt font
        self.load_label = QtWidgets.QLabel("")
        self.load_label.setFont(active_font)
        form.addRow("Status:", self.load_label)

        self.chart_type = QtWidgets.QComboBox()
        # More chart types and clearer labels
        self.chart_type.addItems([
            "Line",
            "Bar (horizontal)",
            "Column (vertical)",
            "Pie",
            "Area",
            "Scatter",
            "Radar",
        ])
        self.chart_type.currentTextChanged.connect(self.on_chart_type_changed)
        form.addRow("Type:", self.chart_type)

        # New: clearer inputs
        self.dim_input = QtWidgets.QLineEdit()
        self.dim_input.setPlaceholderText("e.g. A2:A5 or A")
        dim_label = QtWidgets.QLabel("<b>Dim(X):</b>")
        dim_label.setStyleSheet("color: #d35400; font-size: 12pt;")
        form.addRow(dim_label, self.dim_input)

        self.values_input = QtWidgets.QLineEdit()
        self.values_input.setPlaceholderText("e.g. B2:B5 or B,C or (B,C)*(7:10)")
        values_label = QtWidgets.QLabel("<b>Values:</b>")
        values_label.setStyleSheet("color: #d35400; font-size: 12pt;")
        form.addRow(values_label, self.values_input)

        # Multi-value mode (updated based on chart type)
        self.multi_mode = QtWidgets.QComboBox()
        form.addRow("Multi mode:", self.multi_mode)

        layout.addLayout(form)

        # initialize multi_mode according to default chart type
        QtCore.QTimer.singleShot(0, lambda: self.on_chart_type_changed(self.chart_type.currentText()))

        # poll for active workbook/sheet and auto-refresh status
        self.poll_timer = QtCore.QTimer(self)
        self.poll_timer.setInterval(1500)  # ms
        self.poll_timer.timeout.connect(self.locate_active_sheet)
        self.poll_timer.start()
        # initial locate shortly after show
        QtCore.QTimer.singleShot(100, self.locate_active_sheet)

        btn_layout = QtWidgets.QHBoxLayout()
        btn_layout.setSpacing(5)  # Reduce spacing between buttons
        btn_layout.setContentsMargins(0, 0, 0, 0)  # Remove margins
        
        # Help button
        self.help_btn = QtWidgets.QPushButton("?")
        self.help_btn.setMaximumWidth(40)
        self.help_btn.clicked.connect(self.on_help)
        btn_layout.addWidget(self.help_btn)
        
        # Create / Change chart buttons only (workbook/sheet selection removed)
        self.create_btn = QtWidgets.QPushButton("Create ↵")
        self.create_btn.clicked.connect(self.on_create)
        self.create_btn.setShortcut(QtGui.QKeySequence(QtCore.Qt.Key.Key_Return))
        btn_layout.addWidget(self.create_btn)

        # Change (modify existing) button
        self.change_btn = QtWidgets.QPushButton("Change")
        self.change_btn.clicked.connect(self.on_change)
        btn_layout.addWidget(self.change_btn)

        layout.addLayout(btn_layout)

        # status / loading label
        self.status_label = QtWidgets.QLabel("")
        layout.addWidget(self.status_label)
        
        # Log/Notice board with toggle button
        log_header_layout = QtWidgets.QHBoxLayout()
        self.log_toggle_btn = QtWidgets.QPushButton("▶")
        self.log_toggle_btn.setMaximumWidth(30)
        self.log_toggle_btn.setFlat(True)
        self.log_toggle_btn.clicked.connect(self.toggle_log)
        log_font = QtGui.QFont("Times New Roman", 10)
        self.log_toggle_btn.setFont(log_font)
        log_header_layout.addWidget(self.log_toggle_btn)
        
        log_label = QtWidgets.QLabel("Log:")
        log_label.setFont(log_font)
        log_label.setStyleSheet("color: #555;")
        log_header_layout.addWidget(log_label)
        log_header_layout.addStretch()
        layout.addLayout(log_header_layout)
        
        self.log_board = QtWidgets.QTextEdit()
        self.log_board.setReadOnly(True)
        self.log_board.setFixedHeight(140)
        self.log_board.setFont(log_font)
        self.log_board.setStyleSheet("background-color: #f9f9f9; color: #555;")
        self.log_board.hide()  # Start with log hidden
        layout.addWidget(self.log_board)

        self.setLayout(layout)
        
        # Adjust window size to fit content (log is hidden initially)
        self.adjustSize()
        
        # Log initial message
        self._log("Peel Potato initialized. Ready to create charts!")

    def toggle_log(self):
        """Toggle log board visibility."""
        if self.log_board.isVisible():
            self.log_board.hide()
            self.log_toggle_btn.setText("▶")
        else:
            self.log_board.show()
            self.log_toggle_btn.setText("▼")
        # Let Qt automatically adjust the window height based on visible content
        QtCore.QTimer.singleShot(0, lambda: self.adjustSize())

    def locate_active_sheet(self):
        """Update the active_label and load_label with the currently focused workbook and sheet."""
        try:
            self.load_label.setText("Loading...")
            QtWidgets.QApplication.processEvents()
            app = xw.apps.active
            if app is None:
                self.active_label.setText("(no Excel instance)")
                self.load_label.setText("No Excel")
                return

            # Use COM ActiveWorkbook/ActiveSheet for reliable focus detection
            try:
                active_wb = app.api.ActiveWorkbook
                active_sh = app.api.ActiveSheet
                bname = getattr(active_wb, 'Name', None)
                sname = getattr(active_sh, 'Name', None)
            except Exception:
                # fallback: use first open book/sheet
                bname = None
                sname = None

            # Try to find friendly names via xw Book objects
            try:
                if bname:
                    for b in app.books:
                        if b.name == bname:
                            self._current_book = b
                            break
                else:
                    self._current_book = app.books[0] if app.books else None
            except Exception:
                self._current_book = None

            # Resolve sheet name
            try:
                if self._current_book is not None:
                    if sname and sname in [s.name for s in self._current_book.sheets]:
                        display = f"{self._current_book.name} — {sname}"
                        self.active_label.setText(display)
                        self.load_label.setText("Loaded")
                    else:
                        # fallback to first sheet name
                        first = self._current_book.sheets[0].name if self._current_book.sheets else '(no sheets)'
                        display = f"{self._current_book.name} — {first}"
                        self.active_label.setText(display)
                        self.load_label.setText("Loaded")
                else:
                    self.active_label.setText("(no workbook)")
                    self.load_label.setText("No workbook")
            except Exception:
                self.active_label.setText("(error)")
                self.load_label.setText("Error")

        except Exception:
            # Ensure label shows something even on error
            try:
                self.active_label.setText("(error)")
                self.load_label.setText("Error")
            except Exception:
                pass

    def get_selected_sheet(self):
        try:
            app = xw.apps.active
            if app is None:
                QtWidgets.QMessageBox.critical(self, "Error", "No Excel instance found.")
                return None
            # Always use the focused ActiveWorkbook/ActiveSheet
            try:
                active_wb_api = app.api.ActiveWorkbook
                active_sh_api = app.api.ActiveSheet
                bname = getattr(active_wb_api, 'Name', None)
                sname = getattr(active_sh_api, 'Name', None)
            except Exception:
                bname = None
                sname = None

            # find matching xw Book
            book = None
            if bname:
                for b in app.books:
                    if b.name == bname:
                        book = b
                        break
            if book is None and app.books:
                book = app.books[0]
            if book is None:
                QtWidgets.QMessageBox.critical(self, "Error", "No open workbook found.")
                return None

            try:
                if sname and sname in [s.name for s in book.sheets]:
                    return book.sheets[sname]
                return book.sheets[0]
            except Exception:
                return book.sheets[0]
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Could not get sheet: {e}")
            return None

    def _set_status(self, text, busy=False):
        try:
            self.status_label.setText(text)
            QtWidgets.QApplication.processEvents()
            if busy:
                QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
            else:
                QtWidgets.QApplication.restoreOverrideCursor()
        except Exception:
            pass
    
    def _log(self, message):
        """Append a message to the log board."""
        try:
            import datetime
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            self.log_board.append(f"[{timestamp}] {message}")
            # Auto-scroll to bottom
            self.log_board.verticalScrollBar().setValue(self.log_board.verticalScrollBar().maximum())
        except Exception:
            pass

    def _col_letter_to_index(self, letter):
        # Convert column letter(s) like 'A' or 'AA' to 1-based index
        letter = letter.strip().upper()
        if not letter.isalpha():
            return None
        result = 0
        for ch in letter:
            result = result * 26 + (ord(ch) - ord('A') + 1)
        return result

    def _parse_values_input(self, values_text, sheet, ref_rows=None):
        """Return list of xlwings Range objects for values input.
        values_text can be:
          - single range like B2:B5
          - range span like B:C
          - comma separated columns like B,C (if ref_rows provided the rows will be applied)
          - Cartesian product format like (B,E)*(2,7) or (B:E)*(2:7)
            - Use , for specific items: B,E means just B and E
            - Use : for continuous range: B:E means B, C, D, E
        This function avoids including the header row by default: it uses UsedRange.Row + 1 as data start.
        """
        values_text = values_text.strip()
        if not values_text:
            return []
        
        # Check for Cartesian product format: (cols) * (rows)
        import re
        cartesian_pattern = r'\(([^)]+)\)\s*\*\s*\(([^)]+)\)'
        cartesian_match = re.match(cartesian_pattern, values_text)
        
        if cartesian_match:
            # Parse Cartesian product format
            cols_str = cartesian_match.group(1).strip()
            rows_str = cartesian_match.group(2).strip()
            
            # Parse columns: supports both "B,E" (specific) and "B:E" (continuous range)
            cols = []
            if ':' in cols_str and ',' not in cols_str:
                # Continuous range like B:E
                col_parts = [c.strip().upper() for c in cols_str.split(':')]
                if len(col_parts) == 2:
                    start_col_idx = self._col_letter_to_index(col_parts[0])
                    end_col_idx = self._col_letter_to_index(col_parts[1])
                    if start_col_idx and end_col_idx:
                        # Generate all columns from start to end
                        for col_idx in range(start_col_idx, end_col_idx + 1):
                            # Convert index back to letter
                            col_letter = ''
                            idx = col_idx
                            while idx > 0:
                                idx, remainder = divmod(idx - 1, 26)
                                col_letter = chr(65 + remainder) + col_letter
                            cols.append(col_letter)
            else:
                # Comma-separated (may also have ranges): B,C or B:C,E
                for part in cols_str.split(','):
                    part = part.strip().upper()
                    if ':' in part:
                        # Range within comma-separated list
                        col_parts = [c.strip().upper() for c in part.split(':')]
                        if len(col_parts) == 2:
                            start_col_idx = self._col_letter_to_index(col_parts[0])
                            end_col_idx = self._col_letter_to_index(col_parts[1])
                            if start_col_idx and end_col_idx:
                                for col_idx in range(start_col_idx, end_col_idx + 1):
                                    col_letter = ''
                                    idx = col_idx
                                    while idx > 0:
                                        idx, remainder = divmod(idx - 1, 26)
                                        col_letter = chr(65 + remainder) + col_letter
                                    cols.append(col_letter)
                    else:
                        cols.append(part)
            
            # Parse rows: supports both "2,7" (specific) and "2:7" (continuous range)
            rows = []
            if ':' in rows_str and ',' not in rows_str:
                # Continuous range like 2:7
                row_parts = rows_str.split(':')
                try:
                    start_row = int(row_parts[0].strip())
                    end_row = int(row_parts[1].strip())
                    rows = list(range(start_row, end_row + 1))
                except ValueError:
                    pass
            else:
                # Comma-separated (may also have ranges): 2,7 or 2:5,7
                for part in rows_str.split(','):
                    part = part.strip()
                    if ':' in part:
                        # Range within comma-separated list
                        row_parts = part.split(':')
                        try:
                            start_row = int(row_parts[0].strip())
                            end_row = int(row_parts[1].strip())
                            rows.extend(range(start_row, end_row + 1))
                        except ValueError:
                            pass
                    else:
                        try:
                            rows.append(int(part))
                        except ValueError:
                            pass
            
            # Build ranges for each column-row combination
            if cols and rows:
                ranges = []
                for col in cols:
                    col_idx = self._col_letter_to_index(col)
                    if col_idx:
                        try:
                            # Create range for this column spanning all specified rows
                            start_row = min(rows)
                            end_row = max(rows)
                            ranges.append(sheet.range((start_row, col_idx), (end_row, col_idx)))
                        except Exception:
                            pass
                return ranges
        
        # Regular parsing (backward compatible)
        parts = [p.strip() for p in values_text.split(',') if p.strip()]
        ranges = []
        # determine used rows for the sheet when available
        try:
            used = sheet.api.UsedRange
            header_row = int(used.Row)
            data_start = header_row + 1
            used_end = int(used.Row + used.Rows.Count - 1)
        except Exception:
            header_row = None
            data_start = None
            used_end = None

        for p in parts:
            # direct explicit range with row numbers stays as-is (e.g. B2:B5)
            if any(ch.isdigit() for ch in p):
                try:
                    ranges.append(sheet.range(p))
                except Exception:
                    pass
                continue

            # column span like B:C or single-column like B or C
            if ':' in p:
                left, right = [x.strip() for x in p.split(':', 1)]
                # if both sides are column letters (no digits) then constrain to data rows (exclude header)
                if left.isalpha() and right.isalpha() and data_start is not None and used_end is not None:
                    left_idx = self._col_letter_to_index(left)
                    right_idx = self._col_letter_to_index(right)
                    if left_idx and right_idx and left_idx <= right_idx:
                        for col in range(left_idx, right_idx + 1):
                            try:
                                ranges.append(sheet.range((data_start, col), (used_end, col)))
                            except Exception:
                                pass
                    continue
                else:
                    # fallback: treat as general range
                    try:
                        ranges.append(sheet.range(p))
                    except Exception:
                        pass
                    continue

            # single column letter like 'B'
            col_idx = self._col_letter_to_index(p)
            if col_idx is None:
                continue
            if ref_rows:
                start_row, end_row = ref_rows
                ranges.append(sheet.range((start_row, col_idx), (end_row, col_idx)))
            elif data_start is not None and used_end is not None:
                # use data rows only, excluding header
                ranges.append(sheet.range((data_start, col_idx), (used_end, col_idx)))
            else:
                # fallback: pick a long-ish span starting at row 2 to avoid header at row1
                ranges.append(sheet.range((2, col_idx), (10000, col_idx)))
        return ranges

    def _compute_source_block(self, sheet, ranges_list):
        """Given list of xlwings Range objects, return a combined source COM Range covering them all."""
        api = sheet.api
        min_row = None
        max_row = None
        min_col = None
        max_col = None
        for r in ranges_list:
            try:
                ra = r.api
                r_row = ra.Row
                r_rows = ra.Rows.Count
                r_col = ra.Column
                r_cols = ra.Columns.Count
                r_min_row = r_row
                r_max_row = r_row + r_rows - 1
                r_min_col = r_col
                r_max_col = r_col + r_cols - 1
                if min_row is None or r_min_row < min_row:
                    min_row = r_min_row
                if max_row is None or r_max_row > max_row:
                    max_row = r_max_row
                if min_col is None or r_min_col < min_col:
                    min_col = r_min_col
                if max_col is None or r_max_col > max_col:
                    max_col = r_max_col
            except Exception:
                pass
        if min_row is None:
            return None
        return api.Range(api.Cells(min_row, min_col), api.Cells(max_row, max_col))

    def on_create(self):
        self.create_chart(preview=False)

    def create_chart(self, preview=True, modify=False):
        try:
            chart_type = self.chart_type.currentText()
            dim_text = self.dim_input.text().strip()
            values_text = self.values_input.text().strip()

            sheet = self.get_selected_sheet()
            if sheet is None:
                return

            # show busy status and cursor while chart is created
            self._set_status("Creating chart…", busy=True)
        except Exception as e:
            self._show_potato_error("Chart creation", e)
            self._set_status("", busy=False)
            return
        
        try:
            # Validate inputs
            ct_lower = chart_type.lower()
            if 'scatter' in ct_lower:
                if not dim_text or not values_text:
                    QtWidgets.QMessageBox.warning(self, "Input required", "Scatter needs Dim (X) and Values (Y).")
                    return
            else:
                # default: require Dim and Values
                if not dim_text or not values_text:
                    QtWidgets.QMessageBox.warning(self, "Input required", "Please fill Dim and Values.")
                    return

            # COM objects
            sht_api = sheet.api
            # place chart at fixed position
            left = 50
            top = 20
            width = 520
            height = 320

            # Ensure win32 constants available
            if xlconst is None:
                from win32com.client import constants as xlconst_local
                _xl = xlconst_local
            else:
                _xl = xlconst

            # Create or reuse a chartobject
            chart = None
            if modify and hasattr(self, '_last_chart') and self._last_chart is not None:
                # Validate that the chart still exists
                try:
                    _ = self._last_chart.ChartType
                    chart = self._last_chart
                except Exception:
                    # Chart no longer exists, set to None to create new
                    chart = None
                    self._last_chart = None
            
            if chart is None:
                chart_objects = sht_api.ChartObjects()
                chart_obj = chart_objects.Add(left, top, width, height)
                chart = chart_obj.Chart
                # remember last created chart for later modifications
                self._last_chart = chart

            # Determine Excel chart constant and set type before adding series
            chart_const = self._chart_constant_for(chart_type, self.multi_mode.currentText(), _xl)
            self._log(f"Creating {chart_type} chart with {self.multi_mode.currentText()} mode")
            try:
                chart.ChartType = chart_const
            except Exception:
                pass

            # Build ranges for dim and values (support multi-values)
            # dim_range
            self._log(f"Parsing Dim input: {dim_text}")
            try:
                dim_range = sheet.range(dim_text) if (any(ch.isdigit() for ch in dim_text) or ':' in dim_text) else None
            except Exception:
                dim_range = None
            ref_rows = None
            if dim_range is not None:
                dra = dim_range.api
                ref_rows = (dra.Row, dra.Row + dra.Rows.Count - 1)

            self._log(f"Parsing Values input: {values_text}")
            value_ranges = self._parse_values_input(values_text, sheet, ref_rows=ref_rows)
            
            # Log detected value names
            if value_ranges:
                self._log(f"Found {len(value_ranges)} value range(s)")

            # If dim was provided as a column letter (e.g. 'A') build dim_range from value rows or used range
            if dim_range is None and dim_text and dim_text.strip().isalpha():
                # infer rows from first value range if available
                try:
                    if value_ranges:
                        vr0 = value_ranges[0].api
                        ref_rows = (vr0.Row, vr0.Row + vr0.Rows.Count - 1)
                    else:
                        used = sheet.api.UsedRange
                        ref_rows = (used.Row, used.Row + used.Rows.Count - 1)
                except Exception:
                    ref_rows = None
                if ref_rows:
                    col_idx = self._col_letter_to_index(dim_text.strip())
                    if col_idx:
                        try:
                            dim_range = sheet.range((ref_rows[0], col_idx), (ref_rows[1], col_idx))
                        except Exception:
                            dim_range = None

            # Now detect and log all names (after dim_range is fully built)
            if dim_range:
                try:
                    dim_name = peel_potato_prettify.reset_title_name(dim_range)
                    if dim_name:
                        self._log(f"✓ Dim name: <b>{dim_name}</b>")
                except Exception:
                    pass
            
            if value_ranges:
                for idx, vr in enumerate(value_ranges):
                    try:
                        value_name = peel_potato_prettify.reset_title_name(vr)
                        if value_name:
                            self._log(f"  ✓ Value {idx+1}: <b>{value_name}</b>")
                    except Exception:
                        pass

            # ...histogram support removed. Other chart types handled below.

            # Scatter: use first two ranges as X and Y
            if chart_type == "Scatter":
                if len(value_ranges) == 0:
                    QtWidgets.QMessageBox.warning(self, "Input error", "No value ranges parsed for scatter.")
                    return
                x_range = dim_range.api if dim_range is not None else value_ranges[0].api
                y_range = value_ranges[0].api if dim_range is not None else (value_ranges[1].api if len(value_ranges) > 1 else None)
                if y_range is None:
                    QtWidgets.QMessageBox.warning(self, "Input error", "Scatter needs two columns (X and Y).")
                    return
                series = chart.SeriesCollection().NewSeries()
                series.XValues = x_range
                series.Values = y_range
                chart.ChartType = _xl.xlXYScatter

            else:
                # For line/bar/column/area/pie/etc. build series explicitly so Dim is used as category axis
                if not value_ranges:
                    QtWidgets.QMessageBox.warning(self, "Input error", "No value ranges parsed for chart.")
                    return

                # Try to determine header row (assume header is the row immediately above data rows if ref_rows present)
                header_row = None
                if ref_rows:
                    header_row = max(1, ref_rows[0] - 1)
                else:
                    try:
                        used = sheet.api.UsedRange
                        header_row = used.Row
                    except Exception:
                        header_row = 1

                # If modifying an existing chart, clear existing series first
                try:
                    if modify:
                        while chart.SeriesCollection().Count > 0:
                            chart.SeriesCollection(1).Delete()
                except Exception:
                    pass

                # Create a series per value_range (Pie will use only the first)
                for idx, vr in enumerate(value_ranges):
                    try:
                        # For Pie charts only use first value column
                        if 'pie' in chart_type.lower() and idx > 0:
                            break
                        s = chart.SeriesCollection().NewSeries()
                        s.Values = vr.api
                        if dim_range is not None:
                            s.XValues = dim_range.api
                        # Try to set a sensible series name from header row
                        try:
                            name_cell = sheet.api.Cells(header_row, vr.api.Column)
                            name_val = name_cell.Value
                            if name_val is not None:
                                s.Name = str(name_val)
                        except Exception:
                            pass
                    except Exception:
                        pass

                # determine chart subtype
                chart.ChartType = self._chart_constant_for(chart_type, self.multi_mode.currentText(), _xl)

            # Give the chart a title
            chart.HasTitle = True
            chart.ChartTitle.Text = f"{chart_type} — Peel Potato"

            # Apply default formatting with dim and value ranges for auto-naming
            names_result = peel_potato_prettify.apply_chart_formatting(chart, dim_range, value_ranges)
            
            # Display the detected names prominently
            if names_result and names_result[0] and names_result[1]:
                dim_name, value_names = names_result
                self._log(f"📊 Chart title set: <b>{value_names[0]} by {dim_name}</b>")
            
            # Log completion
            action = "modified" if modify else "created"
            self._log(f"Chart {action} successfully! Applied formatting.")

            # remember last chart even when modifying
            try:
                self._last_chart = chart
            except Exception:
                pass

            # No separate sheet required on Create: preview already places the chart on the sheet
            # Keep the created chart embedded where preview showed it.

        except Exception as e:
            self._show_potato_error("Chart operation", e)
            self._set_status("", busy=False)
        finally:
            # clear status and restore cursor
            self._set_status("", busy=False)

    def on_chart_type_changed(self, text):
        # Update multi_mode options to match chart type
        t = text.lower()
        self.multi_mode.clear()
        if 'line' in t:
            self.multi_mode.addItems(["Standard", "Stacked", "100% Stacked"]) 
            self.multi_mode.setEnabled(True)
        elif 'column' in t:
            self.multi_mode.addItems(["Clustered", "Stacked", "100% Stacked"]) 
            self.multi_mode.setEnabled(True)
        elif 'bar' in t:
            self.multi_mode.addItems(["Clustered", "Stacked", "100% Stacked"]) 
            self.multi_mode.setEnabled(True)
        elif 'area' in t:
            self.multi_mode.addItems(["Stacked", "100% Stacked"]) 
            self.multi_mode.setEnabled(True)
        elif 'pie' in t:
            self.multi_mode.addItems(["Pie", "Pie of", "Doughnut"]) 
            self.multi_mode.setEnabled(True)
        elif 'scatter' in t:
            self.multi_mode.addItems(["Scatter"]) 
            self.multi_mode.setEnabled(False)
        elif 'radar' in t:
            self.multi_mode.addItems(["Radar"]) 
            self.multi_mode.setEnabled(False)
        # histogram removed from modes
        else:
            self.multi_mode.setEnabled(False)

    def on_help(self):
        """Show help dialog with usage information loaded from help.html."""
        try:
            # Load help.html from the same directory as the script
            help_path = os.path.join(os.path.dirname(__file__), 'help.html')
            
            # Read the HTML content
            with open(help_path, 'r', encoding='utf-8') as f:
                help_text = f.read()
            
            msg = QtWidgets.QMessageBox(self)
            msg.setWindowTitle("Help — Peel Potato")
            msg.setTextFormat(QtCore.Qt.TextFormat.RichText)
            msg.setText(help_text)
            msg.setIcon(QtWidgets.QMessageBox.Icon.Information)
            msg.setStandardButtons(QtWidgets.QMessageBox.StandardButton.Ok)
            
            # Set font for the message box
            msg_font = QtGui.QFont("Microsoft YaHei UI", 12)
            msg.setFont(msg_font)
            
            msg.exec()
        except Exception as e:
            # Fallback if help.html cannot be loaded
            QtWidgets.QMessageBox.information(
                self,
                "Help — Peel Potato",
                "Help file not found. Please ensure help.html is in the application directory."
            )
    
    def on_change(self):
        # Modify the currently selected chart in Excel
        try:
            sheet = self.get_selected_sheet()
            if sheet is None:
                return
            
            # Try to get the selected chart
            selected_chart = None
            try:
                app = xw.apps.active
                if app is not None:
                    selection = app.api.Selection
                    
                    # Try different ways to get the chart
                    # 1. Selection might be a ChartObject directly
                    try:
                        if hasattr(selection, 'Chart'):
                            selected_chart = selection.Chart
                    except Exception:
                        pass
                    
                    # 2. Selection might be a chart area (when clicking inside the chart)
                    if selected_chart is None:
                        try:
                            if hasattr(selection, 'Parent') and hasattr(selection.Parent, 'ChartType'):
                                selected_chart = selection.Parent
                        except Exception:
                            pass
                    
                    # 3. Try ActiveChart as fallback
                    if selected_chart is None:
                        try:
                            active_chart = app.api.ActiveChart
                            if active_chart is not None:
                                selected_chart = active_chart
                        except Exception:
                            pass
                    
                    if selected_chart is not None:
                        self._last_chart = selected_chart
            except Exception:
                pass
            
            if selected_chart is None:
                self._log("No chart selected in Excel.")
                QtWidgets.QMessageBox.warning(
                    self, 
                    "No chart selected", 
                    "Please select a chart in Excel to modify, or use Create to create a new one."
                )
                return
            
            self._log("Selected chart detected, modifying...")
            # Modify the selected chart
            self.create_chart(preview=False, modify=True)
            
        except Exception as e:
            self._show_potato_error("Change chart", e)

    def _show_potato_error(self, operation, error):
        """Show a potato-themed error message."""
        potato_messages = [
            "🥔 Oops! The potato got mashed!",
            "🥔 The potato peeler hit a snag!",
            "🥔 Potato malfunction detected!",
            "🥔 The potato needs a moment...",
            "🥔 Chart potato overcooked!"
        ]
        import random
        title = random.choice(potato_messages)
        message = f"{operation} encountered an issue:\n\n{str(error)[:200]}"
        
        try:
            QtWidgets.QMessageBox.warning(self, title, message)
        except Exception:
            pass
    
    def _chart_constant_for(self, chart_text, mode_text, _xl):
        # Map chart type + mode to Excel ChartType constant
        ct = chart_text.lower()
        m = mode_text.lower() if mode_text else ''
        try:
            if 'line' in ct:
                if 'stacked' in m and '100' in m:
                    return getattr(_xl, 'xlLineStacked100', getattr(_xl, 'xlLineStacked', _xl.xlLine))
                if 'stacked' in m:
                    return getattr(_xl, 'xlLineStacked', _xl.xlLine)
                return getattr(_xl, 'xlLine', getattr(_xl, 'xlLineMarkers', _xl.xlLine))
            if 'column' in ct:
                if '100' in m:
                    return getattr(_xl, 'xlColumnStacked100', getattr(_xl, 'xlColumnStacked', _xl.xlColumnClustered))
                if 'stack' in m:
                    return getattr(_xl, 'xlColumnStacked', _xl.xlColumnClustered)
                return getattr(_xl, 'xlColumnClustered', _xl.xlColumnClustered)
            if 'bar' in ct:
                if '100' in m:
                    return getattr(_xl, 'xlBarStacked100', getattr(_xl, 'xlBarStacked', _xl.xlBarClustered))
                if 'stack' in m:
                    return getattr(_xl, 'xlBarStacked', _xl.xlBarClustered)
                return getattr(_xl, 'xlBarClustered', _xl.xlBarClustered)
            if 'area' in ct:
                if '100' in m:
                    return getattr(_xl, 'xlAreaStacked100', getattr(_xl, 'xlAreaStacked', _xl.xlArea))
                if 'stack' in m:
                    return getattr(_xl, 'xlAreaStacked', _xl.xlArea)
                return getattr(_xl, 'xlArea', _xl.xlArea)
            if 'pie' in ct:
                if 'doughnut' in m:
                    return getattr(_xl, 'xlDoughnut', _xl.xlPie)
                if 'pie of' in m:
                    return getattr(_xl, 'xlPieOfPie', _xl.xlPie)
                return getattr(_xl, 'xlPie', _xl.xlPie)
            if 'scatter' in ct:
                return getattr(_xl, 'xlXYScatter', _xl.xlXYScatter)
            if 'radar' in ct:
                return getattr(_xl, 'xlRadar', _xl.xlRadar)
            # histogram removed from mapping
        except Exception:
            pass
        # fallback
        return _xl.xlColumnClustered


def main():
    """Main entry point with crash recovery."""
    max_restarts = 3
    restart_count = 0
    
    while restart_count < max_restarts:
        try:
            app = QtWidgets.QApplication(sys.argv)
            w = PeelPotato()
            w.show()
            sys.exit(app.exec())
        except Exception as e:
            restart_count += 1
            print(f"Crash detected (attempt {restart_count}/{max_restarts}): {e}")
            
            # Show error dialog if possible
            try:
                import random
                potato_titles = [
                    "🥔 Potato Emergency!",
                    "🥔 The Potato Tumbled!",
                    "🥔 Potato Overload!"
                ]
                msg = QtWidgets.QApplication.instance()
                if msg is None:
                    msg = QtWidgets.QApplication(sys.argv)
                
                error_box = QtWidgets.QMessageBox()
                error_box.setWindowTitle(random.choice(potato_titles))
                error_box.setIcon(QtWidgets.QMessageBox.Icon.Warning)
                error_box.setText(f"Peel Potato encountered an error and will restart.\n\nAttempt {restart_count}/{max_restarts}")
                error_box.setDetailedText(str(e))
                error_box.setStandardButtons(QtWidgets.QMessageBox.StandardButton.Ok)
                error_box.exec()
                
                # Clean up before restart
                try:
                    QtWidgets.QApplication.instance().quit()
                except Exception:
                    pass
            except Exception:
                pass
            
            if restart_count >= max_restarts:
                print("Max restart attempts reached. Exiting.")
                break
            
            # Brief pause before restart
            time.sleep(1)
    
    sys.exit(1)


if __name__ == '__main__':
    main()
