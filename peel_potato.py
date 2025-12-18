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
        self.resize(420, 220)

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

        form = QtWidgets.QFormLayout()

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
        form.addRow("Chart type:", self.chart_type)

        # Active workbook/sheet notification (read-only)
        self.active_label = QtWidgets.QLabel("(no Excel detected)")
        self.active_label.setStyleSheet("font-weight: bold")
        form.addRow("Active workbook/sheet:", self.active_label)

        # Loading / status indicator
        self.load_label = QtWidgets.QLabel("")
        self.load_label.setStyleSheet("color: #666; font-size: 11px")
        form.addRow("Status:", self.load_label)

        # New: clearer inputs
        self.dim_input = QtWidgets.QLineEdit()
        self.dim_input.setPlaceholderText("e.g. A2:A5 or A (header will be used) — category / X")
        form.addRow("Dim (category / X):", self.dim_input)

        self.values_input = QtWidgets.QLineEdit()
        self.values_input.setPlaceholderText("e.g. B2:B5 or B,C or B:C — support multiple values")
        form.addRow("Values:", self.values_input)

        # Multi-value mode (updated based on chart type)
        self.multi_mode = QtWidgets.QComboBox()
        form.addRow("Multi-series mode:", self.multi_mode)

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
        # Create / Change chart buttons only (workbook/sheet selection removed)
        self.create_btn = QtWidgets.QPushButton("Create Chart")
        self.create_btn.clicked.connect(self.on_create)
        btn_layout.addWidget(self.create_btn)

        # Change (modify existing) button
        self.change_btn = QtWidgets.QPushButton("Change Chart")
        self.change_btn.clicked.connect(self.on_change)
        btn_layout.addWidget(self.change_btn)

        layout.addLayout(btn_layout)

        # status / loading label
        self.status_label = QtWidgets.QLabel("")
        self.status_label.setStyleSheet("color: #333; font-size: 11px")
        layout.addWidget(self.status_label)

        self.setLayout(layout)

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
        This function avoids including the header row by default: it uses UsedRange.Row + 1 as data start.
        """
        values_text = values_text.strip()
        if not values_text:
            return []
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

    def on_preview(self):
        self.create_chart(preview=True)

    def on_create(self):
        self.create_chart(preview=False)

    def on_create_pivot(self):
        self.create_pivot()

    def create_pivot(self):
        sheet = self.get_selected_sheet()
        if sheet is None:
            return
        # show busy status
        self._set_status("Creating pivot…", busy=True)
        dim_text = self.dim_input.text().strip()
        values_text = self.values_input.text().strip()
        if not dim_text or not values_text:
            QtWidgets.QMessageBox.warning(self, "Input required", "Pivot needs Dim and Values inputs.")
            return
        try:
            # parse dim to get rows
            dim_range = sheet.range(dim_text) if (any(ch.isdigit() for ch in dim_text) or ':' in dim_text) else None
        except Exception:
            dim_range = None
        # If dim_range has rows, get start/end row
        ref_rows = None
        if dim_range is not None:
            ra = dim_range.api
            ref_rows = (ra.Row, ra.Row + ra.Rows.Count - 1)

        value_ranges = self._parse_values_input(values_text, sheet, ref_rows=ref_rows)
        if not value_ranges:
            QtWidgets.QMessageBox.warning(self, "Input error", "Could not parse Values input.")
            return

        # Build source block from dim + value ranges (include headers row)
        all_ranges = []
        if dim_range is not None:
            all_ranges.append(dim_range)
        else:
            # if dim_text is a column letter like A, try to construct via ref_rows
            if ref_rows:
                col_letter = dim_text.strip()
                col_idx = self._col_letter_to_index(col_letter)
                if col_idx:
                    all_ranges.append(sheet.range((ref_rows[0], col_idx), (ref_rows[1], col_idx)))
        all_ranges.extend(value_ranges)

        source_com_range = self._compute_source_block(sheet, all_ranges)
        if source_com_range is None:
            QtWidgets.QMessageBox.warning(self, "Error", "Could not determine source range for pivot.")
            return

        # Expand pivot source to include all contiguous used columns to the right (auto-select all columns)
        try:
            used = sheet.api.UsedRange
            used_last_row = used.Row + used.Rows.Count - 1
            used_last_col = used.Column + used.Columns.Count - 1
            header_row = source_com_range.Row
            first_col = source_com_range.Column
            # build expanded source to include all used cols up to used_last_col
            source_com_range = sheet.api.Range(sheet.api.Cells(header_row, first_col), sheet.api.Cells(used_last_row, used_last_col))
        except Exception:
            pass

        # Determine destination: far right of the source block + 2 columns
        src_cols = source_com_range.Columns.Count
        src_rows = source_com_range.Rows.Count
        src_first_col = source_com_range.Column
        src_last_col = src_first_col + src_cols - 1
        dest_col = src_last_col + 2
        dest_cell = sheet.api.Cells(1, dest_col)

        # Create pivot cache and table
        wb_api = self._current_book.api if hasattr(self, '_current_book') else sheet.book.api
        if xlconst is None:
            from win32com.client import constants as xlconst_local
            _xl = xlconst_local
        else:
            _xl = xlconst

        try:
            pc = wb_api.PivotCaches().Create(SourceType=_xl.xlDatabase, SourceData=source_com_range)
            name = f"PeelPotatoPivot_{int(time.time())}"
            pt = pc.CreatePivotTable(TableDestination=dest_cell, TableName=name)

            # Add fields: assume first column is Dim, remaining are Values
            header_row = source_com_range.Row
            first_col = source_com_range.Column
            last_col = src_last_col
            # Row field:
            dim_field_name = sheet.api.Cells(header_row, first_col).Value
            try:
                pt.PivotFields(dim_field_name).Orientation = _xl.xlRowField
            except Exception:
                pass

            # Data fields (default aggregation: Sum)
            func = _xl.xlSum
            for col in range(first_col + 1, last_col + 1):
                fld_name = sheet.api.Cells(header_row, col).Value
                try:
                    pt.AddDataField(pt.PivotFields(fld_name), f"Sum of {fld_name}", func)
                except Exception:
                    try:
                        # fallback: set orientation and use default
                        pt.PivotFields(fld_name).Orientation = _xl.xlDataField
                    except Exception:
                        pass

            QtWidgets.QMessageBox.information(self, "Pivot created", f"Pivot table '{name}' created at column {dest_col}.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Pivot error", str(e))
        finally:
            # clear busy status
            self._set_status("", busy=False)

    def create_chart(self, preview=True, modify=False):
        chart_type = self.chart_type.currentText()
        dim_text = self.dim_input.text().strip()
        values_text = self.values_input.text().strip()

        sheet = self.get_selected_sheet()
        if sheet is None:
            return

        # show busy status and cursor while chart is created
        self._set_status("Creating chart…", busy=True)
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
            try:
                chart.ChartType = chart_const
            except Exception:
                pass

            # Build ranges for dim and values (support multi-values)
            # dim_range
            try:
                dim_range = sheet.range(dim_text) if (any(ch.isdigit() for ch in dim_text) or ':' in dim_text) else None
            except Exception:
                dim_range = None
            ref_rows = None
            if dim_range is not None:
                dra = dim_range.api
                ref_rows = (dra.Row, dra.Row + dra.Rows.Count - 1)

            value_ranges = self._parse_values_input(values_text, sheet, ref_rows=ref_rows)

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

            # Apply default formatting
            peel_potato_prettify.apply_chart_formatting(chart)

            # remember last chart even when modifying
            try:
                self._last_chart = chart
            except Exception:
                pass

            # No separate sheet required on Create: preview already places the chart on the sheet
            # Keep the created chart embedded where preview showed it.

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Chart error", str(e))
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
                QtWidgets.QMessageBox.warning(
                    self, 
                    "No chart selected", 
                    "Please select a chart in Excel to modify, or use Create Chart to create a new one."
                )
                return
            
            # Modify the selected chart
            self.create_chart(preview=False, modify=True)
            
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to modify chart: {e}")

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
    app = QtWidgets.QApplication(sys.argv)
    w = PeelPotato()
    w.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
