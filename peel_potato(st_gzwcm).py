"""
Peel Potato V3.2.6(gzwcm) - FastBI for Excel

Version: V3.2.6 (gzwcm batch)
Purpose: Excel data processing and chart creation tool with specialized gzwcm utilities
Features:
  - FastBI charting: Quick chart creation/modification with intuitive range syntax
  - gzwcm_auto tools:
    * info: Match and enrich employee data from emp.xlsx by emp_nm or emp_id
    * slc: Select columns per dict.xlsx mapping, keep default columns (data_dt, grp, emp_id, emp_nm)
    * sum: Filter by emp.xlsx (emp_nm), sum by emp_id, group by grp with totals
Architecture: 
  - Modular design with separate controller, parser, builder, and adapter layers
  - String-based emp_id formatting (8-digit with leading zeros)
  - Pre-formatted Excel text columns to preserve emp_id integrity
"""
# Version and batch information
VERSION = "V3.2.6"
BATCH = "gzwcm"

import sys
import os
import datetime
import random
from PyQt6 import QtWidgets, QtCore, QtGui

# Import our refactored services
from peel_potato_adapter import ExcelAdapter
from peel_potato_parser import RangeParser
from peel_potato_chart_builder import ChartBuilder
from peel_potato_controller import ChartController

# Import ST_GZWCM utilities
from st_gzwcm_info import info
from st_gzwcm_sum import sum as gzwcm_sum
from st_gzwcm_slc import slc


class PeelPotatoWindow(QtWidgets.QWidget):
    """Main UI window for Peel Potato - thin UI layer only."""
    
    def __init__(self): 
        super().__init__()
        
        # Initialize controller with services (Dependency Injection)
        self.controller = ChartController(
            excel_adapter=ExcelAdapter(),
            range_parser=RangeParser(),
            chart_builder=ChartBuilder()
        )
        
        self._setup_ui()
        self._setup_polling()
        
        # Log initial message
        self._log(f"Peel Potato {VERSION}({BATCH}) initialized. Ready to create charts!")
    
    def _setup_ui(self):
        """Setup all UI widgets and layout."""
        self.setWindowTitle(f"Peel Potato {VERSION}({BATCH}) — FastBI for Excel")
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowType.WindowStaysOnTopHint)
        self.setFixedWidth(420)
        
        # Set base font
        base_font = QtGui.QFont("Microsoft YaHei UI", 12)
        self.setFont(base_font)
        
        # Set window icon
        self._setup_icon()
        
        # Main layout
        layout = QtWidgets.QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        
        # Form layout for inputs
        form = QtWidgets.QFormLayout()
        
        # Active workbook/sheet notification
        self.active_label = QtWidgets.QLabel("(no Excel detected)")
        active_font = QtGui.QFont("Microsoft YaHei UI", 10)
        self.active_label.setFont(active_font)
        form.addRow("Active at:", self.active_label)
        
        # Status indicator
        self.load_label = QtWidgets.QLabel("")
        self.load_label.setFont(active_font)
        form.addRow("Status:", self.load_label)
        
        # Chart type selector
        self.chart_type = QtWidgets.QComboBox()
        self.chart_type.addItems([
            "Line", "Bar (horizontal)", "Column (vertical)",
            "Pie", "Area", "Scatter", "Radar"
        ])
        self.chart_type.currentTextChanged.connect(self._on_chart_type_changed)
        form.addRow("Type:", self.chart_type)
        
        # Dimension input
        self.dim_input = QtWidgets.QLineEdit()
        self.dim_input.setPlaceholderText("e.g. A2:A5 or A")
        dim_label = QtWidgets.QLabel("<b>Dim(X):</b>")
        dim_label.setStyleSheet("color: #d35400; font-size: 12pt;")
        form.addRow(dim_label, self.dim_input)
        
        # Values input
        self.values_input = QtWidgets.QLineEdit()
        self.values_input.setPlaceholderText("e.g. B2:B5 or B,C or (B,C)*(7:10)")
        values_label = QtWidgets.QLabel("<b>Values:</b>")
        values_label.setStyleSheet("color: #d35400; font-size: 12pt;")
        form.addRow(values_label, self.values_input)
        
        # Multi-value mode
        self.multi_mode = QtWidgets.QComboBox()
        form.addRow("Multi mode:", self.multi_mode)
        
        layout.addLayout(form)
        
        # Initialize multi_mode for default chart type
        QtCore.QTimer.singleShot(0, lambda: self._on_chart_type_changed(self.chart_type.currentText()))
        
        # Button layout
        btn_layout = QtWidgets.QHBoxLayout()
        btn_layout.setSpacing(5)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        
        # Help button
        self.help_btn = QtWidgets.QPushButton("?")
        self.help_btn.setMaximumWidth(40)
        self.help_btn.clicked.connect(self._on_help)
        btn_layout.addWidget(self.help_btn)
        
        # Create button
        self.create_btn = QtWidgets.QPushButton("Create ↵")
        self.create_btn.clicked.connect(self._on_create)
        self.create_btn.setShortcut(QtGui.QKeySequence(QtCore.Qt.Key.Key_Return))
        btn_layout.addWidget(self.create_btn)
        
        # Change button
        self.change_btn = QtWidgets.QPushButton("Change")
        self.change_btn.clicked.connect(self._on_change)
        btn_layout.addWidget(self.change_btn)
        
        layout.addLayout(btn_layout)
        
        # GZWCM Auto Tools section
        gzwcm_label = QtWidgets.QLabel("<b>gzwcm_auto tools</b>")
        gzwcm_label.setStyleSheet("color: #2980b9; font-size: 11pt;")
        gzwcm_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(gzwcm_label)
        
        # 4 buttons in one line: info, slc, sum, cat
        gzwcm_btn_layout = QtWidgets.QHBoxLayout()
        gzwcm_btn_layout.setSpacing(5)
        gzwcm_btn_layout.setContentsMargins(0, 0, 0, 0)
        
        self.info_btn = QtWidgets.QPushButton("info")
        self.info_btn.clicked.connect(self._on_info)
        gzwcm_btn_layout.addWidget(self.info_btn)
        
        self.slc_btn = QtWidgets.QPushButton("slc")
        self.slc_btn.clicked.connect(self._on_slc)
        gzwcm_btn_layout.addWidget(self.slc_btn)
        
        self.sum_btn = QtWidgets.QPushButton("sum")
        self.sum_btn.clicked.connect(self._on_sum)
        gzwcm_btn_layout.addWidget(self.sum_btn)
        
        self.cat_btn = QtWidgets.QPushButton("cat")
        self.cat_btn.clicked.connect(self._on_cat)
        self.cat_btn.setEnabled(False)  # TODO: Implement cat functionality
        gzwcm_btn_layout.addWidget(self.cat_btn)
        
        layout.addLayout(gzwcm_btn_layout)
        
        # Status label
        self.status_label = QtWidgets.QLabel("")
        layout.addWidget(self.status_label)
        
        # Log/Notice board with toggle
        log_header_layout = QtWidgets.QHBoxLayout()
        self.log_toggle_btn = QtWidgets.QPushButton("▶")
        self.log_toggle_btn.setMaximumWidth(30)
        self.log_toggle_btn.setFlat(True)
        self.log_toggle_btn.clicked.connect(self._toggle_log)
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
        self.log_board.hide()
        layout.addWidget(self.log_board)
        
        self.setLayout(layout)
        self.adjustSize()
    
    def _setup_icon(self):
        """Setup window icon."""
        try:
            icon_dir = os.path.dirname(__file__)
            candidates = [
                os.path.join(icon_dir, 'media', 'icon_app.ico'),
                os.path.join(icon_dir, 'media', 'icon_exe.ico'),
            ]
            for icon_path in candidates:
                if os.path.exists(icon_path):
                    try:
                        pix = QtGui.QPixmap(icon_path)
                        if not pix.isNull():
                            pix = pix.scaled(60, 60, QtCore.Qt.AspectRatioMode.KeepAspectRatio,
                                           QtCore.Qt.TransformationMode.SmoothTransformation)
                            self.setWindowIcon(QtGui.QIcon(pix))
                        else:
                            self.setWindowIcon(QtGui.QIcon(icon_path))
                        break
                    except Exception:
                        pass
        except Exception:
            pass
    
    def _setup_polling(self):
        """Setup polling timer for active Excel sheet."""
        self.poll_timer = QtCore.QTimer(self)
        self.poll_timer.setInterval(5000)  # 5 seconds
        self.poll_timer.timeout.connect(self._poll_active_sheet)
        self.poll_timer.start()
        
        # Initial poll shortly after startup
        QtCore.QTimer.singleShot(100, self._poll_active_sheet)
    
    def _poll_active_sheet(self):
        """Poll for active Excel workbook and sheet."""
        try:
            self.load_label.setText("Loading...")
            QtWidgets.QApplication.processEvents()
            
            workbook_name, sheet_name = self.controller.get_active_sheet_info()
            
            if workbook_name and sheet_name:
                self.active_label.setText(f"{workbook_name} → {sheet_name}")
                self.load_label.setText("✓ Ready")
            else:
                self.active_label.setText("(no Excel detected)")
                self.load_label.setText("(waiting)")
        except Exception:
            self.active_label.setText("(error detecting Excel)")
            self.load_label.setText("(error)")
    
    def _on_create(self):
        """Handle Create button click."""
        self._set_status("Creating chart…", busy=True)
        self._log("Starting chart creation...")
        
        try:
            result = self.controller.create_chart(
                dim_text=self.dim_input.text(),
                values_text=self.values_input.text(),
                chart_type=self.chart_type.currentText(),
                multi_mode=self.multi_mode.currentText()
            )
            
            self._handle_chart_result(result, "create")
            
        except Exception as e:
            self._show_error("Chart Creation", str(e))
        finally:
            self._set_status("", busy=False)
    
    def _on_change(self):
        """Handle Change button click."""
        self._set_status("Modifying chart…", busy=True)
        self._log("Starting chart modification...")
        
        try:
            result = self.controller.modify_chart(
                dim_text=self.dim_input.text(),
                values_text=self.values_input.text(),
                chart_type=self.chart_type.currentText(),
                multi_mode=self.multi_mode.currentText()
            )
            
            self._handle_chart_result(result, "modify")
            
        except Exception as e:
            self._show_error("Chart Modification", str(e))
        finally:
            self._set_status("", busy=False)
    
    def _on_sum(self):
        """Handle Sum button click."""
        self._set_status("Running Sum…", busy=True)
        self._log("Starting Sum operation...")
        
        try:
            result = gzwcm_sum(logger=self._log)
            self._log(result)
            self._set_status("Sum completed!", busy=False)
            QtWidgets.QMessageBox.information(self, "Sum", result)
        except Exception as e:
            self._log(f"✗ Sum failed: {str(e)}")
            self._show_error("Sum", str(e))
        finally:
            self._set_status("", busy=False)
    
    def _on_slc(self):
        """Handle SLC button click."""
        self._set_status("Running SLC…", busy=True)
        self._log("Starting SLC operation...")
        
        try:
            result = slc(logger=self._log)
            self._log(result)
            self._set_status("SLC completed!", busy=False)
            QtWidgets.QMessageBox.information(self, "SLC", result)
        except Exception as e:
            self._log(f"✗ SLC failed: {str(e)}")
            self._show_error("SLC", str(e))
        finally:
            self._set_status("", busy=False)
    
    def _on_info(self):
        """Handle Info button click."""
        self._set_status("Running Info…", busy=True)
        self._log("Starting Info operation...")
        
        try:
            result = info(logger=self._log)
            self._log(result)
            self._set_status("Info completed!", busy=False)
            QtWidgets.QMessageBox.information(self, "Info", result)
        except Exception as e:
            self._log(f"✗ Info failed: {str(e)}")
            self._show_error("Info", str(e))
        finally:
            self._set_status("", busy=False)
    
    def _on_cat(self):
        """Handle Cat button click."""
        self._set_status("Running Cat…", busy=True)
        self._log("Starting Cat operation...")
        
        try:
            # TODO: Implement cat functionality
            result = "Cat functionality not yet implemented"
            self._log(result)
            self._set_status("Cat completed!", busy=False)
            QtWidgets.QMessageBox.information(self, "Cat", result)
        except Exception as e:
            self._log(f"✗ Cat failed: {str(e)}")
            self._show_error("Cat", str(e))
        finally:
            self._set_status("", busy=False)
    
    def _handle_chart_result(self, result, action):
        """Handle the result of a chart operation."""
        # Log all messages
        for msg in result.log_messages:
            self._log(msg)
        
        if result.success:
            action_text = "created" if action == "create" else "modified"
            self._log(f"✓ Chart {action_text} successfully!")
            self._set_status(f"Chart {action_text}!", busy=False)
        else:
            self._log(f"✗ Failed to {action} chart: {result.error_message}")
            self._show_error(f"Chart {action.title()}", result.error_message)
            self._set_status("", busy=False)
    
    def _on_chart_type_changed(self, text):
        """Update multi_mode options when chart type changes."""
        t = text.lower()
        self.multi_mode.clear()
        
        if 'line' in t:
            self.multi_mode.addItems(["Normal", "Stacked", "100% Stacked"])
        elif 'column' in t:
            self.multi_mode.addItems(["Clustered", "Stacked", "100% Stacked"])
        elif 'bar' in t:
            self.multi_mode.addItems(["Clustered", "Stacked", "100% Stacked"])
        elif 'area' in t:
            self.multi_mode.addItems(["Normal", "Stacked", "100% Stacked"])
        elif 'pie' in t:
            self.multi_mode.addItems(["Pie", "Doughnut", "Pie of Pie"])
        elif 'scatter' in t:
            self.multi_mode.addItems(["Scatter", "Scatter with lines"])
        elif 'radar' in t:
            self.multi_mode.addItems(["Radar", "Filled Radar"])
        else:
            self.multi_mode.addItems(["Default"])
    
    def _on_help(self):
        """Show help dialog."""
        try:
            help_path = os.path.join(os.path.dirname(__file__), 'media', 'help_st_gzwcm.html')
            if os.path.exists(help_path):
                with open(help_path, 'r', encoding='utf-8') as f:
                    help_html = f.read()
                
                dialog = QtWidgets.QDialog(self)
                dialog.setWindowTitle("Peel Potato Help")
                dialog.setMinimumSize(600, 400)
                
                layout = QtWidgets.QVBoxLayout()
                text_browser = QtWidgets.QTextBrowser()
                text_browser.setHtml(help_html)
                text_browser.setOpenExternalLinks(True)
                layout.addWidget(text_browser)
                
                close_btn = QtWidgets.QPushButton("Close")
                close_btn.clicked.connect(dialog.accept)
                layout.addWidget(close_btn)
                
                dialog.setLayout(layout)
                dialog.exec()
            else:
                QtWidgets.QMessageBox.information(self, "Help", 
                    "Help file not found. Please check the installation.")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Help Error", f"Could not load help: {e}")
    
    def _toggle_log(self):
        """Toggle log board visibility."""
        if self.log_board.isVisible():
            self.log_board.hide()
            self.log_toggle_btn.setText("▶")
        else:
            self.log_board.show()
            self.log_toggle_btn.setText("▼")
        QtCore.QTimer.singleShot(0, lambda: self.adjustSize())
    
    def _log(self, message):
        """Append message to log board."""
        try:
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            self.log_board.append(f"[{timestamp}] {message}")
            self.log_board.verticalScrollBar().setValue(
                self.log_board.verticalScrollBar().maximum()
            )
        except Exception:
            pass
    
    def _set_status(self, text, busy=False):
        """Set status label text and cursor."""
        try:
            self.status_label.setText(text)
            QtWidgets.QApplication.processEvents()
            if busy:
                QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
            else:
                QtWidgets.QApplication.restoreOverrideCursor()
        except Exception:
            pass
    
    def _show_error(self, title, message):
        """Show error message dialog with potato theme."""
        potato_titles = [
            "🥔 Oops! The potato got mashed!",
            "🥔 The potato peeler hit a snag!",
            "🥔 Potato malfunction detected!",
            "🥔 The potato needs a moment...",
            "🥔 Chart potato overcooked!"
        ]
        
        error_title = random.choice(potato_titles)
        error_msg = f"{title} encountered an issue:\n\n{str(message)[:200]}"
        
        try:
            QtWidgets.QMessageBox.warning(self, error_title, error_msg)
        except Exception:
            QtWidgets.QMessageBox.warning(self, "Error", error_msg)


def main():
    """Main entry point with crash recovery."""
    max_restarts = 3
    restart_count = 0
    
    while restart_count < max_restarts:
        try:
            app = QtWidgets.QApplication(sys.argv)
            window = PeelPotatoWindow()
            window.show()
            sys.exit(app.exec())
        except Exception as e:
            restart_count += 1
            print(f"Error: {e}. Restart {restart_count}/{max_restarts}")
            if restart_count >= max_restarts:
                break
    
    sys.exit(1)


if __name__ == '__main__':
    main()
