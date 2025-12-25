"""
Excel Adapter for Peel Potato.
Handles all xlwings and COM interactions, isolating Excel operations.
Renamed from PeelPotatoEngine to better reflect its purpose.
"""
import sys

# Initialize pywin32 for frozen exe
if getattr(sys, 'frozen', False):
    import pywintypes
    import pythoncom
    # Initialize COM for this thread
    try:
        pythoncom.CoInitialize()
    except:
        pass

import xlwings as xw

try:
    from win32com.client import constants as xlconst
except Exception:
    xlconst = None


class ExcelAdapter:
    """Responsible for all Excel/COM API interactions."""
    
    def __init__(self):
        pass
    
    def get_active_sheet(self):
        """Get the currently focused xlwings Sheet object.
        
        Returns:
            xlwings Sheet object or None if no Excel instance is active
        """
        try:
            app = xw.apps.active
            if app is None:
                return None
            
            try:
                active_wb_api = app.api.ActiveWorkbook
                active_sh_api = app.api.ActiveSheet
                bname = getattr(active_wb_api, 'Name', None)
                sname = getattr(active_sh_api, 'Name', None)
            except Exception:
                bname = None
                sname = None
            
            book = None
            if bname:
                for b in app.books:
                    if b.name == bname:
                        book = b
                        break
            if book is None and app.books:
                book = app.books[0]
            if book is None:
                return None
            
            try:
                if sname and sname in [s.name for s in book.sheets]:
                    return book.sheets[sname]
                return book.sheets[0]
            except Exception:
                return book.sheets[0]
        except Exception:
            return None
    
    def get_active_workbook_info(self):
        """Get information about the active workbook and sheet.
        
        Returns:
            tuple: (workbook_name, sheet_name) or (None, None) if no Excel
        """
        try:
            app = xw.apps.active
            if app is None:
                return None, None
            
            try:
                active_wb_api = app.api.ActiveWorkbook
                active_sh_api = app.api.ActiveSheet
                bname = getattr(active_wb_api, 'Name', None)
                sname = getattr(active_sh_api, 'Name', None)
                return bname, sname
            except Exception:
                return None, None
        except Exception:
            return None, None
    
    def create_chart_object(self, sheet, left=50, top=20, width=520, height=320):
        """Create a new chart object on the sheet.
        
        Args:
            sheet: xlwings Sheet object
            left: Left position in points
            top: Top position in points
            width: Width in points
            height: Height in points
            
        Returns:
            COM Chart object
        """
        try:
            sht_api = sheet.api
            chart_objects = sht_api.ChartObjects()
            chart_obj = chart_objects.Add(left, top, width, height)
            return chart_obj.Chart
        except Exception as e:
            raise RuntimeError(f"Failed to create chart object: {e}")
    
    def begin_performance_mode(self, sheet):
        """Disable ScreenUpdating/Events/Alerts and set manual calculation.
        
        This speeds up Excel operations by preventing screen updates and
        recalculations during bulk operations.
        
        Args:
            sheet: xlwings Sheet object
            
        Returns:
            tuple: (app_api, saved_state_dict) for later restoration
        """
        sht_api = sheet.api
        app_api = None
        _excel_saved = {}
        
        try:
            app_api = sht_api.Application
            
            # Save current state
            try:
                _excel_saved['ScreenUpdating'] = app_api.ScreenUpdating
            except Exception:
                _excel_saved['ScreenUpdating'] = None
            try:
                _excel_saved['EnableEvents'] = app_api.EnableEvents
            except Exception:
                _excel_saved['EnableEvents'] = None
            try:
                _excel_saved['DisplayAlerts'] = app_api.DisplayAlerts
            except Exception:
                _excel_saved['DisplayAlerts'] = None
            try:
                _excel_saved['Calculation'] = app_api.Calculation
            except Exception:
                _excel_saved['Calculation'] = None
            
            # Disable for performance
            try:
                app_api.ScreenUpdating = False
            except Exception:
                pass
            try:
                app_api.EnableEvents = False
            except Exception:
                pass
            try:
                app_api.DisplayAlerts = False
            except Exception:
                pass
            try:
                manual_calc = getattr(xlconst, 'xlCalculationManual', -4135) if xlconst is not None else -4135
                app_api.Calculation = manual_calc
            except Exception:
                pass
        except Exception:
            app_api = None
        
        return app_api, _excel_saved
    
    def end_performance_mode(self, app_api, saved):
        """Restore Excel UI and calculation settings.
        
        Args:
            app_api: Excel Application COM object
            saved: Dictionary of saved settings from begin_performance_mode
        """
        try:
            if app_api is None or not isinstance(saved, dict):
                return
            
            try:
                if saved.get('ScreenUpdating') is not None:
                    app_api.ScreenUpdating = saved.get('ScreenUpdating')
            except Exception:
                pass
            try:
                if saved.get('EnableEvents') is not None:
                    app_api.EnableEvents = saved.get('EnableEvents')
            except Exception:
                pass
            try:
                if saved.get('DisplayAlerts') is not None:
                    app_api.DisplayAlerts = saved.get('DisplayAlerts')
            except Exception:
                pass
            try:
                if saved.get('Calculation') is not None:
                    app_api.Calculation = saved.get('Calculation')
            except Exception:
                pass
        except Exception:
            pass
    
    def get_selected_chart(self, sheet):
        """Get the currently selected chart on the sheet.
        
        Args:
            sheet: xlwings Sheet object
            
        Returns:
            COM Chart object or None if no chart is selected
        """
        try:
            sht_api = sheet.api
            app_api = sht_api.Application
            
            # Check if selection is a ChartObject
            try:
                selection = app_api.Selection
                if hasattr(selection, 'Chart'):
                    return selection.Chart
            except Exception:
                pass
            
            # Check if active chart exists
            try:
                if hasattr(sht_api, 'ChartObjects'):
                    chart_objects = sht_api.ChartObjects()
                    if chart_objects.Count > 0:
                        # Return most recently created chart
                        return chart_objects(chart_objects.Count).Chart
            except Exception:
                pass
            
            return None
        except Exception:
            return None
    
    def validate_excel_available(self):
        """Check if Excel is available and accessible.
        
        Returns:
            bool: True if Excel is available, False otherwise
        """
        try:
            app = xw.apps.active
            return app is not None
        except Exception:
            return False
    
    def get_range(self, sheet, address):
        """Get a range object from the sheet.
        
        Args:
            sheet: xlwings Sheet object
            address: Range address like "A1:A10"
            
        Returns:
            xlwings Range object
        """
        try:
            return sheet.range(address)
        except Exception as e:
            raise ValueError(f"Invalid range address '{address}': {e}")
