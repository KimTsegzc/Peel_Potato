"""
Backend engine for Peel Potato: handles xlwings / COM interactions.
This module isolates Excel-specific operations so the UI code stays thin.
"""
from typing import List, Tuple
import xlwings as xw

try:
    from win32com.client import constants as xlconst
except Exception:
    xlconst = None


class PeelPotatoEngine:
    def __init__(self):
        pass

    def get_selected_sheet(self):
        """Return the currently focused xlwings Sheet object or None."""
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

    def parse_values_input(self, values_text: str, sheet, ref_rows=None) -> List:
        """Return list of xlwings Range objects for the given values_text, using
        logic similar to the original implementation but isolated here.
        """
        values_text = (values_text or '').strip()
        if not values_text:
            return []

        import re
        cartesian_pattern = r'\(([^)]+)\)\s*\*\s*\(([^)]+)\)'
        cartesian_match = re.match(cartesian_pattern, values_text)

        if cartesian_match:
            cols_str = cartesian_match.group(1).strip()
            rows_str = cartesian_match.group(2).strip()

            # Parse columns
            cols = []
            if ':' in cols_str and ',' not in cols_str:
                col_parts = [c.strip().upper() for c in cols_str.split(':')]
                if len(col_parts) == 2:
                    try:
                        from peel_potato_logic import col_letter_to_index
                        start_col_idx = col_letter_to_index(col_parts[0])
                        end_col_idx = col_letter_to_index(col_parts[1])
                        if start_col_idx and end_col_idx:
                            for col_idx in range(start_col_idx, end_col_idx + 1):
                                col_letter = ''
                                idx = col_idx
                                while idx > 0:
                                    idx, remainder = divmod(idx - 1, 26)
                                    col_letter = chr(65 + remainder) + col_letter
                                cols.append(col_letter)
                    except Exception:
                        pass
            else:
                for part in cols_str.split(','):
                    part = part.strip().upper()
                    if ':' in part:
                        col_parts = [c.strip().upper() for c in part.split(':')]
                        if len(col_parts) == 2:
                            try:
                                from peel_potato_logic import col_letter_to_index
                                start_col_idx = col_letter_to_index(col_parts[0])
                                end_col_idx = col_letter_to_index(col_parts[1])
                                if start_col_idx and end_col_idx:
                                    for col_idx in range(start_col_idx, end_col_idx + 1):
                                        col_letter = ''
                                        idx = col_idx
                                        while idx > 0:
                                            idx, remainder = divmod(idx - 1, 26)
                                            col_letter = chr(65 + remainder) + col_letter
                                        cols.append(col_letter)
                            except Exception:
                                pass
                    else:
                        cols.append(part)

            # Parse rows
            rows = []
            if ':' in rows_str and ',' not in rows_str:
                row_parts = rows_str.split(':')
                try:
                    start_row = int(row_parts[0].strip())
                    end_row = int(row_parts[1].strip())
                    rows = list(range(start_row, end_row + 1))
                except Exception:
                    pass
            else:
                for part in rows_str.split(','):
                    part = part.strip()
                    if ':' in part:
                        row_parts = part.split(':')
                        try:
                            start_row = int(row_parts[0].strip())
                            end_row = int(row_parts[1].strip())
                            rows.extend(range(start_row, end_row + 1))
                        except Exception:
                            pass
                    else:
                        try:
                            rows.append(int(part))
                        except Exception:
                            pass

            if cols and rows:
                ranges = []
                for col in cols:
                    try:
                        col_idx = None
                        try:
                            from peel_potato_logic import col_letter_to_index
                            col_idx = col_letter_to_index(col)
                        except Exception:
                            pass
                        if col_idx:
                            try:
                                start_row = min(rows)
                                end_row = max(rows)
                                ranges.append(sheet.range((start_row, col_idx), (end_row, col_idx)))
                            except Exception:
                                pass
                    except Exception:
                        pass
                return ranges

        # Regular parsing
        parts = [p.strip() for p in values_text.split(',') if p.strip()]
        ranges = []
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
            if any(ch.isdigit() for ch in p):
                try:
                    ranges.append(sheet.range(p))
                except Exception:
                    pass
                continue

            if ':' in p:
                left, right = [x.strip() for x in p.split(':', 1)]
                if left.isalpha() and right.isalpha() and data_start is not None and used_end is not None:
                    left_idx = None
                    right_idx = None
                    try:
                        from peel_potato_logic import col_letter_to_index
                        left_idx = col_letter_to_index(left)
                        right_idx = col_letter_to_index(right)
                    except Exception:
                        pass
                    if left_idx and right_idx and left_idx <= right_idx:
                        for col in range(left_idx, right_idx + 1):
                            try:
                                ranges.append(sheet.range((data_start, col), (used_end, col)))
                            except Exception:
                                pass
                    continue
                else:
                    try:
                        ranges.append(sheet.range(p))
                    except Exception:
                        pass
                    continue

            col_idx = None
            try:
                from peel_potato_logic import col_letter_to_index
                col_idx = col_letter_to_index(p)
            except Exception:
                col_idx = None
            if col_idx is None:
                continue
            if ref_rows:
                start_row, end_row = ref_rows
                ranges.append(sheet.range((start_row, col_idx), (end_row, col_idx)))
            elif data_start is not None and used_end is not None:
                ranges.append(sheet.range((data_start, col_idx), (used_end, col_idx)))
            else:
                ranges.append(sheet.range((2, col_idx), (10000, col_idx)))
        return ranges

    def compute_source_block(self, sheet, ranges_list):
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

    def begin_performance_mode(self, sheet):
        """Disable ScreenUpdating/Events/Alerts and set manual calculation if possible.
        Returns (app_api, saved_state_dict).
        """
        sht_api = sheet.api
        app_api = None
        _excel_saved = {}
        try:
            app_api = sht_api.Application
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
