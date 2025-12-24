"""
Range Parser for Peel Potato.
Handles parsing of Excel range specifications into xlwings Range objects.
Consolidates all range parsing logic in one place.
"""
import re


class RangeParser:
    """Responsible for parsing range specifications and converting to xlwings Range objects."""
    
    def __init__(self):
        self._cartesian_pattern = re.compile(r'\(([^)]+)\)\s*\*\s*\(([^)]+)\)')
    
    def parse_dim(self, dim_text, sheet):
        """Parse dimension (X-axis/category) range.
        
        Args:
            dim_text: String like "A2:A5", "A", "A2", etc.
            sheet: xlwings Sheet object
            
        Returns:
            xlwings Range object or None
        """
        if not dim_text:
            return None
        
        dim_text = dim_text.strip()
        
        # If contains digits or colon, treat as explicit range
        if any(ch.isdigit() for ch in dim_text) or ':' in dim_text:
            try:
                return sheet.range(dim_text)
            except Exception:
                return None
        
        # Otherwise, it's just a column letter - will be expanded later with ref_rows
        return None
    
    def parse_values(self, values_text, sheet, ref_rows=None):
        """Parse values input into list of xlwings Range objects.
        
        Supports multiple formats:
        - Single range: "B2:B5"
        - Column span: "B:C"
        - Comma-separated: "B,C"
        - Cartesian product: "(B,E)*(2,7)" or "(B:E)*(2:7)"
        
        Args:
            values_text: String specification
            sheet: xlwings Sheet object
            ref_rows: Optional tuple (start_row, end_row) for inferring row range
            
        Returns:
            List of xlwings Range objects
        """
        values_text = (values_text or '').strip()
        if not values_text:
            return []
        
        # Check for Cartesian product format: (cols)*(rows)
        cartesian_match = self._cartesian_pattern.match(values_text)
        if cartesian_match:
            return self._parse_cartesian(cartesian_match, sheet)
        
        # Regular parsing
        return self._parse_regular(values_text, sheet, ref_rows)
    
    def _parse_cartesian(self, match, sheet):
        """Parse Cartesian product format like (B,E)*(2,7)."""
        cols_str = match.group(1).strip()
        rows_str = match.group(2).strip()
        
        # Expand columns
        cols = self.expand_column_range(cols_str)
        
        # Expand rows
        rows = self._expand_row_range(rows_str)
        
        # Build ranges for each column-row combination
        if not cols or not rows:
            return []
        
        ranges = []
        for col in cols:
            try:
                col_idx = self.col_letter_to_index(col)
                if col_idx:
                    start_row = min(rows)
                    end_row = max(rows)
                    ranges.append(sheet.range((start_row, col_idx), (end_row, col_idx)))
            except Exception:
                pass
        
        return ranges
    
    def _parse_regular(self, values_text, sheet, ref_rows):
        """Parse regular comma-separated range specifications."""
        parts = [p.strip() for p in values_text.split(',') if p.strip()]
        ranges = []
        
        # Get used range info for default row ranges
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
            # Explicit range with row numbers (e.g., B2:B5)
            if any(ch.isdigit() for ch in p):
                try:
                    ranges.append(sheet.range(p))
                except Exception:
                    pass
                continue
            
            # Column span (e.g., B:C)
            if ':' in p:
                left, right = [x.strip() for x in p.split(':', 1)]
                if left.isalpha() and right.isalpha() and data_start is not None and used_end is not None:
                    left_idx = self.col_letter_to_index(left)
                    right_idx = self.col_letter_to_index(right)
                    if left_idx and right_idx and left_idx <= right_idx:
                        for col in range(left_idx, right_idx + 1):
                            try:
                                ranges.append(sheet.range((data_start, col), (used_end, col)))
                            except Exception:
                                pass
                    continue
                else:
                    # Try as explicit range
                    try:
                        ranges.append(sheet.range(p))
                    except Exception:
                        pass
                    continue
            
            # Single column letter (e.g., B)
            col_idx = self.col_letter_to_index(p)
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
        """Compute a combined source COM Range covering all ranges in the list.
        
        Args:
            sheet: xlwings Sheet object
            ranges_list: List of xlwings Range objects
            
        Returns:
            COM Range object covering all ranges, or None
        """
        if not ranges_list:
            return None
        
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
    
    def col_letter_to_index(self, letter):
        """Convert Excel column letter to 1-based index.
        
        Args:
            letter: String like 'A', 'B', 'AA', etc.
            
        Returns:
            Integer 1-based index, or None for invalid input
        """
        if letter is None:
            return None
        letter = letter.strip().upper()
        if not letter.isalpha():
            return None
        result = 0
        for ch in letter:
            result = result * 26 + (ord(ch) - ord('A') + 1)
        return result
    
    def expand_column_range(self, cols_str):
        """Expand column specification into list of column letters.
        
        Supports:
        - Single: "B" → ["B"]
        - List: "B,E" → ["B", "E"]
        - Range: "B:E" → ["B", "C", "D", "E"]
        - Mixed: "B:C,E" → ["B", "C", "E"]
        
        Args:
            cols_str: String specification
            
        Returns:
            List of uppercase column letters
        """
        if not cols_str:
            return []
        
        cols = []
        cols_str = cols_str.strip().upper()
        
        # Continuous range like B:E (no comma)
        if ':' in cols_str and ',' not in cols_str:
            parts = [p.strip() for p in cols_str.split(':')]
            if len(parts) == 2:
                start_idx = self.col_letter_to_index(parts[0])
                end_idx = self.col_letter_to_index(parts[1])
                if start_idx and end_idx and start_idx <= end_idx:
                    for i in range(start_idx, end_idx + 1):
                        cols.append(self._index_to_col_letter(i))
            return cols
        
        # Comma-separated parts
        for part in cols_str.split(','):
            part = part.strip()
            if not part:
                continue
            
            if ':' in part:
                # Sub-range within comma list
                sub = [s.strip() for s in part.split(':')]
                if len(sub) == 2:
                    s_idx = self.col_letter_to_index(sub[0])
                    e_idx = self.col_letter_to_index(sub[1])
                    if s_idx and e_idx and s_idx <= e_idx:
                        for i in range(s_idx, e_idx + 1):
                            cols.append(self._index_to_col_letter(i))
            else:
                if part.isalpha():
                    cols.append(part)
        
        return cols
    
    def _expand_row_range(self, rows_str):
        """Expand row specification into list of row numbers."""
        rows = []
        
        # Continuous range like 2:7 (no comma)
        if ':' in rows_str and ',' not in rows_str:
            parts = rows_str.split(':')
            try:
                start_row = int(parts[0].strip())
                end_row = int(parts[1].strip())
                rows = list(range(start_row, end_row + 1))
            except Exception:
                pass
            return rows
        
        # Comma-separated parts
        for part in rows_str.split(','):
            part = part.strip()
            if not part:
                continue
            
            if ':' in part:
                # Sub-range within comma list
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
        
        return rows
    
    def _index_to_col_letter(self, idx):
        """Convert 1-based column index to Excel column letter."""
        col = ''
        while idx > 0:
            idx, remainder = divmod(idx - 1, 26)
            col = chr(65 + remainder) + col
        return col
    
    def infer_dim_range_from_column(self, dim_text, sheet, value_ranges):
        """Infer dim range when only a column letter is provided.
        
        Uses row range from value_ranges or sheet's used range.
        
        Args:
            dim_text: Column letter like "A"
            sheet: xlwings Sheet object
            value_ranges: List of xlwings Range objects (for inferring rows)
            
        Returns:
            xlwings Range object or None
        """
        if not dim_text or not dim_text.strip().isalpha():
            return None
        
        # Infer rows from first value range if available
        ref_rows = None
        try:
            if value_ranges:
                vr0 = value_ranges[0].api
                ref_rows = (vr0.Row, vr0.Row + vr0.Rows.Count - 1)
            else:
                used = sheet.api.UsedRange
                ref_rows = (used.Row, used.Row + used.Rows.Count - 1)
        except Exception:
            ref_rows = None
        
        if not ref_rows:
            return None
        
        col_idx = self.col_letter_to_index(dim_text.strip())
        if not col_idx:
            return None
        
        try:
            return sheet.range((ref_rows[0], col_idx), (ref_rows[1], col_idx))
        except Exception:
            return None
