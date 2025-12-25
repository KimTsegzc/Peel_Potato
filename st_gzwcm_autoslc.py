"""
ST_GZWCM AutoSLC - Auto Select Columns based on dictionary.
Loads dict.xlsx (or dict_embed.xlsx as fallback) and filters/renames columns in active workbook.
"""
import os
import sys
import re
import xlwings as xw
import pandas as pd

# Import shared constants
from st_gzwcm_constants import EMP_COLUMN_NAMES, DATE_COLUMN_NAMES, GRP_COLUMN_NAMES


def sanitize_sheet_name(name, suffix=''):
    """Sanitize sheet name to comply with Excel restrictions.
    
    Args:
        name: Original sheet name
        suffix: Suffix to add (e.g., '_slc')
    
    Returns:
        Valid Excel sheet name (max 31 chars, no invalid chars)
    """
    # Remove invalid characters: : \ / ? * [ ]
    clean_name = re.sub(r'[:\\/*?\[\]]', '_', str(name))
    
    # Add suffix
    full_name = f"{clean_name}{suffix}"
    
    # Limit to 31 characters
    if len(full_name) > 31:
        # Keep as much of the original name as possible
        max_base_len = 31 - len(suffix)
        full_name = f"{clean_name[:max_base_len]}{suffix}"
    
    return full_name


def load_column_dict(logger=None):
    """Load column dictionary from dict.xlsx or dict_embed.xlsx.
    
    Args:
        logger: Optional callback function to log messages (e.g., for UI)
    
    Returns:
        DataFrame with columns: old, new
    """
    def log(msg):
        if logger:
            logger(msg)
        else:
            print(msg)
    
    # Get project root (where this script is located)
    project_root = os.path.dirname(os.path.abspath(__file__))
    
    # User file: dict.xlsx in project root (same dir as .py/.exe)
    dict_file = os.path.join(project_root, 'dict.xlsx')
    
    # Embedded file: dict_embed.xlsx in data/ folder (packed with exe)
    # When running as PyInstaller exe, bundled files are in sys._MEIPASS
    if getattr(sys, 'frozen', False):
        # Running as compiled exe
        bundle_dir = sys._MEIPASS
    else:
        # Running as script
        bundle_dir = project_root
    dict_embed_file = os.path.join(bundle_dir, 'data', 'dict_embed.xlsx')
    
    # Try user file first (dict.xlsx in project root)
    if os.path.exists(dict_file):
        log(f"[DEBUG] Found dict file: {dict_file}")
        try:
            df = pd.read_excel(dict_file, sheet_name='dict')
            log(f"[DEBUG] Loaded dict.xlsx successfully. Shape: {df.shape}")
            log(f"[DEBUG] First 3 rows:\n{df.head(3)}")
            if 'old' in df.columns and 'new' in df.columns:
                return df[['old', 'new']]
            else:
                raise Exception(f"dict.xlsx must have 'old' and 'new' columns. Found columns: {list(df.columns)}")
        except Exception as e:
            raise Exception(f"Error loading dict.xlsx: {e}")
    
    # Fallback to embedded file (dict_embed.xlsx in data/ folder)
    if os.path.exists(dict_embed_file):
        log(f"[DEBUG] Found embedded dict file: {dict_embed_file}")
        try:
            df = pd.read_excel(dict_embed_file, sheet_name='dict')
            log(f"[DEBUG] Loaded dict_embed.xlsx successfully. Shape: {df.shape}")
            log(f"[DEBUG] First 3 rows:\n{df.head(3)}")
            if 'old' in df.columns and 'new' in df.columns:
                return df[['old', 'new']]
            else:
                raise Exception(f"dict_embed.xlsx must have 'old' and 'new' columns. Found columns: {list(df.columns)}")
        except Exception as e:
            raise Exception(f"Error loading dict_embed.xlsx: {e}")
    
    raise Exception(f"Column dictionary file not found.\nSearched for:\n  User file: {dict_file}\n  Embedded file: {dict_embed_file}")


def autoslc(logger=None):
    """Main autoslc function.
    
    Args:
        logger: Optional callback function to log messages (e.g., for UI)
    
    Filters active workbook to keep only columns in columnlist (where new is not null),
    and renames them according to the new column names.
    """
    try:
        # Get active Excel application and workbook
        app = xw.apps.active
        if app is None:
            raise Exception("No active Excel application found")
        
        wb = app.books.active
        if wb is None:
            raise Exception("No active workbook found")
        
        active_sheet = app.api.ActiveSheet
        if active_sheet is None:
            raise Exception("No active sheet found")
        
        # Load column dictionary
        columnlist = load_column_dict(logger=logger)
        if columnlist.empty:
            raise Exception("Column dictionary is empty")
        
        # Filter to only rows where 'new' is not null
        columnlist = columnlist[columnlist['new'].notna()]
        
        if columnlist.empty:
            raise Exception("No valid column mappings found (all 'new' values are null)")
        
        # Create mapping dictionary
        col_mapping = dict(zip(columnlist['old'], columnlist['new']))
        
        # Read data from active sheet
        sheet = wb.sheets[active_sheet.Name]
        data_range = sheet.used_range
        data = data_range.value
        
        if not data or len(data) < 1:
            raise Exception("Active sheet has insufficient data")
        
        # Convert to DataFrame
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
        else:
            df = pd.DataFrame(columns=data[0])
        
        # Find and validate employee column first (required)
        emp_col = None
        for col in df.columns:
            if col and str(col).lower() in EMP_COLUMN_NAMES:
                emp_col = col
                break
        
        if emp_col is None:
            raise Exception(f"Could not find employee column in active sheet. Expected column named: {', '.join(EMP_COLUMN_NAMES)}")
        
        # Find columns to keep (case-insensitive matching)
        columns_to_keep = []
        rename_map = {}
        
        for old_col_name in columnlist['old']:
            # Try exact match first
            if old_col_name in df.columns:
                columns_to_keep.append(old_col_name)
                rename_map[old_col_name] = col_mapping[old_col_name]
            else:
                # Try case-insensitive match
                for col in df.columns:
                    if col and str(col).lower() == str(old_col_name).lower():
                        columns_to_keep.append(col)
                        rename_map[col] = col_mapping[old_col_name]
                        break
        
        if not columns_to_keep:
            raise Exception("No matching columns found in active sheet")
        
        # Detect date column before filtering
        date_col = None
        for col in df.columns:
            if col and str(col).lower() in DATE_COLUMN_NAMES:
                date_col = col
                break
        
        # Always include date and emp columns if they exist and aren't already included
        if date_col and date_col not in columns_to_keep:
            columns_to_keep.insert(0, date_col)
        # emp_col is guaranteed to exist (validated earlier), include it if not already there
        if emp_col not in columns_to_keep:
            columns_to_keep.insert(0, emp_col)
        
        # Filter columns
        filtered_df = df[columns_to_keep].copy()
        
        # Rename columns according to mapping
        filtered_df = filtered_df.rename(columns=rename_map)
        
        # Rename date column to dt_date
        if date_col:
            # Check if date_col was renamed or is still original
            if date_col in filtered_df.columns:
                filtered_df = filtered_df.rename(columns={date_col: 'dt_date'})
            elif date_col in rename_map and rename_map[date_col] in filtered_df.columns:
                filtered_df = filtered_df.rename(columns={rename_map[date_col]: 'dt_date'})
        
        # Rename emp column to 'emp' for consistency (emp_col is guaranteed to exist)
        emp_renamed = rename_map.get(emp_col, emp_col)
        if emp_renamed != 'emp' and emp_renamed in filtered_df.columns:
            filtered_df = filtered_df.rename(columns={emp_renamed: 'emp'})
        elif emp_col != 'emp' and emp_col in filtered_df.columns:
            filtered_df = filtered_df.rename(columns={emp_col: 'emp'})
        
        # Check if grp column exists in source data and include it if present
        grp_col = None
        for col in df.columns:
            if col and str(col).lower() in GRP_COLUMN_NAMES:
                grp_col = col
                break
        
        if grp_col and grp_col not in columns_to_keep:
            # Add grp column from source data (don't merge from emp.xlsx)
            if 'dt_date' in filtered_df.columns:
                grp_pos = list(filtered_df.columns).index('dt_date') + 1
            else:
                grp_pos = 0
            filtered_df.insert(grp_pos, 'grp', df[grp_col])
        elif grp_col and grp_col in columns_to_keep:
            # Rename grp column to 'grp' if it exists but has different name
            grp_renamed = rename_map.get(grp_col, grp_col)
            if grp_renamed != 'grp' and grp_renamed in filtered_df.columns:
                filtered_df = filtered_df.rename(columns={grp_renamed: 'grp'})
        
        # Reorder columns to ensure dt_date, grp, emp order
        priority_cols = []
        if 'dt_date' in filtered_df.columns:
            priority_cols.append('dt_date')
        if 'grp' in filtered_df.columns:
            priority_cols.append('grp')
        if 'emp' in filtered_df.columns:
            priority_cols.append('emp')
        
        other_cols = [c for c in filtered_df.columns if c not in priority_cols]
        filtered_df = filtered_df[priority_cols + other_cols]
        
        # Create new sheet with filtered and renamed data
        new_sheet_name = 'autoslc'
        
        # Remove existing sheet with the same name if it exists
        try:
            if new_sheet_name in [s.name for s in wb.sheets]:
                wb.sheets[new_sheet_name].delete()
        except Exception:
            pass
        
        # Create new sheet
        new_sheet = wb.sheets.add(name=new_sheet_name, after=sheet)
        
        # Write filtered data with new column names
        new_sheet.range('A1').value = filtered_df.columns.tolist()
        if not filtered_df.empty:
            new_sheet.range('A2').value = filtered_df.values.tolist()
        
        # Auto-fit columns
        new_sheet.autofit()
        
        return f"✓ Created sheet '{new_sheet_name}' with {len(filtered_df.columns)} selected columns"
        
    except Exception as e:
        raise Exception(f"AutoSLC failed: {str(e)}")


if __name__ == '__main__':
    """For testing purposes."""
    try:
        result = autoslc()
        print(result)
    except Exception as e:
        print(f"Error: {e}")
