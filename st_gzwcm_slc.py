"""
ST_GZWCM SLC - Select Columns based on dictionary.
Loads dict.xlsx (or dict_embed.xlsx as fallback) and filters/renames columns in active workbook.
"""
import os
import sys
import re

# Initialize pywin32 for frozen exe
if getattr(sys, 'frozen', False):
    import pywintypes
    import pythoncom

import xlwings as xw
import pandas as pd

# Import shared constants
from st_gzwcm_constants import EMP_NAME_COLUMN_NAMES, EMP_ID_COLUMN_NAMES, DATE_COLUMN_NAMES, GRP_COLUMN_NAMES


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
    
    # Get directory for user files
    # When running as exe, use exe directory; when running as script, use script directory
    if getattr(sys, 'frozen', False):
        # Running as compiled exe - user files are in the same dir as the exe
        user_files_dir = os.path.dirname(sys.executable)
        # Bundled files are in sys._MEIPASS
        bundle_dir = sys._MEIPASS
    else:
        # Running as script
        user_files_dir = os.path.dirname(os.path.abspath(__file__))
        bundle_dir = user_files_dir
    
    # User file: dict.xlsx in same directory as exe/script
    dict_file = os.path.join(user_files_dir, 'dict.xlsx')
    # Embedded file: dict_embed.xlsx in data/ folder (packed with exe)
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


def slc(logger=None):
    """Main slc function.
    
    Args:
        logger: Optional callback function to log messages (e.g., for UI)
    
    Filters active workbook to keep only columns in columnlist (where new is not null),
    and renames them according to the new column names.
    """
    # Initialize COM for frozen exe
    if getattr(sys, 'frozen', False):
        import pythoncom
        pythoncom.CoInitialize()
    
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
        
        # Find default columns (date, grp, emp_id, emp_nm)
        date_col = None
        grp_col = None
        emp_id_col = None
        emp_nm_col = None
        
        for col in df.columns:
            if col:
                col_lower = str(col).lower()
                if not date_col and col_lower in [c.lower() for c in DATE_COLUMN_NAMES]:
                    date_col = col
                if not grp_col and col_lower in [c.lower() for c in GRP_COLUMN_NAMES]:
                    grp_col = col
                if not emp_id_col and col_lower in [c.lower() for c in EMP_ID_COLUMN_NAMES]:
                    emp_id_col = col
                if not emp_nm_col and col_lower in [c.lower() for c in EMP_NAME_COLUMN_NAMES]:
                    emp_nm_col = col
        
        # At least emp_nm or emp_id must exist
        if not emp_nm_col and not emp_id_col:
            raise Exception(f"Could not find employee column in active sheet. Expected emp_nm or emp_id column")
        
        # Start with default columns that exist
        default_columns = []
        if date_col:
            default_columns.append(date_col)
        if grp_col:
            default_columns.append(grp_col)
        if emp_id_col:
            default_columns.append(emp_id_col)
        if emp_nm_col:
            default_columns.append(emp_nm_col)
        
        # Find columns to keep from dict (case-insensitive matching)
        columns_to_keep = []
        rename_map = {}
        
        for old_col_name in columnlist['old']:
            # Skip if this is a default column
            if old_col_name in default_columns:
                continue
            
            # Try exact match first
            if old_col_name in df.columns:
                columns_to_keep.append(old_col_name)
                rename_map[old_col_name] = col_mapping[old_col_name]
            else:
                # Try case-insensitive match
                for col in df.columns:
                    if col and str(col).lower() == str(old_col_name).lower():
                        # Don't include if it's a default column
                        if col not in default_columns:
                            columns_to_keep.append(col)
                            rename_map[col] = col_mapping[old_col_name]
                        break
        
        # Combine default columns with dict columns
        all_columns = default_columns + columns_to_keep
        
        # Filter to keep only selected columns
        filtered_df = df[all_columns].copy()
        
        # Rename columns from dict according to mapping
        filtered_df = filtered_df.rename(columns=rename_map)
        
        # Standardize default column names
        default_rename = {}
        if date_col and date_col in filtered_df.columns:
            default_rename[date_col] = 'data_dt'
        if grp_col and grp_col in filtered_df.columns:
            default_rename[grp_col] = 'grp'
        if emp_id_col and emp_id_col in filtered_df.columns:
            default_rename[emp_id_col] = 'emp_id'
        if emp_nm_col and emp_nm_col in filtered_df.columns:
            default_rename[emp_nm_col] = 'emp_nm'
        
        filtered_df = filtered_df.rename(columns=default_rename)
        
        # Keep emp_id as string without padding - just remove .0 suffix if present
        if 'emp_id' in filtered_df.columns:
            filtered_df['emp_id'] = filtered_df['emp_id'].fillna('').astype(str).str.replace('.0', '', regex=False)
        
        # Reorder columns to ensure data_dt, grp, emp_id, emp_nm order
        priority_cols = []
        if 'data_dt' in filtered_df.columns:
            priority_cols.append('data_dt')
        if 'grp' in filtered_df.columns:
            priority_cols.append('grp')
        if 'emp_id' in filtered_df.columns:
            priority_cols.append('emp_id')
        if 'emp_nm' in filtered_df.columns:
            priority_cols.append('emp_nm')
        
        other_cols = [c for c in filtered_df.columns if c not in priority_cols]
        filtered_df = filtered_df[priority_cols + other_cols]
        
        # Create new sheet with filtered and renamed data
        new_sheet_name = 'slc'
        
        # Remove existing sheet with the same name if it exists
        try:
            if new_sheet_name in [s.name for s in wb.sheets]:
                wb.sheets[new_sheet_name].delete()
        except Exception:
            pass
        
        # Create new sheet
        new_sheet = wb.sheets.add(name=new_sheet_name, after=sheet)
        
        # Write headers
        new_sheet.range('A1').value = filtered_df.columns.tolist()
        
        # Pre-format emp_id column as text before writing data
        if not filtered_df.empty and 'emp_id' in filtered_df.columns:
            emp_id_col_idx = filtered_df.columns.tolist().index('emp_id') + 1  # 1-based
            emp_id_col_letter = chr(64 + emp_id_col_idx) if emp_id_col_idx <= 26 else 'A' + chr(64 + emp_id_col_idx - 26)
            last_row = len(filtered_df) + 1
            emp_id_range = new_sheet.range(f'{emp_id_col_letter}2:{emp_id_col_letter}{last_row}')
            emp_id_range.number_format = '@'  # Set as text format
        
        # Write filtered data with new column names
        if not filtered_df.empty:
            # Convert to list of lists to preserve string types
            data_to_write = []
            for _, row in filtered_df.iterrows():
                data_to_write.append(row.tolist())
            new_sheet.range('A2').value = data_to_write
        
        # Auto-fit columns
        new_sheet.autofit()
        
        return f"✓ Created sheet '{new_sheet_name}' with {len(filtered_df.columns)} selected columns"
        
    except Exception as e:
        raise Exception(f"SLC failed: {str(e)}")
    finally:
        # Cleanup COM for frozen exe
        if getattr(sys, 'frozen', False):
            import pythoncom
            pythoncom.CoUninitialize()


if __name__ == '__main__':
    """For testing purposes."""
    try:
        result = slc()
        print(result)
    except Exception as e:
        print(f"Error: {e}")
