"""
ST_GZWCM Info - Match and fulfill employee information from employee list.
Loads emp.xlsx (or emp_embed.xlsx as fallback) and enriches active workbook data.
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


def load_employee_info(logger=None):
    """Load employee information from emp.xlsx or emp_embed.xlsx.
    
    Args:
        logger: Optional callback function to log messages (e.g., for UI)
    
    Returns:
        DataFrame with columns including: grp, emp_id, emp_nm
    """
    def log(msg):
        if logger:
            logger(msg)
        else:
            print(msg)
    
    # Get directory for user files
    if getattr(sys, 'frozen', False):
        user_files_dir = os.path.dirname(sys.executable)
        bundle_dir = sys._MEIPASS
    else:
        user_files_dir = os.path.dirname(os.path.abspath(__file__))
        bundle_dir = user_files_dir
    
    # User file: emp.xlsx in same directory as exe/script
    emp_file = os.path.join(user_files_dir, 'emp.xlsx')
    # Embedded file: emp_embed.xlsx in data/ folder (packed with exe)
    emp_embed_file = os.path.join(bundle_dir, 'data', 'emp_embed.xlsx')
    
    # Try user file first
    if os.path.exists(emp_file):
        log(f"[DEBUG] Found emp file: {emp_file}")
        try:
            df = pd.read_excel(emp_file, sheet_name='emp')
            log(f"[DEBUG] Loaded emp.xlsx successfully. Shape: {df.shape}")
            log(f"[DEBUG] Columns: {list(df.columns)}")
            return df
        except Exception as e:
            raise Exception(f"Error loading emp.xlsx: {e}")
    
    # Fallback to embedded file
    if os.path.exists(emp_embed_file):
        log(f"[DEBUG] Found embedded emp file: {emp_embed_file}")
        try:
            df = pd.read_excel(emp_embed_file, sheet_name='emp')
            log(f"[DEBUG] Loaded emp_embed.xlsx successfully. Shape: {df.shape}")
            log(f"[DEBUG] Columns: {list(df.columns)}")
            return df
        except Exception as e:
            raise Exception(f"Error loading emp_embed.xlsx: {e}")
    
    raise Exception(f"Employee info file not found.\nSearched for:\n  User file: {emp_file}\n  Embedded file: {emp_embed_file}")


def find_column(df, column_name_list, logger=None):
    """Find a column in DataFrame by checking against a list of possible names (case-insensitive).
    
    Args:
        df: DataFrame to search in
        column_name_list: List of possible column names
        logger: Optional callback function to log messages
    
    Returns:
        Column name if found, None otherwise
    """
    def log(msg):
        if logger:
            logger(msg)
        else:
            print(msg)
    
    for col in df.columns:
        if col and str(col).lower() in [name.lower() for name in column_name_list]:
            log(f"[DEBUG] Found column '{col}' matching {column_name_list}")
            return col
    return None


def info(logger=None):
    """Main info function.
    
    Args:
        logger: Optional callback function to log messages (e.g., for UI)
    
    Matches employee information from emp.xlsx and enriches the active workbook data.
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
        
        # Load employee information
        emp_info = load_employee_info(logger=logger)
        if emp_info.empty:
            raise Exception("Employee information is empty")
        
        # Find emp_nm and emp_id columns in emp_info
        emp_info_nm_col = find_column(emp_info, EMP_NAME_COLUMN_NAMES, logger)
        emp_info_id_col = find_column(emp_info, EMP_ID_COLUMN_NAMES, logger)
        emp_info_grp_col = find_column(emp_info, GRP_COLUMN_NAMES, logger)
        
        if not emp_info_nm_col and not emp_info_id_col:
            raise Exception(f"Employee info file must have either emp_nm or emp_id column. Found columns: {list(emp_info.columns)}")
        
        # Read data from active sheet
        sheet = wb.sheets[active_sheet.Name]
        data_range = sheet.used_range
        data = data_range.value
        
        if not data or len(data) < 2:
            raise Exception("Active sheet has insufficient data")
        
        # Convert to DataFrame
        df = pd.DataFrame(data[1:], columns=data[0])
        
        # Find emp_nm or emp_id column in active sheet
        sheet_emp_nm_col = find_column(df, EMP_NAME_COLUMN_NAMES, logger)
        sheet_emp_id_col = find_column(df, EMP_ID_COLUMN_NAMES, logger)
        
        if not sheet_emp_nm_col and not sheet_emp_id_col:
            raise Exception(f"Active sheet must have either emp_nm or emp_id column.\nExpected emp_nm: {', '.join(EMP_NAME_COLUMN_NAMES)}\nExpected emp_id: {', '.join(EMP_ID_COLUMN_NAMES)}")
        
        # Determine which column to use as key for matching
        if sheet_emp_nm_col:
            key_col = sheet_emp_nm_col
            emp_info_key_col = emp_info_nm_col if emp_info_nm_col else None
            key_type = "emp_nm"
        elif sheet_emp_id_col:
            key_col = sheet_emp_id_col
            emp_info_key_col = emp_info_id_col if emp_info_id_col else None
            key_type = "emp_id"
        else:
            raise Exception("No valid employee column found")
        
        if not emp_info_key_col:
            raise Exception(f"Employee info file does not have matching {key_type} column")
        
        # Standardize column names in emp_info for merging
        emp_info_renamed = emp_info.copy()
        rename_dict = {}
        if emp_info_nm_col:
            rename_dict[emp_info_nm_col] = 'emp_nm'
        if emp_info_id_col:
            rename_dict[emp_info_id_col] = 'emp_id'
        if emp_info_grp_col:
            rename_dict[emp_info_grp_col] = 'grp'
        emp_info_renamed = emp_info_renamed.rename(columns=rename_dict)
        
        # Convert emp_id to string and pad to 8 digits with leading zeros in emp_info
        if 'emp_id' in emp_info_renamed.columns:
            emp_info_renamed['emp_id'] = emp_info_renamed['emp_id'].astype(str).str.zfill(8)
        
        # Merge data based on key column
        df_renamed = df.copy()
        
        # Keep track of original emp_id if exists in active sheet
        has_original_emp_id = sheet_emp_id_col is not None
        if has_original_emp_id and sheet_emp_id_col in df_renamed.columns:
            # Convert and format original emp_id
            df_renamed['emp_id_original'] = df_renamed[sheet_emp_id_col].astype(str).str.zfill(8)
        
        # Rename key column for merging
        df_renamed = df_renamed.rename(columns={key_col: key_type})
        
        # Convert emp_id to string and pad to 8 digits with leading zeros in active sheet (for merging)
        if key_type == 'emp_id':
            df_renamed['emp_id'] = df_renamed['emp_id'].astype(str).str.zfill(8)
        
        # Determine which columns to bring from emp_info
        merge_cols = []
        if key_type == 'emp_nm':
            # If matching by emp_nm, bring grp and emp_id from emp_info
            if 'grp' in emp_info_renamed.columns:
                merge_cols.append('grp')
            if 'emp_id' in emp_info_renamed.columns:
                merge_cols.append('emp_id')
        else:
            # If matching by emp_id, bring grp and emp_nm from emp_info
            if 'grp' in emp_info_renamed.columns:
                merge_cols.append('grp')
            if 'emp_nm' in emp_info_renamed.columns:
                merge_cols.append('emp_nm')
        
        # Merge with employee info
        if merge_cols:
            merged_df = df_renamed.merge(
                emp_info_renamed[[key_type] + merge_cols],
                left_on=key_type,
                right_on=key_type,
                how='left',
                suffixes=('_old', '')
            )
        else:
            merged_df = df_renamed
        
        # If active sheet had emp_id originally, restore it (overwrite merged emp_id)
        if has_original_emp_id and 'emp_id_original' in merged_df.columns:
            merged_df['emp_id'] = merged_df['emp_id_original']
            merged_df = merged_df.drop(columns=['emp_id_original'])
        
        # Remove old columns if they exist
        for col in merged_df.columns:
            if col.endswith('_old'):
                merged_df = merged_df.drop(columns=[col])
        
        # Detect and rename date column
        date_col = find_column(df, DATE_COLUMN_NAMES, logger)
        if date_col and date_col in merged_df.columns:
            merged_df = merged_df.rename(columns={date_col: 'data_dt'})
        
        # Rearrange columns: data_dt (if exists), grp, emp_id, emp_nm, others
        priority_cols = []
        if 'data_dt' in merged_df.columns:
            priority_cols.append('data_dt')
        if 'grp' in merged_df.columns:
            priority_cols.append('grp')
        if 'emp_id' in merged_df.columns:
            priority_cols.append('emp_id')
        if 'emp_nm' in merged_df.columns:
            priority_cols.append('emp_nm')
        
        other_cols = [c for c in merged_df.columns if c not in priority_cols]
        merged_df = merged_df[priority_cols + other_cols]
        
        # Ensure emp_id is string and padded to 8 digits in final output
        if 'emp_id' in merged_df.columns:
            # Convert to string, handle NaN/None, and pad to 8 digits
            merged_df['emp_id'] = merged_df['emp_id'].fillna('').astype(str).str.replace('.0', '', regex=False).str.zfill(8)
        
        # Create new sheet with enriched data
        new_sheet_name = 'info'
        
        # Remove existing sheet with the same name if it exists
        try:
            if new_sheet_name in [s.name for s in wb.sheets]:
                wb.sheets[new_sheet_name].delete()
        except Exception:
            pass
        
        # Create new sheet
        new_sheet = wb.sheets.add(name=new_sheet_name, after=sheet)
        
        # Write headers
        new_sheet.range('A1').value = merged_df.columns.tolist()
        
        # Pre-format emp_id column as text before writing data
        if not merged_df.empty and 'emp_id' in merged_df.columns:
            emp_id_col_idx = merged_df.columns.tolist().index('emp_id') + 1  # 1-based
            emp_id_col_letter = chr(64 + emp_id_col_idx) if emp_id_col_idx <= 26 else 'A' + chr(64 + emp_id_col_idx - 26)
            last_row = len(merged_df) + 1
            emp_id_range = new_sheet.range(f'{emp_id_col_letter}2:{emp_id_col_letter}{last_row}')
            emp_id_range.number_format = '@'  # Set as text format first
        
        # Write data in bulk (convert DataFrame to list to preserve strings)
        if not merged_df.empty:
            # Convert to list of lists to preserve data types
            data_to_write = []
            for _, row in merged_df.iterrows():
                data_to_write.append(row.tolist())
            new_sheet.range('A2').value = data_to_write
        
        # Auto-fit columns
        new_sheet.autofit()
        
        return f"✓ Created sheet '{new_sheet_name}' with {len(merged_df)} records enriched with employee info"
        
    except Exception as e:
        raise Exception(f"Info failed: {str(e)}")
    finally:
        # Cleanup COM for frozen exe
        if getattr(sys, 'frozen', False):
            import pythoncom
            pythoncom.CoUninitialize()


if __name__ == '__main__':
    """For testing purposes."""
    try:
        result = info()
        print(result)
    except Exception as e:
        print(f"Error: {e}")
