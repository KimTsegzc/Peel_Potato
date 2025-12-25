"""
ST_GZWCM Sum - Filter and summarize employees based on employee list.
Loads emp.xlsx (or emp_embed.xlsx as fallback) and filters active workbook data.
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
        suffix: Suffix to add (e.g., '_filtered')
    
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


def load_employee_list(logger=None):
    """Load employee list from emp.xlsx or emp_embed.xlsx.
    
    Args:
        logger: Optional callback function to log messages (e.g., for UI)
    
    Returns:
        DataFrame with available columns from emp file
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
    
    # User file: emp.xlsx in same directory as exe/script
    emp_file = os.path.join(user_files_dir, 'emp.xlsx')
    # Embedded file: emp_embed.xlsx in data/ folder (packed with exe)
    emp_embed_file = os.path.join(bundle_dir, 'data', 'emp_embed.xlsx')
    
    # Try user file first (emp.xlsx in project root)
    if os.path.exists(emp_file):
        log(f"[DEBUG] Found emp file: {emp_file}")
        try:
            df = pd.read_excel(emp_file, sheet_name='emp')
            log(f"[DEBUG] Loaded emp.xlsx successfully. Shape: {df.shape}")
            log(f"[DEBUG] Columns: {list(df.columns)}")
            return df
        except Exception as e:
            raise Exception(f"Error loading emp.xlsx: {e}")
    
    # Fallback to embedded file (emp_embed.xlsx in data/ folder)
    if os.path.exists(emp_embed_file):
        log(f"[DEBUG] Found embedded emp file: {emp_embed_file}")
        try:
            df = pd.read_excel(emp_embed_file, sheet_name='emp')
            log(f"[DEBUG] Loaded emp_embed.xlsx successfully. Shape: {df.shape}")
            log(f"[DEBUG] Columns: {list(df.columns)}")
            return df
        except Exception as e:
            raise Exception(f"Error loading emp_embed.xlsx: {e}")
    
    raise Exception(f"Employee list file not found.\nSearched for:\n  User file: {emp_file}\n  Embedded file: {emp_embed_file}")


def sum(logger=None):
    """Main sum function.
    
    Args:
        logger: Optional callback function to log messages (e.g., for UI)
    
    Sums up data by emp_id, keeping data_dt, grp, emp_id, emp_nm and adding up numeric columns.
    Adds group sums and total sum rows.
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
        
        # Load employee list
        emplist = load_employee_list(logger=logger)
        if emplist.empty:
            raise Exception("Employee list is empty")
        
        # Get emp_nm list from emplist to filter
        # Find emp_nm column in emplist
        emplist_emp_nm_col = None
        for col in emplist.columns:
            if col and str(col).lower() in [c.lower() for c in EMP_NAME_COLUMN_NAMES]:
                emplist_emp_nm_col = col
                break
        
        if not emplist_emp_nm_col:
            raise Exception(f"Could not find emp_nm column in emp.xlsx. Expected: {', '.join(EMP_NAME_COLUMN_NAMES)}")
        
        # Get set of emp_names to keep
        emp_names_to_keep = set(emplist[emplist_emp_nm_col].dropna().unique())
        
        # Read data from active sheet
        sheet = wb.sheets[active_sheet.Name]
        data_range = sheet.used_range
        data = data_range.value
        
        if not data or len(data) < 2:
            raise Exception("Active sheet has insufficient data")
        
        # Convert to DataFrame
        df = pd.DataFrame(data[1:], columns=data[0])
        
        # Find required columns
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
        
        # Must have emp_id
        if not emp_id_col:
            raise Exception(f"Could not find emp_id column in active sheet. Expected: {', '.join(EMP_ID_COLUMN_NAMES)}")
        
        # Standardize column names
        rename_dict = {}
        if date_col:
            rename_dict[date_col] = 'data_dt'
        if grp_col:
            rename_dict[grp_col] = 'grp'
        if emp_id_col:
            rename_dict[emp_id_col] = 'emp_id'
        if emp_nm_col:
            rename_dict[emp_nm_col] = 'emp_nm'
        
        df = df.rename(columns=rename_dict)
        
        # Convert emp_id to string, remove .0 suffix, and format as 8-digit string
        if 'emp_id' in df.columns:
            df['emp_id'] = df['emp_id'].fillna('').astype(str).str.replace('.0', '', regex=False).str.zfill(8)
        
        # Filter to keep only emp_names in emplist
        if 'emp_nm' not in df.columns:
            raise Exception("Active sheet must have emp_nm column for filtering")
        
        df = df[df['emp_nm'].isin(emp_names_to_keep)].copy()
        
        if df.empty:
            raise Exception("No matching employees found in active sheet")
        
        # Define priority columns order
        priority_cols = []
        if 'data_dt' in df.columns:
            priority_cols.append('data_dt')
        if 'grp' in df.columns:
            priority_cols.append('grp')
        if 'emp_id' in df.columns:
            priority_cols.append('emp_id')
        if 'emp_nm' in df.columns:
            priority_cols.append('emp_nm')
        
        # Get other columns (numeric and non-numeric)
        other_cols = [c for c in df.columns if c not in priority_cols]
        
        # Identify numeric columns for summing
        numeric_cols = []
        for col in other_cols:
            if pd.api.types.is_numeric_dtype(df[col]):
                numeric_cols.append(col)
        
        if not numeric_cols:
            raise Exception("No numeric columns found to sum")
        
        # Get sample data_dt value for sum rows
        sample_dt = df['data_dt'].iloc[0] if 'data_dt' in df.columns and not df.empty else None
        
        # Group by emp_id and sum numeric columns, keeping first values of non-numeric
        agg_dict = {}
        for col in priority_cols:
            if col != 'emp_id':
                agg_dict[col] = 'first'
        for col in numeric_cols:
            agg_dict[col] = 'sum'
        for col in other_cols:
            if col not in numeric_cols:
                agg_dict[col] = 'first'
        
        # Sum by emp_id
        summed_df = df.groupby('emp_id', as_index=False).agg(agg_dict)
        
        # Sort by grp and first numeric column descending
        sort_cols = []
        if 'grp' in summed_df.columns:
            sort_cols.append('grp')
        if numeric_cols:
            sort_cols.append(numeric_cols[0])
        
        if sort_cols:
            ascending = [True] * len(sort_cols)
            ascending[-1] = False  # Last column (first numeric) descending
            summed_df = summed_df.sort_values(by=sort_cols, ascending=ascending)
        
        # Ensure consistent column order
        final_columns = priority_cols + other_cols
        summed_df = summed_df[final_columns]
        
        # Create total sum row first
        total_row = {}
        if 'data_dt' in priority_cols:
            total_row['data_dt'] = sample_dt
        if 'grp' in priority_cols:
            total_row['grp'] = 'all'
        total_row['emp_id'] = 'all'
        if 'emp_nm' in priority_cols:
            total_row['emp_nm'] = 'all'
        
        for col in numeric_cols:
            total_row[col] = summed_df[col].sum()
        for col in other_cols:
            if col not in numeric_cols:
                total_row[col] = None
        
        # Build result with total on top, then group sum before each group
        result_rows = [total_row]
        
        if 'grp' in summed_df.columns:
            # Get unique groups in sorted order
            unique_grps = summed_df['grp'].dropna().unique()
            
            for grp_name in unique_grps:
                # Add group sum row first
                grp_data = summed_df[summed_df['grp'] == grp_name]
                grp_sum_row = {}
                if 'data_dt' in priority_cols:
                    grp_sum_row['data_dt'] = sample_dt
                grp_sum_row['grp'] = grp_name
                grp_sum_row['emp_id'] = f"{grp_name}_sum"
                grp_sum_row['emp_nm'] = f"{grp_name}_sum"
                
                for col in numeric_cols:
                    grp_sum_row[col] = grp_data[col].sum()
                for col in other_cols:
                    if col not in numeric_cols:
                        grp_sum_row[col] = None
                
                result_rows.append(grp_sum_row)
                
                # Then add all employees in this group
                result_rows.extend(grp_data.to_dict('records'))
        else:
            # No group column, just use summed data
            result_rows.extend(summed_df.to_dict('records'))
        
        # Create final result DataFrame
        result_df = pd.DataFrame(result_rows, columns=final_columns)
        
        # Create new sheet with filtered data
        new_sheet_name = 'sum'
        
        # Remove existing sheet with the same name if it exists
        try:
            if new_sheet_name in [s.name for s in wb.sheets]:
                wb.sheets[new_sheet_name].delete()
        except Exception:
            pass
        
        # Create new sheet
        new_sheet = wb.sheets.add(name=new_sheet_name, after=sheet)
        
        # Pre-format emp_id column as text
        if 'emp_id' in result_df.columns:
            emp_id_col_idx = result_df.columns.tolist().index('emp_id') + 1
            emp_id_col_letter = chr(64 + emp_id_col_idx)
            new_sheet.range(f'{emp_id_col_letter}:{emp_id_col_letter}').number_format = '@'
        
        # Write filtered data
        new_sheet.range('A1').value = result_df.columns.tolist()
        new_sheet.range('A2').value = result_df.values.tolist()
        
        # Auto-fit columns
        new_sheet.autofit()
        
        # Count records and groups
        records_count = len([r for r in result_rows if not str(r.get('emp_id', '')).endswith('_sum') and r.get('emp_id') != 'all'])
        groups_count = len([r for r in result_rows if str(r.get('emp_id', '')).endswith('_sum')])
        return f"✓ Created sheet '{new_sheet_name}' with {records_count} employee records, {groups_count} group sums, and 1 total"
        
    except Exception as e:
        raise Exception(f"Sum failed: {str(e)}")
    finally:
        # Cleanup COM for frozen exe
        if getattr(sys, 'frozen', False):
            import pythoncom
            pythoncom.CoUninitialize()


if __name__ == '__main__':
    """For testing purposes."""
    try:
        result = sum()
        print(result)
    except Exception as e:
        print(f"Error: {e}")
