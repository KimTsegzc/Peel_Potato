"""
ST_GZWCM AutoSum - Filter and summarize employees based on employee list.
Loads emp.xlsx (or emp_embed.xlsx as fallback) and filters active workbook data.
"""
import os
import sys
import re
import xlwings as xw
import pandas as pd

# Import shared constants
from st_gzwcm_constants import EMP_COLUMN_NAMES, DATE_COLUMN_NAMES


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
        DataFrame with columns: grp, emp
    """
    def log(msg):
        if logger:
            logger(msg)
        else:
            print(msg)
    
    # Get project root (where this script is located)
    project_root = os.path.dirname(os.path.abspath(__file__))
    
    # User file: emp.xlsx in project root (same dir as .py/.exe)
    emp_file = os.path.join(project_root, 'emp.xlsx')
    
    # Embedded file: emp_embed.xlsx in data/ folder (packed with exe)
    # When running as PyInstaller exe, bundled files are in sys._MEIPASS
    if getattr(sys, 'frozen', False):
        # Running as compiled exe
        bundle_dir = sys._MEIPASS
    else:
        # Running as script
        bundle_dir = project_root
    emp_embed_file = os.path.join(bundle_dir, 'data', 'emp_embed.xlsx')
    
    # Try user file first (emp.xlsx in project root)
    if os.path.exists(emp_file):
        log(f"[DEBUG] Found emp file: {emp_file}")
        try:
            df = pd.read_excel(emp_file, sheet_name='emp')
            log(f"[DEBUG] Loaded emp.xlsx successfully. Shape: {df.shape}")
            log(f"[DEBUG] First 3 rows:\n{df.head(3)}")
            if 'grp' in df.columns and 'emp' in df.columns:
                return df[['grp', 'emp']]
            else:
                raise Exception(f"emp.xlsx must have 'grp' and 'emp' columns. Found columns: {list(df.columns)}")
        except Exception as e:
            raise Exception(f"Error loading emp.xlsx: {e}")
    
    # Fallback to embedded file (emp_embed.xlsx in data/ folder)
    if os.path.exists(emp_embed_file):
        log(f"[DEBUG] Found embedded emp file: {emp_embed_file}")
        try:
            df = pd.read_excel(emp_embed_file, sheet_name='emp')
            log(f"[DEBUG] Loaded emp_embed.xlsx successfully. Shape: {df.shape}")
            log(f"[DEBUG] First 3 rows:\n{df.head(3)}")
            if 'grp' in df.columns and 'emp' in df.columns:
                return df[['grp', 'emp']]
            else:
                raise Exception(f"emp_embed.xlsx must have 'grp' and 'emp' columns. Found columns: {list(df.columns)}")
        except Exception as e:
            raise Exception(f"Error loading emp_embed.xlsx: {e}")
    
    raise Exception(f"Employee list file not found.\nSearched for:\n  User file: {emp_file}\n  Embedded file: {emp_embed_file}")


def autosum(logger=None):
    """Main autosum function.
    
    Args:
        logger: Optional callback function to log messages (e.g., for UI)
    
    Filters active workbook data to include only employees in the emplist,
    and creates a new sheet with the filtered results.
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
        
        # Load employee list
        emplist = load_employee_list(logger=logger)
        if emplist.empty:
            raise Exception("Employee list is empty")
        
        # Get employee names to filter
        emp_names = set(emplist['emp'].dropna().unique())
        
        # Read data from active sheet
        sheet = wb.sheets[active_sheet.Name]
        data_range = sheet.used_range
        data = data_range.value
        
        if not data or len(data) < 2:
            raise Exception("Active sheet has insufficient data")
        
        # Convert to DataFrame
        df = pd.DataFrame(data[1:], columns=data[0])
        
        # Try to find employee column (case-insensitive search for common names)
        emp_col = None
        for col in df.columns:
            if col and str(col).lower() in EMP_COLUMN_NAMES:
                emp_col = col
                break
        
        if emp_col is None:
            raise Exception(f"Could not find employee column in active sheet. Expected column named: {', '.join(EMP_COLUMN_NAMES)}")
        
        # Filter data to include only employees in emplist
        filtered_df = df[df[emp_col].isin(emp_names)].copy()
        
        if filtered_df.empty:
            raise Exception("No matching employees found in active sheet")
        
        # Merge grp information from emplist based on employee name
        # Create mapping from emp to grp
        emp_to_grp = dict(zip(emplist['emp'], emplist['grp']))
        
        # Add or update grp column
        filtered_df['grp'] = filtered_df[emp_col].map(emp_to_grp)
        
        # Detect and rename date column
        date_col = None
        for col in filtered_df.columns:
            if col and str(col).lower() in DATE_COLUMN_NAMES:
                date_col = col
                break
        
        # If date column found, rename it to dt_date
        if date_col and date_col != 'dt_date':
            filtered_df = filtered_df.rename(columns={date_col: 'dt_date'})
        
        # Rename emp column to 'emp' for consistency
        if emp_col != 'emp':
            filtered_df = filtered_df.rename(columns={emp_col: 'emp'})
        
        # Ensure dt_date, grp, emp are at the front (in this order)
        priority_cols = []
        if 'dt_date' in filtered_df.columns:
            priority_cols.append('dt_date')
        if 'grp' in filtered_df.columns:
            priority_cols.append('grp')
        if 'emp' in filtered_df.columns:
            priority_cols.append('emp')
        
        # Reorder columns: priority columns first, then rest
        other_cols = [c for c in filtered_df.columns if c not in priority_cols]
        filtered_df = filtered_df[priority_cols + other_cols]
        
        # Identify numeric columns for summing (exclude dt_date, grp, emp)
        numeric_cols = []
        for col in other_cols:
            if pd.api.types.is_numeric_dtype(filtered_df[col]):
                numeric_cols.append(col)
        
        # Calculate sum by group
        if numeric_cols:
            # Get a sample dt_date value to use for summary rows
            sample_dt_date = filtered_df['dt_date'].iloc[0] if 'dt_date' in filtered_df.columns and not filtered_df.empty else None
            
            # Group by grp and sum numeric columns
            grp_sums = filtered_df.groupby('grp')[numeric_cols].sum().reset_index()
            
            # Add grp and emp columns with specific naming
            grp_sum_rows = []
            for _, row in grp_sums.iterrows():
                grp_id = row['grp']
                sum_row = {}
                if 'dt_date' in priority_cols:
                    sum_row['dt_date'] = sample_dt_date
                sum_row['grp'] = grp_id
                sum_row['emp'] = f"{grp_id}_sum"
                for col in numeric_cols:
                    sum_row[col] = row[col]
                # Fill non-numeric columns with None
                for col in other_cols:
                    if col not in numeric_cols:
                        sum_row[col] = None
                grp_sum_rows.append(sum_row)
            
            # Calculate total sum across all groups
            total_sums = filtered_df[numeric_cols].sum()
            total_row = {}
            if 'dt_date' in priority_cols:
                total_row['dt_date'] = sample_dt_date
            total_row['grp'] = 'all'
            total_row['emp'] = 'all'
            for col in numeric_cols:
                total_row[col] = total_sums[col]
            for col in other_cols:
                if col not in numeric_cols:
                    total_row[col] = None
            
            # Combine original data, group sums, and total
            grp_sum_df = pd.DataFrame(grp_sum_rows)
            total_df = pd.DataFrame([total_row])
            
            # Ensure all dataframes have the same column order
            final_columns = priority_cols + other_cols
            filtered_df = filtered_df[final_columns]
            grp_sum_df = grp_sum_df[final_columns]
            total_df = total_df[final_columns]
            
            # Concatenate all data
            result_df = pd.concat([filtered_df, grp_sum_df, total_df], ignore_index=True)
            
            # Sort by grp, then by column 4 (index 3) descending
            # Column 4 is the first column after dt_date, grp, emp
            if len(result_df.columns) > 3:
                col_4_name = result_df.columns[3]
                result_df = result_df.sort_values(by=['grp', col_4_name], ascending=[True, False])
        else:
            result_df = filtered_df
        
        # Create new sheet with filtered data
        new_sheet_name = 'autosum'
        
        # Remove existing sheet with the same name if it exists
        try:
            if new_sheet_name in [s.name for s in wb.sheets]:
                wb.sheets[new_sheet_name].delete()
        except Exception:
            pass
        
        # Create new sheet
        new_sheet = wb.sheets.add(name=new_sheet_name, after=sheet)
        
        # Write filtered data
        new_sheet.range('A1').value = result_df.columns.tolist()
        new_sheet.range('A2').value = result_df.values.tolist()
        
        # Auto-fit columns
        new_sheet.autofit()
        
        records_count = len(filtered_df)
        groups_count = len(grp_sum_rows) if numeric_cols else 0
        return f"✓ Created sheet '{new_sheet_name}' with {records_count} records, {groups_count} group sums, and 1 total"
        
    except Exception as e:
        raise Exception(f"AutoSum failed: {str(e)}")


if __name__ == '__main__':
    """For testing purposes."""
    try:
        result = autosum()
        print(result)
    except Exception as e:
        print(f"Error: {e}")
