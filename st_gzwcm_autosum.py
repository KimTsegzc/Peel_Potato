"""
ST_GZWCM AutoSum - Filter and summarize employees based on employee list.
Loads emp.xlsx (or emp_embbed.xlsx as fallback) and filters active workbook data.
"""
import os
import xlwings as xw
import pandas as pd


def load_employee_list():
    """Load employee list from emp.xlsx or emp_embbed.xlsx.
    
    Returns:
        DataFrame with columns: grp, emp
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Try emp.xlsx first
    emp_file = os.path.join(script_dir, 'data', 'emp.xlsx')
    if os.path.exists(emp_file):
        try:
            df = pd.read_excel(emp_file, sheet_name='emplist')
            if 'grp' in df.columns and 'emp' in df.columns:
                return df[['grp', 'emp']]
        except Exception as e:
            print(f"Error loading emp.xlsx: {e}")
    
    # Fallback to emp_embbed.xlsx
    emp_embbed_file = os.path.join(script_dir, 'data', 'emp_embbed.xlsx')
    if os.path.exists(emp_embbed_file):
        try:
            df = pd.read_excel(emp_embbed_file, sheet_name='emplist')
            if 'grp' in df.columns and 'emp' in df.columns:
                return df[['grp', 'emp']]
        except Exception as e:
            print(f"Error loading emp_embbed.xlsx: {e}")
            raise Exception("Could not load employee list from emp.xlsx or emp_embbed.xlsx")
    
    raise Exception("Employee list file not found (emp.xlsx or emp_embbed.xlsx)")


def autosum():
    """Main autosum function.
    
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
        emplist = load_employee_list()
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
            if col and str(col).lower() in ['emp', 'employee', 'name', '员工', '姓名']:
                emp_col = col
                break
        
        if emp_col is None:
            raise Exception("Could not find employee column in active sheet. Expected column named 'emp', 'employee', 'name', '员工', or '姓名'")
        
        # Filter data to include only employees in emplist
        filtered_df = df[df[emp_col].isin(emp_names)]
        
        if filtered_df.empty:
            raise Exception("No matching employees found in active sheet")
        
        # Create new sheet with filtered data
        new_sheet_name = f"{active_sheet.Name}_filtered"
        
        # Remove existing sheet with the same name if it exists
        try:
            if new_sheet_name in [s.name for s in wb.sheets]:
                wb.sheets[new_sheet_name].delete()
        except Exception:
            pass
        
        # Create new sheet
        new_sheet = wb.sheets.add(name=new_sheet_name, after=sheet)
        
        # Write filtered data
        new_sheet.range('A1').value = filtered_df.columns.tolist()
        new_sheet.range('A2').value = filtered_df.values.tolist()
        
        # Auto-fit columns
        new_sheet.autofit()
        
        return f"✓ Created sheet '{new_sheet_name}' with {len(filtered_df)} filtered employees"
        
    except Exception as e:
        raise Exception(f"AutoSum failed: {str(e)}")


if __name__ == '__main__':
    """For testing purposes."""
    try:
        result = autosum()
        print(result)
    except Exception as e:
        print(f"Error: {e}")
