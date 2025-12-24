"""
ST_GZWCM AutoSLC - Auto Select Columns based on dictionary.
Loads dict.xlsx (or dict_embbed.xlsx as fallback) and filters/renames columns in active workbook.
"""
import os
import xlwings as xw
import pandas as pd


def load_column_dict():
    """Load column dictionary from dict.xlsx or dict_embbed.xlsx.
    
    Returns:
        DataFrame with columns: old, new
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Try dict.xlsx first
    dict_file = os.path.join(script_dir, 'data', 'dict.xlsx')
    if os.path.exists(dict_file):
        try:
            df = pd.read_excel(dict_file, sheet_name='columnlist')
            if 'old' in df.columns and 'new' in df.columns:
                return df[['old', 'new']]
        except Exception as e:
            print(f"Error loading dict.xlsx: {e}")
    
    # Fallback to dict_embbed.xlsx
    dict_embbed_file = os.path.join(script_dir, 'data', 'dict_embbed.xlsx')
    if os.path.exists(dict_embbed_file):
        try:
            df = pd.read_excel(dict_embbed_file, sheet_name='columnlist')
            if 'old' in df.columns and 'new' in df.columns:
                return df[['old', 'new']]
        except Exception as e:
            print(f"Error loading dict_embbed.xlsx: {e}")
            raise Exception("Could not load column dictionary from dict.xlsx or dict_embbed.xlsx")
    
    raise Exception("Column dictionary file not found (dict.xlsx or dict_embbed.xlsx)")


def autoslc():
    """Main autoslc function.
    
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
        columnlist = load_column_dict()
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
        
        # Filter columns
        filtered_df = df[columns_to_keep]
        
        # Rename columns
        filtered_df = filtered_df.rename(columns=rename_map)
        
        # Create new sheet with filtered and renamed data
        new_sheet_name = f"{active_sheet.Name}_slc"
        
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
