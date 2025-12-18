"""
Chart formatting module for Peel Potato.
Handles default styling and formatting for Excel charts.
"""

try:
    from win32com.client import constants as xlconst
except Exception:
    xlconst = None


def apply_chart_formatting(chart, dim_range=None, value_ranges=None):
    """Apply default formatting to charts:
    1) Legend on, position at top
    2) Data labels on, above, number format 0.0
    3) All fonts 12pt, Microsoft YaHei UI
    4) Chart title to left/top, formatted as "Value" by Dim
    
    Returns:
        tuple: (dim_name, value_names_list) or (None, None) if names couldn't be extracted
    """
    dim_name = None
    value_names = []
    
    try:
        # Ensure win32 constants available
        if xlconst is None:
            from win32com.client import constants as xlconst_local
            _xl = xlconst_local
        else:
            _xl = xlconst

        # Set chart title: "Value" by Dim with left/top positioning
        try:
            dim_name = reset_title_name(dim_range) if dim_range else None
            value_name = reset_title_name(value_ranges[0]) if value_ranges and len(value_ranges) > 0 else None
            
            # Collect all value names for logging
            if value_ranges:
                for vr in value_ranges:
                    vn = reset_title_name(vr)
                    if vn:
                        value_names.append(vn)
            
            if value_name and dim_name:
                chart.HasTitle = True
                chart.ChartTitle.Text = f'{value_name} by {dim_name}'
                # Position title to left/top
                chart.ChartTitle.Left = 0
                chart.ChartTitle.Top = 0
            elif chart.HasTitle:
                # Keep existing title but position it
                chart.ChartTitle.Left = 0
                chart.ChartTitle.Top = 0
        except Exception:
            pass

        # 1) Legend: on and position at top
        try:
            chart.HasLegend = True
            chart.Legend.Position = _xl.xlLegendPositionTop
            # Set legend font
            chart.Legend.Font.Size = 12
            chart.Legend.Font.Name = "Microsoft YaHei UI"
        except Exception:
            pass

        # 2) Data labels: on, above, number format 0.0
        try:
            series_count = chart.SeriesCollection().Count
            for i in range(1, series_count + 1):  # COM collections are 1-indexed
                try:
                    series = chart.SeriesCollection(i)
                    series.HasDataLabels = True
                except Exception:
                    pass
            
            # Apply formatting after all series have labels enabled
            for i in range(1, series_count + 1):
                try:
                    series = chart.SeriesCollection(i)
                    datalabels = series.DataLabels()
                    datalabels.Position = _xl.xlLabelPositionAbove
                    datalabels.NumberFormat = "0.0"
                    datalabels.Font.Size = 12
                    datalabels.Font.Name = "Microsoft YaHei UI"
                except Exception as e:
                    pass
        except Exception as e:
            pass

        # 3) Chart title font
        try:
            if chart.HasTitle:
                chart.ChartTitle.Font.Size = 12
                chart.ChartTitle.Font.Name = "Microsoft YaHei UI"
        except Exception:
            pass

        # Axes fonts (if applicable)
        try:
            # Category axis
            chart.Axes(_xl.xlCategory).TickLabels.Font.Size = 12
            chart.Axes(_xl.xlCategory).TickLabels.Font.Name = "Microsoft YaHei UI"
        except Exception:
            pass

        try:
            # Value axis
            chart.Axes(_xl.xlValue).TickLabels.Font.Size = 12
            chart.Axes(_xl.xlValue).TickLabels.Font.Name = "Microsoft YaHei UI"
        except Exception:
            pass

    except Exception:
        pass
    
    return (dim_name, value_names)


def reset_title_name(cell_range):
    """Find the first string (non-numeric) value for chart title.
    Searches from row 1 downward in the same column as the provided range.
    For multi-column ranges, only searches the first column.
    """
    if cell_range is None:
        return None
    
    try:
        # Get the range API object
        range_api = cell_range.api if hasattr(cell_range, 'api') else cell_range
        
        # Get the first column of the range
        first_col = range_api.Column
        
        # Get the worksheet
        worksheet = range_api.Worksheet
        
        # Search from row 1 downward until we find a string
        max_search_rows = 100  # Reasonable limit to avoid searching entire sheet
        for row_num in range(1, max_search_rows + 1):
            try:
                cell_value = worksheet.Cells(row_num, first_col).Value
                if isinstance(cell_value, str) and cell_value.strip():
                    return cell_value.strip()
            except Exception:
                continue
        
        return None
    except Exception:
        return None
