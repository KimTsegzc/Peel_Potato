"""
Chart Builder for Peel Potato.
Handles all chart creation and modification logic, isolated from UI.
"""
import peel_potato_prettify

try:
    from win32com.client import constants as xlconst
except Exception:
    xlconst = None


class ChartBuilder:
    """Responsible for creating and modifying Excel charts."""
    
    def __init__(self):
        self._last_chart = None
    
    def create(self, sheet, dim_range, value_ranges, chart_type, multi_mode, modify=False):
        """Create or modify a chart on the given sheet.
        
        Args:
            sheet: xlwings Sheet object
            dim_range: xlwings Range for dimension (X-axis/categories)
            value_ranges: List of xlwings Range objects for values
            chart_type: String chart type (e.g., "Line", "Column", "Pie")
            multi_mode: String mode (e.g., "Clustered", "Stacked")
            modify: If True, modify existing chart instead of creating new
            
        Returns:
            tuple: (chart_object, dim_name, value_names_list, log_messages)
        """
        log_messages = []
        sht_api = sheet.api
        
        # Ensure win32 constants available
        if xlconst is None:
            from win32com.client import constants as xlconst_local
            _xl = xlconst_local
        else:
            _xl = xlconst
        
        # Create or reuse chart object
        chart = None
        if modify and self._last_chart is not None:
            # Validate that the chart still exists
            try:
                _ = self._last_chart.ChartType
                chart = self._last_chart
                log_messages.append("Modifying existing chart")
            except Exception:
                # Chart no longer exists, create new
                chart = None
                self._last_chart = None
                log_messages.append("Previous chart not found, creating new")
        
        if chart is None:
            # Place chart at fixed position
            left = 50
            top = 20
            width = 520
            height = 320
            
            chart_objects = sht_api.ChartObjects()
            chart_obj = chart_objects.Add(left, top, width, height)
            chart = chart_obj.Chart
            log_messages.append(f"Created new chart object at ({left}, {top})")
        
        # Remember last chart
        self._last_chart = chart
        
        # Determine Excel chart constant and set type
        chart_const = self._get_chart_constant(chart_type, multi_mode, _xl)
        log_messages.append(f"Creating {chart_type} chart with {multi_mode} mode")
        
        try:
            chart.ChartType = chart_const
        except Exception as e:
            log_messages.append(f"Warning: Could not set chart type: {e}")
        
        # Build the chart based on type
        chart_type_lower = chart_type.lower()
        
        if 'scatter' in chart_type_lower:
            self._build_scatter_chart(chart, dim_range, value_ranges, _xl, log_messages)
        else:
            self._build_standard_chart(chart, sheet, dim_range, value_ranges, chart_type, 
                                      multi_mode, modify, _xl, log_messages)
        
        # Set chart title
        chart.HasTitle = True
        chart.ChartTitle.Text = f"{chart_type} — Peel Potato"
        
        # Apply formatting and get names
        dim_name, value_names = peel_potato_prettify.apply_chart_formatting(
            chart, dim_range, value_ranges
        )
        
        if dim_name and value_names:
            log_messages.append(f"📊 Chart title set: <b>{value_names[0]} by {dim_name}</b>")
        
        action = "modified" if modify else "created"
        log_messages.append(f"Chart {action} successfully! Applied formatting.")
        
        return chart, dim_name, value_names, log_messages
    
    def _build_scatter_chart(self, chart, dim_range, value_ranges, _xl, log_messages):
        """Build a scatter chart."""
        if len(value_ranges) == 0:
            raise ValueError("No value ranges parsed for scatter chart")
        
        x_range = dim_range.api if dim_range is not None else value_ranges[0].api
        y_range = value_ranges[0].api if dim_range is not None else (
            value_ranges[1].api if len(value_ranges) > 1 else None
        )
        
        if y_range is None:
            raise ValueError("Scatter chart needs two columns (X and Y)")
        
        series = chart.SeriesCollection().NewSeries()
        series.XValues = x_range
        series.Values = y_range
        chart.ChartType = _xl.xlXYScatter
        
        log_messages.append(f"Added scatter series with X and Y ranges")
    
    def _build_standard_chart(self, chart, sheet, dim_range, value_ranges, chart_type, 
                             multi_mode, modify, _xl, log_messages):
        """Build standard charts (line, bar, column, area, pie, etc.)."""
        if not value_ranges:
            raise ValueError("No value ranges parsed for chart")
        
        # Determine header row
        header_row = self._find_header_row(sheet, value_ranges)
        
        # If modifying, clear existing series first
        if modify:
            try:
                while chart.SeriesCollection().Count > 0:
                    chart.SeriesCollection(1).Delete()
                log_messages.append("Cleared existing series")
            except Exception:
                pass
        
        # Create series for each value range
        chart_type_lower = chart_type.lower()
        for idx, vr in enumerate(value_ranges):
            try:
                # Pie charts only use first value column
                if 'pie' in chart_type_lower and idx > 0:
                    break
                
                s = chart.SeriesCollection().NewSeries()
                s.Values = vr.api
                
                if dim_range is not None:
                    s.XValues = dim_range.api
                
                # Set series name from header row
                try:
                    name_cell = sheet.api.Cells(header_row, vr.api.Column)
                    name_val = name_cell.Value
                    if name_val is not None:
                        s.Name = str(name_val)
                        log_messages.append(f"  Series {idx+1}: {name_val}")
                except Exception:
                    pass
                    
            except Exception as e:
                log_messages.append(f"Warning: Could not create series {idx+1}: {e}")
        
        # Set chart subtype
        try:
            chart.ChartType = self._get_chart_constant(chart_type, multi_mode, _xl)
        except Exception:
            pass
        
        log_messages.append(f"Added {len(value_ranges)} series to chart")
    
    def _find_header_row(self, sheet, value_ranges):
        """Find the header row (row immediately above data)."""
        header_row = 1
        
        try:
            if value_ranges and len(value_ranges) > 0:
                first_range_api = value_ranges[0].api
                data_start_row = first_range_api.Row
                header_row = max(1, data_start_row - 1)
            else:
                # Fallback to used range
                used = sheet.api.UsedRange
                header_row = used.Row
        except Exception:
            header_row = 1
        
        return header_row
    
    def _get_chart_constant(self, chart_text, mode_text, _xl):
        """Map chart type + mode to Excel ChartType constant."""
        ct = chart_text.lower()
        m = mode_text.lower() if mode_text else ''
        
        try:
            if 'line' in ct:
                if 'stacked' in m and '100' in m:
                    return getattr(_xl, 'xlLineStacked100', 
                                 getattr(_xl, 'xlLineStacked', _xl.xlLine))
                if 'stacked' in m:
                    return getattr(_xl, 'xlLineStacked', _xl.xlLine)
                return getattr(_xl, 'xlLine', getattr(_xl, 'xlLineMarkers', _xl.xlLine))
            
            if 'column' in ct:
                if '100' in m:
                    return getattr(_xl, 'xlColumnStacked100', 
                                 getattr(_xl, 'xlColumnStacked', _xl.xlColumnClustered))
                if 'stack' in m:
                    return getattr(_xl, 'xlColumnStacked', _xl.xlColumnClustered)
                return getattr(_xl, 'xlColumnClustered', _xl.xlColumnClustered)
            
            if 'bar' in ct:
                if '100' in m:
                    return getattr(_xl, 'xlBarStacked100', 
                                 getattr(_xl, 'xlBarStacked', _xl.xlBarClustered))
                if 'stack' in m:
                    return getattr(_xl, 'xlBarStacked', _xl.xlBarClustered)
                return getattr(_xl, 'xlBarClustered', _xl.xlBarClustered)
            
            if 'area' in ct:
                if '100' in m:
                    return getattr(_xl, 'xlAreaStacked100', 
                                 getattr(_xl, 'xlAreaStacked', _xl.xlArea))
                if 'stack' in m:
                    return getattr(_xl, 'xlAreaStacked', _xl.xlArea)
                return getattr(_xl, 'xlArea', _xl.xlArea)
            
            if 'pie' in ct:
                if 'doughnut' in m:
                    return getattr(_xl, 'xlDoughnut', _xl.xlPie)
                if 'pie of' in m:
                    return getattr(_xl, 'xlPieOfPie', _xl.xlPie)
                return getattr(_xl, 'xlPie', _xl.xlPie)
            
            if 'scatter' in ct:
                return getattr(_xl, 'xlXYScatter', _xl.xlXYScatter)
            
            if 'radar' in ct:
                return getattr(_xl, 'xlRadar', _xl.xlRadar)
                
        except Exception:
            pass
        
        # Fallback
        return _xl.xlColumnClustered
    
    def get_last_chart(self):
        """Get the last created/modified chart object."""
        return self._last_chart
