"""
Chart Controller for Peel Potato.
Orchestrates services to handle chart creation and modification requests.
Acts as the business logic layer between UI and services.
"""
from dataclasses import dataclass
from typing import List, Optional, Any


@dataclass
class ChartResult:
    """Result of a chart operation."""
    success: bool
    chart: Any
    dim_name: Optional[str]
    value_names: List[str]
    log_messages: List[str]
    error_message: Optional[str] = None


class ChartController:
    """Orchestrates chart creation and modification using various services."""
    
    def __init__(self, excel_adapter, range_parser, chart_builder):
        """Initialize controller with required services.
        
        Args:
            excel_adapter: ExcelAdapter instance for Excel operations
            range_parser: RangeParser instance for range parsing
            chart_builder: ChartBuilder instance for chart creation
        """
        self.excel = excel_adapter
        self.parser = range_parser
        self.builder = chart_builder
    
    def create_chart(self, dim_text, values_text, chart_type, multi_mode):
        """Create a new chart based on user inputs.
        
        Args:
            dim_text: Dimension/X-axis range specification
            values_text: Values range specification
            chart_type: Chart type string (e.g., "Line", "Column")
            multi_mode: Multi-value mode (e.g., "Clustered", "Stacked")
            
        Returns:
            ChartResult object with operation results
        """
        return self._execute_chart_operation(
            dim_text, values_text, chart_type, multi_mode, modify=False
        )
    
    def modify_chart(self, dim_text, values_text, chart_type, multi_mode):
        """Modify the existing chart with new data/settings.
        
        Args:
            dim_text: Dimension/X-axis range specification
            values_text: Values range specification
            chart_type: Chart type string (e.g., "Line", "Column")
            multi_mode: Multi-value mode (e.g., "Clustered", "Stacked")
            
        Returns:
            ChartResult object with operation results
        """
        return self._execute_chart_operation(
            dim_text, values_text, chart_type, multi_mode, modify=True
        )
    
    def _execute_chart_operation(self, dim_text, values_text, chart_type, multi_mode, modify=False):
        """Execute chart creation or modification operation.
        
        This is the main orchestration method that coordinates all services.
        """
        log_messages = []
        
        try:
            # Validate inputs
            validation_error = self.validate_inputs(dim_text, values_text, chart_type)
            if validation_error:
                return ChartResult(
                    success=False,
                    chart=None,
                    dim_name=None,
                    value_names=[],
                    log_messages=[],
                    error_message=validation_error
                )
            
            # Get active sheet
            log_messages.append("Getting active Excel sheet...")
            sheet = self.excel.get_active_sheet()
            if sheet is None:
                return ChartResult(
                    success=False,
                    chart=None,
                    dim_name=None,
                    value_names=[],
                    log_messages=log_messages,
                    error_message="No active Excel sheet found. Please open an Excel workbook."
                )
            
            # Enter performance mode
            log_messages.append("Entering performance mode...")
            app_api, saved_state = self.excel.begin_performance_mode(sheet)
            
            try:
                # Parse dimension range
                log_messages.append(f"Parsing dimension: {dim_text}")
                dim_range = self.parser.parse_dim(dim_text, sheet)
                
                # Parse values ranges
                log_messages.append(f"Parsing values: {values_text}")
                ref_rows = None
                if dim_range is not None:
                    dra = dim_range.api
                    ref_rows = (dra.Row, dra.Row + dra.Rows.Count - 1)
                
                value_ranges = self.parser.parse_values(values_text, sheet, ref_rows=ref_rows)
                log_messages.append(f"Found {len(value_ranges)} value range(s)")
                
                # If dim was just a column letter, infer its range from value ranges
                if dim_range is None and dim_text and dim_text.strip().isalpha():
                    log_messages.append(f"Inferring dimension range from column {dim_text}")
                    dim_range = self.parser.infer_dim_range_from_column(
                        dim_text, sheet, value_ranges
                    )
                
                # Log detected names
                if dim_range:
                    try:
                        import peel_potato_prettify
                        dim_name = peel_potato_prettify.reset_title_name(dim_range)
                        if dim_name:
                            log_messages.append(f"✓ Dimension: <b>{dim_name}</b>")
                    except Exception:
                        pass
                
                if value_ranges:
                    try:
                        import peel_potato_prettify
                        for idx, vr in enumerate(value_ranges):
                            value_name = peel_potato_prettify.reset_title_name(vr)
                            if value_name:
                                log_messages.append(f"  ✓ Value {idx+1}: <b>{value_name}</b>")
                    except Exception:
                        pass
                
                # Create/modify chart
                action = "Modifying" if modify else "Creating"
                log_messages.append(f"{action} {chart_type} chart...")
                
                chart, dim_name, value_names, builder_logs = self.builder.create(
                    sheet, dim_range, value_ranges, chart_type, multi_mode, modify=modify
                )
                
                # Merge builder logs
                log_messages.extend(builder_logs)
                
                return ChartResult(
                    success=True,
                    chart=chart,
                    dim_name=dim_name,
                    value_names=value_names,
                    log_messages=log_messages,
                    error_message=None
                )
                
            finally:
                # Always restore Excel settings
                self.excel.end_performance_mode(app_api, saved_state)
                log_messages.append("Performance mode restored")
        
        except ValueError as e:
            # User input errors
            return ChartResult(
                success=False,
                chart=None,
                dim_name=None,
                value_names=[],
                log_messages=log_messages,
                error_message=str(e)
            )
        except Exception as e:
            # Unexpected errors
            return ChartResult(
                success=False,
                chart=None,
                dim_name=None,
                value_names=[],
                log_messages=log_messages,
                error_message=f"Unexpected error: {str(e)}"
            )
    
    def validate_inputs(self, dim_text, values_text, chart_type):
        """Validate user inputs before processing.
        
        Args:
            dim_text: Dimension range specification
            values_text: Values range specification
            chart_type: Chart type string
            
        Returns:
            Error message string if validation fails, None if valid
        """
        dim_text = (dim_text or '').strip()
        values_text = (values_text or '').strip()
        chart_type_lower = chart_type.lower()
        
        # Scatter charts need both X and Y
        if 'scatter' in chart_type_lower:
            if not dim_text or not values_text:
                return "Scatter charts require both Dim (X) and Values (Y) to be specified."
        else:
            # Default charts need at least Dim and Values
            if not dim_text or not values_text:
                return "Please specify both Dim (categories) and Values (data) for the chart."
        
        return None
    
    def get_active_sheet_info(self):
        """Get information about the currently active Excel sheet.
        
        Returns:
            tuple: (workbook_name, sheet_name) or (None, None) if no Excel
        """
        return self.excel.get_active_workbook_info()
    
    def is_excel_available(self):
        """Check if Excel is available.
        
        Returns:
            bool: True if Excel is available
        """
        return self.excel.validate_excel_available()
