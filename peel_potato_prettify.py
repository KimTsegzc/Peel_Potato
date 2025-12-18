"""
Chart formatting module for Peel Potato.
Handles default styling and formatting for Excel charts.
"""

try:
    from win32com.client import constants as xlconst
except Exception:
    xlconst = None


def apply_chart_formatting(chart):
    """Apply default formatting to charts:
    1) Legend on, position at top
    2) Data labels on, above, number format 0.0
    3) All fonts 12pt, Microsoft YaHei UI
    """
    try:
        # Ensure win32 constants available
        if xlconst is None:
            from win32com.client import constants as xlconst_local
            _xl = xlconst_local
        else:
            _xl = xlconst

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
