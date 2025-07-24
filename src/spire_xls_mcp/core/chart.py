import logging
from typing import Any, Dict, Optional, Callable

from spire.xls import Workbook, Worksheet, Chart, ExcelChartType, LegendPositionType

from spire_xls_mcp.utils.exceptions import ChartError
from .workbook import get_or_create_workbook
from spire_xls_mcp.utils.cell_utils import EnumMapper

logger = logging.getLogger(__name__)


def _configure_pie_chart(chart: Chart, sheet: Worksheet, chart_options: Dict[str, Any]):
    """Configures a pie or doughnut chart."""
    x_axis = chart_options.get('x_axis')
    y_axis = chart_options.get('y_axis')
    if not x_axis or not y_axis:
        raise ChartError("Pie/doughnut chart requires both 'x_axis' and 'y_axis' options.")
    cs = chart.Series[0]
    cs.CategoryLabels = sheet.Range[x_axis]
    cs.Values = sheet.Range[y_axis]
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True


def _configure_waterfall_chart(chart: Chart, sheet: Worksheet, chart_options: Dict[str, Any]):
    """Configures a waterfall chart."""
    chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.HasValue = True
    chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.ShowCategoryName = True
    chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.ShowSeriesName = True


def _configure_bubble_chart(chart: Chart, sheet: Worksheet, chart_options: Dict[str, Any]):
    """Configures a bubble chart."""
    chart.Series.Clear()
    chart.Series.Add()
    chart.SeriesDataFromRange = False
    x_axis = chart_options.get('x_axis')
    y_axis = chart_options.get('y_axis')
    bubbles = chart_options.get('bubbles')
    if bubbles:
        chart.Series[0].Bubbles = sheet.Range[bubbles]
    if x_axis:
        chart.Series[0].CategoryLabels = sheet.Range[x_axis]
    if y_axis:
        chart.Series[0].Values = sheet.Range[y_axis]


def _configure_line_chart(chart: Chart, sheet: Worksheet, chart_options: Dict[str, Any]):
    """Configures a line chart with axis titles."""
    for cs in chart.Series:
        cs.Format.Options.IsVaryColor = True
        cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
        
def _configure_scatter_chart(chart: Chart, sheet: Worksheet, chart_options: Dict[str, Any]):
    """Configures a scatter chart."""
    data_range = chart_options.get('data_range')
    x_axis = chart_options.get('x_axis')
    y_axis = chart_options.get('y_axis')

    if chart.Series.Count < 1:
        chart.Series.Add(sheet.Range[data_range])
    if x_axis:
        chart.Series[0].CategoryLabels = sheet.Range[x_axis]
    if y_axis:
        chart.Series[0].Values = sheet.Range[y_axis]


def _configure_area_chart(chart: Chart, sheet: Worksheet, chart_options: Dict[str, Any]):
    """Configures a area chart with axis titles."""
    chart.SeriesDataFromRange = False

# Chart configuration mapping
CHART_CONFIGURATORS: Dict[ExcelChartType, Callable[[Chart, Worksheet, Dict[str, Any]], None]] = {
    ExcelChartType.Pie: _configure_pie_chart,
    ExcelChartType.Doughnut: _configure_pie_chart,
    ExcelChartType.WaterFall: _configure_waterfall_chart,
    ExcelChartType.Bubble: _configure_bubble_chart,
    ExcelChartType.Line: _configure_line_chart,
    ExcelChartType.ScatterLine: _configure_scatter_chart,
    ExcelChartType.Area: _configure_area_chart,
}

def _configure_default_chart(chart: Chart, sheet: Worksheet, chart_options: Dict[str, Any]):
    """Configures a default chart with axis titles."""
    x_axis = chart_options.get('x_axis')
    y_axis = chart_options.get('y_axis')
    for cs in chart.Series:
        cs.Format.Options.IsVaryColor = True
        cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
    
def create_chart_in_sheet(
        filepath: str,
        sheet_name: str,
        data_range: str,
        chart_type: str,
        target_cell: str,
        chart_options: Optional[Dict[str, Any]] = None
) -> dict[str, Any]:
    """Create chart in sheet with enhanced styling options"""
    chart_options = chart_options or {}
    try:
        wb = get_or_create_workbook(filepath)
        sheet = wb.Worksheets[sheet_name] if sheet_name in [ws.Name for ws in wb.Worksheets] else wb.CreateEmptySheet(
            sheet_name)

        chart = sheet.Charts.Add()
        chart.ChartType = EnumMapper.get_chart_type_enum(chart_type)
        chart.DataRange = sheet.Range[data_range]

        target_range = sheet.Range[target_cell]
        chart.LeftColumn = target_range.Column
        chart.TopRow = target_range.Row

        if title := chart_options.get("title"):
            chart.ChartTitle = title

        # Get the specific configurator for the chart type, or the default one
        configurator = CHART_CONFIGURATORS.get(chart.ChartType, _configure_default_chart)
        # Pass all relevant options to the configurator
        options_for_configurator = chart_options.copy()
        options_for_configurator['data_range'] = data_range
        configurator(chart, sheet, options_for_configurator)

        # Apply common styling
        style = chart_options.get('style', {})
        width = style.get('width', 480)
        height = style.get('height', 300)

        if 'legend_position' in style:
            positions = {
                'right': LegendPositionType.Right,
                'left': LegendPositionType.Left,
                'top': LegendPositionType.Top,
                'bottom': LegendPositionType.Bottom
            }
            if style['legend_position'] in positions:
                chart.Legend.Position = positions[style['legend_position']]

        if 'has_legend' in style:
            chart.Legend.Visible = style['has_legend']

        if 'has_data_labels' in style:
            for series in chart.Series:
                series.DataLabels.HasValue = style['has_data_labels']

        chart.Width = width
        chart.Height = height

        wb.SaveToFile(filepath)
        return {"message": "Chart created successfully"}

    except Exception as e:
        logger.error(f"Failed to create chart: {e}")
        raise ChartError(f"Failed to create chart: {e!s}")