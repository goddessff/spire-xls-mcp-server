import logging
from typing import Any, Dict, Optional, Callable

from spire.xls import Workbook, Worksheet, Chart, ExcelChartType, LegendPositionType

from spire_xls_mcp.utils.exceptions import ChartError
from .workbook import get_or_create_workbook
from spire_xls_mcp.utils.cell_utils import EnumMapper

logger = logging.getLogger(__name__)




def _configure_waterfall_chart(chart: Chart, sheet: Worksheet, chart_options: Dict[str, Any]):
    """Configures a waterfall chart."""
    chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.ShowCategoryName = True
    chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.ShowSeriesName = True


def _configure_bubble_chart(chart: Chart, sheet: Worksheet, chart_options: Dict[str, Any]):
    """Configures a bubble chart."""
    chart.Series.Clear()
    chart.Series.Add()
    bubbles = chart_options.get('bubbles')
    if bubbles:
        chart.Series[0].Bubbles = sheet.Range[bubbles]


# Chart configuration mapping
CHART_CONFIGURATORS: Dict[ExcelChartType, Callable[[Chart, Worksheet, Dict[str, Any]], None]] = {
    ExcelChartType.WaterFall: _configure_waterfall_chart,
    ExcelChartType.Bubble: _configure_bubble_chart,
}

def _configure_default_chart(chart: Chart, sheet: Worksheet, chart_options: Dict[str, Any]):
    """Configures a default chart with axis titles."""
    x_axis = chart_options.get('x_axis')
    y_axis = chart_options.get('y_axis')
    
def create_chart_in_sheet(
        input_filepath: str,
        output_filepath: str,
        data_sheet_name: str,
        chart_sheet_name: str,
        data_range: str,
        chart_type: str,
        target_cell: str,
        chart_options: Optional[Dict[str, Any]] = None
) -> dict[str, Any]:
    """Create chart in sheet with enhanced styling options"""
    chart_options = chart_options or {}
    try:
        if not input_filepath and output_filepath:
            input_filepath = output_filepath
        elif input_filepath and not output_filepath:
            output_filepath = input_filepath
        elif not input_filepath and not output_filepath:
            raise ChartError("input_filepath or output_filepath must be provided")
        wb = get_or_create_workbook(input_filepath)
        data_sheet = None
        chart_sheet = None
        # Find or create sheet
        for ws in wb.Worksheets:
            if ws.Name == data_sheet_name:
                data_sheet = ws
            if ws.Name == chart_sheet_name:
                chart_sheet = ws
        if data_sheet is None:
            raise ChartError(f"sheet '{data_sheet_name}' not found")
        if chart_sheet is None:
            chart_sheet = wb.CreateEmptySheet(chart_sheet_name)

        chart = chart_sheet.Charts.Add()
        chart.ChartType = EnumMapper.get_chart_type_enum(chart_type)
        chart.DataRange = data_sheet.Range[data_range]

        target_range = chart_sheet.Range[target_cell]
        chart.LeftColumn = target_range.Column
        chart.TopRow = target_range.Row

        if title := chart_options.get("title"):
            chart.ChartTitle = title

        # Set Series Data
        x_axis = chart_options.get('x_axis')
        y_axis = chart_options.get('y_axis')
        cs = chart.Series[0]
        if x_axis:
            chart.SeriesDataFromRange = False
            cs.CategoryLabels = data_sheet.Range[x_axis]
        if y_axis:
            chart.SeriesDataFromRange = False
            cs.Values = data_sheet.Range[y_axis]
            
        # Get the specific configurator for the chart type, or the default one
        configurator = CHART_CONFIGURATORS.get(chart.ChartType, _configure_default_chart)
        # Pass all relevant options to the configurator
        options_for_configurator = chart_options.copy()
        options_for_configurator['data_range'] = data_range
        configurator(chart, data_sheet, options_for_configurator)
        

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
                series.Format.Options.IsVaryColor = True
                series.DataPoints.DefaultDataPoint.DataLabels.HasValue = style['has_data_labels']

        chart.Width = width
        chart.Height = height

        wb.SaveToFile(output_filepath)
        return {"message": "Chart created successfully"}

    except Exception as e:
        logger.error(f"Failed to create chart: {e}")
        raise ChartError(f"Failed to create chart: {e!s}")