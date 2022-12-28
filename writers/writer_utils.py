from xlsxwriter import Workbook, utility
from writers.tables import Table, Series, Cell

Worksheet = Workbook.worksheet_class


def write_cell(workbook: Workbook, worksheet: Worksheet, cell: Cell):
    """Write `cell` to `worksheet`."""
    format = workbook.add_format(cell.format)
    # default format for all cells
    format.set_shrink()
    worksheet.write(cell.row, cell.col, cell.value, format)


def write_list_of_cells(workbook: Workbook, worksheet: Worksheet, cells: list[Cell]):
    """Write each cell in `cells` to `worksheet`."""
    for cell in cells:
        write_cell(workbook, worksheet, cell)


def write_table(workbook: Workbook, worksheet: Worksheet, table: Table):
    """Write the contents of `table` to `worksheet`."""
    table_cols = table.get_cols_as_lists()
    for col in table_cols:
        write_list_of_cells(workbook, worksheet, col)


def create_line_chart_for_table(workbook: Workbook,
                                worksheet: Worksheet,
                                worksheet_name: str,
                                table: Table,
                                series_list: list[Series],
                                chart_row: int,
                                chart_col: int):
    """Create a line chart for `worksheet` from the series in `series_list` and the x_axis in `table`."""
    chart = workbook.add_chart({'type': 'line'})
    chart.set_size({'width': 900, 'height': 500})

    xl_timespan_col = utility.xl_col_to_name(table.start_col)
    xl_data_start_row = table.start_row + 2
    xl_data_end_row = table.start_row + 1 + table.get_num_data_rows()

    col_reference_str = '={}!${}${}:${}${}'
    x_axis_col = col_reference_str.format(
        worksheet_name, xl_timespan_col, xl_data_start_row, xl_timespan_col, xl_data_end_row)
    for series in series_list:
        xl_col = utility.xl_col_to_name(series.col)
        values_col = col_reference_str.format(
            worksheet_name, xl_col, xl_data_start_row, xl_col, xl_data_end_row)
        chart.add_series({
            'categories':   x_axis_col,
            'values':       values_col,
            'name':         series.category,
            'line':         series.get_line_styles(),
            'marker':       {'type': 'square'}
        })
    chart_cell = utility.xl_rowcol_to_cell(chart_row, chart_col)
    worksheet.insert_chart(chart_cell, chart)
