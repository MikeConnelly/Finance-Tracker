from typing import Dict

from xlsxwriter import Workbook, utility
from writers.category_totals_table import CategoryTotalsTable, DEFAULT_STYLES

Worksheet = Workbook.worksheet_class


def write_row(worksheet, row_index: int, col_index: int, row_data: list[str]):
    """Write all of `row_data` starting at the index `row_index`, `col_index`."""
    for cell in row_data:
        worksheet.write(row_index, col_index, cell)
        col_index += 1


def write_col(worksheet, row_index: int, col_index: int, col_data: list[str]):
    """Write all of `col_data` starting at the index `row_index`, `col_index`."""
    for cell in col_data:
        worksheet.write(row_index, col_index, cell)
        row_index += 1


def write_data_cells_by_col(worksheet, row_start_index: int, col_start_index: int, data: list[list[str]]):
    """Write all of `data` in a grid, column by column."""
    for col_values in data:
        write_col(worksheet, row_start_index, col_start_index, col_values)
        col_start_index += 1


def write_data_cells_by_row(worksheet, row_start_index: int, col_start_index: int, data: list[list[str]]):
    """Write all of `data` in a grid, row by row."""
    for row_values in data:
        write_row(worksheet, row_start_index, col_start_index, row_values)
        row_start_index += 1


def write_table(workbook: Workbook,
                worksheet: Worksheet,
                start_row_index: int,
                start_col_index: int,
                table: CategoryTotalsTable,
                styles_map: Dict[str, Dict[str, str]] = {},
                include_sum_row: bool = True,
                include_sum_col: bool = True):
    """Write all columns of a table, including the header row."""
    # write time col
    default_styles_index = 0
    time_cell_row_index = start_row_index + 1
    for time in table.timespan_col:
        worksheet.write(time_cell_row_index, start_col_index, time)
        time_cell_row_index += 1

    data_col_index = start_col_index + 1
    for col in table.data_cols:
        data_cell_row_index = start_row_index + 1
        header_cell_format = workbook.add_format()
        cell_format = workbook.add_format()
        alt_cell_format = workbook.add_format()
        header = col.header

        if header in styles_map.keys():
            header_cell_format.set_bg_color(styles_map[header]['header_bg_color'])
            cell_format.set_bg_color(styles_map[header]['bg_color'])
            alt_cell_format.set_bg_color(styles_map[header]['alt_bg_color'])
        else:
            header_cell_format.set_bg_color(DEFAULT_STYLES[default_styles_index]['header_bg_color'])
            cell_format.set_bg_color(DEFAULT_STYLES[default_styles_index]['bg_color'])
            alt_cell_format.set_bg_color(DEFAULT_STYLES[default_styles_index]['alt_bg_color'])
            default_styles_index = (default_styles_index + 1) % len(DEFAULT_STYLES)
        worksheet.write(start_row_index, data_col_index, header, header_cell_format)

        for cell_data in col.data:
            if data_cell_row_index % 2 == 0:
                worksheet.write(data_cell_row_index, data_col_index, cell_data, alt_cell_format)
            else:
                worksheet.write(data_cell_row_index, data_col_index, cell_data, cell_format)
            data_cell_row_index += 1
        data_col_index += 1

    if include_sum_row:
        sum_row_index = start_row_index + table.num_data_rows + 1
        worksheet.write(sum_row_index, start_col_index, 'Total')  # TODO: move this to the write_sum_row fun
        write_sum_row(worksheet, sum_row_index, start_row_index + 1,
                      table.num_data_rows, start_col_index + 1, table.num_data_cols)

    if include_sum_col:
        sum_col_index = start_col_index + table.num_data_cols + 1
        worksheet.write(start_row_index, sum_col_index, 'Total')
        write_sum_col(worksheet, sum_col_index, start_row_index + 1,
                      table.num_data_rows, start_col_index + 1, table.num_data_cols)


def write_sum_row(worksheet,
                  sum_row_index: int,
                  row_start_zero_indexed: int,
                  num_rows: int,
                  col_start_zero_indexed: int,
                  num_cols: int):
    """
    Write a row that contains the SUM formula for adding a number of rows.

    :param Any worksheet: Worksheet to write contents.
    :param int sum_row_index: The sum row index.
    :param int row_start_zero_indexed: The row index to start the SUM formula.
    :param int num_sum_rows: The number of rows to sum.
    :param int col_start_zero_indexed: The column to start the sum row from.
    :param int num_cols: The number of columns in the sum row.
    """
    # convert zero-indexed cell values to actual cell values
    row_start_index = row_start_zero_indexed + 1
    row_end_index = row_start_zero_indexed + num_rows

    for col_index in range(col_start_zero_indexed, col_start_zero_indexed + num_cols):
        col = utility.xl_col_to_name(col_index)
        sum_formula = '=SUM({}{}:{}{})'.format(col, row_start_index, col, row_end_index)
        worksheet.write(sum_row_index, col_index, sum_formula)


def write_sum_col(worksheet,
                  sum_col_index: int,
                  row_start_zero_indexed: int,
                  num_rows: int,
                  col_start_zero_indexed: int,
                  num_cols: int):
    """
    Write a column that contains the SUM formula for adding a number of columns.

    :param Any worksheet: Worksheet to write contents.
    :param int sum_col_index: The sum column index.
    :param int row_start_zero_indexed: The row index to start the SUM formula.
    :param int num_rows: The number of rows to sum.
    :param int col_start_zero_indexed: The column to start the sum row from.
    :param int num_cols: The number of columns in the sum row.
    """
    # convert zero-indexed cell values to actual cell values
    start_col = utility.xl_col_to_name(col_start_zero_indexed)
    end_col = utility.xl_col_to_name(col_start_zero_indexed + num_cols - 1)

    for row_index in range(row_start_zero_indexed, row_start_zero_indexed + num_rows):
        row = row_index + 1
        sum_formula = '=SUM({}{}:{}{})'.format(start_col, row, end_col, row)
        worksheet.write(row_index, sum_col_index, sum_formula)


def write_subtract_row(worksheet,
                       row_index: int,
                       row1_zero_indexed: int,
                       row2_zero_indexed: int,
                       col_start_zero_indexed: int,
                       num_cols: int):
    """
    Write a row that contains a subtract formula for two given rows.

    :param int row_index: The subtraction row index.
    :param int row1_zero_indexed: The first operand row.
    :param int row2_zero_indexed: The second operand row.
    :param int col_start_zero_indexed: The column to start the subtraction row from.
    :param int num_cols: The number of columns in the subtraction row.
    """
    # convert zero-indexed cell values to actual cell valeus
    row1_index = row1_zero_indexed + 1
    row2_index = row2_zero_indexed + 1

    for col_index in range(col_start_zero_indexed, col_start_zero_indexed + num_cols):
        col = utility.xl_col_to_name(col_index)
        diff_formula = '={}{}-{}{}'.format(col, row1_index, col, row2_index)
        worksheet.write(row_index, col_index, diff_formula)


def create_line_chart_with_series_as_rows(chart,
                                          worksheet_name: str,
                                          x_axis_row_zero_indexed: int,
                                          series_name_col_zero_indexed: int,
                                          data_start_row_zero_indexed: int,
                                          num_data_rows: int,
                                          data_start_col_zero_indexed: int,
                                          num_data_cols: int,
                                          line_styles: dict = {}):
    """
    Populate a line chart with each series being a row.

    :param Any chart: The chart to populate.
    :param str worksheet_name: The worksheet name.
    :param int x_axis_row_zero_indexed: The row that will be the chart's x-axis
    :param int series_name_col_zero_indexed: The column that contains each series' name.
    :param int data_start_row_zero_indexed: The first row of data.
    :param int num_data_rows: The number of rows in the dataset.
    :param int data_start_col_zero_indexed: The first column of data.
    :param int num_data_cols: The number of columns in the dataset.
    :param line_styles: Line styles.
    :type line_styles: dict or None
    """
    # convert zero-indexed cell values to actual cell values
    x_axis_row_index = x_axis_row_zero_indexed + 1
    data_row_index = data_start_row_zero_indexed + 1
    series_name_col = utility.xl_col_to_name(series_name_col_zero_indexed)
    data_start_col = utility.xl_col_to_name(data_start_col_zero_indexed)
    data_end_col = utility.xl_col_to_name(data_start_col_zero_indexed + num_data_cols - 1)

    row_reference_str = '={}!${}${}:${}${}'
    x_axis_row = row_reference_str.format(worksheet_name, data_start_col, x_axis_row_index, data_end_col,
                                          x_axis_row_index)
    for _ in range(num_data_rows):
        values_row = row_reference_str.format(worksheet_name, data_start_col, data_row_index, data_end_col,
                                              data_row_index)
        series_name = '={}!${}${}'.format(worksheet_name, series_name_col, data_row_index)
        chart.add_series({
            'categories':   x_axis_row,
            'values':       values_row,
            'name':         series_name,
            'line':         line_styles,
            'marker':       {'type': 'square'}
        })
        data_row_index += 1


def create_line_chart_with_series_as_cols(chart,
                                          worksheet_name: str,
                                          series_name_row_zero_indexed: int,
                                          x_axis_col_zero_indexed: int,
                                          data_start_row_zero_indexed: int,
                                          num_data_rows: int,
                                          data_start_col_zero_indexed: int,
                                          num_data_cols: int,
                                          line_styles: dict = {}):
    """
    Populate a line chart with each series being a column.

    :param Any chart: The chart to populate.
    :param str worksheet_name: The worksheet name.
    :param int series_name_row_zero_indexed: The row that contains each series' name.
    :param int x_axis_col_zero_indexed: The column that will be the chart's x-axis.
    :param int data_start_row_zero_indexed: The first row of data.
    :param int num_data_rows: The number of rows in the dataset.
    :param int data_start_col_zero_indexed: The first column of data.
    :param int num_data_cols: The number of columns in the dataset.
    :param line_styles: Line styles.
    :type line_styles: dict or None
    """
    # convert zero-indexed cell values to actual cell values
    series_name_row_index = series_name_row_zero_indexed + 1
    data_start_row_index = data_start_row_zero_indexed + 1
    data_end_row_index = data_start_row_zero_indexed + num_data_rows
    x_axis_col_name = utility.xl_col_to_name(x_axis_col_zero_indexed)
    data_col_index = data_start_col_zero_indexed

    col_reference_str = '={}!${}${}:${}${}'
    x_axis_col = col_reference_str.format(worksheet_name, x_axis_col_name, data_start_row_index, x_axis_col_name,
                                          data_end_row_index)
    for _ in range(num_data_cols):
        data_col_name = utility.xl_col_to_name(data_col_index)
        values_col = col_reference_str.format(worksheet_name, data_col_name, data_start_row_index, data_col_name,
                                              data_end_row_index)
        series_name = '={}!${}${}'.format(worksheet_name, data_col_name, series_name_row_index)
        chart.add_series({
            'categories':   x_axis_col,
            'values':       values_col,
            'name':         series_name,
            'line':         line_styles,
            'marker':       {'type': 'square'}
        })
        data_col_index += 1


# TODO: try to create this from a table object so I can use just the topleft coordinate and know all dimensions
def create_line_chart(workbook: Workbook,
                      worksheet: Worksheet,
                      worksheet_name: str,
                      series_name_row_zero_indexed: int,
                      x_axis_col_zero_indexed: int,
                      data_start_row_zero_indexed: int,
                      num_data_rows: int,
                      data_start_col_zero_indexed: int,
                      num_data_cols: int,
                      chart_cell_row_index: int = 0,
                      chart_cell_col_index: int = 0,
                      line_styles: dict = {}):
    """
    Populate a line chart with each series being a column.

    :param Any chart: The chart to populate.
    :param str worksheet_name: The worksheet name.
    :param int series_name_row_zero_indexed: The row that contains each series' name.
    :param int x_axis_col_zero_indexed: The column that will be the chart's x-axis.
    :param int data_start_row_zero_indexed: The first row of data.
    :param int num_data_rows: The number of rows in the dataset.
    :param int data_start_col_zero_indexed: The first column of data.
    :param int num_data_cols: The number of columns in the dataset.
    :param line_styles: Line styles.
    :type line_styles: dict or None
    """
    chart = workbook.add_chart({'type': 'line'})
    chart.set_size({'width': 900, 'height': 500})

    # convert zero-indexed cell values to actual cell values
    series_name_row_index = series_name_row_zero_indexed + 1
    data_start_row_index = data_start_row_zero_indexed + 1
    data_end_row_index = data_start_row_zero_indexed + num_data_rows
    x_axis_col_name = utility.xl_col_to_name(x_axis_col_zero_indexed)
    data_col_index = data_start_col_zero_indexed

    col_reference_str = '={}!${}${}:${}${}'
    x_axis_col = col_reference_str.format(worksheet_name, x_axis_col_name, data_start_row_index, x_axis_col_name,
                                          data_end_row_index)
    for _ in range(num_data_cols):
        data_col_name = utility.xl_col_to_name(data_col_index)
        values_col = col_reference_str.format(worksheet_name, data_col_name, data_start_row_index, data_col_name,
                                              data_end_row_index)
        series_name = '={}!${}${}'.format(worksheet_name, data_col_name, series_name_row_index)
        chart.add_series({
            'categories':   x_axis_col,
            'values':       values_col,
            'name':         series_name,
            'line':         line_styles,
            'marker':       {'type': 'square'}
        })
        data_col_index += 1

    chart_cell = utility.xl_rowcol_to_cell(chart_cell_row_index, chart_cell_col_index)
    worksheet.insert_chart(chart_cell, chart)
