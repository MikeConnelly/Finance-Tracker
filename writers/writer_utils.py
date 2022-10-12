from xlsxwriter import utility

def write_row(worksheet, row_index: int, col_index: int, row_data: list[str]):
    for cell in row_data:
        worksheet.write(row_index, col_index, cell)
        col_index += 1

def write_col(worksheet, row_index: int, col_index: int, col_data: list[str]):
    for cell in col_data:
        worksheet.write(row_index, col_index, cell)
        row_index += 1

def write_data_cells(worksheet, row_start_index: int, col_start_index: int, data: list[list[str]]):
    for col_values in data:
        write_col(worksheet, row_start_index, col_start_index, col_values)
        col_start_index += 1

def write_sum_row(worksheet,
                  row_index: int,
                  sum_rows_start_zero_indexed: int,
                  num_sum_rows: int,
                  col_start_zero_indexed: int,
                  num_cols: int):
    # convert zero-indexed cell values to actual cell values
    sum_rows_start_index = sum_rows_start_zero_indexed + 1
    sum_rows_end_index = sum_rows_start_zero_indexed + num_sum_rows

    for col_index in range(col_start_zero_indexed, col_start_zero_indexed + num_cols):
        col = utility.xl_col_to_name(col_index)
        sum_formula = '=SUM({}{}:{}{})'.format(col, sum_rows_start_index, col, sum_rows_end_index)
        worksheet.write(row_index, col_index, sum_formula)

def write_subtract_row(worksheet,
                       row_index: int,
                       row1_zero_indexed: int,
                       row2_zero_indexed: int,
                       col_start_zero_indexed: int,
                       num_cols: int):
    # convert zero-indexed cell values to actual cell valeus
    row1_index = row1_zero_indexed + 1
    row2_index = row2_zero_indexed + 1

    for col_index in range(col_start_zero_indexed, col_start_zero_indexed + num_cols):
        col = utility.xl_col_to_name(col_index)
        diff_formula = '={}{}-{}{}'.format(col, row1_index, col, row2_index)
        worksheet.write(row_index, col_index, diff_formula)

def create_line_chart(chart,
                      worksheet_name: str,
                      x_axis_row_zero_indexed: int,
                      line_name_col_zero_indexed: int,
                      data_start_row_zero_indexed: int,
                      num_data_rows: int,
                      data_start_col_zero_indexed: int,
                      num_data_cols: int,
                      line_styles: dict = {}
                     ):
    # convert zero-indexed cell values to actual cell values
    x_axis_row_index = x_axis_row_zero_indexed + 1
    data_row_index = data_start_row_zero_indexed + 1
    line_name_col = utility.xl_col_to_name(line_name_col_zero_indexed)
    data_start_col = utility.xl_col_to_name(data_start_col_zero_indexed)
    data_end_col = utility.xl_col_to_name(data_start_col_zero_indexed + num_data_cols - 1)

    row_reference_str = '={}!${}${}:${}${}'
    x_axis_row = row_reference_str.format(worksheet_name, data_start_col, x_axis_row_index, data_end_col,
            x_axis_row_index)
    for _ in range(num_data_rows):
        values_row = row_reference_str.format(worksheet_name, data_start_col, data_row_index, data_end_col,
                data_row_index)
        line_name = '={}!${}${}'.format(worksheet_name, line_name_col, data_row_index)
        chart.add_series({
            'categories':   x_axis_row,
            'values':       values_row,
            'name':         line_name,
            'line':         line_styles,
            'marker':       { 'type': 'square' }
        })
        data_row_index += 1
