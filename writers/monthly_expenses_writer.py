from typing import Dict

import xlsxwriter
from xlsxwriter import utility

from finance_data import FinanceData
from writers import writer_utils

MONTHLY_EXPENSES_WORKSHEET_NAME = 'Monthly_Expenses_Data'


def create_monthly_expenses_worksheet(workbook: xlsxwriter.Workbook, financeData: FinanceData):
    """Create a new worksheet and populate it with expenses data by month."""
    worksheet = workbook.add_worksheet(MONTHLY_EXPENSES_WORKSHEET_NAME)
    monthly_expenses_data = financeData.get_monthly_expenses_totals()
    # setup sheet indices
    month_row_index = 0
    month_col_start_index = 1
    category_col_index = 0
    category_row_start_index = 1
    data_start_row_index = 1
    data_start_col_index = 1

    months = list(monthly_expenses_data.keys())
    first_month_entry = months[0]
    num_months = len(months)
    num_categories = len(monthly_expenses_data[first_month_entry])

    # write month row
    writer_utils.write_row(worksheet, month_row_index, month_col_start_index, months)
    # write category name column
    writer_utils.write_col(worksheet, category_row_start_index, category_col_index,
                           list(monthly_expenses_data[first_month_entry].keys()))

    # convert monthly data dicts into a list
    data = []
    for _, category_map in monthly_expenses_data.items():
        data.append([value for _, value in category_map.items()])
    # write data cells
    writer_utils.write_data_cells_by_col(
        worksheet, data_start_row_index, data_start_col_index, data)

    # create total income row
    totals_row_index = num_categories + 1
    worksheet.write(totals_row_index, category_col_index, 'Total')
    writer_utils.write_sum_row(worksheet, totals_row_index, data_start_row_index, num_categories, data_start_col_index,
                               num_months)

    # create line chart
    monthly_expenses_data_chart = workbook.add_chart({'type': 'line'})
    monthly_expenses_data_chart.set_size({'width': 900, 'height': 500})
    writer_utils.create_line_chart_with_series_as_rows(monthly_expenses_data_chart, MONTHLY_EXPENSES_WORKSHEET_NAME,
                                                       month_row_index, category_col_index, data_start_row_index,
                                                       num_categories, data_start_col_index, num_months)
    chart_cell = utility.xl_rowcol_to_cell(0, num_months + 2)
    worksheet.insert_chart(chart_cell, monthly_expenses_data_chart)
