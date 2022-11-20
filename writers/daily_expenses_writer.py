from typing import Dict

import xlsxwriter
from xlsxwriter import utility

from finance_data import FinanceData
from writers import writer_utils


def create_daily_expenses_worksheets(workbook: xlsxwriter.Workbook, finance_data: FinanceData):
    """Create a new worksheet for every month and populate it with expeneses data."""
    for year, month in finance_data.get_months():
        daily_expenses = finance_data.get_daily_expenses(year, month)
        create_daily_expenses_worksheet(
            workbook, f'{year}-{month}', daily_expenses)


def create_daily_expenses_worksheet(workbook: xlsxwriter.Workbook,
                                    month: str,
                                    daily_expenses: Dict[str, Dict[str, float]]):
    """Create a worksheet and populate it with expenses by day from the given month."""
    worksheet = workbook.add_worksheet(month)
    # setup sheet indices
    category_row_index = 0
    category_col_start_index = 1
    day_col_index = 0
    day_row_start_index = 1
    data_start_row_index = 1
    data_start_col_index = 1

    days = list(daily_expenses.keys())
    first_month_entry = days[0]
    num_days = len(days)
    num_categories = len(daily_expenses[first_month_entry])

    # write month row
    writer_utils.write_row(worksheet, category_row_index, category_col_start_index,
                           list(daily_expenses[first_month_entry].keys()))
    # write category name column
    writer_utils.write_col(worksheet, day_row_start_index, day_col_index, days)

    # convert monthly data dicts into a list
    data = []
    for _, category_map in daily_expenses.items():
        data.append([value for _, value in category_map.items()])
    # write data cells
    writer_utils.write_data_cells_by_row(
        worksheet, data_start_row_index, data_start_col_index, data)

    # create total income row
    totals_row_index = num_days + 1
    worksheet.write(totals_row_index, day_col_index, 'Total')
    writer_utils.write_sum_row(worksheet, totals_row_index, data_start_row_index, num_days, data_start_col_index,
                               num_categories)

    # create line chart
    credit_card_data_chart = workbook.add_chart({'type': 'line'})
    credit_card_data_chart.set_size({'width': 900, 'height': 500})
    writer_utils.create_line_chart_with_series_as_cols(credit_card_data_chart, month, category_row_index,
                                                       day_col_index, data_start_row_index, num_days,
                                                       data_start_col_index, num_categories)
    chart_cell = utility.xl_rowcol_to_cell(0, num_categories + 2)
    worksheet.insert_chart(chart_cell, credit_card_data_chart)
