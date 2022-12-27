from typing import Dict

import xlsxwriter
from xlsxwriter import utility

from finance_data import FinanceData
from writers import writer_utils
from writers.category_totals_table import CategoryTotalsTable


def create_daily_expenses_worksheets(workbook: xlsxwriter.Workbook,
                                     finance_data: FinanceData,
                                     styles_map: Dict[str, Dict[str, Dict[str, str]]]):
    """Create a new worksheet for every month and populate it with expeneses data."""
    for year, month in finance_data.get_months():
        daily_expenses = finance_data.get_daily_expenses(year, month)
        create_daily_expenses_worksheet(
            workbook, f'{year}-{month}_EXPENSES', daily_expenses, styles_map)


def create_daily_expenses_worksheet(workbook: xlsxwriter.Workbook,
                                    worksheet_name: str,
                                    daily_expenses: Dict[str, Dict[str, float]],
                                    styles_map: Dict[str, Dict[str, Dict[str, str]]]):
    """Create a worksheet and populate it with expenses by day from the given month."""
    worksheet = workbook.add_worksheet(worksheet_name)
    # setup sheet indices
    category_row_index = 0
    category_col_start_index = 1
    day_col_index = 0
    day_row_start_index = 1
    data_start_row_index = 1
    data_start_col_index = 1

    days = list(daily_expenses.keys())
    first_day_map = days[0]
    num_days = len(days)
    num_categories = len(daily_expenses[first_day_map])

    table = CategoryTotalsTable(category_row_index, day_col_index, daily_expenses, styles_map)
    writer_utils.write_table(workbook, worksheet, table)
    writer_utils.create_line_chart(
        workbook, worksheet, worksheet_name, category_row_index, day_col_index, data_start_row_index, num_days,
        data_start_col_index, num_categories, category_row_index, num_categories + 2)
