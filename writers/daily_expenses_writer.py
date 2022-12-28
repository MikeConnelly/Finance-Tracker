from typing import Dict

import xlsxwriter

from finance_data import FinanceData
from writers import writer_utils
from writers.tables import ExpensesTable
from writers.styles import Styles


def create_daily_expenses_worksheets(workbook: xlsxwriter.Workbook, finance_data: FinanceData, styles_map: Styles):
    """Create a new worksheet for every month and populate it with expeneses data."""
    for year, month in finance_data.get_months():
        daily_expenses = finance_data.get_daily_expenses(year, month)
        create_daily_expenses_worksheet(
            workbook, f'{year}-{month}_EXPENSES', daily_expenses, styles_map)


def create_daily_expenses_worksheet(workbook: xlsxwriter.Workbook,
                                    worksheet_name: str,
                                    daily_expenses: Dict[str, Dict[str, float]],
                                    styles_map: Styles):
    """Create a worksheet and populate it with expenses by day from the given month."""
    worksheet = workbook.add_worksheet(worksheet_name)

    table_row = 0
    table_col = 0
    table = ExpensesTable(table_row, table_col, daily_expenses, styles_map)
    writer_utils.write_table(workbook, worksheet, table)

    chart_row = table_row + table.get_height()
    chart_col = table_col
    writer_utils.create_line_chart_for_table(
        workbook, worksheet, worksheet_name, table, table.get_series_for_expenses_chart(),
        chart_row, chart_col)
