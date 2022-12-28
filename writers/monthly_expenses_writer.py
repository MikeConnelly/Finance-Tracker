from typing import Dict

from finance_data import FinanceData
from xlsxwriter import Workbook

from writers import writer_utils
from writers.tables import ExpensesTable
from writers.styles import Styles

Worksheet = Workbook.worksheet_class


def create_monthly_expenses_worksheets(workbook: Workbook, finance_data: FinanceData, styles_map: Styles):
    """Create a new worksheet for every year and populate it with expeneses data."""
    for year in finance_data.get_years():
        monthly_expenses = finance_data.get_monthly_expenses(year)
        create_monthly_expenses_worksheet(
            workbook, f'{year}_EXPENSES', monthly_expenses, styles_map)


def create_monthly_expenses_worksheet(workbook: Workbook,
                                      worksheet_name: str,
                                      monthly_expenses: Dict[str, Dict[str, float]],
                                      styles_map: Styles):
    """Create a worksheet and populate it with expenses by month from the given year."""
    worksheet = workbook.add_worksheet(worksheet_name)

    table_row = 0
    table_col = 0
    table = ExpensesTable(table_row, table_col, monthly_expenses, styles_map)
    writer_utils.write_table(workbook, worksheet, table)

    chart_row = table_row + table.get_height()
    chart_col = table_col
    writer_utils.create_line_chart_for_table(
        workbook, worksheet, worksheet_name, table, table.get_series_for_expenses_chart(),
        chart_row, chart_col)
