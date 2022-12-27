from typing import Dict

from finance_data import FinanceData
from xlsxwriter import Workbook, utility

from writers import writer_utils
from writers.category_totals_table import ExpensesTable

Worksheet = Workbook.worksheet_class


def create_monthly_expenses_worksheets(workbook: Workbook,
                                       finance_data: FinanceData,
                                       styles_map: Dict[str, Dict[str, str]],
                                       description_map: Dict[str, str]):
    """Create a new worksheet for every year and populate it with expeneses data."""
    for year in finance_data.get_years():
        monthly_expenses = finance_data.get_monthly_expenses(year)
        monthly_percent_change = finance_data.get_percent_change_of_monthly_expenses(year)
        create_monthly_expenses_worksheet(
            workbook, f'{year}_EXPENSES', monthly_expenses, monthly_percent_change, styles_map)


def create_monthly_expenses_worksheet(workbook: Workbook,
                                      worksheet_name: str,
                                      monthly_expenses: Dict[str, Dict[str, float]],
                                      monthly_percent_change: Dict[str, Dict[str, float]],
                                      styles_map: Dict[str, Dict[str, str]]):
    """Create a worksheet and populate it with expenses by month from the given year."""
    worksheet = workbook.add_worksheet(worksheet_name)
    table = ExpensesTable(0, 0, monthly_expenses, styles_map)
    writer_utils.write_table(workbook, worksheet, table)
    writer_utils.create_line_chart_for_table(workbook, worksheet, worksheet_name, table)
