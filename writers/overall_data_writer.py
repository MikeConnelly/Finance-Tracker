from xlsxwriter import Workbook

from finance_data import FinanceData
from writers import writer_utils
from writers.tables import OverallTable
from writers.styles import Styles

OVERALL_DATA_WORKSHEET_NAME = 'OVERALL_DATA'


def create_overall_data_worksheet(workbook: Workbook, finance_data: FinanceData, styles_map: Styles):
    """Create a new worksheet and populate it with the overall data by month."""
    worksheet = workbook.add_worksheet(OVERALL_DATA_WORKSHEET_NAME)
    overall_data = finance_data.get_monthly_overall()

    table_row = 0
    table_col = 0
    table = OverallTable(table_row, table_col, overall_data, styles_map)
    writer_utils.write_table(workbook, worksheet, table)

    income_expenses_chart_row = table_row + table.get_height()
    income_expenses_chart_col = table_col
    writer_utils.create_line_chart_for_table(
        workbook, worksheet, OVERALL_DATA_WORKSHEET_NAME, table, table.get_series_for_income_expenses_chart(),
        income_expenses_chart_row, income_expenses_chart_col)

    totals_chart_row = income_expenses_chart_row
    totals_chart_col = income_expenses_chart_col + 15
    writer_utils.create_line_chart_for_table(
        workbook, worksheet, OVERALL_DATA_WORKSHEET_NAME, table, table.get_series_for_totals_chart(),
        totals_chart_row, totals_chart_col)
