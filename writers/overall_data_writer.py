import xlsxwriter
from xlsxwriter import utility

from finance_data import FinanceData
from writers import writer_utils, styles
from writers.tables import OverallTable

OVERALL_DATA_WORKSHEET_NAME = 'OVERALL_DATA'


def create_overall_data_worksheet(workbook: xlsxwriter.Workbook, financeData: FinanceData):
    """Create a new worksheet and populate it with the overall data by month."""
    worksheet = workbook.add_worksheet(OVERALL_DATA_WORKSHEET_NAME)
    overall_data = financeData.get_monthly_overall()
    styles_map = styles.create_styles_map_for_overall_data(overall_data)

    table_row = 0
    table_col = 0
    table = OverallTable(table_row, table_col, overall_data, styles_map)
    writer_utils.write_table(workbook, worksheet, table)

    chart_row = table_row + table.get_height()
    chart_col = table_col
    writer_utils.create_line_chart_for_table(
        workbook, worksheet, OVERALL_DATA_WORKSHEET_NAME, table, chart_row, chart_col)
