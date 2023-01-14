from typing import Dict

from xlsxwriter import Workbook

from finance_data import FinanceData
from writers import writer_utils, sankey
from writers.tables import OverallTable
from writers.styles import Styles

OVERALL_DATA_WORKSHEET_NAME = 'OVERALL_DATA'
DEFAULT_COLUMN_WIDTH = 15


def create_overall_worksheets(workbook: Workbook, finance_data: FinanceData, styles_map: Styles):
    """Create a new worksheet for every year and populate it with the overall data for that year."""
    yearly_totals = finance_data.get_yearly_overall()
    for year in finance_data.get_years():
        monthly_totals = finance_data.get_monthly_overall(year)
        year_totals = yearly_totals[year]
        create_overall_data_worksheet(workbook, f'{year}_SUMMARY', monthly_totals, year_totals, styles_map)


def create_overall_data_worksheet(workbook: Workbook,
                                  worksheet_name: str,
                                  monthly_totals: Dict[str, Dict[str, Dict[str, float]]],
                                  year_totals: Dict[str, Dict[str, float]],
                                  styles_map: Styles):
    """Create a new worksheet and populate it with the overall data by month."""
    worksheet = workbook.add_worksheet(worksheet_name)

    table_row = 0
    table_col = 0
    table = OverallTable(table_row, table_col, monthly_totals, styles_map)
    writer_utils.write_table(workbook, worksheet, table)

    # TODO: setup default and custom worksheet formats
    worksheet.set_column(table_col, table_col + table.get_width() - 1, DEFAULT_COLUMN_WIDTH)

    income_expenses_chart_row = table_row + table.get_height()
    income_expenses_chart_col = table_col
    writer_utils.create_line_chart_for_table(
        workbook, worksheet, worksheet_name, table, table.get_series_for_income_expenses_chart(),
        income_expenses_chart_row, income_expenses_chart_col)

    totals_chart_row = income_expenses_chart_row
    totals_chart_col = income_expenses_chart_col + 10
    writer_utils.create_line_chart_for_table(
        workbook, worksheet, worksheet_name, table, table.get_series_for_totals_chart(),
        totals_chart_row, totals_chart_col)

    img_path = sankey.create_sankey_plot_for_overall_data(year_totals)
    worksheet.insert_image('A45', img_path)
