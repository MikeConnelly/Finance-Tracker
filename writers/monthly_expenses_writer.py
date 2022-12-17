from typing import Dict

from finance_data import FinanceData
from xlsxwriter import Workbook, utility

from writers import writer_utils
from writers.category_totals_table import CategoryTotalsTable

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
    # get table dimensions
    months = list(monthly_expenses.keys())
    num_months = len(months)
    num_categories = len(monthly_expenses[months[0]])
    # setup indices for the expenses section of the worksheet
    expenses_category_row_index = 0
    expenses_category_col_start_index = 1
    expenses_month_col_index = 0
    expenses_month_row_start_index = 1
    expenses_data_start_row_index = 1
    expenses_data_start_col_index = 1
    # setup indices for the percent change section of the worksheet
    offset = 3
    pc_category_row_index = 0
    pc_category_col_start_index = expenses_category_col_start_index + num_categories + offset
    pc_month_col_index = expenses_month_col_index + num_categories + offset
    pc_month_row_start_index = 1
    pc_data_start_row_index = 1
    pc_data_start_col_index = expenses_data_start_col_index + num_categories + offset

    create_expenses_section(
        workbook, worksheet, worksheet_name, monthly_expenses, months, expenses_category_row_index,
        expenses_category_col_start_index, num_categories, expenses_month_col_index,
        expenses_month_row_start_index, num_months, expenses_data_start_row_index,
        expenses_data_start_col_index, styles_map)

    # create_percent_change_section(
    #     workbook, worksheet, worksheet_name, monthly_percent_change, months, pc_category_row_index,
    #     pc_category_col_start_index, num_categories, pc_month_col_index, pc_month_row_start_index, num_months,
    #     pc_data_start_row_index, pc_data_start_col_index)


def create_expenses_section(workbook: Workbook,
                            worksheet: Worksheet,
                            worksheet_name: str,
                            monthly_expenses: Dict[str, Dict[str, float]],
                            months: list[str],
                            category_row_index: int,
                            category_col_start_index: int,
                            num_categories: int,
                            month_col_index: int,
                            month_row_start_index: int,
                            num_months: int,
                            data_start_row_index: int,
                            data_start_col_index: int,
                            styles_map: Dict[str, Dict[str, str]]):
    """Create a worksheet and populate it with expenses by month from the given year."""
    table = CategoryTotalsTable(monthly_expenses)
    writer_utils.write_table(workbook, worksheet, category_row_index, month_col_index, table, styles_map)
    writer_utils.create_line_chart(
        workbook, worksheet, worksheet_name, category_row_index, month_col_index, data_start_row_index, num_months,
        data_start_col_index, num_categories, num_months + 2, month_col_index)


def create_percent_change_section(workbook: Workbook,
                                  worksheet: Worksheet,
                                  worksheet_name: str,
                                  monthly_expenses: Dict[str, Dict[str, float]],
                                  months: list[str],
                                  category_row_index: int,
                                  category_col_start_index: int,
                                  num_categories: int,
                                  month_col_index: int,
                                  month_row_start_index: int,
                                  num_months: int,
                                  data_start_row_index: int,
                                  data_start_col_index: int):
    """Create a worksheet and populate it with expenses by month from the given year."""
    # write category row
    writer_utils.write_row(worksheet, category_row_index, category_col_start_index,
                           list(monthly_expenses[months[0]].keys()))
    # write month column
    writer_utils.write_col(worksheet, month_row_start_index, month_col_index, months)

    # convert monthly data dicts into a list
    data = []
    for _, category_map in monthly_expenses.items():
        data.append([value for _, value in category_map.items()])
    # write data cells
    writer_utils.write_data_cells_by_row(
        worksheet, data_start_row_index, data_start_col_index, data)

    # create line chart
    credit_card_data_chart = workbook.add_chart({'type': 'line'})
    credit_card_data_chart.set_size({'width': 900, 'height': 500})
    writer_utils.create_line_chart_with_series_as_cols(
        credit_card_data_chart, worksheet_name, category_row_index, month_col_index, data_start_row_index, num_months,
        data_start_col_index, num_categories)
    chart_cell = utility.xl_rowcol_to_cell(num_months + 2, month_col_index)
    worksheet.insert_chart(chart_cell, credit_card_data_chart)
