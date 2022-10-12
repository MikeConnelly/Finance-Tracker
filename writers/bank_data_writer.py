import xlsxwriter
from writers import writer_utils
from xlsxwriter import utility
from typing import Dict

BANK_DATA_WORKSHEET_NAME = 'BankData'

def create_bank_data_worksheet(workbook: xlsxwriter.Workbook, bank_data: Dict[str, Dict[str, Dict[str, float]]]):
    """Create a new worksheet and populate it with bank data and create charts."""
    worksheet = workbook.add_worksheet(BANK_DATA_WORKSHEET_NAME)
    # setup sheet indices
    month_row_index = 0
    month_col_start_index = 1
    category_col_index = 0
    category_row_start_index = 1
    data_start_row_index = 1
    data_start_col_index = 1

    months = list(bank_data.keys())
    first_month_entry = months[0]
    num_months = len(months)
    num_income_rows = len(bank_data[first_month_entry]['income'])
    num_expenses_rows = len(bank_data[first_month_entry]['expenses'])
    expenses_start_row_index = data_start_row_index + num_income_rows
    
    # write month row
    writer_utils.write_row(worksheet, month_row_index, month_col_start_index, months)
    # get list of all minor categories
    minor_categories = []
    for _, minor_category_map in bank_data[first_month_entry].items():
        minor_categories.extend(minor_category_map.keys())
    # write category name column
    writer_utils.write_col(worksheet, category_row_start_index, category_col_index, minor_categories)
    
    # create an array for months of arrays for values
    data = []
    for _, major_category_map in bank_data.items():
        data.append([value for major in major_category_map for _, value in major_category_map[major].items()])
    # write data cells
    writer_utils.write_data_cells(worksheet, data_start_row_index, data_start_col_index, data)
    
    # create data chart
    bank_data_chart = workbook.add_chart({ 'type': 'line' })
    bank_data_chart.set_size({ 'width': 800, 'height': 450 })
    # write income lines to chart
    writer_utils.create_line_chart(bank_data_chart, BANK_DATA_WORKSHEET_NAME, month_row_index, category_col_index,
            data_start_row_index, num_income_rows, data_start_col_index, num_months, { 'color': 'green' })
    # write expenses lines to chart
    writer_utils.create_line_chart(bank_data_chart, BANK_DATA_WORKSHEET_NAME, month_row_index, category_col_index,
            expenses_start_row_index, num_expenses_rows, data_start_col_index, num_months, { 'color': 'red' })
    chart_cell = utility.xl_rowcol_to_cell(0, num_months + 4)
    worksheet.insert_chart(chart_cell, bank_data_chart)

    # create total income row
    income_total_row_index = len(minor_categories) + 1
    worksheet.write(income_total_row_index, category_col_index, 'Total Income')
    writer_utils.write_sum_row(worksheet, income_total_row_index, data_start_row_index, num_income_rows,
            data_start_col_index, num_months)
    # create total expenses row
    expenses_total_row_index = income_total_row_index + 1
    worksheet.write(expenses_total_row_index, category_col_index, 'Total Expenses')
    writer_utils.write_sum_row(worksheet, expenses_total_row_index, expenses_start_row_index, num_expenses_rows,
            data_start_col_index, num_months)
    # create surplus row
    surplus_row_index = expenses_total_row_index + 1
    worksheet.write(surplus_row_index, category_col_index, 'Total Surplus')
    writer_utils.write_subtract_row(worksheet, surplus_row_index, income_total_row_index, expenses_total_row_index,
            data_start_col_index, num_months)
    
    # create totals chart
    totals_chart = workbook.add_chart({ 'type': 'line' })
    totals_chart.set_size({ 'width': 800, 'height': 450 })
    writer_utils.create_line_chart(totals_chart, BANK_DATA_WORKSHEET_NAME, month_row_index, category_col_index,
            income_total_row_index, 3, data_start_col_index, num_months)
    chart_cell = utility.xl_rowcol_to_cell(surplus_row_index + 2, 0)
    worksheet.insert_chart(chart_cell, totals_chart)
