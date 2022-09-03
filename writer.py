import xlsxwriter
from xlsxwriter import utility
from typing import Dict

def create_xlsx_file(bank_data: Dict[str, Dict[str, Dict[str, float]]], credit_card_data: Dict[str, Dict[str, float]]):
    """Create xlsx file from bank and credit card data."""
    workbook = xlsxwriter.Workbook('output.xlsx')
    worksheet = workbook.add_worksheet()

    populate_bank_data(workbook, worksheet, bank_data)

    workbook.close()

    # to format bank data for the excel sheet
    # need to populate the first row with an array of months
    # need to make the first column all the minor categories
    # use the major categories to determine the indices of incomes and expenses for use in the graph

def populate_bank_data(workbook, worksheet, bank_data: Dict[str, Dict[str, Dict[str, float]]]):
    """Populate `worksheet` with bank data and create charts."""
    months = list(bank_data.keys())
    first_month_entry = months[0]
    num_months = len(months)
    num_income_rows = len(bank_data[first_month_entry]['income'])
    num_expenses_rows = len(bank_data[first_month_entry]['expenses'])
    
    # first row
    row = 0
    col = 1
    for month in months:
        worksheet.write(row, col, month)
        col += 1
    
    # get list of all minor categories
    minor_categories = []
    for _, minor_category_map in bank_data[first_month_entry].items():
        minor_categories.extend(minor_category_map.keys())

    # first column
    row = 1
    col = 0
    for category in minor_categories:
        worksheet.write(row, col, category)
        row += 1
    
    # create an array for months of arrays for values
    data = []
    for month, major_category_map in bank_data.items():
        data.append([value for major in major_category_map for _, value in major_category_map[major].items()])
    
    # data cells
    row = 1
    col = 1
    for monthly_values in data:
        for value in monthly_values:
            worksheet.write(row, col, value)
            row += 1
        row = 1
        col += 1
    
    # create line chart
    bank_data_chart = workbook.add_chart({ 'type': 'line' })
    bank_data_chart.set_size({ 'width': 900, 'height': 500 })
    data_start_col = utility.xl_col_to_name(1)
    data_end_col = utility.xl_col_to_name(num_months)
    data_row_index = 2
    month_row_ref = f'=Sheet1!${data_start_col}$1:${data_end_col}$1'
    for _ in range(num_income_rows):
        bank_data_chart.add_series({
            'values':       f'=Sheet1!${data_start_col}${data_row_index}:${data_end_col}${data_row_index}',
            'name':         f'=Sheet1!$A${data_row_index}',
            'line':         { 'color': 'green' },
            'marker':       { 'type': 'square' },
            'categories':   month_row_ref
        })
        data_row_index += 1
    for _ in range(num_expenses_rows):
        bank_data_chart.add_series({
            'values':       f'=Sheet1!${data_start_col}${data_row_index}:${data_end_col}${data_row_index}',
            'name':         f'=Sheet1!$A${data_row_index}',
            'line':         { 'color': 'red' },
            'marker':       { 'type': 'square' },
            'categories':   month_row_ref
        })
        data_row_index += 1
    # rest of rows aren't included in the chart
    chart_cell = utility.xl_rowcol_to_cell(0, num_months + 2)
    worksheet.insert_chart(chart_cell, bank_data_chart)
