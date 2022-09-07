import xlsxwriter
from xlsxwriter import utility
from typing import Dict

BANK_DATA_WORKSHEET_NAME = 'BankData'
CREDIT_CARD_DATA_WORKSHEET_NAME = 'CreditCardData'

def create_xlsx_file(bank_data: Dict[str, Dict[str, Dict[str, float]]], credit_card_data: Dict[str, Dict[str, float]]):
    """Create xlsx file from bank and credit card data."""
    workbook = xlsxwriter.Workbook('output.xlsx')
    create_bank_data_worksheet(workbook, bank_data)
    create_credit_card_data_worksheet(workbook, credit_card_data)
    workbook.close()

###################
# PRIVATE METHODS #
###################

def create_bank_data_worksheet(workbook: xlsxwriter.Workbook, bank_data: Dict[str, Dict[str, Dict[str, float]]]):
    """Create a new worksheet and populate it with bank data and create charts."""
    worksheet = workbook.add_worksheet(BANK_DATA_WORKSHEET_NAME)

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
    
    # create data chart
    bank_data_chart = workbook.add_chart({ 'type': 'line' })
    bank_data_chart.set_size({ 'width': 800, 'height': 450 })
    # set dataset limits
    data_start_col = utility.xl_col_to_name(1)
    data_end_col = utility.xl_col_to_name(num_months)
    data_row_index = 2
    month_row_reference = f'={BANK_DATA_WORKSHEET_NAME}!${data_start_col}$1:${data_end_col}$1'
    for _ in range(num_income_rows):
        values = f'={BANK_DATA_WORKSHEET_NAME}!${data_start_col}${data_row_index}:${data_end_col}${data_row_index}'
        name = f'={BANK_DATA_WORKSHEET_NAME}!$A${data_row_index}'
        bank_data_chart.add_series({
            'categories':   month_row_reference,
            'values':       values,
            'name':         name,
            'line':         { 'color': 'green' },
            'marker':       { 'type': 'square' }
        })
        data_row_index += 1
    for _ in range(num_expenses_rows):
        values = f'={BANK_DATA_WORKSHEET_NAME}!${data_start_col}${data_row_index}:${data_end_col}${data_row_index}'
        name = f'={BANK_DATA_WORKSHEET_NAME}!$A${data_row_index}'
        bank_data_chart.add_series({
            'categories':   month_row_reference,
            'values':       values,
            'name':         name,
            'line':         { 'color': 'red' },
            'marker':       { 'type': 'square' }
        })
        data_row_index += 1
    # rest of rows aren't included in the chart
    chart_cell = utility.xl_rowcol_to_cell(0, num_months + 4)
    worksheet.insert_chart(chart_cell, bank_data_chart)

    # create total income row
    col = 0
    totals_row_index = len(minor_categories) + 1
    worksheet.write(totals_row_index, col, 'Total Income')
    for month_index in range(num_months):
        col_name = utility.xl_col_to_name(month_index + 1)
        total_formula = f'=SUM({col_name}2:{col_name}{num_income_rows + 1})'
        worksheet.write(totals_row_index, month_index + 1, total_formula)
    # create total expenses row
    col = 0
    totals_row_index += 1
    worksheet.write(totals_row_index, col, 'Total Expenses')
    for month_index in range(num_months):
        col_name = utility.xl_col_to_name(month_index + 1)
        total_formula = f'=SUM({col_name}{num_income_rows + 2}:{col_name}{num_income_rows + num_expenses_rows + 1})'
        worksheet.write(totals_row_index, month_index + 1, total_formula)
    # create surplus row
    col = 0
    totals_row_index += 1
    worksheet.write(totals_row_index, col, 'Total Expenses')
    for month_index in range(num_months):
        col_name = utility.xl_col_to_name(month_index + 1)
        total_formula = f'={col_name}{totals_row_index - 1}-{col_name}{totals_row_index}'
        worksheet.write(totals_row_index, month_index + 1, total_formula)
    
    # create totals chart
    totals_chart = workbook.add_chart({ 'type': 'line' })
    totals_chart.set_size({ 'width': 800, 'height': 450 })
    data_row_index = totals_row_index - 1
    for _ in range(3):
        values = f'={BANK_DATA_WORKSHEET_NAME}!${data_start_col}${data_row_index}:${data_end_col}${data_row_index}'
        name = f'={BANK_DATA_WORKSHEET_NAME}!$A${data_row_index}'
        totals_chart.add_series({
            'categories':   month_row_reference,
            'values':       values,
            'name':         name,
            'marker':       { 'type': 'square' }
        })
        data_row_index += 1
    chart_cell = utility.xl_rowcol_to_cell(totals_row_index + 2, 0)
    worksheet.insert_chart(chart_cell, totals_chart)

def create_credit_card_data_worksheet(workbook: xlsxwriter.Workbook, credit_card_data: Dict[str, Dict[str, float]]):
    """Create a new worksheet and populate it with credit card data and create charts."""
    worksheet = workbook.add_worksheet(CREDIT_CARD_DATA_WORKSHEET_NAME)

    months = list(credit_card_data.keys())
    first_month_entry = months[0]
    num_months = len(months)
    num_categories = len(credit_card_data[first_month_entry])
    
    # first row
    row = 0
    col = 1
    for month in months:
        worksheet.write(row, col, month)
        col += 1
    
    # first column
    row = 1
    col = 0
    for category in list(credit_card_data[first_month_entry].keys()):
        worksheet.write(row, col, category)
        row += 1
    
    # create an array for months of arrays for values
    data = []
    for _, category_map in credit_card_data.items():
        data.append([value for _, value in category_map.items()])
    
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
    credit_card_data_chart = workbook.add_chart({ 'type': 'line' })
    credit_card_data_chart.set_size({ 'width': 900, 'height': 500 })
    # set dataset limits
    data_start_col = utility.xl_col_to_name(1)
    data_end_col = utility.xl_col_to_name(num_months)
    data_row_index = 2
    month_row_reference = f'={CREDIT_CARD_DATA_WORKSHEET_NAME}!${data_start_col}$1:${data_end_col}$1'
    for _ in range(num_categories):
        values = f'={CREDIT_CARD_DATA_WORKSHEET_NAME}!${data_start_col}${data_row_index}:${data_end_col}${data_row_index}'
        name = f'={CREDIT_CARD_DATA_WORKSHEET_NAME}!$A${data_row_index}'
        credit_card_data_chart.add_series({
            'categories':   month_row_reference,
            'values':       values,
            'name':         name,
            'marker':       { 'type': 'square' }
        })
        data_row_index += 1
    chart_cell = utility.xl_rowcol_to_cell(0, num_months + 2)
    worksheet.insert_chart(chart_cell, credit_card_data_chart)
