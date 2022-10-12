import xlsxwriter
from writers import writer_utils
from xlsxwriter import utility
from typing import Dict

CREDIT_CARD_DATA_WORKSHEET_NAME = 'CreditCardData'

def create_credit_card_data_worksheet(workbook: xlsxwriter.Workbook, credit_card_data: Dict[str, Dict[str, float]]):
    """Create a new worksheet and populate it with credit card data and create charts."""
    worksheet = workbook.add_worksheet(CREDIT_CARD_DATA_WORKSHEET_NAME)
    # setup sheet indices
    month_row_index = 0
    month_col_start_index = 1
    category_col_index = 0
    category_row_start_index = 1
    data_start_row_index = 1
    data_start_col_index = 1

    months = list(credit_card_data.keys())
    first_month_entry = months[0]
    num_months = len(months)
    num_categories = len(credit_card_data[first_month_entry])
    
    # write month row
    writer_utils.write_row(worksheet, month_row_index, month_col_start_index, months)
    # write category name column
    writer_utils.write_col(worksheet, category_row_start_index, category_col_index,
            list(credit_card_data[first_month_entry].keys()))
    
    # convert monthly data dicts into a list
    data = []
    for _, category_map in credit_card_data.items():
        data.append([value for _, value in category_map.items()])
    # write data cells
    writer_utils.write_data_cells(worksheet, data_start_row_index, data_start_col_index, data)
    
    # create total income row
    totals_row_index = num_categories + 1
    worksheet.write(totals_row_index, category_col_index, 'Total')
    writer_utils.write_sum_row(worksheet, totals_row_index, data_start_row_index, num_categories, data_start_col_index,
            num_months)
    
    # create line chart
    credit_card_data_chart = workbook.add_chart({ 'type': 'line' })
    credit_card_data_chart.set_size({ 'width': 900, 'height': 500 })
    writer_utils.create_line_chart(credit_card_data_chart, CREDIT_CARD_DATA_WORKSHEET_NAME, month_row_index,
            category_col_index, data_start_row_index, num_categories, data_start_col_index, num_months)
    chart_cell = utility.xl_rowcol_to_cell(0, num_months + 2)
    worksheet.insert_chart(chart_cell, credit_card_data_chart)
