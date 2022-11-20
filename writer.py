from typing import Dict

import xlsxwriter

from finance_data import FinanceData
from writers import monthly_expenses_writer, overall_data_writer, daily_expenses_writer, writer_utils

FILE_NAME = 'output.xlsx'


def create_xlsx_file(finance_data: FinanceData):
    """Create xlsx file from parsed finance data."""
    workbook = xlsxwriter.Workbook(FILE_NAME)
    overall_data_writer.create_overall_data_worksheet(workbook, finance_data)
    monthly_expenses_writer.create_monthly_expenses_worksheet(workbook, finance_data)
    daily_expenses_writer.create_daily_expenses_worksheets(workbook, finance_data)
    workbook.close()
