from typing import Dict

import xlsxwriter

from finance_data import FinanceData
from writers import overall_data_writer, monthly_expenses_writer, daily_expenses_writer

FILE_NAME = 'output.xlsx'


def create_xlsx_file(finance_data: FinanceData, styles_map: Dict[str, str], description_map: Dict[str, str]):
    """Create xlsx file from parsed finance data."""
    workbook = xlsxwriter.Workbook(FILE_NAME)
    overall_data_writer.create_overall_data_worksheet(workbook, finance_data)
    monthly_expenses_writer.create_monthly_expenses_worksheets(workbook, finance_data, styles_map, description_map)
    daily_expenses_writer.create_daily_expenses_worksheets(workbook, finance_data, styles_map)
    workbook.close()
