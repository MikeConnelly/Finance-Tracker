from typing import Dict

import xlsxwriter

from finance_data import FinanceData
from writers import overall_data_writer, monthly_expenses_writer, daily_expenses_writer, sankey
from writers.styles import Styles, create_styles_map_for_overall_data, merge_styles_with_defaults

FILE_NAME = 'output.xlsx'


def create_xlsx_file(finance_data: FinanceData, custom_styles: Styles, description_map: Dict[str, str]):
    """Create xlsx file from parsed finance data."""
    overall_styles = create_styles_map_for_overall_data(finance_data.get_categories())
    expenses_styles = merge_styles_with_defaults(finance_data.get_minor_categories('expenses'), custom_styles)

    workbook = xlsxwriter.Workbook(FILE_NAME)
    overall_data_writer.create_overall_worksheets(workbook, finance_data, overall_styles)
    monthly_expenses_writer.create_monthly_expenses_worksheets(workbook, finance_data, expenses_styles)
    daily_expenses_writer.create_daily_expenses_worksheets(workbook, finance_data, expenses_styles)
    workbook.close()
