import xlsxwriter
from typing import Dict
from writers import bank_data_writer, credit_card_data_writer

def create_xlsx_file(bank_data: Dict[str, Dict[str, Dict[str, float]]], credit_card_data: Dict[str, Dict[str, float]]):
    """Create xlsx file from bank and credit card data."""
    workbook = xlsxwriter.Workbook('output.xlsx')
    bank_data_writer.create_bank_data_worksheet(workbook, bank_data)
    credit_card_data_writer.create_credit_card_data_worksheet(workbook, credit_card_data)
    # maybe combine the bank expenses with cc expenses into a total expenses graph
    workbook.close()
