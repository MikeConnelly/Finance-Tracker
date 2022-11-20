import csv
import os
from datetime import datetime
from typing import List, Tuple

from finance_data import FinanceData


def parse_bank_data(finance_data: FinanceData,
                    description_map: List[Tuple[str, str, str]],
                    bank_activity_dir: str):
    """
    Parse all transactions from files in `bank_activity_dir`.
    Add values to `finance_data` to sum all transactions of the same date and categories.
    Attempt to match each transaction description to a substring in `description_map`.
    If a match is found, the transaction value will be added to `finance_data` under it's date, major category,
    and minor category. Otherwise, it will be added to the data under it's date, unknown category, credit or debit.
    """
    for bank_file in os.listdir(bank_activity_dir):
        file_path = f'{bank_activity_dir}/{bank_file}'
        with open(file_path, 'r') as f:
            reader = csv.reader(f)
            for index, row in enumerate(reader):
                parse_row(finance_data, description_map, index, row)


def parse_row(finance_data: FinanceData,
              description_map: List[Tuple[str, str, str]],
              index: int,
              row: List[str]):
    """Parse a row of a bank file."""
    # skip header line
    if index == 0:
        return
    # format row values
    [date_str, amount_str, desc, _, _, transaction_type] = row
    date = datetime.strptime(date_str, '%Y/%m/%d')
    desc = desc.lower()
    amount = float(amount_str)
    
    # put amount into it's category
    substring_found = False
    for major, minor, substring in description_map:
        if substring in desc:
            finance_data.add_value(date, major, minor, amount)
            substring_found = True
            break
    # if description did not match any category, put amount into other
    if not substring_found:
        # assume credit = income, debit = expense
        if transaction_type == 'CREDIT':
            finance_data.add_value(date, 'unknown', 'credit', amount)
        else:
            finance_data.add_value(date, 'unknown', 'debit', amount)
