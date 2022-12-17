import csv
import os
from datetime import datetime
from typing import List, Tuple

from finance_data import FinanceData


def parse_credit_card_data(finance_data: FinanceData,
                           substring_map: List[Tuple[str, str, str]],
                           credit_card_activity_dir: str):
    """
    Parse all transactions from files in `credit_card_activity_dir`.
    Add values to `finance_data` to sum all transactions of the same date and categories.
    Attempt to match each transaction description to a substring in `substring_map`.
    If a match is found, the transaction value will be added to `finance_data` under it's date, major category, and
    minor category. Otherwise, it will be added to the data under it's date, expenses category, and unknown category.
    """
    for credit_card_file in os.listdir(credit_card_activity_dir):
        file_path = f'{credit_card_activity_dir}/{credit_card_file}'
        with open(file_path, 'r') as f:
            reader = csv.reader(f)
            for index, row in enumerate(reader):
                parse_row(finance_data, substring_map, index, row)


def parse_row(finance_data: FinanceData,
              substring_map: List[Tuple[str, str, str]],
              index: int,
              row: List[str]):
    """Parse a row of a credit card file."""
    if index == 0:
        return
    # avoid errors with improperly formatted files
    if len(row) > 3:
        row = row[0:3]
    if len(row) < 3:
        print(f'file_path: line {index} invalid')
        return
    [date_str, desc, amount_str] = row
    # remove header and extra lines
    if len(amount_str) == 0:
        return
    # format row values
    date = datetime.strptime(date_str.strip(), '%m/%d/%Y')
    desc = desc.lower()
    # format amount, skip improperly formatted lines
    # this actually skips lines with negative values like credit card payments
    # should probably check for that explicitly
    if amount_str[0] != '$':
        print(f'file_path: line {index} has invalid amount value')
        return
    amount = float(amount_str[1::])

    # put amount into it's category
    substring_found = False
    for major, minor, substring in substring_map:
        if substring in desc:
            finance_data.add_value(date, major, minor, amount)
            substring_found = True
            break
    # if description did not match any category, put amount into other expenses
    if not substring_found:
        finance_data.add_value(date, 'expenses', 'unknown', amount)
        print('unknown category for payment: ' + desc)
