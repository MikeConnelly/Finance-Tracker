import csv
import os
from datetime import datetime
from typing import List, Tuple

from finance_data import FinanceData


def parse_bank_data(finance_data: FinanceData,
                    substring_map: List[Tuple[str, str, str]],
                    bank_activity_dir: str):
    """
    Parse all transactions from files in `bank_activity_dir`.
    Add values to `finance_data` to sum all transactions of the same date and categories.
    """
    for bank_file in os.listdir(bank_activity_dir):
        file_path = f'{bank_activity_dir}/{bank_file}'
        parse_file(finance_data, substring_map, file_path)


def parse_file(finance_data: FinanceData,
               substring_map: List[Tuple[str, str, str]],
               file_path: str):
    """Parse a single credit card file."""
    with open(file_path, 'r') as f:
        reader = csv.reader(f)
        for index, row in enumerate(reader):
            parse_row(finance_data, substring_map, file_path, index, row)


def parse_row(finance_data: FinanceData,
              substring_map: List[Tuple[str, str, str]],
              file_path: str,
              index: int,
              row: List[str]):
    """Parse a row of a bank file."""
    # skip header line
    if index == 0:
        return
    date_str = None
    desc = None
    value_str = None
    transaction_type = None
    category_overwrite = None
    row_len = len(row)
    # parse the row differently depending on if it contains a specified category
    if row_len == 6:
        [date_str, value_str, desc, _, _, transaction_type] = row
    elif row_len == 7:
        [date_str, value_str, desc, _, _, transaction_type, category_overwrite] = row
    else:
        # invalid row format
        print(row_len)
        print(f'{file_path}: line {index + 1} invalid')
        return
    # format row values
    date = datetime.strptime(date_str, '%Y/%m/%d')
    value = float(value_str)
    
    add_value_to_finance_data(
        finance_data, substring_map, date, desc, value, transaction_type, category_overwrite, file_path, index)


def add_value_to_finance_data(finance_data: FinanceData,
                              substring_map: List[Tuple[str, str, str]],
                              date: str,
                              desc: str,
                              value: float,
                              transaction_type: str,
                              category_overwrite: str,
                              file_path: str,
                              index: int):
    """
    Add `value` to `finance_data`.
    Use `category_overwrite` or search for a category that matches `desc` in `substring_map`.
    """
    # put value into it's category
    if category_overwrite:
        major_category = finance_data.get_major_category(category_overwrite)
        if major_category:
            finance_data.add_value(date, major_category, category_overwrite, value)
            return
        else:
            print(f'{file_path}: line {index + 1} has an invalid category overwrite value')
    # check if the description contains any substrings
    desc_lowercase = desc.lower()
    for major, minor, substring in substring_map:
        if substring.lower() in desc_lowercase:
            finance_data.add_value(date, major, minor, value)
            return
    # if description did not match any category, put value into other
    # assume credit = income, debit = expense
    print(f'unknown category for {desc}')
    if transaction_type == 'CREDIT':
        finance_data.add_value(date, 'unknown', 'credit', value)
    else:
        finance_data.add_value(date, 'unknown', 'debit', value)
