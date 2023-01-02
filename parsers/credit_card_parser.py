import csv
import os
import re
from datetime import datetime
from typing import List, Tuple

from finance_data import FinanceData

value_pattern = re.compile(r'^-?\$\d+\.\d\d$')
positive_value_pattern = re.compile(r'^\$(\d+\.\d\d)$')


def parse_credit_card_data(finance_data: FinanceData,
                           substring_map: List[Tuple[str, str, str]],
                           credit_card_activity_dir: str):
    """
    Parse all transactions from files in `credit_card_activity_dir`.
    Add values to `finance_data` to sum all transactions of the same date and categories.
    """
    for credit_card_file in os.listdir(credit_card_activity_dir):
        file_path = f'{credit_card_activity_dir}/{credit_card_file}'
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
              row: list[str]):
    """Parse a single row of a credit card file."""
    # skip header line
    if index == 0:
        return
    date_str = None
    desc_list = None
    desc = None
    value_str = None
    category_overwrite = None
    row_len = len(row)
    # parse the row differently depending on if it contains a specified category
    # and if the description was parsed incorrectly due to containing commas
    if row_len == 4:
        # row has no category_overwrite and a correctly formatted description
        date_str, desc, value_str, _ = row
    elif row_len == 5:
        # row could have two description entries due to containing a comma OR the row has a category_overwrite
        date_str, desc, value_str, _, category_overwrite = row
        value_is_valid = bool(value_pattern.match(value_str))  # check that value is not part of the description
        if not value_is_valid:
            # desc has multiple entires and overwrite does NOT exist
            date_str, *desc_list, value_str, _ = row
    elif row_len > 5:
        # row has multiple description entires due to containing a comma AND the row has a category_overwrite
        date_str, *desc_list, value_str, _, category_overwrite = row
    elif row_len == 1:
        # skip final row of the file
        return
    else:
        # invalid row format
        print(f'{file_path}: line {index + 1} invalid')
        return

    # format row values
    parsed_value = positive_value_pattern.match(value_str)
    if not parsed_value:
        print(f'{file_path}: line {index + 1} contains a negative value.')
        return
    value = float(parsed_value.group(1))
    date = datetime.strptime(date_str.strip(), '%m/%d/%Y')
    if desc_list:
        desc = ','.join(desc_list)

    add_value_to_finance_data(finance_data, substring_map, date, desc, value, category_overwrite, file_path, index)


def add_value_to_finance_data(finance_data: FinanceData,
                              substring_map: List[Tuple[str, str, str]],
                              date: str,
                              desc: str,
                              value: float,
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
    # if description did not match any category, put value into other expenses
    print('unknown category for payment: ' + desc)
    finance_data.add_value(date, 'expenses', 'unknown', value)
