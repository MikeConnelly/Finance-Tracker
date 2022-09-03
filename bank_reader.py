import os
import csv
import copy
from typing import Dict, List, Tuple

def parse_bank_data(bank_config: Dict[str, Dict[str, List[str]]],
                    bank_activity_dir: str
                   ) -> Dict[str, Dict[str, Dict[str, float]]]:
    """
    Parse all transactions from files in `bank_activity_dir`.
    Create a `dict` of bank data that contains sums of all transactions of the same month and category.
    Major and minor transaction categories are determined by `bank_config`.

    `returns` a `dict` of all transaction data mapping month to major categories, to minor categories, to the sum
    of all transactions within this category.
    """
    bank_default_values = create_default_value_map(bank_config)
    description_map = create_description_map(bank_config)

    bank_data = parse_bank_files(bank_activity_dir, bank_default_values, description_map)

    # round all values to 2 decimal points
    for month, major_category_map in bank_data.items():
        for major, minor_category_map in major_category_map.items():
            for minor, value in minor_category_map.items():
                bank_data[month][major][minor] = round(value, 2)

    return bank_data

###################
# PRIVATE METHODS #
###################

def create_default_value_map(bank_config: Dict[str, Dict[str, List[str]]]) -> Dict[str, Dict[str, float]]:
    """Create default value `dict` by replacing the substring lists in `bank_config` with zeros."""
    bank_default_values = copy.deepcopy(bank_config)
    for major_category in bank_default_values.keys():
        for minor_category in bank_default_values[major_category].keys():
            bank_default_values[major_category][minor_category] = 0
    return bank_default_values

def create_description_map(bank_config: Dict[str, Dict[str, List[str]]]) -> List[Tuple[str, str, str]]:
    """Create a `list` of all substrings in `bank_config` tupled with their categories."""
    return [(major, minor, substring)
            for major in bank_config
            for minor in bank_config[major]
            for substring in bank_config[major][minor]]

def parse_bank_files(bank_activity_dir: str,
                     bank_default_values: Dict[str, Dict[str, float]],
                     description_map: List[Tuple[str, str, str]]
                    ) -> Dict[str, Dict[str, Dict[str, float]]]:
    """
    Read all lines from files in `bank_activity_dir`.
    Attempt to match each transaction description to a substring in `description_map`.
    If a match is found, the transaction value will be added to the data `dict` under it's month, major category,
    and minor category. Otherwise, it will be added to the data under it's month, unknown category, credit or debit.

    `bank_default_values` is used to populate default values each time a new month is added to the data.

    `description_map` should contain tuples of major category, minor category, and substring.
    """
    bank_data = {}
    for bank_file in os.listdir(bank_activity_dir):
        file_path = f'{bank_activity_dir}/{bank_file}'
        with open(file_path, 'r') as f:
            reader = csv.reader(f)
            line = 0
            for row in reader:
                line += 1

                # skip header line
                if line == 1:
                    continue
                [date, amount, desc, _, _, transaction_type] = row

                # format date from YYYY/MM/DD to YYYY/MM
                month = date[:7]
                # format description
                desc = desc.lower()
                # format amount
                amount = float(amount)

                # check if month already exists in bank_data. If not, then create a new entry for it
                if month not in bank_data.keys():
                    bank_data[month] = copy.deepcopy(bank_default_values)

                # put amount into it's category
                substring_found = False
                for major, minor, substring in description_map:
                    if substring in desc:
                        bank_data[month][major][minor] += amount
                        substring_found = True
                        break
                # if description did not match any category, put amount into other
                if not substring_found:
                    # assume credit = income, debit = expense
                    if transaction_type == 'CREDIT':
                        bank_data[month]['unknown']['credit'] += amount
                    else:
                        bank_data[month]['unknown']['debit'] += amount
    return bank_data
