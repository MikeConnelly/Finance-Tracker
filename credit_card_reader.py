import os
import csv
import copy
from typing import Dict, List, Tuple

def parse_credit_card_data(credit_card_config: Dict[str, List[str]],
                           credit_card_activity_dir: str
                          ) -> Dict[str, Dict[str, float]]:
    """
    Parse all transactions from files in `credit_card_activity_dir`.
    Create a `dict` of credit card data that contains sums of all transactions of the same month and category.
    categories are determined by `credit_card_config`.

    `returns` a `dict` of all transaction data mapping month to categories, to the sum of all
    transactions within this category.
    """
    credit_card_default_values = create_default_value_map(credit_card_config)
    description_map = create_description_map(credit_card_config)

    credit_card_data = parse_credit_card_files(credit_card_activity_dir, credit_card_default_values, description_map)
    
    # round all values to 2 decimal points
    for month, category_map in credit_card_data.items():
        for category, value in category_map.items():
            credit_card_data[month][category] = round(value, 2)
    
    return credit_card_data

###################
# PRIVATE METHODS #
###################

def create_default_value_map(credit_card_config: Dict[str, List[str]]) -> Dict[str, float]:
    """Create default value `dict` by replacing the substring lists in `credit_card_config` with zeros."""
    credit_card_default_values = copy.deepcopy(credit_card_config)
    for category in credit_card_default_values.keys():
        credit_card_default_values[category] = 0
    return credit_card_default_values

def create_description_map(credit_card_config: Dict[str, List[str]]) -> List[Tuple[str, str]]:
    """Create a `list` of all substrings in `credit_card_config` tupled with their categories."""
    return [(category, substring) for category in credit_card_config for substring in credit_card_config[category]]

def parse_credit_card_files(credit_card_activity_dir: str,
                            credit_card_default_values: Dict[str, float],
                            description_map: List[Tuple[str, str]]
                           ) -> Dict[str, Dict[str, float]]:
    """
    Read all lines from files in `credit_card_activity_dir`.
    Attempt to match each transaction description to a substring in `description_map`.
    If a match is found, the transaction value will be added to the data `dict` under it's month and category.
    Otherwise, it will be added to the data under it's month, unknown category.

    `credit_card_default_values` is used to populate default values each time a new month is added to the data.

    `description_map` should contain tuples of category and substring.
    """
    credit_card_data = {}
    for credit_card_file in os.listdir(credit_card_activity_dir):
        file_path = f'{credit_card_activity_dir}/{credit_card_file}'
        with open(file_path, 'r') as f:
            reader = csv.reader(f)
            line = 0
            for row in reader:
                line += 1

                # avoid errors with improperly formatted files
                if len(row) > 3:
                    row = row[0:3]
                if len(row) < 3:
                    continue
                [date, desc, amount] = row
                # remove header and extra lines
                if line == 1 or len(amount) == 0:
                    continue
                
                # format date into MM/YYYY
                date = date.strip()
                month = date[0:2] + date[5:]
                # format description
                desc = desc.lower()
                # format amount, skip improperly formatted lines
                # this actually skips lines with negative values like credit card payments
                # should probably check for that explicitly
                if amount[0] != '$':
                    continue
                amount = float(amount[1::])

                # check if month already exists in cc_data. If not, then create a new entry for it
                if month not in credit_card_data.keys():
                    credit_card_data[month] = copy.deepcopy(credit_card_default_values)

                # put amount into it's category
                substring_found = False
                for category, substring in description_map:
                    if substring in desc:
                        credit_card_data[month][category] += amount
                        substring_found = True
                        break
                # if description did not match any category, put amount into other expenses
                if not substring_found:
                    credit_card_data[month]['unknown'] += amount
    return credit_card_data
