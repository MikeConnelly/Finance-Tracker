import copy
import json
from typing import Dict, List, Tuple, Any

from bank_reader import parse_bank_data
from credit_card_reader import parse_credit_card_data
from finance_data import FinanceData
from writer import create_xlsx_file

CONFIG_FILE = './config/config.json'
BANK_ACTIVITY_DIR = './bank_activity'
CREDIT_CARD_ACTIVITY_DIR = './credit_card_activity'


def main():
    config = load_config_file(CONFIG_FILE)
    default_values = create_default_value_map(config)
    substring_map = create_substring_map(config)
    styles_map = create_styles_map(config)
    description_map = create_description_map(config)
    finance_data = FinanceData(default_values)

    parse_bank_data(finance_data, substring_map, BANK_ACTIVITY_DIR)
    parse_credit_card_data(finance_data, substring_map, CREDIT_CARD_ACTIVITY_DIR)

    create_xlsx_file(finance_data, styles_map, description_map)


def load_config_file(config_file: str) -> Dict[str, Dict[str, List[str]]]:
    """Load `config_file` into a `dict` mapping categories to lists of substrings."""
    with open(config_file) as config_json:
        return json.load(config_json)


def create_default_value_map(config: Dict[str, Dict[str, List[str]]]) -> Dict[str, Dict[str, float]]:
    """Create default value `dict` by replacing the substring lists in `config` with zeros."""
    default_values = copy.deepcopy(config)
    for major_category in default_values.keys():
        for minor_category in default_values[major_category].keys():
            default_values[major_category][minor_category] = 0
    return default_values


def create_substring_map(config: Dict[str, Dict[str, Any]]) -> List[Tuple[str, str, str]]:
    """Create a `list` of all substrings in `bank_config` tupled with their categories."""
    return [(major, minor, substring.lower())
            for major in config
            for minor in config[major]
            for substring in config[major][minor]['substrings']]


def create_styles_map(config: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Dict[str, str]]]:
    """"""
    styles_map = {}
    for major in config.keys():
        for minor in config[major].keys():
            if 'styles' in config[major][minor]:
                styles_map[minor] = config[major][minor]['styles']
    return styles_map


def create_description_map(config: Dict[str, Dict[str, Any]]) -> Dict[str, str]:
    """"""
    return {minor: description
            for major in config
            for minor in config[major]
            for description in config[major][minor]['description']}


if __name__ == '__main__':
    main()
