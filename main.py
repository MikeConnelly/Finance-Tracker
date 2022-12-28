import copy
import json
from typing import Dict, List, Tuple

from bank_reader import parse_bank_data
from credit_card_reader import parse_credit_card_data
from finance_data import FinanceData
from writer import create_xlsx_file
from writers.styles import Styles

"""
Config Type Structure
{
    major_category: {
        minor_category: {
            description: 'description'
            substrings: ['substring 1', 'substring 2']
            styles: {
                cell_type: {
                    property: value
                }
            }
        }
    }
}
"""
Config = Dict[str, Dict[str, Dict[str, str | list[str] | Dict[str, Dict[str, str]]]]]

CONFIG_FILE = './config/config.json'
BANK_ACTIVITY_DIR = './bank_activity'
CREDIT_CARD_ACTIVITY_DIR = './credit_card_activity'


def main():
    config = load_config_file(CONFIG_FILE)
    default_values = create_default_value_map(config)
    substring_map = create_substring_map(config)
    description_map = create_description_map(config)
    custom_styles = create_custom_styles_map(config)
    finance_data = FinanceData(default_values)

    parse_bank_data(finance_data, substring_map, BANK_ACTIVITY_DIR)
    parse_credit_card_data(finance_data, substring_map, CREDIT_CARD_ACTIVITY_DIR)

    create_xlsx_file(finance_data, custom_styles, description_map)


def load_config_file(config_file: str) -> Config:
    """Load `config_file` into a dict of type `Config`."""
    with open(config_file) as config_json:
        return json.load(config_json)


def create_default_value_map(config: Config) -> Dict[str, Dict[str, float]]:
    """Create a `dict` that maps all categories in `config` to zeros."""
    default_values = copy.deepcopy(config)
    for major_category in default_values.keys():
        for minor_category in default_values[major_category].keys():
            default_values[major_category][minor_category] = 0
    return default_values


def create_substring_map(config: Config) -> List[Tuple[str, str, str]]:
    """Create a `list` of all substrings in `config` tupled with their categories."""
    return [(major, minor, substring.lower())
            for major in config
            for minor in config[major]
            for substring in config[major][minor]['substrings']]


def create_custom_styles_map(config: Config) -> Styles:
    """Creata a `dict` that maps all minor categories in `config` to their styles object."""
    custom_styles = {}
    for major in config.keys():
        for minor in config[major].keys():
            if 'styles' in config[major][minor]:
                custom_styles[minor] = config[major][minor]['styles']
    return custom_styles


def create_description_map(config: Config) -> Dict[str, str]:
    """Create a `dict` taht maps all minor categories in `config` to their description string."""
    return {minor: description
            for major in config
            for minor in config[major]
            for description in config[major][minor]['description']}


if __name__ == '__main__':
    main()
