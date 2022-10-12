import json
from bank_reader import parse_bank_data
from credit_card_reader import parse_credit_card_data
from writer import create_xlsx_file

BANK_CONFIG_FILE = './config/bank_config.json'
BANK_ACTIVITY_DIR = './bank_activity'
CREDIT_CARD_CONFIG_FILE = './config/credit_card_config.json'
CREDIT_CARD_ACTIVITY_DIR = './credit_card_activity'

def main():
    bank_data = parse_bank_data(load_config_file(BANK_CONFIG_FILE), BANK_ACTIVITY_DIR)
    credit_card_data = parse_credit_card_data(load_config_file(CREDIT_CARD_CONFIG_FILE), CREDIT_CARD_ACTIVITY_DIR)

    print(json.dumps(bank_data, indent=4))
    print(json.dumps(credit_card_data, indent=4))

    create_xlsx_file(bank_data, credit_card_data)

def load_config_file(config_file: str) -> dict:
    """
    Load `config_file` into a `dict` mapping categories to lists of substrings.
    Structure of the returned dictionary depends on the contents of the given config file.
    """
    with open(config_file) as config_json:
        return json.load(config_json)

if __name__ == '__main__':
    main()
