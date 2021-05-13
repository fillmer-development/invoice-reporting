from openpyxl.worksheet.worksheet import Worksheet
import openpyxl
import os
import csv
from configparser import ConfigParser


def main():
    config = ConfigParser()
    config.read('./config.ini')
    input = config['PATHS'].get('source_directory')
    output = config['PATHS'].get('output_path')
    mapping = dict(config['MAPPING'])

    with open(output, 'w') as out:
        header = True
        w = None
        for item in read_directory(input):
            row = parse_excel_file(load_excel_file(item), mapping=mapping)

            if not type(row) == dict:
                continue
            if header:
                w = csv.DictWriter(out, row.keys())
                w.writeheader()
                header = False

            w.writerow(row)


def load_excel_file(file):
    return openpyxl.load_workbook(file, data_only=True).active


def parse_excel_file(book: Worksheet, mapping={}):
    default_mapping = {
        'name': 'A9',
        'street': 'A10',
        'city': 'A11',
        'date': 'C16',
        'id': 'C17',
        'desc': 'B34',
        'sum': 'F100',
        **mapping
    }

    invoice_meta = {}
    for (key, value) in default_mapping.items():
        invoice_meta[key] = str(book[value].value)

    return invoice_meta


def read_directory(path='.'):
    for root, dirs, files in os.walk(path):
        for filename in files:
            if filename.endswith('.xlsx'):
                yield os.path.join(path, filename)
