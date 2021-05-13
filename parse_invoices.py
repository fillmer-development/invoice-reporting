from collections import defaultdict
import csv
import os
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet


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
