import csv
from parse_invoices import load_excel_file, parse_excel_file, read_directory
from sys import argv
from configparser import ConfigParser

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
