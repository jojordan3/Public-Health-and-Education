import csv
import openpyxl
import sys
from openpyxl import load_workbook
from pathlib import Path

try:
    CSV_PATH = sys.argv[2]
except:
    CSV_PATH = 'data/raw_csvs/'

FILEPATH = sys.argv[1]


def convert_multisheet(FILEPATH=FILEPATH, CSV_PATH=CSV_PATH):
    prefix = Path(FILEPATH).stem
    wb = load_workbook(FILEPATH)
    sheets = wb.get_sheet_names()
    for sheet in sheets[1:]:
        df = wb[sheet]
        with open(f'{CSV_PATH}{prefix}{sheet}.csv', 'w', newline='') as csvfile:
            c = csv.writer(csvfile)
            data =[]
            for row in df.rows:
                data.append([cell.value for cell in row])
            c.writerows(data)

if __name__ == '__main__':
    convert_multisheet()
    
