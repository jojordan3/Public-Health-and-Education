import os, sys, subprocess

# Also used sas7bdat_to_csv *.sas7bdat

#and openpyxl
#import csv
#import openpyxl
#from openpyxl import load_workbook
# for file, prefix in excelfiles.items():
#     wb = load_workbook(file)
#     sheets = wb.get_sheet_names()
#     for sheet in sheets[1:]:
#         df = wb[sheet]
#         with open(f'{prefix}{sheet}.csv', 'w', newline='') as csvfile:
#             c = csv.writer(csvfile)
#             data =[]
#             for row in df.rows:
#                 data.append([cell.value for cell in row])
#             c.writerows(data)

# for data conversion

# Get database name from arguments passed to the script
# Alternative you could set explicitly e.g. `DATABASE = 'my-access-db.mdb'`

DATABASE = sys.argv[1]

# Get table names using mdb-tables
table_names = subprocess.Popen(['mdb-tables', '-1', DATABASE], stdout=subprocess.PIPE).communicate()[0]
tables = table_names.decode("utf-8").split('\n')

# Walk through each table and dump as CSV file using 'mdb-export'
# Replace ' ' in table names with '_' when generating CSV filename
for table in tables:
    if table != '':
        filename = table.replace(' ','_') + '.csv'
        print('Exporting ' + table)
        with open(filename, 'wb') as f:
            subprocess.check_call(['mdb-export', DATABASE, table], stdout=f)
