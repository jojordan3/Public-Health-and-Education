import os, sys, subprocess

# Also used sas7bdat_to_csv *.sas7bdat for data conversion

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
