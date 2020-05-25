from openpyxl import load_workbook
import sqlite3
import sys
import argparse
import helper

parser = argparse.ArgumentParser(formatter_class=argparse.RawDescriptionHelpFormatter,
                                description = 'DBTool : Convert spreadsheet to SQlite3 Database',
                                epilog='''Examples
            python dbtool.py -i /home/test.xlsx -o my

            python dbtool.py --in /home/example.xlsx

            python dbtool --in /home/this_file.xlsz --out mydatabase
            ''')

parser.add_argument('-i',
                    '--infile',
                    help="Input file path (spreadsheet)",
                    required=True
                    )
parser.add_argument('-o',
                    '--outfile',
                    help="output file name (database)",
                    required=False
                    )

args = parser.parse_args()

path, db_name = helper.retrieve_path_and_dbname(args.infile, args.outfile)

workbook = load_workbook(path, read_only=True)
table_name = workbook.sheetnames[0].upper()
sheet = workbook.active
sheet_iter = sheet.values #iterator containing every row

#first row of spreadsheet
column_names = tuple(next(sheet_iter))
column_names = tuple(
            map(
                lambda elem: elem.upper().replace(" ", "_"), #make column names uppercase and
                column_names                                 #replace spaces with underscore
               )
            )

#:arg: - [tuple] second row of spreadsheet 
datatypes = helper.get_datatypes(tuple(next(sheet_iter)))

create_table_query = f'create table {table_name} ('
joined_colName_and_types = ', '.join(f"{i} {j}" for (i,j) in zip(column_names, datatypes))
create_table_query += joined_colName_and_types + ');'

#reset iterator
sheet_iter = sheet.values
next(sheet_iter) #ignore first row

#establish db connection
conn = sqlite3.connect(db_name)
cur = conn.cursor()

#execute all sql queries
try:
    cur.execute(create_table_query)
    while(True):
        row = map(lambda x: helper.format_value(x), next(sheet_iter))
        insert_query = f'insert into {table_name} values('
        joined_rows  = ', '.join(f'{i}' for i in row)
        insert_query += joined_rows + ');'
        cur.execute(insert_query)

except StopIteration:
    conn.commit()
finally:
    conn.close()
    workbook.close()
