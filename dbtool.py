from openpyxl import load_workbook
import datetime
import sqlite3
import sys
import os

def foo(col_value):
    if type(col_value) == datetime.datetime:
        return f"'{str(col_value.date())}'"
    elif type(col_value) == str:
        return f"'{col_value}'"
    else:
        return col_value

path = sys.argv[1]
db_name = os.path.split(path)[1].split(".")[0]

conn = sqlite3.connect(db_name + ".db")
cur = conn.cursor()

workbook = load_workbook('test.xlsx', read_only=True)
table_name = workbook.sheetnames[0].upper()
sheet = workbook.active
sheet_iter = sheet.values

column_names = tuple(next(sheet_iter))
column_names = tuple(map(lambda elem: elem.upper().replace(" ", "_"), column_names))

datatypes = list()
for item in next(sheet_iter):
    type_of_item = type(item)
    
    if type_of_item == int:
        datatypes.append('INTEGER')
    elif type_of_item == float:
        datatypes.append('REAL')
    else:
        datatypes.append('STRING')

create_table_query = f'create table {table_name} ('
joined_colName_and_types = ', '.join(f"{i} {j}" for (i,j) in zip(column_names, datatypes))
create_table_query += joined_colName_and_types + ');'

cur.execute(create_table_query)

sheet_iter = sheet.values
next(sheet_iter)
try:
    while(True):
        row = map(lambda x: foo(x), next(sheet_iter))
        insert_query = f'insert into {table_name} values('
        joined_rows  = ', '.join(f'{i}' for i in row)
        insert_query += joined_rows + ');'
        cur.execute(insert_query)

except StopIteration:
    conn.commit()
finally:
    conn.close()
    workbook.close()
