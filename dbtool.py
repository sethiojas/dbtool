from openpyxl import load_workbook
import sqlite3
import sys
import helper

path, db_name = helper.retrieve_path_and_dbname(sys.argv)

workbook = load_workbook(path, read_only=True)
table_name = workbook.sheetnames[0].upper()
sheet = workbook.active
sheet_iter = sheet.values

column_names = tuple(next(sheet_iter))
column_names = tuple(map(lambda elem: elem.upper().replace(" ", "_"), column_names))

datatypes = helper.get_datatypes(sheet_iter)

create_table_query = f'create table {table_name} ('
joined_colName_and_types = ', '.join(f"{i} {j}" for (i,j) in zip(column_names, datatypes))
create_table_query += joined_colName_and_types + ');'

conn = sqlite3.connect(db_name)
cur = conn.cursor()

cur.execute(create_table_query)

sheet_iter = sheet.values
next(sheet_iter)
try:
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
