import os
import sys
import datetime

def format_value(col_value):
    '''
    param :col_value: - value of a column in a row
    returns           - column value formatted in a way which is
                        suitable for insertion into database
    '''
    if type(col_value) == datetime.datetime:
        return f"'{str(col_value.date())}'"
    elif type(col_value) == str:
        return f"'{col_value}'"
    else:
        return col_value

def retrieve_path_and_dbname(in_path, in_db):
    '''
    param :in_path:   - entered file path (spreadsheet)
    param :in_db:     - entered database name (can be `None`)
    returns           - [tuple] path to spreadsheet, name of database

    handles following Exceptions:
        * path does not exists/path is not a file
        * format of file is not supported
    '''
    try:
        if(not os.path.isfile(in_path)):
            raise Exception('Path does not exists or is not of a file')
        
        filename = os.path.split(in_path)[1]
        dbname, extension = filename.split(".")

        #if in_db is not null make it dbname otherwise
        #the above derived name will be used
        dbname = in_db if in_db is not None else dbname

        supported_formats = tuple({'xlsx', 'xlsm', 'xlxt', 'xltm'})
        if (extension not in supported_formats):
            raise Exception(f'''Format not supported '{extension}'.
Supported formats are: .xlsx,.xlsm,.xltx,.xltm''')

        return (in_path, dbname + '.db')

    except Exception as e:
        print(e)
        sys.exit(1)

def get_datatypes(row):
    '''
    param :iter:  - single row of data
    returns       - [list] datatype of each column for sql query
    '''
    datatypes = list()
   
    for item in row:  
        type_of_item = type(item)
        
        if type_of_item == int:
            datatypes.append('INTEGER')
        elif type_of_item == float:
            datatypes.append('REAL')
        else:
            datatypes.append('STRING')
    
    return datatypes
