import os
import sys
import datetime

def format_value(col_value):
    if type(col_value) == datetime.datetime:
        return f"'{str(col_value.date())}'"
    elif type(col_value) == str:
        return f"'{col_value}'"
    else:
        return col_value

def retrieve_path_and_dbname(arguments):
    try:
        supported_formats = tuple({'xlsx', 'xlsm', 'xlxt', 'xltm'})

        if (len(arguments) < 2):
            raise Exception('Path to file not specified')
        elif (len(arguments) > 2):
            raise Exception('Too many input arguments')

        path = arguments[1]
        
        if(not os.path.isfile(path)):
            raise Exception('Path does not exists or is not of a file')
        
        filename = os.path.split(path)[1]
        dbname, extension = filename.split(".")

        if (extension not in supported_formats):
            raise Exception(f'''Format not supported '{extension}'.
Supported formats are: .xlsx,.xlsm,.xltx,.xltm''')

        return (path, dbname + '.db')

    except Exception as e:
        print(e)
        sys.exit(1)

def get_datatypes(iter):
    datatypes = list()
   
    for item in next(iter):  
        type_of_item = type(item)
        
        if type_of_item == int:
            datatypes.append('INTEGER')
        elif type_of_item == float:
            datatypes.append('REAL')
        else:
            datatypes.append('STRING')
    
    return datatypes
