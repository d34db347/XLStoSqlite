'''
@author: Natalie
@attention: Will not work for password protected workbooks. This is a limitation of xlrd.
@attention: As far as I know, you also cannot password protect sqlite databases...
@attention: This program creates a very BASIC table in a database.
    Database integrity issues such as assigning a primary/foreign key, etc. are not considered.
    Interactive version may be forked from this version. Stay tuned!
@todo: Database integrity
@attention: Header row must be consistent across sheets. Either all sheets have a header row, or all do not. 
    This is a noted design flaw that will be considered in future implementations.
    This may require an interactive version of this program.
@todo: All or none situation with header rows
@attention: At this time, only allowing alpha-numeric names for tables. If workbook is named something like
    'book.123.xls', the tables created will be prefixed by 'book123'
@param -i: str, This is the name of the Excel file to be copied into the table. 
@param -o: str, This is the name of the sqlite database the file is to be copied to.
@param -nohdr: boolean, Indicates sheets in excel workbook have no header
@param -buffer: int, number of rows to insert at a time, default = 100
@raise ValueError: Raises value error if filetype is not xls or xlsx
@raise IOError: Raises IO error if cannot open Excel file. A bad location may have been supplied, or the file is corrupted or 
    password protected.
@raise Any error that can be raised by sqlite3. http://legacy.python.org/dev/peps/pep-0249/

    
'''
import sqlite3
import xlrd
import argparse
import createTable
import fillTable
import os

# set up command line arguments
parser = argparse.ArgumentParser(description='Create new table in database with an excel file.')
parser.add_argument('-i', metavar='Input File', required=True,
                   help='Full path of Excel file to be used for input. Note, file must not be password protected.')
parser.add_argument('-o', metavar='Output Database', required=True,
                   help='Full path to database to write file to.')
parser.add_argument('-nohdr', default=False, action='store_true',
                   help='Use -nohdr if your data does not have a header row (Default: data has header row.)')
parser.add_argument('-buffer', default=100,
                   help='Specify the number of rows to be inserted at a time. (Default: 100 rows')

# parse out the args
args = parser.parse_args()
# database
db = args.o
# input xls file
xls = args.i
(xls_path, xlsfile) = os.path.split(xls)

# parse xls file name to see if it's good
goodext = ['xls', 'xlsx']
xt = xlsfile.split('.')
if xt[len(xt) - 1] not in goodext:
    raise ValueError('Input file must be .xls or .xlsx filetype')
# get name of workbook
workbook_name = xt[0] if len(xt)==2 else ''.join(xt[:-1])
# does this data have header rows?
header = not args.nohdr
# how many rows to be inserted at a time
buff = args.buffer

# open database
conn = sqlite3.connect(db)
c = conn.cursor()

try:
    # open xls file
    workbook = xlrd.open_workbook(xls)
except:
    raise IOError('Cannot open Excel file. ' + xls)
    
# for every sheet on this workbook
for sheet in workbook.sheet_names():
    try:
        (startrow, tablename, datatypes) = createTable.create(workbook_name, workbook.sheet_by_name(sheet), sheet, c, header)
        conn.commit()
    except:
        raise 
    
    # if there is data to move
    if tablename is not None:
        # fill the table
        try:
            fillTable.fill(workbook.sheet_by_name(sheet), startrow, c, tablename, workbook.datemode, datatypes, buff, conn)
        except:
            raise
        
# commit changes and close db
conn.commit()
conn.close()



