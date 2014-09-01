'''
Created on Aug 31, 2014

@author: Natalie
@attention: If a table already exists with the name 'workbook_name'_'sheet_name', existing will be DROPPED and new table created.
    No distinction is made regarding filetype (ie - workbook.xls and workbook.xlsx will be treated the same)
@attention: Duplicate headers are not allowed at this time. This is because sqlite3 throws an OperationalError for both 
    1) Creating a table that already exists (Can work around this one)
    2) Creating columns with duplicate names (But not this one)
    The answer to this problem is to create nested try statements.
@bug: Fix for duplicate column names.
@param workbook_name: str, the name of the workbook, extension not included. This is used to name the table.
@param worksheet: an xlrd sheet object, 
    https://secure.simplistix.co.uk/svn/xlrd/trunk/xlrd/doc/xlrd.html?p=4966#sheet.Sheet-class
@param sheet_name: str, the name of this worksheet. This is used to create the name of the table.
@param cursor: sqlite3 cursor object, used to execute CREATE statement. 
    https://docs.python.org/2/library/sqlite3.html#sqlite3.Cursor
@param header: boolean, true if sheet has a header row

@return tuple(startDataRow, tablename, datatypes): 
    startDataRow: row in worksheet to begin copying data (0 or 1), None if there is no data to be copied
    tablename: 'workbook_name'_'sheet_name', None if there is no data to be copied
    datatypes: sqlite datatypes to be used for each row

'''

# helper to scrub column names
# todo THIS NEEDS WORK!, what if a sheet is named something all crazy like '((('
# research what is allowed in names of xls sheets 
def scrub(colname):
    scrubus = [';', ',', '(', ')', ' ', '!']
    scrubbed = colname
    for s in scrubus:
        scrubbed = scrubbed.split(s)
        scrubbed = ''.join(scrubbed)
    return scrubbed
        
def create(workbook_name, worksheet, sheet_name, cursor, header):
    # is this an empty sheet?
    if worksheet.nrows == 0:
        return (None, None, None)
    # only the header row
    elif worksheet.nrows == 1 and header:
        return (None, None, None)
    # header row plus some data or just data
    else:
        # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
        # corresponding map to sqlite data types
        typemap = ['NULL', 'TEXT', 'REAL', 'DATE', 'TEXT', 'TEXT', 'TEXT' ]
        # name this table
        tablename = scrub(workbook_name + '_' + sheet_name)
        
        # get column names and data types
        firstRow = worksheet.row(0)
        col_names = []
        datatypes = []
        if header:
            # checked earlier, if there is a header row
            # there must be at least one additional row
            # column names are the first row
            col_names = [scrub(cell.value) for cell in firstRow]
            
            # datatypes are the second row
            # map cell type to sqlite data types
            datatypes = [typemap[cell.ctype] for cell in worksheet.row(1)]
            startDataRow = 1
        else:
            # datatypes in first row
            # name columns
            x = 1
            for cell in firstRow:
                datatypes.append(typemap[cell.ctype])
                col_names.append('COL' + str(x))
                x = x + 1
            startDataRow = 0
        # if one or more of the datatypes is null, scan the column until a non null datatype is found
        nullcols = [i for i, j in enumerate(datatypes) if j == 'NULL']
        for n in nullcols:
            # condense a column slice to a set
            # ignore the first row - it is either null, or a header(TEXT)
            ctypes = set(worksheet.col_types(n)[1::])
            # remove the null type if it's still  there
            ctypes.discard('NULL')
            # if all the values are null, use TEXT
            # if there are mixed types, use TEXT... ain't nobody got time for that
            newDT = 'TEXT'
            if len(ctypes) == 1:
                newDT = typemap[ctypes.pop()]
            datatypes[n] = newDT    
            
            
        # create table
        # NOTE no scubbing done here! THIS IS VERY DANGEROUS!!
        createStr = 'CREATE TABLE ' + tablename + '(' + ','.join(map(lambda (x): str(x[0]) + ' ' + str(x[1]), zip(col_names, datatypes))) + ')'
        try:
            cursor.execute(createStr)  
        except:
            try: 
                # drop the table and create a new one
                dropStr = 'DROP TABLE ' + tablename 
                cursor.execute(dropStr)
                cursor.execute(createStr)
            except:
                raise
        return (startDataRow, tablename, datatypes)