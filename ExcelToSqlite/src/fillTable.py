'''
Created on Aug 31, 2014

@author: Natalie
@param sheet: an xlrd sheet object, 
    https://secure.simplistix.co.uk/svn/xlrd/trunk/xlrd/doc/xlrd.html?p=4966#sheet.Sheet-class
@param start: int, first row in sheet that contains data. Values are 0 or 1, generated in createTable.py
@param cursor: sqlite3 cursor object, used to execute SQL statements. 
    https://docs.python.org/2/library/sqlite3.html#sqlite3.Cursor
@param tablename: the name of the table
@param datemode: attribute of xlrd book object
    http://www.lexicon.net/sjmachin/xlrd.html#xlrd.Book-class
@param datatypes: sqllite data types for each of the columns in this sheet, determined by createTable.py
@param buffer: int, user defined number of rows to be inserted at a time, default is 100 (see readFromXls.py parameters)
@param conn: sqlite3 connection object, used to commit transactions

'''
import xlrd

# convert date to string
def convertToDate(d, datemode):
    tup = xlrd.xldate_as_tuple(d, datemode)
    # xlrd.xldate_as_tuple returns Gregorian (year, month, day, hour, minute, nearest_second). 
    return str(tup[0]) + '-' + str(tup[1]) + '-' + str(tup[2]) + ' ' + str(tup[3]) + ':' + str(tup[4]) + ':' + str(tup[5])

def fill(sheet, start, cursor, tablename, datemode, datatypes, buff, conn):
    # build insert string
    qmarks = '?,' * sheet.ncols
    qmarks = qmarks[:-1]
    insertStr = 'INSERT INTO ' + tablename + ' VALUES (' + qmarks + ')'
    
    # formatting note
    # Larger example that inserts many records at a time
    #purchases = [('2006-03-28', 'BUY', 'IBM', 1000, 45.00),
    #        ('2006-04-05', 'BUY', 'MSFT', 1000, 72.00),
    #        ('2006-04-06', 'SELL', 'IBM', 500, 53.00),
    #      ]
    #c.executemany('INSERT INTO stocks VALUES (?,?,?,?,?)', purchases)
    
    # which columns require formatting?
    datecols = [i for i, j in enumerate(datatypes) if j == 'DATE']   
    
    # this is a little ugly, and I'm sorry. No elegant way to deal with startrow being either 0 or 1
    # number of insert statements (executemany)
    ninserts = (sheet.nrows/buff) + 1
    # EXAMPLE if sheet.numrows = 276
    #ranges = [[start, 100], [100, 200], [200, sheet.numrows]]
    ranges = [[i*100,(i+1)*100] for i in range(ninserts)]
    # modify start entry
    ranges[0][0] = start
    # modify last entry
    ranges[ninserts - 1][1] = ((ninserts - 1)*buff) + sheet.nrows%buff
    for i in range(ninserts):
        vals = []
        for n in range(ranges[i][0], ranges[i][1]):
            l =  [cell.value for cell in sheet.row(n)]
            #convert dates
            for d in datecols:
                l[d] = convertToDate(l[d], datemode)
            vals.append(tuple(l))
        cursor.executemany(insertStr, vals)
        conn.commit()
   
    
    
    
    
        
    
        
    

