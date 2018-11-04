# -*- coding: utf-8 -*-

#TODO: passing date as a tuple; column name from column number

import uno

# getStringFromShiftedDate( date, shift ) -
# returns date that is "shift" month past
# the "date" given. Day is automatically set to 01
#
# Parameters:
#   date - represents initial date
#       Type: tuple of 2 elements: ( year, month )
#
#   shift - number of months to move forward
#       Type: integer
#
# Return value:
#   String of the format: yyyy;mm;01
#   Type: string
def getStringFromShiftedDate( date, shift ):
    return (
    #calculate year
    str(date[0]+((date[1]-1+shift) // 12)) +
    ";" +
    #calculate month
    str(((date[1]-1+shift) % 12) + 1) +
    ";01"
    )

# getColumnNameFromNumber( columnNumber ) -
# converts a column number to the corresponding
# column name in OpenOffice column naming fashion
# (the same as in MS Excel: A, B, C,.., Y, Z, AA, AB,..)
#
# Parameters:
#   columnNumber - number to be converted
#   Type: integer
#
# Return Value:
#   Name of the column with the given number
#   Type: string
def getColumnNameFromNumber( columnNumber ):
    retval = ""
    
    if( columnNumber < 26 ):
        return chr(ord('A')+columnNumber)
    
    while columnNumber >= 26:
        retval = chr((columnNumber % 26)+ord('A')) + retval
        columnNumber = columnNumber // 26
    
    retval = chr(ord('A')+columnNumber-1) + retval
    return retval

# setCell( cell, data ) -
#   puts "data" into the given "cell"
#
# Parameters:
#   cell - the cell to be modified
#       Type: internal UNO type
#
#   data - data to be inserted
#       Type: OO TYPED CELL (see createTable function description)
#
# Notes:
#   This function uses dictionary to call the method,
#   depending on the type of the data that will be inserted
def setCell( cell, data ):
    { "formula":cell.setFormula, "string":cell.setString, "value":cell.setValue }[data[0]](data[1])
    
# createTable( sheet, offset, columns, items, values ) -
# creates an open office table in the given sheet, of the format:
#
# EmptyCell | columnName1 | columnName2 | columnName3 ...
# ----------+-------------+-------------+-------------
# itemName1 | value[0][0] | value[0][2] | value[0][3] ...
# ----------+-------------+-------------+-------------
# itemName2 | value[1][0] | value[1][2] | value[1][3] ...
# ----------+-------------+-------------+-------------
# itemName3 | value[2][0] | value[2][2] | value[2][3] ...
# ----------+-------------+-------------+-------------
# ...
#
# PLEASE, READ PARAMETERS SECTION TO GET TO KNOW
# WHAT columns, items AND values ARE
#
# Parameters:
#   sheet - a sheet in an openoffice calc file, where
#       the table will be created
#       Type: internal UNO type
#    
#   offset - offset of the EmptyCell from the top left cell
#       Type: tuple of 2 items: ( leftOffset, topOffset )
#
#   columns - names of each column in the table
#       Type: SEE SECTION BELOW
#
#   items - names of each item in the table
#       Type: SEE SECTION BELOW
#
#   values - the very values of the table
#       Type: SEE SECTION BELOW
#
#   SECTION BELOW:
#       "columns" and "items" have the same type -
#       list of tuples of 2 elements in each
#       and the only difference between them
#       and "values" is that values is a 2D list,
#       whereas the first two are 1D lists.
#       Tuples conform to this format: ( type, value )
#       Type is either "string" or "formula" - depending
#       on that, either setString or setFormula method
#       is called in setCell function.
#       This format is referred to below as "OO TYPED CELL"
def createTable( sheet, offset, columns, items, values ):
    for i in range(len(columns)):
        setCell(sheet.getCellByPosition( offset[0]+i+1, offset[1] ), columns[i])

    for i in range(len(items)):
        setCell(sheet.getCellByPosition( offset[0], offset[1]+i+1 ), items[i])
        
    for i in range(len(values)):
        for j in range(len(values[i])):
            setCell(sheet.getCellByPosition( offset[0]+j+1, offset[1]+i+1 ), values[i][j])

# createCashflowTable( sheet, offset, items, timePeriod, beforeSummaryShift, values = None ) -
# creates a table of the format given in description to the createTable function,
# after listing all the "items", leaves "beforeSummaryShift" rows blank and right below them
# puts table summary information.
#
# TODO: 1. Add more summary items
#       2. Add charts creating
#
# Parameters:
#   sheet - sheet in the OpenOffice file
#       Type: internal UNO type
#
#   offset - location of the top-left cell of the table, 0-indexed
#       Type: tuple of 2 items: ( leftShift, topShift )
#
#   items - names of the table contents (rows)
#       Type: list of OO TYPED CELLs
#
#   timePeriod - names of table time periods (columns)
#       Type: list of OO TYPED CELLs
#
#   beforeSummaryShift - number of rows to leave blank before
#       summary info
#       Type: integer
#
#   values - table contents (optional), if omitted, table contents left empty
#       Type: 2D list of OO TYPED CELLs
def createCashflowTable( sheet, offset, items, timePeriod, beforeSummaryShift, values = None ):
    summaryData = [ ( "formula", "=SUM(" + getColumnNameFromNumber(offset[0]+i+1)+str(offset[1]+2) + ":" + getColumnNameFromNumber(offset[0]+i+1) + str(offset[1]+len(items)+1) + ")" ) for i in range(len(timePeriod)) ]
    
    items.extend(([ ("string", "") ] * beforeSummaryShift) + [("string", "Total")])
    
    if values is not None:
        values.extend( [summaryData] )
    else:
        values = ([[]] * (len(items) - 1)) + [summaryData]
        
    createTable( sheet, offset, timePeriod, items, values )
        

# generateFinancialPlanTables( offset, itemsNow, itemsSoon, startDate, months, beforeSummaryShift, spaceBetweenTables ) -
#   the very function to be called to generate financial plans in OpenOffice Calc sheets.
#
# TODO: 1. Add a possibility to rename sheets and put each table on its own sheet,
#       2. Add document saving
#
# Parameters:
#   offset - top-left position of the first table
#       Type: a tuple of type: ( leftOffset, topOffset )
#
#   itemsNow - items of the first table
#       Type: list of OO TYPED CELLs
#
#   itemsSoon - items of the second table
#       Type: list of OO TYPED CELLs
#
#   startDate - first time period segment
#       Type: tuple of 2 items ( year, month )
#
#   months - numbers of months to generate the tables for
#       Type: integer
#
#   beforeSummaryShift - space (number of rows) to leave (blank)
#       between table contents and table summary info
#
#   spaceBetweenTables - used when both tables are put into
#       the same sheet - number of rows between them
#       Type: integer
def generateFinancialPlanTables( offset, itemsNow, itemsSoon, startDate, months, beforeSummaryShift, spaceBetweenTables ):
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    
    oDoc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
    oSheet = oDoc.getSheets().getByIndex( 0 ) # get the zero'th sheet 
    #oSheet = oDoc.getSheets().getByName( "Sheet3" ) # get by name 
    
    monthsData = [ ("formula", "=DATE(" + getStringFromShiftedDate(startDate, i) + ")") for i in range(months) ]
    
    createCashflowTable( oSheet, offset, itemsNow, monthsData, beforeSummaryShift )
    createCashflowTable( oSheet, (offset[0], offset[1] + len(itemsNow) + beforeSummaryShift + 1 + spaceBetweenTables), itemsSoon, monthsData, beforeSummaryShift )
    
    #oDoc.close( True )

# Testing call    
generateFinancialPlanTables(( 1, 2 ),[("string","Salary"), ("string","Electricity"), ("string","Other"), ("string","Taxes")],[("string","Robot Electricity"), ("string","Robot Maintaining")],( 2018, 11 ),80,2,3)