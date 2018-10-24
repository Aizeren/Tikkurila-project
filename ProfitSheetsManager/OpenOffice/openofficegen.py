# -*- coding: utf-8 -*-

#TODO: passing date as a tuple; column name from column number

import uno

def getStringFromShiftedDate( date, shift ):
    return (
    #calculate year
    str(date[0]+((date[1]-1+shift) // 12)) +
    ";" +
    #calculate month
    str(((date[1]-1+shift) % 12) + 1) +
    ";01"
    )

def getColumnNameFromNumber( columnNumber ):
    retval = ""
    columnNumber += 26
    while columnNumber > 1:
        retval = chr((columnNumber % 27)+ord('A')) + retval
        columnNumber = columnNumber // 27
    return retval

def genOpenOfficeCalc( offset, items, startDate, months, beforeSummaryShift ):
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    
    oDoc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
    oSheet = oDoc.getSheets().getByIndex( 0 ) # get the zero'th sheet 
    #oSheet = oDoc.getSheets().getByName( "Sheet3" ) # get by name 

    #----- 
    # Put some sales figures onto the sheet. 
    for i in range(len(items)):
        oSheet.getCellByPosition( offset[1], i+offset[0]+1 ).setString( items[i] )
    
    for i in range(months):
        oSheet.getCellByPosition( offset[1]+i+1, offset[0] ).setFormula( "=DATE(" + getStringFromShiftedDate(startDate, i) + ")" )
        oSheet.getCellByPosition( offset[1]+i+1, offset[0]+len(items)+beforeSummaryShift+1 ).setFormula(
        "=SUM(" + getColumnNameFromNumber( offset[1]+i+1 ) +
        str(offset[0]+2) +
        ":" + getColumnNameFromNumber( offset[1]+i+1 ) +
        str(offset[0]+len(items)+1) +
        ")"
        )
        
    oSheet.getCellByPosition( offset[1], offset[0]+len(items)+beforeSummaryShift+1 ).setString( "Total" )
    
    #oDoc.close( True )
    
genOpenOfficeCalc(( 2, 1 ),["Salary", "Electricity", "Other", "Taxes"],( 2018, 11 ),80,2)