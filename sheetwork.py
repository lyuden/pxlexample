# coding: utf-8


import uno
import subprocess

COMMAND = ['libreoffice', '--headless', '--calc', '--accept="socket,host=localhost,port=8100;urp;"' ]


def open_calc_file(path_to_file):

    
    p = subprocess.Popen(COMMAND + [path_to_file, ], shell = True)
    

    return p
    

def set_cells(cellist,  sheet_index = 0):
    
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
				"com.sun.star.bridge.UnoUrlResolver", localContext )
    localContext = uno.getComponentContext()
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
				"com.sun.star.bridge.UnoUrlResolver", localContext )
    ctx = resolver.resolve( "uno:socket,host=localhost,port=8100;urp;StarOffice.ComponentContext" )
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
    model = desktop.getCurrentComponent()

    sheet1=model.Sheets.getByIndex(sheet_index)

    for cell_row, cell_column, cell_value in cellist:
        
        cell =sheet1.getCellByPosition(cell_row,cell_column)
        cell.setValue(cell_value)

    model.calculateAll()
    model.storeSelf(())

    
