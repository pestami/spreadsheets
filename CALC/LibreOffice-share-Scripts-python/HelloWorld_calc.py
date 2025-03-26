# Created by Gweno 16/08/2019 for tutolibro.tech
# This program displays 'Hello World!" in cell A1 of the
# current Calc document.
def HelloWorld():
    """Write 'Hello World!' in Cell A1"""
# get the doc from the scripting context
# which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
# access the active sheet
    active_sheet = model.CurrentController.ActiveSheet
# write 'Hello World' in A1
    active_sheet.getCellRangeByName("J4").String = "Hello World!"

# get the Cell Range (nLeft, nTop, nRight, nBottom)  
#           =StartColumn =StartRow  =EndColumn  =EndRow
    oRangeSource1 = active_sheet.getCellRangeByPosition(8, 11, 22, 11)
      # example by name:
# ~ oRangeSource = active_sheet.getCellRangeByName('A1:C10')
    oRangeSource2 = active_sheet.getCellRangeByName('H12:V12')
    oRangeSource3 = active_sheet.getCellRangeByName('rng_CHILDREN')
    
       # get data from the Range of cells and store in a tuple
    oDataSource1 = oRangeSource1.getDataArray()
    oDataSource2 = oRangeSource2.getDataArray()
    oDataSource3 = oRangeSource3.getDataArray()
    
    # print to console
    print(oDataSource1)
    print(oDataSource2)
    print(oDataSource3)