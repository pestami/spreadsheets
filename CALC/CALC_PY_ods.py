# Created by Gweno 16/08/2019 for tutolibro.tech
# https://github.com/Gweno/tutolibro.tech/blob/master/lopy/part9/ReverseRange.py
#C:\ProgrammeApps\LibreOffice\share\Scripts\python\
#===HELPER FUNCTIONS=================================================================
def getSelectionAddresses(horizontalOffset = 0 , verticalOffset = 0):
    # get the range of addresses from selection
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    oSelection = model.getCurrentSelection()
    oArea = oSelection.getRangeAddress()
    return oArea.StartColumn + horizontalOffset, oArea.StartRow + verticalOffset, oArea.EndColumn + horizontalOffset, oArea.EndRow + verticalOffset

#====================================================================    
def _Ranges():
 
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
    print('--Content of Ranges--------------')
    print(oDataSource1)
    print(oDataSource2)
    print(oDataSource3)
    print('----------------')
    oSelection = model.getCurrentSelection()   # is object
    oArea1 = oSelection.getRangeAddress()
    x1,y1,x2,y2 = getSelectionAddresses()
    
    print('--Selection Address--------------')
    print('Range Start Column= ' + str(oArea1.StartColumn) )
    print('Selection Start Column= ' + str(x1)+' '+ str(y1) +' '+  str(x2) +' '+  str(y2)  )
    print('----------------')
    #The *l idiom is to unpack argument lists when calling a function
    oRangeTarget1 = active_sheet.getCellRangeByPosition(*getSelectionAddresses(horizontalOffset = 3))
    oRangeTarget2 = active_sheet.getCellRangeByName('rng_Output1')
    
    print(oDataSource3[0])
    oRangeTarget2.setDataArray(oDataSource3)
    
#====================================================================
def _Execute():  
    import os
    #os.system('"C:/Windows/System32/notepad.exe"')
    os.system('"C:\PortableApps\SQLiteStudio\SQLiteStudio.exe C:/Users/sesa237770/Documents/Projects_DIY/2024-ROMS-DB/ROM_DB_PY/DB/RetroRoms_20240301.db"')
    
    
#====================================================================
def TEST_Calc_PY_Macro():
    _Ranges()
    #_Execute()
#====================================================================
g_exportedScripts = (TEST_Calc_PY_Macro,)
#====================================================================