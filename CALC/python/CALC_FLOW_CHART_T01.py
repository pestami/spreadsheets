#==================================================================================== 
# Created by Gweno 16/08/2019 for tutolibro.tech
# https://github.com/Gweno/tutolibro.tech/blob/master/lopy/part9/ReverseRange.py
#C:\ProgrammeApps\LibreOffice\share\Scripts\python\
    #HELP=============  
    #print(dir(oCell))
    #Help(oCell)
    #print("\n".join(sorted(dir(oCell), key=lambda s: s.lower())))
#====================================================================================    
#===HELPER FUNCTIONS=================================================================
def _getSelectionAddresses(horizontalOffset = 0 , verticalOffset = 0):
    # get the range of addresses from selection
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    oSelection = model.getCurrentSelection()
    oArea = oSelection.getRangeAddress()
    return oArea.StartColumn + horizontalOffset, oArea.StartRow + verticalOffset, oArea.EndColumn + horizontalOffset, oArea.EndRow + verticalOffset

#====================================================================    
#===LIBRARY FUNCTIONS=================================================================
def _getVariables():

    print('\n--_getVariables()--------------')
    # get the range of addresses from selection
    oDesktop = XSCRIPTCONTEXT.getDesktop()
    oModel = oDesktop.getCurrentComponent()
    
       # print("\n".join(sorted(dir(oModel.CurrentController), key=lambda s: s.lower())))
       # oSheet = model.CurrentController.ActiveSheet
       #oModel.CurrentController.setActiveSheet.getByName("SETUP")
       
    #oSheet = oModel.CurrentController.getByName("SETUP")
    
    oSheet = oModel.CurrentController.ActiveSheet
    oSheet2 = oModel.getSheets().getByName('SETUP')
    oModel.CurrentController.setActiveSheet(oSheet2)
    
    oRange = oSheet2.getCellRangeByName('rngVARIABLE')
    
    tRange = oRange.getDataArray()
    print('\n--_getVariables()--------------')
    #print(tRange)
    
    
    oModel.CurrentController.setActiveSheet(oSheet)
    
    tVARS=tRange
    lVARS=[]
    for i in list(tVARS):
        lVARS.append(list(i))

    dVARS_WIN={}
    for i in lVARS:
        dVARS_WIN[i[0]] = i[1]

    dVARS_LIN={}
    for i in lVARS:
        dVARS_LIN[i[0]] = i[2]
    print('Dictionary WIN:')
    print(dVARS_WIN)

    print('-------------------------------')
    
    return dVARS_WIN , dVARS_LIN
#--------------------------------------------------------------------    
def _ResolveVariables(sCMD):
    
    print('\n--_ResolveVariables(sCMD)--------------')
    import re 
    #sCMD=sCMD.replace('%ROOT%', 'C:/Users/sesa237770/Documents/ArcadeMeta')
    
    dVARS_WIN,dVARS_LIN, = _getVariables()

    sPath=sCMD
    ReString='%[A-Za-z0-9_-]*%'
    resVAR= re.search(ReString,sPath,0)
    if resVAR:
        sVAR=resVAR.group(0)
        print('RE Result=' + sVAR)    
        sCMD=sCMD.replace(sVAR, dVARS_WIN[sVAR])
    
    sCMD=sCMD.replace('\n', '')
    sCMD=sCMD.replace('\r', '')
    sCMD=sCMD.replace('/', '\\')      

    return sCMD
#============================================================================
def _getProgramPath():

    print('\n--_getProgramPath()--------------')
    # get the range of addresses from selection
    oDesktop = XSCRIPTCONTEXT.getDesktop()
    oModel = oDesktop.getCurrentComponent()    
    
    oSheet = oModel.CurrentController.ActiveSheet
    oSheet2 = oModel.getSheets().getByName('SETUP')
    oModel.CurrentController.setActiveSheet(oSheet2)
    
    oRange = oSheet2.getCellRangeByName('rngPROGRAMS')
    
    tRange = oRange.getDataArray()
    print('\n--_getVariables()--------------')
    #print(tRange)
    
    
    oModel.CurrentController.setActiveSheet(oSheet)
    
    tVARS=tRange
    lVARS=[]
    for i in list(tVARS):
        lVARS.append(list(i))

    dVARS_WIN={}
    for i in lVARS:
        dVARS_WIN[i[0]] = i[1]

    dVARS_LIN={}
    for i in lVARS:
        dVARS_LIN[i[0]] = i[2]
    print('Dictionary WIN:')
    print(dVARS_WIN)

    print('-------------------------------')
    
    return dVARS_WIN , dVARS_LIN
#--------------------------------------------------------------------    
def _ResolveProgramPath(sAPPNAME):
    print('\n--_ResolveVariables(sCMD)--------------')
    import re 
    
    #sCMD=sCMD.replace('%ROOT%', 'C:/Users/sesa237770/Documents/ArcadeMeta')
    dVARS_WIN,dVARS_LIN, = _getProgramPath()
    sPathProgramName=dVARS_WIN[sAPPNAME]   

    return sPathProgramName
#====================================================================
# 'PyScripter'  
def ExecuteProgram(sAPPNAME):     
    import os
    #os.system('"C:/Windows/System32/notepad.exe"')
    print('\n--ExecuteProgram()--------------')
    print('sAPPNAME= ' + sAPPNAME)
    
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    oSheet = model.CurrentController.ActiveSheet
    oSelection = model.getCurrentSelection()
    
    sSelectionType=oSelection.ElementType.typeName
    print("ExecuteProgram Selected: " +sSelectionType)
    
    if sSelectionType=='com.sun.star.text.XTextRange':     
        sCMD = oSelection.String        #setString            
    
    if sSelectionType=='com.sun.star.drawing.XShape':     
        oShape = oSelection.getByIndex(0)
        sCMD = oShape.String        #setString   
     
    sCMD=_ResolveVariables(sCMD)               
        
    sAPP='\"'+ _ResolveProgramPath(sAPPNAME) + ''     
    sCMD = sAPP + ' ' + sCMD +'\"'   
    
    print('sCMD= ' + sCMD)
    os.system(sCMD)
   print('-------------------------------------\n')
    
#====================================================================
#====================================================================  
# EXECUTE PROGRAMS IN FLOWCHART
#====================================================================
#====================================================================   
def Execute_Selection():  
    ExecuteProgram(''): 
#====================================================================
def Open_Explorer():  
   ExecuteProgram(''): 
    
#====================================================================
def Open_Notepad_PP():  
   ExecuteProgram(''): 
    
#====================================================================
def Open_SQLite():  
    ExecuteProgram(''): 
    
#====================================================================
def Open_PythonIDE():  
   ExecuteProgram(''): 
    
#====================================================================
def _Execute():  
    import os
    #os.system('"C:/Windows/System32/notepad.exe"')
    os.system('"C:\PortableApps\SQLiteStudio\SQLiteStudio.exe C:/Users/sesa237770/Documents/Projects_DIY/2024-ROMS-DB/ROM_DB_PY/DB/RetroRoms_20240301.db"')
    
def TEST_Macro():
    #_Ranges()
    #_Execute()    
    _getVariables()
#====================================================================
g_exportedScripts = (TEST_Macro,Execute_Selection,Open_Explorer,Open_Notepad_PP,Open_SQLite,Open_PythonIDE)
#====================================================================

# oShapes = model.getCurrentSelection() 
# sSelectionType= oShapes.ElementType.typeName 
#if sSelectionType='# oShapes.ElementType.typeName'
# oShape = oShapes.getByIndex(0)
# sCMD =oShape.String	

# oShapes.ElementType
# <Type instance com.sun.star.drawing.XShape (<Enum instance com.sun.star.uno.TypeClass ('INTERFACE')>)>   
