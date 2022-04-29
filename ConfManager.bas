Attribute VB_Name = "ConfManager"
'Define a carrega variáveis globais
Public sInPath As String
Public sOutPath As String
Public sURL As String
Public sUser As String
Public sLanguage As String
Public sObjId As String
Public sStatus As String
Public sEvent As String
Public sAlias As String
Public bDebugger As Boolean
Public sCookie As String
Public bAutoLog As Boolean
Public bAutoService As Boolean
Public plConf As Worksheet
Public plMat As Worksheet
Public plOut As Worksheet
Public plOutLog As Worksheet
Public dplMatGrouperColumn As Double
Public dplMatMaterialColumn As Double
Public dplMatPlantColumn As Double
Public rNode As Range
Public rChar As Range
Public rName As Range
Public rReturnType As Range
Public dI As Date
Public sLogFile As String


Sub LoadConfiguration()
    Set plOut = ThisWorkbook.Worksheets(3)
    Set plOutLog = ThisWorkbook.Worksheets(4)
    Set plConf = ThisWorkbook.Worksheets(1)
    Set plMat = ThisWorkbook.Worksheets(2)
    Set plOut = ThisWorkbook.Worksheets(3)
    sURL = plConf.Range("URL").Value
    sUser = plConf.Range("USER").Value
    sLanguage = plConf.Range("LANGUAGE").Value
    sObjId = plConf.Range("OBJ_TYPE_ID").Value
    sStatus = plConf.Range("RULE_STATUS").Value
    sEvent = plConf.Range("EVENT").Value
    sCookie = plConf.Range("COOKIE").Value
    sInPath = plConf.Range("IN_PATH").Value
    sOutPath = plConf.Range("OUT_PATH").Value
    sAlias = plConf.Range("ALIAS").Value
    Set rNode = plConf.Range("NODE")
    Set rChar = plConf.Range("CHAR")
    Set rName = plConf.Range("NOME")
    Set rReturnType = plConf.Range("RETURN")
    bDebugger = False
    bAutoLog = False
    bAutoService = False
    If plConf.Range("DEBUGGER").Value <> "Não" Then bDebugger = True
    If plConf.Range("AUTOLOG").Value <> "Não" Then bAutoLog = True
    If plConf.Range("AUTOSERVICE").Value <> "Não" Then bAutoService = True
    dplMatMaterialColumn = 1
    dplMatPlantColumn = 2
    dplMatGrouperColumn = 3
    sLogFile = plConf.Range("LOG_FILE").Value
End Sub
