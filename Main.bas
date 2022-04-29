Attribute VB_Name = "Main"
Public dLogColumn As Double

Sub MainData()
    ConfManager.LoadConfiguration
    InitializeData
    FileManager.LoadDataFromDirectoryFiles
    If bAutoService Then MainService
End Sub
Sub InitializeData()
    plMat.Cells.Delete
    plMat.Cells.NumberFormat = "@"
    plMat.Cells(1, dplMatMaterialColumn).Value = "Material"
    plMat.Cells(1, dplMatPlantColumn).Value = "Centro"
    plMat.Cells(1, dplMatGrouperColumn).Value = "Agrupador"
End Sub

Sub MainService()
    ConfManager.LoadConfiguration
    InitializeService
    MaterialManager.LoadNodes
    If sLogFile <> "" Then
        FileManager.EscreverTxt "#####################  Início:" & Now & Chr(9) & Chr(9) & Chr(9) & "#####################"
        FileManager.EscreverTxt "#####################  Usuário:" & sUser & Chr(9) & Chr(9) & Chr(9) & "#####################"
        FileManager.EscreverTxt "#####################  Tipo de objeto:" & sObjId & Chr(9) & Chr(9) & Chr(9) & "#####################"
        FileManager.EscreverTxt "#####################  Tipo de objeto:" & sAlias & Chr(9) & Chr(9) & Chr(9) & "#####################"
    End If
    MaterialManager.LoadMaterial
    If bAutoLog Then Main.MainLog
    If sLogFile <> "" Then FileManager.EscreverTxt "#####################  Fim:" & Now & " #####################"
End Sub

Sub InitializeService()
    plOut.Cells.Delete
    plOut.Cells.NumberFormat = "@"
    plOut.Cells(1, 1).Value = "Material"
    plOut.Cells(1, 2).Value = "Status"
End Sub

Sub MainLog()
    ConfManager.LoadConfiguration
    InitializeLog
    LogManager.CreateLogOutput
End Sub

Sub InitializeLog()
On Error GoTo ErrorHandler
    dLogColumn = WorksheetFunction.Match("log", plOut.Rows(1), 0)
    plOutLog.Cells.Delete
    plOutLog.Cells.NumberFormat = "@"
    plOutLog.Cells(1, 1).Value = "Material"
    plOutLog.Cells(1, 2).Value = "Grupo"
    plOutLog.Cells(1, 3).Value = "Regra"
    plOutLog.Cells(1, 4).Value = "Valor material"
    plOutLog.Cells(1, 5).Value = "Valor esperado"
    Exit Sub
ErrorHandler:
    MsgBox "Não possui a coluna 'LOG' na aba 'Saída'"
    End
End Sub

