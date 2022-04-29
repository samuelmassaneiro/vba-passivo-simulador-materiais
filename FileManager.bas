Attribute VB_Name = "FileManager"
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function LoadDataFromDirectoryFiles()
Dim FSO As New FileSystemObject
Dim cFiles As Collection
Dim vFile As Variant
Dim fFile As File
Dim regEx As New RegExp
On Error GoTo ErrorH
    Set cFiles = ListDirectoryOutPutFiles
    For Each vFile In cFiles
        ReadTxtFile (CStr(vFile))
        If FSO.FolderExists(sOutPath) Then
            FSO.MoveFile CStr(vFile), FSO.GetFolder(sOutPath).Path & "\"
        Else
            FSO.CreateFolder sOutPath
            FSO.MoveFile CStr(vFile), FSO.GetFolder(sOutPath).Path & "\"
        End If
    Next
    Exit Function
ErrorH:
    Set fFile = FSO.GetFile(vFile)
    regEx.Global = True
    regEx.MultiLine = True
    regEx.IgnoreCase = True
    regEx.Pattern = "\D"
    fFile.Name = FSO.GetBaseName(fFile.Name) & "_Reprocessed_" & regEx.Replace(Now, "") & "." & FSO.GetExtensionName(fFile.Name)
    FSO.MoveFile CStr(fFile.Path), FSO.GetFolder(sOutPath).Path & "\"
End Function

Private Function ListDirectoryOutPutFiles() As Collection
Dim FSO As New FileSystemObject
Dim fDirectory As Folder
Dim fArchive As File
Dim sPath As String
Dim cFiles As New Collection
    sPath = sInPath
    If FSO.FolderExists(sPath) Then
        Set fDirectory = FSO.GetFolder(sPath)
        For Each fArchive In fDirectory.Files
            If UCase(FSO.GetExtensionName(fArchive.Path)) = "TXT" Then cFiles.Add (fArchive.Path)
        Next
    Else
        MsgBox "Pasta de origem não encontrada"
        End
    End If
    Set ListDirectoryOutPutFiles = cFiles
End Function

Private Sub ReadTxtFile(ByVal sFile As String)
Dim FSO As New FileSystemObject
Dim txt As TextStream
        Set txt = FSO.OpenTextFile(sFile)
        Do While Not txt.AtEndOfStream
            WriteLineAtPlMat (txt.ReadLine)
        Loop
End Sub

Private Sub WriteLineAtPlMat(ByVal sLine As String)
Dim c As Double, r As Double
Dim vColumns As Variant
Dim sMaterial As String
Dim sPlant As String
Dim sGrouper As String
Dim sCharacteristic As String
Dim sValue As String
    vColumns = Split(sLine, "|")
    If UBound(vColumns) > 1 Then
        If IsNumeric(CVar(vColumns(1))) Then
            sMaterial = Trim(CStr(vColumns(2)))
            sPlant = Trim(CStr(vColumns(3)))
            sGrouper = Trim(CStr(vColumns(4)))
            sCharacteristic = Trim(CStr(vColumns(6)))
            sValue = GetCharacteristicValue(Trim(CStr(vColumns(7))), Trim(CStr(vColumns(8))))
            r = GetGrouperRow(sMaterial & ";" & sGrouper & ";" & sPlant)
            c = GetCharacteristicColumn(sCharacteristic)
            plMat.Cells(r, dplMatMaterialColumn).Value = sMaterial & ";" & sGrouper & ";" & sPlant
            If plMat.Cells(r, c).Value = "" Then
                plMat.Cells(r, c).Value = sValue
            Else
                If InStr(1, plMat.Cells(r, c).Value, sValue) = 0 Then
                    plMat.Cells(r, c).Value = plMat.Cells(r, c).Value & ";" & sValue
                End If
            End If
        End If
    End If
End Sub

Private Function GetCharacteristicValue(ByVal sValue As String, ByVal sTyped As String) As String
    If InStr(1, sValue, "ZADI") > 0 Then
        GetCharacteristicValue = sTyped
    Else
        GetCharacteristicValue = sValue
    End If
End Function

Private Function GetGrouperRow(ByVal sGrouper As String) As Double
On Error GoTo ErrorH
    GetGrouperRow = WorksheetFunction.Match(sGrouper, plMat.Columns(1), 0)
    Exit Function
ErrorH:
    GetGrouperRow = plMat.Cells(plMat.Cells.Rows.Count, 1).End(xlUp).Row + 1
End Function

Private Function GetCharacteristicColumn(ByVal sCharacteristic As String) As Double
On Error GoTo ErrorH
    GetCharacteristicColumn = WorksheetFunction.Match(sCharacteristic, plMat.Rows(1), 0)
    Exit Function
ErrorH:
    GetCharacteristicColumn = plMat.Cells(1, plMat.Cells.Columns.Count).End(xlToLeft).Column + 1
    plMat.Cells(1, GetCharacteristicColumn).Value = sCharacteristic
End Function

Public Function LogFile(ByVal sText As String)
Dim txt As TextStream
Dim FSO As New FileSystemObject
    Set txt = FSO.CreateTextFile("C:\carga\logOut.txt", True)
    txt.Write (sText)
    txt.Close
    Set txt = Nothing
    Set FSO = Nothing
End Function

Public Sub EscreverTxt(ByVal Texto As String)
Dim sArq As String
    On Error GoTo erro
    Open sLogFile For Append As #2 'Abre o arquivo
    Print #2, Texto 'Escreve no arquivo
    Close #2 'Fecha o arquivo
    Exit Sub
erro:
    MsgBox "Diretório inexistente", vbCritical
End Sub



