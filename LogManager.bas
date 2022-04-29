Attribute VB_Name = "LogManager"
Sub CreateLogOutput()
Dim dR1 As Double, dR2 As Double, r As Double, c As Double
Dim sMaterial As String
Dim sLog As String
Dim vLogs As Variant
Dim vLogColumns As Variant
Dim vLog As Variant
Dim vLogColumn As Variant

Dim clog As Long, c2 As Long
    dR1 = plOut.Cells(plOut.Cells.Rows.Count, 1).End(xlUp).Row
    r = 1
    c = 1
    For dR2 = 2 To dR1
        sMaterial = plOut.Cells(dR2, 1).Value
        sLog = plOut.Cells(dR2, dLogColumn).Value
        vLogs = Split(sLog, ";")
        clog = plOut.Cells(1, plOut.Cells.Columns.Count).End(xlToLeft).Column
        For Each vLog In vLogs
            r = r + 1
            vLogColumns = Split(CStr(vLog), "#")
            plOutLog.Cells(r, c).Value = sMaterial
            For Each vLogColumn In vLogColumns
                c = c + 1
                plOutLog.Cells(r, c).Value = CStr(vLogColumn)
            Next
            c2 = 1
            For c2 = 1 To clog
                plOutLog.Cells(r, c2 + c) = plOut.Cells(dR2, c2)
            Next
            c = 1
        Next
    Next
End Sub
