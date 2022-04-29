Attribute VB_Name = "MaterialManager"
Public dR1 As Double
Public dC1 As Double
Public cNodes As Collection
Function printNodes(xmlNodes As IXMLDOMNodeList)
Dim xmlCharNode As IXMLDOMNode
    Debug.Print ("* ")
    For Each xmlCharNode In xmlNodes
        Debug.Print (xmlCharNode.SelectSingleNode("Node").Text)
    Next
End Function
Sub LoadMaterial()
Dim r As Double
Dim sEnv As String
Dim xmlRequest As MSXML2.XMLHTTP60
Dim xmlResponse As MSXML2.DOMDocument60
Dim sMaterialStatus As String
Dim xmlNodes As IXMLDOMNodeList
Dim xmlCharac As IXMLDOMNodeList
Dim xmlCharNode As IXMLDOMNode
Dim sResponse As String
Dim nNode As Node
Dim sSelectNode As String
Dim dT As Date
Dim j As Integer
    On Error Resume Next
    ' AQUI ESTÁ O CÓDIGO PARA PARAR E VERIFICAR APENAS DEPOIS DE ALGUM TEMPO
    dR1 = plMat.Cells(plMat.Cells.Rows.Count, 1).End(xlUp).Row
    dC1 = plMat.Cells(1, plMat.Cells.Columns.Count).End(xlToLeft).Column
    For r = 2 To dR1
        dI = Now
        On Error Resume Next
        sEnv = ""
        sEnv = ServiceManager.CreateEnvelope(r, dC1)
        sSelectNode = ""
        sResponse = ""
request:
        Set xmlRequest = ServiceManager.ExecuteService(sEnv)
        If xmlRequest.Status = 200 Then
            Set xmlResponse = xmlRequest.responseXML
            sMaterialStatus = xmlRequest.statusText
            For Each nNode In cNodes
                Set xmlNodes = xmlResponse.SelectNodes("//ObjectContext")
                If (xmlNodes.Item(0).Text = nNode.Node) Then
                    sSelectNode = "//ObjectContext/ObjectHeader/ObjectVariant/ObjectValue[Characteristic/Name[text()]=""" & nNode.Char & """]/PropertyValue"
                    If nNode.ReturnType = "Descrição" Then sSelectNode = sSelectNode & "/Description"
                    Set xmlCharac = xmlResponse.SelectNodes(sSelectNode)
                    sSelectNode = ""
                Else
                    Set xmlNodes = xmlResponse.SelectNodes("//ObjectContext")
                    For Each xmlCharNode In xmlNodes
                        Debug.Print (xmlCharNode.SelectSingleNode("Node").Text)
                        If xmlCharNode.SelectSingleNode("Node").Text = nNode.Node Then
                            sSelectNode = "//ObjectContext[Node[text()]=""" & nNode.Node & """]/ObjectHeader/ObjectVariant/ObjectValue[Characteristic/Name[text()]=""" & nNode.Char & """]/PropertyValue"
                            If nNode.ReturnType = "Descrição" Then sSelectNode = sSelectNode & "/Description"
                            Set xmlCharac = xmlCharNode.SelectNodes(sSelectNode)
                            sSelectNode = ""
                        End If
                    Next
                End If
                For Each xmlCharNode In xmlCharac
                    If Not xmlCharNode.ChildNodes(0) Is Nothing Then sResponse = sResponse & ";" & xmlCharNode.ChildNodes(0).Text
                Next
                plOut.Cells(r, WorksheetFunction.Match(nNode.Name, plOut.Rows(1), 0)).Value = Replace(sResponse, ";", "", 1, 1)
                sResponse = ""
                Set xmlCharac = Nothing
            Next
            Set xmlCharac = xmlResponse.SelectNodes("//RuleAction/Message/Description")
            For Each xmlCharNode In xmlCharac
                If Not xmlCharNode.ChildNodes(0) Is Nothing Then sResponse = sResponse & ";" & xmlCharNode.ChildNodes(0).Text
            Next
            plOut.Cells(r, WorksheetFunction.Match("Mensagem", plOut.Rows(1), 0)).Value = Replace(sResponse, ";", "", 1, 1)
            Set xmlCharac = xmlResponse.SelectNodes("//Error")
            For Each xmlCharNode In xmlCharac
                If Not xmlCharNode.ChildNodes(0) Is Nothing Then sResponse = sResponse & ";" & xmlCharNode.ChildNodes(0).Text
            Next
            plOut.Cells(r, WorksheetFunction.Match("Erro", plOut.Rows(1), 0)).Value = Replace(sResponse, ";", "", 1, 1)
        Else
            'LINHAS EDITADAS
            Application.Wait (Now + TimeValue("0:01:00"))
            GoTo request
            sMaterialStatus = "Erro na requisição: " & xmlRequest.statusText
        End If
        plOut.Cells(r, 1).Value = plMat.Cells(r, 1).Value
        plOut.Cells(r, 2).Value = sMaterialStatus
        sMaterialStatus = ""
        sSelectNode = ""
        sResponse = ""
        Set xmlRequest = Nothing
        Set xmlResponse = Nothing
        If sLogFile <> "" Then
            dT = (Now - dI)
            FileManager.EscreverTxt (r - 1) & "/" & (dR1 - 1) & Chr(9) & "#" & Chr(9) & dT & "#" & Chr(9) & Now
        End If
        j = j + 1
        If j = 100 Then
            ThisWorkbook.Save
            j = 0
        End If
    Next
End Sub

Sub LoadNodes()
Dim ncNode As New NodeCtrl
Dim sNode As String
Dim i As Long, r As Long

    i = plConf.Cells(plConf.Cells.Rows.Count, rNode.Column).End(xlUp).Row
    For r = rNode.Row + 1 To i
        ncNode.Add plConf.Cells(r, rNode.Column).Value, _
                    plConf.Cells(r, rChar.Column).Value, _
                    plConf.Cells(r, rName.Column).Value, _
                    plConf.Cells(r, rReturnType.Column).Value
    Next
    Set cNodes = ncNode.mCol
    r = 3
    For Each v In cNodes
        plOut.Cells(1, r).Value = v.Name
        r = r + 1
    Next
    plOut.Cells(1, r).Value = "Mensagem"
    plOut.Cells(1, r + 1).Value = "Erro"
End Sub


