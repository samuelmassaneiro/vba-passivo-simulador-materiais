Attribute VB_Name = "ServiceManager"
Function CreateEnvelope(ByVal r As Double, ByVal c As Double) As String
Dim c1 As Double
Dim sCharacteristic As String
Dim vValue As Variant
Dim vValues As Variant
Dim sEnv As String
Dim vEnvents As Variant
Dim vEnvent As Variant

    sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:maes=""http://soa.weg.net/eng/rule/maestro"">"
    sEnv = sEnv & " <soapenv:Header/>"
    sEnv = sEnv & "     <soapenv:Body>"
    sEnv = sEnv & "         <maes:MaestroRuleRequest>"
    sEnv = sEnv & "             <maes:User>" & sUser & "</maes:User>"
    sEnv = sEnv & "             <maes:Language>" & sLanguage & "</maes:Language>"
    sEnv = sEnv & "             <maes:ObjectContext>"
    sEnv = sEnv & "                 <ObjectHeader>"
    sEnv = sEnv & "                     <ObjectType>"
    sEnv = sEnv & "                         <Id>" & sObjId & "</Id>"
    sEnv = sEnv & "                     </ObjectType>"
    sEnv = sEnv & "                 <ObjectVariant>"
    For c1 = 2 To c
        sCharacteristic = plMat.Cells(1, c1).Value
        vValues = Split(plMat.Cells(r, c1).Value, ";")
        sEnv = sEnv & "                 <ObjectValue>"
        sEnv = sEnv & "                     <Characteristic>"
        sEnv = sEnv & "                         <Name>" & sCharacteristic & "</Name>"
        sEnv = sEnv & "                     </Characteristic>"
        For Each vValue In vValues
            sEnv = sEnv & "                     <PropertyValue>"
            sEnv = sEnv & "                         <Value>" & CStr(vValue) & "</Value>"
            sEnv = sEnv & "                     </PropertyValue>"
        Next
        sEnv = sEnv & "                 </ObjectValue>"
    Next
    sEnv = sEnv & "                 </ObjectVariant>"
    sEnv = sEnv & "                 </ObjectHeader>"
    sEnv = sEnv & "                 <RuleStatus>" & sStatus & "</RuleStatus>"
    vEnvents = Split(sEvent, ";")
    For Each vEnvent In vEnvents
        sEnv = sEnv & "                 <Event>"
        sEnv = sEnv & "                     <EventType>" & CStr(vEnvent) & "</EventType>"
        sEnv = sEnv & "                 </Event>"
    Next
    sEnv = sEnv & "                 <Node>" & sAlias & "</Node>"
    sEnv = sEnv & "                 <Alias>" & sAlias & "</Alias>"
    sEnv = sEnv & "             </maes:ObjectContext>"
    sEnv = sEnv & "     </maes:MaestroRuleRequest>"
    sEnv = sEnv & " </soapenv:Body>"
    sEnv = sEnv & "</soapenv:Envelope>"
    CreateEnvelope = sEnv
End Function

Function ExecuteService(ByVal sEnv As String) As MSXML2.XMLHTTP60
Dim xmlRequest As New MSXML2.XMLHTTP60
    'Executa a requisição
    xmlRequest.Open "POST", sURL, False
    xmlRequest.setRequestHeader "Content-Type", "text/xml;charset=UTF-8"
    xmlRequest.setRequestHeader "SOAPAction", "http://sap.com/xi/WebService/soap1.1"
    xmlRequest.setRequestHeader "Operation", "execute"
    If bDebugger Then xmlRequest.setRequestHeader "Cookie", sCookie
    xmlRequest.send sEnv
    Set ExecuteService = xmlRequest
    Set xmlRequest = Nothing
End Function
