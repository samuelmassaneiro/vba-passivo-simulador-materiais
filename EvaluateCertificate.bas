Attribute VB_Name = "EvaluateCertificate"
Sub avalia_certificacoes()

On Error GoTo OnError

'Verifica preenchimento Saída
Sheets("Saida").Select
Range("A1").Select
If (Selection = "") Then
OnError:
    MsgBox "Saída Mal Configurada"
    Application.ScreenUpdating = True
    Exit Sub
End If

'Dim Ranges
Dim certificacao_antiga As String
Dim certificacao_nova As String

'Dim Arrays
Dim certificacoes_novas() As String
Dim certificacoes_antigas() As String

'Dim length
Dim lenght As Integer

'Select CertificateArrays
Dim ArrayOldCertificates As Range
Dim ArrayNewCertificates As Range

'Dim Colunas
Dim column1 As String
Dim column2 As String
Dim columnNewCertificates
Dim columnOldCertificates

Worksheets("Saida").Select

'Inputbox Letra Coluna
column1 = InputBox("Digite qual a Coluna da Certificação Nova", "Coluna Certificação Nova", "ex: 'Y'")
If column1 = "" Then
    MsgBox "Avaliação Cancelada"
    Application.ScreenUpdating = True
    Exit Sub
End If
column2 = InputBox("Digite qual a Coluna da Certificação Antiga", "Coluna Certificação Antiga", "ex: 'Y'")
If column2 = "" Then
    MsgBox "Avaliação Cancelada"
    Application.ScreenUpdating = True
    Exit Sub
End If

'========================COLUNA CERTIFICACAO ANTIGA=======================

       columnOldCertificates = column2 & "2"
       Range(columnOldCertificates).Offset(0, 1).Select
       Set ArrayOldCertificates = Range(Selection, Selection.End(xlDown))
       'MsgBox columnOldCertificates
       
'========================COLUNA CERTIFICACAO NOVA==========================

       columnNewCertificates = column1 & "2"
       Range(columnNewCertificates).Select
       Set ArrayNewCertificates = Range(Selection, Selection.End(xlDown))
       'MsgBox columnNewCertificates
       
'==========================================================================

'Executa DE/PARA
CopyAndReplaceVcCertificates

'Tamanho do Passivo
passiveLength = ArrayOldCertificates.Count
'MsgBox passiveLength

ArrayNewCertificates.Interior.Color = xlNone

For x = 1 To passiveLength - 1

    certificacao_antiga = ArrayOldCertificates.Cells(x).Value
    certificacao_nova = ArrayNewCertificates.Cells(x).Value
    'MsgBox certificacao_antiga

    certificacoes_novas = Split(certificacao_nova, ";", -1)
    certificacoes_antigas = Split(certificacao_antiga, "/", -1)
    
    'MsgBox UBound(certificacoes_novas)

    lenght = UBound(certificacoes_antigas)
    
    For y = 0 To lenght
        
        If Not IsInArray(certificacoes_antigas(y), certificacoes_novas) Then
            ArrayNewCertificates.Cells(x).Activate
            ActiveCell.Interior.ColorIndex = 3
        End If
        
    Next
Next

'Range("AY2").EntireColumn.Delete

End Sub
Function IsInArray(var As String, arrayy As Variant) As Boolean

  IsInArray = (UBound(Filter(arrayy, var)) > -1)
  
End Function
Sub CopyAndReplaceVcCertificates()
    
    Columns("AY").Insert
    Range("AX2").Select
    Range(Selection, Selection.End(xlDown)).Copy
    Range("AY2").Select
    Range(Selection, Selection.End(xlDown)).PasteSpecial

    Dim deCertificacao, deCertificacao1, paraCertificacao, paraCertificacao1
    Dim arrayLength As Integer
    
    '******************************DE/PARA*1***********************************'
    
    deCertificacao = Array( _
        "CE/EAC/CEL", _
        "CE/UL", _
        "CERTIFICACAO UL", _
        "CSA CLASS/UL CLASS", _
        "CSA CLASS/UL CLASS/EAC SEGURA", _
        "CSA CLASS/UL SEG/EAC/CE/SABS", _
        "CSA", _
        "CSA/UL - SEM AREA CLASSIFICADA", _
        "CSA/UL", _
        "UL", _
        "UL/CSA/CE", _
        "UL/FIRE PUMP", _
        "CSA CL/UL SEG/ATEX/IECEX/EAC", _
        "ATEX/IECEX/EAC", _
        "CSA SEG/UL SEG/EAC/CE/MASC", _
        "BUREAU VERITAS", _
        "CSA SEG/UL SEG/INMETRO/CE", _
        "CSA CL/UL SEG/ATEX/IECEX/CE", _
        "CSA SEG/UL SEG/CE/EAC/INMETRO", _
        "CSA SEG/UL SEG/EAC/CE/NOM-ANCE", _
        "CE/IRAM/UL", _
        "UL SEGURA/Ex ec", _
        "CSA SEG/UL SEG/IRAM/RETIE", _
        "CSA SEG/UL SEG/CE/IRAM/RETIE")
    
    paraCertificacao = Array( _
        "CE/EAC", _
        "UL SEGURA/CE", _
        "UL SEGURA", _
        "CSA CLASS/UL", _
        "CSA CLASS/UL/EAC", _
        "CSA SEGURA/UL SEGURA/CE/SABS/EAC", _
        "CSA CLASS", _
        "CSA SEGURA/UL SEGURA", _
        "CSA CLASS/UL SEGURA", _
        "UL SEGURA", _
        "CSA CLASS/UL SEGURA/CE", _
        "UL FIRE PUMP", _
        "CSA CLASS/UL SEGURA/ATEX/IECEX/EAC", _
        "ATEX/IECEX/EAC", _
        "CSA SEGURA/UL SEGURA/EAC/CE/MASC", _
        "BV", _
        "CSA SEGURA/UL SEGURA/INMETRO/CE", _
        "CSA CLASS/UL SEGURA/CE/ATEX/IECEX", _
        "CSA SEGURA/UL SEGURA/CE/EAC/INMETRO", _
        "CSA SEGURA/UL SEGURA/CE/NOM-ANCE/EAC", _
        "IRAM/CE/UL SEGURA", _
        "UL SEGURA", _
        "CSA SEGURA/UL SEGURA/IRAM/RETIE", _
        "CSA SEGURA/UL SEGURA/CE/IRAM/RETIE")
        
    '****************************FIM DE/PARA*1*********************************'
        
    '******************************DE/PARA*2***********************************'
    
    deCertificacao1 = Array( _
        "CSA CL/UL SE/ATEX/IECEX/EAC/CE", _
        "CSA CLASS/UL SEG/CE/UKCA", _
        "CSA CLASS/UL SEG/CE/EAC/UKCA", _
        "CSA SEG/UL SEG/CE/UKCA", _
        "CSA SEG/UL SEG/CE/EAC/UKCA", _
        "CSA SEG/UL SEG/CE/MASC/UKCA", _
        "CSA SE/UL SE/CE/EAC/MASC/UKCA", _
        "CSA SE/UL SE/CE/EAC/Ex ec/UKCA", _
        "CSA SE/UL SE/CE/EAC/INMET/UKCA", _
        "CSA SEG/UL SEG/CE/EAC/UA/UKCA", _
        "CSA SEG/UL FIRE PUMP/CE/UKCA", _
        "UL SEG/CE/UKCA", _
        "UL SEG/CE/EAC/UKCA", _
        "UL SEG/CE/EAC/NOM-ANCE/UKCA", _
        "IECEX/EAC Ex/CCC Ex")
    
    paraCertificacao1 = Array( _
        "CSA CLASS/UL SEGURA/ATEX/IECEX/EAC/CE", _
        "CE/UKCA/UL SEGURA/CSA CLASS", _
        "CE/UKCA/UL SEGURA/CSA CLASS/EAC", _
        "CE/UKCA/UL SEGURA/CSA SEGURA", _
        "CE/UKCA/UL SEGURA/CSA SEGURA/EAC", _
        "CE/UKCA/UL SEGURA/CSA SEGURA/MASC", _
        "CE/UKCA/UL SEGURA/CSA SEGURA/MASC/EAC", _
        "CE/UKCA/UL SEGURA/CSA SEGURA/EAC/Ex ex", _
        "CE/UKCA/UL SEGURA/CSA SEGURA/EAC/INMETRO", _
        "CE/UKCA/UL SEGURA/CSA SEGURA/EAC/UA", _
        "CE/UKCA/CSA SEGURA/UL FIRE PUMP", _
        "CE/UKCA/UL SEGURA", _
        "CE/UKCA/UL SEGURA/EAC", _
        "CE/UKCA/UL SEGURA/EAC/NON-ANCE", _
        "IECEX/EAC Ex/CCC ex")
        
    '****************************FIM DE/PARA*2*********************************'
    
    For x = 0 To UBound(deCertificacao)
        'DE/PARA
        Range("AY2").Select
        Range(Selection, Selection.End(xlDown)).Replace What:=deCertificacao(x), Replacement:=paraCertificacao(x), LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=True
    Next
    
    For x = 0 To UBound(deCertificacao1)
        'DE/PARA
        Range("AY2").Select
        Range(Selection, Selection.End(xlDown)).Replace What:=deCertificacao1(x), Replacement:=paraCertificacao1(x), LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=True
    Next
    
End Sub
