EXCEL


Option Explicit

'========================
' Macro principal (robusta e corrigida p/ NF)
'========================
Sub ExtrairNaoLidosSaparado()
    Dim olApp As Object, olNamespace As Object
    Dim olFolder As Object, olSubFolder As Object
    Dim items As Object, olMail As Object
    Dim ws As Worksheet
    Dim linhaAtual As Long
    
    Dim corpoRaw As String
    Dim corpoLinhas As String
    Dim corpoCompacto As String

    Dim pedido As String, oc As String
    Dim pedidoValido As Boolean, ocValido As Boolean

    Dim notaFiscal As String
    Dim valorTxt As String, valorNum As Variant
    Dim vencTxt As String, vencData As Variant
    Dim formaPgto As String, comprador As String, fornecedor As String, obs As String
    
    Dim nomesIgnorados As Variant
    nomesIgnorados = Array("Lucas", "Marcelo", "Francisco", "Erivaldo")
    
    On Error GoTo TratarErro
    
    ' --- Outlook
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(6) ' Inbox
    Set olSubFolder = GetFolder(olFolder, "NFE\01.NOTAS")
    If olSubFolder Is Nothing Then
        MsgBox "Pasta 'NFE\01.NOTAS' não encontrada.", vbCritical
        GoTo Finalizar
    End If
    
    ' --- Planilha "Dados"
    Set ws = ObterOuCriarPlanilha("Dados")
    ws.Cells.Clear
    With ws.Range("A1:K1")
        .Value = Array("DATA RECEBIMENTO", "REMETENTE", "PEDIDO", _
                       "Nº NOTA FISCAL", "VALOR", "VENCIMENTO", _
                       "FORMA DE PAGAMENTO", "COMPRADOR", "FORNECEDOR", "OBS", "IARA")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    linhaAtual = 2
    
    ' --- Itens
    Set items = olSubFolder.items
    On Error Resume Next
    items.Sort "[ReceivedTime]", True
    On Error GoTo TratarErro
    
    Dim senderName As String
    
    For Each olMail In items
        If Not olMail Is Nothing Then
            If HasProperty(olMail, "Class") And olMail.Class = 43 Then
                If olMail.UnRead = True And olMail.Attachments.Count > 0 Then
                    
                    senderName = CStr(olMail.senderName)
                    If IsInArray(senderName, nomesIgnorados) Then GoTo MarcarELerProximo
                    
                    corpoRaw = CStr(olMail.Body)
                    If Len(corpoRaw) = 0 Then GoTo MarcarELerProximo
                    
                    ' Normalizações
                    corpoLinhas = NormalizeForLines(corpoRaw)
                    corpoCompacto = RemoveAllWhitespace(corpoRaw)
                                      
                    ' Extrair PEDIDO/OC
                    pedido = ""
                    oc = ""
                    pedidoValido = False
                    ocValido = False
                    
                    pedido = OnlyDigits(RegexGetFirst(corpoLinhas, LabelValuePatternTolerant("PEDIDO")))
                    oc = OnlyDigits(RegexGetFirst(corpoLinhas, LabelValuePatternTolerant("OC")))
                    
                    If Len(pedido) = 0 Then pedido = OnlyDigits(RegexGetFirst(corpoCompacto, LabelValuePatternAfterCompact("PEDIDO")))
                    If Len(oc) = 0 Then oc = OnlyDigits(RegexGetFirst(corpoCompacto, LabelValuePatternAfterCompact("OC")))
                    
                    If EhPedidoOcValido(pedido) Then pedidoValido = True
                    If EhPedidoOcValido(oc) Then ocValido = True
                    
                    ' Se nenhum válido, descarta (mas marca como lido)
                    If Not pedidoValido And Not ocValido Then
MarcarELerProximo:
                        On Error Resume Next
                        olMail.UnRead = False
                        olMail.Save
                        On Error GoTo TratarErro
                        GoTo Proximo
                    End If
                    
                    ' Preenche linha
                    ws.Cells(linhaAtual, 1).Value = olMail.ReceivedTime
                    ws.Cells(linhaAtual, 1).NumberFormat = "dd/mm/yyyy hh:mm"
                    
                    ws.Cells(linhaAtual, 2).Value = senderName
                    
                    If pedidoValido Then
                        ws.Cells(linhaAtual, 3).Value = pedido
                    ElseIf ocValido Then
                        ws.Cells(linhaAtual, 3).Value = oc
                    End If
                    
                    ' Nº NOTA FISCAL
                    notaFiscal = ""
                    notaFiscal = RegexGetFirst( _
                        corpoLinhas, _
                        "(?:N[º°o]?\s*(?:da\s*)?Nota\s*Fiscal|Nota\s*Fiscal|N[º°o]?\s*NF(?:-?e)?|NF-?e)\s*[:\-–—]?\s*(\d[\dA-Za-z\.\-\/]*)" _
                    )
                    If Len(notaFiscal) = 0 Then
                        notaFiscal = RegexGetFirst( _
                            corpoCompacto, _
                            "(?:N[º°o]?NOTAFISCAL|NOTAFISCAL|NF-?E)\s*[:\-–—]*\s*(\d[\dA-Za-z\.\-\/]*)" _
                        )
                    End If
                    If Len(notaFiscal) > 0 Then ws.Cells(linhaAtual, 4).Value = notaFiscal
                    
                    ' Valor
                    valorTxt = RegexGetFirst(corpoLinhas, "Valor\s*[:\-–—]?\s*(?:R\$\s*)?(\d{1,3}(?:\.\d{3})*,\d{2})")
                    If Len(valorTxt) = 0 Then
                        valorTxt = RegexGetFirst(corpoCompacto, "VALOR[:\-–—]*\s*([0-9]{1,3}(?:\.[0-9]{3})*,[0-9]{2})")
                    End If
                    If Len(valorTxt) > 0 Then
                        valorNum = ToNumberBR(valorTxt)
                        If Not IsError(valorNum) And Not IsEmpty(valorNum) Then
                            ws.Cells(linhaAtual, 5).Value = CDbl(valorNum)
                            ws.Cells(linhaAtual, 5).NumberFormat = "$ #.##0,00"
                        End If
                    End If
                    
                    ' Vencimento
                    vencTxt = RegexGetFirst(corpoLinhas, "Vencimento\s*[:\-–—]?\s*([0-3]?\d[\/\-\.\s][01]?\d[\/\-\.\s]\d{2,4})")
                    If Len(vencTxt) = 0 Then
                        vencTxt = RegexGetFirst(corpoCompacto, "VENCIMENTO[:\-–—]*\s*([0-3]?\d[\/\-\.\s][01]?\d[\/\-\.\s]\d{2,4})")
                    End If
                    If Len(vencTxt) > 0 Then
                        vencData = ToDateBR(vencTxt)
                        If IsDate(vencData) Then
                            ws.Cells(linhaAtual, 6).Value = CDate(vencData)
                            ws.Cells(linhaAtual, 6).NumberFormat = "dd/mm/yyyy"
                        End If
                    End If
                    
                    ' Outros campos
                    formaPgto = RegexGetLineValue(corpoLinhas, "Forma\s*de\s*Pagamento")
                    If Len(formaPgto) > 0 Then ws.Cells(linhaAtual, 7).Value = formaPgto
                    
                    comprador = RegexGetLineValue(corpoLinhas, "Comprador")
                    If Len(comprador) > 0 Then ws.Cells(linhaAtual, 8).Value = comprador
                    
                    fornecedor = RegexGetLineValue(corpoLinhas, "Fornecedor")
                    If Len(fornecedor) > 0 Then ws.Cells(linhaAtual, 9).Value = fornecedor
                    
                    obs = RegexGetLineValue(corpoLinhas, "Obs")
                    If Len(obs) > 0 Then ws.Cells(linhaAtual, 10).Value = obs
                    
                    ' Marcar como lido
                    On Error Resume Next
                    olMail.UnRead = False
                    olMail.Save
                    On Error GoTo TratarErro
                    
                    linhaAtual = linhaAtual + 1
                End If
            End If
        End If
Proximo:
    Next olMail

    '=========================================================
    '   BLOCO DO PROCX (COLUNA K - "IARA")
    '=========================================================
    If linhaAtual > 2 Then
        With ws
            ' Cabeçalho já criado como "IARA" em K1; se quiser outra legenda, ajuste aqui:
            .Cells(1, 11).Value = "IARA"
            
            ' Fórmula PROCX em PT-BR (FormulaLocal)
            .Range("K2").FormulaLocal = _
                "=PROCX(" & _
                "C2&D2;" & _
                "'[Planilha de Chamados IARA+ 2025 v8 - 012026.xlsx]Preechimento'!$O:$O&" & _
                "'[Planilha de Chamados IARA+ 2025 v8 - 012026.xlsx]Preechimento'!$S:$S;" & _
                "'[Planilha de Chamados IARA+ 2025 v8 - 012026.xlsx]Preechimento'!$V:$V;" & _
                """AUSENTE""" & _
                ")"
            
            ' Preenche até a última linha de dados (linhaAtual - 1)
            .Range("K2").AutoFill Destination:=.Range("K2:K" & (linhaAtual - 1))
            
            ' Formatação e ajuste de largura
            .Columns("K").NumberFormat = "@" ' deixe texto; remova se a coluna V for numérica
            .Columns("A:K").AutoFit
        End With
    End If
    '=========================================================
    
    ' --- Formatação final
    With ws.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Columns("A:K").AutoFit
    ws.Range("A1:K1").AutoFilter
    
Finalizar:
    On Error Resume Next
    Set olMail = Nothing
    Set items = Nothing
    Set olSubFolder = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    On Error GoTo 0
    Exit Sub

TratarErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
    Resume Finalizar
End Sub

'========================
' Funções auxiliares
'========================

' Obter subpasta por caminho (ex.: "NFE\01.NOTAS")
Function GetFolder(ByVal rootFolder As Object, ByVal folderPath As String) As Object
    Dim arrFolders() As String, folder As Object
    Dim i As Long
    If rootFolder Is Nothing Or Len(folderPath) = 0 Then Exit Function
    arrFolders = Split(folderPath, "\")
    Set folder = rootFolder
    For i = LBound(arrFolders) To UBound(arrFolders)
        On Error Resume Next
        Set folder = folder.Folders(arrFolders(i))
        On Error GoTo 0
        If folder Is Nothing Then Exit For
    Next i
    Set GetFolder = folder
End Function

' Cria ou retorna uma planilha pelo nome
Function ObterOuCriarPlanilha(ByVal nome As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = nome
    End If
    Set ObterOuCriarPlanilha = ws
End Function

' Preserva QUEBRAS de linha, mas normaliza espaços/tabs/NBSP e colapsa múltiplos espaços
Function NormalizeForLines(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    t = Replace(t, vbTab, " ")
    t = Replace(t, Chr$(160), " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormalizeForLines = t
End Function

' Remove TODOS os espaços/tabs/quebras
Function RemoveAllWhitespace(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, vbCrLf, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")
    t = Replace(t, Chr$(160), "")
    t = Replace(t, " ", "")
    RemoveAllWhitespace = t
End Function

' Retorna o primeiro grupo capturado (ou o match inteiro)
Function RegexGetFirst(ByVal texto As String, ByVal pattern As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    re.Multiline = True
    If re.Test(texto) Then
        Set m = re.Execute(texto)(0)
        If m.SubMatches.Count > 0 Then
            RegexGetFirst = Trim$(m.SubMatches(0))
        Else
            RegexGetFirst = Trim$(m.Value)
        End If
    Else
        RegexGetFirst = ""
    End If
End Function

' Padrão tolerante do tipo: LABEL : valor
Function LabelValuePatternTolerant(ByVal label As String) As String
    LabelValuePatternTolerant = label & "\s*[:\-–—]?\s*([^\r\n]+)"
End Function

' Para o corpo compacto (sem espaços): LABEL seguido de pontuação opcional e dígitos/sinais
Function LabelValuePatternAfterCompact(ByVal label As String) As String
    LabelValuePatternAfterCompact = label & "[:\-–—]*\s*([0-9\.\-\/]+)"
End Function

' Lê um valor até o fim da linha após uma label
Function RegexGetLineValue(ByVal texto As String, ByVal label As String) As String
    RegexGetLineValue = RegexGetFirst(texto, LabelValuePatternTolerant(label))
End Function

' Mantém apenas dígitos
Function OnlyDigits(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then r = r & ch
    Next i
    OnlyDigits = r
End Function

' Valida PEDIDO/OC: 10 dígitos começando com "45"
Function EhPedidoOcValido(ByVal codigo As String) As Boolean
    EhPedidoOcValido = (Len(codigo) = 10 And Left$(codigo, 2) = "45")
End Function

' Converte "1.234,56" -> 1234.56 (Double)
Function ToNumberBR(ByVal s As String) As Variant
    Dim t As String
    t = Trim$(s)
    If t = "" Then
        ToNumberBR = Empty
        Exit Function
    End If
    t = Replace(t, ".", ".")
    t = Replace(t, ",", ",")
    If IsNumeric(t) Then
        ToNumberBR = CDbl(t)
    Else
        ToNumberBR = Empty
    End If
End Function

' Converte "dd/mm/aaaa" (ou dd-mm-aaaa, dd.mm.aaaa, com espaço) para Date
Function ToDateBR(ByVal s As String) As Variant
    Dim t As String, d As Variant, dia As Integer, mes As Integer, ano As Integer
    t = Trim$(s)
    If t = "" Then
        ToDateBR = Empty
        Exit Function
    End If
    t = Replace(t, "-", "/")
    t = Replace(t, ".", "/")
    t = Replace(t, " ", "/")
    Do While InStr(t, "//") > 0
        t = Replace(t, "//", "/")
    Loop
    d = Split(t, "/")
    If UBound(d) <> 2 Then
        ToDateBR = Empty
        Exit Function
    End If
    dia = CInt(val(d(0))): mes = CInt(val(d(1))): ano = CInt(val(d(2)))
    If ano < 100 Then ano = 2000 + ano
    If dia >= 1 And dia <= 31 And mes >= 1 And mes <= 12 And ano >= 1900 And ano <= 2100 Then
        ToDateBR = DateSerial(ano, mes, dia)
    Else
        ToDateBR = Empty
    End If
End Function

' Evita erro ao acessar propriedades em itens que não são MailItem
Function HasProperty(ByVal obj As Object, ByVal propName As String) As Boolean
    On Error Resume Next
    Dim tmp As Variant
    tmp = CallByName(obj, propName, VbGet)
    HasProperty = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' Checa se valor está no array (case-insensitive)
Function IsInArray(ByVal val As String, ByVal arr As Variant) As Boolean
    Dim element As Variant
    IsInArray = False
    For Each element In arr
        If StrComp(val, element, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next
End Function












Sub Main_ExtrairTodosNaoLidosNotas()
    Dim OutlookApp As Object, OutlookNamespace As Object, olSubFolder As Object, olMail As Object
    Dim ws As Worksheet, linhaAtual As Long
    Dim conteudoExtraido As String
    Dim nomesIgnorados As Variant
    
    On Error GoTo TratarErro
    
    ' Inicializar Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    Set olSubFolder = GetFolder(OutlookNamespace.GetDefaultFolder(6), "NFE\01.NOTAS")
    If olSubFolder Is Nothing Then MsgBox "Pasta não encontrada!", vbCritical: Exit Sub
    
    ' Configurar planilha
    Set ws = ThisWorkbook.Sheets("Dados")
    ws.Cells.Clear
    ws.Range("A1:E1").Value = Array("REMETENTE", "CONTEÚDO EXTRAÍDO", _
                                    "NÚMERO DE PEDIDO", "DATA DE VENCIMENTO", "DATA DE RECEBIMENTO")
    ws.Range("A1:E1").Font.Bold = True
    linhaAtual = 2
    
    ' Centralizar cabeçalho
    ws.Range("A1:E1").HorizontalAlignment = xlCenter
    
    ' Remetentes ignorados
    nomesIgnorados = Array("Lucas", "Marcelo", "Francisco", "Erivaldo")
    
    ' Processar e-mails não lidos
    Dim olItems As Object: Set olItems = olSubFolder.items
    olItems.Sort "[ReceivedTime]", True
    
    For Each olMail In olItems
        If olMail.Class = 43 And olMail.UnRead Then
            If IsInArray(olMail.senderName, nomesIgnorados) Then GoTo ProximoEmail
            
            ' Limpar corpo e consolidar como conteúdo extraído
            conteudoExtraido = LimparTextoSemEspacos(olMail.Body)
            
            ' Preencher planilha
            ws.Cells(linhaAtual, 1).Value = olMail.senderName
            ws.Cells(linhaAtual, 2).Value = conteudoExtraido
            
            ' Extrair pedido no padrão correto (45 + 8 dígitos)
            Dim pedido As String
            pedido = RegexExtrairPrimeiro(conteudoExtraido, "45\d{8}")
            If pedido = "" Then pedido = "A DEFINIR"
            ws.Cells(linhaAtual, 3).Value = pedido
            
            ' Extrair primeira data e formatar como data
            Dim vencimento As String
            vencimento = RegexExtrairPrimeiro(conteudoExtraido, "\d{2}[/-]\d{2}[/-]\d{2,4}")
            If vencimento = "" Then
                ws.Cells(linhaAtual, 4).Value = "A DEFINIR"
            Else
                On Error Resume Next
                ws.Cells(linhaAtual, 4).Value = CDate(Replace(Replace(vencimento, "-", "/"), ".", "/"))
                ws.Cells(linhaAtual, 4).NumberFormat = "dd/mm/yyyy"
                On Error GoTo 0
            End If
            
            ' Adicionar data de recebimento do e-mail
            ws.Cells(linhaAtual, 5).Value = olMail.ReceivedTime
            ws.Cells(linhaAtual, 5).NumberFormat = "dd/mm/yyyy hh:mm"
            
            ' Centralizar linha atual
            ws.Range(ws.Cells(linhaAtual, 1), ws.Cells(linhaAtual, 5)).HorizontalAlignment = xlCenter
            
            linhaAtual = linhaAtual + 1
                End If
                
ProximoEmail:
    Next
    
    ws.UsedRange.Columns.AutoFit
    Exit Sub
    
TratarErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' Funções auxiliares
Function GetFolder(rootFolder As Object, folderPath As String) As Object
    Dim arrFolders() As String, folder As Object, i As Integer
    arrFolders = Split(folderPath, "\")
    Set folder = rootFolder
    For i = LBound(arrFolders) To UBound(arrFolders)
        Set folder = folder.Folders(arrFolders(i))
        If folder Is Nothing Then Exit For
    Next
    Set GetFolder = folder
End Function

Function LimparTextoSemEspacos(texto As String) As String
    Dim t As String
    t = Replace(texto, vbCrLf, "")
    t = Replace(t, vbTab, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, " ", "")
    LimparTextoSemEspacos = t
End Function

Function RegexExtrairPrimeiro(texto As String, padrao As String) As String
    Dim regex As Object, matches As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = padrao
    regex.Global = True
    If regex.Test(texto) Then
        Set matches = regex.Execute(texto)
        RegexExtrairPrimeiro = matches(0).Value
    Else
        RegexExtrairPrimeiro = ""
    End If
End Function

Function IsInArray(val As String, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If StrComp(val, element, vbTextCompare) = 0 Then IsInArray = True: Exit Function
    Next
End Function








'obs:
'o que n for alta prioridade ou n for boleto e estiver como com vencimento 11/2 ou próximo jogar para 25/2 - fpp
'padrão sap "100*****"
'padrão cc "2025'41'ou'42'****"

Option Explicit

' ==========================================================
' MÓDULO PRINCIPAL
' ==========================================================
Sub Marcacomolidos()

    Dim OutlookApp As Object
    Dim OutlookNamespace As Object
    Dim olSubFolder As Object
    Dim olMail As Object
    Dim olItems As Object
    
    Dim ws As Worksheet
    Dim linhaAtual As Long
    Dim conteudoExtraido As String
    Dim nomesIgnorados As Variant
    
    On Error GoTo TratarErro

    ' --- Inicializar Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    ' 6 = olFolderInbox. Ajuste o caminho abaixo conforme sua árvore no Outlook:
    Set olSubFolder = GetFolder(OutlookNamespace.GetDefaultFolder(6), "NFE\01.NOTAS")
    If olSubFolder Is Nothing Then
        MsgBox "Pasta não encontrada! Verifique o caminho NFE\01.NOTAS.", vbCritical
        Exit Sub
    End If

    ' --- Configurar planilha
    Set ws = ThisWorkbook.Sheets("Dados1")
    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array( _
        "REMETENTE", "CONTEÚDO EXTRAÍDO", _
        "NÚMERO DE PEDIDO", "DATA DE VENCIMENTO", _
        "DATA DE RECEBIMENTO", "Nº NF" _
    )
    With ws.Range("A1:F1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    linhaAtual = 2

    ' --- Remetentes ignorados
    nomesIgnorados = Array("Lucas", "Marcelo", "Francisco", "Erivaldo")

    ' --- Processar e-mails (ordem: mais recentes primeiro)
    Set olItems = olSubFolder.items
    olItems.Sort "[ReceivedTime]", True

    For Each olMail In olItems
        If olMail.Class = 43 Then ' 43 = MailItem
            If olMail.UnRead Then

                ' Pular remetentes ignorados
                If IsInArray(olMail.senderName, nomesIgnorados) Then GoTo ProximoEmail

                ' Limpa texto removendo quebras/espacos (mantém a sua estratégia)
                conteudoExtraido = LimparTextoSemEspacos(olMail.Body)

                ' --- Preenche colunas base
                ws.Cells(linhaAtual, 1).Value = olMail.senderName
                ws.Cells(linhaAtual, 2).Value = conteudoExtraido

                ' --- Nº do Pedido (padrão: 45 + 8 dígitos)
                Dim pedido As String
                pedido = RegexExtrairPrimeiro(conteudoExtraido, "45\d{8}")
                If pedido = "" Then pedido = "A DEFINIR"
                ws.Cells(linhaAtual, 3).Value = pedido

                ' --- Primeira Data encontrada -> Vencimento
                Dim vencimento As String
                vencimento = RegexExtrairPrimeiro(conteudoExtraido, "\d{2}[/-]\d{2}[/-]\d{2,4}")
                If vencimento <> "" Then
                    On Error Resume Next
                    ws.Cells(linhaAtual, 4).Value = CDate(Replace(Replace(vencimento, "-", "/"), ".", "/"))
                    ws.Cells(linhaAtual, 4).NumberFormat = "dd/mm/yyyy"
                    On Error GoTo TratarErro
                Else
                    ws.Cells(linhaAtual, 4).Value = "A DEFINIR"
                End If

                ' --- Data/Hora de recebimento do e-mail
                ws.Cells(linhaAtual, 5).Value = olMail.ReceivedTime
                ws.Cells(linhaAtual, 5).NumberFormat = "dd/mm/yyyy hh:mm"

                ' ======================================================
                ' Nº da NF — CORRIGIDO (evita confundir com nº do pedido)
                ' Observação: como o texto foi "compactado", procuramos tokens
                ' típicos de NF que aparecem SEM espaço: NF, NOTA, NºNOTAFISCAL, NOTAFISCAL
                ' Ex.: "Nº NOTA FISCAL" -> "NºNOTAFISCAL"
                ' ======================================================
                Dim numNFTok As String
                numNFTok = RegexExtrairPrimeiro( _
                    conteudoExtraido, _
                    "(?:Nº?NOTAFISCAL|NOTAFISCAL|NOTA|NF)[:\- ]*(\d+)" _
                )

                Dim numNF As String
                If numNFTok <> "" Then
                    ' O helper retorna a correspondência completa.
                    ' Vamos “limpar” o prefixo e manter somente os dígitos finais.
                    numNF = RegexExtrairPrimeiro(numNFTok, "\d+")
                Else
                    ' Fallback (opcional): tenta achar "Nº:" seguido de número
                    numNF = RegexExtrairPrimeiro(conteudoExtraido, "Nº[:\- ]*(\d+)")
                    If numNF = "" Then numNF = "NÃO ENCONTRADO"
                End If
                ws.Cells(linhaAtual, 6).Value = numNF

                ' --- Valor da NF em formato BR (ex.: 1.234,56 | 11.180,00)
                Dim valorNFMatch As String, valorNFnum As Double
                valorNFMatch = RegexExtrairPrimeiro(conteudoExtraido, "R?\$?\d{1,3}(\.\d{3})*,\d{2}")
               
                Else
                    ws.Cells(linhaAtual, 7).Value = "NÃO ENCONTRADO"
                End If

                ' --- Centralizar linha
                ws.Range(ws.Cells(linhaAtual, 1), ws.Cells(linhaAtual, 7)).HorizontalAlignment = xlCenter

                ' --- Marcar como lido e salvar
                olMail.UnRead = False
                olMail.Save

                linhaAtual = linhaAtual + 1
            End If
        
ProximoEmail:
    Next

    ws.UsedRange.Columns.AutoFit
    Exit Sub

TratarErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical

End Sub


' ==========================================================
' FUNÇÕES AUXILIARES
' ==========================================================

' Navega até uma subpasta a partir de um rootFolder (ex.: Inbox)
Function GetFolder(rootFolder As Object, folderPath As String) As Object
    Dim arrFolders() As String
    Dim folder As Object
    Dim i As Long

    arrFolders = Split(folderPath, "\")
    Set folder = rootFolder

    For i = LBound(arrFolders) To UBound(arrFolders)
        If folder Is Nothing Then Exit For
        Set folder = folder.Folders(arrFolders(i))
    Next i

    Set GetFolder = folder
End Function

' Remove quebras de linha, tabs e espaços (mantém sua estratégia de compactar texto)
Function LimparTextoSemEspacos(texto As String) As String
    Dim t As String
    t = Replace(texto, vbCrLf, "")
    t = Replace(t, vbTab, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, " ", "")
    LimparTextoSemEspacos = t
End Function

' Retorna o PRIMEIRO match da expressão (como string do próprio match)
Function RegexExtrairPrimeiro(texto As String, padrao As String) As String
    Dim regex As Object
    Dim matches As Object

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .pattern = padrao
        .Global = True
        .IgnoreCase = True
    End With

    If regex.Test(texto) Then
        Set matches = regex.Execute(texto)
        RegexExtrairPrimeiro = matches(0).Value
    Else
        RegexExtrairPrimeiro = ""
    End If
End Function

' Checa se um valor existe em um array (case-insensitive)
Function IsInArray(val As String, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If StrComp(val, element, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next
    IsInArray = False
End Function









Sub ExtrairTodos()
    Dim olApp As Object, olNamespace As Object, olFolder As Object, olSubFolder As Object
    Dim olMail As Object
    Dim ws As Worksheet
    Dim i As Long, linhaAtual As Long
    Dim linhas() As String
    Dim pedido As String, oc As String, notaFiscal As String
    Dim pedidoValido As Boolean, ocValido As Boolean
    Dim linhaLimpa As String
    Dim campo As Variant, valor As String
    Dim dados As Variant
    Dim colunas As Variant
    
    On Error GoTo TratarErro
    
    ' Inicializar Outlook
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' Obter pasta "NFE\01.NOTAS"
    Set olFolder = olNamespace.GetDefaultFolder(6) ' Inbox
    Set olSubFolder = GetFolder(olFolder, "NFE\01.NOTAS")
    If olSubFolder Is Nothing Then
        MsgBox "Pasta 'NFE\01.NOTAS' não encontrada!", vbCritical
        Exit Sub
    End If
    
    ' Configurar planilha
    Set ws = ThisWorkbook.Sheets("Dados")
    ws.Cells.Clear
    
    With ws.Range("A1:J1")
        .Value = Array("DATA RECEBIMENTO", "REMETENTE", "PEDIDO", _
                       "Nº NOTA FISCAL", "VALOR", "VENCIMENTO", _
                       "FORMA DE PAGAMENTO", "COMPRADOR", "FORNECEDOR", "OBS")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    linhaAtual = 2
    dados = Array("PEDIDO:", "OC:", "Nº NOTA FISCAL:", "Nº da NOTA FISCAL:", _
                  "VALOR:", "VENCIMENTO:", "FORMA DE PAGAMENTO:", _
                  "COMPRADOR:", "FORNECEDOR:", "OBS:")
    
    ' Loop nos e-mails
    For Each olMail In olSubFolder.items
        If olMail.Class = 43 And olMail.UnRead = False And olMail.Attachments.Count > 0 Then
            linhas = Split(olMail.Body, vbCrLf)
            
          ' Inicializar variáveis
pedido = ""
oc = ""
pedidoValido = False
ocValido = False

' Loop para extrair e validar PEDIDO e OC
For i = LBound(linhas) To UBound(linhas)
    linhaLimpa = Trim(Replace(linhas(i), "?", ""))
    
    If InStr(1, linhaLimpa, "PEDIDO:", vbTextCompare) > 0 Then
        pedido = ExtrairValor(linhaLimpa)
        If Left(pedido, 2) = "45" And Len(pedido) = 10 Then
            pedidoValido = True
        Else
            pedido = ""
            pedidoValido = False
        End If
        
    ElseIf InStr(1, linhaLimpa, "OC:", vbTextCompare) > 0 Then
        oc = ExtrairValor(linhaLimpa)
        If Left(oc, 2) = "45" And Len(oc) = 10 Then
            ocValido = True
        Else
            oc = ""
            ocValido = False
        End If
    End If
Next i

' Inserir na planilha
If pedidoValido Then
    ws.Cells(linhaAtual, 3).Value = pedido
ElseIf ocValido Then
    ws.Cells(linhaAtual, 3).Value = oc
End If
            ' Ignorar e-mails sem pedido ou OC válido
            ' If Not pedidoValido And Not ocValido Then
                ' olMail.UnRead = False
                ' olMail.Save
             '    GoTo ProximoEmail
         '    End If
            
            ' Preencher dados básicos
            With ws.Cells(linhaAtual, 1)
                If IsDate(olMail.ReceivedTime) Then
                    .Value = olMail.ReceivedTime
                    .NumberFormat = "dd/mm/yyyy"
                Else
                    MsgBox "Data inválida no e-mail de " & olMail.senderName, vbExclamation
                    GoTo ProximoEmail
                End If
            End With
            
            ws.Cells(linhaAtual, 2).Value = olMail.senderName
            
            If pedidoValido Then
                ws.Cells(linhaAtual, 3).Value = pedido
            ElseIf ocValido Then
                ws.Cells(linhaAtual, 3).Value = oc
            End If
            
            ' Extrair demais campos
            For i = LBound(linhas) To UBound(linhas)
                linhaLimpa = Trim(Replace(linhas(i), "?", ""))
                
                ' Nota Fiscal
                If InStr(1, linhaLimpa, "Nº NOTA FISCAL:", vbTextCompare) > 0 Or _
                   InStr(1, linhaLimpa, "Nº da NOTA FISCAL:", vbTextCompare) > 0 Then
                    notaFiscal = ExtrairValor(linhaLimpa)
                    ws.Cells(linhaAtual, 4).Value = notaFiscal
                End If
                
                ' Outros campos
                For Each campo In dados
                    If campo <> "PEDIDO:" And campo <> "OC:" And _
                       campo <> "Nº NOTA FISCAL:" And campo <> "Nº da NOTA FISCAL:" Then
                       
                        If InStr(1, linhaLimpa, campo, vbTextCompare) > 0 Then
                            valor = ExtrairValor(linhaLimpa)
                            Select Case campo
                                Case "VALOR:"
                                    ws.Cells(linhaAtual, 5).Value = valor
                                Case "VENCIMENTO:"
                                    ws.Cells(linhaAtual, 6).Value = valor
                                Case "FORMA DE PAGAMENTO:"
                                    ws.Cells(linhaAtual, 7).Value = valor
                                Case "COMPRADOR:"
                                    ws.Cells(linhaAtual, 8).Value = valor
                                Case "FORNECEDOR:"
                                    ws.Cells(linhaAtual, 9).Value = valor
                                Case "OBS:"
                                    ws.Cells(linhaAtual, 10).Value = valor
                            End Select
                        End If
                    End If
                Next campo
            Next i
            
            ' Marcar e-mail como lido e salvar
            ' olMail.UnRead = False
            ' olMail.Save
            
            linhaAtual = linhaAtual + 1
        End If
ProximoEmail:
    Next olMail
    
    ' Formatar planilha
    With ws.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ' Limpar objetos
    Set olMail = Nothing
    Set olSubFolder = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' Função para extrair valor após ":"
Function ExtrairValor(texto As String) As String
    Dim pos As Long
    pos = InStr(texto, ":")
    If pos > 0 Then
        ExtrairValor = Trim(Mid(texto, pos + 1))
    Else
        ExtrairValor = ""
    End If
End Function

' Função para obter pasta via caminho
Function GetFolder(ByVal rootFolder As Object, ByVal folderPath As String) As Object
    Dim arrFolders() As String
    Dim i As Integer
    Dim folder As Object
    
    arrFolders = Split(folderPath, "\")
    Set folder = rootFolder
    
    For i = LBound(arrFolders) To UBound(arrFolders)
        On Error Resume Next
        Set folder = folder.Folders(arrFolders(i))
        On Error GoTo 0
        If folder Is Nothing Then Exit For
    Next i
    
    Set GetFolder = folder
End Function









'========================
' Macro principal
'========================
Sub ExtrairNaoLidosANTIGO()
    Dim olApp As Object, olNamespace As Object
    Dim olFolder As Object, olSubFolder As Object
    Dim olMail As Object, items As Object
    Dim ws As Worksheet
    Dim i As Long, linhaAtual As Long
    Dim linhas() As String
    Dim pedido As String, oc As String, notaFiscal As String
    Dim pedidoValido As Boolean, ocValido As Boolean
    Dim linhaLimpa As String
    Dim campo As Variant, valor As String
    Dim corpo As String
    Dim dictMap As Object ' Scripting.Dictionary
    Dim chave As Variant
    
    On Error GoTo TratarErro

    ' --- Inicializar Outlook (late binding)
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")

    ' --- Obter pasta: Caixa de Entrada -> NFE\01.NOTAS
    Set olFolder = olNamespace.GetDefaultFolder(6) ' 6 = olFolderInbox
    Set olSubFolder = GetFolder(olFolder, "NFE\01.NOTAS")
    If olSubFolder Is Nothing Then
        MsgBox "Pasta 'NFE\01.NOTAS' não encontrada!", vbCritical
        GoTo Finalizar
    End If

    ' --- Preparar planilha "Dados"
    Set ws = ObterOuCriarPlanilha("Dados")
    ws.Cells.Clear

    With ws.Range("A1:J1")
        .Value = Array("DATA RECEBIMENTO", "REMETENTE", "PEDIDO", _
                       "Nº NOTA FISCAL", "VALOR", "VENCIMENTO", _
                       "FORMA DE PAGAMENTO", "COMPRADOR", "FORNECEDOR", "OBS")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' --- Mapa campo -> coluna
    Set dictMap = CreateObject("Scripting.Dictionary")
    dictMap.CompareMode = 1 ' TextCompare
    dictMap.Add "VALOR:", 5
    dictMap.Add "VENCIMENTO:", 6
    dictMap.Add "FORMA DE PAGAMENTO:", 7
    dictMap.Add "COMPRADOR:", 8
    dictMap.Add "FORNECEDOR:", 9
    dictMap.Add "OBS:", 10
    ' NOTA FISCAL é tratada à parte (coluna 4) pois possui 2 variações de rótulo

    linhaAtual = 2

    ' --- Iterar itens da pasta
    Set items = olSubFolder.items
    ' Opcional: ordenar por data de recebimento
    On Error Resume Next
    items.Sort "[ReceivedTime]", True
    On Error GoTo TratarErro

    For Each olMail In items
        ' Somente e-mails (MailItem = Class 43)
        If Not olMail Is Nothing Then
            If HasProperty(olMail, "Class") Then
                If olMail.Class = 43 Then
                    ' Apenas NÃO lidos com anexos
                    If olMail.UnRead = True And olMail.Attachments.Count > 0 Then

                        ' Corpo e linhas
                        corpo = olMail.Body
                        If Len(corpo) = 0 Then GoTo MarcarELimpar

                        ' Normalizar quebras para CRLF e dividir
                        corpo = Replace(corpo, vbCr, vbLf)
                        corpo = Replace(corpo, vbLf & vbLf, vbLf)
                        linhas = Split(corpo, vbLf)

                        ' Inicializar variáveis do e-mail
                        pedido = ""
                        oc = ""
                        notaFiscal = ""
                        pedidoValido = False
                        ocValido = False

                        ' --- 1) Descobrir PEDIDO/OC e validar
                        For i = LBound(linhas) To UBound(linhas)
                            linhaLimpa = NormalizarLinha(linhas(i))

                            If InStr(1, linhaLimpa, "PEDIDO:", vbTextCompare) > 0 Then
                                valor = ExtrairAposDoisPontos(linhaLimpa)
                                valor = ExtrairNumeros(valor)
                                If EhPedidoOcValido(valor) Then
                                    pedido = valor
                                    pedidoValido = True
                                End If

                            ElseIf InStr(1, linhaLimpa, "OC:", vbTextCompare) > 0 Then
                                valor = ExtrairAposDoisPontos(linhaLimpa)
                                valor = ExtrairNumeros(valor)
                                If EhPedidoOcValido(valor) Then
                                    oc = valor
                                    ocValido = True
                                End If
                            End If
                        Next i

                        ' Se não achou nem PEDIDO nem OC válidos, ignorar (marcar como lido e seguir)
                        If Not pedidoValido And Not ocValido Then
MarcarELimpar:
                            On Error Resume Next
                            olMail.UnRead = False
                            olMail.Save
                            On Error GoTo TratarErro
                            GoTo Proximo
                        End If

                        ' --- 2) Preencher linha
                        ' A) Data recebimento
                        ws.Cells(linhaAtual, 1).Value = olMail.ReceivedTime
                        ws.Cells(linhaAtual, 1).NumberFormat = "dd/mm/yyyy hh:mm"

                        ' B) Remetente
                        ws.Cells(linhaAtual, 2).Value = olMail.senderName

                        ' C) PEDIDO (prioriza PEDIDO sobre OC)
                        If pedidoValido Then
                            ws.Cells(linhaAtual, 3).Value = pedido
                        ElseIf ocValido Then
                            ws.Cells(linhaAtual, 3).Value = oc
                        End If

                        ' --- 3) Demais campos (inclui NOTA FISCAL e mapeados)
                        For i = LBound(linhas) To UBound(linhas)
                            linhaLimpa = NormalizarLinha(linhas(i))

                            ' Nota Fiscal (duas variações de rótulo)
                            If InStr(1, linhaLimpa, "Nº NOTA FISCAL:", vbTextCompare) > 0 _
                               Or InStr(1, linhaLimpa, "Nº da NOTA FISCAL:", vbTextCompare) > 0 Then
                                notaFiscal = ExtrairAposDoisPontos(linhaLimpa)
                                notaFiscal = Trim(notaFiscal)
                                ws.Cells(linhaAtual, 4).Value = notaFiscal
                            Else
                                ' Campos mapeados
                                For Each chave In dictMap.Keys
                                    If InStr(1, linhaLimpa, CStr(chave), vbTextCompare) > 0 Then
                                        valor = ExtrairAposDoisPontos(linhaLimpa)
                                        valor = Trim(valor)
                                        ws.Cells(linhaAtual, CLng(dictMap(chave))).Value = valor
                                    End If
                                Next chave
                            End If
                        Next i

                        ' Marcar como lido e salvar após processar
                        On Error Resume Next
                        olMail.UnRead = False
                        olMail.Save
                        On Error GoTo TratarErro

                        linhaAtual = linhaAtual + 1
                    End If
                End If
            End If
        End If
Proximo:
    Next olMail

    ' --- Formatação final
    With ws.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' Larguras de coluna e AutoFilter
    ws.Columns("A:J").AutoFit
    ws.Range("A1").AutoFilter

Finalizar:
    ' Limpar objetos
    On Error Resume Next
    Set olMail = Nothing
    Set items = Nothing
    Set olSubFolder = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    On Error GoTo 0
    Exit Sub

TratarErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
    Resume Finalizar
End Sub

'========================
' Funções auxiliares
'========================

' Obtém subpasta via caminho relativo a uma pasta raiz (ex.: "NFE\01.NOTAS")
Function GetFolder(ByVal rootFolder As Object, ByVal folderPath As String) As Object
    Dim arrFolders() As String
    Dim i As Long
    Dim folder As Object

    If rootFolder Is Nothing Then Exit Function
    If Len(folderPath) = 0 Then Exit Function

    arrFolders = Split(folderPath, "\")
    Set folder = rootFolder

    For i = LBound(arrFolders) To UBound(arrFolders)
        On Error Resume Next
        Set folder = folder.Folders(arrFolders(i))
        On Error GoTo 0
        If folder Is Nothing Then Exit For
    Next i

    Set GetFolder = folder
End Function

' Cria ou retorna uma planilha pelo nome
Function ObterOuCriarPlanilha(ByVal nome As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = nome
    End If
    Set ObterOuCriarPlanilha = ws
End Function

' Normaliza uma linha: remove TAB, quebras, múltiplos espaços, "?" e NBSP
Function NormalizarLinha(ByVal s As String) As String
    If Len(s) = 0 Then
        NormalizarLinha = ""
        Exit Function
    End If

    s = Replace(s, ChrW(&HA0), " ")   ' NBSP -> espaço
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, "?", "")           ' alguns e-mails vêm com "?" por encoding

    ' Colapsar espaços
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizarLinha = Trim$(s)
End Function

' Retorna o texto após o primeiro ":" (após normalizar)
Function ExtrairAposDoisPontos(ByVal texto As String) As String
    Dim pos As Long
    pos = InStr(1, texto, ":")
    If pos > 0 Then
        ExtrairAposDoisPontos = Trim$(Mid$(texto, pos + 1))
    Else
        ExtrairAposDoisPontos = ""
    End If
End Function

' Extrai somente dígitos de uma string
Function ExtrairNumeros(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then r = r & ch
    Next i
    ExtrairNumeros = r
End Function

' Valida PEDIDO/OC (deve iniciar por "45" e ter 10 dígitos)
Function EhPedidoOcValido(ByVal codigo As String) As Boolean
    EhPedidoOcValido = (Len(codigo) = 10 And Left$(codigo, 2) = "45")
End Function

' Verifica se a propriedade existe (para evitar erros em alguns itens)
Function HasProperty(ByVal obj As Object, ByVal propName As String) As Boolean
    On Error Resume Next
    Dim tmp As Variant
    tmp = CallByName(obj, propName, VbGet)
    HasProperty = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function





OUTLOOK

Sub ProcurarTextoNaColunaO()

    Dim mailItem As Object
    Dim selectedText As String
    Dim excelApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim cell As Object
    Dim wbName1 As String
    Dim wbName2 As String
    Dim encontrado As Boolean
    Dim procurouPrimeiro As Boolean

    ' === Obter item ativo e texto selecionado ===
    On Error Resume Next
    Set mailItem = Application.ActiveInspector.CurrentItem
    On Error GoTo 0

    If mailItem Is Nothing Then
        MsgBox "Nenhum e-mail ativo.", vbExclamation
        Exit Sub
    End If

    selectedText = Trim(Application.ActiveInspector.WordEditor.Application.Selection.Text)

    ' Limpar caracteres invisíveis
    selectedText = Replace(selectedText, vbCrLf, "")
    selectedText = Replace(selectedText, vbLf, "")
    selectedText = Replace(selectedText, vbCr, "")
    selectedText = Replace(selectedText, Chr(160), "") ' espaço não separável
    selectedText = Trim(selectedText)

    If Len(selectedText) = 0 Then
        MsgBox "Selecione um texto válido no corpo do e-mail.", vbExclamation
        Exit Sub
    End If

    ' === Conectar ao Excel ===
    Set excelApp = GetObject(, "Excel.Application")
    If excelApp Is Nothing Then
        MsgBox "Excel não está aberto.", vbCritical
        Exit Sub
    End If

    ' === Nomes das planilhas (workbooks) ===
    wbName1 = "Planilha de Chamados IARA+ 2025 v8 - 122025.xlsx"
    wbName2 = "Planilha de Chamados IARA+ 2025 v8 - 012026.xlsx"

    ' ===== Função local para obter workbook por nome exato entre os abertos =====
    Dim FunctionGetWb As Object
    ' (usaremos Set wb = Nothing e loop direto abaixo para evitar criar função separada)

    ' === Tenta procurar no primeiro workbook (122025) ===
    Set wb = Nothing
    Dim wbLoop As Object
    For Each wbLoop In excelApp.Workbooks
        If wbLoop.Name = wbName1 Then
            Set wb = wbLoop
            Exit For
        End If
    Next wbLoop

    procurouPrimeiro = False
    encontrado = False

    If Not wb Is Nothing Then
        procurouPrimeiro = True

        ' Tenta obter a aba
        On Error Resume Next
        Set ws = wb.Sheets("Preechimento") ' ajuste se o nome correto for "Preenchimento"
        On Error GoTo 0

        If Not ws Is Nothing Then
            For Each cell In ws.Range("O1:O3000")
                If StrComp(Trim(CStr(cell.Value)), selectedText, vbTextCompare) = 0 Then
                    encontrado = True
                    MsgBox "Texto encontrado na célula " & cell.Address & _
                           " da aba '" & ws.Name & "' em '" & wb.Name & "'.", vbInformation
                    excelApp.Visible = True
                    wb.Activate
                    ws.Activate
                    cell.Select
                    cell.Application.GoTo cell, True
                    Exit Sub
                End If
            Next cell
        End If
        ' Se chegou aqui, não encontrou no primeiro workbook (ou aba inexistente)
    End If

    ' === Tenta procurar no segundo workbook (012026) ===
    Set wb = Nothing
    For Each wbLoop In excelApp.Workbooks
        If wbLoop.Name = wbName2 Then
            Set wb = wbLoop
            Exit For
        End If
    Next wbLoop

    If Not wb Is Nothing Then
        On Error Resume Next
        Set ws = wb.Sheets("Preechimento") ' ajuste se o nome correto for "Preenchimento"
        On Error GoTo 0

        If Not ws Is Nothing Then
            For Each cell In ws.Range("O1:O3000")
                If StrComp(Trim(CStr(cell.Value)), selectedText, vbTextCompare) = 0 Then
                    encontrado = True
                    MsgBox "Texto encontrado na célula " & cell.Address & _
                           " da aba '" & ws.Name & "' em '" & wb.Name & "'.", vbInformation
                    excelApp.Visible = True
                    wb.Activate
                    ws.Activate
                    cell.Select
                    cell.Application.GoTo cell, True
                    Exit Sub
                End If
            Next cell
        End If
    End If

    ' === Mensagens finais conforme resultados ===
    If Not encontrado Then
        Dim infoPlanilhas As String

        ' Monta informação sobre disponibilidade dos arquivos
        Dim aberto1 As Boolean, aberto2 As Boolean
        aberto1 = False: aberto2 = False
        For Each wbLoop In excelApp.Workbooks
            If wbLoop.Name = wbName1 Then aberto1 = True
            If wbLoop.Name = wbName2 Then aberto2 = True
        Next wbLoop

        If Not aberto1 And Not aberto2 Then
            infoPlanilhas = " (Observação: nenhuma das duas planilhas está aberta no Excel.)"
        ElseIf Not aberto1 Then
            infoPlanilhas = " (Observação: a planilha '" & wbName1 & "' não está aberta.)"
        ElseIf Not aberto2 Then
            infoPlanilhas = " (Observação: a planilha '" & wbName2 & "' não está aberta.)"
        Else
            infoPlanilhas = ""
        End If

        MsgBox "Texto '" & selectedText & "' NÃO foi encontrado em nenhuma das duas planilhas:" & vbCrLf & _
               " - " & wbName1 & vbCrLf & " - " & wbName2 & infoPlanilhas, vbInformation
    End If

End Sub


                                                                                                                                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                                                                                                                                Sub BaixarAnexosNaoLidosOLD()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olSubFolder As Outlook.MAPIFolder
    Dim olMail As Outlook.mailItem
    Dim olAttachment As Outlook.Attachment
    Dim pastaDestino As String
    Dim anexosBaixados As Boolean
    Dim item As Object
    Dim nomeArquivo As String
    Dim filteredItems As Outlook.items
    
    On Error GoTo TratamentoErro
    
    Set olApp = New Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    pastaDestino = Environ$("USERPROFILE") & "\OneDrive - ALLOS\Área de Trabalho\Anexos\"
    If Dir(pastaDestino, vbDirectory) = "" Then MkDir pastaDestino
    
    Set olFolder = olNamespace.GetDefaultFolder(olFolderInbox)
    
    ' Acesso seguro à subpasta
    On Error Resume Next
    Set olSubFolder = olNamespace.Folders("NFE")
    If olSubFolder Is Nothing Then
        Set olSubFolder = olFolder.Folders("NFE")
    End If
    If Not olSubFolder Is Nothing Then
        Set olSubFolder = olSubFolder.Folders("01.NOTAS")
    End If
    On Error GoTo TratamentoErro
    
    If olSubFolder Is Nothing Then
        MsgBox "Subpasta 'NFE > 01.NOTAS' não encontrada.", vbExclamation, "Erro"
        GoTo Finalizar
    End If
    
    anexosBaixados = False
    
    ' Filtro otimizado para não lidos
    Set filteredItems = olSubFolder.items.Restrict("@SQL=""urn:schemas:httpmail:read"" = 0")
    
    For Each item In filteredItems
        If TypeOf item Is Outlook.mailItem Then
            Set olMail = item
            For Each olAttachment In olMail.Attachments
                If LCase(Right(olAttachment.fileName, 4)) = ".xml" Or LCase(Right(olAttachment.fileName, 4)) = ".pdf" Then
                    nomeArquivo = pastaDestino & olAttachment.fileName
                    
                    If Dir(nomeArquivo) <> "" Then
                        Dim contador As Integer
                        contador = 1
                        Do While Dir(Left(nomeArquivo, InStrRev(nomeArquivo, ".") - 1) & "_" & contador & Mid(nomeArquivo, InStrRev(nomeArquivo, "."))) <> ""
                            contador = contador + 1
                        Loop
                        nomeArquivo = Left(nomeArquivo, InStrRev(nomeArquivo, ".") - 1) & "_" & contador & Mid(nomeArquivo, InStrRev(nomeArquivo, "."))
                    End If
                    
                    olAttachment.SaveAsFile nomeArquivo
                    anexosBaixados = True
                End If
            Next olAttachment
            
            ' Descomentar para marcar como lido
            ' olMail.UnRead = False
            ' olMail.Save
        End If
    Next item
    
    If anexosBaixados Then
        MsgBox "Anexos baixados com sucesso!", vbInformation, "Sucesso"
    Else
        MsgBox "Nenhum anexo XML/PDF encontrado.", vbInformation, "Aviso"
    End If
 
Finalizar:
    Set olAttachment = Nothing
    Set olMail = Nothing
    Set olSubFolder = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    If Not filteredItems Is Nothing Then Set filteredItems = Nothing
    Exit Sub
 
TratamentoErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Falha"
    Resume Finalizar
End Sub






Sub BaixarAnexosNaoLidosOLD()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olSubFolder As Outlook.MAPIFolder
    Dim olMail As Outlook.mailItem
    Dim olAttachment As Outlook.Attachment
    Dim pastaDestino As String
    Dim anexosBaixados As Boolean
    Dim item As Object
    Dim nomeArquivo As String
    Dim filteredItems As Outlook.items
    
    On Error GoTo TratamentoErro
    
    Set olApp = New Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    pastaDestino = Environ$("USERPROFILE") & "\OneDrive - ALLOS\Área de Trabalho\Anexos\"
    If Dir(pastaDestino, vbDirectory) = "" Then MkDir pastaDestino
    
    Set olFolder = olNamespace.GetDefaultFolder(olFolderInbox)
    
    ' Acesso seguro à subpasta
    On Error Resume Next
    Set olSubFolder = olNamespace.Folders("NFE")
    If olSubFolder Is Nothing Then
        Set olSubFolder = olFolder.Folders("NFE")
    End If
    If Not olSubFolder Is Nothing Then
        Set olSubFolder = olSubFolder.Folders("01.NOTAS")
    End If
    On Error GoTo TratamentoErro
    
    If olSubFolder Is Nothing Then
        MsgBox "Subpasta 'NFE > 01.NOTAS' não encontrada.", vbExclamation, "Erro"
        GoTo Finalizar
    End If
    
    anexosBaixados = False
    
    ' Filtro otimizado para não lidos
    Set filteredItems = olSubFolder.items.Restrict("@SQL=""urn:schemas:httpmail:read"" = 0")
    
    For Each item In filteredItems
        If TypeOf item Is Outlook.mailItem Then
            Set olMail = item
            For Each olAttachment In olMail.Attachments
                If LCase(Right(olAttachment.fileName, 4)) = ".xml" Or LCase(Right(olAttachment.fileName, 4)) = ".pdf" Then
                    nomeArquivo = pastaDestino & olAttachment.fileName
                    
                    If Dir(nomeArquivo) <> "" Then
                        Dim contador As Integer
                        contador = 1
                        Do While Dir(Left(nomeArquivo, InStrRev(nomeArquivo, ".") - 1) & "_" & contador & Mid(nomeArquivo, InStrRev(nomeArquivo, "."))) <> ""
                            contador = contador + 1
                        Loop
                        nomeArquivo = Left(nomeArquivo, InStrRev(nomeArquivo, ".") - 1) & "_" & contador & Mid(nomeArquivo, InStrRev(nomeArquivo, "."))
                    End If
                    
                    olAttachment.SaveAsFile nomeArquivo
                    anexosBaixados = True
                End If
            Next olAttachment
            
            ' Descomentar para marcar como lido
            ' olMail.UnRead = False
            ' olMail.Save
        End If
    Next item
    
    If anexosBaixados Then
        MsgBox "Anexos baixados com sucesso!", vbInformation, "Sucesso"
    Else
        MsgBox "Nenhum anexo XML/PDF encontrado.", vbInformation, "Aviso"
    End If
 
Finalizar:
    Set olAttachment = Nothing
    Set olMail = Nothing
    Set olSubFolder = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    If Not filteredItems Is Nothing Then Set filteredItems = Nothing
    Exit Sub
 
TratamentoErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Falha"
    Resume Finalizar
End Sub
 







Option Explicit

' ============================
' Configurações
' ============================
Private Const MARK_AS_READ As Boolean = False   ' True para marcar e-mail como lido após salvar
Private Const FILTRAR_APENAS_PDF_XML As Boolean = True
Private Const LOG_DEBUG As Boolean = True       ' Exibe motivos no Immediate (Ctrl+G)

' Ajuste seu caminho padrão
Private Function PASTA_BASE_DESTINO() As String
    PASTA_BASE_DESTINO = Environ$("USERPROFILE") & "\OneDrive - ALLOS\Área de Trabalho\Anexos\"
End Function

' ============================
' Entrada principal
' ============================
Public Sub BaixarAnexosRenomeados()
    On Error GoTo TratamentoErro
    
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olInbox As Outlook.MAPIFolder
    Dim olSubFolder As Outlook.MAPIFolder
    Dim filteredItems As Outlook.items
    Dim item As Object
    Dim mail As Outlook.mailItem
    Dim att As Outlook.Attachment
    
    Dim pastaDestinoBase As String
    pastaDestinoBase = PASTA_BASE_DESTINO()
    EnsureFolderExistsRecursive pastaDestinoBase
    
    Set olApp = Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olInbox = olNamespace.GetDefaultFolder(olFolderInbox)
    
    ' Tenta localizar "NFE > 01.NOTAS" em raiz do mailbox OU dentro da Caixa de Entrada
    Set olSubFolder = TryGetSubFolder(olNamespace, olInbox, "NFE", "01.NOTAS")
    If olSubFolder Is Nothing Then
        MsgBox "Subpasta 'NFE > 01.NOTAS' não encontrada.", vbExclamation, "Aviso"
        GoTo Finalizar
    End If
    
    ' Filtra somente não lidos
    Set filteredItems = olSubFolder.items.Restrict("@SQL=""urn:schemas:httpmail:read"" = 0")
    
    Dim anexosBaixados As Boolean
    anexosBaixados = False
    
    Dim corpoRaw As String, corpoLinhas As String, corpoCompacto As String
    Dim assunto As String, senderName As String
    Dim pedido As String, oc As String, pedidoValido As Boolean, ocValido As Boolean
    Dim numeroNF As String, fornecedor As String
    
    For Each item In filteredItems
        If TypeOf item Is Outlook.mailItem Then
            Set mail = item
            If mail.Attachments.Count > 0 Then
                assunto = NzStr(mail.Subject)
                senderName = NzStr(mail.senderName)
                corpoRaw = NzStr(mail.Body)
                If Len(corpoRaw) = 0 Then GoTo ProximoItem
                
                ' Normalizações para extração
                corpoLinhas = NormalizeForLines(corpoRaw)
                corpoCompacto = RemoveAllWhitespace(corpoRaw)
                
                ' Extrai PEDIDO/OC
                pedido = OnlyDigits(RegexGetFirst(corpoLinhas, LabelValuePatternTolerant("PEDIDO")))
                oc = OnlyDigits(RegexGetFirst(corpoLinhas, LabelValuePatternTolerant("OC")))
                If Len(pedido) = 0 Then pedido = OnlyDigits(RegexGetFirst(corpoCompacto, LabelValuePatternAfterCompact("PEDIDO")))
                If Len(oc) = 0 Then oc = OnlyDigits(RegexGetFirst(corpoCompacto, LabelValuePatternAfterCompact("OC")))
                
                pedidoValido = EhPedidoOcValido(pedido)
                ocValido = EhPedidoOcValido(oc)
                
                If Not pedidoValido And Not ocValido Then
                    If LOG_DEBUG Then Debug.Print "Ignorado e-mail sem pedido/OC válido: "; assunto
                    GoTo ProximoItem
                End If
                
                Dim pedidoUsar As String
                pedidoUsar = IIf(pedidoValido, pedido, oc)
                
                ' Extrai NF (se houver; senão tentamos pelo nome do anexo)
                numeroNF = Trim$(RegexGetFirst( _
                    corpoLinhas, _
                    "(?:N[º°o]?\s*(?:da\s*)?Nota\s*Fiscal|Nota\s*Fiscal|N[º°o]?\s*NF(?:-?e)?|NF-?e|NF)\s*[:\-–—]?\s*(\d[\dA-Za-z\.\-\/]*)" _
                ))
                
                ' Extrai fornecedor; fallback = remetente
                fornecedor = RegexGetLineValue(corpoLinhas, "Fornecedor")
                If Len(Trim$(fornecedor)) = 0 Then fornecedor = senderName
                
                ' Prefixo CG x FPP (ajuste a regra se quiser)
                Dim prefixo As String
                prefixo = DetectarPrefixo(corpoLinhas, assunto)    ' "FPP" ou "CG"
                
                ' Pasta do pedido
                Dim pastaPedido As String
                pastaPedido = pastaDestinoBase
                'pastaPedido = pastaDestinoBase & pedidoUsar & "\"
                EnsureFolderExistsRecursive pastaPedido
                
                ' Percorre anexos
                Dim i As Long
                For i = 1 To mail.Attachments.Count
                    Set att = mail.Attachments(i)
                    
                    Dim nomeArq As String, ext As String
                    nomeArq = NzStr(att.fileName)
                    ext = "." & LCase$(ObterExtensao(nomeArq))
                    
                    ' Filtrar extensões
                    If FILTRAR_APENAS_PDF_XML Then
                        If Not (ext = ".xml" Or ext = ".pdf") Then
                            If LOG_DEBUG Then Debug.Print "Ignorado (extensão não permitida): "; nomeArq
                            GoTo ProximoAnexo
                        End If
                    End If
                    
                    ' ---- BLOQUEIO: não baixar se for PO (Purchase Order) pelo NOME ----
                    If IsArquivoPO(LCase$(nomeArq)) Then
                        If LOG_DEBUG Then Debug.Print "Ignorado (PO detectado no nome): "; nomeArq
                        GoTo ProximoAnexo
                    End If
                    
                    ' Detectar tipo (XML / BOLETO / NF / INVALIDO_PO / DESCONHECIDO)
                    Dim tipo As String
                    tipo = DetectarTipoAnexo(nomeArq, ext)
                    
                    ' Se a detecção indicar PO, ou desconhecido (PDF que não é NF nem Boleto) -> pular
                    If tipo = "INVALIDO_PO" Then
                        If LOG_DEBUG Then Debug.Print "Ignorado (PO detectado pela detecção): "; nomeArq
                        GoTo ProximoAnexo
                    End If
                    If tipo = "DESCONHECIDO" Then
                        If LOG_DEBUG Then Debug.Print "Ignorado (PDF não reconhecido como NF/BOLETO): "; nomeArq
                        GoTo ProximoAnexo
                    End If
                    
                    Dim fornecedorUsar As String
                    fornecedorUsar = SafeFileComponent(fornecedor)
                    
                    Dim nfUsar As String
                    nfUsar = Trim$(numeroNF)
                    If Len(nfUsar) = 0 Then
                        nfUsar = ExtrairNumeroProvavelDeNF(nomeArq)
                    End If
                    
                    Dim novoNome As String
                    Select Case tipo
                        Case "XML"
                            novoNome = prefixo & "_" & pedidoUsar & "_XML" & ext
                        Case "BOLETO"
                            novoNome = prefixo & "_" & pedidoUsar & "_BOLETO" & ext
                        Case "NF"
                            If Len(nfUsar) > 0 Then
                                novoNome = prefixo & "_" & pedidoUsar & "_NF " & nfUsar & "_" & fornecedorUsar & ext
                            Else
                                novoNome = prefixo & "_" & pedidoUsar & "_NF_" & fornecedorUsar & ext
                            End If
                        Case Else
                            ' Segurança
                            If LOG_DEBUG Then Debug.Print "Ignorado (tipo inesperado): "; nomeArq
                            GoTo ProximoAnexo
                    End Select
                    
                    Dim destino As String
                    destino = GetUniqueFilePath(pastaPedido & novoNome)
                    
                    att.SaveAsFile destino
                    anexosBaixados = True
                    If LOG_DEBUG Then Debug.Print "Salvo: "; destino
ProximoAnexo:
                Next i
                
                ' Marcar como lido se desejado (apenas se baixou algo deste e-mail)
                If MARK_AS_READ And anexosBaixados Then
                    mail.UnRead = False
                    mail.Save
                End If
            End If
        End If
ProximoItem:
    Next item
    
    If anexosBaixados Then
        MsgBox "Anexos baixados e renomeados com sucesso!", vbInformation, "Sucesso"
    Else
        MsgBox "Nenhum anexo elegível (PDF/XML) encontrado em não lidos.", vbInformation, "Aviso"
    End If

Finalizar:
    On Error Resume Next
    Set att = Nothing
    Set mail = Nothing
    Set filteredItems = Nothing
    Set olSubFolder = Nothing
    Set olInbox = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    Exit Sub

TratamentoErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Falha"
    Resume Finalizar
End Sub

' ============================
' ======= Funções util =======
' ============================

Private Function TryGetSubFolder(ns As Outlook.NameSpace, inbox As Outlook.MAPIFolder, _
                                 topLevel As String, subLevel As String) As Outlook.MAPIFolder
    On Error Resume Next
    Dim fld As Outlook.MAPIFolder
    
    ' 1) No topo do mailbox
    Set fld = ns.Folders(topLevel)
    If Not fld Is Nothing Then
        Set fld = fld.Folders(subLevel)
        If Not fld Is Nothing Then
            Set TryGetSubFolder = fld
            Exit Function
        End If
    End If
    
    ' 2) Dentro da Caixa de Entrada
    Set fld = inbox.Folders(topLevel)
    If Not fld Is Nothing Then
        Set fld = fld.Folders(subLevel)
        If Not fld Is Nothing Then
            Set TryGetSubFolder = fld
            Exit Function
        End If
    End If
    On Error GoTo 0
End Function

Private Function NzStr(ByVal s As Variant) As String
    If IsNull(s) Or IsEmpty(s) Then
        NzStr = ""
    Else
        NzStr = CStr(s)
    End If
End Function

' -------- Normalização --------

Private Function NormalizeForLines(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    t = Replace(t, vbTab, " ")
    t = Replace(t, Chr$(160), " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormalizeForLines = t
End Function

Private Function RemoveAllWhitespace(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, vbCrLf, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")
    t = Replace(t, Chr$(160), "")
    t = Replace(t, " ", "")
    RemoveAllWhitespace = t
End Function

' -------- RegEx helpers --------

Private Function RegexGetFirst(ByVal texto As String, ByVal pattern As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .pattern = pattern
        .IgnoreCase = True
        .Global = False
        .Multiline = True
    End With
    If re.Test(texto) Then
        Set m = re.Execute(texto)(0)
        If m.SubMatches.Count > 0 Then
            RegexGetFirst = Trim$(m.SubMatches(0))
        Else
            RegexGetFirst = Trim$(m.Value)
        End If
    Else
        RegexGetFirst = ""
    End If
End Function

Private Function LabelValuePatternTolerant(ByVal label As String) As String
    LabelValuePatternTolerant = label & "\s*[:\-–—]?\s*([^\r\n]+)"
End Function

Private Function LabelValuePatternAfterCompact(ByVal label As String) As String
    LabelValuePatternAfterCompact = label & "[:\-–—]*\s*([0-9\.\-\/]+)"
End Function

Private Function RegexGetLineValue(ByVal texto As String, ByVal label As String) As String
    RegexGetLineValue = RegexGetFirst(texto, LabelValuePatternTolerant(label))
End Function

' -------- Extras --------

Private Function OnlyDigits(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then r = r & ch
    Next i
    OnlyDigits = r
End Function

Private Function EhPedidoOcValido(ByVal codigo As String) As Boolean
    EhPedidoOcValido = (Len(codigo) = 10 And Left$(codigo, 2) = "45")
End Function

Private Function DetectarPrefixo(ByVal corpo As String, ByVal assunto As String) As String
    ' Regra simples: achou "FPP" no assunto/corpo -> "FPP", senão "CG"
    If InStr(1, assunto, "FPP", vbTextCompare) > 0 Or InStr(1, corpo, "FPP", vbTextCompare) > 0 Then
        DetectarPrefixo = "FPP"
    Else
        DetectarPrefixo = "CG"
    End If
End Function

Private Function ObterExtensao(ByVal fileName As String) As String
    Dim p As Long: p = InStrRev(fileName, ".")
    If p > 0 Then
        ObterExtensao = Mid$(fileName, p + 1)
    Else
        ObterExtensao = ""
    End If
End Function

' -------- Checagem específica de PO (Purchase Order) --------
' Cobre: "PO123", "reqPO123", "PO_123", "PO-123", "PO 123",
'        "purchase order", "pedido de compra", "ordem de compra"
Private Function IsArquivoPO(ByVal nmLower As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .IgnoreCase = True
        .Global = False
        .Multiline = False
        .pattern = "po\d+|(^|[^A-Za-z0-9])po([^A-Za-z0-9]|$)|purchase[\s_\-]*order|pedido[\s_\-]*de[\s_\-]*compra|ordem[\s_\-]*de[\s_\-]*compra"
    End With
    IsArquivoPO = re.Test(nmLower)
End Function

' -------- Detecção do tipo do anexo --------
Private Function DetectarTipoAnexo(ByVal nomeArq As String, ByVal ext As String) As String
    Dim nm As String: nm = LCase$(nomeArq)
    
    ' XML -> NFe
    If ext = ".xml" Then
        DetectarTipoAnexo = "XML"
        Exit Function
    End If
    
    ' Se o nome indicar PO -> invalidar
    If IsArquivoPO(nm) Then
        DetectarTipoAnexo = "INVALIDO_PO"
        Exit Function
    End If
    
    If ext = ".pdf" Then
        ' BOLETO
        If InStr(nm, "boleto") > 0 Or InStr(nm, "linha digit") > 0 Then
            DetectarTipoAnexo = "BOLETO"
            Exit Function
        End If
        
        ' NF: usa regex pra evitar falso positivo em "info"
        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        With re
            .IgnoreCase = True
            .Global = False
            .Multiline = False
            .pattern = "danfe|nfe|nota[\s_\-]*fiscal|(^|[^A-Za-z])nf(\d|[^A-Za-z]|$)"
        End With
        If re.Test(nm) Then
            DetectarTipoAnexo = "NF"
            Exit Function
        End If
        
        ' PDF que não é BOLETO nem NF -> não baixar
        DetectarTipoAnexo = "DESCONHECIDO"
        Exit Function
    End If
    
    ' Outras extensões (se passarem pelo filtro) -> não baixar
    DetectarTipoAnexo = "DESCONHECIDO"
End Function

Private Function SafeFileComponent(ByVal s As String) As String
    Dim t As String, i As Long, ch As Integer
    t = s
    ' Remove caracteres inválidos para nome de arquivo
    t = Replace(t, "\", " ")
    t = Replace(t, "/", " ")
    t = Replace(t, ":", " ")
    t = Replace(t, "*", " ")
    t = Replace(t, "?", " ")
    t = Replace(t, """", " ")
    t = Replace(t, "<", " ")
    t = Replace(t, ">", " ")
    t = Replace(t, "|", " ")
    ' Remove caracteres de controle (ASCII < 32)
    Dim sb As String
    For i = 1 To Len(t)
        ch = Asc(Mid$(t, i, 1))
        If ch >= 32 Then sb = sb & Chr$(ch)
    Next i
    sb = Trim$(sb)
    Do While InStr(sb, "  ") > 0
        sb = Replace(sb, "  ", " ")
    Loop
    SafeFileComponent = sb
End Function

Private Function GetUniqueFilePath(ByVal fullPath As String) As String
    Dim p As Long, base As String, ext As String, candidate As String, n As Long
    p = InStrRev(fullPath, ".")
    If p > 0 Then
        base = Left$(fullPath, p - 1)
        ext = Mid$(fullPath, p)
    Else
        base = fullPath
        ext = ""
    End If
    candidate = fullPath
    n = 1
    Do While FileExists(candidate)
        n = n + 1
        candidate = base & " (" & n & ")" & ext
    Loop
    GetUniqueFilePath = candidate
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(fullPath, vbNormal)) > 0)
End Function

Private Sub EnsureFolderExistsRecursive(ByVal path As String)
    ' Cria a árvore de pastas se não existir
    Dim fso As Object, parts() As String, cur As String, i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(path) = 0 Then Exit Sub
    If Right$(path, 1) = "\" Then path = Left$(path, Len(path) - 1)
    parts = Split(path, "\")
    If UBound(parts) < 1 Then Exit Sub
    cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Not fso.FolderExists(cur) Then
            On Error Resume Next
            fso.CreateFolder cur
            On Error GoTo 0
        End If
    Next i
End Sub

Private Function ExtrairNumeroProvavelDeNF(ByVal nomeArquivo As String) As String
    ' Busca sequências de 6–10 dígitos no nome do arquivo (heurística)
    Dim re As Object, mc As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .pattern = "(\d{6,10})"
        .Global = True
        .IgnoreCase = True
    End With
    If re.Test(nomeArquivo) Then
        Set mc = re.Execute(nomeArquivo)
        ExtrairNumeroProvavelDeNF = mc(0).SubMatches(0)
    Else
        ExtrairNumeroProvavelDeNF = ""
    End If
End Function









