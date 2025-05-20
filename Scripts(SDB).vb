Sub ExtrairNaoLidos()
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
    For Each olMail In olSubFolder.Items
        If olMail.Class = 43 And olMail.UnRead = True And olMail.Attachments.Count > 0 Then
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
                    MsgBox "Data inválida no e-mail de " & olMail.SenderName, vbExclamation
                    GoTo ProximoEmail
                End If
            End With
            
            ws.Cells(linhaAtual, 2).Value = olMail.SenderName
            
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

------------------------------------------------------------------------------------------------------------------------------------------------------
Sub ExtrairNaoLidosNotasCamposNoPedidoSeparado()
    Dim olApp As Object, olNamespace As Object, olFolder As Object, olSubFolder As Object
    Dim olMail As Object
    Dim ws As Worksheet
    Dim i As Long, linhaAtual As Long
    Dim linhas() As String
    Dim linhaLimpa As String
    Dim conteudoExtraido As String
    Dim campos As Variant
    Dim campo As Variant
    Dim encontrouInfo As Boolean
    
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
    
    ' Cabeçalho simplificado - sem coluna PEDIDO separada
    With ws.Range("A1:C1")
        .Value = Array("DATA RECEBIMENTO", "REMETENTE", "CONTEÚDO DO E-MAIL")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    linhaAtual = 2
    
    ' Campos que queremos extrair para o conteúdo (incluindo PEDIDO)
    campos = Array("PEDIDO:", "OC:", "Nº NOTA FISCAL:", "VALOR:", "VENCIMENTO:", _
                   "FORMA DE PAGAMENTO:", "COMPRADOR:", "FORNECEDOR:", "OBS.")
    
    ' Loop em todos os e-mails não lidos da pasta
    For Each olMail In olSubFolder.Items
        If olMail.Class = 43 And olMail.UnRead = True Then
            conteudoExtraido = ""
            encontrouInfo = False
            
            ' Quebra o corpo em linhas para buscar os campos específicos
            linhas = Split(olMail.Body, vbCrLf)
            For i = LBound(linhas) To UBound(linhas)
                linhaLimpa = Trim(Replace(linhas(i), "?", ""))
                
                For Each campo In campos
                    If InStr(1, linhaLimpa, campo, vbTextCompare) > 0 Then
                        ' Limpar espaços/tabulações entre ":" e valor
                        linhaLimpa = LimparEspacosAposDoisPontos(linhaLimpa)
                        
                        ' Adiciona a linha limpa ao conteúdo extraído, com quebra de linha
                        If conteudoExtraido <> "" Then
                            conteudoExtraido = conteudoExtraido & vbCrLf
                        End If
                        conteudoExtraido = conteudoExtraido & linhaLimpa
                        encontrouInfo = True
                        Exit For
                    End If
                Next campo
            Next i
            
            ' Se não encontrou nenhum campo, colocar mensagem padrão
            If Not encontrouInfo Then
                conteudoExtraido = "Nenhuma informação relevante encontrada"
            End If
            
            ' Preencher dados na planilha
            With ws.Cells(linhaAtual, 1)
                If IsDate(olMail.ReceivedTime) Then
                    .Value = olMail.ReceivedTime
                    .NumberFormat = "dd/mm/yyyy"
                Else
                    .Value = ""
                End If
            End With
            
            ws.Cells(linhaAtual, 2).Value = olMail.SenderName
            ws.Cells(linhaAtual, 3).Value = conteudoExtraido
            
            linhaAtual = linhaAtual + 1
        End If
    Next olMail
    
    ' Formatar planilha
    With ws.UsedRange
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Columns.AutoFit
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

' Função para limpar espaços e tabulações entre ":" e o valor
Function LimparEspacosAposDoisPontos(texto As String) As String
    Dim pos As Long
    Dim campo As String, valor As String
    
    pos = InStr(texto, ":")
    If pos > 0 Then
        campo = Left(texto, pos) ' inclui os dois pontos
        valor = Mid(texto, pos + 1)
        ' Remove todos os espaços e tabs do início do valor
        valor = LTrim(valor)
        LimparEspacosAposDoisPontos = campo & valor
    Else
        LimparEspacosAposDoisPontos = texto
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


------------------------------------------------------------------------------------------------------------------------------------------------------

    
    Sub ExtrairTodos()
    Dim olApp As Object
    Dim olNamespace As Object
    Dim olFolder As Object
    Dim olSubFolder As Object
    Dim olMail As Object
    Dim ws As Worksheet
    Dim i As Long
    Dim corpo As String
    Dim linhas() As String
    Dim dados As Variant
    Dim linhaAtual As Long
    Dim campo As Variant
    Dim valor As String
    Dim pedido As String, oc As String, notaFiscal As String
    Dim linhaLimpa As String
    Dim colunas As Variant

    On Error GoTo TratarErro

    ' Inicializa Outlook
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")

    ' Acessa a pasta padrão Itens Enviados (6) e verifica subpasta "NFE\01.NOTAS"
    Set olFolder = olNamespace.GetDefaultFolder(6) ' 6 = olFolderSentMail
    If olFolder.Folders("NFE").Folders("01.NOTAS").Name = "01.NOTAS" Then
        Set olSubFolder = olFolder.Folders("NFE").Folders("01.NOTAS")
    Else
        MsgBox "Pasta 'NFE\01.NOTAS' não encontrada!", vbCritical
        Exit Sub
    End If

    ' Define a planilha "Dados" e limpa conteúdo anterior
    Set ws = ThisWorkbook.Sheets("Dados")
    ws.Cells.Clear

    ' Configura cabeçalho da planilha
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

    ' Campos a buscar no corpo do e-mail
    dados = Array("PEDIDO:", "OC:", "Nº NOTA FISCAL:", "Nº da NOTA FISCAL:", _
                  "VALOR:", "VENCIMENTO:", "FORMA DE PAGAMENTO:", _
                  "COMPRADOR:", "FORNECEDOR:", "OBS:")

    ' Percorre todos os e-mails da subpasta
    For Each olMail In olSubFolder.Items
        ' Verifica se é um e-mail e possui anexos
        If olMail.Class = 43 And olMail.Attachments.Count > 0 Then
            corpo = olMail.Body
            linhas = Split(corpo, vbCrLf)

            ' Data de recebimento formatada
            If IsDate(olMail.ReceivedTime) Then
                With ws.Cells(linhaAtual, 1)
                    .Value = Format(olMail.ReceivedTime, "dd/mm/yyyy")
                    .NumberFormat = "mm/dd/yyyy"
                End With
            Else
                MsgBox "Erro: Data inválida no e-mail", vbExclamation
                Exit Sub
            End If

            ws.Cells(linhaAtual, 2) = olMail.SenderName

            ' Inicializa variáveis para capturar PEDIDO, OC e Nº NOTA FISCAL
            pedido = ""
            oc = ""
            notaFiscal = ""

            ' Percorre linhas do corpo para capturar dados
            For i = LBound(linhas) To UBound(linhas)
                linhaLimpa = Trim(Replace(linhas(i), "?", ""))

                ' Captura PEDIDO, OC e Nº NOTA FISCAL
                If InStr(1, linhaLimpa, "PEDIDO:", vbTextCompare) > 0 Then
                    pedido = Trim(Split(linhaLimpa, ":")(1))
                ElseIf InStr(1, linhaLimpa, "OC:", vbTextCompare) > 0 Then
                    oc = Trim(Split(linhaLimpa, ":")(1))
                ElseIf InStr(1, linhaLimpa, "Nº NOTA FISCAL:", vbTextCompare) > 0 Or _
                       InStr(1, linhaLimpa, "Nº da NOTA FISCAL:", vbTextCompare) > 0 Then
                    notaFiscal = Trim(Split(linhaLimpa, ":")(1))
                End If

                ' Captura outros campos conforme array dados
                For Each campo In dados
                    Dim pos As Long
                    pos = InStr(1, linhaLimpa, campo, vbTextCompare)
                    If pos > 0 Then
                        valor = Trim(Split(linhaLimpa, ":")(1))
                        Select Case campo
                            Case "VALOR:"
                                ws.Cells(linhaAtual, 5) = valor
                            Case "VENCIMENTO:"
                                ws.Cells(linhaAtual, 6) = valor
                            Case "FORMA DE PAGAMENTO:"
                                ws.Cells(linhaAtual, 7) = valor
                            Case "COMPRADOR:"
                                ws.Cells(linhaAtual, 8) = valor
                            Case "FORNECEDOR:"
                                ws.Cells(linhaAtual, 9) = valor
                            Case "OBS:"
                                ws.Cells(linhaAtual, 10) = valor
                        End Select
                    End If
                Next campo
            Next i

            ' Preenche coluna PEDIDO com pedido ou OC
            If pedido <> "" Then
                ws.Cells(linhaAtual, 3) = pedido
            ElseIf oc <> "" Then
                ws.Cells(linhaAtual, 3) = oc
            End If

            ' Preenche Nº NOTA FISCAL
            If notaFiscal <> "" Then
                ws.Cells(linhaAtual, 4) = notaFiscal
            End If

            ' Ignora e-mails sem pedido ou OC
            If pedido <> "" Or oc <> "" Then
                linhaAtual = linhaAtual + 1
            End If
        End If
    Next olMail

    ' Formatação final da planilha
    With ws.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' Libera objetos
    Set olMail = Nothing
    Set olSubFolder = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing

    Exit Sub

TratarErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
End Sub
------------------------------------------------------------------------------------------------------------------------------------------------------    
    
Sub BaixarAnexosNaoLidos()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.Namespace
    Dim olFolder As Outlook.MAPIFolder
    Dim olSubFolder As Outlook.MAPIFolder
    Dim olMail As Outlook.MailItem
    Dim olAttachment As Outlook.Attachment
    Dim pastaDestino As String
    Dim anexosBaixados As Boolean
    Dim item As Object
    Dim nomeArquivo As String
    
    On Error GoTo TratamentoErro
    
    Set olApp = New Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' Defina a pasta de destino
    pastaDestino = Environ$("USERPROFILE") & "\OneDrive - ALLOS\Área de Trabalho\Anexos\"
    
    ' Cria a pasta se não existir
    If Dir(pastaDestino, vbDirectory) = "" Then MkDir pastaDestino
    
    ' Acesse a pasta padrão de entrada (Caixa de Entrada)
    Set olFolder = olNamespace.GetDefaultFolder(olFolderInbox)
    
    ' Verifique se a pasta "NFE" existe
    On Error Resume Next
    Set olSubFolder = olFolder.Folders("NFE").Folders("01.NOTAS")
    On Error GoTo TratamentoErro
    
    If olSubFolder Is Nothing Then
        MsgBox "A subpasta '01.NOTAS' não foi encontrada.", vbExclamation, "Pasta Não Encontrada"
        GoTo Finalizar
    End If
    
    anexosBaixados = False
    
    ' Itera somente sobre os e-mails não lidos
    For Each item In olSubFolder.Items.Restrict("[UnRead] = True")
        If TypeOf item Is Outlook.MailItem Then
            Set olMail = item
            For Each olAttachment In olMail.Attachments
                If LCase(Right(olAttachment.FileName, 4)) = ".xml" Or LCase(Right(olAttachment.FileName, 4)) = ".pdf" Then
                    ' Para evitar sobrescrever arquivos com mesmo nome, adiciona timestamp
                    nomeArquivo = pastaDestino & Format(Now, "yyyymmdd_hhnnss_") & olAttachment.FileName
                    olAttachment.SaveAsFile nomeArquivo
                    anexosBaixados = True
                End If
            Next olAttachment
            ' Opcional: marcar o e-mail como lido após baixar anexos
            ' olMail.UnRead = False
            ' olMail.Save
        End If
    Next item
    
    If anexosBaixados Then
        MsgBox "Anexos baixados com sucesso!", vbInformation, "Anexos Baixados"
    Else
        MsgBox "Nenhum anexo XML ou PDF encontrado em e-mails não lidos.", vbInformation, "Sem Anexos"
    End If

Finalizar:
    Set olAttachment = Nothing
    Set olMail = Nothing
    Set olSubFolder = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    Exit Sub

TratamentoErro:
    MsgBox "Erro: " & Err.Description, vbCritical, "Erro no VBA"
    Resume Finalizar
End Sub
------------------------------------------------------------------------------------------------------------------------------------------------------
    
    
Sub MarcarEmailsNaoLidosComoLidos()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.Namespace
    Dim olFolder As Outlook.MAPIFolder
    Dim olSubFolder As Outlook.MAPIFolder
    Dim olMail As Outlook.MailItem
    Dim itensNaoLidos As Outlook.Items
    Dim item As Object
    
    Set olApp = New Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(olFolderInbox)
    
    ' Tenta acessar a subpasta "01.NOTAS" dentro da pasta "NFE"
    On Error Resume Next
    Set olSubFolder = olFolder.Folders("NFE").Folders("01.NOTAS")
    On Error GoTo 0
    
    If olSubFolder Is Nothing Then
        MsgBox "A subpasta '01.NOTAS' não foi encontrada.", vbExclamation, "Pasta Não Encontrada"
        Exit Sub
    End If
    
    ' Obter somente os e-mails não lidos
    Set itensNaoLidos = olSubFolder.Items.Restrict("[UnRead] = True")
    
    If itensNaoLidos Is Nothing Or itensNaoLidos.Count = 0 Then
        MsgBox "Nenhum e-mail não lido encontrado nesta pasta.", vbInformation, "Sem E-mails Não Lidos"
        Exit Sub
    End If
    
    Dim i As Long
    For i = itensNaoLidos.Count To 1 Step -1
        Set item = itensNaoLidos(i)
        If TypeOf item Is Outlook.MailItem Then
            Set olMail = item
            olMail.UnRead = False
            olMail.Save
        End If
    Next i
    
    MsgBox "Todos os e-mails não lidos foram marcados como lidos.", vbInformation, "Processo Concluído"
    
    ' Limpeza
    Set olMail = Nothing
    Set itensNaoLidos = Nothing
    Set olSubFolder = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
End Sub
------------------------------------------------------------------------------------------------------------------------------------------------------

Sub ExportarParaTXT()
    Dim FilePath As String
    Dim FileNum As Integer
    Dim Linha As Range
    
    ' Definir o caminho do arquivo na ï¿½rea de trabalho
    FilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Exportado.txt"
    
    ' Obter um nï¿½mero de arquivo livre
    FileNum = FreeFile
    
    ' Abrir o arquivo para escrita
    Open FilePath For Output As FileNum
    
    ' Iterar pelas cï¿½lulas na planilha ativa (ajuste conforme necessï¿½rio)
    For Each Linha In ActiveSheet.UsedRange.Rows
        Dim TextoLinha As String
        TextoLinha = ""
        
        ' Concatenar os valores das cï¿½lulas da linha com tabulaï¿½ï¿½o como separador
        For Each Celula In Linha.Cells
            TextoLinha = TextoLinha & Celula.Value & vbTab
        Next Celula
        
        ' Remover o ï¿½ltimo separador e escrever no arquivo
        Print #FileNum, Left(TextoLinha, Len(TextoLinha) - 1)
    Next Linha
    
    ' Fechar o arquivo
    Close FileNum
    
    MsgBox "Arquivo exportado para: " & FilePath, vbInformation, "Exportaï¿½ï¿½o Concluï¿½da"
End Sub


