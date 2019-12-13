Attribute VB_Name = "NFCeAPI"
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'activate Microsoft XML, v6.0 in references

'Atributo privado da classe
Private Const tempoResposta = 500
Private Const impressaoParam = """impressao"":{" & """tipo"":""pdf""," & """ecologica"":false," & """itemLinhas"":""1""," & """itemDesconto"":false," & """larguraPapel"":""80mm""}"
Private Const token = "SEU_TOKEN"

'Esta função envia um conteï¿½do para uma URL, em requisições do tipo POST
Function enviaConteudoParaAPI(conteudo As String, url As String, tpConteudo As String) As String
On Error GoTo SAI
    Dim contentType As String
    
    If (tpConteudo = "txt") Then
        contentType = "text/plain;charset=utf-8"
    ElseIf (tpConteudo = "xml") Then
        contentType = "application/xml;charset=utf-8"
    Else
        contentType = "application/json;charset=utf-8"
    End If
    
    Dim obj As MSXML2.ServerXMLHTTP60
    Set obj = New MSXML2.ServerXMLHTTP60
    obj.Open "POST", url
    obj.setRequestHeader "Content-Type", contentType
    If Trim(token) <> "" Then
        obj.setRequestHeader "X-AUTH-TOKEN", token
    End If
    obj.send conteudo
    Dim resposta As String
    resposta = obj.responseText
    
    Select Case obj.status
        Case 401
            MsgBox ("Token não enviado ou invï¿½lido")
        Case 403
            MsgBox ("Token sem permissão")
    End Select
    
    enviaConteudoParaAPI = resposta
    Exit Function
SAI:
  enviaConteudoParaAPI = "{" & """status"":""" & Err.Number & """," & """motivo"":""" & Err.Description & """" & "}"
End Function

'Esta função realiza o processo completo de emissãoo: envio e download do documento
Public Function emitirNFCeSincrono(conteudo As String, tpConteudo As String, tpAmb As String, caminho As String, exibeNaTela As Boolean) As String
    Dim retorno As String
    Dim resposta As String
    Dim statusEnvio As String
    Dim statusDownload As String
    Dim motivo As String
    Dim erros As String
    Dim chNFe As String
    Dim cStat As String
    Dim nProt As String

    statusEnvio = ""
    statusDownload = ""
    motivo = ""
    erros = ""
    chNFe = ""
    cStat = ""
    nProt = ""
    
    gravaLinhaLog ("[EMISSAO_SINCRONA_INICIO]")
    
    resposta = emitirNFCe(conteudo, tpConteudo)
    statusEnvio = LerDadosJSON(resposta, "status", "", "")

    'Testa se o envio foi feito com sucesso (200) ou se ï¿½ para reconsultar (-6)
    If (statusEnvio = "100") Or (statusEnvio = "-100") Then
    
        cStat = LerDadosJSON(resposta, "nfeProc", "cStat", "")

        'Testa se o cStat ï¿½ igual a 100 ou 150, pois ambos significam "Autorizado"
        If (cStat = "100") Or (cStat = "150") Then
        
            chNFe = LerDadosJSON(resposta, "nfeProc", "chNFe", "")
            nProt = LerDadosJSON(resposta, "nfeProc", "nProt", "")
            motivo = LerDadosJSON(resposta, "nfeProc", "xMotivo", "")

            Sleep (tempoResposta)

            resposta = downloadNFCeESalvar(chNFe, tpAmb, caminho, exibeNaTela)
            statusDownload = LerDadosJSON(resposta, "status", "", "")
            
            'Testa se houve problema no download
            If (statusDownload <> "100") Then
            
                motivo = LerDadosJSON(resposta, "motivo", "", "")
                
            End If
        Else
        
            motivo = LerDadosJSON(resposta, "nfeProc", "xMotivo", "")
            
        End If
        
    ElseIf (status = "-995") Then

        motivo = LerDadosJSON(resposta, "motivo", "", "")
        erros = LerDadosJSON(resposta, "erros", "", "")
        
    Else
    
        motivo = LerDadosJSON(resposta, "motivo", "", "")
        
    End If
    
    'Monta o JSON de retorno
    retorno = "{"
    retorno = retorno & """statusEnvio"":""" & statusEnvio & ""","
    retorno = retorno & """statusDownload"":""" & statusDownload & ""","
    retorno = retorno & """cStat"":""" & cStat & ""","
    retorno = retorno & """chNFe"":""" & chNFe & ""","
    retorno = retorno & """nProt"":""" & nProt & ""","
    retorno = retorno & """motivo"":""" & motivo & ""","
    retorno = retorno & """erros"":""" & erros & """"
    retorno = retorno & "}"
    
    gravaLinhaLog ("[JSON_RETORNO]")
    gravaLinhaLog (retorno)
    gravaLinhaLog ("[EMISSAO_SINCRONA_FIM]")

    emitirNFCeSincrono = retorno
End Function

'Esta função realiza o envio de uma NFC-e
Public Function emitirNFCe(conteudo As String, tpConteudo As String) As String

    Dim url As String
    Dim resposta As String

    url = "https://nfce.ns.eti.br/v1/nfce/issue"

    gravaLinhaLog ("[ENVIO_DADOS]")
    gravaLinhaLog (conteudo)
        
    resposta = enviaConteudoParaAPI(conteudo, url, tpConteudo)
    
    gravaLinhaLog ("[ENVIO_RESPOSTA]")
    gravaLinhaLog (resposta)

    emitirNFCe = resposta
End Function

'Esta função realiza o download de documentos de uma NFC-e
Public Function downloadNFCe(chNFe As String, tpAmb As String) As String

    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String

    'Monta o JSON
    json = "{"
    json = json & """chNFe"":""" & chNFe & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & impressaoParam
    json = json & "}"

    url = "https://nfce.ns.eti.br/v1/nfce/get"

    gravaLinhaLog ("[DOWNLOAD_NFCE_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    status = LerDadosJSON(resposta, "status", "", "")
        
    'O retorno da API serão gravado somente em caso de erro,
    'para não gerar um log extenso com o PDF e XML
    If (status <> "100") Then
    
        gravaLinhaLog ("[DOWNLOAD_NFCE_RESPOSTA]")
        gravaLinhaLog (resposta)
        
    Else

        gravaLinhaLog ("[DOWNLOAD_NFCE_STATUS]")
        gravaLinhaLog (status)
        
    End If

    downloadNFCe = resposta
End Function

'Esta função realiza o download de documentos de uma NFC-e e salva-os
Public Function downloadNFCeESalvar(chNFe As String, tpAmb As String, caminho As String, exibeNaTela As Boolean) As String

    Dim xml As String
    Dim pdf As String
    Dim status As String
    Dim resposta As String
    
    resposta = downloadNFCe(chNFe, tpAmb)
    status = LerDadosJSON(resposta, "status", "", "")

    If status = "100" Then
        
        'Cria o diretório, caso não exista
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
    
        xml = LerDadosJSON(resposta, "nfeProc", "xml", "")
        Call salvarXML(xml, caminho, chNFe)
        
        If InStr(1, impressaoParam, "pdf") Then
        
            pdf = LerDadosJSON(resposta, "pdf", "", "")
            Call salvarPDF(pdf, caminho, chNFe)
            
            If exibeNaTela Then
            
                ShellExecute 0, "open", caminho & chNFe & "-procNFCe.pdf", "", "", vbNormalFocus
            
            End If
        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informaï¿½ï¿½es")
        gravaLinhaLog ("[Ocorreu um erro, veja o Retorno da API para mais informações  - Metodo: downloadNFCeESalvar]")
    End If

    downloadNFCeESalvar = resposta
End Function


'Esta função realiza o download de eventos de uma NFC-e e salva-os
Public Function downloadEventoNFCeESalvar(chNFe As String, tpAmb As String, caminho As String, exibeNaTela As Boolean) As String
    Dim xml As String
    Dim chNFeCanc As String
    Dim pdf As String
    Dim status As String
    Dim resposta As String
    
    resposta = downloadNFCe(chNFe, tpAmb)
    status = LerDadosJSON(resposta, "status", "", "")

    If status = "100" Then
    
        'Cria o diretório, caso não exista
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
        
        xml = LerDadosJSON(resposta, "retEvento", "xml", "")
        chNFeCanc = LerDadosJSON(resposta, "retEvento", "chNFeCanc", "")
        Call salvarXML(xml, caminho, chNFeCanc, "CANC")

        If InStr(1, impressaoParam, "pdf") Then
        
            pdf = LerDadosJSON(resposta, "pdfCancelamento", "", "")
            Call salvarPDF(pdf, caminho, chNFeCanc, "CANC")
            
            If exibeNaTela Then
    
                ShellExecute 0, "open", caminho & chNFeCanc & "-procEvenNFe.pdf", "", "", vbNormalFocus
            
            End If
        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informaï¿½ï¿½es")
         gravaLinhaLog ("[Ocorreu um erro, veja o Retorno da API para mais informações  - Metodo: downloadEventoNFCeESalvar]")
    End If

    downloadEventoNFCeESalvar = resposta
End Function

'Esta função realiza o cancelamento de uma NFC-e
Public Function cancelarNFCe(chNFe As String, tpAmb As String, dhEvento As String, nProt As String, xJust As String, caminho As String, exibeNaTela As Boolean) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String
    Dim respostaDownload As String

    'Monta o JSON
    json = "{"
    json = json & """chNFe"":""" & chNFe & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """dhEvento"":""" & dhEvento & ""","
    json = json & """nProt"":""" & nProt & ""","
    json = json & """xJust"":""" & xJust & """"
    json = json & "}"
    
    url = "https://nfce.ns.eti.br/v1/nfce/cancel"
    
    gravaLinhaLog ("[CANCELAMENTO_DADOS]")
    gravaLinhaLog (json)
    
    resposta = enviaConteudoParaAPI(json, url, "json")
        
    gravaLinhaLog ("[CANCELAMENTO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    status = LerDadosJSON(resposta, "status", "", "")
    
    'Se houve sucesso no evento, realiza o download
    If (status = "135") Then
    
        respostaDownload = downloadEventoNFCeESalvar(chNFe, tpAmb, caminho, exibeNaTela)
        status = LerDadosJSON(respostaDownload, "status", "", "")
        
        If (status <> "100") Then
            MsgBox ("Ocorreu um erro ao fazer o download. Verifique os logs.")
        End If
    End If
    
    cancelarNFCe = resposta
End Function

'Esta função realiza a consulta de situação de uma NFC-e
Public Function consultarSituacao(chNFe As String, tpAmb As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """chNFe"":""" & chNFe & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"

    url = "https://nfce.ns.eti.br/v1/nfce/status"
    
    gravaLinhaLog ("[CONSULTA_SITUACAO_DADOS]")
    gravaLinhaLog (json)

    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[CONSULTA_SITUACAO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    consultarSituacao = resposta
End Function

'Esta função realiza a inutilização de um intervalo de numeração de NFC-e
Public Function inutilizar(cUF As String, tpAmb As String, ano As String, CNPJ As String, serie As String, nNFIni As String, nNFFin As String, xJust As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """cUF"":""" & cUF & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """ano"":""" & ano & ""","
    json = json & """CNPJ"":""" & CNPJ & ""","
    json = json & """serie"":""" & serie & ""","
    json = json & """nNFIni"":""" & nNFIni & ""","
    json = json & """nNFFin"":""" & nNFFin & ""","
    json = json & """xJust"":""" & xJust & """"
    json = json & "}"

    url = "https://nfce.ns.eti.br/v1/nfce/inut"
    
    gravaLinhaLog ("[INUTILIZACAO_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[INUTILIZACAO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    inutilizar = resposta
End Function

'Esta função realiza o envio de e-mail de uma NFC-e
Public Function enviarEmail(chNFe As String, enviaEmailDoc As String, email) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """chNFe"":""" & chNFe & ""","
    json = json & """enviaEmailDoc"":" & enviaEmailDoc & ","
    json = json & """email"":["
    
    Dim emails() As String
    Dim i, quantidade As Integer
    
    emails = Split(email, ",")
    
    quantidade = UBound(emails)
    
    For i = 0 To quantidade
        If (i = quantidade) Then
            json = json & """" & emails(i) & """"
        Else
            json = json & """" & emails(i) & ""","
        End If
    Next
    
    json = json & "]"
    json = json & "}"

    url = "https://nfce.ns.eti.br/v1/util/resendemail"
    
    gravaLinhaLog ("[ENVIO_EMAIL_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")

    gravaLinhaLog ("[ENVIO_EMAIL_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    enviarEmail = resposta
End Function

'Esta função salva um XML
Public Sub salvarXML(xml As String, caminho As String, chNFe As String, Optional tipo As String = "")
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String
    Dim extensao As String
    
    If (tipo = "CANC") Then
        extensao = "-procEvenNFe.xml"
    Else
        extensao = "-procNFe.xml"
    End If
    'Seta o caminho para o arquivo XML
    localParaSalvar = caminho & chNFe & nSeqEvento & extensao

    'Remove as contrabarras
    conteudoSalvar = Replace(xml, "\""", "")

    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open
    fsT.WriteText conteudoSalvar
    fsT.SaveToFile localParaSalvar
End Sub

'Esta função salva um PDF
Public Function salvarPDF(pdf As String, caminho As String, chNFe As String, Optional tipo As String = "") As Boolean
On Error GoTo SAI
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String
    Dim extensao As String
    If (tipo = "CANC") Then
        extensao = "-procEvenNFe.pdf"
    Else
        extensao = "-procNFe.pdf"
    End If
    'Seta o caminho para o arquivo PDF
    localParaSalvar = caminho & chNFe & nSeqEvento & extensao

    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Binary As #fnum
    Put #fnum, 1, Base64Decode(pdf)
    Close fnum
    Exit Function
SAI:
    MsgBox (Err.Number & " - " & Err.Description), vbCritical
End Function

'Esta função lê os dados de um JSON
Public Function LerDadosJSON(sJsonString As String, key1 As String, key2 As String, key3 As String, Optional key4 As String, Optional key5 As String) As String
On Error GoTo err_handler
    Dim oScriptEngine As ScriptControl
    Set oScriptEngine = New ScriptControl
    oScriptEngine.Language = "JScript"
    Dim objJSON As Object
    Set objJSON = oScriptEngine.Eval("(" + sJsonString + ")")
    If key1 <> "" And key2 <> "" And key3 <> "" And key4 <> "" And key5 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet), key4, VbGet), key5, VbGet)
    ElseIf key1 <> "" And key2 <> "" And key3 <> "" And key4 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet), key4, VbGet)
    ElseIf key1 <> "" And key2 <> "" And key3 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet)
    ElseIf key1 <> "" And key2 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet)
    ElseIf key1 <> "" Then
        LerDadosJSON = VBA.CallByName(objJSON, key1, VbGet)
    End If
Err_Exit:
    Exit Function
err_handler:
    LerDadosJSON = "Error: " & Err.Description
    Resume Err_Exit
End Function

'Esta função lê os dados de um XML
Public Function LerDadosXML(sXml As String, key1 As String, key2 As String) As String
    On Error Resume Next
    LerDadosXML = ""
    
    Set xml = New DOMDocument60
    xml.async = False
    
    If xml.loadXML(sXml) Then
        'Tentar pegar o strCampoXML
        Set objNodeList = xml.getElementsByTagName(key1 & "//" & key2)
        Set objNode = objNodeList.nextNode
        
        Dim valor As String
        valor = objNode.Text
        
        If Len(Trim(valor)) > 0 Then 'CONSEGUI LER O XML NODE
            LerDadosXML = valor
        End If
        Else
        MsgBox "Nï¿½o foi possï¿½vel ler o conteï¿½do do XML da NFe especificado para leitura.", vbCritical, "ERRO"
    End If
End Function

'Esta função grava uma linha de texto em um arquivo de log
Public Sub gravaLinhaLog(conteudoSalvar As String)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim localParaSalvar As String
    Dim data As String
    
    'Diretório para salvar os logs
    localParaSalvar = App.Path & "\log\"
    
    'Checa se existe o caminho passado para salvar os arquivos
    If Dir(localParaSalvar, vbDirectory) = "" Then
        MkDir (localParaSalvar)
    End If
    
    'Pega data atual
    data = Format(Date, "yyyyMMdd")
    
    'Diretório + nome do arquivo para salvar os logs
    localParaSalvar = App.Path & "\log\" & data & ".txt"
    
    'Pega data e hora atual
    data = DateTime.Now
    
    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Append Shared As #fnum
    Print #fnum, data & " - " & conteudoSalvar
    Close fnum
End Sub
