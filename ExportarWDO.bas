Attribute VB_Name = "ModuloWDO"

' ============================================================
'  ExportarWDO.bas - Modulo VBA para o Monitor WDO
' ------------------------------------------------------------
'  Fork de ExportarWIN.bas (Modulo1) adaptado para o DOLFUT. Os dois
'  modulos convivem na MESMA planilha (WDO_Master_RTD.xlsm) - por isso
'  todo nome Public/Private aqui leva o sufixo WDO, para nao colidir com
'  o Modulo1 do WIN.
'
'  Linha do DOLFUT na planilha DADOS: localizada AUTOMATICAMENTE
'  pela coluna A (imune a insercao/remocao de linhas). Colunas:
'  BLINDAGEM: as colunas NAO sao fixas. A funcao ColPorCampoWDO localiza
'  cada campo pela PROPRIA formula RTD (casa o ticker DOLFUT_F_0 + o
'  codigo do campo), entao um deslocamento de colunas na planilha nao
'  quebra o export. Codigos usados:
'    ULT=ultimo ABE=abertura MAX=maxima MIN=minima FEC=fec_ant
'    VOL=volume  98=agr_compra  99=agr_venda  67=vwap
'  Fallback (se a formula sumir): colunas fixas atuais da linha DOLFUT
'    ultimo4 abertura5 maxima6 minima7 fec_ant8 volume10
'    agr_compra27 agr_venda29 vwap30
'
'  Formula RTD de exemplo (celula D10, locale BR):
'    =RTD("rtdtrading.rtdserver";;"DOLFUT_F_0";"ULT")
'
'  INICIO AUTOMATICO: o evento Workbook_Open (em EstaPastaDeTrabalho)
'  precisa chamar IniciarExportWDO alem do IniciarExportWIN ja existente,
'  10s apos abrir a planilha. Rodar manualmente via Alt+F8 so e necessario
'  se as macros forem bloqueadas na abertura.
'  O CSV e gravado na pasta do monitor: ajuste CAMINHO_CSV abaixo.
'
'  SERVIDOR: IniciarExportWDO tambem sobe o server_wdo.py em segundo
'  plano (sem janela), caso a porta 8002 ainda nao esteja respondendo.
'  Um server ja no ar NUNCA e reiniciado (preserva historico intradiario).
'
'  22/09/2026 - BUG CONHECIDO, NAO CORRIGIDO AO VIVO AINDA: URL_SERVIDOR
'  abaixo aponta pra porta 8002, mas server_wdo.py escuta na 8003 (ver
'  PORT em server_wdo.py). ServidorNoArWDO() portanto nunca acha o server
'  já no ar e tenta subir outro a cada IniciarExportWDO - inofensivo ate
'  agora porque o proprio "pythonw.exe" resolvido via PATH e' o stub da
'  Microsoft Store, que falha em silencio quando chamado via WScript.Shell
'  por automacao COM (mesma causa-raiz documentada abaixo). Uma correcao
'  (trocar 8002->8003 e apontar pythonw pro caminho REAL do interpretador,
'  tipo PYTHONW_REAL) foi tentada nesta sessao e quebrou a compilacao do
'  modulo ao vivo (erro "Somente comentarios podem aparecer apos End Sub"
'  logo apos adicionar os Consts PYTHONW_REAL/PYTHON_REAL) por motivo NAO
'  diagnosticado - o bloco sozinho compila limpo isolado num modulo novo,
'  mas falha especificamente dentro deste modulo com o resto do codigo
'  dele junto. Isso derrubou o Excel (crash + autorecover) e ficou
'  revertido pra essa versao simples (porta 8002, sem PYTHONW_REAL) de
'  proposito. Antes de tentar de novo: isolar por bisseccao qual parte
'  exata do bloco colide (ja descartado: nome duplicado de Sub/Const,
'  caracteres nao-ASCII, Sub/End Sub desbalanceado - todos vieram limpos).
'
'  FLUXO (BOOK/FITA/VAP): este modulo le os topicos BOOK0/T&T0/VAP0 -
'  os MESMOS que o WIN usava. Em 24/08/2026 o usuario reapontou essas
'  janelas de Livro/Fita/VAP no Profit do WINFUTV para o DOLPRO (dolar):
'  o topico RTD (BOOK0 etc.) nao muda quando o ATIVO exibido na janela
'  muda, entao as formulas ja existentes na planilha (A22:H.../J22:N.../
'  W22:X...) passam a refletir o DOLFUT sozinhas, sem precisar colar
'  nada de novo no Excel. CONSEQUENCIA: o Monitor WIN fica SEM fluxo
'  (livro/fita/VAP) ate alguem abrir janelas novas e separadas pra ele -
'  foi decisao consciente do usuario, nao bug. Se um dia o WIN precisar
'  do fluxo de volta, o padrao seria abrir uma janela nova (que vira
'  BOOK1/T&T1/VAP1) e linkar essas na planilha para o WIN.
' ============================================================
Option Explicit

Private Const PASTA_MONITOR As String = "C:\Users\rodri\Desktop\Day trade\Dolar monitor"
Private Const CAMINHO_CSV As String = "C:\Users\rodri\Desktop\Day trade\Dolar monitor\dados_wdo.csv"
Private Const URL_SERVIDOR As String = "http://127.0.0.1:8002/ultimo"
Private Const PLANILHA As String = "DADOS"
Private Const ATIVO As String = "DOLFUT"   ' <== procurado na coluna A
Private Const INTERVALO_SEG As Long = 2

' --- Macro RTD (DI futuro + DOLFUT + WINFUTV) para os cards MACRO/CASADO --
' Ate 21/09/2026 este modulo NAO escrevia esse CSV (a suposicao era que o
' Modulo1, ExportarWIN.bas, ja mantinha o arquivo fresco). Descoberto em
' 22/09/2026 que o Modulo1 escreve na pasta monitor_win\ (o path original
' dele, correto pro server_win.py) - NAO na pasta Dolar monitor\ que o
' server_wdo.py le. Ou seja "Dolar monitor\dados_macro_rtd.csv" ficava
' PARADO (ultima escrita real 24/08/2026) e o cupom implicito do card
' CASADO rodava com DI desatualizado havia quase um mes, sem erro visivel.
' Corrigido fazendo o proprio Modulo3 (WDO) escrever sua copia, no caminho
' certo (ExportarMacroRTDWDO). Aproveitado pra incluir o WINFUTV (mini
' indice, linha 14/coluna D "ULT" na aba DADOS) - server_wdo.py passa a
' usar o preco REAL do WIN via RTD no card MACRO, em vez do proxy Yahoo
' ^BVSP (Ibovespa a vista) usado ate entao.
Private Const CAMINHO_CSV_MACRO As String = "C:\Users\rodri\Desktop\Day trade\Dolar monitor\dados_macro_rtd.csv"

' --- Fluxo: livro de ofertas (BOOK0) e fita (T&T0) -----------------
' Formulas ja existentes na planilha (herdadas do WIN) - so o ATIVO
' exibido na janela do Profit mudou para DOLPRO (ver nota no topo).
' A ancora de cada bloco e localizada pela formula do indice 0 e fica em
' cache (varrer a planilha a cada 2s seria caro demais). Se a ancora sair
' do lugar, a proxima leitura relocaliza sozinha.
' A ALTURA de cada bloco e detectada (nao fixa): basta acrescentar linhas de
' RTD na planilha que o export passa a levar tudo, sem mexer no codigo.
Private Const CAMINHO_CSV_BOOK As String = "C:\Users\rodri\Desktop\Day trade\Dolar monitor\dados_book.csv"
Private Const CAMINHO_CSV_TT As String = "C:\Users\rodri\Desktop\Day trade\Dolar monitor\dados_tt.csv"
Private Const CAMINHO_CSV_VAP As String = "C:\Users\rodri\Desktop\Day trade\Dolar monitor\dados_vap.csv"
Private Const MAX_LINHAS_BLOCO As Long = 2000   ' trava de seguranca da varredura

Private bookLin As Long, bookCol As Long, bookN As Long   ' BOOK0 (HORC idx 0)
Private ttLin As Long, ttCol As Long, ttN As Long         ' T&T0  (DAT  idx 0)
Private vapLin As Long, vapCol As Long, vapN As Long      ' VAP0  (VOL  idx 0)

' --- Ranking acumulado -> aba AcumWDO -------------------------------
' O server compila a fita do dia inteiro e espelha em ranking_acum.csv;
' a cada ACUM_A_CADA ciclos (~30s) o VBA cola na aba AcumWDO (sheet
' PROPRIA do WDO - a aba Acum ja e do WIN, colar no mesmo lugar
' sobrescreveria o ranking do WIN). Crie a aba AcumWDO na planilha se
' quiser usar esse recurso; sem ela, a colagem so falha em silencio.
Private Const CAMINHO_CSV_ACUM As String = "C:\Users\rodri\Desktop\Day trade\Dolar monitor\ranking_acum.csv"
Private Const ABA_ACUM As String = "AcumWDO"
Private Const ACUM_A_CADA As Long = 15      ' ciclos de 2s -> ~30s
Private ciclosAcum As Long

Private proximaExecucao As Date
Private exportando As Boolean

Public Sub IniciarExportWDO()
    Call IniciarServidorWDO
    exportando = True
    Call ExportarWDO
End Sub

' Sobe o server_wdo.py em segundo plano se a porta 8002 nao responder.
' Nao mexe em um server que ja esteja no ar.
Public Sub IniciarServidorWDO()
    If ServidorNoArWDO() Then Exit Sub

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.CurrentDirectory = PASTA_MONITOR

    On Error Resume Next
    ' pythonw = roda sem janela de console
    sh.Run "pythonw.exe """ & PASTA_MONITOR & "\server_wdo.py""", 0, False
    If Err.Number <> 0 Then
        Err.Clear
        ' fallback caso pythonw nao esteja no PATH (janela oculta)
        sh.Run "python.exe """ & PASTA_MONITOR & "\server_wdo.py""", 0, False
    End If
    On Error GoTo 0
End Sub

' Testa se o servidor responde na porta 8002 (timeout de 500ms).
Private Function ServidorNoArWDO() As Boolean
    Dim http As Object
    On Error GoTo Fora
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 500, 500, 500, 500
    http.Open "GET", URL_SERVIDOR, False
    http.send
    ServidorNoArWDO = True
    Exit Function
Fora:
    ServidorNoArWDO = False
End Function

Public Sub PararExportWDO()
    exportando = False
    On Error Resume Next
    Application.OnTime proximaExecucao, "ExportarWDO", , False
End Sub

Public Sub ExportarWDO()
    If Not exportando Then Exit Sub

    Dim ws As Worksheet
    Dim linha As String
    Dim fnum As Integer
    Dim lin As Long

    On Error GoTo Reagendar
    Set ws = ThisWorkbook.Sheets(PLANILHA)

    lin = LinhaDoAtivoWDO(ws)
    If lin = 0 Then GoTo Reagendar      ' ativo nao encontrado na coluna A

    ' BLINDAGEM: cada coluna e localizada pela PROPRIA formula RTD da linha
    ' do ativo (casa "DOLFUT_F_0" + o codigo do campo), imune a deslocamento
    ' de colunas na planilha. O ultimo argumento e a coluna fixa de fallback
    ' (usada so se a formula nao for encontrada).
    Dim cUlt As Long, cAbe As Long, cMax As Long, cMin As Long
    Dim cVol As Long, cAgC As Long, cAgV As Long, cVwap As Long, cFec As Long
    cUlt = ColPorCampoWDO(ws, lin, ATIVO, "ULT", 4)
    cAbe = ColPorCampoWDO(ws, lin, ATIVO, "ABE", 5)
    cMax = ColPorCampoWDO(ws, lin, ATIVO, "MAX", 6)
    cMin = ColPorCampoWDO(ws, lin, ATIVO, "MIN", 7)
    cVol = ColPorCampoWDO(ws, lin, ATIVO, "VOL", 10)
    cAgC = ColPorCampoWDO(ws, lin, ATIVO, "98", 27)
    cAgV = ColPorCampoWDO(ws, lin, ATIVO, "99", 29)
    cVwap = ColPorCampoWDO(ws, lin, ATIVO, "67", 30)
    cFec = ColPorCampoWDO(ws, lin, ATIVO, "FEC", 8)

    With ws
        ' ultimo;abertura;maxima;minima;volume;agr_compra;agr_venda;vwap;fec_ant;timestamp
        linha = _
            NumBR(.Cells(lin, cUlt).Value) & ";" & _
            NumBR(.Cells(lin, cAbe).Value) & ";" & _
            NumBR(.Cells(lin, cMax).Value) & ";" & _
            NumBR(.Cells(lin, cMin).Value) & ";" & _
            NumBR(.Cells(lin, cVol).Value) & ";" & _
            NumBR(.Cells(lin, cAgC).Value) & ";" & _
            NumBR(.Cells(lin, cAgV).Value) & ";" & _
            NumBR(.Cells(lin, cVwap).Value) & ";" & _
            NumBR(.Cells(lin, cFec).Value) & ";" & _
            Format(Now, "hh:nn:ss")
    End With

    ' Sobrescreve o arquivo (o servidor le sempre a ultima linha)
    fnum = FreeFile
    Open CAMINHO_CSV For Output As #fnum
    Print #fnum, "ultimo;abertura;maxima;minima;volume;agr_compra;agr_venda;vwap;fec_ant;timestamp"
    Print #fnum, linha
    Close #fnum

    ' 22/09/2026: passou a rodar aqui tambem - ver comentario na declaracao
    ' de CAMINHO_CSV_MACRO acima (o Modulo1 escrevia esse CSV na pasta
    ' ERRADA para o WDO).
    Call ExportarMacroRTDWDO(ws)
    Call ExportarFluxo(ws)

    ciclosAcum = ciclosAcum + 1
    If ciclosAcum >= ACUM_A_CADA Then
        ciclosAcum = 0
        Call AtualizarAcumWDO
    End If

Reagendar:
    If exportando Then
        proximaExecucao = Now + TimeSerial(0, 0, INTERVALO_SEG)
        Application.OnTime proximaExecucao, "ExportarWDO"
    End If
End Sub

' Exporta DI futuros (DI1*), DOLFUT e WINFUTV (mini indice) para
' dados_macro_rtd.csv - ver comentario na declaracao de CAMINHO_CSV_MACRO
' acima. Colunas fixas (D=4 ultimo, H=8 fec_ant, J=10 volume), mesmo
' desenho simples do Modulo1 (sem ColPorCampoWDO): nunca derruba o export
' principal se a planilha nao tiver os tickers esperados.
Private Sub ExportarMacroRTDWDO(ws As Worksheet)
    Dim r As Long
    Dim fnum As Integer
    Dim tk As String
    Dim linhas As String
    On Error GoTo Fim
    For r = 1 To 100
        tk = Trim(CStr(ws.Cells(r, 1).Value))
        If tk Like "DI1*" Or tk = "DOLFUT" Or tk = "WINFUTV" Then
            linhas = linhas & tk & ";" & _
                NumBR(ws.Cells(r, 4).Value) & ";" & _
                NumBR(ws.Cells(r, 8).Value) & ";" & _
                NumBR(ws.Cells(r, 10).Value) & ";" & _
                Format(Now, "hh:nn:ss") & vbCrLf
        End If
    Next r
    If Len(linhas) = 0 Then Exit Sub
    fnum = FreeFile
    Open CAMINHO_CSV_MACRO For Output As #fnum
    Print #fnum, "ticker;ultimo;fec_ant;volume;timestamp"
    Print #fnum, linhas;
    Close #fnum
Fim:
End Sub

' ================= FLUXO: BOOK0 (livro) e T&T0 (fita) =================
' Exporta os dois blocos como CSV. Nunca derruba o export principal: se o
' bloco nao existir na planilha, sai em silencio.
Private Sub ExportarFluxo(ws As Worksheet)
    On Error GoTo Fim
    Call ExportarBook(ws)
    Call ExportarTT(ws)
    Call ExportarVAP(ws)
Fim:
End Sub

' VAP0 -> preco;volume   (perfil de volume por preco, do topo para baixo)
' ATENCAO: se o RTD trocar campos aqui como fez no VAP0 do WIN (conferido
' 21/07/2026: "VOL" devolvia o PRECO e "PRC" o VOLUME) confira contra o
' preco do DOLFUT antes de confiar no export - o tick aqui e de 0,5 ponto,
' nao 5 como no WIN.
Private Sub ExportarVAP(ws As Worksheet)
    Dim dados As Variant, linhas As String, i As Long, fnum As Integer
    If Not AncoraValida(ws, vapLin, vapCol, "VAP0", "VOL") Then
        If Not LocalizarAncora(ws, "VAP0", "VOL", vapLin, vapCol) Then Exit Sub
        vapN = 0
    End If
    If Not AlturaOk(ws, vapLin, vapCol, "VAP0", "VOL", vapN) Then _
        vapN = AlturaBloco(ws, vapLin, vapCol, "VAP0", "VOL")
    If vapN = 0 Then Exit Sub
    dados = ws.Range(ws.Cells(vapLin, vapCol), _
                     ws.Cells(vapLin + vapN - 1, vapCol + 1)).Value
    For i = 1 To vapN
        ' preco (campo VOL) ; volume (campo PRC, pode vir como "3,04k")
        linhas = linhas & NumBR(dados(i, 1)) & ";" & Txt(dados(i, 2)) & vbCrLf
    Next i
    fnum = FreeFile
    Open CAMINHO_CSV_VAP For Output As #fnum
    Print #fnum, "preco;volume"
    Print #fnum, linhas;
    Close #fnum
End Sub

' Quantas linhas seguidas, a partir da ancora, ainda sao do mesmo bloco RTD?
Private Function AlturaBloco(ws As Worksheet, linha As Long, coluna As Long, _
                             topico As String, campo As String) As Long
    Dim r As Long, alvo As String, f As String
    alvo = Chr(34) & topico & Chr(34) & "," & Chr(34) & campo & Chr(34) & ","
    For r = linha To linha + MAX_LINHAS_BLOCO - 1
        f = Replace(CStr(ws.Cells(r, coluna).Formula), " ", "")
        If InStr(1, f, alvo, vbTextCompare) = 0 Then Exit For
    Next r
    AlturaBloco = r - linha
End Function

' A altura em cache ainda bate? Checagem de 2 celulas (borda do bloco):
' a ultima linha do bloco tem que SER do topico e a seguinte NAO pode ser.
' Detecta extensao/encolhimento do bloco na planilha sem varredura completa.
Private Function AlturaOk(ws As Worksheet, linha As Long, coluna As Long, _
                          topico As String, campo As String, n As Long) As Boolean
    Dim alvo As String
    If n <= 0 Then Exit Function
    alvo = Chr(34) & topico & Chr(34) & "," & Chr(34) & campo & Chr(34) & ","
    If InStr(1, Replace(CStr(ws.Cells(linha + n - 1, coluna).Formula), " ", ""), _
             alvo, vbTextCompare) = 0 Then Exit Function
    If InStr(1, Replace(CStr(ws.Cells(linha + n, coluna).Formula), " ", ""), _
             alvo, vbTextCompare) > 0 Then Exit Function
    AlturaOk = True
End Function

' BOOK0 -> hora_c;agente_c;qtd_c;preco_c;preco_v;qtd_v;agente_v;hora_v
' Uma linha por nivel do livro (ofertas individuais, nao agregadas por preco).
Private Sub ExportarBook(ws As Worksheet)
    Dim dados As Variant, linhas As String, i As Long, fnum As Integer
    If Not AncoraValida(ws, bookLin, bookCol, "BOOK0", "HORC") Then
        If Not LocalizarAncora(ws, "BOOK0", "HORC", bookLin, bookCol) Then Exit Sub
        bookN = 0
    End If
    If Not AlturaOk(ws, bookLin, bookCol, "BOOK0", "HORC", bookN) Then _
        bookN = AlturaBloco(ws, bookLin, bookCol, "BOOK0", "HORC")
    If bookN = 0 Then Exit Sub
    ' 8 colunas a partir da ancora: HORC ACP VOC OCP OVD VOV AVD HORV
    dados = ws.Range(ws.Cells(bookLin, bookCol), _
                     ws.Cells(bookLin + bookN - 1, bookCol + 7)).Value
    For i = 1 To bookN
        linhas = linhas & _
            Txt(dados(i, 1)) & ";" & Txt(dados(i, 2)) & ";" & _
            NumBR(dados(i, 3)) & ";" & NumBR(dados(i, 4)) & ";" & _
            NumBR(dados(i, 5)) & ";" & NumBR(dados(i, 6)) & ";" & _
            Txt(dados(i, 7)) & ";" & Txt(dados(i, 8)) & vbCrLf
    Next i
    fnum = FreeFile
    Open CAMINHO_CSV_BOOK For Output As #fnum
    Print #fnum, "hora_c;agente_c;qtd_c;preco_c;preco_v;qtd_v;agente_v;hora_v"
    Print #fnum, linhas;
    Close #fnum
End Sub

' T&T0 -> hora;qtd;preco;comprador;vendedor   (indice 0 = negocio mais recente)
Private Sub ExportarTT(ws As Worksheet)
    Dim dados As Variant, linhas As String, i As Long, fnum As Integer
    If Not AncoraValida(ws, ttLin, ttCol, "T&T0", "DAT") Then
        If Not LocalizarAncora(ws, "T&T0", "DAT", ttLin, ttCol) Then Exit Sub
        ttN = 0
    End If
    If Not AlturaOk(ws, ttLin, ttCol, "T&T0", "DAT", ttN) Then _
        ttN = AlturaBloco(ws, ttLin, ttCol, "T&T0", "DAT")
    If ttN = 0 Then Exit Sub
    ' 5 colunas a partir da ancora: DAT QUL PRE ACP AVD
    dados = ws.Range(ws.Cells(ttLin, ttCol), _
                     ws.Cells(ttLin + ttN - 1, ttCol + 4)).Value
    For i = 1 To ttN
        linhas = linhas & _
            Txt(dados(i, 1)) & ";" & NumBR(dados(i, 2)) & ";" & _
            NumBR(dados(i, 3)) & ";" & Txt(dados(i, 4)) & ";" & _
            Txt(dados(i, 5)) & vbCrLf
    Next i
    fnum = FreeFile
    Open CAMINHO_CSV_TT For Output As #fnum
    Print #fnum, "hora;qtd;preco;comprador;vendedor"
    Print #fnum, linhas;
    Close #fnum
End Sub

' Cola o ranking acumulado (ranking_acum.csv, escrito pelo server) na aba
' AcumWDO. Publico para poder rodar via Alt+F8. Silencioso em qualquer
' falha: sem arquivo / sem aba / server desligado nao podem derrubar o export.
Public Sub AtualizarAcumWDO()
    Dim ws As Worksheet, fnum As Integer, linha As String
    Dim partes() As String, r As Long, diaArq As String
    On Error GoTo Fim
    If Dir(CAMINHO_CSV_ACUM) = "" Then Exit Sub
    Set ws = ThisWorkbook.Sheets(ABA_ACUM)

    fnum = FreeFile
    Open CAMINHO_CSV_ACUM For Input As #fnum
    r = 1
    Do While Not EOF(fnum)
        Line Input #fnum, linha
        linha = Trim(linha)
        If Left(linha, 6) = "# dia=" Then
            ' arquivo de outro pregao: limpa a Acum e nao cola nada
            diaArq = Mid(linha, 7)
            If diaArq <> Format(Date, "yyyy-mm-dd") Then
                ws.Range("A2:E500").ClearContents
                Close #fnum
                Exit Sub
            End If
        ElseIf InStr(linha, ";") > 0 And Left(linha, 9) <> "Corretora" Then
            partes = Split(linha, ";")
            If UBound(partes) >= 4 Then
                r = r + 1
                ws.Cells(r, 1).Value = partes(0)
                ws.Cells(r, 2).Value = Val(partes(1))
                ws.Cells(r, 3).Value = Val(partes(2))
                ws.Cells(r, 4).Value = Val(partes(3))
                ws.Cells(r, 5).Value = Val(partes(4))
            End If
        End If
    Loop
    Close #fnum
    ' remove sobras de uma lista anterior maior e carimba a atualizacao
    If r < 500 Then ws.Range(ws.Cells(r + 1, 1), ws.Cells(500, 5)).ClearContents
    ws.Range("G2").Value = Format(Now, "hh:nn:ss")
    Exit Sub
Fim:
    On Error Resume Next
    Close #fnum
End Sub

' A ancora em cache ainda aponta para a formula certa? (checagem de 1 celula)
Private Function AncoraValida(ws As Worksheet, linha As Long, coluna As Long, _
                              topico As String, campo As String) As Boolean
    If linha = 0 Or coluna = 0 Then Exit Function
    AncoraValida = (InStr(1, Replace(CStr(ws.Cells(linha, coluna).Formula), " ", ""), _
                          AlvoRTD(topico, campo), vbTextCompare) > 0)
End Function

' Varre a planilha atras da formula do indice 0 do bloco. So roda quando o
' cache esta vazio ou invalido.
Private Function LocalizarAncora(ws As Worksheet, topico As String, _
                                 campo As String, ByRef linha As Long, _
                                 ByRef coluna As Long) As Boolean
    Dim r As Long, c As Long, alvo As String, f As String
    alvo = AlvoRTD(topico, campo)
    For r = 1 To 120
        For c = 1 To 60
            f = Replace(CStr(ws.Cells(r, c).Formula), " ", "")
            If InStr(1, f, alvo, vbTextCompare) > 0 Then
                linha = r: coluna = c
                LocalizarAncora = True
                Exit Function
            End If
        Next c
    Next r
    linha = 0: coluna = 0
End Function

' Assinatura do indice 0 de um bloco: "TOPICO","CAMPO",0)
' (o indice vem SEM aspas nessas formulas, diferente dos campos do ativo)
Private Function AlvoRTD(topico As String, campo As String) As String
    AlvoRTD = Chr(34) & topico & Chr(34) & "," & Chr(34) & campo & Chr(34) & ",0)"
End Function

' Texto limpo para CSV (remove o separador e quebras de linha).
Private Function Txt(v As Variant) As String
    Dim s As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    s = Trim(CStr(v))
    s = Replace(Replace(Replace(s, ";", ","), vbCr, ""), vbLf, "")
    Txt = s
End Function

' Localiza a linha do ativo na coluna A (0 se nao encontrado).
' Busca nas primeiras 100 linhas - suficiente para a tabela de ativos.
Private Function LinhaDoAtivoWDO(ws As Worksheet) As Long
    LinhaDoAtivoWDO = LinhaDoTickerWDO(ws, ATIVO)
End Function

Private Function LinhaDoTickerWDO(ws As Worksheet, ticker As String) As Long
    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = ticker Then
            LinhaDoTickerWDO = r
            Exit Function
        End If
    Next r
    LinhaDoTickerWDO = 0
End Function

' Localiza a coluna de um campo RTD na linha do ativo pela PROPRIA formula
' (imune a deslocamento de colunas). Casa ticker + codigo do campo, ex.:
'   =RTD("rtdtrading.rtdserver",, "DOLFUT_F_0", "67")  -> campo "67" (VWAP)
' Ignora espacos na formula (tolerante a variacoes de digitacao). Se nao
' encontrar, devolve 'fallback' (a coluna fixa antiga, como ultimo recurso).
Private Function ColPorCampoWDO(ws As Worksheet, linha As Long, _
                             ticker As String, campo As String, _
                             fallback As Long) As Long
    Dim c As Long, f As String, alvo As String
    ' alvo (sem espacos): "TICKER_F_0","CAMPO")
    alvo = Chr(34) & ticker & "_F_0" & Chr(34) & "," & Chr(34) & campo & Chr(34) & ")"
    For c = 1 To 256
        f = Replace(CStr(ws.Cells(linha, c).Formula), " ", "")
        If InStr(1, f, alvo, vbTextCompare) > 0 Then
            ColPorCampoWDO = c
            Exit Function
        End If
    Next c
    ColPorCampoWDO = fallback
End Function

' Converte para string numerica com ponto decimal (formato US),
' evitando ambiguidade de locale no Python.
Private Function NumBR(v As Variant) As String
    If IsNumeric(v) Then
        NumBR = Replace(CStr(CDbl(v)), ",", ".")
    Else
        NumBR = ""
    End If
End Function




