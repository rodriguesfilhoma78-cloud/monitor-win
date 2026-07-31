Attribute VB_Name = "ModuloBlueChips"
' ============================================================
'  ExportarBlueChips.bas - Modulo VBA para o Painel Blue Chips
' ------------------------------------------------------------
'  Linha de cada ativo localizada AUTOMATICAMENTE pela coluna A.
'  Nao usa porta propria - o CSV alimenta direto o server_win.py
'  (porta 8001), que ja esta no ar.
'
'  BLINDAGEM (desde 31/07/2026, portada do Modulo1/WIN): as colunas NAO sao
'  mais fixas. ColPorCampo localiza cada campo pela PROPRIA formula RTD da
'  linha do ativo (casa "TICKER_B_0" + o codigo do campo), entao um
'  deslocamento de colunas na planilha (ex.: novos indicadores inseridos no
'  meio) nao quebra mais o export silenciosamente. Codigos usados:
'    ULT=ultimo ABE=abertura MAX=maxima MIN=minima FEC=fec_ant VOL=volume
'    98=agr_compra  99=agr_venda  67=vwap
'  Fallback (se a formula sumir): colunas fixas atuais
'    ultimo4 abertura5 maxima6 minima7 fec_ant8 agr_compra27 agr_venda29
'    vwap30 volume10
'  Contexto: esse export ja quebrou 2x por indice fixo (30/07 e 31/07/2026,
'  ver CHANGELOG.md) - a segunda vez foi a MESMA correcao de colunas do dia
'  anterior nao ter sido reimportada no workbook aberto. A blindagem existe
'  pra essa classe de bug (deslocamento de coluna) parar de silenciosamente
'  quebrar o export, independente de o codigo estar ou nao sincronizado.
'
'  Formula RTD de exemplo (linha da VALE3, locale BR):
'    =RTD("rtdtrading.rtdserver";;"VALE3_B_0";"98")
'
'  INICIO AUTOMATICO: adicionar ao Workbook_Open junto com os demais:
'    Application.OnTime Now + TimeSerial(0, 0, 16), "IniciarExportBlueChips"
' ============================================================
Option Explicit

Private Const CAMINHO_CSV As String = "C:\Users\rodri\Desktop\Day trade\monitor_win\dados_blue_chips.csv"
Private Const PLANILHA As String = "DADOS"
Private Const INTERVALO_SEG As Long = 2
Private Const LISTA_TICKERS As String = "VALE3,PETR4,ITUB4,BBDC4,BBAS3"
Private Const SUFIXO_RTD As String = "_B_0"

Private proximaExecucao As Date
Private exportando As Boolean

Public Sub IniciarExportBlueChips()
    exportando = True
    Call ExportarBlueChips
End Sub

Public Sub PararExportBlueChips()
    exportando = False
    On Error Resume Next
    Application.OnTime proximaExecucao, "ExportarBlueChips", , False
End Sub

Public Sub ExportarBlueChips()
    If Not exportando Then Exit Sub

    Dim ws As Worksheet
    Dim fnum As Integer
    Dim ts As String
    Dim arrTickers() As String
    Dim i As Long

    On Error GoTo Reagendar
    Set ws = ThisWorkbook.Sheets(PLANILHA)
    ts = Format(Now, "hh:nn:ss")
    arrTickers = Split(LISTA_TICKERS, ",")

    fnum = FreeFile
    Open CAMINHO_CSV For Output As #fnum
    Print #fnum, "ticker;ultimo;abertura;maxima;minima;fec_ant;agr_compra;agr_venda;vwap;volume;timestamp"
    For i = LBound(arrTickers) To UBound(arrTickers)
        Print #fnum, LinhaAtivo(ws, arrTickers(i), ts)
    Next i
    Close #fnum

Reagendar:
    If exportando Then
        proximaExecucao = Now + TimeSerial(0, 0, INTERVALO_SEG)
        Application.OnTime proximaExecucao, "ExportarBlueChips"
    End If
End Sub

Private Function LinhaAtivo(ws As Worksheet, ticker As String, ts As String) As String
    Dim lin As Long
    lin = LinhaDoAtivo(ws, ticker)
    If lin = 0 Then
        LinhaAtivo = ticker & ";;;;;;;;;;" & ts
        Exit Function
    End If

    ' BLINDAGEM: cada coluna e localizada pela PROPRIA formula RTD da linha
    ' do ativo (ver ColPorCampo). O ultimo argumento e a coluna fixa de
    ' fallback, usada so se a formula nao for encontrada.
    Dim cUlt As Long, cAbe As Long, cMax As Long, cMin As Long, cFec As Long
    Dim cAgC As Long, cAgV As Long, cVwap As Long, cVol As Long
    cUlt = ColPorCampo(ws, lin, ticker, "ULT", 4)
    cAbe = ColPorCampo(ws, lin, ticker, "ABE", 5)
    cMax = ColPorCampo(ws, lin, ticker, "MAX", 6)
    cMin = ColPorCampo(ws, lin, ticker, "MIN", 7)
    cFec = ColPorCampo(ws, lin, ticker, "FEC", 8)
    cAgC = ColPorCampo(ws, lin, ticker, "98", 27)
    cAgV = ColPorCampo(ws, lin, ticker, "99", 29)
    cVwap = ColPorCampo(ws, lin, ticker, "67", 30)
    cVol = ColPorCampo(ws, lin, ticker, "VOL", 10)

    LinhaAtivo = ticker & ";" & _
        NumBR(ws.Cells(lin, cUlt).Value) & ";" & _
        NumBR(ws.Cells(lin, cAbe).Value) & ";" & _
        NumBR(ws.Cells(lin, cMax).Value) & ";" & _
        NumBR(ws.Cells(lin, cMin).Value) & ";" & _
        NumBR(ws.Cells(lin, cFec).Value) & ";" & _
        NumBR(ws.Cells(lin, cAgC).Value) & ";" & _
        NumBR(ws.Cells(lin, cAgV).Value) & ";" & _
        NumBR(ws.Cells(lin, cVwap).Value) & ";" & _
        NumBR(ws.Cells(lin, cVol).Value) & ";" & ts
End Function

Private Function LinhaDoAtivo(ws As Worksheet, ticker As String) As Long
    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = ticker Then
            LinhaDoAtivo = r
            Exit Function
        End If
    Next r
    LinhaDoAtivo = 0
End Function

' Localiza a coluna de um campo RTD na linha do ativo pela PROPRIA formula
' (imune a deslocamento de colunas). Casa ticker + sufixo RTD + codigo do
' campo, ex.: =RTD("rtdtrading.rtdserver",, "VALE3_B_0", "98") -> campo "98"
' (agr_compra). Ignora espacos na formula. Se nao encontrar, devolve
' 'fallback' (a coluna fixa atual, como ultimo recurso).
Private Function ColPorCampo(ws As Worksheet, linha As Long, _
                             ticker As String, campo As String, _
                             fallback As Long) As Long
    Dim c As Long, f As String, alvo As String
    alvo = Chr(34) & ticker & SUFIXO_RTD & Chr(34) & "," & Chr(34) & campo & Chr(34) & ")"
    For c = 1 To 256
        f = Replace(CStr(ws.Cells(linha, c).Formula), " ", "")
        If InStr(1, f, alvo, vbTextCompare) > 0 Then
            ColPorCampo = c
            Exit Function
        End If
    Next c
    ColPorCampo = fallback
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
