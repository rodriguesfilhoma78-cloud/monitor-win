Attribute VB_Name = "ModuloBlueChips"
' ============================================================
'  ExportarBlueChips.bas - Modulo VBA para o Painel Blue Chips
' ------------------------------------------------------------
'  Linha de cada ativo localizada AUTOMATICAMENTE pela coluna A.
'  Nao usa porta propria - o CSV alimenta direto o server_win.py
'  (porta 8001), que ja esta no ar.
'  Colunas RTD: D=ultimo E=abertura F=maxima G=minima H=fec_ant
'    AA(27)=agr_compra [TR Volume de Agressao - Compra]
'    AC(29)=agr_venda [TR Volume de Agressao - Venda]
'    AD(30)=vwap [VWAP]   J(10)=volume [Volume]
'  (corrigido 30/07: colunas 24/26/27/28 antigas passaram a apontar
'  para Prior Cote/Saldo Agressao apos a planilha ganhar colunas de
'  MACD - o fluxo de compra/venda saia sempre "venda" por causa disso)
'
'  INICIO AUTOMATICO: adicionar ao Workbook_Open junto com os demais:
'    Application.OnTime Now + TimeSerial(0, 0, 16), "IniciarExportBlueChips"
' ============================================================
Option Explicit

Private Const CAMINHO_CSV As String = "C:\Users\rodri\Desktop\Day trade\monitor_win\dados_blue_chips.csv"
Private Const PLANILHA As String = "DADOS"
Private Const INTERVALO_SEG As Long = 2
Private Const LISTA_TICKERS As String = "VALE3,PETR4,ITUB4,BBDC4,BBAS3"

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
    LinhaAtivo = ticker & ";" & _
        NumBR(ws.Cells(lin, 4).Value) & ";" & _
        NumBR(ws.Cells(lin, 5).Value) & ";" & _
        NumBR(ws.Cells(lin, 6).Value) & ";" & _
        NumBR(ws.Cells(lin, 7).Value) & ";" & _
        NumBR(ws.Cells(lin, 8).Value) & ";" & _
        NumBR(ws.Cells(lin, 27).Value) & ";" & _
        NumBR(ws.Cells(lin, 29).Value) & ";" & _
        NumBR(ws.Cells(lin, 30).Value) & ";" & _
        NumBR(ws.Cells(lin, 10).Value) & ";" & ts
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

Private Function NumBR(v As Variant) As String
    If IsNumeric(v) Then
        NumBR = Replace(CStr(CDbl(v)), ",", ".")
    Else
        NumBR = ""
    End If
End Function

