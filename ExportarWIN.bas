' ============================================================
'  ExportarWIN.bas - Modulo VBA para o Monitor WIN
' ------------------------------------------------------------
'  Linha do WINFUTV na planilha DADOS: localizada AUTOMATICAMENTE
'  pela coluna A (imune a insercao/remocao de linhas). Colunas:
'    D(4)=ultimo  E(5)=abertura  F(6)=maxima  G(7)=minima
'    H(8)=fec_ant["FEC"]     X(24)=agr_compra["98"]
'    Z(26)=agr_venda["99"]   AA(27)=vwap["67"]  AB(28)=volume["VOL"]
'
'  Formula RTD de exemplo (celula D14, locale BR):
'    =RTD("rtdtrading.rtdserver";;"WINFUT_F_0";"2")
'    (ajuste o campo numerico conforme sua tabela de campos)
'
'  INICIO AUTOMATICO: o evento Workbook_Open (em EstaPastaDeTrabalho)
'  chama IniciarExportWIN 10s apos abrir a planilha. Rodar manualmente
'  via Alt+F8 so e necessario se as macros forem bloqueadas na abertura.
'  O CSV e gravado na pasta do monitor: ajuste CAMINHO_CSV abaixo.
'
'  SERVIDOR: IniciarExportWIN tambem sobe o server_win.py em segundo
'  plano (sem janela), caso a porta 8001 ainda nao esteja respondendo.
'  Um server ja no ar NUNCA e reiniciado (preserva historico intradiario).
' ============================================================
Option Explicit

Private Const PASTA_MONITOR As String = "C:\Users\rodri\OneDrive\Apps\monitor_win"
Private Const CAMINHO_CSV As String = "C:\Users\rodri\OneDrive\Apps\monitor_win\dados_win.csv"
Private Const URL_SERVIDOR As String = "http://127.0.0.1:8001/ultimo"
Private Const PLANILHA As String = "DADOS"
Private Const ATIVO As String = "WINFUTV"   ' <== procurado na coluna A
Private Const INTERVALO_SEG As Long = 2

Private proximaExecucao As Date
Private exportando As Boolean

Public Sub IniciarExportWIN()
    Call IniciarServidorWIN
    exportando = True
    Call ExportarWIN
End Sub

' Sobe o server_win.py em segundo plano se a porta 8001 nao responder.
' Nao mexe em um server que ja esteja no ar.
Public Sub IniciarServidorWIN()
    If ServidorNoAr() Then Exit Sub

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.CurrentDirectory = PASTA_MONITOR

    On Error Resume Next
    ' pythonw = roda sem janela de console
    sh.Run "pythonw.exe """ & PASTA_MONITOR & "\server_win.py""", 0, False
    If Err.Number <> 0 Then
        Err.Clear
        ' fallback caso pythonw nao esteja no PATH (janela oculta)
        sh.Run "python.exe """ & PASTA_MONITOR & "\server_win.py""", 0, False
    End If
    On Error GoTo 0
End Sub

' Testa se o servidor responde na porta 8001 (timeout de 500ms).
Private Function ServidorNoAr() As Boolean
    Dim http As Object
    On Error GoTo Fora
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 500, 500, 500, 500
    http.Open "GET", URL_SERVIDOR, False
    http.send
    ServidorNoAr = True
    Exit Function
Fora:
    ServidorNoAr = False
End Function

Public Sub PararExportWIN()
    exportando = False
    On Error Resume Next
    Application.OnTime proximaExecucao, "ExportarWIN", , False
End Sub

Public Sub ExportarWIN()
    If Not exportando Then Exit Sub

    Dim ws As Worksheet
    Dim linha As String
    Dim fnum As Integer
    Dim lin As Long

    On Error GoTo Reagendar
    Set ws = ThisWorkbook.Sheets(PLANILHA)

    lin = LinhaDoAtivo(ws)
    If lin = 0 Then GoTo Reagendar      ' ativo nao encontrado na coluna A

    With ws
        ' ultimo;abertura;maxima;minima;volume;agr_compra;agr_venda;vwap;fec_ant;timestamp
        linha = _
            NumBR(.Cells(lin, 4).Value) & ";" & _
            NumBR(.Cells(lin, 5).Value) & ";" & _
            NumBR(.Cells(lin, 6).Value) & ";" & _
            NumBR(.Cells(lin, 7).Value) & ";" & _
            NumBR(.Cells(lin, 28).Value) & ";" & _
            NumBR(.Cells(lin, 24).Value) & ";" & _
            NumBR(.Cells(lin, 26).Value) & ";" & _
            NumBR(.Cells(lin, 27).Value) & ";" & _
            NumBR(.Cells(lin, 8).Value) & ";" & _
            Format(Now, "hh:nn:ss")
    End With

    ' Sobrescreve o arquivo (o servidor le sempre a ultima linha)
    fnum = FreeFile
    Open CAMINHO_CSV For Output As #fnum
    Print #fnum, "ultimo;abertura;maxima;minima;volume;agr_compra;agr_venda;vwap;fec_ant;timestamp"
    Print #fnum, linha
    Close #fnum

Reagendar:
    If exportando Then
        proximaExecucao = Now + TimeSerial(0, 0, INTERVALO_SEG)
        Application.OnTime proximaExecucao, "ExportarWIN"
    End If
End Sub

' Localiza a linha do ativo na coluna A (0 se nao encontrado).
' Busca nas primeiras 100 linhas — suficiente para a tabela de ativos.
Private Function LinhaDoAtivo(ws As Worksheet) As Long
    Dim r As Long
    For r = 1 To 100
        If Trim(CStr(ws.Cells(r, 1).Value)) = ATIVO Then
            LinhaDoAtivo = r
            Exit Function
        End If
    Next r
    LinhaDoAtivo = 0
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
