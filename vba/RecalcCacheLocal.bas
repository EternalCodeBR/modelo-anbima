Attribute VB_Name = "modRecalc"
' ============================================================================
'  Recalcula o cache de PU das calculadoras FORA do OneDrive (AutoSave é o gargalo).
'  Copia cada calculadora p/ C:\PU_Recalc (local), recalcula/salva lá (~1s), e devolve
'  ao caminho original. Lê os caminhos da aba "Lista" (coluna A, a partir da linha 2).
'
'  DOIS pontos de entrada SEM parâmetro (por isso aparecem no Alt+F8 e no Application.Run):
'    RecalcCacheLocal      -> manual: mostra resumo em MsgBox
'    RecalcCacheLocalAuto  -> automático (COM/rodar_rotina.py): silencioso, grava log
'
'  IMPORTANTE: Subs COM parâmetro NÃO aparecem na caixa de Macros nem são chamáveis
'  direto pelo Application.Run — por isso o núcleo fica num Sub PRIVADO (RecalcCore) e
'  os dois pontos de entrada são parameterless. O nome do MÓDULO (modRecalc) é diferente
'  dos Subs para não haver ambiguidade no Application.Run.
' ============================================================================
Option Explicit

' Ponto de entrada MANUAL (Alt+F8) — mostra o resumo em MsgBox.
Public Sub RecalcCacheLocal()
    RecalcCore False
End Sub

' Ponto de entrada AUTOMÁTICO (COM) — silencioso; grava %TEMP%\recalc_log.txt.
Public Sub RecalcCacheLocalAuto()
    RecalcCore True
End Sub

' Núcleo. silent=True -> sem MsgBox, grava o resultado num log que o Python lê.
Private Sub RecalcCore(ByVal silent As Boolean)
    Const TMP As String = "C:\PU_Recalc\"
    Dim ws As Worksheet
    Dim ultima As Long, r As Long
    Dim origem As String, nomeArq As String, arqLocal As String
    Dim wb As Workbook
    Dim ok As Long, erros As Long
    Dim t0 As Double, tArq As Double, total As Double
    Dim maisLento As Double, arqLento As String
    Dim salvou As Boolean, copiou As Boolean
    Dim calcAntigo As XlCalculation
    Dim logF As Integer

    calcAntigo = Application.Calculation
    If Dir(TMP, vbDirectory) = "" Then MkDir TMP

    Set ws = ThisWorkbook.Worksheets("Lista")
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual

    t0 = Timer
    For r = 2 To ultima
        origem = Trim(ws.Cells(r, 1).Value)
        If Len(origem) > 0 Then
            If Dir(origem) = "" Then
                erros = erros + 1
            Else
                nomeArq = Mid$(origem, InStrRev(origem, "\") + 1)
                arqLocal = TMP & r & "_" & nomeArq
                tArq = Timer
                Application.StatusBar = "RecalcLocal " & (r - 1) & "/" & (ultima - 1) & _
                    " | " & nomeArq & " | " & Format(Timer - t0, "0.0") & "s"

                copiou = False
                On Error Resume Next
                If Dir(arqLocal) <> "" Then Kill arqLocal
                FileCopy origem, arqLocal
                copiou = (Err.Number = 0)
                Err.Clear
                On Error GoTo 0

                If Not copiou Then
                    erros = erros + 1
                Else
                    Set wb = Nothing
                    On Error Resume Next
                    Set wb = Workbooks.Open(Filename:=arqLocal, UpdateLinks:=0, ReadOnly:=False)
                    On Error GoTo 0

                    If wb Is Nothing Then
                        erros = erros + 1
                    Else
                        Application.CalculateFull
                        salvou = False
                        On Error Resume Next
                        wb.Save
                        salvou = (Err.Number = 0)
                        Err.Clear
                        wb.Close SaveChanges:=False
                        On Error GoTo 0

                        If salvou Then
                            On Error Resume Next
                            FileCopy arqLocal, origem
                            If Err.Number = 0 Then
                                ok = ok + 1
                            Else
                                erros = erros + 1
                            End If
                            Err.Clear
                            On Error GoTo 0
                        Else
                            erros = erros + 1
                        End If
                    End If

                    On Error Resume Next
                    If Dir(arqLocal) <> "" Then Kill arqLocal
                    On Error GoTo 0
                End If

                tArq = Timer - tArq
                If tArq < 0 Then tArq = tArq + 86400
                If tArq > maisLento Then
                    maisLento = tArq
                    arqLento = nomeArq
                End If
            End If
        End If
    Next r

    total = Timer - t0
    If total < 0 Then total = total + 86400

    Application.StatusBar = False
    Application.Calculation = calcAntigo
    Application.AskToUpdateLinks = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If silent Then
        On Error Resume Next
        logF = FreeFile
        Open Environ$("TEMP") & "\recalc_log.txt" For Output As #logF
        Print #logF, "OK=" & ok & " | erros=" & erros & _
              " | total=" & Format(total, "0.0") & "s" & _
              " | lento=" & Format(maisLento, "0.0") & "s (" & arqLento & ")"
        Close #logF
        On Error GoTo 0
    Else
        MsgBox "RecalcCacheLocal concluido" & vbCrLf & vbCrLf & _
               "OK: " & ok & " | erros: " & erros & vbCrLf & _
               "Tempo total: " & Format(total, "0.0") & " s (" & Format(total / 60, "0.0") & " min)" & vbCrLf & _
               "Mais lento: " & Format(maisLento, "0.0") & " s (" & arqLento & ")", _
               vbInformation, "RecalcCacheLocal"
    End If
End Sub
