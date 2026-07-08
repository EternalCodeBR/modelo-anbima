' ============================================================================
'  COLE ESTE CÓDIGO NO MÓDULO "ThisWorkbook" do Rotina_Recalc.xlsm
'  (VBE > painel esquerdo > EstaPastaDeTrabalho / ThisWorkbook), NÃO num módulo comum.
'
'  Faz o Excel rodar o RecalcCacheLocal SOZINHO quando o Python o abre pela rotina.
'  Comunicação por arquivos (sem win32com/COM, sem antivírus):
'    - o Python cria  C:\PU_Recalc\_run.flag  e abre este .xlsm
'    - ao abrir, se a flag existe: apaga a flag, recalcula, cria _done.marker e fecha
'    - se a flag NÃO existe (abertura manual normal): não faz nada
'
'  PRÉ-REQUISITO: a pasta do Rotina_Recalc.xlsm deve estar como LOCAL CONFIÁVEL
'  (Central de Confiabilidade > Locais Confiáveis), senão o Excel bloqueia a macro
'  ao abrir e o Workbook_Open não roda.
' ============================================================================
Private Sub Workbook_Open()
    Const FLAG As String = "C:\PU_Recalc\_run.flag"
    Const DONE As String = "C:\PU_Recalc\_done.marker"
    Dim f As Integer

    If Dir(FLAG) <> "" Then
        Kill FLAG                       ' consome o gatilho (não repete)
        RecalcCacheLocal True           ' recalcula em modo silencioso (sem MsgBox)
        f = FreeFile
        Open DONE For Output As #f       ' avisa o Python que terminou
        Print #f, "ok"
        Close #f
        ThisWorkbook.Saved = True        ' evita prompt de "salvar?" ao sair
        Application.Quit                 ' fecha o Excel
    End If
End Sub
