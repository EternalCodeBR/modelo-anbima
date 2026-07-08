"""Oráculo: abre a calculadora no Excel, recalcula e lê a célula de PU.
Reusa a abordagem do protótipo COM do usuário (não roda na rotina; só verificação).

Sem teste unitário — depende de Excel instalado (pywin32). Validado manualmente.

Nota (Fase 0): linhas projetadas no futuro de algumas calculadoras erram (#VALUE!,
devolvido por COM como ~-2.147e9). Por isso a reconciliação lê o PU numa **data-alvo**
(D-1, com valor limpo), não a "última linha preenchida", e ignora células de erro.

LIMITAÇÃO CONHECIDA (a resolver antes do modo-sombra diário): num `CalculateFull` em
instância COM fresca, o **link externo** `[1]SERIE_VALOR` (BaseDadosMercado) não resolve →
o SUMIFS erra → o PU vira #VALUE! em todas as linhas. Tentar pré-abrir a BaseDadosMercado
na mesma instância falhou aqui porque o Excel COM não acessa o caminho sincronizado de
SharePoint/OneDrive (openpyxl lê, COM bloqueia). Caminhos a investigar: (a) apontar o link
para uma cópia local/UNC acessível ao COM; (b) `UpdateLinks` + workbook da base pré-aberto;
(c) injetar a curva via `SetLinkOnData`/valores. **A prova do centavo da Fase 1 não depende
disto** — usa o oráculo em cache (openpyxl), que reflete o que a calculadora exibe.
"""
from datetime import date
from pathlib import Path

# códigos de erro de célula que o COM devolve como inteiros grandes negativos
_ERRO_COM = -2146820000  # qualquer valor <= este é sentinela de erro do Excel


def _valido(v) -> bool:
    return isinstance(v, (int, float)) and v > _ERRO_COM


def pu_via_excel(caminho_xlsx: str, aba: str, coluna_pu: str,
                 data_ref: date | None = None, coluna_data: str = "C") -> float:
    """Recalcula a planilha e devolve o PU. Se `data_ref` for dado, lê o PU na linha
    cuja coluna de data bate com `data_ref`; senão, a última linha com PU numérico válido."""
    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(Path(caminho_xlsx)), UpdateLinks=0, ReadOnly=True)
        wb.Application.CalculateFull()
        ws = wb.Worksheets(aba)
        col_pu = ws.Range(f"{coluna_pu}1").Column
        col_dt = ws.Range(f"{coluna_data}1").Column
        ult = ws.Cells(ws.Rows.Count, col_dt).End(-4162).Row  # xlUp na coluna de data

        if data_ref is not None:
            for r in range(2, ult + 1):
                v = ws.Cells(r, col_dt).Value
                d = v.date() if hasattr(v, "date") else v
                if d == data_ref:
                    pu = ws.Cells(r, col_pu).Value
                    if not _valido(pu):
                        raise ValueError(f"PU em {data_ref} é erro/vazio ({pu})")
                    return float(pu)
            raise ValueError(f"data {data_ref} não encontrada na coluna {coluna_data}")

        # sem data: última linha com PU numérico válido (pula erros das projeções)
        for r in range(ult, 1, -1):
            pu = ws.Cells(r, col_pu).Value
            if _valido(pu):
                return float(pu)
        raise ValueError("nenhum PU numérico válido encontrado")
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()
