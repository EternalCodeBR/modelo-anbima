# pu_mtm/dados/cadastro.py
"""Leitura do cadastro_ativos.csv e do livro de eventos por ativo.

Lê via `csvio` (detecta separador ',' ou ';' e aceita datas BR/ISO) porque os CSVs
do projeto passam por limpeza no Excel pt-BR. Ver csvio.py.

ler_eventos prefere {id_serie}.xlsx (novo formato com Previsto/Realizado) e cai de
volta para {id_serie}.csv (formato legado) se o xlsx não existir."""
from pathlib import Path
from pu_mtm.dados import csvio
from pu_mtm.dominio.modelos import Ativo, Evento

def _data(s: str):
    return csvio.parse_data(s)  # ISO ou DD/MM/AAAA

def _float(s: str) -> float:
    return float(s) if s not in (None, "") else 0.0

def _int(s: str, padrao: int) -> int:
    return int(float(s)) if s not in (None, "") else padrao

def _str(s, padrao: str) -> str:
    return s.strip() if isinstance(s, str) and s.strip() else padrao

def ler_cadastro(caminho_csv: str) -> dict[str, Ativo]:
    out = {}
    for r in csvio.ler_dict(caminho_csv):
            out[r["IdSerie"]] = Ativo(
                id_serie=r["IdSerie"], apelido=r["Apelido"], familia=r["Familia"],
                vne=_float(r["VNe"]), data_emissao=_data(r["DataEmissao"]),
                spread=_float(r.get("Spread", "")), taxa_fixa=_float(r.get("TaxaFixa", "")),
                percentual_cdi=(_float(r.get("PercentualCDI", "")) or 100.0),
                # parâmetros de convenção (colunas opcionais; usam o default do Ativo se ausentes)
                base=_int(r.get("Base", ""), 252),
                fator_diario_arred=_str(r.get("FatorDiarioArred"), "nenhum"),
                juros_arred=_str(r.get("JurosArred"), "trunc8"),
                evjuros_fonte=_str(r.get("EvJurosFonte"), "nenhum"),
                evamort_fonte=_str(r.get("EvAmortFonte"), "nenhum"),
                pre_dc=_int(r.get("PreDC", ""), 30),
                pre_daycount=_str(r.get("PreDaycount"), "30360"))
    return out

_FAMILIA_PASTA = {
    "di_puro":    "CDI",
    "di_spread":  "CDI + SPREAD",
    "prefixado":  "PREFIXADO",
    "ipca_spread":"IPCA",
}

def _parse_data_opt(s: str) -> "date | None":
    """Retorna date ou None se vazio."""
    from datetime import date as _date
    s = (s or "").strip()
    return csvio.parse_data(s) if s else None


def _ler_xlsx(p: "Path") -> list[Evento]:
    """Lê planilha de eventos (formato novo: Previsto/Realizado)."""
    import openpyxl
    wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
    ws = wb["Eventos"]
    rows = ws.iter_rows(values_only=True)
    header = [str(c).strip() if c is not None else "" for c in next(rows)]

    def _col(row, nome):
        try:
            v = row[header.index(nome)]
        except (ValueError, IndexError):
            return None
        return v

    def _str(v) -> str:
        return str(v).strip() if v is not None else ""

    def _fv(v) -> float:
        if v in (None, ""):
            return 0.0
        try:
            return float(v)
        except (ValueError, TypeError):
            return 0.0

    def _dv(v) -> "date | None":
        if v is None:
            return None
        from datetime import date as _date, datetime as _dt
        if isinstance(v, _date) and not isinstance(v, _dt):
            return v
        if isinstance(v, _dt):
            return v.date()
        return _parse_data_opt(str(v))

    evs = []
    for row in rows:
        data_prev = _dv(_col(row, "DataPrevista"))
        data_real = _dv(_col(row, "DataRealizada"))
        if data_prev is None:
            continue                             # linha vazia
        j_prev = _fv(_col(row, "JurosPrevisto"))
        j_real = _col(row, "JurosRealizado")
        a_prev = _fv(_col(row, "AmortizacaoPrevista"))
        a_real = _col(row, "AmortizacaoRealizada")
        tipo = _str(_col(row, "Tipo")) or "agendado"
        obs  = _str(_col(row, "Observacao"))

        # data e valores efetivos: realizado se preenchido, senão previsto
        data_ef = data_real if data_real is not None else data_prev
        j_ef    = _fv(j_real) if j_real not in (None, "") else j_prev
        a_ef    = _fv(a_real) if a_real not in (None, "") else a_prev

        evs.append(Evento(
            data=data_ef,
            evento_juros=j_ef,
            evento_amortizacao=a_ef,
            data_prevista=data_prev if data_real is not None else None,
            tipo=tipo,
            obs=obs,
        ))
    wb.close()
    return evs


def _ler_csv(p: "Path") -> list[Evento]:
    """Lê CSV legado (Data/EventoJuros/EventoAmortizacao)."""
    evs = []
    for r in csvio.ler_dict(str(p)):
        evs.append(Evento(data=_data(r["Data"]),
                          evento_juros=_float(r["EventoJuros"]),
                          evento_amortizacao=_float(r["EventoAmortizacao"])))
    return evs


def ler_eventos(dir_eventos: str, id_serie: str, familia: str = "") -> list[Evento]:
    base = Path(dir_eventos)
    pasta = _FAMILIA_PASTA.get(familia, "")
    pasta_ativo = (base / pasta / id_serie) if pasta else (base / id_serie)

    xlsx = pasta_ativo / f"{id_serie}.xlsx"
    if xlsx.exists():
        return _ler_xlsx(xlsx)

    csv = pasta_ativo / f"{id_serie}.csv"
    if csv.exists():
        return _ler_csv(csv)

    return []
