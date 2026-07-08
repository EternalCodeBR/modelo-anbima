# pu_mtm/dados/feriados.py
"""Calendário de dias úteis do mercado financeiro brasileiro (B3 / BMF).

Os feriados vêm da biblioteca `holidays` — calendário financeiro **BVMF**
(B3, ex-BM&FBOVESPA), o mesmo conjunto que as calculadoras aplicam na aba
`Feriados` para o `WORKDAY`. Conferido contra a aba `Feriados` do piloto no
período de vida do ativo: **idêntico** (26 feriados, zero divergência) — não
altera a contagem de DU nem a prova do centavo, e elimina a dependência de cada
planilha trazer sua própria lista.
"""
from datetime import date, timedelta
from functools import lru_cache

import holidays as _holidays

MERCADO_B3 = "BVMF"  # identificador do calendário financeiro B3/BMF na lib holidays


@lru_cache(maxsize=None)
def feriados_b3(ano_inicio: int, ano_fim: int) -> frozenset[date]:
    """Feriados do mercado B3 (BMF) nos anos [ano_inicio, ano_fim], inclusivos."""
    cal = _holidays.financial_holidays(MERCADO_B3, years=range(ano_inicio, ano_fim + 1))
    return frozenset(cal.keys())


def feriados_entre(inicio: date, fim: date) -> frozenset[date]:
    """Conveniência: feriados B3 cobrindo o intervalo de datas [inicio, fim]."""
    return feriados_b3(inicio.year, fim.year)


def dias_uteis_entre(emissao: date, fim: date, feriados) -> list[date]:
    """Dias úteis em (emissao, fim], exclusivo na emissao.
    Dia útil = dia de semana (seg–sex) que não é feriado B3."""
    out, d = [], emissao + timedelta(days=1)
    while d <= fim:
        if d.weekday() < 5 and d not in feriados:
            out.append(d)
        d += timedelta(days=1)
    return out


def contar_du(emissao: date, fim: date, feriados) -> int:
    return len(dias_uteis_entre(emissao, fim, feriados))
