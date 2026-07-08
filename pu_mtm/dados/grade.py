"""Montagem da grade diária por família.

Prefixado (base 360, 30/360): grade de **dias corridos**. O fator de um segmento é
`(1+taxa)^(DCP/30)`, onde DCP é a contagem 30/360 desde a **âncora de reset** — o
último cupom pago antes do dia (ou a emissão, se ainda não houve cupom). DC=30 fixo.

Equivalência provada na 476 (bater_476.py, divergência 1e-15): o produto mês a mês
`Π (1,02)^(Gₘ/30)` colapsa em `(1,02)^(DCP_30/360(âncora, dia)/30)`, porque
`Σ Gₘ = DCP_30/360(âncora, dia)`. Por isso a âncora basta — não é preciso varrer mês a mês.
"""
import calendar
from datetime import date, timedelta
from pu_mtm.dominio.modelos import Ativo, DiaCalc, Evento
from pu_mtm.dominio.familias.ipca import days360_us


def _add_meses(d: date, n: int) -> date:
    """Soma n meses preservando o dia (clampa ao último dia do mês quando necessário)."""
    m = d.month - 1 + n
    y = d.year + m // 12
    m = m % 12 + 1
    return date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def _expo_mensal(emissao: date, d: date) -> float:
    """Expoente per-período mensal: nº de aniversários mensais ≤ d + fração do período
    corrente (dias corridos reais). Espelha (1+taxa)^(meses + dias/dias_do_mês) das
    calculadoras Ativo Mensal mensais (ex.: 738), cujo DC = dias reais de cada mês."""
    k = 0
    while _add_meses(emissao, k + 1) <= d:
        k += 1
    anchor = emissao if k == 0 else _add_meses(emissao, k)
    prox = _add_meses(emissao, k + 1)
    return k + (d - anchor).days / (prox - anchor).days


def dcp_30360(inicio: date, fim: date) -> int:
    """Dias corridos na convenção 30/360 (US): cada mês conta 30; espelha a coluna
    DCP da calculadora prefixada com DC=30 fixo."""
    return (fim.year - inicio.year) * 360 + (fim.month - inicio.month) * 30 + (fim.day - inicio.day)


def _ancora_reset(emissao: date, cupons: list[date], dia: date) -> date:
    """Último cupom **estritamente antes** de `dia` (o dia do cupom ainda pertence ao
    segmento anterior — acumula e paga); na ausência de cupom anterior, a emissão."""
    ancora = emissao
    for c in cupons:
        if c < dia:
            ancora = c
    return ancora


def montar_dias_prefixado(ativo: Ativo, data_ref: date, eventos: list[Evento],
                          feriados=None) -> list[DiaCalc]:
    """Grade prefixada da emissão a `data_ref`. O numerador do fator depende de
    `pre_daycount` e o denominador de `pre_dc`; o `dcp` reinicia a cada cupom (âncora):

    - "30360": dias corridos 30/360 (ex.: 476, dc=30 mensal).
    - "actual": dias corridos reais / ACT (ex.: 755/759, dc=365 anual).
    - "du": dias ÚTEIS (ex.: 750, dc=252 anual) — grade só de dias úteis; exige `feriados`.
    """
    cupons = sorted(e.data for e in eventos)
    out = []
    if ativo.pre_daycount == "mensal":
        # per-período mensal (738/722/762): fator = (1+taxa)^(meses_completos+fração),
        # DC = dias reais de cada mês; capitaliza (carry) através dos aniversários.
        d = ativo.data_emissao
        while d <= data_ref:
            out.append(DiaCalc(data=d, du=0, dcp=0, dc=1, cdi=0.0,
                               expo=_expo_mensal(ativo.data_emissao, d)))
            d += timedelta(days=1)
        return out
    if ativo.pre_daycount == "du":
        from pu_mtm.dados.feriados import dias_uteis_entre, contar_du
        grade = [ativo.data_emissao] + dias_uteis_entre(ativo.data_emissao, data_ref, feriados)
        for d in grade:
            ancora = _ancora_reset(ativo.data_emissao, cupons, d)
            dcp = contar_du(ancora, d, feriados)
            out.append(DiaCalc(data=d, du=dcp, dcp=dcp, dc=ativo.pre_dc, cdi=0.0))
        return out
    d = ativo.data_emissao
    while d <= data_ref:
        ancora = _ancora_reset(ativo.data_emissao, cupons, d)
        dcp = (d - ancora).days if ativo.pre_daycount == "actual" else dcp_30360(ancora, d)
        out.append(DiaCalc(data=d, du=0, dcp=dcp, dc=ativo.pre_dc, cdi=0.0))
        d += timedelta(days=1)
    return out


# --------------------------------------------------------------------------- #
#  IPCA (VNA pro-rata + spread) — parcelas mensais 15→15, defasagem de 2 meses
# --------------------------------------------------------------------------- #
def _proximo_quinze(d: date) -> date:
    """1º dia-15 estritamente após o período natural que abre em `d`."""
    return date(d.year, d.month, 15) if d.day < 15 else _add_meses(date(d.year, d.month, 15), 1)


def _quinze_abre(d: date) -> date:
    """O dia-15 que ABRE o período natural mensal contendo `d` (data-base do IPCA)."""
    if d.day >= 15:
        return date(d.year, d.month, 15)
    y, m = (d.year - 1, 12) if d.month == 1 else (d.year, d.month - 1)
    return date(y, m, 15)


def _ref_ipca(parc_inicio: date) -> tuple[int, int]:
    """Mês de referência do IPCA da parcela: mês do 15-de-abertura − 1 (defasagem total
    de 2 meses do padrão NTN-B). Ex.: parcela [15/10→15/11] usa IPCA de setembro."""
    p = _quinze_abre(parc_inicio)
    y, m = (p.year - 1, 12) if p.month == 1 else (p.year, p.month - 1)
    return (y, m)


def montar_dias_ipca(ativo: Ativo, data_ref: date, eventos: list[Evento],
                     ipca_var: dict[tuple[int, int], float], feriados=None) -> list[DiaCalc]:
    """Grade diária IPCA da emissão a `data_ref`. Cada dia carrega `fator_ipca` (VNA
    acumulado relativo à âncora de reset/cupom — Π mensal de ROUND((1+H/100)^(g/30),8))
    e `dcp` (30/360 desde a âncora) para o fator de spread `(1+spread)^(dcp/360)`.

    `ipca_var`: {(ano,mes): var%} do IPCA (BaseDadosMercado, INDICE_MENSAL_VALOR).
    Datas de cupom (eventos) entram como bordas extras de parcela e resetam a âncora.
    """
    cupons = sorted(e.data for e in eventos)
    bordas = {ativo.data_emissao}
    q = _proximo_quinze(ativo.data_emissao)
    while q <= data_ref:
        bordas.add(q)
        q = _add_meses(q, 1)
    bordas.add(q)                                  # 15 que fecha a parcela corrente
    for c in cupons:
        if ativo.data_emissao <= c <= data_ref:
            bordas.add(c)
    bordas = sorted(bordas)
    parcelas = list(zip(bordas, bordas[1:]))       # [(ini, fim), ...]
    H = [ipca_var.get(_ref_ipca(ini), 0.0) for ini, _ in parcelas]

    out = []
    d = ativo.data_emissao
    while d <= data_ref:
        ancora = _ancora_reset(ativo.data_emissao, cupons, d)
        f = 1.0
        for (pi, pf), h in zip(parcelas, H):
            if pi < ancora:
                continue                           # parcela anterior à âncora: já paga
            if d >= pf:                            # parcela encerrada: I cheio
                f *= round((1.0 + h / 100.0) ** (days360_us(pi, pf) / 30.0), 8)
            elif d >= pi:                          # parcela corrente: pro-rata
                f *= round((1.0 + h / 100.0) ** (days360_us(pi, d) / 30.0), 8)
                break
            else:
                break
        dcp = dcp_30360(ancora, d)
        out.append(DiaCalc(data=d, du=0, dcp=dcp, dc=360, cdi=0.0, fator_ipca=f))
        d += timedelta(days=1)
    return out
