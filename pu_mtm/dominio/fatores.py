"""Fatores ANBIMA. DI base 252; spread 30/360."""
from pu_mtm.dominio.anbima import round_anbima

def fator_di_diario(cdi_aa: float, percentual: float = 100.0) -> float:
    """Fator diário DI, SEM arredondar. Para `percentual` (P) % do CDI:

        TDIk = (1+CDI)^(1/252) − 1   (taxa DI diária, CDI puro)
        fator = 1 + (P/100)·TDIk

    Com P=100 retornamos exatamente `(1+CDI)^(1/252)` (caminho idêntico ao di_puro;
    bit-a-bit, garantido pelo atalho abaixo). P=105 (RDB 464/474) multiplica o TDIk:
    o prêmio escala com o CDI — distinto do spread, que é um fator anual à parte.

    A calculadora-piloto (Ativo DI-piloto) usa o fator cru: `H = (1+G/100)^(1/252)` sem
    ROUND. Arredondar a 8 casas acumula >R$0,20 sobre ~130 DU e estoura o centavo
    (verificado na Fase 0). Espelhamos o Excel: cru.
    """
    base = (1.0 + cdi_aa) ** (1.0 / 252.0)
    if percentual == 100.0:
        return base                       # idêntico ao di_puro (sem regressão)
    return 1.0 + (percentual / 100.0) * (base - 1.0)

def fator_di_acumulado(cdis: list[float]) -> float:
    """Produto dos fatores diários crus."""
    fator = 1.0
    for cdi in cdis:
        fator *= fator_di_diario(cdi)
    return fator

def fator_spread_acumulado(spread_aa: float, dcp: int, dc: int) -> float:
    """ROUND((1+spread)^(DCP/DC), 9)."""
    return round_anbima((1.0 + spread_aa) ** (dcp / dc), 9)
