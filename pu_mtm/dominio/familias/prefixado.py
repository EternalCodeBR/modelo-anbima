"""Fator de juros prefixado: (1+taxa)^(dcp/dc), ou (1+taxa)^expo no modo per-período."""
from pu_mtm.dominio.modelos import Ativo, DiaCalc

def fator_juros_acumulado(ativo: Ativo, dias: list[DiaCalc]) -> float:
    ultimo = dias[-1]
    if ultimo.expo is not None:                       # per-período (ex.: 738 mensal)
        return (1.0 + ativo.taxa_fixa) ** ultimo.expo
    return (1.0 + ativo.taxa_fixa) ** (ultimo.dcp / ultimo.dc)
