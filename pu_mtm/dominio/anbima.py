"""Arredondamento/truncamento que espelha ROUND/TRUNC do Excel (cálculo em float)."""
import math

def round_anbima(x: float, casas: int) -> float:
    """ROUND do Excel: meio para cima (away from zero)."""
    f = 10 ** casas
    sinal = 1.0 if x >= 0 else -1.0
    val = abs(x) * f
    # round(val, 6) limpa a poeira de ponto flutuante (ex: 0.49999999999 -> 0.5)
    return sinal * math.floor(round(val, 6) + 0.5) / f

def trunc(x: float, casas: int) -> float:
    """TRUNC do Excel: corta em direção a zero."""
    f = 10 ** casas
    val = x * f
    # round(val, 6) limpa a poeira (ex: 0.99999999999 -> 1.0) para que o trunc bata com o Excel
    return math.trunc(round(val, 6)) / f
