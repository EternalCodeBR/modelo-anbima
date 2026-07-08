"""Fator de juros das famílias DI (puro, % de CDI e spread)."""
from pu_mtm.dominio.modelos import Ativo, DiaCalc
from pu_mtm.dominio.fatores import fator_di_diario, fator_spread_acumulado
from pu_mtm.dominio.anbima import round_anbima, trunc


def _fator_di_spread(ativo: Ativo, dias: list[DiaCalc]) -> float:
    """Pipeline di_spread, provado bit-a-bit na calculadora 741 (Ativo DI+Spread A):

        TDIk diário  G = ROUND((1+CDI)^(1/252) − 1, 8)
        Fator DI     H = TRUNC(H_ant · (1+G), 16)            (acumulado)
        Fator Spread K = ROUND((1+spread)^(DP/base), 9)      DP = nº de DU do segmento
        Fator juros  L = ROUND(H · K, 9)

    O juros do dia (TRUNC 8) entra via `arred_juros` no núcleo. Distinto do di_puro:
    a TDIk é arredondada a 8 e o fator DI acumula truncando a 16 — espelha o Excel."""
    fator_di = 1.0
    for d in dias:
        tdik = round_anbima((1.0 + d.cdi) ** (1.0 / 252.0) - 1.0, 8)
        fator_di = trunc(fator_di * (1.0 + tdik), 16)
    if ativo.base == 360:
        # Variante Ativo DI+Spread B/FOCUS (691/763): Fator Spread = ROUND((1+s)^(DCP/360),9)
        # sobre dias CORRIDOS (30/360); juros = TRUNC(VNe·(DI·spread−1),8), sem ROUND9 no produto.
        ultimo = dias[-1]
        fator_spread = fator_spread_acumulado(ativo.spread, ultimo.dcp, ultimo.dc)
        return fator_di * fator_spread
    # Variante Ativo DI+Spread A (741): Fator Spread = ROUND((1+s)^(DP/252),9), DP = nº de DU;
    # Fator de juros = ROUND(DI·spread, 9).
    dp = len(dias)  # contador de dias úteis (coluna DP), reinicia a cada segmento
    fator_spread = fator_spread_acumulado(ativo.spread, dcp=dp, dc=ativo.base)
    return round_anbima(fator_di * fator_spread, 9)


def _fator_cru(ativo: Ativo, dias: list[DiaCalc]) -> float:
    """Ativo Cru (Confissão de Dívida): fatores CRUS (sem ROUND/TRUNC), provado na
    'Calculadora OF'. DI e Spread acumulam dentro do segmento e AMBOS reiniciam no
    aniversário COM pagamento (via reset do núcleo) — aniversários de carência não
    resetam. Juros do dia sem arredondar.

        Fator DI     = Π (1+CDI)^(1/252)            (segmento; cru)
        Fator Spread = (1+spread)^(n_spread/252)    (segmento; cru)

    No 1º segmento (o aditamento, âncora 07/04/2025) o spread leva 1 dia de vantagem
    sobre o DI (no dia-âncora o DI=0 mas o spread já conta 1 dia, DP=1)."""
    fator_di = 1.0
    for d in dias:
        fator_di *= (1.0 + d.cdi) ** (1.0 / 252.0)
    n_seg = len(dias)
    n_spread = n_seg + 1 if n_seg == dias[-1].du else n_seg   # +1 só no 1º segmento
    fator_spread = (1.0 + ativo.spread) ** (n_spread / 252.0)
    return fator_di * fator_spread


def fator_juros_acumulado(ativo: Ativo, dias: list[DiaCalc]) -> float:
    if ativo.familia == "di_spread":
        if ativo.fator_diario_arred == "cru":
            return _fator_cru(ativo, dias)
        return _fator_di_spread(ativo, dias)
    fator_di = 1.0
    for d in dias:
        fator_di *= fator_di_diario(d.cdi, ativo.percentual_cdi)
    if ativo.fator_diario_arred == "round8":
        # 464/474 (RDB): PU = VNe × ROUND(fator_acumulado, 8). A acumulação é crua;
        # o ROUND a 8 entra no fator usado para montar o PU (juros_arred=nenhum).
        fator_di = round_anbima(fator_di, 8)
    return fator_di
