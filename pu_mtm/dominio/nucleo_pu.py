"""Montagem universal do PU dia a dia: PU = Saldo + Juros - EvJuros - EvAmort.

Modelo validado calculadora por calculadora (ver docs/.../modelo-pu-validado.md):
- bullet: fator acumula do início; saldo = VNe constante.
- evento (aniversário com pagamento): paga EvJuros/EvAmort e, no dia seguinte,
  **Saldo <- PU do evento** e o **fator reinicia** (acúmulo a partir do dia seguinte).
  Essa regra única reproduz amortização, cupom integral e cupom parcial (capitalização).

O arredondamento do juros é parâmetro por ativo (alguns truncam em 8, outros não).
"""
from typing import Callable
from pu_mtm.dominio.modelos import Ativo, DiaCalc, Evento, ResultadoPU
from pu_mtm.dominio.anbima import trunc, round_anbima
from pu_mtm.dominio.amortizacao import indexar_eventos, evento_na_data


def _trunc8(j: float) -> float:
    return trunc(j, 8)


def arred_juros_por_nome(nome: str) -> Callable[[float], float]:
    """Mapeia o parâmetro `juros_arred` do cadastro para a função de arredondamento
    do juros do dia. `nenhum` (ex.: prefixado 476) não arredonda."""
    return {
        "nenhum": lambda j: j,
        "trunc8": _trunc8,
        "round8": lambda j: round_anbima(j, 8),
    }[nome]


def calcular_pu(ativo: Ativo, dias: list[DiaCalc], eventos: list[Evento],
                fator_fn: Callable[[Ativo, list[DiaCalc]], float],
                arred_juros: Callable[[float], float] = _trunc8) -> ResultadoPU:
    """`fator_fn(ativo, segmento)` devolve o fator de juros acumulado sobre o
    segmento de dias desde o último reset (inclusive). `arred_juros` arredonda o
    juros do dia (default: TRUNC 8 casas)."""
    idx = indexar_eventos(eventos)
    saldo = ativo.vne
    pu = ativo.vne
    serie = []
    inicio_seg = 0  # índice do 1º dia do segmento de acúmulo corrente (reinicia no evento)
    for i, dia in enumerate(dias):
        segmento = dias[inicio_seg: i + 1]
        fator = fator_fn(ativo, segmento)
        juros = arred_juros(saldo * (fator - 1.0))
        ev_juros, ev_amort = evento_na_data(idx, dia.data)
        pu = saldo + juros - ev_juros - ev_amort
        if ev_juros or ev_amort:           # aniversário com pagamento
            saldo = pu                     # regra universal: Saldo <- PU do evento
            inicio_seg = i + 1             # fator reinicia no dia seguinte
        serie.append((dia.data, pu))
    data_ref = dias[-1].data if dias else ativo.data_emissao
    return ResultadoPU(id_serie=ativo.id_serie, data_ref=data_ref, pu=pu, serie_diaria=serie)
