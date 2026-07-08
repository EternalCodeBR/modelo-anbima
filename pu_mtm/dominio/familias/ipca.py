"""Família ipca_spread (CRI/CRA/NC/Debênture corrigidos por IPCA): VNA + spread.

Espelha a calculadora NTN-B da casa (validado na Ativo IPCA 367, divergência ~1e-15):

- **VNA (Valor Nominal Atualizado)** — a calculadora trabalha em **parcelas mensais
  15→15** (a data-base do IPCA). Em cada parcela aplica a variação mensal `H` (%) do
  índice pro-rata 30/360: `I = ROUND((1+H/100)^(g/30), 8)`, com `g = DAYS360(início, dia)`
  (g=30 na parcela cheia). O saldo atualiza `J = E·I` e, na virada de mês, `E ← J` —
  ou seja o VNA **compõe mês a mês com o I arredondado em 8 casas**. Logo o fator VNA
  acumulado é `Π das parcelas cheias ROUND((1+H/100)^(g/30),8) × parcela corrente`.
  A última parcela usa a **projeção ANBIMA** do índice (número-índice ainda não
  divulgado). Datas de cupom dividem o mês numa parcela extra (reset do saldo).

- **spread** — prêmio anual sobre o VNA, capitalizado 30/360 igual ao prefixado:
  `(1+spread)^(dcp/dc)`, com `dc=360`. Equivale ao `(1+A1)^(g/30)` mensal da planilha
  (`A1 = (1+spread)^(1/12)-1`), que telescopa em `(1+spread)^(dcp_30360/360)`.

PU = Saldo·fator_ipca + juros − eventos, com juros = Saldo·fator_ipca·(spread−1).
Como `fator_fn` devolve `fator_ipca × spread`, o núcleo universal monta
`PU = Saldo·(fator−1) + Saldo − eventos = VNA·spread − eventos` — exatamente a coluna O.
"""
from datetime import date
from pu_mtm.dominio.modelos import Ativo, DiaCalc


def days360_us(d1: date, d2: date) -> int:
    """DAYS360 do Excel (método US/NASD, 4º arg FALSE): cada mês conta 30 dias."""
    a, b = d1.day, d2.day
    if a == 31:
        a = 30
    if b == 31 and a == 30:
        b = 30
    return (d2.year - d1.year) * 360 + (d2.month - d1.month) * 30 + (b - a)


def fator_vna(parcelas: list[tuple[date, date]], h_por_parcela: list[float],
              dia: date) -> float:
    """Fator VNA acumulado da 1ª parcela até `dia` (replica E/I da calculadora).

    `parcelas[k]` é o par (início, fim) da k-ésima parcela mensal (já incluídos os
    splits de cupom); `h_por_parcela[k]` é a variação do IPCA (%) dessa parcela.
    Multiplica o I cheio (g=DAYS360(início,fim)) de cada parcela encerrada e o I
    pro-rata (g=DAYS360(início,dia)) da parcela corrente — cada I arredondado em 8.
    """
    f = 1.0
    for (ini, fim), h in zip(parcelas, h_por_parcela):
        if dia >= fim:                                   # parcela encerrada: I cheio
            g = days360_us(ini, fim)
            f *= round((1.0 + h / 100.0) ** (g / 30.0), 8)
        elif dia >= ini:                                 # parcela corrente: pro-rata
            g = days360_us(ini, dia)
            f *= round((1.0 + h / 100.0) ** (g / 30.0), 8)
            break
        else:
            break
    return f


def fator_juros_acumulado(ativo: Ativo, dias: list[DiaCalc]) -> float:
    """fator = VNA acumulado (já pronto em `dia.fator_ipca`, relativo à âncora de
    reset) × fator de spread 30/360 `(1+spread)^(dcp/dc)`."""
    ult = dias[-1]
    spread = (1.0 + ativo.spread) ** (ult.dcp / ult.dc)
    return ult.fator_ipca * spread
