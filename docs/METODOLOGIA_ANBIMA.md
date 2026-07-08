# Metodologia de Precificação (ANBIMA)

Descrição pública das convenções usadas pelo motor para marcar a mercado ativos de
crédito privado. Baseada nas convenções de mercado ANBIMA / B3. Não contém dados de
clientes.

## Calendário e contagem de dias

- **Dias úteis** pelo calendário **B3/BM&FBOVESPA**, obtido da biblioteca `holidays`
  (`financial_holidays("BVMF")`) — nunca derivado de planilha.
- **DU/252** para indexadores em dias úteis (DI); **30/360** e **ACT/365** para variantes
  prefixadas; **DC** (dias corridos) para o VNA do IPCA.

## Famílias

### DI (% do CDI) — `di_puro`
PU acumula o fator diário do CDI (convenção **DI-over**: a taxa aplicada em um dia é a do
**dia útil anterior**):

```
fator_dia = (1 + CDI_ano/100) ^ (1/252)
PU        = VNe × Π fator_dia   (opcionalmente × percentual do CDI)
```

### DI + spread — `di_spread`
Fator do CDI multiplicado por um fator de spread. Há variantes de convenção observadas no
mercado (todas suportadas):
- spread em **DU/252**;
- spread em **30/360**;
- fatores **crus** (sem arredondamento intermediário).

### Prefixado — `prefixado`
Taxa fixa capitalizada pela contagem de dias do papel (30/360, ACT/365, DU/252 ou mensal),
descontada até a data de marcação.

### IPCA + spread — `ipca_spread`
Espelha a NTN-B: **VNA** corrigido pelo IPCA (com a defasagem de projeção usual) mais um
spread real, seguindo a convenção ANBIMA de NTN-B.

## Accrual e eventos

- **Accrual diário** entre eventos; em datas de **juros/amortização** o saldo é reduzido e
  o PU reinicia sobre o novo saldo.
- **Amortização** aplicada na âncora do fluxo (coluna de evento), não por aproximação de
  aniversário.

## Batimento

- O motor é uma **segunda fonte de verdade**, independente da planilha.
- Comparação **no centavo**, **apenas em dia útil** — em dia não-útil a planilha acumula
  pelo calendário e o motor segura no último dia útil; as duas reconvergem no próximo dia
  útil, então comparar nesses dias geraria falso erro.
