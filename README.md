# Motor de Marcação a Mercado (MtM) — Crédito Privado

Motor em Python que calcula o **Preço Unitário (PU)** marcado a mercado de uma carteira
de crédito privado brasileiro (CDI, CDI+spread, prefixado e IPCA+spread), seguindo as
convenções **ANBIMA**, e valida cada ativo no centavo contra a planilha de referência.

> Projeto de portfólio. Todos os dados de clientes, nomes de contrapartes e caminhos
> internos foram removidos; os exemplos em `exemplos/` são **fictícios**.

---

## Problema

A marcação a mercado da carteira era **100% manual**, planilha por planilha:

- Cada ativo tinha sua própria calculadora em Excel (dezenas de arquivos).
- Para atualizar o MtM do dia, era preciso **abrir cada planilha**, colar a taxa de CDI,
  forçar o recálculo e salvar — repetido para toda a carteira.
- O processo levava **horas**, dependia de **uma pessoa** conhecer cada planilha, e não
  tinha rastreabilidade.
- Estava sujeito a **erro humano**: fórmula colada errada, data trocada, taxa aplicada em
  dia não-útil, convenção de contagem de dias inconsistente entre planilhas.
- Divergências de precificação só apareciam **tarde**, quando alguém conferia à mão.

Em uma instituição regulada, isso é risco operacional: um PU errado é uma cota errada.

---

## Solução

Um **motor de precificação independente** + um **pipeline de validação automatizado**:

- **Motor de PU (Python).** Recalcula cada ativo do zero pelas convenções ANBIMA —
  fator DI acumulado, spread, indexação IPCA/prefixado, contagem de dias úteis (calendário
  B3), amortização e eventos. É uma **segunda fonte de verdade**, independente da planilha.
- **Batimento no centavo.** Compara, ativo a ativo, o PU do motor com o da planilha, **só
  em dia útil** (em dia não-útil as duas convenções divergem legitimamente e reconvergem).
- **Painel de validação em 3 camadas** (estrutura da planilha, dados de mercado em dia,
  preço batendo) que mostra **verde/vermelho por ativo** — separa erro real de ruído.
- **Correções cirúrgicas** nas planilhas Excel direto no XML, sem abrir o Excel e **sem
  corromper** desenhos, comentários ou fórmulas — usado para otimizar fórmulas lentas e
  limpar dependências mortas.
- **Rotina de um comando** que encadeia: atualizar dados de mercado → recalcular o cache
  das planilhas → rodar o motor → bater. O que era um turno de trabalho vira um comando.

---

## Resultado

| Métrica | Antes | Depois |
|---|---|---|
| Recálculo da carteira inteira | ~10 min (e antes disso, minutos por arquivo) | **21 segundos** |
| Recálculo de fórmula por planilha | ~70 s | **0,3 s** |
| Batimento motor × planilha | conferência manual, tardia | **no centavo, automático** |
| Dependência de pessoa | uma pessoa por planilha | **rotina auditável de 1 comando** |

Ganhos concretos obtidos no caminho:

- **Diagnóstico guiado por medição, não por chute.** O gargalo do recálculo não era o
  cálculo — era o **AutoSave do OneDrive** no salvamento (descoberto isolando as fases
  Abrir/Calcular/Salvar). Recalcular em cópia local derrubou 10 min → 21 s.
- **Fórmula lenta corrigida.** Funções `WORKDAY` varrendo uma **coluna inteira** de
  feriados (1 milhão de linhas) faziam o recálculo levar 70 s; limitar o intervalo levou a
  0,3 s — sem alterar nenhum resultado.
- **Erros eliminados do batimento:** comparação restrita a dia útil e ao escopo correto de
  famílias acabou com falsos positivos que poluíam a conferência.
- **Achados reais de auditoria:** uma calculadora estava **congelada** por um vínculo
  externo morto que retornava zero; outras carregavam vínculos órfãos — tudo detectado
  pelo painel de validação e corrigido cirurgicamente.

---

## Arquitetura (visão geral)

```
Dados de mercado (CDI/IPCA)  ─►  Motor de PU (Python, convenções ANBIMA)
                                        │
Calculadoras Excel  ◄─ recálculo ─►  Batimento (motor × planilha, no centavo)
                                        │
                                 Painel de validação (verde/vermelho por ativo)
```

- `pu_mtm/` — o motor: domínio de precificação (famílias, fatores, amortização, ANBIMA),
  leitura de dados, verificação/batimento.
- `mtm_skills/` — ferramentas operacionais (atualização de dados, otimização e correção
  cirúrgica de planilhas, painel de validação, geração de saída).
- `vba/` — macro nativa de recálculo do cache (recalcula fora da nuvem para evitar o
  gargalo do AutoSave).
- `docs/METODOLOGIA_ANBIMA.md` — a metodologia pública de precificação.
- `exemplos/` — cadastro de ativos **fictício** para ilustrar o formato.

## Stack

Python 3.12 · openpyxl · holidays (calendário B3/BVMF) · VBA (Excel) · win32com (disparo
nativo da macro).
