# Modelo de Marcação a Mercado (MtM) — Migração VBA → Python

Migração do modelo de **marcação a mercado de crédito privado** de uma planilha
Excel/VBA para um motor em Python que precifica a carteira seguindo as convenções
**ANBIMA** (CDI, CDI+spread, prefixado e IPCA+spread) e valida cada ativo no centavo.

> Projeto profissional. Dados de clientes e contrapartes foram removidos; os exemplos
> em `exemplos/` são fictícios.

## Problema

A marcação a mercado da carteira era **manual, planilha por planilha**. Para fechar o
MtM do dia, era preciso abrir dezenas de calculadoras em Excel, colar a taxa de CDI,
recalcular e salvar — uma a uma. Levava **horas**, dependia de uma única pessoa conhecer
cada planilha e era vulnerável a **erro operacional**: taxa aplicada em dia não-útil,
data trocada, convenção de contagem de dias inconsistente. Em uma instituição regulada,
um preço unitário errado é uma cota errada.

## Solução

Um **motor de precificação independente** que recalcula cada ativo do zero pelas
convenções ANBIMA e serve como **segunda fonte de verdade** contra a planilha. Sobre ele,
uma rotina que atualiza os dados de mercado, recalcula toda a carteira e faz o
**batimento automático** (motor × planilha) no centavo, com um painel que mostra, ativo
a ativo, o que está conforme e o que precisa de atenção.

## Resultado

- Fechamento da carteira inteira: de **~horas de trabalho manual** para **21 segundos**.
- **Batimento no centavo** em toda a carteira, automático e auditável.
- **Erros operacionais eliminados**: convenções de contagem de dias e de dia útil
  padronizadas; divergências antes invisíveis passaram a ser detectadas na hora
  (ex.: uma calculadora estava "congelada" por uma fonte de dados quebrada, e o painel
  a sinalizou).
- Processo antes dependente de uma pessoa virou **rotina de um comando**.

## Tecnologias

Python · Excel/VBA · openpyxl · calendário de dias úteis B3/ANBIMA · automação de rotina
