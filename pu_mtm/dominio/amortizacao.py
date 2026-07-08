"""Indexa o livro de eventos realizados por data."""
from datetime import date
from pu_mtm.dominio.modelos import Evento

def indexar_eventos(eventos: list[Evento]) -> dict:
    return {e.data: (e.evento_juros, e.evento_amortizacao) for e in eventos}

def evento_na_data(indice: dict, dia: date) -> tuple[float, float]:
    return indice.get(dia, (0.0, 0.0))
