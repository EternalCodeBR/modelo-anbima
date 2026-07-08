# pu_mtm/verificacao/comparador.py
"""Compara PU Python x PU Excel com régua RELATIVA única (escala-invariante).

Régua unificada motor x calculadora: aprova se o erro relativo ao |excel| ficar
dentro de _TOL_REL. Relativa (não absoluta) por dois motivos:
- (a) não estoura no piso de precisão do float64 em PU grande — um absoluto fixo
  de 1e-8 vira ~1e-15 relativo num PU de milhões (futuros IPCA), gerando falso
  positivo só por ruído de ponto flutuante;
- (b) mede de fato PU pequeno (unitário ~0,01) e PU ~1 (prefixados VNe=1), onde o
  antigo "centavo" dava uma folga de ~1% inútil.

Limiar 2e-8: 2 na oitava casa decimal (≡ 2×10⁻⁶%), folga segura para o pior ativo
limpo do portfólio. `bate_no_centavo` fica só como referência do conceito de negócio
"bate no centavo"; NÃO é mais a régua de aprovação.
"""

_TOL_REL = 2e-8  # tolerância relativa única: 2 na 8ª casa decimal ≡ 2×10⁻⁶% (escala-invariante)


def bate_no_centavo(py: float, excel: float) -> bool:
    """Conceito de negócio: diferença < 1 centavo. Mantido para referência/testes."""
    return round(abs(py - excel), 2) < 0.01


def bate(py: float, excel: float, escala: float | None = None) -> bool:
    """Aprova se |py - excel| <= |excel|·_TOL_REL. `escala` é ignorado (mantido por
    compatibilidade de chamada); a régua agora é puramente relativa."""
    return abs(py - excel) <= max(abs(excel), 1e-12) * _TOL_REL


def comparar(py: float, excel: float, id_serie: str, vne: float | None = None) -> dict:
    return {"id_serie": id_serie, "py": py, "excel": excel,
            "dif": py - excel, "ok": bate(py, excel)}
