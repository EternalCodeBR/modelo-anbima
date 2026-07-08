"""Reconstrói as calculadoras de homologação SEM NUNCA usar openpyxl.save().

Motivação: ``openpyxl.Workbook.save()`` reescreve o pacote .xlsx inteiro e, ao
fazê-lo, degrada silenciosamente partes que ele não sabe reproduzir fielmente —
VML (``vmlDrawing*.vml`` = "Forma de desenho"), DrawingML, comentários threaded,
links externos, deixa ``<v></v>`` vazio, etc. Cada save introduz uma classe nova
de corrupção; remendar depois é jogo de gato-e-rato interminável.

Este pipeline é 100% cirúrgico no ZIP:

    1. copia a ORIGINAL do SharePoint byte-a-byte para a pasta de homologação
       (``<CDI|CDI + Spread>/<Apelido> - Homologação/<arquivo>``);
    2. em cada ``xl/worksheets/sheetN.xml``, troca APENAS o texto das fórmulas de
       taxa: ``SUMIFS([n]SERIE_VALOR!...)`` -> ``_xlfn.XLOOKUP(<data>,CDI!A:A,CDI!B:B,0)``
       (preservando envelope IF/``/100`` e a célula de data), e descarta o valor
       cacheado obsoleto da célula (força recálculo limpo);
    3. reempacota o .xlsx com TODAS as demais partes idênticas ao original.

Resultado: o arquivo é bit-a-bit igual ao original (que abre limpo no Excel),
exceto pelo texto das fórmulas de taxa. Nenhuma parte visual/estrutural é tocada.

A aba CDI já existe em 22/23 originais (a Ativo DI-D não tem — tratada à
parte). A atualização de valores do CDI é uma etapa separada
(``atualizar_cdi_calculadoras``), também sem openpyxl no futuro.

Escrita atômica (temp + os.replace). Uso:
    python -m mtm_skills.rebuild_surgical            # reconstrói os 23
    python -m mtm_skills.rebuild_surgical --check     # só relata SUMIFS por ativo
"""
from __future__ import annotations

import os
import re
import shutil
import sys
import time
import warnings
import zipfile
from pathlib import Path

warnings.filterwarnings("ignore")

_RAIZ = Path(__file__).resolve().parents[1]
if str(_RAIZ) not in sys.path:
    sys.path.insert(0, str(_RAIZ))

from mtm_skills.rebuild_homolog import carregar_ativos, sumifs_para_xlookup, FAMILIA_LABEL

DELAY = float(os.environ.get("REBUILD_DELAY", "0"))

# Célula <c ...> cuja fórmula <f> contém SUMIFS. A fórmula não tem '<' literal
# (os operadores viram &lt;/&gt;), então [^<]* captura o texto inteiro.
_RE_CELL_SUMIFS = re.compile(
    r'<c\b([^>]*)>\s*<f([^>]*)>([^<]*SUMIFS[^<]*)</f>\s*(?:<v[^>]*>[^<]*</v>)?\s*</c>',
    re.IGNORECASE,
)


def _swap_sumifs_no_xml(xml: str) -> tuple[str, int]:
    """Troca SUMIFS->XLOOKUP no texto das fórmulas de uma planilha (XML cru).

    Descarta ``t="..."`` e o ``<v>`` cacheado da célula: a fórmula mudou, o valor
    antigo ficou obsoleto; sem ``<v>`` a célula é "ainda não calculada" (estado
    que o Excel sempre aceita, e o ``calcPr fullCalcOnLoad`` já força recálculo).
    """
    total = 0

    def _repl(m: re.Match) -> str:
        nonlocal total
        attrs, fattrs, formula = m.group(1), m.group(2), m.group(3)
        nova, n, _ = sumifs_para_xlookup(formula)
        if n == 0:
            return m.group(0)
        total += n
        attrs_sem_t = re.sub(r'\s+t="[^"]*"', "", attrs)
        return f"<c{attrs_sem_t}><f{fattrs}>{nova}</f></c>"

    return _RE_CELL_SUMIFS.sub(_repl, xml), total


def _reconstruir(origem: Path, destino: Path) -> tuple[int, list[str]]:
    """Copia origem->destino e troca SUMIFS->XLOOKUP no XML. Retorna (n_trocas, abas)."""
    destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(str(origem), str(destino))  # byte-a-byte

    with zipfile.ZipFile(str(destino)) as zin:
        infos = zin.infolist()
        conteudo = {i.filename: zin.read(i.filename) for i in infos}

    total = 0
    abas_alteradas: list[str] = []
    for nome in list(conteudo):
        if re.match(r"^xl/worksheets/sheet\d+\.xml$", nome):
            txt = conteudo[nome].decode("utf-8")
            nova, n = _swap_sumifs_no_xml(txt)
            if n:
                conteudo[nome] = nova.encode("utf-8")
                total += n
                abas_alteradas.append(nome)

    if total:
        tmp = destino.with_name(destino.name + ".tmp")
        with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
            for i in infos:
                zout.writestr(i, conteudo[i.filename])
        os.replace(str(tmp), str(destino))
    return total, abas_alteradas


def _apagar_familias() -> None:
    for label in FAMILIA_LABEL.values():
        d = _RAIZ / "data" / "Calculadoras - Homologação" / label
        if d.exists():
            shutil.rmtree(d, onexc=lambda f, p, e: None)


def _contar_sumifs(origem: Path) -> int:
    with zipfile.ZipFile(str(origem)) as z:
        n = 0
        for nome in z.namelist():
            if re.match(r"^xl/worksheets/sheet\d+\.xml$", nome):
                n += z.read(nome).decode("utf-8", "replace").count("SUMIFS")
    return n


def check() -> None:
    ativos = carregar_ativos()
    for a in ativos:
        if a.origem.exists():
            print(f"  {a.idserie:>10} {a.apelido:<26} SUMIFS na origem: {_contar_sumifs(a.origem)}")
        else:
            print(f"  {a.idserie:>10} {a.apelido:<26} ORIG FALTA")


def rebuild() -> None:
    ativos = carregar_ativos()
    print(f"Reconstrução cirúrgica de {len(ativos)} calculadora(s) (sem openpyxl.save)\n")
    _apagar_familias()
    ok = falta = 0
    for a in ativos:
        if not a.origem.exists():
            print(f"  [{a.idserie:>10}] {a.apelido:<26} ORIG FALTA")
            falta += 1
            continue
        n, abas = _reconstruir(a.origem, a.destino)
        print(f"  [{a.idserie:>10}] {a.apelido:<26} {n:>5} SUMIFS->XLOOKUP  "
              f"({a.familia_label})")
        ok += 1
        if DELAY:
            time.sleep(DELAY)
    print(f"\nReconstruídas: {ok} | Faltando: {falta}")
    print("As partes visuais (VML, desenhos, comentários) foram preservadas byte-a-byte.")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    if "--check" in sys.argv:
        check()
    else:
        rebuild()
