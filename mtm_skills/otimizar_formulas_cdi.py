"""Otimiza as fórmulas das calculadoras — elimina o gargalo de recálculo (70s → ~1s).

Causa raiz (medida com cronômetro por fase: Abrir 0,5s | Calcular 70,6s | Salvar 0,2s):
o cálculo é 100% do custo, e o vilão é **referência de COLUNA INTEIRA** passada para
funções que NÃO otimizam o intervalo para o "usado":

  * ``WORKDAY(C3, 1, Feriados!A:A)`` — 1.500+ por planilha. ``Feriados!A:A`` =
    1.048.576 células. WORKDAY processa TODAS para montar o conjunto de feriados.
    ~1.500 × 1M = ~1,5 bilhão de leituras por recálculo → os 70s.
  * Também: ``Planilha2!D:J``, ``SERIE_VALOR!$A:$A``, ``Fluxo!E:F`` (hogs locais).

(O XLOOKUP com ``CDI!A:A`` é a exceção: o Excel já reduz ao intervalo usado, então
não pesava — mas limitar também não custa e deixa tudo consistente.)

Correção CIRÚRGICA (sem Excel, sem openpyxl): para cada aba, descobre a última linha
REAL usada e troca ``Aba!X:X`` (coluna inteira) por ``Aba!$X$1:$X$N`` (N = última
linha + folga). Reempacota o .xlsx com as demais partes byte-a-byte idênticas.
NÃO altera nenhum resultado — só a forma de calcular. Cache de PU permanece válido.

Fórmulas compartilhadas (``t="shared"``): o intervalo TEM que ser ABSOLUTO
(``$X$1:$X$N``), senão desliza por linha. Por isso o ``$``.

Arquivo ABERTO no Excel é PULADO (não trava o lote).

Uso:
    python -m mtm_skills.otimizar_formulas_cdi            # aplica em todas
    python -m mtm_skills.otimizar_formulas_cdi --check    # só relata, não grava
"""
from __future__ import annotations

import os
import re
import sys
import warnings
import zipfile
from pathlib import Path

warnings.filterwarnings("ignore")

_RAIZ = Path(__file__).resolve().parents[1]
if str(_RAIZ) not in sys.path:
    sys.path.insert(0, str(_RAIZ))

from pu_mtm.app import config

# Folga de linhas acima da última usada, para o append diário nunca ultrapassar o teto.
_FOLGA = 2000

_RE_WS = re.compile(r"xl/worksheets/sheet\d+\.xml$")
_RE_FORCE_FULL = re.compile(r'\sforceFullCalc="1"')
_RE_ROW = re.compile(r'<row r="(\d+)"')


def _mapa_abas(conteudo: dict[str, bytes]) -> dict[str, str]:
    """nome da aba -> caminho da parte (xl/worksheets/sheetN.xml)."""
    wb = conteudo.get("xl/workbook.xml", b"").decode("utf-8")
    rels = conteudo.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8")
    rid2tgt = {m.group(1): m.group(2) for m in re.finditer(
        r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels)}
    out: dict[str, str] = {}
    for m in re.finditer(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        tgt = rid2tgt.get(m.group(2), "")
        if tgt:
            out[m.group(1)] = "xl/" + tgt.lstrip("/")
    return out


def _ultima_linha(sheet_xml: str) -> int:
    linhas = [int(x) for x in _RE_ROW.findall(sheet_xml)]
    return max(linhas) if linhas else 1


def _limitar(xml: str, tetos: dict[str, int]) -> tuple[str, int]:
    """Troca `Aba!X:X` (col. inteira, 1+ colunas) por `Aba!$X$1:$Y$teto`."""
    n = 0
    for aba, teto in tetos.items():
        # Aba!  [ $? COL ] : [ $? COL ]   sem dígitos = coluna inteira
        pat = re.compile(
            r"(?<![A-Za-z0-9_.!'\[])" + re.escape(aba)
            + r"!\$?([A-Z]{1,3}):\$?([A-Z]{1,3})")

        def _sub(m: re.Match) -> str:
            nonlocal n
            n += 1
            c1, c2 = m.group(1), m.group(2)
            return f"{aba}!${c1}$1:${c2}${teto}"

        xml = pat.sub(_sub, xml)
    return xml, n


def _processar(path: Path, apenas_check: bool) -> str:
    try:
        with zipfile.ZipFile(str(path)) as zin:
            infos = zin.infolist()
            conteudo = {i.filename: zin.read(i.filename) for i in infos}
    except PermissionError:
        return "ABERTO no Excel — pulado"
    except zipfile.BadZipFile:
        return "ZIP inválido — pulado"

    abas = _mapa_abas(conteudo)
    if not abas:
        return "sem abas mapeáveis — pulado"

    # teto por aba = última linha real + folga
    tetos: dict[str, int] = {}
    for nome, parte in abas.items():
        if parte in conteudo:
            tetos[nome] = _ultima_linha(conteudo[parte].decode("utf-8")) + _FOLGA

    total = 0
    for nome, parte in abas.items():
        if parte in conteudo and _RE_WS.match(parte):
            novo, k = _limitar(conteudo[parte].decode("utf-8"), tetos)
            if k:
                conteudo[parte] = novo.encode("utf-8")
                total += k

    wbxml = conteudo.get("xl/workbook.xml", b"").decode("utf-8")
    tinha_ffc = bool(_RE_FORCE_FULL.search(wbxml))
    if tinha_ffc:
        conteudo["xl/workbook.xml"] = _RE_FORCE_FULL.sub("", wbxml).encode("utf-8")

    # calcChain.xml obsoleto (após mudar o texto das fórmulas) faz o Excel reconstruir
    # a árvore inteira no LOAD (abertura lenta). Removê-lo → Excel reconstrói no 1º
    # cálculo, rápido e incremental. Precisa sumir das 3 partes que o referenciam.
    tem_chain = "xl/calcChain.xml" in conteudo

    if total == 0 and not tinha_ffc and not tem_chain:
        return "já otimizada"

    dropou_chain = tem_chain
    if dropou_chain:
        del conteudo["xl/calcChain.xml"]
        infos = [i for i in infos if i.filename != "xl/calcChain.xml"]
        ct = conteudo.get("[Content_Types].xml", b"").decode("utf-8")
        conteudo["[Content_Types].xml"] = re.sub(
            r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct).encode("utf-8")
        wr = conteudo.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8")
        conteudo["xl/_rels/workbook.xml.rels"] = re.sub(
            r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', "", wr).encode("utf-8")

    resumo = (f"{total} ref limitada(s)"
              + (" | forceFullCalc removido" if tinha_ffc else "")
              + (" | calcChain limpo" if dropou_chain else ""))
    if apenas_check:
        return f"A OTIMIZAR: {resumo}"

    tmp = path.with_name(path.name + ".tmp")
    try:
        with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
            for i in infos:
                zout.writestr(i, conteudo[i.filename])
        os.replace(str(tmp), str(path))
    except PermissionError:
        if tmp.exists():
            tmp.unlink()
        return "ABERTO no Excel — pulado"
    return f"OTIMIZADA: {resumo}"


def _arquivos() -> list[Path]:
    out: list[Path] = []
    for label in ("CDI", "CDI + Spread"):
        base = config.HOMOLOG_CALC_ROOT / label
        if base.exists():
            out += sorted(p for p in base.rglob("*")
                          if p.suffix.lower() in (".xlsx", ".xlsm"))
    return out


def main(apenas_check: bool) -> None:
    feitas = puladas = ok = 0
    for p in _arquivos():
        status = _processar(p, apenas_check)
        print(f"  {p.parent.name:<36} {status}")
        if "OTIMIZADA" in status or "A OTIMIZAR" in status:
            feitas += 1
        elif "pulado" in status:
            puladas += 1
        else:
            ok += 1
    verbo = "a otimizar" if apenas_check else "otimizada(s)"
    print(f"\n{feitas} {verbo} | {ok} já OK | {puladas} pulada(s) (abertas)")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    main(apenas_check="--check" in sys.argv)
