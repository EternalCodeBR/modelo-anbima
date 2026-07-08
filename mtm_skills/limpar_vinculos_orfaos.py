"""Remove vínculos externos ÓRFÃOS das calculadoras (flag `estrut:ext` do validador).

Órfão = o pacote tem `xl/externalLinks/externalLinkN.xml`, mas NENHUMA fórmula viva
referencia `[n]…!`. São links mortos que sobraram (ex.: BaseDadosMercado, Calculadora
CRI). Removê-los NÃO muda resultado — não precisa recalcular.

TRAVA DE SEGURANÇA: se alguma fórmula ainda usa `[n]…!` (vínculo VIVO, como o 765 que
lia `[2]CDI!` morto e congelava), o arquivo é PULADO — remover quebraria o cálculo.
Esse caso pede migração para fonte local (não limpeza cega).

Correção cirúrgica (sem Excel/openpyxl): remove as partes externalLinks + o bloco
`<externalReferences>` + as relações/Overrides + os nomes definidos que apontam para
`[n]`. calcChain é preservado (não mexemos em fórmula). Reempacota byte-a-byte.

Arquivo ABERTO no Excel é PULADO.

Uso:
    python -m mtm_skills.limpar_vinculos_orfaos            # aplica
    python -m mtm_skills.limpar_vinculos_orfaos --check    # só relata
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

_RE_WS = re.compile(r"xl/worksheets/sheet\d+\.xml$")
_RE_EXTLINK = re.compile(r"xl/externalLinks/")
_RE_REF_EXTERNA = re.compile(r"\[\d+\][A-Za-z0-9_ ]+!")  # [n]Aba! em fórmula


def _refs_vivas(cont: dict[str, bytes]) -> int:
    n = 0
    for nome, raw in cont.items():
        if _RE_WS.match(nome):
            for f in re.findall(r"<f[^>]*>([^<]*)</f>", raw.decode("utf-8", "replace")):
                n += len(_RE_REF_EXTERNA.findall(f))
    return n


def _tem_ext(cont: dict[str, bytes]) -> bool:
    return any(_RE_EXTLINK.match(n) for n in cont)


def _limpar(cont: dict[str, bytes], infos: list) -> tuple[dict, list, int]:
    # workbook.xml: remove <externalReferences> e definedNames com [n]
    wb = cont["xl/workbook.xml"].decode("utf-8")
    wb = re.sub(r"<externalReferences>.*?</externalReferences>", "", wb, flags=re.S)
    wb = re.sub(r"<definedName\b[^>]*>[^<]*\[\d+\][^<]*</definedName>", "", wb)
    cont["xl/workbook.xml"] = wb.encode("utf-8")

    # workbook.xml.rels: remove relações de externalLink
    rels = cont["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rels = re.sub(r'<Relationship[^>]*Target="externalLinks/[^"]+"[^>]*/>', "", rels)
    cont["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    # [Content_Types].xml: remove Overrides de externalLink
    ct = cont["[Content_Types].xml"].decode("utf-8")
    ct = re.sub(r'<Override PartName="/xl/externalLinks/[^"]+"[^>]*/>', "", ct)
    cont["[Content_Types].xml"] = ct.encode("utf-8")

    # remove as partes físicas externalLinks (+ seus _rels)
    remover = {n for n in list(cont) if _RE_EXTLINK.match(n)}
    infos = [i for i in infos if i.filename not in remover]
    for r in remover:
        cont.pop(r, None)
    return cont, infos, len(remover)


def _processar(path: Path, apenas_check: bool) -> str:
    try:
        with zipfile.ZipFile(str(path)) as z:
            infos = z.infolist()
            cont = {i.filename: z.read(i.filename) for i in infos}
    except PermissionError:
        return "ABERTO no Excel — pulado"
    except zipfile.BadZipFile:
        return "ZIP inválido — pulado"

    if not _tem_ext(cont):
        return "sem vínculo externo"

    vivas = _refs_vivas(cont)
    if vivas > 0:
        return f"PULADO: {vivas} ref externa VIVA em fórmula (precisa migrar, não limpar)"

    n_partes = sum(1 for n in cont if _RE_EXTLINK.match(n)
                   and re.match(r"xl/externalLinks/externalLink\d+\.xml$", n))
    if apenas_check:
        return f"A LIMPAR: {n_partes} vínculo(s) órfão(s)"

    cont, infos, _ = _limpar(cont, infos)
    tmp = path.with_name(path.name + ".tmp")
    try:
        with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
            for i in infos:
                zout.writestr(i, cont[i.filename])
        os.replace(str(tmp), str(path))
    except PermissionError:
        if tmp.exists():
            tmp.unlink()
        return "ABERTO no Excel — pulado"
    return f"LIMPO: {n_partes} vínculo(s) órfão(s) removido(s)"


def _arquivos() -> list[Path]:
    out: list[Path] = []
    for label in ("CDI", "CDI + Spread"):
        base = config.HOMOLOG_CALC_ROOT / label
        if base.exists():
            out += sorted(p for p in base.rglob("*")
                          if p.suffix.lower() in (".xlsx", ".xlsm"))
    return out


def main(apenas_check: bool) -> None:
    limpos = pulados = sem = 0
    for p in _arquivos():
        status = _processar(p, apenas_check)
        if "sem vínculo" in status:
            sem += 1
            continue
        print(f"  {p.parent.name:<36} {status}")
        if "LIMPO" in status or "A LIMPAR" in status:
            limpos += 1
        elif "PULADO" in status or "pulado" in status:
            pulados += 1
    verbo = "a limpar" if apenas_check else "limpo(s)"
    print(f"\n{limpos} {verbo} | {pulados} pulado(s) | {sem} já sem vínculo")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    main(apenas_check="--check" in sys.argv)
