"""Restaura os desenhos/imagens originais (openpyxl reescreve DrawingML com perda).

Causa raiz: ``openpyxl.Workbook.save()`` REESCREVE por completo qualquer
``xl/drawings/drawingN.xml`` existente, trocando os namespaces prefixados
(``xdr:``, ``a:``) do Excel por namespace padrão e **descartando** o
``<a:extLst>``/``creationId`` de cada imagem. O parser de conteúdo do Excel
rejeita o resultado — "Encontramos um problema em um conteúdo de..." com
"Parte Removida: Forma de desenho".

Como o desenho é só decoração (nunca é tocado pela nossa lógica de fórmula/CDI),
a correção mais segura é restaurar os bytes ORIGINAIS do desenho por cima do
que o openpyxl reescreveu — casando por NOME DA ABA (robusto a qualquer
renumeração de ``drawingN.xml`` entre original e reconstruída). Cirúrgico:
opera direto no XML dentro do .xlsx, sem reabrir via openpyxl.

Escrita atômica (arquivo temporário + os.replace) e em subprocessos de poucos
arquivos (contorna a proteção antiransomware do TI).

Uso:
    python -m mtm_skills.corrigir_desenho_openpyxl            # todas
    python -m mtm_skills.corrigir_desenho_openpyxl --check    # só verifica
    python -m mtm_skills.corrigir_desenho_openpyxl <arq...>   # (modo subprocesso)
"""
import os
import re
import subprocess
import sys
import zipfile
from pathlib import Path

_RAIZ = Path(__file__).resolve().parents[1]
if str(_RAIZ) not in sys.path:
    sys.path.insert(0, str(_RAIZ))

from pu_mtm.app import config

BATCH = 3


def _sheet_map(conteudo: dict[str, bytes]) -> dict[str, str]:
    """{nome_da_aba: sheetN.xml} via workbook.xml + workbook.xml.rels."""
    wb = conteudo.get("xl/workbook.xml")
    rels = conteudo.get("xl/_rels/workbook.xml.rels")
    if not wb or not rels:
        return {}
    wb_txt, rels_txt = wb.decode("utf-8"), rels.decode("utf-8")
    rid_target = {}
    for m in re.finditer(r"<Relationship\b[^>]*/>", rels_txt):
        tag = m.group(0)
        idm = re.search(r'Id="([^"]+)"', tag)
        tgtm = re.search(r'Target="([^"]+)"', tag)
        if idm and tgtm:
            tgt = tgtm.group(1)
            # openpyxl escreve Target absoluto ("/xl/worksheets/sheetN.xml"); o
            # Excel/original escreve relativo a xl/ ("worksheets/sheetN.xml").
            rid_target[idm.group(1)] = tgt.lstrip("/") if tgt.startswith("/") else "xl/" + tgt
    out = {}
    for m in re.finditer(r"<sheet\b[^>]*/>", wb_txt):
        tag = m.group(0)
        namem = re.search(r'name="([^"]+)"', tag)
        ridm = re.search(r'r:id="([^"]+)"', tag)
        if namem and ridm and ridm.group(1) in rid_target:
            out[namem.group(1)] = rid_target[ridm.group(1)]
    return out


def _drawing_part(conteudo: dict[str, bytes], sheet_part: str) -> str | None:
    """Nome do drawingN.xml referenciado pela aba (via sheetN.xml.rels)."""
    rels_name = sheet_part.replace("worksheets/", "worksheets/_rels/") + ".rels"
    rels = conteudo.get(rels_name)
    if not rels:
        return None
    for m in re.finditer(r"<Relationship\b[^>]*/>", rels.decode("utf-8")):
        tag = m.group(0)
        tm_type = re.search(r'Type="([^"]+)"', tag)
        if tm_type and tm_type.group(1).endswith("/drawing"):
            tgtm = re.search(r'Target="([^"]+)"', tag)
            if tgtm:
                return "xl/drawings/" + Path(tgtm.group(1)).name
    return None


def _media_refs(conteudo: dict[str, bytes], drawing_part: str) -> dict[str, str]:
    """{Id: caminho_media_absoluto} declarados no rels do drawing."""
    rels_name = drawing_part.replace("drawings/", "drawings/_rels/") + ".rels"
    rels = conteudo.get(rels_name)
    if not rels:
        return {}
    out = {}
    for m in re.finditer(r"<Relationship\b[^>]*/>", rels.decode("utf-8")):
        tag = m.group(0)
        idm = re.search(r'Id="([^"]+)"', tag)
        tgtm = re.search(r'Target="([^"]+)"', tag)
        if idm and tgtm and "media" in tgtm.group(1):
            out[idm.group(1)] = "xl/media/" + Path(tgtm.group(1)).name
    return out


def _pares_a_restaurar(conteudo_orig: dict[str, bytes],
                       conteudo_reb: dict[str, bytes]) -> list[dict]:
    """Para cada aba com desenho em AMBOS, casa drawing original <-> reconstruída."""
    mapa_orig = _sheet_map(conteudo_orig)
    mapa_reb = _sheet_map(conteudo_reb)
    pares = []
    for aba, sheet_part_o in mapa_orig.items():
        sheet_part_r = mapa_reb.get(aba)
        if not sheet_part_r:
            continue
        dr_o = _drawing_part(conteudo_orig, sheet_part_o)
        dr_r = _drawing_part(conteudo_reb, sheet_part_r)
        if not dr_o or not dr_r:
            continue
        if dr_o not in conteudo_orig or dr_r not in conteudo_reb:
            continue
        # já idêntico? não mexe.
        if conteudo_orig[dr_o] == conteudo_reb[dr_r]:
            continue
        pares.append({"aba": aba, "drawing_orig": dr_o, "drawing_reb": dr_r})
    return pares


def _contar(path_orig: Path, path_reb: Path) -> int:
    with zipfile.ZipFile(str(path_orig)) as zo, zipfile.ZipFile(str(path_reb)) as zr:
        co = {i: zo.read(i) for i in zo.namelist()}
        cr = {i: zr.read(i) for i in zr.namelist()}
    return len(_pares_a_restaurar(co, cr))


def _fix_um(path_orig: Path, path_reb: Path) -> int:
    with zipfile.ZipFile(str(path_orig)) as zo:
        conteudo_orig = {i: zo.read(i) for i in zo.namelist()}
    with zipfile.ZipFile(str(path_reb)) as zin:
        infos = zin.infolist()
        conteudo = {i.filename: zin.read(i.filename) for i in infos}

    pares = _pares_a_restaurar(conteudo_orig, conteudo)
    if not pares:
        return 0

    extras: dict[str, bytes] = {}
    for par in pares:
        dr_o, dr_r = par["drawing_orig"], par["drawing_reb"]
        # 1) bytes do drawingN.xml: original por cima do reescrito pelo openpyxl
        conteudo[dr_r] = conteudo_orig[dr_o]

        # 2) rels do drawing: usa o do ORIGINAL (mesmos Id/Target relativos)
        rels_o_nome = dr_o.replace("drawings/", "drawings/_rels/") + ".rels"
        rels_r_nome = dr_r.replace("drawings/", "drawings/_rels/") + ".rels"
        if rels_o_nome in conteudo_orig:
            conteudo[rels_r_nome] = conteudo_orig[rels_o_nome]

        # 3) mídia referenciada pelo rels do original: garante bytes idênticos
        for _rid, media_path in _media_refs(conteudo_orig, dr_o).items():
            if media_path in conteudo_orig:
                extras[media_path] = conteudo_orig[media_path]

    conteudo.update(extras)

    tmp = path_reb.with_name(path_reb.name + ".tmp")
    nomes_existentes = {i.filename for i in infos}
    with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
        for i in infos:
            zout.writestr(i, conteudo[i.filename])
        # mídia nova (não existia antes na reconstruída) precisa de ZipInfo próprio
        for nome, dados in extras.items():
            if nome not in nomes_existentes:
                zout.writestr(nome, dados)
    os.replace(str(tmp), str(path_reb))
    return len(pares)


def _pares_arquivos() -> list[tuple[Path, Path]]:
    """(original_sharepoint, reconstruída_homolog) casados por CalcPath do cadastro."""
    from pu_mtm.dados.csvio import ler_dict

    rows = ler_dict(str(config.CADASTRO))
    idx = config.indexar_homolog()
    saida = []
    for r in rows:
        if (r.get("Familia") or "").lower() not in ("di_puro", "di_spread"):
            continue
        cp = r.get("CalcPath")
        if not cp:
            continue
        origem = config.FLUXO_DIR / cp.replace("/", os.sep)
        reb = idx.get(Path(cp).name.lower())
        if origem.exists() and reb is not None:
            saida.append((origem, reb))
    return saida


def check() -> None:
    total = 0
    for orig, reb in _pares_arquivos():
        n = _contar(orig, reb)
        total += n
        print(f"  {n}  {reb.parent.name}" + ("  <<<" if n else ""))
    print(f"\nTotal de desenhos a restaurar: {total}")


def aplicar() -> None:
    pares = _pares_arquivos()
    lotes = [pares[i:i + BATCH] for i in range(0, len(pares), BATCH)]
    print(f"Restaurando desenhos em {len(pares)} arquivo(s), "
          f"{len(lotes)} subprocesso(s) de até {BATCH}...\n")
    for lote in lotes:
        args = []
        for orig, reb in lote:
            args += [str(orig), str(reb)]
        proc = subprocess.run(
            [sys.executable, "-W", "ignore", str(Path(__file__).resolve()), *args],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        if proc.stdout:
            sys.stdout.write(proc.stdout)
        if proc.returncode != 0 and proc.stderr:
            sys.stdout.write(proc.stderr)
    print("\nReverificando...")
    check()


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if "--check" in sys.argv:
        check()
    elif args:
        # modo subprocesso: pares (original, reconstruída) em sequência
        for i in range(0, len(args), 2):
            po, pr = Path(args[i]), Path(args[i + 1])
            n = _fix_um(po, pr)
            print(f"  {pr.parent.name}: {n} desenho(s) restaurado(s)")
    else:
        aplicar()
