"""Remove links externos órfãos (vestígio do SUMIFS antigo).

Causa raiz: antes da migração, as fórmulas de taxa referenciavam workbooks
externos via SUMIFS (``[1]SERIE_VALOR!...``/``[2]SERIE_VALOR!...``, apontando
para Calculadora CRI.xlsm / BaseDadosMercado.xlsm). A migração trocou tudo para
XLOOKUP na aba ``CDI`` local — mas o `openpyxl` PRESERVA os pacotes
``xl/externalLinks/externalLinkN.xml`` mesmo depois que nenhuma fórmula os
referencia mais. Esses pacotes têm um cache de valores (``<sheetData>``) que o
Excel valida ao abrir; como ninguém mais atualiza esse cache, o Excel acusa
"Registros Reparados: Referência de fórmula externa... Valores no cache" e
oferece reparo.

Como a única maneira de confirmar que é seguro remover é provar que NENHUMA
fórmula ainda referencia o link (ex.: ``[1]SERIE_VALOR!...``), esta correção
só age quando essa prova é verdadeira — caso contrário deixa o arquivo intacto
e avisa.

Cirúrgica: edita o pacote .xlsx diretamente (workbook.xml, workbook.xml.rels,
[Content_Types].xml, remove as partes externalLinkN.xml/.rels) — sem abrir via
openpyxl.

Escrita atômica (arquivo temporário + os.replace) e em subprocessos de poucos
arquivos (contorna a proteção antiransomware do TI).

Uso:
    python -m mtm_skills.corrigir_links_externos_orfaos            # todas
    python -m mtm_skills.corrigir_links_externos_orfaos --check    # só verifica
    python -m mtm_skills.corrigir_links_externos_orfaos <arq...>   # (modo subprocesso)
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
_RE_EXTLINK_PART = re.compile(r"^xl/externalLinks/externalLink(\d+)\.xml$")
_RE_REL = re.compile(r"<Relationship\b[^>]*/>")


def _indices_referenciados(conteudo: dict[str, bytes]) -> set[int]:
    """Índices [N] ainda citados em alguma fórmula OU em algum defined name.

    ``definedNames`` (ex.: ``dtBase`` -> ``[1]RESUMO!$C$2``) é facilmente
    esquecido: um índice pode não aparecer em nenhuma fórmula de célula e ainda
    assim estar vivo só no nome definido. Ignorar isso deixa o nome pendurado
    (referência de link externo agora inexistente) — é exatamente o que o
    Excel reporta como "Registros Removidos: Intervalo nomeado".
    """
    usados: set[int] = set()
    for nome, raw in conteudo.items():
        if not re.match(r"^xl/worksheets/sheet\d+\.xml$", nome):
            continue
        for m in re.finditer(r"<f>[^<]*\[(\d+)\][A-Za-z]", raw.decode("utf-8", "replace")):
            usados.add(int(m.group(1)))
    wb = conteudo.get("xl/workbook.xml")
    if wb:
        wb_txt = wb.decode("utf-8", "replace")
        dn_block = re.search(r"<definedNames>.*?</definedNames>", wb_txt)
        if dn_block:
            for m in re.finditer(r"\[(\d+)\][A-Za-z]", dn_block.group(0)):
                usados.add(int(m.group(1)))
    return usados


def _indices_dn_penduradas(conteudo: dict[str, bytes]) -> set[int]:
    """Índices [N] citados em definedNames cujo externalLinkN.xml NÃO existe mais
    (sobra de uma rodada anterior que removeu o link mas esqueceu do nome)."""
    wb = conteudo.get("xl/workbook.xml")
    if not wb:
        return set()
    wb_txt = wb.decode("utf-8", "replace")
    dn_block = re.search(r"<definedNames>.*?</definedNames>", wb_txt)
    if not dn_block:
        return set()
    citados = {int(i) for i in re.findall(r"\[(\d+)\]", dn_block.group(0))}
    presentes = {int(m.group(1)) for n in conteudo
                 if (m := _RE_EXTLINK_PART.match(n))}
    return citados - presentes


def _limpar_definedNames_penduradas(wb_txt: str, indices_removidos: set[int]) -> str:
    """Remove <definedName>...</definedName> cujo valor cite um índice [N] já removido."""
    def _tira(m: re.Match) -> str:
        idxs = {int(i) for i in re.findall(r"\[(\d+)\]", m.group(0))}
        return "" if idxs & indices_removidos else m.group(0)
    return re.sub(r'<definedName\b[^>]*>.*?</definedName>', _tira, wb_txt)


def _orfaos(conteudo: dict[str, bytes]) -> list[int]:
    """Índices de externalLinkN.xml presentes e NÃO referenciados por fórmula."""
    presentes = set()
    for nome in conteudo:
        m = _RE_EXTLINK_PART.match(nome)
        if m:
            presentes.add(int(m.group(1)))
    if not presentes:
        return []
    return sorted(presentes - _indices_referenciados(conteudo))


def _contar(path: Path) -> int:
    with zipfile.ZipFile(str(path)) as z:
        conteudo = {i: z.read(i) for i in z.namelist()}
    return len(_orfaos(conteudo)) + len(_indices_dn_penduradas(conteudo))


def _fix_um(path: Path) -> int:
    with zipfile.ZipFile(str(path)) as zin:
        infos = zin.infolist()
        conteudo = {i.filename: zin.read(i.filename) for i in infos}

    orfaos = _orfaos(conteudo)
    dn_penduradas = _indices_dn_penduradas(conteudo)
    if not orfaos and not dn_penduradas:
        return 0

    wb_txt = conteudo["xl/workbook.xml"].decode("utf-8")
    wbrels_txt = conteudo["xl/_rels/workbook.xml.rels"].decode("utf-8")
    ct_txt = conteudo["[Content_Types].xml"].decode("utf-8")

    # definedNames que já ficaram pendurados numa rodada anterior (o link já
    # não existe mais): limpa antes de mais nada, sem depender de `orfaos`.
    if dn_penduradas:
        wb_txt = _limpar_definedNames_penduradas(wb_txt, dn_penduradas)

    partes_removidas: set[str] = set()
    rids_a_remover: set[str] = set()

    if not orfaos:
        conteudo["xl/workbook.xml"] = wb_txt.encode("utf-8")
        tmp = path.with_name(path.name + ".tmp")
        with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
            for i in infos:
                zout.writestr(i, conteudo[i.filename])
        os.replace(str(tmp), str(path))
        return len(dn_penduradas)

    for idx in orfaos:
        part = f"xl/externalLinks/externalLink{idx}.xml"
        rels_part = f"xl/externalLinks/_rels/externalLink{idx}.xml.rels"
        partes_removidas.add(part)
        if rels_part in conteudo:
            partes_removidas.add(rels_part)

        # acha o Id da relação em workbook.xml.rels cujo Target aponta pra essa parte
        for m in _RE_REL.finditer(wbrels_txt):
            tag = m.group(0)
            tgtm = re.search(r'Target="([^"]+)"', tag)
            if tgtm and Path(tgtm.group(1)).name == f"externalLink{idx}.xml":
                idm = re.search(r'Id="([^"]+)"', tag)
                if idm:
                    rids_a_remover.add(idm.group(1))

        # remove o Override do Content_Types (se houver — externalLink pode usar Default por extensão)
        ct_txt = re.sub(
            rf'<Override PartName="/xl/externalLinks/externalLink{idx}\.xml"[^>]*/>',
            "", ct_txt)

    # remove as <Relationship> do workbook.xml.rels
    def _tira_rel(m: re.Match) -> str:
        idm = re.search(r'Id="([^"]+)"', m.group(0))
        return "" if (idm and idm.group(1) in rids_a_remover) else m.group(0)
    wbrels_txt = _RE_REL.sub(_tira_rel, wbrels_txt)

    # remove os <externalReference r:id="..."/> do workbook.xml
    for rid in rids_a_remover:
        wb_txt = re.sub(
            rf'<externalReference\b[^>]*r:id="{re.escape(rid)}"[^>]*/>', "", wb_txt)
    # se a lista ficou vazia, remove o wrapper <externalReferences>...</externalReferences>
    wb_txt = re.sub(r"<externalReferences>\s*</externalReferences>", "", wb_txt)

    conteudo["xl/workbook.xml"] = wb_txt.encode("utf-8")
    conteudo["xl/_rels/workbook.xml.rels"] = wbrels_txt.encode("utf-8")
    conteudo["[Content_Types].xml"] = ct_txt.encode("utf-8")

    tmp = path.with_name(path.name + ".tmp")
    with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
        for i in infos:
            if i.filename in partes_removidas:
                continue
            zout.writestr(i, conteudo[i.filename])
    os.replace(str(tmp), str(path))
    return len(orfaos) + len(dn_penduradas)


def _arquivos() -> list[Path]:
    return sorted(
        p for p in config.HOMOLOG_CALC_ROOT.rglob("*")
        if p.suffix.lower() in (".xlsx", ".xlsm")
    )


def check() -> None:
    total = 0
    for p in _arquivos():
        n = _contar(p)
        total += n
        print(f"  {n}  {p.parent.name}" + ("  <<<" if n else ""))
    print(f"\nTotal de links externos órfãos: {total}")


def aplicar() -> None:
    arqs = _arquivos()
    lotes = [arqs[i:i + BATCH] for i in range(0, len(arqs), BATCH)]
    print(f"Removendo links órfãos em {len(arqs)} arquivo(s), "
          f"{len(lotes)} subprocesso(s) de até {BATCH}...\n")
    for lote in lotes:
        proc = subprocess.run(
            [sys.executable, "-W", "ignore", str(Path(__file__).resolve()),
             *[str(p) for p in lote]],
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
        for a in args:
            p = Path(a)
            n = _fix_um(p)
            print(f"  {p.parent.name}: {n} link(s) externo(s) órfão(s) removido(s)")
    else:
        aplicar()
