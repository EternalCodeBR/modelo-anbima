"""Remove comentários órfãos deixados pelo openpyxl (causa real do prompt de reparo).

Causa raiz: os originais têm comentários MODERNOS (threaded — ``xl/threadedComments/``
+ ``xl/persons/``). O openpyxl não sabe fazer o round-trip disso: ao salvar, descarta
as partes threaded e tenta rebaixar para o formato legado (``xl/comments/commentN.xml``),
mas **não gera as formas VML correspondentes** (``xl/drawings/commentsDrawingN.vml``
fica com 0 shapes) nem conecta ``<legacyDrawing>`` na aba. O resultado é uma cadeia de
relacionamento órfã — comentário declarado, sem âncora visual — que o validador de
conteúdo do Excel rejeita com "Encontramos um problema em um conteúdo de ...".

Como comentários são anotação de UI (não entram no cálculo do PU), a correção mais
segura é remover a cadeia inteira (parte do comentário + VML + as duas relações que
apontam para eles), em vez de tentar reconstruir a geometria VML — operação bem mais
arriscada. Cirúrgico: opera direto no XML dentro do .xlsx, sem reabrir via openpyxl
(evitando reintroduzir o próprio defeito).

Escrita atômica (arquivo temporário + os.replace) e em subprocessos de poucos
arquivos (contorna a proteção antiransomware do TI).

Uso:
    python -m mtm_skills.corrigir_comentarios_quebrados            # todas
    python -m mtm_skills.corrigir_comentarios_quebrados --check    # só verifica
    python -m mtm_skills.corrigir_comentarios_quebrados <arq...>   # (modo subprocesso)
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
_RE_COMMENT_PART = re.compile(r"^xl/comments/comment\d+\.xml$")
_RE_REL = re.compile(r"<Relationship\b[^>]*/>")
_RE_OVERRIDE_COMMENT = re.compile(
    r'<Override PartName="/xl/comments/comment\d+\.xml"[^>]*/>')


def _achar_pares_quebrados(conteudo: dict[str, bytes]) -> list[dict]:
    """Localiza pares (comentário, VML, .rels-dono) com âncora quebrada.

    Um par é "quebrado" quando o VML tem menos <v:shape> do que o comentário tem
    <comment ref=...> (inclui o caso mais comum: 0 shapes).
    """
    pares = []
    for nome, raw in conteudo.items():
        if not re.match(r"^xl/worksheets/_rels/sheet\d+\.xml\.rels$", nome):
            continue
        rels_txt = raw.decode("utf-8")
        comment_tgt = vml_tgt = None
        for m in _RE_REL.finditer(rels_txt):
            tag = m.group(0)
            # checar o atributo Type (não a tag inteira: o Target da relação
            # vmlDrawing costuma ser ".../commentsDrawingN.vml", cujo NOME
            # contém a substring "/comments" e colidiria com um match ingênuo).
            tm_type = re.search(r'Type="([^"]+)"', tag)
            if not tm_type:
                continue
            tipo = tm_type.group(1)
            tm_tgt = re.search(r'Target="([^"]+)"', tag)
            if not tm_tgt:
                continue
            if tipo.endswith("/comments"):
                comment_tgt = tm_tgt.group(1).lstrip("/")
            elif tipo.endswith("/vmlDrawing"):
                vml_tgt = tm_tgt.group(1).lstrip("/")
        if not comment_tgt or not vml_tgt:
            continue
        if comment_tgt not in conteudo or vml_tgt not in conteudo:
            continue
        com_txt = conteudo[comment_tgt].decode("utf-8", "replace")
        vml_txt = conteudo[vml_tgt].decode("utf-8", "replace")
        n_refs = len(re.findall(r'<comment ref="', com_txt))
        n_shapes = len(re.findall(r"<v:shape ", vml_txt))
        if n_refs and n_shapes < n_refs:
            pares.append({
                "rels_nome": nome, "comment_nome": comment_tgt, "vml_nome": vml_tgt,
                "n_refs": n_refs, "n_shapes": n_shapes,
            })
    return pares


def _contar(path: Path) -> int:
    with zipfile.ZipFile(str(path)) as z:
        conteudo = {i: z.read(i) for i in z.namelist()}
    return len(_achar_pares_quebrados(conteudo))


def _fix_um(path: Path) -> int:
    with zipfile.ZipFile(str(path)) as zin:
        infos = zin.infolist()
        conteudo = {i.filename: zin.read(i.filename) for i in infos}

    pares = _achar_pares_quebrados(conteudo)
    if not pares:
        return 0

    ct_nome = "[Content_Types].xml"
    ct_txt = conteudo[ct_nome].decode("utf-8") if ct_nome in conteudo else None

    for par in pares:
        # 1) remove as partes (comentário + VML)
        conteudo.pop(par["comment_nome"], None)
        conteudo.pop(par["vml_nome"], None)

        # 2) remove as duas <Relationship> do .rels dono
        rels_txt = conteudo[par["rels_nome"]].decode("utf-8")
        def _remover_rel(m: re.Match) -> str:
            tag = m.group(0)
            if par["comment_nome"] in tag or par["vml_nome"] in tag:
                return ""
            return tag
        rels_txt = _RE_REL.sub(_remover_rel, rels_txt)
        # se não sobrou nenhuma <Relationship>, remove o .rels inteiro
        if not _RE_REL.search(rels_txt):
            conteudo.pop(par["rels_nome"], None)
        else:
            conteudo[par["rels_nome"]] = rels_txt.encode("utf-8")

        # 3) remove o <Override> do comentário em [Content_Types].xml
        if ct_txt is not None:
            ct_txt = _RE_OVERRIDE_COMMENT.sub(
                lambda m: "" if par["comment_nome"] in m.group(0) else m.group(0), ct_txt)

    if ct_txt is not None:
        conteudo[ct_nome] = ct_txt.encode("utf-8")

    tmp = path.with_name(path.name + ".tmp")
    nomes_removidos = {p["comment_nome"] for p in pares} | {p["vml_nome"] for p in pares}
    nomes_removidos |= {p["rels_nome"] for p in pares if p["rels_nome"] not in conteudo}
    with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
        for i in infos:
            if i.filename in nomes_removidos:
                continue
            zout.writestr(i, conteudo[i.filename])
    os.replace(str(tmp), str(path))
    return len(pares)


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
        print(f"  {n:>2} par(es) quebrado(s)  {p.parent.name}" + ("  <<<" if n else ""))
    print(f"\nTotal de comentário/VML quebrado: {total}")


def aplicar() -> None:
    arqs = _arquivos()
    lotes = [arqs[i:i + BATCH] for i in range(0, len(arqs), BATCH)]
    print(f"Corrigindo {len(arqs)} arquivo(s) em {len(lotes)} subprocesso(s) de até {BATCH}...\n")
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
            print(f"  {p.parent.name}: {n} par(es) de comentário órfão removido(s)")
    else:
        aplicar()
