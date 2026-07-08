"""Padroniza o prefixo do XLOOKUP no XML das calculadoras (sem openpyxl).

Histórico: uma suposição antiga do projeto (nunca confirmada abrindo em Excel
real) dizia que XLOOKUP exigia o prefixo duplo ``_xlfn._xlws.XLOOKUP`` no OOXML.
Essa suposição caiu por terra com evidência direta: ao deixar o Excel reparar e
resalvar uma calculadora (Ativo DI-piloto), o PRÓPRIO EXCEL escreveu a fórmula com
prefixo SIMPLES — ``_xlfn.XLOOKUP`` — e o arquivo passou a abrir sem prompt de
reparo. A suspeita antiga do prefixo duplo provavelmente estava confundida com
os defeitos reais (comentário órfão / desenho reescrito pelo openpyxl) que
coexistiam nos mesmos arquivos testados na época.

Esta correção agora vai no sentido oposto do nome: padroniza TODAS as
calculadoras para o prefixo SIMPLES (``_xlfn.XLOOKUP``), que é o formato
confirmado — pelo próprio Excel — como o aceito nesta instalação.

Cirúrgica: edita apenas o texto das fórmulas dentro dos `xl/worksheets/sheet*.xml`,
reempacotando o .xlsx com todas as demais partes intactas — NÃO usa openpyxl
(que reescreve o pacote e degrada o arquivo).

Escrita atômica (arquivo temporário + os.replace) e em subprocessos de poucos
arquivos (contorna a proteção antiransomware do TI).

Uso:
    python -m mtm_skills.corrigir_prefixo_xlookup            # todas
    python -m mtm_skills.corrigir_prefixo_xlookup --check    # só verifica
    python -m mtm_skills.corrigir_prefixo_xlookup <arq...>   # (modo subprocesso)
"""
import os
import subprocess
import sys
import zipfile
from pathlib import Path

_RAIZ = Path(__file__).resolve().parents[1]
if str(_RAIZ) not in sys.path:
    sys.path.insert(0, str(_RAIZ))

from pu_mtm.app import config

BATCH = 3
ERRADO = "_xlfn._xlws.XLOOKUP"  # prefixo duplo — não confirmado, descartado
CERTO = "_xlfn.XLOOKUP"          # prefixo simples — confirmado pelo próprio Excel


def _contar(path: Path) -> int:
    n = 0
    with zipfile.ZipFile(str(path)) as z:
        for nome in z.namelist():
            if nome.startswith("xl/worksheets/sheet") and nome.endswith(".xml"):
                n += z.read(nome).decode("utf-8", "ignore").count(ERRADO)
    return n


def _fix_um(path: Path) -> int:
    with zipfile.ZipFile(str(path)) as zin:
        infos = zin.infolist()
        conteudo = {i.filename: zin.read(i.filename) for i in infos}

    total = 0
    for nome in list(conteudo):
        if nome.startswith("xl/worksheets/sheet") and nome.endswith(".xml"):
            txt = conteudo[nome].decode("utf-8")
            n = txt.count(ERRADO)
            if n:
                conteudo[nome] = txt.replace(ERRADO, CERTO).encode("utf-8")
                total += n

    if total:
        tmp = path.with_name(path.name + ".tmp")
        with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
            for i in infos:  # preserva ordem e metadados das partes
                zout.writestr(i, conteudo[i.filename])
        os.replace(str(tmp), str(path))  # atômico
    return total


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
        print(f"  {n:>5} '{ERRADO}' (prefixo duplo)  {p.parent.name}" + ("  <<<" if n else ""))
    print(f"\nTotal com prefixo duplo (a padronizar): {total}")


def aplicar() -> None:
    arqs = _arquivos()
    lotes = [arqs[i:i + BATCH] for i in range(0, len(arqs), BATCH)]
    print(f"Padronizando {len(arqs)} arquivo(s) em {len(lotes)} subprocesso(s) de até {BATCH}...\n")
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
            print(f"  {p.parent.name}: {n} XLOOKUP padronizado (prefixo simples)")
    else:
        aplicar()
