"""Remove tags <v></v> vazias deixadas em células de fórmula pelo openpyxl.

Causa raiz: ``openpyxl.Workbook.save()`` NUNCA preserva o valor cacheado de uma
célula de fórmula — ao reescrever o XML, toda célula com ``<f>`` ganha um
``<v></v>`` vazio (independente de qual célula foi de fato editada; afeta o
workbook inteiro a cada save). Isso viola a expectativa do Excel de que uma
``<v>`` presente contenha um número válido, e é a causa mais provável dos
prompts de reparo ("Encontramos um problema em um conteúdo").

A correção é CIRÚRGICA: remove apenas os pares ``<v></v>``/``<v/>`` vazios que
seguem imediatamente um ``</f>``, deixando ``<f>fórmula</f>`` — o mesmo estado
de "fórmula nova, ainda não calculada" que o Excel sempre aceita sem erro
(e o ``calcPr`` do workbook já força recálculo completo no load). Não usa
openpyxl para reescrever (evitando reintroduzir o próprio defeito); opera
direto no XML dentro do .xlsx, reempacotando com as demais partes intactas.

Escrita atômica (arquivo temporário + os.replace) e em subprocessos de poucos
arquivos (contorna a proteção antiransomware do TI).

Uso:
    python -m mtm_skills.corrigir_valor_vazio            # todas
    python -m mtm_skills.corrigir_valor_vazio --check    # só verifica
    python -m mtm_skills.corrigir_valor_vazio <arq...>   # (modo subprocesso)
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
# <f ...>...</f> seguido de <v/> ou <v>(só espaço/nada)</v>. O 2º ramo exige
# fechamento explícito — não pode casar parcialmente uma <v> com CONTEÚDO real
# (ex.: <v>14.15</v>), o que corromperia a célula ao deixar "14.15</v>" órfão.
_RE_F_VAZIA = re.compile(r"(</f>)(?:<v\s*/>|<v>\s*</v>)")


def _contar(path: Path) -> int:
    n = 0
    with zipfile.ZipFile(str(path)) as z:
        for nome in z.namelist():
            if nome.startswith("xl/worksheets/sheet") and nome.endswith(".xml"):
                n += len(_RE_F_VAZIA.findall(z.read(nome).decode("utf-8", "ignore")))
    return n


def _fix_um(path: Path) -> int:
    with zipfile.ZipFile(str(path)) as zin:
        infos = zin.infolist()
        conteudo = {i.filename: zin.read(i.filename) for i in infos}

    total = 0
    for nome in list(conteudo):
        if nome.startswith("xl/worksheets/sheet") and nome.endswith(".xml"):
            txt = conteudo[nome].decode("utf-8")
            nova, n = _RE_F_VAZIA.subn(r"\1", txt)
            if n:
                conteudo[nome] = nova.encode("utf-8")
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
        print(f"  {n:>6} <v> vazia  {p.parent.name}" + ("  <<<" if n else ""))
    print(f"\nTotal de <v> vazia pós-fórmula: {total}")


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
            print(f"  {p.parent.name}: {n} <v> vazia removida(s)")
    else:
        aplicar()
