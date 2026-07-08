"""Corrige datas de CDI gravadas com HORA (serial fracionário) — causa de #N/D.

Defeito: quando o CDI foi escrito via Excel/pywin32, a conversão datetime→COM
adicionou o fuso (Brasil UTC-3) ao serial, gravando ex.: ``46199.125`` (26/06 às
03:00) em vez de ``46199`` (26/06). O XLOOKUP faz match EXATO: a data do
cronograma (serial inteiro) não acha o serial fracionário → ``#N/D``.

Correção CIRÚRGICA (sem Excel, sem openpyxl): na coluna A da aba CDI, arredonda
todo serial fracionário para o inteiro correspondente (a hora era só o offset de
fuso; ``round`` recupera o dia certo). Só reescreve o arquivo se houver o que
corrigir; reempacota o .xlsx com todas as demais partes byte-a-byte idênticas.

Se um arquivo estiver ABERTO no Excel, é PULADO (não trava o lote nem mexe no
que você está usando).

Uso:
    python -m mtm_skills.corrigir_hora_cdi            # corrige todos os afetados
    python -m mtm_skills.corrigir_hora_cdi --check    # só relata, não grava
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

from mtm_skills.atualizar_cdi_surgical import _cdi_part
from pu_mtm.app import config

# Célula da coluna A com serial fracionário (tem hora).
_RE_A_FRAC = re.compile(r'(<c r="A\d+"[^>]*><v>)(\d+\.\d+)(</v>)')


def _corrigir_xml(xml: str) -> tuple[str, int]:
    n = 0

    def _repl(m: re.Match) -> str:
        nonlocal n
        n += 1
        return f"{m.group(1)}{round(float(m.group(2)))}{m.group(3)}"

    return _RE_A_FRAC.sub(_repl, xml), n


def _processar(path: Path, apenas_check: bool) -> str:
    try:
        with zipfile.ZipFile(str(path)) as zin:
            infos = zin.infolist()
            conteudo = {i.filename: zin.read(i.filename) for i in infos}
    except PermissionError:
        return "ABERTO no Excel — pulado"
    except zipfile.BadZipFile:
        return "ZIP inválido — pulado"

    part = _cdi_part(conteudo)
    if not part or part not in conteudo:
        return "sem aba CDI"

    xml = conteudo[part].decode("utf-8")
    novo, n = _corrigir_xml(xml)
    if n == 0:
        return "OK (nenhuma data com hora)"
    if apenas_check:
        return f"{n} data(s) com hora (a corrigir)"

    conteudo[part] = novo.encode("utf-8")
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
    return f"CORRIGIDO: {n} data(s)"


def _arquivos() -> list[Path]:
    out: list[Path] = []
    for label in ("CDI", "CDI + Spread"):
        base = config.HOMOLOG_CALC_ROOT / label
        if base.exists():
            out += sorted(p for p in base.rglob("*")
                          if p.suffix.lower() in (".xlsx", ".xlsm"))
    return out


def main(apenas_check: bool) -> None:
    corrigidos = pulados = limpos = 0
    for p in _arquivos():
        status = _processar(p, apenas_check)
        print(f"  {p.parent.name:<36} {status}")
        if "CORRIGIDO" in status or "a corrigir" in status:
            corrigidos += 1
        elif "pulado" in status:
            pulados += 1
        else:
            limpos += 1
    verbo = "a corrigir" if apenas_check else "corrigido(s)"
    print(f"\n{corrigidos} {verbo} | {limpos} já OK | {pulados} pulado(s) (abertos)")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    main(apenas_check="--check" in sys.argv)
