"""Atualiza a aba CDI das calculadoras SEM Excel e SEM openpyxl (cirúrgico no XML).

O pywin32 (Excel) preserva tudo, mas é lento: abrir + recalcular + salvar 23
arquivos no OneDrive leva minutos. O openpyxl é rápido mas corrompe VML/desenhos.
Este módulo é o meio-termo correto: anexa as datas de CDI que faltam DIRETO no
``<sheetData>`` da aba CDI (que é só dado — sem desenho/VML/comentário), e
reempacota o .xlsx com TODAS as demais partes byte-a-byte idênticas. Milissegundos
por arquivo, zero corrupção.

Não recalcula o PU em cache (isso exigiria o Excel). Mas:
  * o motor Python calcula o PU de forma independente (é a fonte de verdade);
  * ao abrir a calculadora no Excel, o ``fullCalcOnLoad`` recalcula sozinho.

Incremental: lê a última data já na aba e anexa só o rabo (normalmente 1 linha).

Uso:
    python -m mtm_skills.atualizar_cdi_surgical            # todos
    python -m mtm_skills.atualizar_cdi_surgical --check    # só relata, não grava
    python -m mtm_skills.atualizar_cdi_surgical --ids 531  # um ativo
"""
from __future__ import annotations

import os
import re
import sys
import warnings
import zipfile
from datetime import date, timedelta
from pathlib import Path

warnings.filterwarnings("ignore")

_RAIZ = Path(__file__).resolve().parents[1]
if str(_RAIZ) not in sys.path:
    sys.path.insert(0, str(_RAIZ))

from mtm_skills.atualizar_cdi_calculadoras import (
    ler_cdi_mercado, ler_calculadoras_cdi, indexar_homolog, resolver_calc,
    _normalizar_data, MERCADO_XLSX, CADASTRO,
)

_EPOCH = date(1899, 12, 30)  # base do serial do Excel (datas pós-1900-03 são exatas)


def _serial(d: date) -> int:
    return (d - _EPOCH).days


def _from_serial(n: int) -> date:
    return _EPOCH + timedelta(days=int(n))


def _num(v: float) -> str:
    """Número como o Excel gravaria (sem notação científica, round-trip curto)."""
    return repr(float(v))


# ---------------------------------------------------------------------------
# Localização da aba CDI dentro do pacote
# ---------------------------------------------------------------------------
def _cdi_part(conteudo: dict[str, bytes]) -> str | None:
    wb = conteudo.get("xl/workbook.xml")
    rels = conteudo.get("xl/_rels/workbook.xml.rels")
    if not wb or not rels:
        return None
    wb_txt, rels_txt = wb.decode("utf-8"), rels.decode("utf-8")
    rid = None
    for m in re.finditer(r"<sheet\b[^>]*/>", wb_txt):
        tag = m.group(0)
        if re.search(r'name="CDI"', tag, re.I):
            rm = re.search(r'r:id="([^"]+)"', tag)
            rid = rm.group(1) if rm else None
            break
    if not rid:
        return None
    for m in re.finditer(r"<Relationship\b[^>]*/>", rels_txt):
        tag = m.group(0)
        if re.search(rf'Id="{re.escape(rid)}"', tag):
            tgt = re.search(r'Target="([^"]+)"', tag)
            if tgt:
                t = tgt.group(1)
                return t.lstrip("/") if t.startswith("/") else "xl/" + t
    return None


def _estilo_col(sheet_xml: str, col: str) -> str:
    """Atributo de estilo (' s="N"') de uma célula da coluna dada, ou ''."""
    m = re.search(rf'<c r="{col}\d+"([^>]*)>', sheet_xml)
    if m:
        sm = re.search(r'\ss="\d+"', m.group(1))
        return sm.group(0) if sm else ""
    return ""


def _ultima_data(sheet_xml: str) -> tuple[str | None, int]:
    """(iso_da_última_data, número_da_última_linha) lendo a coluna A."""
    ult_row = 1
    ult_serial = None
    for m in re.finditer(r'<c r="A(\d+)"[^>]*>(?:<v>([^<]*)</v>)?', sheet_xml):
        r = int(m.group(1))
        val = m.group(2)
        if r >= 2 and val:
            if r > ult_row:
                ult_row = r
                ult_serial = val
    if ult_serial is None:
        return None, 1
    try:
        d = _from_serial(float(ult_serial))
        return d.strftime("%Y-%m-%d"), ult_row
    except ValueError:
        return None, ult_row


# ---------------------------------------------------------------------------
# Núcleo
# ---------------------------------------------------------------------------
def _plano(sheet_xml: str, cdi_hist: dict, emissao_iso: str) -> list[tuple[str, tuple]]:
    """Datas a anexar: [(iso, (cdi, fator)), ...] após a última já presente."""
    ult_iso, _ = _ultima_data(sheet_xml)
    corte = ult_iso if ult_iso else None
    plano = []
    for d in sorted(cdi_hist):
        if d < emissao_iso:
            continue
        if corte is not None and d <= corte:
            continue
        plano.append((d, cdi_hist[d]))
    return plano


def _contar(path: Path, cdi_hist: dict, emissao_iso: str) -> int:
    with zipfile.ZipFile(str(path)) as z:
        conteudo = {i: z.read(i) for i in z.namelist()}
    part = _cdi_part(conteudo)
    if not part or part not in conteudo:
        return -1  # sem aba CDI
    return len(_plano(conteudo[part].decode("utf-8"), cdi_hist, emissao_iso))


def _fix_um(path: Path, cdi_hist: dict, emissao_iso: str) -> tuple[int, str]:
    with zipfile.ZipFile(str(path)) as zin:
        infos = zin.infolist()
        conteudo = {i.filename: zin.read(i.filename) for i in infos}

    part = _cdi_part(conteudo)
    if not part or part not in conteudo:
        return -1, "sem aba CDI (precisa criar)"

    xml = conteudo[part].decode("utf-8")
    ult_iso, ult_row = _ultima_data(xml)
    plano = _plano(xml, cdi_hist, emissao_iso)
    if not plano:
        return 0, f"em dia (última {ult_iso})"

    s_a = _estilo_col(xml, "A")  # estilo de data (ex.: ' s="1"')
    s_b = _estilo_col(xml, "B")
    s_c = _estilo_col(xml, "C")

    linhas = []
    r = ult_row
    for iso, (cdi, fat) in plano:
        r += 1
        d = date.fromisoformat(iso)
        cel_a = f'<c r="A{r}"{s_a}><v>{_serial(d)}</v></c>'
        cel_b = f'<c r="B{r}"{s_b}><v>{_num(cdi)}</v></c>'
        cel_c = (f'<c r="C{r}"{s_c}><v>{_num(fat)}</v></c>'
                 if fat not in (None, "") else f'<c r="C{r}"{s_c}/>')
        linhas.append(f'<row r="{r}" spans="1:3">{cel_a}{cel_b}{cel_c}</row>')
    bloco = "".join(linhas)

    # insere antes de </sheetData>
    if "</sheetData>" in xml:
        xml = xml.replace("</sheetData>", bloco + "</sheetData>", 1)
    else:  # sheetData vazio: <sheetData/>
        xml = xml.replace("<sheetData/>", "<sheetData>" + bloco + "</sheetData>", 1)

    # atualiza <dimension ref="A1:C{r}"/>
    xml = re.sub(r'(<dimension ref="A1:)[A-Z]+\d+(")', rf'\g<1>C{r}\g<2>', xml, count=1)

    conteudo[part] = xml.encode("utf-8")
    tmp = path.with_name(path.name + ".tmp")
    with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
        for i in infos:
            zout.writestr(i, conteudo[i.filename])
    os.replace(str(tmp), str(path))
    return len(plano), f"+{len(plano)} linha(s), até {plano[-1][0]}"


# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------
def main(ids: list[str] | None, apenas_check: bool) -> None:
    cdi_hist = ler_cdi_mercado(MERCADO_XLSX)
    datas = sorted(cdi_hist)
    print(f"CDI mercado: {len(datas)} datas ({datas[0]} -> {datas[-1]})\n")

    calcs = {r["IdSerie"]: r for r in ler_calculadoras_cdi(CADASTRO)}
    idx = indexar_homolog()
    alvos = ids if ids else list(calcs.keys())

    ok = em_dia = sem_cdi = erro = 0
    for sid in alvos:
        r = calcs.get(sid)
        if not r:
            print(f"  [{sid:>10}] NÃO NO CADASTRO"); erro += 1; continue
        cam = resolver_calc(r, idx)
        if cam is None or not Path(cam).exists():
            print(f"  [{sid:>10}] {r.get('Apelido', sid):<26} NÃO ENCONTRADO"); erro += 1; continue
        emissao = _normalizar_data(r["DataEmissao"])
        ape = r.get("Apelido", sid)
        try:
            if apenas_check:
                n = _contar(Path(cam), cdi_hist, emissao)
                tag = "sem aba CDI" if n < 0 else ("em dia" if n == 0 else f"faltam {n}")
                print(f"  [{sid:>10}] {ape:<26} {tag}")
            else:
                n, msg = _fix_um(Path(cam), cdi_hist, emissao)
                print(f"  [{sid:>10}] {ape:<26} {msg}")
                if n < 0:
                    sem_cdi += 1
                elif n == 0:
                    em_dia += 1
                else:
                    ok += 1
        except Exception as e:
            print(f"  [{sid:>10}] {ape:<26} ERRO: {e}"); erro += 1

    if not apenas_check:
        print(f"\nConcluído: {ok} atualizada(s) | {em_dia} em dia | "
              f"{sem_cdi} sem aba CDI | {erro} erro(s)")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    argv = sys.argv[1:]
    ids = None
    if "--ids" in argv:
        ids = [s for s in argv[argv.index("--ids") + 1].split(",") if s]
    main(ids, apenas_check="--check" in argv)
