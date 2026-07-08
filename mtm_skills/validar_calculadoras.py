"""Validador consolidado das calculadoras (relatório único verde/vermelho).

Roda, por ativo no escopo (di_puro + di_spread), três camadas de verificação:

  ESTRUTURA  — coluna-inteira em fórmula, forceFullCalc, vínculo externo,
               serial de data com hora (#N/D), <v> vazio (corrupção openpyxl).
  CDI        — última data de CDI na planilha vs. no mercado (defasagem = ponta).
  PREÇO      — motor × cache, só em DIA ÚTIL e só em datas COM CDI disponível
               (isola erro de precificação de mera defasagem de dado).

Determinístico: não abre Excel, não altera nada. É o insumo para decidir onde
vale despachar agente de investigação (só os vermelhos de PREÇO).

Uso:
    python -m mtm_skills.validar_calculadoras
    python -m mtm_skills.validar_calculadoras --ids 629,691
"""
from __future__ import annotations

import re
import sys
import warnings
import zipfile
from datetime import date
from pathlib import Path

warnings.filterwarnings("ignore")

_RAIZ = Path(__file__).resolve().parents[1]
if str(_RAIZ) not in sys.path:
    sys.path.insert(0, str(_RAIZ))

from pu_mtm.app import bater, config
from pu_mtm.app.rodar_verificacao import calcular_pu_piloto
from pu_mtm.dados import csvio
from pu_mtm.dados.cadastro import ler_cadastro
from pu_mtm.dados.indice_mercado import cdi_por_data
from pu_mtm.verificacao.comparador import comparar
from mtm_skills.atualizar_cdi_surgical import _cdi_part, _ultima_data

_RE_WS = re.compile(r"xl/worksheets/sheet\d+\.xml$")
_RE_FULLCOL = re.compile(r"[A-Za-z0-9_]+!\$?[A-Z]{1,3}:\$?[A-Z]{1,3}")
_RE_A_FRAC = re.compile(r'<c r="A\d+"[^>]*><v>\d+\.\d+</v>')


def _estrutural(cont: dict[str, bytes]) -> dict:
    fullcol = frac = vazio = 0
    for nome, raw in cont.items():
        if _RE_WS.match(nome):
            xml = raw.decode("utf-8", "replace")
            fullcol += len(_RE_FULLCOL.findall(xml))
            vazio += xml.count("<v></v>")
    part = _cdi_part(cont)
    if part and part in cont:
        frac = len(_RE_A_FRAC.findall(cont[part].decode("utf-8", "replace")))
    ffc = 'forceFullCalc="1"' in cont.get("xl/workbook.xml", b"").decode("utf-8", "replace")
    extlinks = sum(1 for n in cont if re.match(r"xl/externalLinks/externalLink\d+\.xml$", n))
    ok = (fullcol == 0 and not ffc and extlinks == 0 and frac == 0 and vazio == 0)
    return {"ok": ok, "fullcol": fullcol, "ffc": ffc, "ext": extlinks,
            "frac": frac, "vazio": vazio}


def _cdi_ultima(cont: dict[str, bytes]) -> str | None:
    part = _cdi_part(cont)
    if not part or part not in cont:
        return None
    iso, _ = _ultima_data(cont[part].decode("utf-8", "replace"))
    return iso


def _preco(id_serie: str, meta: dict, ativo, cdi_calc_iso: str | None,
           mercado_teto: date) -> dict:
    """Bate motor × cache em dia útil, só em datas com CDI disponível na planilha."""
    pu_excel = bater._pu_calc(meta)  # cache do Excel
    teto = mercado_teto
    if cdi_calc_iso:
        teto = min(teto, date.fromisoformat(cdi_calc_iso))
    datas = bater._datas(pu_excel, teto)
    ok = xx = 0
    piores = []
    for d in datas:
        if d not in pu_excel or not bater._eh_dia_util(d):
            continue
        pu = calcular_pu_piloto(id_serie, d)
        rel = comparar(py=pu, excel=pu_excel[d], id_serie=id_serie, vne=ativo.vne)
        if rel["ok"]:
            ok += 1
        else:
            xx += 1
            piores.append((d.isoformat(), rel["py"] - rel["excel"]))
    return {"ok": xx == 0 and ok > 0, "n_ok": ok, "n_xx": xx, "piores": piores}


def _linha(id_serie: str, mercado_teto: date) -> dict:
    meta = bater._meta(id_serie)
    ativo = ler_cadastro(str(config.CADASTRO))[id_serie]
    cam = bater._calc_path(meta)
    if not cam.exists():
        return {"id": id_serie, "ape": ativo.apelido, "fam": ativo.familia,
                "estrut": None, "cdi": None, "preco": None, "erro": "não encontrada"}
    try:
        with zipfile.ZipFile(str(cam)) as z:
            cont = {n: z.read(n) for n in z.namelist()}
    except PermissionError:
        return {"id": id_serie, "ape": ativo.apelido, "fam": ativo.familia,
                "estrut": None, "cdi": None, "preco": None, "erro": "aberta no Excel"}

    est = _estrutural(cont)
    cdi_iso = _cdi_ultima(cont)
    defasagem = None
    if cdi_iso:
        defasagem = (mercado_teto - date.fromisoformat(cdi_iso)).days
    try:
        prc = _preco(id_serie, meta, ativo, cdi_iso, mercado_teto)
        erro = None
    except Exception as e:
        prc = None
        erro = f"preço: {e}"
    return {"id": id_serie, "ape": ativo.apelido, "fam": ativo.familia,
            "estrut": est, "cdi": cdi_iso, "defas": defasagem, "preco": prc, "erro": erro}


def _tag(b: bool | None) -> str:
    return "OK " if b else ("XX " if b is False else "?? ")


def main(ids: list[str] | None) -> None:
    mercado_teto = max(cdi_por_data(str(config.BASE_MERCADO), serie=config.SERIE_DI))
    print(f"Mercado: CDI até {mercado_teto.isoformat()}\n")

    linhas = list(csvio.ler_dict(str(config.CADASTRO)))
    escopo = bater.FAMILIAS_ESCOPO
    alvos = ids or [
        r["IdSerie"] for r in linhas
        if str(r.get("Familia", "")).strip().lower() in escopo
        and "liquidad" not in str(r.get("Status", "")).lower()
        and "congelad" not in str(r.get("Status", "")).lower()
        and r.get("CalcPath")
    ]

    print(f"  {'Id':>6}  {'Apelido':<28}{'Estrut':>7}{'CDI':>7}{'Preço':>8}   Observação")
    print("  " + "-" * 82)
    reds = []
    for sid in alvos:
        L = _linha(sid, mercado_teto)
        if L["erro"] and L["estrut"] is None:
            print(f"  {sid:>6}  {L['ape']:<28}{'??':>7}{'??':>7}{'??':>8}   {L['erro']}")
            reds.append((sid, L["ape"], L["erro"]))
            continue
        est_ok = L["estrut"]["ok"]
        cdi_ok = (L["defas"] is not None and L["defas"] <= 0)
        prc = L["preco"]
        prc_ok = prc["ok"] if prc else None
        obs = []
        if not est_ok:
            e = L["estrut"]
            det = [k for k, v in (("colInt", e["fullcol"]), ("FFC", e["ffc"]),
                    ("ext", e["ext"]), ("hora", e["frac"]), ("vVazio", e["vazio"])) if v]
            obs.append("estrut:" + ",".join(det))
        if not cdi_ok and L["defas"] is not None:
            obs.append(f"CDI D-{L['defas']}")
        if prc and not prc["ok"]:
            obs.append(f"preço {prc['n_xx']}XX " + ",".join(
                f"{d}({dif:+.6f})" for d, dif in prc["piores"][:2]))
        if L["erro"]:
            obs.append(L["erro"])
        print(f"  {sid:>6}  {L['ape']:<28}{_tag(est_ok):>7}{_tag(cdi_ok):>7}"
              f"{_tag(prc_ok):>8}   {' | '.join(obs)}")
        if (prc_ok is False) or (prc_ok is None):
            reds.append((sid, L["ape"], " | ".join(obs)))

    print("  " + "-" * 82)
    print(f"\nTotal: {len(alvos)} | vermelhos de PREÇO (candidatos a agente): {len(reds)}")
    for sid, ape, obs in reds:
        print(f"  - {sid} {ape}: {obs}")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    argv = sys.argv[1:]
    ids = None
    if "--ids" in argv:
        ids = [s for s in argv[argv.index("--ids") + 1].split(",") if s]
    main(ids)
