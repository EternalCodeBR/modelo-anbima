"""Reproduz o bloqueio do antivírus (Acronis) para demonstrar ao TI.

O antivírus mata o python.exe que SOBRESCREVE vários documentos Office em
sequência (assinatura de ransomware). Este script força esse cenário de forma
SEGURA e SEM EXCEL:

  1. copia N calculadoras para uma pasta temporária (copiar é permitido);
  2. num ÚNICO processo, abre e re-salva cada uma (openpyxl);
  3. o antivírus deve matar o processo por volta do 5º arquivo — é aí que a
     notificação do Acronis aparece na tela (tire o print para o TI).

Não altera as calculadoras reais. Não usa Excel.

Uso:
    python -m mtm_skills.reproduzir_bloqueio_ti
"""
import shutil
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")

_RAIZ = Path(__file__).resolve().parents[1]
if str(_RAIZ) not in sys.path:
    sys.path.insert(0, str(_RAIZ))

import openpyxl

from pu_mtm.app import config

N = 12  # arquivos a modificar (acima do limiar ~5 do antivírus)
# NÃO usar %TEMP% (AppData\Local\Temp): o Acronis normalmente exclui a pasta temp.
# Uma pasta comum do AppData\Local (fora de Temp) reproduz o bloqueio.
TMP = Path.home() / "AppData" / "Local" / "PU_MTM_repro_ti"


def main() -> None:
    arqs = sorted(
        p for p in config.HOMOLOG_CALC_ROOT.rglob("*")
        if p.suffix.lower() in (".xlsx", ".xlsm")
    )[:N]

    print("=" * 64)
    print("REPRODUCAO DO BLOQUEIO DO ANTIVIRUS (para o TI)")
    print("=" * 64)
    print("Processo unico modificando arquivos Office em sequencia.")
    print("Se o antivirus bloquear, este processo sera MORTO no meio")
    print("e a notificacao do Acronis aparecera na tela.\n")

    # 1) copiar para pasta temporaria (operacao permitida)
    if TMP.exists():
        shutil.rmtree(TMP, ignore_errors=True)
    TMP.mkdir(parents=True, exist_ok=True)
    copias = []
    for p in arqs:
        dst = TMP / p.name
        shutil.copy2(str(p), str(dst))
        copias.append(dst)
    print(f"[copia] {len(copias)} arquivos copiados para {TMP}\n")

    # 2) modificar em sequencia, num unico processo
    print("[modificacao] abrindo e re-salvando cada arquivo:")
    for i, dst in enumerate(copias, 1):
        wb = openpyxl.load_workbook(str(dst))
        wb.save(str(dst))          # SOBRESCREVE = dispara o antivirus
        wb.close()
        print(f"   {i:>2}/{len(copias)}  OK  {dst.name}", flush=True)

    # 3) se chegou aqui, o antivirus NAO bloqueou
    print("\nConcluido SEM bloqueio (o antivirus nao matou o processo).")
    shutil.rmtree(TMP, ignore_errors=True)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    main()
