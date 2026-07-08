r"""Rotina diária UNIFICADA de PU MtM — orquestra tudo, incluindo o recálculo VBA.

Ordem (a ordem importa: o cache tem de ser recalculado ANTES do batimento):
  1. Atualiza o CDI nas 23 calculadoras (cirúrgico, Python)  -> atualizar_cdi_surgical
  2. Recalcula o cache de PU no Excel (VBA RecalcCacheLocal)  -> via COM (excel.Run)
  3. Calcula os PUs pelo motor Python                         -> gerar_planilha
  4. Batimento motor x cache (já recalculado)                 -> rodar_todos

Passo 2 (COM): abre uma instância dedicada do Excel (DispatchEx), roda a macro
RecalcCacheLocal em modo silencioso e fecha. excel.Run é SÍNCRONO — o Python só
segue quando a macro termina. AutomationSecurity=1 habilita a macro sem prompt (não
precisa de Local Confiável). A macro em si recalcula em cópia local (C:\PU_Recalc),
fora do OneDrive, por isso é rápida (~21s nas 23).

PRÉ-REQUISITOS (uma vez):
  * pywin32 instalado:  pip install pywin32
  * RecalcCacheLocal (vba/RecalcCacheLocal.bas) colada num módulo do Rotina_Recalc.xlsm
  (o Workbook_Open de vba/ThisWorkbook_Workbook_Open.bas é OPCIONAL — só serve para o
   modo duplo-clique; no fluxo COM ele não é usado.)

Uso:
    python rodar_rotina.py            # rotina completa (dispara o VBA via COM)
    python rodar_rotina.py --manual   # pula o COM; pausa p/ você rodar o VBA à mão
"""
import os
import sys
import time
from pathlib import Path

_RAIZ = Path(__file__).resolve().parent
if str(_RAIZ) not in sys.path:
    sys.path.insert(0, str(_RAIZ))

from mtm_skills.atualizar_cdi_surgical import main as _atualizar_cdi
from mtm_skills.gerar_planilha_pu import gerar_planilha
from pu_mtm.app.bater import rodar_todos

_SEP = "=" * 64
# O .xlsm precisa morar em LOCAL CONFIÁVEL para a GPO corporativa liberar a macro no
# COM (a pasta pessoal do OneDrive é bloqueada). base é confiável (e é
# a zona excluída do Acronis). Ver memória antivirus-acronis-modify-lote.
ROTINA_XLSM = Path(r"C:\caminho\para\Rotina_Recalc.xlsm")
MACRO = "RecalcCacheLocalAuto"   # ponto de entrada SEM parâmetro (silencioso, grava log)
_LOG = os.path.join(os.environ.get("TEMP", "."), "recalc_log.txt")


def _com_retry(func, *args, tentativas: int = 5, espera: float = 1.0):
    """Reexecuta uma chamada COM que o Excel rejeita quando está ocupado
    (RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER). Espelha o padrão do
    Importacao_PU's.py."""
    ultimo = None
    for _ in range(tentativas):
        try:
            return func(*args)
        except Exception as e:  # pywin32 levanta pythoncom.com_error (subclasse de Exception)
            ultimo = e
            time.sleep(espera)
    raise ultimo


def _recalcular_via_excel() -> bool:
    """Abre o Excel via COM e dispara o RecalcCacheLocal (síncrono). True se ok."""
    if not ROTINA_XLSM.exists():
        print(f"  !! não encontrei {ROTINA_XLSM}")
        return False
    try:
        import win32com.client as win32
        import pythoncom
    except ImportError:
        print("  pywin32 não instalado (pip install pywin32). Use --manual.")
        return False

    if os.path.exists(_LOG):        # apaga log anterior p/ distinguir execução nova
        try:
            os.remove(_LOG)
        except Exception:
            pass

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")     # instância dedicada
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            excel.AutomationSecurity = 1                  # habilita macro sem prompt
        except Exception:
            pass
        try:
            excel.AskToUpdateLinks = False
        except Exception:
            pass

        print(f"  Abrindo {ROTINA_XLSM.name} e disparando {MACRO} (síncrono)...")
        wb = _com_retry(excel.Workbooks.Open, str(ROTINA_XLSM), 0, False)  # UpdateLinks=0, ReadOnly=False
        _com_retry(excel.Run, f"'{ROTINA_XLSM.name}'!{MACRO}")             # sem argumento

        # a macro grava o resultado em %TEMP%\recalc_log.txt — fonte autoritativa
        if os.path.exists(_LOG):
            with open(_LOG, "r", encoding="cp1252", errors="replace") as f:
                print("  " + f.read().strip())
        else:
            print("  cache recalculado (sem log).")
        return True
    except Exception as e:
        print(f"  falha ao disparar via COM: {e}")
        return False
    finally:
        try:
            if wb is not None:
                _com_retry(wb.Close, False, tentativas=3)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def _pausa_manual() -> None:
    input("  >> Rode o RecalcCacheLocal no Excel e tecle ENTER para continuar... ")


def main(manual: bool = False) -> None:
    print(_SEP); print("1/4  Atualizando CDI nas calculadoras (cirurgico, sem Excel)"); print(_SEP)
    _atualizar_cdi(ids=None, apenas_check=False)

    print(f"\n{_SEP}"); print("2/4  Recalculando o cache de PU no Excel"); print(_SEP)
    if manual:
        _pausa_manual()
    elif not _recalcular_via_excel():
        print("  (fallback) disparo automatico falhou.")
        _pausa_manual()

    print(f"\n{_SEP}"); print("3/4  Calculando PUs (motor Python)"); print(_SEP)
    arquivo_pu = gerar_planilha()

    print(f"\n{_SEP}"); print("4/4  Batimento motor x cache"); print(_SEP)
    arquivo_bat = rodar_todos()

    print(f"\n{_SEP}"); print("CONCLUIDO")
    print(f"  PUs:       {arquivo_pu}")
    print(f"  Batimento: {arquivo_bat or '(nenhum resultado gravado)'}")
    print(_SEP)


if __name__ == "__main__":
    main(manual="--manual" in sys.argv[1:])
