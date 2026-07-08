"""Rotina diária de PU MtM.

Fluxo:
  1. Atualiza o CDI nas calculadoras — CIRÚRGICO (sem Excel, sem openpyxl):
     anexa só as datas que faltam no XML da aba CDI → mtm_skills.atualizar_cdi_surgical
  2. Calcula os PUs pelo motor Python                    → data/Saída/...
  3. Batimento motor × calculadora                       → data/Batimento/...

A leitura dos dados de mercado (CDI/IPCA) via API alimenta o mercado.xlsx, que é
a fonte lida no passo 1.

IMPORTANTE: NUNCA gravar as calculadoras com openpyxl (corrompe VML/desenhos/
comentários) nem via Excel/pywin32 (lento e sujeito ao antivírus). O método
oficial é o cirúrgico no ZIP — ver docstring de atualizar_cdi_surgical.

Uso:
    python rodar_diario.py
"""
import sys
from pathlib import Path

_RAIZ = Path(__file__).resolve().parent
if str(_RAIZ) not in sys.path:
    sys.path.insert(0, str(_RAIZ))

from mtm_skills.atualizar_cdi_surgical import main as _atualizar_cdi
from mtm_skills.gerar_planilha_pu import gerar_planilha
from pu_mtm.app.bater import rodar_todos

_SEP = "=" * 60


def main() -> tuple[str, str | None]:
    print(_SEP)
    print("1/3  Atualizando CDI nas calculadoras (cirúrgico, sem Excel)")
    print(_SEP)
    _atualizar_cdi(ids=None, apenas_check=False)

    print(f"\n{_SEP}")
    print("2/3  Calculando PUs (motor)")
    print(_SEP)
    arquivo_pu = gerar_planilha()

    print(f"\n{_SEP}")
    print("3/3  Batimento motor x calculadora")
    print(_SEP)
    arquivo_bat = rodar_todos()

    print(f"\n{_SEP}")
    print("CONCLUIDO")
    print(f"  PUs:       {arquivo_pu}")
    print(f"  Batimento: {arquivo_bat or '(nenhum resultado gravado)'}")
    print(_SEP)
    return arquivo_pu, str(arquivo_bat) if arquivo_bat else None


if __name__ == "__main__":
    main()
