"""Caminhos centrais. Sobrescrevíveis por variável de ambiente."""
import os
from pathlib import Path

PROJ = Path(__file__).resolve().parents[2]
DATA = PROJ / "data"

CALC_ROOT = Path(os.environ.get(
    "CALC_ROOT", r"C:\caminho\para\Calculadoras"))
FLUXO_DIR = CALC_ROOT / "Fluxo de pagamento"

# Fonte única de dados de mercado (gerada por 14-Base_Dados_Mercado)
BASE_MERCADO = Path(os.environ.get(
    "BASE_MERCADO",
    r"C:\caminho\para\Rotinas\Codigos"
    r"\14-Base_Dados_Mercado\output\mercado.xlsx",
))

CADASTRO = DATA / "cadastro_ativos.xlsx"  # xlsx é a fonte; csv mantido como backup
EVENTOS_DIR = DATA / "Eventos"

# código de série do CDI na aba "CDI" do mercado.xlsx
SERIE_DI = 2

# ---------------------------------------------------------------------------
# Fase de homologação
# ---------------------------------------------------------------------------
# Durante a homologação a rotina roda sobre as CÓPIAS locais em
# data/Calculadoras - Homologação (o SharePoint não é mais tocado) e grava
# saída/batimento em pastas próprias. Desligável com PU_HOMOLOG=0.
HOMOLOG = os.environ.get("PU_HOMOLOG", "1") == "1"
HOMOLOG_CALC_ROOT = Path(os.environ.get(
    "HOMOLOG_ROOT", str(DATA / "Calculadoras - Homologação")))

# Raízes de saída/batimento conforme a fase (mesma estrutura, sufixo distinto).
SAIDA_ROOT = DATA / ("Saída - Homologação" if HOMOLOG else "Saída")
BATIMENTO_ROOT = DATA / ("Batimento - Homologação" if HOMOLOG else "Batimento")

_HOMOLOG_IDX: dict | None = None


def indexar_homolog() -> dict:
    """Indexa os arquivos de homologação por nome de arquivo (minúsculo).

    Estrutura: HOMOLOG_CALC_ROOT/<familia>/<pasta - Homologação>/<arquivo>.
    Casar por nome de arquivo é robusto a divergências no nome da pasta.
    """
    idx: dict[str, Path] = {}
    if not HOMOLOG_CALC_ROOT.exists():
        return idx
    for sub in HOMOLOG_CALC_ROOT.iterdir():
        if not sub.is_dir():
            continue
        for pasta in sub.iterdir():
            if not pasta.is_dir():
                continue
            for arq in pasta.iterdir():
                if arq.suffix.lower() in (".xlsx", ".xlsm"):
                    idx[arq.name.lower()] = arq
    return idx


def homolog_idx() -> dict:
    """Índice de homologação com cache em memória."""
    global _HOMOLOG_IDX
    if _HOMOLOG_IDX is None:
        _HOMOLOG_IDX = indexar_homolog()
    return _HOMOLOG_IDX


def resolver_calc_path(calcpath: str) -> Path:
    """Resolve o caminho da calculadora conforme a fase.

    - Absoluto: usado como está.
    - Homolog: procura a cópia local pelo nome do arquivo; se não achar, cai
      para o SharePoint (FLUXO_DIR).
    - Produção: FLUXO_DIR / CalcPath.
    """
    p = Path(calcpath)
    if p.is_absolute():
        return p
    if HOMOLOG:
        homo = homolog_idx().get(Path(calcpath).name.lower())
        if homo is not None:
            return homo
    return FLUXO_DIR / calcpath.replace("/", "\\")
