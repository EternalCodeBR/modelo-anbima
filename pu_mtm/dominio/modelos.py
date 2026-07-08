"""Estruturas de dados puras do domínio (sem I/O)."""
from dataclasses import dataclass, field
from datetime import date

@dataclass(frozen=True)
class Ativo:
    id_serie: str
    apelido: str
    familia: str            # di_puro | di_spread | ipca_spread | prefixado
    vne: float              # valor nominal de emissão
    data_emissao: date
    spread: float = 0.0     # a.a., só di_spread (prêmio ADITIVO, fator anual à parte)
    percentual_cdi: float = 100.0  # % do CDI (ex.: 105 = RDB 464/474); MULTIPLICA o TDIk
    taxa_fixa: float = 0.0  # prefixado: taxa do período (ex.: 0.02 a.m.) ou a.a., conforme base
    # --- parâmetros de convenção por ativo (Fase 2; palpite em data/triagem/nuances.csv,
    #     confirmados pela prova do centavo) ---
    base: int = 252                       # 252 (dias úteis) | 360 (dias corridos)
    fator_diario_arred: str = "nenhum"    # nenhum | round8 | round9 | trunc16
    juros_arred: str = "trunc8"           # nenhum | trunc8 | round8
    evjuros_fonte: str = "nenhum"         # nenhum | acumulado | agendado
    evamort_fonte: str = "nenhum"         # nenhum | agendado
    # --- prefixado: convenção do fator (1+taxa)^(dcp/dc) ---
    pre_dc: int = 30                      # denominador: 30 (mensal) | 360 | 365 (anual)
    pre_daycount: str = "30360"           # contagem do numerador: 30360 | actual (ACT)

@dataclass(frozen=True)
class DiaCalc:
    """Uma linha da grade diária, já preparada pela camada de dados."""
    data: date
    du: int                 # dias úteis acumulados desde a emissão (base 252)
    dcp: int                # dias corridos no período (30/360)
    dc: int                 # base de dias corridos (360)
    cdi: float              # taxa CDI a.a. vigente na data (di_*); ignorada em prefixado
    expo: float | None = None  # prefixado per-período: expoente já pronto (meses+fração); fator=(1+taxa)^expo
    fator_ipca: float = 1.0    # ipca_spread: fator VNA acumulado no dia (Π mensal ROUND8), relativo à âncora

@dataclass(frozen=True)
class Evento:
    data: date                              # data efetiva (realizada se preenchida, senão prevista)
    evento_juros: float = 0.0              # valor efetivo (realizado se preenchido, senão previsto)
    evento_amortizacao: float = 0.0        # valor efetivo (realizado se preenchido, senão previsto)
    data_prevista: date | None = None      # data originalmente programada (None = sem desvio)
    tipo: str = "agendado"                 # agendado | extraordinario | parcial | antecipado | atrasado
    obs: str = ""

@dataclass
class ResultadoPU:
    id_serie: str
    data_ref: date
    pu: float
    serie_diaria: list = field(default_factory=list)  # [(data, pu), ...] p/ auditoria
