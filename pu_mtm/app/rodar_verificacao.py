"""Calcula o PU Python do piloto e (CLI) compara com o oráculo Excel."""
from datetime import date
from pu_mtm.app import config
from pu_mtm.dados.cadastro import ler_cadastro, ler_eventos
from pu_mtm.dados.feriados import feriados_b3, dias_uteis_entre, contar_du
from pu_mtm.dados.indice_mercado import cdi_por_data
from pu_mtm.dominio.modelos import DiaCalc
from pu_mtm.dominio.familias.di import fator_juros_acumulado
from pu_mtm.dominio.familias import prefixado as fam_prefixado
from pu_mtm.dados.grade import montar_dias_prefixado
from pu_mtm.dominio.nucleo_pu import calcular_pu, arred_juros_por_nome

# caminho relativo da calculadora do piloto (no cadastro completo viria do campo CalcPath)
_CALC_REL = {
    "531": r"Contratos de Mútuo e Acordos\Acordo de Investimento\Ativo DI-piloto\calculadora_exemplo.xlsx",
}


def _dcp_30360(inicio: date, fim: date) -> int:
    """DAYS360 método europeu (TRUE): dia 31 é tratado como 30 em ambas as pontas.
    Espelha `=DAYS360(inicio, fim, TRUE)` das calculadoras (ex.: spread 30/360 do
    Ativo DI+Spread B/FOCUS). Difere do cálculo cru só quando uma data cai no dia 31."""
    d_ini = min(inicio.day, 30)
    d_fim = min(fim.day, 30)
    return (fim.year - inicio.year) * 360 + (fim.month - inicio.month) * 30 + (d_fim - d_ini)


def montar_dias(ativo, data_ref, feriados, cdi_map) -> list[DiaCalc]:
    """Monta a grade diária. Convenção DI-over (confirmada na Fase 0): a taxa que
    rende no dia `d` é a publicada no **dia útil anterior** — a calculadora faz
    `I_r = H_{r-1} * I_{r-1}`. Por isso cada passo usa o CDI do dia anterior."""
    dias_uteis = dias_uteis_entre(ativo.data_emissao, data_ref, feriados)  # (emissao, ref]
    dias_anteriores = [ativo.data_emissao] + dias_uteis[:-1]               # dia útil anterior de cada passo
    out = []
    for d, anterior in zip(dias_uteis, dias_anteriores):
        du = contar_du(ativo.data_emissao, d, feriados)
        dcp = _dcp_30360(ativo.data_emissao, d)
        cdi = cdi_map.get(anterior)
        if cdi is None:
            raise ValueError(f"CDI ausente para {anterior} (serie {config.SERIE_DI}, ativo {ativo.id_serie})")
        out.append(DiaCalc(data=d, du=du, dcp=dcp, dc=360, cdi=cdi))
    return out


def calcular_pu_piloto(id_serie: str, data_ref: date) -> float:
    ativos = ler_cadastro(str(config.CADASTRO))
    ativo = ativos[id_serie]
    eventos = ler_eventos(str(config.EVENTOS_DIR), id_serie, ativo.familia)
    if ativo.familia == "prefixado":
        # grade prefixada (30/360, ACT ou DU desde o cupom-âncora); não usa CDI/BaseMercado.
        # feriados só são usados no modo "du" (ex.: 750); baratos de calcular sempre.
        feriados = feriados_b3(ativo.data_emissao.year, data_ref.year)
        dias = montar_dias_prefixado(ativo, data_ref, eventos, feriados)
        fator_fn = fam_prefixado.fator_juros_acumulado
    else:
        # famílias DI: grade de dias úteis + curva CDI (DI-over)
        feriados = feriados_b3(ativo.data_emissao.year, data_ref.year)
        cdi_map = cdi_por_data(str(config.BASE_MERCADO), serie=config.SERIE_DI)
        dias = montar_dias(ativo, data_ref, feriados, cdi_map)
        fator_fn = fator_juros_acumulado
    r = calcular_pu(ativo, dias, eventos, fator_fn,
                    arred_juros_por_nome(ativo.juros_arred))
    return r.pu


if __name__ == "__main__":
    import sys
    from datetime import datetime
    ids = sys.argv[1] if len(sys.argv) > 1 else "531"
    dref = datetime.fromisoformat(sys.argv[2]).date() if len(sys.argv) > 2 else date.today()
    print(f"PU Python {ids} @ {dref} = {calcular_pu_piloto(ids, dref)}")
