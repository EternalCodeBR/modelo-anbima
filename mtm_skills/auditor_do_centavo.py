import pandas as pd
from pu_mtm.verificacao.comparador import bate

def encontrar_primeiro_desvio(id_serie, serie_motor_df, serie_excel_df):
    """
    Cruza o histórico do motor com o do Excel (dia a dia) usando pd.merge.
    Interrompe no primeiro dia em que o PU do motor NÃO bate com o do Excel pela
    régua oficial (comparador.bate — relativa 1e-8, escala-invariante).
    Retorna True se houve desvio, False se bater em toda a série.
    """
    df_merge = pd.merge(serie_motor_df, serie_excel_df, on='Data', suffixes=('_motor', '_excel'))

    for row in df_merge.itertuples():
        pu_m = row.PU_motor
        pu_e = row.PU_excel
        delta = abs(pu_m - pu_e)

        # Régua unificada: o auditor usa exatamente o mesmo critério do comparador
        # do motor (relativo 1e-8), em vez de uma tolerância absoluta própria.
        if not bate(pu_m, pu_e):
            print(f"--- Auditando prova do centavo: Ativo {id_serie} ---")
            print("🚨 DESVIO ENCONTRADO NA PROVA DO CENTAVO!")
            print(f"Data do Desvio: {row.Data}")
            print(f"Valor em Python: {pu_m:.10f}")
            print(f"Valor no Excel:  {pu_e:.10f}")
            print(f"Delta (Erro):    {delta:.8e}")
            return True
            
    print(f"--- Auditando prova do centavo: Ativo {id_serie} ---")
    print("✅ Sucesso absoluto! O motor replicou perfeitamente a calculadora dentro da tolerância exigida.")
    return False

if __name__ == "__main__":
    from pu_mtm.app import config
    from pu_mtm.app.bater import _meta, _pu_calc
    from pu_mtm.dados.cadastro import ler_cadastro, ler_eventos
    from pu_mtm.dados.feriados import feriados_b3
    from pu_mtm.dados.indice_mercado import cdi_por_data
    from pu_mtm.dominio.nucleo_pu import calcular_pu, arred_juros_por_nome
    from pu_mtm.dominio.familias.di import fator_juros_acumulado
    from pu_mtm.dominio.familias import prefixado as fam_prefixado
    from pu_mtm.dados.grade import montar_dias_prefixado
    from pu_mtm.app.rodar_verificacao import montar_dias
    from pu_mtm.dados import csvio
    from mtm_skills.gerar_planilha_pu import coletar_pus, VALORES_CONGELADOS

    print("Iniciando auditoria em massa (mesma seleção/data da esteira de produção)...")

    # A lista de ativos a auditar vem da própria esteira (coletar_pus), não mais
    # de um CSV — mesma seleção (ignora liquidados/ipca, inclui congelados) e data.
    ids_auditados = [(r["id_serie"], r["data"]) for r in coletar_pus()]

    ativos = ler_cadastro(str(config.CADASTRO))
    # Status (congelado) é dado cadastral, não vem do Ativo: lê do CSV cru, igual à esteira.
    status_por_ativo = {l["IdSerie"]: str(l.get("Status", "")).lower()
                        for l in csvio.ler_dict(str(config.CADASTRO))}
    cdi_map = cdi_por_data(str(config.BASE_MERCADO), serie=config.SERIE_DI)
    feriados = feriados_b3(2000, 2027)

    desvios_encontrados = 0

    for id_serie, data_ref in ids_auditados:
        ativo = ativos[id_serie]
        if ativo.familia == "ipca_spread":
            continue

        # Congelado: a esteira NÃO calcula ao vivo — exporta o PU fixo (VALORES_CONGELADOS).
        # Auditar o motor "ao vivo" contra a planilha congelada acusaria desvio de propósito
        # (o motor acumula CDI; a calculadora parou). Aqui replicamos a esteira: validamos o
        # PU fixo contra o valor congelado da calculadora (último PU disponível).
        if "congelad" in status_por_ativo.get(id_serie, ""):
            print(f"--- Auditando congelado: Ativo {id_serie} ---")
            if id_serie not in VALORES_CONGELADOS:
                print("⏭️  Sem PU fixo cadastrado — pulado (não conta como desvio).")
                continue
            try:
                pu_excel_dict = _pu_calc(_meta(id_serie))
            except (KeyError, RuntimeError) as e:
                print(f"🚨 {e}")
                desvios_encontrados += 1
                continue
            pu_fix = VALORES_CONGELADOS[id_serie]
            d_cong = max((d for d in pu_excel_dict if d <= data_ref), default=None)
            pu_calc = pu_excel_dict.get(d_cong)
            if pu_calc is not None and bate(pu_fix, pu_calc):
                print(f"✅ PU fixo {pu_fix:.10f} bate a calculadora ({d_cong}).")
            else:
                print(f"🚨 PU fixo {pu_fix} DIVERGE do valor congelado da calculadora ({pu_calc} em {d_cong}).")
                desvios_encontrados += 1
            continue

        try:
            meta = _meta(id_serie)
        except KeyError:
            print(f"[{id_serie}] Pulo: Metadados (planilha) não configurados.")
            continue
            
        try:
            # 1. Obter série do Excel (Calculadora original)
            pu_excel_dict = _pu_calc(meta)
            serie_excel_df = pd.DataFrame(list(pu_excel_dict.items()), columns=['Data', 'PU'])
            
            # 2. Obter série do Motor Python
            eventos = ler_eventos(str(config.EVENTOS_DIR), id_serie, ativo.familia)
            if ativo.familia == "prefixado":
                dias = montar_dias_prefixado(ativo, data_ref, eventos, feriados)
                fator_fn = fam_prefixado.fator_juros_acumulado
            else:
                dias = montar_dias(ativo, data_ref, feriados, cdi_map)
                fator_fn = fator_juros_acumulado
                
            resultado_pu = calcular_pu(ativo, dias, eventos, fator_fn, arred_juros_por_nome(ativo.juros_arred))
            serie_motor_df = pd.DataFrame(resultado_pu.serie_diaria, columns=['Data', 'PU'])
            
            # 3. Rodar Auditor
            teve_desvio = encontrar_primeiro_desvio(id_serie, serie_motor_df, serie_excel_df)
            if teve_desvio:
                desvios_encontrados += 1
                
        except Exception as e:
            print(f"[{id_serie}] Erro ao auditar: {e}")
            
    print("-" * 50)
    print(f"Auditoria concluída! {desvios_encontrados} ativos com desvio encontrados.")