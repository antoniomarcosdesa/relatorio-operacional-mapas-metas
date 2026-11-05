# ============================================================
# PROCESSAR_MAPA_ANALITICO.PY
#
# Novo pipeline de dados para os clientes que enviam o Excel/CSV
# manualmente (mapa analítico detalhado — ~10 linhas por minuto de
# operação), em vez de coletar via SGPA/PBI como no projeto original.
#
# Mesma filosofia do processar_dados.py original:
#   - área operacional só entra no relatório se > área mínima (aqui,
#     por TALHÃO dentro de cada processo, conforme pedido: > 5 ha);
#   - indicadores de qualidade operacional são médias ponderadas
#     pelo tempo de cada leitura, calculadas só no estado EFETIVO
#     (cd_estado == "E"), que é o equivalente a "TRABALHANDO" no
#     projeto antigo;
#   - quando uma coluna de qualidade não tem NENHUM dado real para
#     aquele Processo x Talhão (sensor não existe/não se aplica —
#     ex.: vazão de bico numa operação de plantio), o indicador fica
#     em branco ("-"), em vez de mostrar 0.00 (que sugeriria "mediu
#     zero", não "não se aplica").
#
# ENTRADA:
#   Um CSV/XLSX no formato "mapa analítico" (ponto-e-vírgula, ~10
#   linhas por minuto), com as colunas:
#     cd_fazenda, cd_zona, cd_talhao, dt_hr_local_inicial,
#     cd_operacao, cd_equipamento, cd_estado (E/M/F/D/A),
#     vl_largura_implemento, vl_tempo_segundos, cd_operacao_parada,
#     desc_parada, cd_operador, desc_operador, "Ordem de serviço",
#     wkt, + as colunas de qualidade operacional (vl_velocidade,
#     vl_rpm, vl_temperatura_motor, vl_consumo_instantaneo,
#     vl_rendimento_colheita, vl_umidade_graos, vl_vazao_litros_ha,
#     vl_vazao_litros_min, "vl_pressao_bomba (kPa)",
#     vl_qtde_sementes_ha_total, vl_hectares_hora,
#     vl_velocidade_vento, vl_temperatura, vl_umidade,
#     vl_ponto_orvalho, vl_taxa_aplicacao_solido_ha, vl_torque).
#
# SAÍDA:
#   resultados/qualidade_operacional_processo_talhao.xlsx
#     - Resumo_Processo_Talhao : 1 linha por Processo x Talhão
#     - Resumo_Processo        : 1 linha por Processo (todos os talhões)
#     - Estados                : horas por estado (E/M/F/D/A) por processo/talhão
#     - Validacao               : rastreabilidade (linhas lidas, filtros, período)
#
# CÓDIGOS DE OPERAÇÃO (cd_operacao):
#   O arquivo bruto não traz descrição do processo, só o código.
#   Preencha CATALOGO_OPERACOES abaixo (código -> nome) assim que
#   tiver a tabela de-para do cliente; enquanto isso, o relatório usa
#   "Processo <código>" como rótulo, então nada quebra por falta do
#   de-para.
# ============================================================

from pathlib import Path
import numpy as np
import pandas as pd

from leitura_csv import ler_csv_bruto, parse_datas

# ------------------------------------------------------------
# CONFIGURAÇÃO
# ------------------------------------------------------------

AREA_MINIMA_TALHAO = 5.0  # ha — regra pedida: só entra talhão com > 5 ha no período

# Limiares de alerta de temperatura do motor (°C) — contados por processo
# (somando todos os talhões que entram no relatório daquele processo),
# só em estado Efetivo. Aparecem como texto informativo na página do PDF.
LIMIARES_TEMP_MOTOR = [100, 105]

ESTADO_EFETIVO = "E"          # equivalente a "TRABALHANDO" no projeto antigo
ESTADOS_LABEL = {
    "E": "Efetivo (trabalhando)",
    "M": "Manobra",
    "F": "Parada",
    "D": "Deslocamento",
    "A": "Deslocamento p/ reabastecimento",
}

# De-para código -> nome do processo. Preencher conforme o cliente
# fornecer a tabela oficial (ex.: 1403: "GRADAGEM", 1401: "PLANTIO"...).
CATALOGO_OPERACOES = {
    # 1403: "GRADAGEM NIVELADORA",
    # 1401: "PLANTIO",
}

# Colunas de QUALIDADE OPERACIONAL DO EQUIPAMENTO (média ponderada por
# tempo, só em estado Efetivo). Ficam em branco quando o sensor não se
# aplica àquele processo (coluna inteiramente zerada/nula no grupo).
# Tupla: (nome exibido, unidade, casas decimais, permite valor negativo)
#
# permite_negativo=False (padrão do equipamento): nenhuma dessas grandezas
# é fisicamente negativa, então valor < 0 é tratado como sentinela de
# "sem leitura" (ex.: vl_taxa_aplicacao_solido_ha vem com -1/-0.9/-0.81
# misturado com leituras reais — confirmado nos dados do Ourives) e é
# DESCARTADO da média, não contado como zero.
# "vl_consumo_instantaneo" e "vl_hectares_hora" NÃO entram aqui — os dois
# ganham cálculo próprio (não é média ponderada simples):
#   - Consumo: total de litros consumidos (estado Efetivo) / área operacional
#     do talhão -> "Consumo (L/ha)", calculado em calcular_consumo_l_ha().
#   - "Rendimento operacional (ha/h)" foi substituído por "Rendimento Real (%)"
#     = Horas efetivas / Horas totais (motor ligado), que já existia como
#     "Eficiência (efetivo/total)" — mesmo cálculo, renomeado.
# vl_rendimento_colheita vem da fonte em toneladas/ha; o relatório usa sacas/ha (1 saca = 60 kg).
FATOR_COLHEITA_T_HA_PARA_SC_HA = 1000.0 / 60.0

COLUNAS_QUALIDADE_EQUIP = {
    "vl_velocidade": ("Velocidade média", "km/h", 2, False),
    "vl_rpm": ("RPM médio", "RPM", 0, False),
    "vl_temperatura_motor": ("Temperatura do motor média", "°C", 1, False),
    "vl_torque": ("Torque médio", "%", 1, False),
    "vl_rendimento_colheita": ("Rendimento de colheita", "sc/ha", 1, False),
    "vl_umidade_graos": ("Umidade dos grãos", "%", 1, False),
    "vl_vazao_litros_ha": ("Vazão de calda", "L/ha", 1, False),
    "vl_vazao_litros_min": ("Vazão de calda (L/min)", "L/min", 2, False),
    "vl_pressao_bomba (kPa)": ("Pressão da bomba", "kPa", 1, False),
    "vl_qtde_sementes_ha_total": ("Densidade de plantio", "sementes/ha", 0, False),
    "vl_taxa_aplicacao_solido_ha": ("Taxa de aplicação (sólido)", "kg/ha", 1, False),
}

# Colunas de CONDIÇÕES AMBIENTAIS (contexto, não é qualidade do
# equipamento em si — mas ajuda a explicar desvios). Mesma regra de
# ponderação/branco, mas aqui negativo É um valor válido (temperatura e
# ponto de orvalho podem legitimamente ficar abaixo de zero).
COLUNAS_AMBIENTE = {
    "vl_temperatura": ("Temperatura ambiente média", "°C", 1, True),
    "vl_umidade": ("Umidade relativa média", "%", 1, True),
    "vl_ponto_orvalho": ("Ponto de orvalho médio", "°C", 1, True),
    "vl_velocidade_vento": ("Velocidade do vento média", "km/h", 1, False),
}

TODAS_COLUNAS_MEDIA = {**COLUNAS_QUALIDADE_EQUIP, **COLUNAS_AMBIENTE}


# ------------------------------------------------------------
# LEITURA
# ------------------------------------------------------------

def carregar_bruto(caminho: Path) -> pd.DataFrame:
    caminho = Path(caminho)
    if caminho.suffix.lower() == ".csv":
        df = ler_csv_bruto(caminho, incluir_wkt=False)
    else:
        df = pd.read_excel(caminho)

    df.columns = [str(c).strip() for c in df.columns]

    obrigatorias = [
        "cd_talhao", "dt_hr_local_inicial", "cd_operacao", "cd_equipamento",
        "cd_estado", "vl_tempo_segundos", "vl_hectares_hora",
    ]
    faltantes = [c for c in obrigatorias if c not in df.columns]
    if faltantes:
        raise ValueError(
            "Colunas obrigatórias ausentes no mapa analítico:\n"
            + "\n".join(f"- {c}" for c in faltantes)
        )

    df["dt_hr_local_inicial"] = parse_datas(df["dt_hr_local_inicial"])

    # Talhão/fazenda/zona não-definidos ficam de fora do relatório
    # (mesma lógica do multicliente: nunca "adivinha" unidade/talhão).
    df["cd_talhao"] = df["cd_talhao"].astype(str).str.strip()
    df = df[
        df["cd_talhao"].notna()
        & (df["cd_talhao"] != "")
        & (df["cd_talhao"].str.upper() != "*NAO_DEFINIDO*")
    ].copy()

    for c in TODAS_COLUNAS_MEDIA:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    if "vl_rendimento_colheita" in df.columns:  # t/ha -> sacas/ha (negativos/zero seguem como "sem dado")
        df["vl_rendimento_colheita"] = df["vl_rendimento_colheita"] * FATOR_COLHEITA_T_HA_PARA_SC_HA

    df["vl_tempo_segundos"] = pd.to_numeric(df["vl_tempo_segundos"], errors="coerce").fillna(0)
    df["vl_hectares_hora"] = pd.to_numeric(df["vl_hectares_hora"], errors="coerce").fillna(0)

    df["cd_estado"] = df["cd_estado"].astype(str).str.strip().str.upper()
    df["eh_efetivo"] = df["cd_estado"] == ESTADO_EFETIVO

    # Área da leitura: taxa instantânea (ha/h) x duração da leitura (h).
    df["area_ha"] = np.where(
        df["eh_efetivo"], df["vl_hectares_hora"] * df["vl_tempo_segundos"] / 3600.0, 0.0
    )

    print(f"Arquivo: {caminho.name}")
    print(f"Linhas lidas: {len(df):,}".replace(",", "."))
    print(f"Período: {df['dt_hr_local_inicial'].min()} a {df['dt_hr_local_inicial'].max()}")
    print(f"Processos (códigos) encontrados: {sorted(df['cd_operacao'].unique())}")
    print(f"Talhões encontrados: {sorted(df['cd_talhao'].unique())}")

    return df


def rotulo_processo(codigo) -> str:
    return CATALOGO_OPERACOES.get(codigo, f"Processo {codigo}")


# ------------------------------------------------------------
# MÉDIA PONDERADA COM REGRA DE "SEM DADO"
# ------------------------------------------------------------

def media_ponderada_ou_branco(valores: pd.Series, pesos: pd.Series, permite_negativo: bool = False,
                              zero_sem_dado: bool = False):
    """Média ponderada pelo tempo. Se a coluna não tiver NENHUM dado
    real no grupo (tudo NaN ou tudo exatamente zero), devolve NaN —
    que o relatório final imprime como "-", em vez de 0.00 (que
    sugeriria uma medição real de zero).

    Quando permite_negativo=False (padrão para as colunas do
    equipamento, que não têm grandeza física negativa), valores < 0
    são tratados como sentinela de "sem leitura" e excluídos da
    média — não contam nem como zero."""
    v = pd.to_numeric(valores, errors="coerce")
    p = pd.to_numeric(pesos, errors="coerce")

    if not permite_negativo:
        v = v.where(v >= 0)
    if zero_sem_dado:
        # Indicadores do equipamento: 0 (ou -1) = "esse equipamento não tem o
        # dado" -> fica fora da média, em vez de puxá-la para baixo.
        v = v.where(v > 0)

    tem_dado_real = v.notna().any() and (v.fillna(0) != 0).any()
    if not tem_dado_real:
        return np.nan

    mask = v.notna() & p.notna() & (p > 0)
    if not mask.any():
        return np.nan

    return float((v[mask] * p[mask]).sum() / p[mask].sum())


# Indicadores "por ha": a média geral é ponderada pela ÁREA trabalhada em cada
# leitura (média por hectare), só entre equipamentos que têm o dado. Os demais
# (velocidade, RPM, torque, pressão...) seguem ponderados pelo tempo.
PESO_POR_AREA = {
    "vl_taxa_aplicacao_solido_ha", "vl_vazao_litros_ha",
    "vl_qtde_sementes_ha_total", "vl_rendimento_colheita",
}

# Indicadores para os quais o PDF mostra o texto de "variação em torno da média".
INDICADORES_VARIACAO = {
    "vl_vazao_litros_ha": ("Vazão de calda", "L/ha", 1),
    "vl_pressao_bomba (kPa)": ("Pressão da bomba", "kPa", 1),
    "vl_qtde_sementes_ha_total": ("Densidade de plantio", "sem./ha", 0),
    "vl_taxa_aplicacao_solido_ha": ("Taxa de aplicação", "kg/ha", 1),
    "vl_rendimento_colheita": ("Rendimento de colheita", "sc/ha", 1),
    "vl_umidade_graos": ("Umidade dos grãos", "%", 1),
    "vl_torque": ("Torque", "%", 1),
}

# Limites de alerta de clima (% do tempo em que ocorreu), por talhão.
LIMITE_VENTO_KMH = 15      # vento acima de
LIMITE_UMIDADE_PCT = 45    # umidade relativa abaixo de
LIMITE_TEMP_AMB_C = 35     # temperatura ambiente acima de
GAP_EPISODIO_MIN = 10      # lacuna (min) que separa dois acionamentos de temperatura


# Faixa fisicamente plausível por indicador; fora dela a leitura é erro de sensor
# e fica fora da média. Sementes/ha: ~40 a 90 mil (milho/sorgo/algodão) e ~200 a
# 400 mil (soja). Não descartamos leituras por faixa (isso escondia a densidade quando o
# dado vinha em outra escala); o PDF avisa quando a média fica fora de 40 a 400 mil.
FAIXA_VALIDA = {}  # sem corte: o dado é mostrado como vem; o PDF só ALERTA se a média sair da faixa usual de sementes/ha


def serie_valida(efetivo: pd.DataFrame, coluna: str) -> pd.Series:
    v = pd.to_numeric(efetivo[coluna], errors="coerce")
    if coluna in FAIXA_VALIDA:
        lo, hi = FAIXA_VALIDA[coluna]
        v = v.where((v >= lo) & (v <= hi))
    return v


def pesos_para(coluna: str, efetivo: pd.DataFrame) -> pd.Series:
    """Peso da média: área (ha) para indicadores por hectare, senão tempo."""
    if coluna in PESO_POR_AREA and efetivo["area_ha"].sum() > 0:
        return efetivo["area_ha"]
    return efetivo["vl_tempo_segundos"]


def media_do_indicador(efetivo: pd.DataFrame, coluna: str, permite_negativo: bool):
    return media_ponderada_ou_branco(
        serie_valida(efetivo, coluna), pesos_para(coluna, efetivo), permite_negativo,
        zero_sem_dado=coluna in COLUNAS_QUALIDADE_EQUIP,
    )


def calcular_consumo_l_ha(efetivo: pd.DataFrame, area_ha: float):
    """Consumo em L/ha = total de litros consumidos no estado Efetivo
    dividido pela área operacional do talhão — diferente de uma média
    simples de vl_consumo_instantaneo (que é uma taxa em L/h): aqui
    somamos o volume real (taxa instantânea x duração de cada leitura)
    e dividimos pela área, para dar litros por hectare trabalhado."""
    consumo = pd.to_numeric(efetivo["vl_consumo_instantaneo"], errors="coerce")
    tempo = pd.to_numeric(efetivo["vl_tempo_segundos"], errors="coerce")
    consumo = consumo.where(consumo >= 0)  # negativo = sem leitura, mesma regra das demais colunas

    tem_dado_real = consumo.notna().any() and (consumo.fillna(0) != 0).any()
    if not tem_dado_real or not area_ha or area_ha <= 0:
        return np.nan

    mask = consumo.notna() & tempo.notna()
    litros_totais = (consumo[mask] * tempo[mask] / 3600.0).sum()
    return litros_totais / area_ha


def contar_alertas_temperatura(efetivo: pd.DataFrame):
    """Quantas leituras (estado Efetivo) ultrapassaram cada limiar de
    temperatura do motor em LIMIARES_TEMP_MOTOR."""
    temp = pd.to_numeric(efetivo["vl_temperatura_motor"], errors="coerce")
    return {limiar: int((temp > limiar).sum()) for limiar in LIMIARES_TEMP_MOTOR}


# ------------------------------------------------------------
# RESUMO POR PROCESSO x TALHÃO
# ------------------------------------------------------------

def gerar_resumo_processo_talhao(df: pd.DataFrame) -> pd.DataFrame:
    linhas = []

    for (cod_processo, talhao), grupo in df.groupby(["cd_operacao", "cd_talhao"], dropna=False):
        efetivo = grupo[grupo["eh_efetivo"]]

        area = efetivo["area_ha"].sum()
        if area <= AREA_MINIMA_TALHAO:
            continue  # regra: só talhões com > 5 ha no período, por processo

        dias = grupo.loc[grupo["dt_hr_local_inicial"].notna(), "dt_hr_local_inicial"].dt.date.nunique()
        equipamentos = sorted(grupo["cd_equipamento"].dropna().unique().tolist())
        horas_efetivas = efetivo["vl_tempo_segundos"].sum() / 3600.0
        horas_totais = grupo["vl_tempo_segundos"].sum() / 3600.0

        registro = {
            "Código processo": cod_processo,
            "Processo": rotulo_processo(cod_processo),
            "Talhão": talhao,
            "Área operacional (ha)": round(area, 2),
            "Dias com dados": dias,
            "Equipamentos": len(equipamentos),
            "Equipamentos (códigos)": ", ".join(str(e) for e in equipamentos),
            "Horas efetivas (h)": round(horas_efetivas, 2),
            "Horas totais (h)": round(horas_totais, 2),
            # "Rendimento Real" = horas efetivas / horas totais (motor ligado).
            # Mesmo cálculo que antes se chamava "Eficiência (efetivo/total)".
            "Rendimento Real (%)": (
                round(horas_efetivas / horas_totais * 100, 1) if horas_totais > 0 else np.nan
            ),
        }

        consumo_l_ha = calcular_consumo_l_ha(efetivo, area)
        registro["Consumo (L/ha)"] = round(consumo_l_ha, 1) if not pd.isna(consumo_l_ha) else np.nan

        for coluna, (nome, unidade, casas, permite_negativo) in TODAS_COLUNAS_MEDIA.items():
            if coluna not in grupo.columns:
                registro[nome] = np.nan
                continue
            media = media_do_indicador(efetivo, coluna, permite_negativo)
            registro[nome] = round(media, casas) if not pd.isna(media) else np.nan

        linhas.append(registro)

    if not linhas:
        raise ValueError(
            f"Nenhuma combinação Processo x Talhão ultrapassou {AREA_MINIMA_TALHAO:g} ha no período."
        )

    resumo = pd.DataFrame(linhas)
    return resumo.sort_values(
        ["Processo", "Área operacional (ha)"], ascending=[True, False]
    ).reset_index(drop=True)


def gerar_resumo_processo(resumo_talhao: pd.DataFrame, df: pd.DataFrame) -> pd.DataFrame:
    """Consolidado por processo (só os talhões que entraram no relatório,
    ou seja, > área mínima) — útil como cabeçalho de página, no mesmo
    espírito do Resumo_Processos do projeto antigo."""
    linhas = []
    for cod_processo, grupo_resumo in resumo_talhao.groupby("Código processo"):
        # Só os talhões >20ha desse processo (os mesmos que aparecem no
        # relatório) — não o bruto inteiro, que incluiria talhões pequenos
        # já descartados. Vale tanto para as médias quanto para os
        # alertas de temperatura ("considerando todos os talhões do
        # processo" = todos os que estão sendo mostrados).
        talhoes_incluidos = grupo_resumo["Talhão"].unique().tolist()
        grupo_bruto = df[(df["cd_operacao"] == cod_processo) & (df["cd_talhao"].isin(talhoes_incluidos))]
        efetivo = grupo_bruto[grupo_bruto["eh_efetivo"]]

        area_total = grupo_resumo["Área operacional (ha)"].sum()
        horas_efetivas = efetivo["vl_tempo_segundos"].sum() / 3600.0
        horas_totais = grupo_bruto["vl_tempo_segundos"].sum() / 3600.0

        registro = {
            "Código processo": cod_processo,
            "Processo": rotulo_processo(cod_processo),
            "Talhões (>5 ha)": len(grupo_resumo),
            "Área operacional (ha)": round(area_total, 2),
            "Equipamentos": grupo_bruto["cd_equipamento"].nunique(),
            "Horas efetivas (h)": round(horas_efetivas, 2),
            "Rendimento Real (%)": round(horas_efetivas / horas_totais * 100, 1) if horas_totais > 0 else np.nan,
        }

        consumo_l_ha = calcular_consumo_l_ha(efetivo, area_total)
        registro["Consumo (L/ha)"] = round(consumo_l_ha, 1) if not pd.isna(consumo_l_ha) else np.nan

        alertas = contar_alertas_temperatura(efetivo)
        for limiar, contagem in alertas.items():
            registro[f"Leituras Temp.Motor > {limiar}°C"] = contagem

        for coluna, (nome, unidade, casas, permite_negativo) in TODAS_COLUNAS_MEDIA.items():
            if coluna not in grupo_bruto.columns:
                registro[nome] = np.nan
                continue
            media = media_do_indicador(efetivo, coluna, permite_negativo)
            registro[nome] = round(media, casas) if not pd.isna(media) else np.nan
        linhas.append(registro)

    return pd.DataFrame(linhas).sort_values("Área operacional (ha)", ascending=False).reset_index(drop=True)


def _mediana_ponderada(v, w):
    o = np.argsort(v)
    v, w = v[o], w[o]
    c = np.cumsum(w)
    return float(v[np.searchsorted(c, c[-1] / 2.0)])


def gerar_variacao(df: pd.DataFrame, resumo_talhao: pd.DataFrame) -> pd.DataFrame:
    """Por Processo x Talhão e indicador: a MESMA média da tabela do PDF e a
    variação dos dados em torno dela. A variação é ROBUSTA (1,4826 x MAD em
    torno da mediana, ponderada): arranque/parada e picos isolados de sensor
    não inflam a variação quando a maior parte das leituras está perto da
    média. Também informa % das leituras (ponderadas) dentro de ±10% da média."""
    linhas = []
    for _, r in resumo_talhao.iterrows():
        g = df[(df["cd_operacao"] == r["Código processo"]) & (df["cd_talhao"] == str(r["Talhão"])) & df["eh_efetivo"]]
        for coluna, (nome, unidade, casas) in INDICADORES_VARIACAO.items():
            if coluna not in g.columns:
                continue
            media_tab = media_do_indicador(g, coluna, False)
            if pd.isna(media_tab):
                continue
            v = serie_valida(g, coluna)
            w = pesos_para(coluna, g)
            ok = v.notna() & (v > 0) & w.notna() & (w > 0)
            if ok.sum() < 5:
                continue
            v, w = v[ok].to_numpy(), w[ok].to_numpy()
            med = _mediana_ponderada(v, w)
            mad = _mediana_ponderada(np.abs(v - med), w)
            sigma = 1.4826 * mad
            media = float(media_tab)
            linhas.append({
                "Código processo": r["Código processo"], "Talhão": str(r["Talhão"]),
                "Indicador": nome, "Unidade": unidade, "Casas": casas,
                "Média": media, "Desvio": sigma,
                "CV (%)": sigma / media * 100 if media else np.nan,
                "Dentro ±10% (%)": float(w[np.abs(v - media) <= 0.10 * media].sum() / w.sum() * 100),
                "Leituras": int(ok.sum()),
            })
    return pd.DataFrame(linhas, columns=[
        "Código processo", "Talhão", "Indicador", "Unidade", "Casas", "Média", "Desvio",
        "CV (%)", "Dentro ±10% (%)", "Leituras"])


def _pct_tempo(g: pd.DataFrame, coluna: str, validar, condicao):
    """% do tempo (entre as leituras válidas) em que condicao(valor) foi verdadeira."""
    if coluna not in g.columns:
        return np.nan
    v = pd.to_numeric(g[coluna], errors="coerce")
    t = pd.to_numeric(g["vl_tempo_segundos"], errors="coerce").fillna(0)
    ok = v.notna() & validar(v) & (t > 0)
    if not ok.any():
        return np.nan
    return float(t[ok & condicao(v)].sum() / t[ok].sum() * 100)


def gerar_clima(df: pd.DataFrame, resumo_talhao: pd.DataFrame) -> pd.DataFrame:
    """% do tempo (de todos os dados gerados no talhão, em qualquer estado) com
    vento > 15 km/h, umidade < 45% e temperatura ambiente > 35°C."""
    linhas = []
    for _, r in resumo_talhao.iterrows():
        g = df[(df["cd_operacao"] == r["Código processo"]) & (df["cd_talhao"] == str(r["Talhão"]))]
        linhas.append({
            "Código processo": r["Código processo"], "Talhão": str(r["Talhão"]),
            f"% tempo vento > {LIMITE_VENTO_KMH}": _pct_tempo(g, "vl_velocidade_vento", lambda v: v >= 0, lambda v: v > LIMITE_VENTO_KMH),
            f"% tempo umidade < {LIMITE_UMIDADE_PCT}": _pct_tempo(g, "vl_umidade", lambda v: v > 0, lambda v: v < LIMITE_UMIDADE_PCT),
            f"% tempo temp. > {LIMITE_TEMP_AMB_C}": _pct_tempo(g, "vl_temperatura", lambda v: v != 0, lambda v: v > LIMITE_TEMP_AMB_C),
        })
    return pd.DataFrame(linhas)


def gerar_temp_motor(df: pd.DataFrame, resumo_talhao: pd.DataFrame) -> pd.DataFrame:
    """Temperatura do motor POR EQUIPAMENTO, uma vez por processo (todos os
    talhões que entraram no relatório): nº de acionamentos (episódios contínuos)
    acima de 100°C e de 105°C, no estado Efetivo."""
    linhas = []
    for cod, grupo_resumo in resumo_talhao.groupby("Código processo"):
        talhoes = [str(t) for t in grupo_resumo["Talhão"].unique()]
        base = df[(df["cd_operacao"] == cod) & (df["cd_talhao"].isin(talhoes))]
        for equip, ge in base.groupby("cd_equipamento"):
            ge = ge[ge["eh_efetivo"]].sort_values("dt_hr_local_inicial", kind="stable")
            temp = pd.to_numeric(ge.get("vl_temperatura_motor"), errors="coerce") if "vl_temperatura_motor" in ge else pd.Series(dtype=float)
            temp = temp.where(temp > 0)
            reg = {"Código processo": cod, "Processo": rotulo_processo(cod), "Equipamento": str(equip),
                   "Temp. máx (°C)": float(temp.max()) if temp.notna().any() else np.nan}
            for lim in LIMIARES_TEMP_MOTOR:
                if not temp.notna().any():
                    reg[f"Acionamentos > {lim}°C"] = np.nan
                    reg[f"Leituras > {lim}°C"] = np.nan
                    continue
                acima = (temp > lim).fillna(False)
                dt = ge["dt_hr_local_inicial"]
                gap = dt.diff() > pd.Timedelta(minutes=GAP_EPISODIO_MIN)
                inicio = acima & (~acima.shift(fill_value=False) | gap)
                reg[f"Acionamentos > {lim}°C"] = int(inicio.sum())
                reg[f"Leituras > {lim}°C"] = int(acima.sum())
            linhas.append(reg)
    return pd.DataFrame(linhas)


def gerar_estados(df: pd.DataFrame, resumo_talhao: pd.DataFrame) -> pd.DataFrame:
    """Horas por estado (E/M/F/D/A) por Processo x Talhão — só para as
    combinações que entraram no resumo (>20 ha), para dar contexto de
    quanto tempo foi manobra/deslocamento/parada em cada talhão."""
    validos = set(zip(resumo_talhao["Código processo"], resumo_talhao["Talhão"]))
    base = df[df.apply(lambda r: (r["cd_operacao"], r["cd_talhao"]) in validos, axis=1)]

    g = base.groupby(["cd_operacao", "cd_talhao", "cd_estado"])["vl_tempo_segundos"].sum().reset_index()
    g["Horas"] = round(g["vl_tempo_segundos"] / 3600.0, 2)
    g["Estado"] = g["cd_estado"].map(ESTADOS_LABEL).fillna(g["cd_estado"])
    g["Processo"] = g["cd_operacao"].map(rotulo_processo)
    g = g.rename(columns={"cd_talhao": "Talhão"})
    return g[["Processo", "Talhão", "Estado", "Horas"]].sort_values(["Processo", "Talhão", "Estado"])


# ------------------------------------------------------------
# EXPORTAÇÃO
# ------------------------------------------------------------

def formatar_excel(arquivo):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = load_workbook(arquivo)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="0A4F2E")
            cell.alignment = Alignment(horizontal="center")
        for col in range(1, ws.max_column + 1):
            letra = get_column_letter(col)
            maior = max(
                (len(str(ws.cell(r, col).value)) if ws.cell(r, col).value is not None else 0
                 for r in range(1, min(ws.max_row, 200) + 1)),
                default=10,
            )
            ws.column_dimensions[letra].width = min(max(maior + 2, 12), 42)
    wb.save(arquivo)


def salvar_resultados(pasta_resultados: Path, resumo_talhao, resumo_processo, estados, df,
                      variacao=None, clima=None, temp_motor=None, secoes=None):
    pasta_resultados.mkdir(parents=True, exist_ok=True)
    arquivo = pasta_resultados / "qualidade_operacional_processo_talhao.xlsx"

    with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
        resumo_processo.to_excel(writer, sheet_name="Resumo_Processo", index=False)
        resumo_talhao.to_excel(writer, sheet_name="Resumo_Processo_Talhao", index=False)
        estados.to_excel(writer, sheet_name="Estados", index=False)
        if variacao is not None:
            variacao.to_excel(writer, sheet_name="Variacao_Talhao", index=False)
        if clima is not None:
            clima.to_excel(writer, sheet_name="Clima_Talhao", index=False)
        if temp_motor is not None:
            temp_motor.to_excel(writer, sheet_name="TempMotor_Equip", index=False)
        # Seções Operação / Paradas / Equipamentos / Operadores (todas as leituras, sem metas)
        for aba, tabela in (secoes or {}).items():
            if not aba.startswith("_") and tabela is not None and not tabela.empty:
                tabela.to_excel(writer, sheet_name=aba, index=False)

        pd.DataFrame({
            "Informação": [
                "Linhas lidas do bruto",
                "Data inicial",
                "Data final",
                "Área mínima por talhão (ha)",
                "Combinações Processo x Talhão > área mínima",
                "Processos distintos no resultado",
                "Talhões distintos no resultado",
            ],
            "Valor": [
                len(df),
                df["dt_hr_local_inicial"].min(),
                df["dt_hr_local_inicial"].max(),
                AREA_MINIMA_TALHAO,
                len(resumo_talhao),
                resumo_talhao["Processo"].nunique(),
                resumo_talhao["Talhão"].nunique(),
            ],
        }).to_excel(writer, sheet_name="Validacao", index=False)

    formatar_excel(arquivo)
    return arquivo


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

def processar(caminho_bruto, pasta_resultados):
    df = carregar_bruto(caminho_bruto)
    resumo_talhao = gerar_resumo_processo_talhao(df)
    resumo_processo = gerar_resumo_processo(resumo_talhao, df)
    estados = gerar_estados(df, resumo_talhao)
    resumo_talhao["Talhão"] = resumo_talhao["Talhão"].astype(str)
    variacao = gerar_variacao(df, resumo_talhao)
    clima = gerar_clima(df, resumo_talhao)
    temp_motor = gerar_temp_motor(df, resumo_talhao)
    secoes = None
    if Path(caminho_bruto).suffix.lower() == ".csv":
        try:
            from operacao_secoes import gerar_secoes
            secoes = gerar_secoes(caminho_bruto)
            print(f"Seções Operação/Paradas/Equipamentos/Operadores: {len(secoes['Operacao'])} processo(s) > 20 ha")
        except Exception as e:  # noqa: BLE001 — não derruba o relatório de qualidade
            print(f"AVISO: não foi possível gerar as seções Operação/Paradas/Equipamentos/Operadores: {e}")
    saida = salvar_resultados(Path(pasta_resultados), resumo_talhao, resumo_processo, estados, df,
                              variacao, clima, temp_motor, secoes)
    print(f"\nIndicadores gerados: {saida}")
    print(f"Processos: {resumo_processo['Processo'].nunique()} | "
          f"Combinações Processo x Talhão (>{AREA_MINIMA_TALHAO:g} ha): {len(resumo_talhao)}")
    return saida, resumo_talhao, resumo_processo, estados


if __name__ == "__main__":
    import sys
    entrada = sys.argv[1] if len(sys.argv) > 1 else "dados_brutos/mapa_analitico.csv"
    saida_pasta = sys.argv[2] if len(sys.argv) > 2 else "resultados"
    processar(entrada, saida_pasta)
