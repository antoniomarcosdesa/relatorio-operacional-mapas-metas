# ============================================================
# OPERACAO_SECOES.PY
#
# Agregações das seções Operação / Paradas / Equipamentos / Operadores
# (portadas do pipeline SGPA) a partir do CSV bruto "mapa analítico".
#
# Diferenças em relação à fonte SGPA (já agregada):
#   - Tempo (h)      = soma de vl_tempo_segundos / 3600.
#   - Motor ligado   = tempo das leituras com vl_rpm > 0 (aproximação).
#   - Área           = vl_hectares_hora x tempo, só no estado Efetivo (E).
#   - Estados        = E trabalhando, F parada, M manobra, D/A deslocamento.
#   - Fim de turno / troca de turno / fim de expediente (desc_parada):
#       EXCLUÍDOS dos indicadores operacionais, MANTIDOS no
#       "Motor ligado parado" (mesma regra do SGPA).
#   - Parada sem motivo (*NAO_DEFINIDO*) conta como "SEM APONTAMENTO".
#   - Quantidade de paradas = nº de episódios (sequência contínua do mesmo
#     motivo no mesmo equipamento), não nº de leituras de ~5 s.
#   - Usa TODAS as linhas do arquivo (inclusive fora de talhão): deslocamento
#     e paradas acontecem muito fora dos talhões.
#   - Sem metas e sem comparações.
# ============================================================

import unicodedata
import numpy as np
import pandas as pd

from leitura_csv import parse_datas

AREA_MINIMA_PROCESSO = 20.0
AREA_MINIMA_OPERADOR_NAO_DEFINIDO = 10.0
GAP_EPISODIO_S = 300  # intervalo sem leituras que quebra um episódio de parada

PADROES_FIM_TURNO = [
    "FIM DE TURNO", "FIM TURNO", "FINAL DE TURNO", "FINAL TURNO",
    "FIM DE EXPEDIENTE", "FIM EXPEDIENTE", "FINAL DE EXPEDIENTE", "FINAL EXPEDIENTE",
    "TROCA DE TURNO", "TROCA TURNO",
]
SEM_APONTAMENTO = "SEM APONTAMENTO"

COLUNAS_LEITURA = [
    "cd_operacao", "cd_equipamento", "cd_estado", "dt_hr_local_inicial", "vl_tempo_segundos",
    "vl_hectares_hora", "vl_rpm", "vl_velocidade", "vl_consumo_instantaneo",
    "desc_parada", "cd_operador", "desc_operador",
]


def _norm(v) -> str:
    t = unicodedata.normalize("NFKD", str(v).strip().upper())
    return "".join(c for c in t if not unicodedata.combining(c))


def _media_pond(v, w):
    v = pd.to_numeric(v, errors="coerce")
    w = pd.to_numeric(w, errors="coerce")
    m = v.notna() & w.notna() & (w > 0)
    return float((v[m] * w[m]).sum() / w[m].sum()) if m.any() else np.nan


def _div(a, b):
    return float(a) / float(b) if b and not pd.isna(b) and float(b) != 0 else np.nan


# ------------------------------------------------------------
# LEITURA
# ------------------------------------------------------------

def carregar_leituras(caminho_csv) -> pd.DataFrame:
    """Todas as linhas do CSV (inclusive sem talhão), só com as colunas necessárias."""
    usecols = lambda c: str(c).strip() in COLUNAS_LEITURA
    df = pd.read_csv(caminho_csv, sep=";", encoding="utf-8", low_memory=False, usecols=usecols)
    df.columns = [str(c).strip() for c in df.columns]
    for c in COLUNAS_LEITURA:
        if c not in df.columns:
            df[c] = np.nan

    df["dt"] = parse_datas(df["dt_hr_local_inicial"])
    df["tempo_h"] = pd.to_numeric(df["vl_tempo_segundos"], errors="coerce").fillna(0) / 3600.0
    df["estado"] = df["cd_estado"].astype(str).str.strip().str.upper()
    for c in ("vl_rpm", "vl_velocidade", "vl_consumo_instantaneo", "vl_hectares_hora"):
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df["eh_trab"] = df["estado"] == "E"
    df["eh_parada"] = df["estado"] == "F"
    df["eh_manobra"] = df["estado"] == "M"
    df["eh_desloc"] = df["estado"].isin(["D", "A"])
    df["motor_on"] = df["vl_rpm"].fillna(0) > 0
    df["motor_h"] = np.where(df["motor_on"], df["tempo_h"], 0.0)
    df["area_ha"] = np.where(df["eh_trab"], df["vl_hectares_hora"].fillna(0) * df["tempo_h"], 0.0)
    # litros = taxa (L/h) x tempo (h); negativo/zero = sem leitura
    consumo = df["vl_consumo_instantaneo"].where(df["vl_consumo_instantaneo"] > 0)
    df["litros"] = (consumo * df["tempo_h"]).fillna(0.0)

    desc = df["desc_parada"].fillna("").astype(str).map(_norm)
    df["eh_fim_turno"] = desc.map(lambda t: bool(t) and any(p in t for p in PADROES_FIM_TURNO))
    sem_motivo = desc.isin(["", "*NAO_DEFINIDO*", "NAN", "-1"])
    df["motivo"] = np.where(
        df["eh_parada"],
        np.where(sem_motivo, SEM_APONTAMENTO, df["desc_parada"].fillna("").astype(str).str.strip()),
        "",
    )
    df["eh_sem_ap"] = df["eh_parada"] & (df["motivo"].map(_norm) == SEM_APONTAMENTO)
    df["entra_op"] = ~df["eh_fim_turno"]

    df["data"] = df["dt"].dt.date
    df["equip"] = df["cd_equipamento"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    cod_op = df["cd_operador"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    nome_op = df["desc_operador"].fillna("").astype(str).str.strip()
    nao_def = nome_op.map(_norm).isin(["", "*NAO_DEFINIDO*", "NAN"]) | cod_op.isin(["-1", "nan", ""])
    df["operador"] = np.where(nao_def, "-1 - * não definido *", cod_op + " - " + nome_op)
    df["cd_operacao"] = pd.to_numeric(df["cd_operacao"], errors="coerce")
    df = df[df["cd_operacao"].notna()].copy()
    df["cd_operacao"] = df["cd_operacao"].astype(int)
    return df


# ------------------------------------------------------------
# INDICADORES DE UM GRUPO (processo, equipamento ou operador)
# ------------------------------------------------------------

def indicadores_grupo(g: pd.DataFrame) -> dict:
    op = g[g["entra_op"]]
    trab = op[op["eh_trab"]]
    area = op["area_ha"].sum()
    tempo = op["tempo_h"].sum()
    motor = op["motor_h"].sum()
    h_trab = trab["tempo_h"].sum()
    dias = op["data"].nunique() if len(op) else 0

    # Consumo: só equipamentos que têm dado de consumo (não dilui a média).
    lit_eq = op.groupby("equip")["litros"].sum()
    com_dado = lit_eq[lit_eq > 0].index
    base_c = op[op["equip"].isin(com_dado)]
    consumo = _div(base_c["litros"].sum(), base_c["area_ha"].sum())

    # Motor ligado parado: inclui fim de turno (regra SGPA).
    motor_parada = g.loc[g["eh_parada"], "motor_h"].sum()
    motor_total = g["motor_h"].sum()

    par = op[op["eh_parada"]]
    desloc = op[op["eh_desloc"]]
    rpm_ok = lambda d: d[d["motor_on"]]
    return {
        "Área operacional (ha)": area,
        "Dias com dados": dias,
        "Equipamentos": op["equip"].nunique(),
        "Tempo total (h)": tempo,
        "Horas trabalhadas (h)": h_trab,
        "Motor ligado (h)": motor,
        "Motor ligado em parada (h)": motor_parada,
        "Motor ligado parado (%)": _div(motor_parada, motor_total) * 100,
        "Consumo total (l)": op["litros"].sum(),
        "Consumo (L/ha)": consumo,
        "Eficiência geral (%)": _div(h_trab, tempo) * 100,
        "Eficiência do motor (%)": _div(h_trab, motor) * 100,
        "Rendimento real (ha/h)": _div(area, motor),
        "Área/dia/equip (ha/dia/equip)": _div(area, dias * op["equip"].nunique()),
        "Jornada média (h/dia)": _div(tempo, dias),
        "Velocidade trabalhando (km/h)": _media_pond(
            trab["vl_velocidade"].where(trab["vl_velocidade"] > 0), trab["tempo_h"]),
        "RPM efetivo": _media_pond(rpm_ok(trab)["vl_rpm"], rpm_ok(trab)["tempo_h"]),
        "RPM parada": _media_pond(rpm_ok(par)["vl_rpm"], rpm_ok(par)["tempo_h"]),
        "RPM deslocamento": _media_pond(rpm_ok(desloc)["vl_rpm"], rpm_ok(desloc)["tempo_h"]),
        "Sem apontamento (h)": op.loc[op["eh_sem_ap"], "tempo_h"].sum(),
        "Sem apontamento (%)": _div(op.loc[op["eh_sem_ap"], "tempo_h"].sum(), par["tempo_h"].sum()) * 100,
        "Horas manobra (h)": op.loc[op["eh_manobra"], "tempo_h"].sum(),
        "Horas deslocamento (h)": desloc["tempo_h"].sum(),
    }


# ------------------------------------------------------------
# PARADAS (por motivo, em episódios)
# ------------------------------------------------------------

def _episodios(df: pd.DataFrame) -> pd.Series:
    """Id de episódio por linha: muda quando troca equipamento, estado ou motivo,
    ou quando há um intervalo > GAP_EPISODIO_S entre leituras do mesmo equipamento."""
    d = df.sort_values(["cd_operacao", "equip", "dt"], kind="mergesort")
    fim_ant = d["dt"].shift() + pd.to_timedelta(d["tempo_h"].shift() * 3600, unit="s")
    gap = (d["dt"] - fim_ant).dt.total_seconds()
    novo = (
        (d["equip"] != d["equip"].shift())
        | (d["cd_operacao"] != d["cd_operacao"].shift())
        | (d["estado"] != d["estado"].shift())
        | (d["motivo"] != d["motivo"].shift())
        | (gap > GAP_EPISODIO_S)
    )
    return novo.cumsum().reindex(df.index)


def gerar_paradas(df: pd.DataFrame, processos_validos) -> pd.DataFrame:
    base = df[df["cd_operacao"].isin(processos_validos) & df["eh_parada"] & df["entra_op"]].copy()
    cols = ["Código processo", "Operação", "Tempo (h)", "Motor Ligado (h)", "Motor Desligado (h)",
            "% do tempo de parada", "Quantidade", "Tempo médio por parada (min)"]
    if base.empty:
        return pd.DataFrame(columns=cols)

    base["ep"] = _episodios(base)
    g = base.groupby(["cd_operacao", "motivo"]).agg(
        **{"Tempo (h)": ("tempo_h", "sum"), "Motor Ligado (h)": ("motor_h", "sum"),
           "Quantidade": ("ep", "nunique")}
    ).reset_index().rename(columns={"cd_operacao": "Código processo", "motivo": "Operação"})
    g["Motor Desligado (h)"] = (g["Tempo (h)"] - g["Motor Ligado (h)"]).clip(lower=0)
    total = g.groupby("Código processo")["Tempo (h)"].transform("sum")
    g["% do tempo de parada"] = np.where(total > 0, g["Tempo (h)"] / total * 100, np.nan)
    g["Tempo médio por parada (min)"] = np.where(g["Quantidade"] > 0, g["Tempo (h)"] / g["Quantidade"] * 60, np.nan)
    return g.sort_values(["Código processo", "Tempo (h)"], ascending=[True, False])[cols].reset_index(drop=True)


# ------------------------------------------------------------
# DIÁRIO (área por dia)
# ------------------------------------------------------------

def gerar_diario(df: pd.DataFrame, processos_validos) -> pd.DataFrame:
    b = df[df["cd_operacao"].isin(processos_validos) & df["entra_op"] & df["data"].notna()]
    g = b.groupby(["cd_operacao", "data"])["area_ha"].sum().reset_index()
    g.columns = ["Código processo", "Data", "Área (ha)"]
    return g.sort_values(["Código processo", "Data"]).reset_index(drop=True)


# ------------------------------------------------------------
# API PRINCIPAL
# ------------------------------------------------------------

def gerar_secoes(caminho_csv) -> dict:
    """Devolve {'Operacao','Paradas','Equipamentos_Op','Operadores_Op','Diario_Op'} (DataFrames)."""
    df = carregar_leituras(caminho_csv)

    linhas = []
    for cod, g in df.groupby("cd_operacao"):
        r = indicadores_grupo(g)
        if r["Área operacional (ha)"] > AREA_MINIMA_PROCESSO:
            linhas.append({"Código processo": cod, **r})
    if not linhas:
        raise ValueError(f"Nenhum processo ultrapassou {AREA_MINIMA_PROCESSO:g} ha (todas as leituras).")
    operacao = pd.DataFrame(linhas).sort_values("Área operacional (ha)", ascending=False).reset_index(drop=True)
    validos = set(operacao["Código processo"])

    equip, oper = [], []
    sub = df[df["cd_operacao"].isin(validos)]
    for (cod, eq), g in sub.groupby(["cd_operacao", "equip"]):
        if not eq or eq.lower() == "nan":
            continue
        r = indicadores_grupo(g)
        if r["Tempo total (h)"] <= 0:
            continue
        equip.append({"Código processo": cod, "Equipamento (código)": eq, **r})
    for (cod, op), g in sub.groupby(["cd_operacao", "operador"]):
        r = indicadores_grupo(g)
        if r["Tempo total (h)"] <= 0:
            continue
        if op.startswith("-1") and r["Área operacional (ha)"] <= AREA_MINIMA_OPERADOR_NAO_DEFINIDO:
            continue
        oper.append({"Código processo": cod, "Operador": op, **r})

    ordenar = lambda d: d.sort_values(["Código processo", "Área operacional (ha)"], ascending=[True, False]).reset_index(drop=True)
    inicio, fim = df["dt"].min(), df["dt"].max()
    return {
        "Operacao": operacao,
        "Paradas": gerar_paradas(df, validos),
        "Equipamentos_Op": ordenar(pd.DataFrame(equip)) if equip else pd.DataFrame(),
        "Operadores_Op": ordenar(pd.DataFrame(oper)) if oper else pd.DataFrame(),
        "Diario_Op": gerar_diario(df, validos),
        "_periodo": (inicio, fim),
    }
