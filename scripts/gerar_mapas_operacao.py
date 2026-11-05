# ============================================================
# GERAR_MAPAS_OPERACAO.PY
#
# Para cada Processo (cd_operacao), gera uma grade de mapas — um par
# [MAPA DE ESTADO | MAPA DE VELOCIDADE] por Talhão (>20 ha) — no mesmo
# espírito do bloco "MAPAS" do relatório SGPA de referência:
#   nome do processo + datas do período, depois quadros de mapas.
#
# MAPA DE ESTADO: rastro colorido por cd_estado
#   E = Efetivo/Trabalhando (verde) · D = Deslocamento (azul) ·
#   M = Manobra (amarelo) · F = Parada (vermelho) ·
#   A = Deslocamento p/ reabastecimento (roxo)
#
# MAPA DE VELOCIDADE: rastro colorido por vl_velocidade (estado
# Efetivo), RdYlGn — mesmo que gerar_mapa_velocidade.py já validado.
#
# Usa a coluna "wkt" (POINT/LINESTRING) do mapa analítico para extrair
# as coordenadas GPS reais de cada leitura.
# ============================================================

from pathlib import Path
import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection
from matplotlib.lines import Line2D

RE_COORD = re.compile(r"(-?\d+\.\d+)\s+(-?\d+\.\d+)")

AREA_MINIMA_TALHAO = 20.0

ESTADO_COR = {
    "E": "#1E8E3E",   # verde — efetivo/trabalhando
    "D": "#1A73E8",   # azul — deslocamento
    "M": "#F2B400",   # amarelo — manobra
    "F": "#D93025",   # vermelho — parada
    "A": "#9334E6",   # roxo — deslocamento p/ reabastecimento
}
ESTADO_LABEL = {
    "E": "Trabalhando",
    "D": "Deslocamento",
    "M": "Manobra",
    "F": "Parada",
    "A": "Reabastecimento",
}

CATALOGO_OPERACOES = {
    # 1403: "GRADAGEM NIVELADORA",
    # 1401: "PLANTIO",
}


def rotulo_processo(codigo) -> str:
    return CATALOGO_OPERACOES.get(codigo, f"Processo {codigo}")


# ------------------------------------------------------------
# LEITURA E GEOMETRIA
# ------------------------------------------------------------

def extrair_ultimo_ponto(wkt: str):
    coords = RE_COORD.findall(str(wkt))
    if not coords:
        return np.nan, np.nan
    lon, lat = coords[-1]
    return float(lon), float(lat)


def carregar_base(caminho_csv: str) -> pd.DataFrame:
    df = pd.read_csv(caminho_csv, sep=";", encoding="utf-8")
    df.columns = [str(c).strip() for c in df.columns]
    df["cd_talhao"] = df["cd_talhao"].astype(str).str.strip()
    df = df[df["cd_talhao"].str.upper() != "*NAO_DEFINIDO*"].copy()

    df["dt_hr_local_inicial"] = pd.to_datetime(df["dt_hr_local_inicial"], dayfirst=True, errors="coerce")
    df["cd_estado"] = df["cd_estado"].astype(str).str.strip().str.upper()
    df["vl_tempo_segundos"] = pd.to_numeric(df["vl_tempo_segundos"], errors="coerce").fillna(0)
    df["vl_hectares_hora"] = pd.to_numeric(df["vl_hectares_hora"], errors="coerce").fillna(0)
    df["vl_velocidade"] = pd.to_numeric(df["vl_velocidade"], errors="coerce")

    df["area_ha"] = np.where(
        df["cd_estado"] == "E", df["vl_hectares_hora"] * df["vl_tempo_segundos"] / 3600.0, 0.0
    )

    pontos = df["wkt"].apply(extrair_ultimo_ponto)
    df["lon"] = pontos.apply(lambda p: p[0])
    df["lat"] = pontos.apply(lambda p: p[1])
    df = df.dropna(subset=["lon", "lat"]).reset_index(drop=True)

    return df


def combinacoes_validas(df: pd.DataFrame) -> pd.DataFrame:
    """Processo x Talhão com área > 20 ha — mesma regra do relatório de qualidade."""
    g = df.groupby(["cd_operacao", "cd_talhao"])["area_ha"].sum().reset_index()
    g = g[g["area_ha"] > AREA_MINIMA_TALHAO]
    return g.sort_values(["cd_operacao", "area_ha"], ascending=[True, False]).reset_index(drop=True)


def projetar_metros(lon, lat, lon0, lat0):
    lat0r = np.radians(lat0)
    x = (lon - lon0) * 111_320 * np.cos(lat0r)
    y = (lat - lat0) * 110_540
    return x, y


def remover_outliers_gps(x, y):
    mediana_x, mediana_y = np.median(x), np.median(y)
    dist = np.hypot(x - mediana_x, y - mediana_y)
    limite = max(np.percentile(dist, 99) * 4, 200)
    return dist <= limite


def preparar_talhao(df: pd.DataFrame, cd_operacao, cd_talhao):
    """Devolve x, y (metros), cd_estado e vl_velocidade de todas as
    leituras do Processo x Talhão, ordenadas no tempo, com outliers
    de GPS removidos (um único filtro, usado nos dois mapas para
    manter os dois na mesma escala/orientação)."""
    sub = df[(df["cd_operacao"] == cd_operacao) & (df["cd_talhao"] == str(cd_talhao))].copy()
    sub = sub.sort_values("dt_hr_local_inicial", kind="stable").reset_index(drop=True)

    lon0, lat0 = sub["lon"].mean(), sub["lat"].mean()
    x, y = projetar_metros(sub["lon"].to_numpy(), sub["lat"].to_numpy(), lon0, lat0)

    valido = remover_outliers_gps(x, y)
    sub = sub.loc[valido].reset_index(drop=True)
    x, y = x[valido], y[valido]

    return sub, x, y


# ------------------------------------------------------------
# DESENHO DE UM MAPA (eixo único) — reaproveitado pelos dois tipos
# ------------------------------------------------------------

def _limites_com_margem(x, y, margem_rel=0.04):
    largura = x.max() - x.min()
    altura = y.max() - y.min()
    m = max(largura, altura) * margem_rel + 8
    return x.min() - m, x.max() + m, y.min() - m, y.max() + m


def desenhar_mapa_estado(ax, sub, x, y):
    estados = sub["cd_estado"].to_numpy()
    segmentos = np.stack([np.column_stack([x[:-1], y[:-1]]), np.column_stack([x[1:], y[1:]])], axis=1)

    dist = np.hypot(segmentos[:, 1, 0] - segmentos[:, 0, 0], segmentos[:, 1, 1] - segmentos[:, 0, 1])
    limite_salto = max(np.nanpercentile(dist, 99) * 3, 30) if len(dist) else 30
    valido = dist <= limite_salto

    cores = np.array([ESTADO_COR.get(e, "#9AA5B1") for e in estados[:-1]])

    ax.set_facecolor("#F4F6F8")
    for estado, cor in ESTADO_COR.items():
        mask = valido & (estados[:-1] == estado)
        if not mask.any():
            continue
        lc = LineCollection(segmentos[mask], colors=cor, linewidths=1.6, capstyle="round", zorder=2)
        ax.add_collection(lc)

    xmin, xmax, ymin, ymax = _limites_com_margem(x, y)
    ax.set_xlim(xmin, xmax)
    ax.set_ylim(ymin, ymax)
    ax.set_aspect("equal")
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_color("#CBD5E1")
        spine.set_linewidth(0.8)


def desenhar_mapa_velocidade(ax, sub, x, y, vmax_percentil=97):
    vel = sub["vl_velocidade"].to_numpy()
    eh_efetivo = (sub["cd_estado"] == "E").to_numpy()

    segmentos = np.stack([np.column_stack([x[:-1], y[:-1]]), np.column_stack([x[1:], y[1:]])], axis=1)
    cor_segmento = (vel[:-1] + vel[1:]) / 2
    so_efetivo = eh_efetivo[:-1] & eh_efetivo[1:]

    dist = np.hypot(segmentos[:, 1, 0] - segmentos[:, 0, 0], segmentos[:, 1, 1] - segmentos[:, 0, 1])
    limite_salto = max(np.nanpercentile(dist, 99) * 3, 30) if len(dist) else 30
    valido = (dist <= limite_salto) & so_efetivo & ~np.isnan(cor_segmento)

    vel_validas = vel[eh_efetivo]
    vmax = np.nanpercentile(vel_validas, vmax_percentil) if len(vel_validas) else 10

    ax.set_facecolor("#F4F6F8")
    # trajeto não-efetivo em cinza claro, de fundo, pra dar contexto do talhão inteiro
    nao_efetivo = valido_fundo = dist <= limite_salto
    lc_fundo = LineCollection(segmentos[nao_efetivo & ~so_efetivo], colors="#D7DCE2", linewidths=1.2, capstyle="round", zorder=1)
    ax.add_collection(lc_fundo)

    lc = LineCollection(
        segmentos[valido], array=cor_segmento[valido], cmap="RdYlGn",
        norm=plt.Normalize(vmin=0, vmax=vmax), linewidths=1.8, capstyle="round", zorder=2,
    )
    ax.add_collection(lc)

    xmin, xmax, ymin, ymax = _limites_com_margem(x, y)
    ax.set_xlim(xmin, xmax)
    ax.set_ylim(ymin, ymax)
    ax.set_aspect("equal")
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_color("#CBD5E1")
        spine.set_linewidth(0.8)

    return lc, vmax


# ------------------------------------------------------------
# GRADE POR PROCESSO
# ------------------------------------------------------------

def gerar_grade_processo(df, cd_operacao, lista_talhoes, caminho_saida):
    """1 linha de título (processo + datas) + N linhas de [mapa estado | mapa velocidade],
    uma linha por talhão (>5ha), igual ao bloco MAPAS do modelo SGPA."""
    n = len(lista_talhoes)
    fig_largura = 9.0
    altura_por_linha = 3.5
    fig_altura = 0.9 + altura_por_linha * n

    fig = plt.figure(figsize=(fig_largura, fig_altura), dpi=160)
    fig.patch.set_facecolor("white")

    gs = fig.add_gridspec(
        n, 2, left=0.05, right=0.95, top=1 - 0.95 / fig_altura, bottom=0.15 / fig_altura,
        hspace=0.32, wspace=0.08,
    )

    datas_min, datas_max = [], []

    for i, cd_talhao in enumerate(lista_talhoes):
        sub, x, y = preparar_talhao(df, cd_operacao, cd_talhao)
        if sub.empty:
            continue

        d0, d1 = sub["dt_hr_local_inicial"].min(), sub["dt_hr_local_inicial"].max()
        datas_min.append(d0)
        datas_max.append(d1)
        area = sub.loc[sub["cd_estado"] == "E", "area_ha"].sum()

        ax_estado = fig.add_subplot(gs[i, 0])
        desenhar_mapa_estado(ax_estado, sub, x, y)
        ax_estado.set_title(
            f"Talhão {cd_talhao} — Estado  ·  {area:,.1f} ha".replace(",", "."),
            fontsize=8.5, color="#334155", pad=4,
        )

        ax_vel = fig.add_subplot(gs[i, 1])
        desenhar_mapa_velocidade(ax_vel, sub, x, y)
        vel_media = sub.loc[sub["cd_estado"] == "E", "vl_velocidade"].mean()
        ax_vel.set_title(
            f"Talhão {cd_talhao} — Velocidade  ·  méd. {vel_media:.1f} km/h",
            fontsize=8.5, color="#334155", pad=4,
        )

    # ---- legenda de estado, uma única vez, no rodapé ----
    handles = [
        Line2D([0], [0], color=cor, lw=3, label=ESTADO_LABEL[est])
        for est, cor in ESTADO_COR.items()
    ]
    fig.legend(
        handles=handles, loc="lower center", ncol=len(handles), frameon=False,
        fontsize=7.5, bbox_to_anchor=(0.5, 0.0),
    )

    # ---- título com nome do processo + datas ----
    if datas_min and datas_max:
        ini = min(datas_min).strftime("%d/%m/%Y")
        fim = max(datas_max).strftime("%d/%m/%Y")
        periodo = f"{ini} a {fim}" if ini != fim else ini
    else:
        periodo = "-"

    fig.suptitle(
        f"{rotulo_processo(cd_operacao)}  ·  {periodo}",
        fontsize=13, fontweight="bold", color="#0F172A", x=0.05, y=1 - 0.42 / fig_altura, ha="left",
    )

    fig.savefig(caminho_saida, facecolor="white", bbox_inches="tight")
    plt.close(fig)
    return caminho_saida


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

def gerar_todos(caminho_csv, pasta_saida="relatorios"):
    pasta_saida = Path(pasta_saida)
    pasta_saida.mkdir(exist_ok=True, parents=True)

    df = carregar_base(caminho_csv)
    validos = combinacoes_validas(df)

    arquivos = []
    for cd_operacao, grupo in validos.groupby("cd_operacao"):
        talhoes = grupo["cd_talhao"].tolist()
        saida = pasta_saida / f"mapas_processo_{cd_operacao}.png"
        gerar_grade_processo(df, cd_operacao, talhoes, saida)
        arquivos.append(saida)
        print(f"Processo {cd_operacao} ({len(talhoes)} talhão/ões) -> {saida}")

    return arquivos


if __name__ == "__main__":
    import sys
    caminho_csv = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/OURIVES_DADOS_MAPA_ANALITICO_BRUTOS_31-08-2026_18_55.csv"
    gerar_todos(caminho_csv)
