# ============================================================
# GERAR_MAPAS_VELOCIDADE_TALHAO.PY
#
# Gera UM arquivo PNG por combinação Processo x Talhão (>20 ha) com
# SÓ o mapa de velocidade (sem o mapa de estado) — usado para montar
# o slide "Mapas" de cada processo no PDF de qualidade operacional.
#
# Reaproveita a mesma lógica geométrica (projeção, filtro de outlier
# de GPS, filtro de salto entre passadas) já validada em
# gerar_mapas_operacao.py, mas desenha cada talhão em uma figura
# própria, dimensionada para ficar nítida quando embutida em grade de
# 2 a 4 imagens por página no PDF.
#
# Pode ser chamado:
#   - como script: gera todos os PNGs e lista no terminal
#   - como módulo: gerar_mapas_velocidade(caminho_csv) -> dict
#     {(cd_operacao, talhao_str): caminho_png}
# ============================================================

from pathlib import Path
import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from leitura_csv import ler_csv_bruto, parse_datas, achar_coluna
from dados_cliente import carregar_cliente
from matplotlib.collections import LineCollection
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.colors import LinearSegmentedColormap

# Tons de verde: verde mais escuro = velocidade menor, mais claro = maior
# Azul = valor menor -> tons de verde -> tons de vermelho = valor maior (cores vivas, sem amarelo/claro)
CMAP_VERDES = LinearSegmentedColormap.from_list(
    "azul_verde_vermelho", ["#0B2FC4", "#1E78E6", "#12A150", "#3DBE2B", "#F26B1D", "#D61F1F", "#8F0A1A"])

RE_COORD = re.compile(r"(-?\d+\.\d+)\s+(-?\d+\.\d+)")

AREA_MINIMA_TALHAO = 20.0  # ha — mesma regra do restante do pipeline

# Verde mais escuro = velocidade menor; verde mais claro = velocidade maior.
PALETA_VERDE = LinearSegmentedColormap.from_list("verde_vel", ["#05431F", "#2E8B4F", "#7ECF95", "#C2F0CD"])

# Indicador usado no mapa quando o talhão (no processo) tem o dado, na ordem de
# prioridade; sem nenhum deles, mapa de velocidade. A densidade de plantio só vale
# em processo de PLANTIO (o dado vem errado nas demais operações).
INDICADORES_MAPA = [
    ("vl_qtde_sementes_ha_total", "Densidade de plantio", "sem./ha", True),
    ("vl_rendimento_colheita", "Produtividade", "sc/ha", False),
    ("vl_taxa_aplicacao_solido_ha", "Taxa de aplicação", "kg/ha", False),
    ("vl_vazao_litros_ha", "Vazão", "L/ha", False),
]
COR_SEM_DADO = "#8A94A0"  # rastro efetivo de equipamento sem o dado do indicador

DPI_IMAGEM = 170  # alta o bastante pra ficar nítido em grades de 2-4 por página de PDF


# ------------------------------------------------------------
# LEITURA E GEOMETRIA (mesma lógica de gerar_mapas_operacao.py)
# ------------------------------------------------------------

class ColunaGeometriaAusente(Exception):
    """CSV sem a coluna wkt (coordenadas) — mapas não podem ser gerados."""


def extrair_pontos(wkt: str):
    """Todos os pontos (lon, lat) do WKT, na ordem."""
    return [(float(a), float(b)) for a, b in RE_COORD.findall(str(wkt))]


def extrair_ultimo_ponto(wkt: str):
    coords = RE_COORD.findall(str(wkt))
    if not coords:
        return np.nan, np.nan
    lon, lat = coords[-1]
    return float(lon), float(lat)


def carregar_base(caminho_csv: str) -> pd.DataFrame:
    df = ler_csv_bruto(caminho_csv)
    col_wkt = achar_coluna(df, "wkt")
    if col_wkt is None:
        raise ColunaGeometriaAusente(
            "O CSV não tem a coluna 'wkt' (coordenadas GPS) — não é possível desenhar os mapas."
        )
    if col_wkt != "wkt":
        df = df.rename(columns={col_wkt: "wkt"})
    df["cd_talhao"] = df["cd_talhao"].astype(str).str.strip()
    df = df[df["cd_talhao"].str.upper() != "*NAO_DEFINIDO*"].copy()
    # Mapa só usa o estado Efetivo: descarta o resto ANTES de interpretar as coordenadas (memória)
    df = df[df["cd_estado"].astype(str).str.strip().str.upper() == "E"].copy()

    df["dt_hr_local_inicial"] = parse_datas(df["dt_hr_local_inicial"])
    df["cd_estado"] = df["cd_estado"].astype(str).str.strip().str.upper()
    df["vl_tempo_segundos"] = pd.to_numeric(df["vl_tempo_segundos"], errors="coerce").fillna(0)
    df["vl_hectares_hora"] = pd.to_numeric(df["vl_hectares_hora"], errors="coerce").fillna(0)
    df["vl_velocidade"] = pd.to_numeric(df["vl_velocidade"], errors="coerce")
    for col, *_ in INDICADORES_MAPA:
        c = achar_coluna(df, col)
        df[col] = pd.to_numeric(df[c], errors="coerce") if c else np.nan
    df["vl_rendimento_colheita"] = df["vl_rendimento_colheita"] * (1000.0 / 60.0)  # t/ha -> sacas/ha (60 kg)

    df["area_ha"] = np.where(
        df["cd_estado"] == "E", df["vl_hectares_hora"] * df["vl_tempo_segundos"] / 3600.0, 0.0
    )

    df["pontos"] = df["wkt"].apply(extrair_pontos)
    pontos = df["wkt"].apply(extrair_ultimo_ponto)
    df["lon"] = pontos.apply(lambda p: p[0])
    df["lat"] = pontos.apply(lambda p: p[1])
    df = df.dropna(subset=["lon", "lat"]).reset_index(drop=True)

    return df


def combinacoes_validas(df: pd.DataFrame) -> pd.DataFrame:
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
    # Mapa SÓ com o estado Efetivo (trabalhando): deslocamento/manobra/parada
    # ficam de fora (nem desenho nem enquadramento).
    sub = df[(df["cd_operacao"] == cd_operacao) & (df["cd_talhao"] == str(cd_talhao)) & (df["cd_estado"] == "E")].copy()
    if sub.empty:
        return sub, np.array([]), np.array([]), 0.0, 0.0
    sub = sub.sort_values("dt_hr_local_inicial", kind="stable").reset_index(drop=True)

    lon0, lat0 = sub["lon"].mean(), sub["lat"].mean()
    x, y = projetar_metros(sub["lon"].to_numpy(), sub["lat"].to_numpy(), lon0, lat0)

    valido = remover_outliers_gps(x, y)
    sub = sub.loc[valido].reset_index(drop=True)
    x, y = x[valido], y[valido]

    return sub, x, y, lon0, lat0


def _limites_com_margem(x, y, margem_rel=0.04):
    largura = x.max() - x.min()
    altura = y.max() - y.min()
    m = max(largura, altura) * margem_rel + 8
    return x.min() - m, x.max() + m, y.min() - m, y.max() + m


# ------------------------------------------------------------
# DESENHO (só velocidade — sem o mapa de estado)
# ------------------------------------------------------------

def _br(v, casas=1):
    return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def desenhar_e_salvar(df, cd_operacao, cd_talhao, caminho_saida, vmax_percentil=97, cliente=None):
    sub, x, y, lon0, lat0 = preparar_talhao(df, cd_operacao, cd_talhao)
    if sub.empty:
        return None

    eh_efetivo = (sub["cd_estado"] == "E").to_numpy()

    # Indicador do mapa: 1º da lista com algum dado válido (>0) neste talhão/processo
    indicador = None
    for col, nome_ind, un_ind, so_plantio in INDICADORES_MAPA:
        if so_plantio and not (cliente is not None and cliente.processo_eh_plantio(cd_operacao)):
            continue
        if (sub[col] > 0).any():
            indicador = (col, nome_ind, un_ind)
            break
    if indicador:
        vel = sub[indicador[0]].where(sub[indicador[0]] > 0).to_numpy()  # sem dado (<=0) -> NaN
    else:
        vel = sub["vl_velocidade"].to_numpy()

    # Cada linha do mapa analítico é um LINESTRING (início -> fim) de ~5 s do
    # equipamento. Desenhamos cada linha como o SEU próprio trecho, sem ligar
    # linhas diferentes entre si (ligar por ordem de horário cruza equipamentos
    # e passadas e gerava os "riscos" retos atravessando o talhão).
    seg_ini, seg_fim, seg_vel, seg_ef = [], [], [], []
    for pts, v, ef in zip(sub["pontos"], vel, eh_efetivo):
        for (lo1, la1), (lo2, la2) in zip(pts[:-1], pts[1:]):
            seg_ini.append((lo1, la1)); seg_fim.append((lo2, la2)); seg_vel.append(v); seg_ef.append(ef)
    if not seg_ini:
        return None
    seg_ini, seg_fim = np.array(seg_ini), np.array(seg_fim)
    x1, y1 = projetar_metros(seg_ini[:, 0], seg_ini[:, 1], lon0, lat0)
    x2, y2 = projetar_metros(seg_fim[:, 0], seg_fim[:, 1], lon0, lat0)
    segmentos = np.stack([np.column_stack([x1, y1]), np.column_stack([x2, y2])], axis=1)
    cor_segmento = np.array(seg_vel, dtype=float)
    so_efetivo = np.array(seg_ef, dtype=bool)

    dist = np.hypot(x2 - x1, y2 - y1)
    limite_salto = max(np.percentile(dist, 99) * 3, 80)  # só descarta erro grosseiro de GPS
    ok_dist = dist <= limite_salto
    valido = ok_dist & so_efetivo & ~np.isnan(cor_segmento)
    sem_dado = ok_dist & so_efetivo & np.isnan(cor_segmento) if indicador else np.zeros_like(valido)

    vel_validas = vel[~np.isnan(vel)]
    if indicador:
        vmin = np.nanpercentile(vel_validas, 3)
        vmax = np.nanpercentile(vel_validas, vmax_percentil)
        if vmax - vmin < 1e-9:
            vmin, vmax = vmin * 0.9, vmax * 1.1 + 1e-9
    else:
        vmin = 0
        vmax = np.nanpercentile(vel_validas, vmax_percentil) if len(vel_validas) else 10

    # figura quadrada-ish, proporcional ao formato real do talhão, com
    # tamanho fixo o bastante pra caber 2-4 por página sem perder nitidez
    largura_dados = max(x.max() - x.min(), 1)
    altura_dados = max(y.max() - y.min(), 1)
    proporcao = np.clip(altura_dados / largura_dados, 0.35, 2.2)
    fig_largura = 5.6
    fig_altura = np.clip(fig_largura * proporcao + 0.55, 3.2, 7.5)

    fig, ax = plt.subplots(figsize=(fig_largura, fig_altura), dpi=DPI_IMAGEM)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("#F4F6F8")

    if sem_dado.any():  # passou no talhão mas esse equipamento não tem o dado
        ax.add_collection(LineCollection(segmentos[sem_dado], colors=COR_SEM_DADO, linewidths=2.4, capstyle="round", zorder=1))

    lc = LineCollection(
        segmentos[valido], array=cor_segmento[valido], cmap=CMAP_VERDES,
        norm=plt.Normalize(vmin=vmin, vmax=vmax), linewidths=2.6, capstyle="round", zorder=2,
    )
    ax.add_collection(lc)

    # Contorno do talhão (KML da fazenda), quando houver e casar com o talhão
    xs_lim, ys_lim = x, y
    if cliente is not None:
        aneis = cliente.contorno_talhao(cd_talhao)
        for anel in aneis:
            lons = np.array([p[0] for p in anel])
            lats = np.array([p[1] for p in anel])
            ax_x, ax_y = projetar_metros(lons, lats, lon0, lat0)
            ax.plot(ax_x, ax_y, color="#1E3A8A", linewidth=1.3, linestyle="--", zorder=3)
            xs_lim = np.concatenate([xs_lim, ax_x])
            ys_lim = np.concatenate([ys_lim, ax_y])

    xmin, xmax, ymin, ymax = _limites_com_margem(xs_lim, ys_lim)
    ax.set_xlim(xmin, xmax)
    ax.set_ylim(ymin, ymax)
    ax.set_aspect("equal")
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_color("#CBD5E1")
        spine.set_linewidth(0.8)

    cbar = fig.colorbar(lc, ax=ax, pad=0.02, fraction=0.045)
    cbar.set_label(indicador[2] if indicador else "km/h", fontsize=7.5, color="#334155")
    cbar.ax.tick_params(labelsize=6.5, colors="#334155")

    area = sub.loc[sub["cd_estado"] == "E", "area_ha"].sum()
    media_ind = float(np.nanmean(vel_validas)) if len(vel_validas) else float("nan")
    rotulo_ind = f"{indicador[1]} méd. {_br(media_ind, 0 if indicador[2] == 'sem./ha' else 1)} {indicador[2]}" if indicador \
        else f"méd. {_br(sub['vl_velocidade'].mean(), 1)} km/h"
    ax.set_title(
        f"Talhão {cliente.rotulo_talhao(cd_talhao) if cliente else cd_talhao}  ·  "
        f"{_br(area, 1)} ha  ·  {rotulo_ind}",
        fontsize=9.5, fontweight="bold", color="#0F172A", pad=6,
    )

    if sem_dado.any():
        fig.text(0.5, 0.005, "cinza = equipamento passou no talhão sem dado do indicador", ha="center", fontsize=6.5, color="#64748B")
    fig.tight_layout()
    fig.savefig(caminho_saida, facecolor="white", bbox_inches="tight")
    plt.close(fig)
    return caminho_saida


# ------------------------------------------------------------
# API PRINCIPAL
# ------------------------------------------------------------

def gerar_mapas_velocidade(caminho_csv, pasta_saida="relatorios/mapas_velocidade"):
    """Gera 1 PNG (só velocidade) por combinação Processo x Talhão
    > área mínima. Devolve {(cd_operacao, talhao_str): caminho_png}."""
    pasta_saida = Path(pasta_saida)
    pasta_saida.mkdir(exist_ok=True, parents=True)

    cliente = carregar_cliente()
    df = carregar_base(caminho_csv)
    cliente.registrar_chaves(df["cd_talhao"].unique())
    validos = combinacoes_validas(df)

    imagens = {}
    for _, row in validos.iterrows():
        cd_operacao, cd_talhao = row["cd_operacao"], row["cd_talhao"]
        caminho = pasta_saida / f"mapa_vel_proc{cd_operacao}_talhao{cd_talhao}.png"
        resultado = desenhar_e_salvar(df, cd_operacao, cd_talhao, caminho, cliente=cliente)
        if resultado is not None:
            imagens[(cd_operacao, str(cd_talhao))] = resultado
            print(f"Mapa de velocidade: Processo {cd_operacao} / Talhão {cd_talhao} -> {caminho}")

    return imagens


if __name__ == "__main__":
    import sys
    caminho_csv = sys.argv[1] if len(sys.argv) > 1 else "dados_brutos/mapa_analitico.csv"
    pasta_saida = sys.argv[2] if len(sys.argv) > 2 else "relatorios/mapas_velocidade"
    imagens = gerar_mapas_velocidade(caminho_csv, pasta_saida)
    print(f"\n{len(imagens)} mapa(s) de velocidade gerado(s) em {pasta_saida}")
