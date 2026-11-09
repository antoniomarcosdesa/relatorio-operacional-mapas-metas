# ============================================================
# GERAR_MAPA_VELOCIDADE.PY
#
# Usa a coluna "wkt" do mapa analítico (POINT/LINESTRING com as
# coordenadas GPS de cada leitura) para desenhar o trajeto real do
# equipamento dentro de um talhão, ligando os pontos em linha e
# colorindo cada trecho pela velocidade — o mesmo tipo de mapa que o
# SGPA mostra na tela de "Detalhes de Frota".
#
# Entrada: CSV bruto (mapa analítico) + cd_operacao (processo) +
#          cd_talhao a visualizar.
# Saída:   PNG do mapa.
# ============================================================

from pathlib import Path
import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection

RE_COORD = re.compile(r"(-?\d+\.\d+)\s+(-?\d+\.\d+)")


def extrair_ultimo_ponto(wkt: str):
    """POINT tem 1 par de coordenadas; LINESTRING tem vários (é o
    rastro desde a última leitura até a atual) — pegamos o ÚLTIMO par,
    que é a posição no momento desta leitura."""
    coords = RE_COORD.findall(str(wkt))
    if not coords:
        return np.nan, np.nan
    lon, lat = coords[-1]
    return float(lon), float(lat)


def projetar_metros(lon, lat):
    """Projeção equiretangular simples, centrada na média da área —
    suficiente para um talhão (poucos km²), evita distorção visual
    entre eixo X e Y num gráfico comum (que trabalha em graus)."""
    lat0 = np.radians(lat.mean())
    lon0 = lon.mean()
    x = (lon - lon0) * 111_320 * np.cos(lat0)
    y = (lat - lat.mean()) * 110_540
    return x, y


def remover_outliers_gps(x, y):
    """Descarta leituras de GPS claramente erradas (salto de vários km
    para fora do talhão) usando distância robusta à mediana — comum
    em rastreadores GPS de campo, um ponto ruim não pode esticar a
    escala do mapa inteiro."""
    mediana_x, mediana_y = np.median(x), np.median(y)
    dist = np.hypot(x - mediana_x, y - mediana_y)
    limite = max(np.percentile(dist, 99) * 4, 200)  # nunca corta abaixo de 200 m do centro
    valido = dist <= limite
    return valido


def carregar_trajeto(caminho_csv, cd_operacao, cd_talhao, estado="E"):
    df = pd.read_csv(caminho_csv, sep=";", encoding="utf-8")
    df.columns = [str(c).strip() for c in df.columns]

    df["cd_talhao"] = df["cd_talhao"].astype(str).str.strip()
    sub = df[
        (df["cd_operacao"] == cd_operacao)
        & (df["cd_talhao"] == str(cd_talhao))
        & (df["cd_estado"] == estado)
    ].copy()

    if sub.empty:
        raise ValueError(f"Nenhuma leitura para processo={cd_operacao}, talhão={cd_talhao}, estado={estado}.")

    sub["dt_hr_local_inicial"] = pd.to_datetime(sub["dt_hr_local_inicial"], dayfirst=True, errors="coerce")
    sub = sub.sort_values("dt_hr_local_inicial", kind="stable").reset_index(drop=True)

    pontos = sub["wkt"].apply(extrair_ultimo_ponto)
    sub["lon"] = pontos.apply(lambda p: p[0])
    sub["lat"] = pontos.apply(lambda p: p[1])
    sub = sub.dropna(subset=["lon", "lat"]).reset_index(drop=True)

    sub["vl_velocidade"] = pd.to_numeric(sub["vl_velocidade"], errors="coerce")

    return sub


def desenhar_mapa_velocidade(sub, titulo, subtitulo, caminho_saida, cmap="RdYlGn", vmax_percentil=97):
    x, y = projetar_metros(sub["lon"], sub["lat"])
    x, y = x.to_numpy(), y.to_numpy()
    vel = sub["vl_velocidade"].to_numpy()

    # remove leituras de GPS com erro grosseiro (salto de km) ANTES de
    # calcular limites do gráfico — senão um ponto ruim estica a escala
    # inteira e o talhão real fica minúsculo no canto.
    valido_gps = remover_outliers_gps(x, y)
    n_removidos = int((~valido_gps).sum())
    if n_removidos:
        print(f"Aviso: {n_removidos} leitura(s) de GPS descartada(s) por erro grosseiro de posição.")
    x, y, vel = x[valido_gps], y[valido_gps], vel[valido_gps]

    # segmentos entre pontos consecutivos, cor = velocidade média do trecho
    pontos = np.column_stack([x, y])
    segmentos = np.stack([pontos[:-1], pontos[1:]], axis=1)
    cor_segmento = (vel[:-1] + vel[1:]) / 2

    # descarta saltos de GPS grandes (equipamento "teleportando" entre
    # passadas não-contíguas do talhão) para não desenhar linhas retas
    # cruzando o talhão inteiro
    dist = np.hypot(segmentos[:, 1, 0] - segmentos[:, 0, 0], segmentos[:, 1, 1] - segmentos[:, 0, 1])
    limite_salto = np.nanpercentile(dist, 99) * 3 if len(dist) else 0
    valido = dist <= max(limite_salto, 30)  # nunca corta trechos normais de até 30 m

    vmin = 0
    vmax = np.nanpercentile(vel, vmax_percentil)

    largura_dados = x.max() - x.min()
    altura_dados = y.max() - y.min()
    proporcao = max(altura_dados / largura_dados, 0.18)
    fig_largura = 10.5
    fig_altura = min(max(fig_largura * proporcao + 1.3, 4.5), 11)

    fig, ax = plt.subplots(figsize=(fig_largura, fig_altura), dpi=160)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("#F4F6F8")

    lc = LineCollection(
        segmentos[valido], array=cor_segmento[valido], cmap=cmap,
        norm=plt.Normalize(vmin=vmin, vmax=vmax), linewidths=2.6, capstyle="round",
    )
    ax.add_collection(lc)

    ax.set_xlim(x.min() - 15, x.max() + 15)
    ax.set_ylim(y.min() - 15, y.max() + 15)
    ax.set_aspect("equal")
    ax.set_xlabel("Leste-Oeste (m)", fontsize=9, color="#475569")
    ax.set_ylabel("Norte-Sul (m)", fontsize=9, color="#475569")
    ax.tick_params(colors="#64748B", labelsize=8)
    for spine in ax.spines.values():
        spine.set_color("#CBD5E1")
    ax.grid(True, color="#E2E8F0", linewidth=0.6, alpha=0.7)

    cbar = fig.colorbar(lc, ax=ax, pad=0.02, fraction=0.046)
    cbar.set_label("Velocidade (km/h)", fontsize=9, color="#334155")
    cbar.ax.tick_params(labelsize=8, colors="#334155")

    fig.suptitle(titulo, fontsize=13, fontweight="bold", color="#0F172A", x=0.06, y=1 - 0.32 / fig_altura, ha="left")
    fig.text(0.06, 1 - 0.62 / fig_altura, subtitulo, fontsize=9.5, color="#64748B", ha="left")

    fig.tight_layout(rect=[0, 0, 1, 1 - 0.85 / fig_altura])
    fig.savefig(caminho_saida, facecolor="white", bbox_inches="tight")
    plt.close(fig)
    return caminho_saida


if __name__ == "__main__":
    import sys

    caminho_csv = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/OURIVES_DADOS_MAPA_ANALITICO_BRUTOS_31-08-2026_18_55.csv"
    cd_operacao = int(sys.argv[2]) if len(sys.argv) > 2 else 1401
    cd_talhao = sys.argv[3] if len(sys.argv) > 3 else "52"

    sub = carregar_trajeto(caminho_csv, cd_operacao, cd_talhao)

    dias = sub["dt_hr_local_inicial"].dt.date.nunique()
    area_aprox = None  # não recalculado aqui — é só o mapa
    titulo = f"Mapa de velocidade — Processo {cd_operacao} · Talhão {cd_talhao}"
    subtitulo = (
        f"{len(sub):,} leituras em estado Efetivo · {dias} dia(s) · "
        f"velocidade média {sub['vl_velocidade'].mean():.2f} km/h"
    ).replace(",", ".")

    saida = Path(f"relatorios/mapa_velocidade_proc{cd_operacao}_talhao{cd_talhao}.png")
    saida.parent.mkdir(exist_ok=True)
    desenhar_mapa_velocidade(sub, titulo, subtitulo, saida)
    print(f"Mapa gerado: {saida}")