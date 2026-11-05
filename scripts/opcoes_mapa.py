# ============================================================
# OPCOES_MAPA.PY
#
# Grupos de operação (PLANTIO, COLHEITA, PULVERIZAÇÃO, PREPARO DE SOLO,
# DISTRIBUIÇÃO A LANÇO) e os indicadores que podem ser escolhidos para os
# MAPAS de cada grupo (opcional). A escolha vem de opcoes_mapa.json
# (o site grava esse arquivo; no uso local pode ser criado à mão):
#
#   {"PLANTIO": ["SEMENTES", "RPM"], "COLHEITA": ["MASSA"]}
#
# - 1 indicador  -> 1 mapa por talhão naquele indicador;
# - 2 indicadores-> 2 mapas por talhão (1 talhão por página, mapas lado a lado);
# - grupo ausente/vazio -> comportamento automático de sempre
#   (densidade/produtividade/taxa/vazão se houver dado; senão velocidade).
#
# O grupo de cada processo vem da coluna "grupo" da aba Processos do Excel do
# cliente (planilha "de-para"); sem a coluna, é deduzido da descrição.
# ============================================================

import json
import unicodedata
from pathlib import Path

GRUPOS = {
    "PLANTIO": "PLANTIO",
    "COLHEITA": "COLHEITA",
    "PULVERIZACAO": "PULVERIZAÇÃO",
    "PREPARO_SOLO": "PREPARO DE SOLO",
    "LANCO": "DISTRIBUIÇÃO A LANÇO",
}

# chave -> (coluna do CSV, nome exibido, unidade)
INDICADORES = {
    "VAZAO": ("vl_vazao_litros_ha", "Vazão", "L/ha"),
    "TAXA": ("vl_taxa_aplicacao_solido_ha", "Taxa de aplicação", "kg/ha"),
    "SEMENTES": ("vl_qtde_sementes_ha_total", "Densidade de plantio", "sem./ha"),
    "MASSA": ("vl_rendimento_colheita", "Produtividade", "sc/ha"),
    "RPM": ("vl_rpm", "RPM", "RPM"),
    "VELOCIDADE": ("vl_velocidade", "Velocidade", "km/h"),
    "CONSUMO": ("vl_consumo_instantaneo", "Consumo", "L/h"),
}

INDICADORES_POR_GRUPO = {
    "PLANTIO": ["SEMENTES", "RPM", "VELOCIDADE", "CONSUMO"],
    "COLHEITA": ["MASSA", "RPM", "VELOCIDADE", "CONSUMO"],
    "PULVERIZACAO": ["VAZAO", "RPM", "VELOCIDADE", "CONSUMO"],
    "PREPARO_SOLO": ["RPM", "VELOCIDADE", "CONSUMO"],
    "LANCO": ["TAXA", "RPM", "VELOCIDADE", "CONSUMO"],
}

ARQUIVO_OPCOES = "opcoes_mapa.json"


def _norm(t) -> str:
    t = unicodedata.normalize("NFKD", str(t).strip().upper())
    return "".join(c for c in t if not unicodedata.combining(c))


def normalizar_grupo(texto) -> str:
    """Texto livre (coluna 'grupo' ou descrição do processo) -> chave do grupo, ou ''."""
    t = _norm(texto)
    if not t:
        return ""
    if "LANCO" in t or "DISTRIBUI" in t:
        return "LANCO"
    if "PLANT" in t or "SEMEA" in t:
        return "PLANTIO"
    if "COLHE" in t:
        return "COLHEITA"
    if "PULVER" in t or "DEFENSIV" in t or "APLICA" in t:
        return "PULVERIZACAO"
    if any(k in t for k in ("PREPARO", "SOLO", "GRADAG", "SUBSOL", "ARACAO", "ESCARIF", "NIVELA", "TERRAPLAN")):
        return "PREPARO_SOLO"
    return ""


def carregar_opcoes(pasta=".") -> dict:
    """{grupo: [chaves de indicador (até 2)]}, só com valores válidos. {} se não houver arquivo."""
    arq = Path(pasta) / ARQUIVO_OPCOES
    if not arq.exists():
        return {}
    try:
        bruto = json.loads(arq.read_text(encoding="utf-8"))
    except Exception as e:  # noqa: BLE001
        print(f"  AVISO: {ARQUIVO_OPCOES} inválido ({e}) — usando mapas automáticos.")
        return {}
    saida = {}
    for grupo, lista in (bruto or {}).items():
        g = normalizar_grupo(grupo) or str(grupo)
        if g not in GRUPOS:
            continue
        validos = [i for i in (lista or []) if i in INDICADORES and i in INDICADORES_POR_GRUPO[g]]
        if validos:
            saida[g] = list(dict.fromkeys(validos))[:2]
    if saida:
        resumo = "; ".join(f"{GRUPOS[g]}: {' + '.join(v)}" for g, v in saida.items())
        print(f"  Mapas por escolha ({ARQUIVO_OPCOES}): {resumo}")
    return saida


ARQUIVO_OPCOES_RELATORIO = "opcoes_relatorio.json"


def carregar_opcoes_relatorio(pasta=".") -> dict:
    """Opções gerais do relatório (ex.: {"comparar_metas": true}). {} se não houver arquivo."""
    arq = Path(pasta) / ARQUIVO_OPCOES_RELATORIO
    if not arq.exists():
        return {}
    try:
        bruto = json.loads(arq.read_text(encoding="utf-8"))
        return bruto if isinstance(bruto, dict) else {}
    except Exception as e:  # noqa: BLE001
        print(f"  AVISO: {ARQUIVO_OPCOES_RELATORIO} inválido ({e}) — ignorado.")
        return {}
