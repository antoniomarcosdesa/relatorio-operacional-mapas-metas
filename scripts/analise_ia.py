# ================================================================
# ANALISE_IA.PY — análise executiva por processo (Gemini), SEM metas.
#
# Portado de gerar_analise_ia.py (pipeline SGPA). Mantidos: prompt com as
# 5 seções (Operação / Paradas / Equipamentos / Operadores / Gargalos e
# causas-raiz), formatação rica (<b> e <font color>) e as simulações
# calculadas em Python (a IA só narra o número pronto):
#   - % de manobra + deslocamento sobre o motor ligado;
#   - área a mais se o deslocamento fosse 20% menor;
#   - área a mais se cada equipamento tivesse a jornada do líder.
# Removidas: tudo que dependia de metas (comparações e simulações de
# eficiência/consumo versus meta).
#
# A chave NUNCA fica no código: vem da variável de ambiente
# GEMINI_API_KEY (ou do arquivo gemini.env na raiz do projeto).
# Se faltar chave ou a chamada falhar, devolve None e o PDF sai normalmente,
# com a página de IA avisando que a análise não está disponível.
# ================================================================

import json
import os
import time
from pathlib import Path

import numpy as np
import pandas as pd

MODELO_GEMINI = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite")
MAX_TENTATIVAS = 3
ESPERA_ENTRE_TENTATIVAS = 5


def _ler_chave():
    chave = os.getenv("GEMINI_API_KEY", "").strip()
    if chave:
        return chave
    for pasta in (Path.cwd(), Path(__file__).resolve().parent.parent):
        arq = pasta / "gemini.env"
        if arq.exists():
            for linha in arq.read_text(encoding="utf-8", errors="ignore").splitlines():
                linha = linha.strip()
                if linha.startswith("GEMINI_API_KEY"):
                    return linha.split("=", 1)[1].strip().strip("'\"")
    return ""


# ------------------------------------------------------------
# DADOS PARA A IA
# ------------------------------------------------------------

def _registros(df, colunas=None, casas=2):
    if df is None or df.empty:
        return []
    saida = []
    for _, linha in df.iterrows():
        reg = {}
        for c in (colunas or df.columns):
            v = linha[c]
            if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))) or (not isinstance(v, str) and pd.isna(v)):
                continue
            reg[str(c)] = round(float(v), casas) if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool) else str(v)
        if reg:
            saida.append(reg)
    return saida


COLS_OPERACAO = [
    "Área operacional (ha)", "Dias com dados", "Equipamentos", "Tempo total (h)", "Horas trabalhadas (h)",
    "Motor ligado (h)", "Motor ligado em parada (h)", "Motor ligado parado (%)", "Consumo (L/ha)",
    "Eficiência geral (%)", "Eficiência do motor (%)", "Rendimento real (ha/h)",
    "Área/dia/equip (ha/dia/equip)", "Velocidade trabalhando (km/h)", "RPM efetivo", "RPM parada",
    "RPM deslocamento", "Sem apontamento (%)", "Horas manobra (h)", "Horas deslocamento (h)",
]
COLS_EQUIP_OPER = [
    "Área operacional (ha)", "Dias com dados", "Horas trabalhadas (h)", "Eficiência geral (%)",
    "Eficiência do motor (%)", "Rendimento real (ha/h)", "Consumo (L/ha)", "RPM efetivo", "RPM parada",
    "RPM deslocamento", "Motor ligado parado (%)", "Velocidade trabalhando (km/h)", "Jornada média (h/dia)",
    "Sem apontamento (%)",
]


def calcular_simulacoes(operacao: pd.DataFrame, equipamentos: pd.DataFrame, rotulo_proc, rotulo_equip):
    """Simulações sem meta, em Python puro. Devolve (por_processo, jornada_equipamentos)."""
    por_proc = []
    for _, r in operacao.iterrows():
        h_man, h_des, motor = r.get("Horas manobra (h)"), r.get("Horas deslocamento (h)"), r.get("Motor ligado (h)")
        rend = r.get("Rendimento real (ha/h)")
        if pd.isna(h_man) or pd.isna(h_des) or pd.isna(motor) or not motor or motor <= 0:
            continue
        pct = (h_man + h_des) / motor * 100
        sim = {"processo": rotulo_proc(int(r["Código processo"])),
               "pct_manobra_deslocamento_sobre_motor_ligado": round(float(pct), 1)}
        if pct > 10 and not pd.isna(rend) and h_des > 0:
            area = h_des * 0.20 * rend
            if area > 0:
                sim["area_potencial_reducao_20pct_deslocamento_ha"] = round(float(area), 1)
        por_proc.append(sim)

    jornada = []
    if equipamentos is not None and not equipamentos.empty:
        for cod, g in equipamentos.groupby("Código processo"):
            v = g[g["Jornada média (h/dia)"].notna()]
            if len(v) < 2:
                continue
            lider = v.loc[v["Jornada média (h/dia)"].idxmax()]
            for _, e in v.iterrows():
                if e["Equipamento (código)"] == lider["Equipamento (código)"]:
                    continue
                dif = lider["Jornada média (h/dia)"] - e["Jornada média (h/dia)"]
                ef, rend, dias = e.get("Eficiência geral (%)"), e.get("Rendimento real (ha/h)"), e.get("Dias com dados")
                if dif <= 0 or pd.isna(ef) or pd.isna(rend) or pd.isna(dias):
                    continue
                area = dif * (ef / 100) * rend * dias
                if area > 0:
                    jornada.append({
                        "processo": rotulo_proc(cod),
                        "equipamento": rotulo_equip(e["Equipamento (código)"]),
                        "jornada_equipamento_h_dia": round(float(e["Jornada média (h/dia)"]), 1),
                        "equipamento_maior_jornada": rotulo_equip(lider["Equipamento (código)"]),
                        "jornada_maior_h_dia": round(float(lider["Jornada média (h/dia)"]), 1),
                        "area_potencial_adicional_ha": round(float(area), 1),
                    })
    return por_proc, jornada


def preparar_dados(secoes, cliente, rotulo_proc, rotulo_equip):
    op, par = secoes["Operacao"], secoes["Paradas"]
    eq, opr = secoes["Equipamentos_Op"], secoes["Operadores_Op"]

    def com_rotulos(df, col_extra=None):
        d = df.copy()
        d.insert(0, "Processo", d["Código processo"].map(rotulo_proc))
        d = d.drop(columns=["Código processo"])
        return d

    eq2 = com_rotulos(eq) if not eq.empty else eq
    if not eq2.empty:
        eq2.insert(1, "Equipamento", eq2["Equipamento (código)"].map(rotulo_equip))
        eq2 = eq2.drop(columns=["Equipamento (código)"])
    por_proc, jornada = calcular_simulacoes(op, eq, rotulo_proc, rotulo_equip)
    return {
        "cliente": cliente,
        "observacao": "Motor ligado é ESTIMADO (tempo com RPM > 0). Não há metas cadastradas.",
        "resumo_processos": _registros(com_rotulos(op), ["Processo"] + COLS_OPERACAO),
        "paradas": _registros(com_rotulos(par).head(400)) if not par.empty else [],
        "equipamentos": _registros(eq2, ["Processo", "Equipamento"] + COLS_EQUIP_OPER) if not eq2.empty else [],
        "operadores": _registros(com_rotulos(opr), ["Processo", "Operador"] + COLS_EQUIP_OPER) if not opr.empty else [],
        "simulacoes_operacao": por_proc,
        "simulacoes_jornada_equipamentos": jornada,
    }


# ------------------------------------------------------------
# PROMPT
# ------------------------------------------------------------

def criar_prompt(dados):
    dados_json = json.dumps(dados, ensure_ascii=False, indent=1)
    return f"""
Você é um ANALISTA OPERACIONAL SÊNIOR especializado em operações agrícolas mecanizadas.
Sua função é interpretar os indicadores operacionais já calculados (você NÃO recalcula os
indicadores principais): interpretar, identificar desvios e oportunidades, destacar bons
resultados, priorizar pontos de atenção e recomendar ações práticas.

CLIENTE: {dados.get('cliente', '')}

============================================================
REGRA FUNDAMENTAL: NÃO HÁ METAS
============================================================
Nenhuma meta está cadastrada. NÃO compare com metas, NÃO invente metas, NÃO cite "meta".
A leitura é sobre o desempenho REALIZADO e a comparação RELATIVA entre equipamentos e entre
operadores do mesmo processo. "Motor ligado" é estimado (tempo com RPM > 0); não trate isso
como erro, apenas não dê precisão exagerada.

============================================================
DIREÇÃO DOS INDICADORES
============================================================
Rendimento real, eficiência geral, eficiência do motor, área: MAIOR = melhor.
Consumo (L/ha), motor ligado parado, sem apontamento: MENOR = melhor.
RPM: sem meta, use apenas para comparar equipamentos (ex.: RPM de deslocamento muito acima
dos demais) — não classifique como bom/ruim sozinho.
Jornada: INFORMATIVA, não classifique como boa ou ruim.

============================================================
ESTRUTURA (uma página por processo)
============================================================
Para CADA processo produza um texto curto e direto para CADA seção, NESTA ORDEM:

1. OPERAÇÃO — leitura dos indicadores gerais (rendimento, eficiência, consumo, motor ligado
   parado, deslocamento). Em "simulacoes_operacao" há valores PRÉ-CALCULADOS:
   - "pct_manobra_deslocamento_sobre_motor_ligado": % do motor ligado gasto em manobra +
     deslocamento. Comente quando for relevante (acima de 10%).
   - "area_potencial_reducao_20pct_deslocamento_ha": só existe se esse % passa de 10%; é a
     área a mais se o deslocamento fosse 20% menor (ex.: "reduzindo o deslocamento em 20%, a
     operação poderia ganhar cerca de X ha").
   Não recalcule; use o valor fornecido. Se não houver registro, não mencione a simulação.

2. PARADAS — principais motivos de parada (horas, % do tempo parado, frequência). Se o
   processo tem mais de um equipamento, use "simulacoes_jornada_equipamentos": cite o
   equipamento de maior jornada e, para os demais listados, a área a mais que fariam se
   tivessem a mesma jornada (valor pré-calculado, não recalcule).

3. EQUIPAMENTOS — comparação entre equipamentos (se houver mais de um): MELHOR e PIOR em
   ÁREA, CONSUMO e EFICIÊNCIA; comente também consumo, RPM efetivo, RPM de deslocamento ou
   motor ligado parado quando algum equipamento destoar dos demais.

4. OPERADORES — comparação entre operadores (se houver mais de um): MELHOR e PIOR em ÁREA,
   CONSUMO e EFICIÊNCIA. Sempre que algum operador tiver SEM APONTAMENTO acima de 10%,
   destaque como ponto negativo (apontamento incompleto). O operador "-1 - * não definido *"
   significa operação sem operador identificado.

5. GARGALOS E CAUSA RAIZ — síntese dos principais gargalos e possíveis causas-raiz, baseada
   nos 4 itens acima (não é uma lista nova de problemas).

REGRAS DE CADA SEÇÃO:
- De 2 a 4 frases curtas por seção. Não repita tabelas; cite só os números necessários.
- Evite diagnósticos mecânicos definitivos; use "avaliar", "verificar", "investigar",
  "acompanhar", "orientar". Prefira ações concretas a frases genéricas.
- Sem dados suficientes (ex.: um único equipamento ou operador), diga isso em 1 frase curta.
- Não crie problemas artificiais: se estiver bem, diga objetivamente.
- Gere também "resumo_executivo" (no máximo 2 frases: principal resultado e principal
  desafio) e "resumo_geral" (2 a 3 frases sobre a fazenda toda).
- O campo "processo" da resposta deve repetir EXATAMENTE o nome do processo recebido.

============================================================
FORMATAÇÃO DO TEXTO (negrito e cor)
============================================================
O texto vira PDF e só entende estas tags: <b>...</b>, <font color='#DC2626'>...</font> e
<font color='#0A8F5A'>...</font>. Nada de markdown (**, *, #, -).
- Envolva em <b> todo nome de EQUIPAMENTO ou OPERADOR e todo VALOR/NÚMERO citado.
- Valor RUIM (pior do grupo; motor ligado parado alto; sem apontamento alto): número em
  <font color='#DC2626'> dentro do <b>.
- Valor BOM (melhor do grupo): <font color='#0A8F5A'> dentro do <b>.
- Exemplo: "<b><font color='#DC2626'>JOSE REINALDO</font></b> teve motor ligado parado de
  <b><font color='#DC2626'>35,9%</font></b>, bem acima dos demais."
- Use números no formato brasileiro (vírgula decimal).

A saída ocupa UMA página A4 horizontal; cada seção precisa ser curta.

============================================================
DADOS DISPONÍVEIS
============================================================
{dados_json}
"""


SCHEMA_RESPOSTA = {
    "type": "object",
    "properties": {
        "cliente": {"type": "string"},
        "resumo_geral": {"type": "string"},
        "processos": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "processo": {"type": "string"},
                    "resumo_executivo": {"type": "string"},
                    "analise_operacao": {"type": "string"},
                    "analise_paradas": {"type": "string"},
                    "analise_equipamentos": {"type": "string"},
                    "analise_operadores": {"type": "string"},
                    "gargalos_causas_raiz": {"type": "string"},
                },
                "required": ["processo", "resumo_executivo", "analise_operacao", "analise_paradas",
                             "analise_equipamentos", "analise_operadores", "gargalos_causas_raiz"],
            },
        },
    },
    "required": ["cliente", "resumo_geral", "processos"],
}

CAMPOS = ("resumo_executivo", "analise_operacao", "analise_paradas", "analise_equipamentos",
          "analise_operadores", "gargalos_causas_raiz")


def _consultar(prompt, chave, log):
    from google import genai
    from google.genai import types
    client = genai.Client(api_key=chave)
    ultimo = None
    for t in range(1, MAX_TENTATIVAS + 1):
        log(f"  IA: enviando ao Gemini ({MODELO_GEMINI}), tentativa {t}/{MAX_TENTATIVAS}...")
        try:
            resp = client.models.generate_content(
                model=MODELO_GEMINI, contents=prompt,
                config=types.GenerateContentConfig(response_mime_type="application/json",
                                                   response_schema=SCHEMA_RESPOSTA),
            )
            if not resp.text:
                raise RuntimeError("resposta vazia")
            return json.loads(resp.text)
        except Exception as e:  # noqa: BLE001
            ultimo = e
            log(f"  IA: erro — {str(e)[:300]}")
            if t < MAX_TENTATIVAS:
                time.sleep(ESPERA_ENTRE_TENTATIVAS)
    raise RuntimeError(f"falha após {MAX_TENTATIVAS} tentativas: {str(ultimo)[:300]}")


def gerar_analise(secoes, cliente, rotulo_proc, rotulo_equip, pasta_saida, log=print):
    """Devolve o dict da análise (e grava resultados/analises_ia.json) ou None se indisponível."""
    chave = _ler_chave()
    if not chave:
        log("  IA: GEMINI_API_KEY não configurada — o PDF sairá sem a análise de IA.")
        return None
    try:
        dados = preparar_dados(secoes, cliente, rotulo_proc, rotulo_equip)
        resultado = _consultar(criar_prompt(dados), chave, log)
        for p in resultado.get("processos", []):
            for c in CAMPOS:
                p[c] = p.get(c) or ""
        pasta = Path(pasta_saida)
        pasta.mkdir(parents=True, exist_ok=True)
        (pasta / "analises_ia.json").write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
        log(f"  IA: análise gerada para {len(resultado.get('processos', []))} processo(s).")
        return resultado
    except Exception as e:  # noqa: BLE001
        log(f"  IA: indisponível ({e}). O PDF sairá sem a análise de IA.")
        return None
