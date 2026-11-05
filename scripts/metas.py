# ============================================================
# METAS.PY — comparação com metas operacionais (opcional)
#
# Portado do pipeline SGPA original (processar_dados.py). Lê o
# metas_operacionais.xlsx da pasta do cliente:
#   aba "Metas da Operação"    : Processo | Indicador | Unidade | Meta | Tipo
#   aba "Metas por Equipamento": Processo | Equipamento | Indicador | Unidade | Meta | Tipo
#                                (ou formato largo: 1 coluna por indicador)
# Regras (iguais às do projeto original):
#   - meta do equipamento tem prioridade sobre a da operação;
#   - área: meta diária = área total planejada / duração; meta do período = meta diária x dias;
#     só meta da operação -> meta por equipamento = meta / nº de equipamentos;
#     metas individuais -> média ponderada pelas horas trabalhadas;
#   - jornada é informativa; RPM: até 50 do alvo = OK, até 100 = ATENÇÃO, acima = FORA DO ALVO.
# Processo e equipamento da planilha são casados pelo CÓDIGO no início do texto
# ("3001 - APLICAÇÃO DEFENSIVOS" -> 3001) ou, se não houver código, pelo nome.
# Se o arquivo não existir ou estiver incompleto, nada quebra: devolve None e o relatório
# sai sem comparação.
# ============================================================

import re
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd

NOME_ARQUIVO = "metas_operacionais.xlsx"
DIAS_PADRAO = 7


# ------------------------------------------------------------
# AUXILIARES
# ------------------------------------------------------------

def norm(valor) -> str:
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
        return ""
    t = unicodedata.normalize("NFKD", str(valor).strip().upper())
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t)


def chave(valor) -> str:
    return re.sub(r"[^A-Z0-9]+", "", norm(valor))


def numero(v):
    try:
        f = float(v)
        return np.nan if np.isnan(f) else f
    except Exception:
        return np.nan


def divisao(a, b):
    a, b = numero(a), numero(b)
    return np.nan if np.isnan(a) or np.isnan(b) or b == 0 else a / b


def media_ponderada(valores, pesos):
    v = pd.Series(pd.to_numeric(valores, errors="coerce"), dtype="float64").reset_index(drop=True)
    p = pd.Series(pd.to_numeric(pesos, errors="coerce"), dtype="float64").reset_index(drop=True)
    m = v.notna() & p.notna() & (p > 0)
    return float((v[m] * p[m]).sum() / p[m].sum()) if m.any() else np.nan


def limpar_numero(valor):
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)) or str(valor).strip() == "":
        return np.nan
    if isinstance(valor, (int, float, np.integer, np.floating)):
        return float(valor)
    s = str(valor).strip().replace("%", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return np.nan


def codigo_inicial(texto):
    m = re.match(r"^\s*(-?\d+)\b", str(texto))
    return str(int(m.group(1))) if m else ""


def mesmo_item(texto_meta, cod, *rotulos):
    """True se o texto da planilha de metas se refere ao item (processo/equipamento) de código `cod`."""
    c = codigo_inicial(texto_meta)
    if c:
        return c == str(int(float(cod))) if re.fullmatch(r"-?\d+(\.0)?", str(cod)) else c == str(cod)
    alvo = norm(texto_meta)
    return bool(alvo) and any(alvo == norm(r) or alvo in norm(r) for r in rotulos if r)


INDICADOR_MAPA = {
    "EFICIENCIA DO MOTOR": "eficiencia_motor", "EFICIENCIA GERAL": "eficiencia_geral",
    "MOTOR LIGADO PARADO": "motor_ligado_parado", "RENDIMENTO REAL": "rendimento_real",
    "CONSUMO": "consumo", "VELOCIDADE MEDIA": "velocidade_media", "VELOCIDADE": "velocidade_media",
    "SEM APONTAMENTO": "sem_apontamento", "AREA POR DIA": "area_por_dia",
    "DURACAO DO PERIODO": "duracao_periodo", "AREA TOTAL PLANEJADA": "area_total_planejada",
    "JORNADA": "jornada", "RPM EFETIVO": "rpm_efetivo", "RPM MEDIO EFETIVO": "rpm_efetivo",
    "RPM TRABALHANDO": "rpm_efetivo", "RPM PARADA": "rpm_parada", "RPM MEDIO PARADA": "rpm_parada",
    "RPM DESLOCAMENTO": "rpm_deslocamento", "RPM MEDIO DESLOCAMENTO": "rpm_deslocamento",
}


def indicador_chave(texto) -> str:
    t = norm(texto)
    for k, v in INDICADOR_MAPA.items():
        if k in t:
            return v
    return chave(texto).lower()


def tipo_padrao(indicador) -> str:
    k = indicador_chave(indicador)
    if k in {"consumo", "motor_ligado_parado", "sem_apontamento"}:
        return "MENOR"
    if k in {"eficiencia_motor", "eficiencia_geral", "rendimento_real", "velocidade_media", "area_por_dia"}:
        return "MAIOR"
    if k in {"rpm_efetivo", "rpm_parada", "rpm_deslocamento"}:
        return "RPM"
    return "INFORMATIVO"


def comparar(realizado, meta, tipo):
    if pd.isna(realizado) or pd.isna(meta):
        return np.nan, "SEM DADO"
    t = norm(tipo)
    if t == "MENOR":
        return realizado - meta, ("OK" if realizado <= meta else "ABAIXO DA META")
    if t == "MAIOR":
        return realizado - meta, ("OK" if realizado >= meta else "ABAIXO DA META")
    if t == "RPM":
        d = abs(realizado - meta)
        return realizado - meta, ("OK" if d <= 50 else "ATENÇÃO" if d <= 100 else "FORA DO ALVO")
    return realizado - meta, "INFORMATIVO"


# ------------------------------------------------------------
# LEITURA DA PLANILHA
# ------------------------------------------------------------

def _aba(excel, candidatos):
    mapa = {chave(x): x for x in excel.sheet_names}
    for c in candidatos:
        if chave(c) in mapa:
            return mapa[chave(c)]
    for k, nome in mapa.items():
        for c in candidatos:
            if chave(c) and (chave(c) in k or k in chave(c)):
                return nome
    return None


def _padronizar(df):
    ren = {}
    for c in df.columns:
        k = chave(c)
        if k in {"PROCESSO", "OPERACAO"}:
            ren[c] = "Processo"
        elif k in {"INDICADOR", "KPI"}:
            ren[c] = "Indicador"
        elif k in {"UNIDADE", "UNID"}:
            ren[c] = "Unidade"
        elif k in {"META", "VALORMETA", "TARGET"}:
            ren[c] = "Meta"
        elif k in {"TIPO", "TIPOMETA"}:
            ren[c] = "Tipo"
        elif k in {"EQUIPAMENTO", "EQUIPAMENTOS"}:
            ren[c] = "Equipamento"
    return df.rename(columns=ren)


# colunas da aba "Metas por Equipamento" (formato largo) -> (indicador, tipo)
_COLUNAS_LARGO = {
    "VELOCIDADEMEDIA": ("Velocidade média", ""), "VELOCIDADE": ("Velocidade média", ""),
    "JORNADATEMPO": ("Jornada", "INFORMATIVO"), "JORNADA": ("Jornada", "INFORMATIVO"),
    "EFMOTOR": ("Eficiência do motor", "MAIOR"), "EFICIENCIADOMOTOR": ("Eficiência do motor", "MAIOR"),
    "RENDREAL": ("Rendimento real", "MAIOR"), "RENDIMENTOREAL": ("Rendimento real", "MAIOR"),
    "AREADIAMETA": ("Área por dia", "MAIOR"), "AREADIA": ("Área por dia", "MAIOR"),
    "RPMEFETIVO": ("RPM efetivo", "RPM"), "RPMPARADA": ("RPM parada", "RPM"),
    "RPMDESLOCAMENTO": ("RPM deslocamento", "RPM"),
}


def _sem_unidade(col):
    return chave(re.sub(r"\([^)]*\)\s*$", "", str(col)))


def _largo_para_longo(df):
    if "Equipamento" not in df.columns:
        return pd.DataFrame()
    regs = []
    for _, r in df.iterrows():
        for col in df.columns:
            alvo = _COLUNAS_LARGO.get(_sem_unidade(col))
            if alvo is None:
                continue
            v = limpar_numero(r.get(col))
            if not pd.isna(v):
                regs.append({"Processo": r.get("Processo", ""), "Equipamento": r.get("Equipamento", ""),
                             "Indicador": alvo[0], "Meta": v, "Unidade": "", "Tipo": alvo[1]})
    return pd.DataFrame(regs)


def _parametros_base_equip(x, a_base):
    """Aba 'Base e Obs. - Equipamento': Ef. Geral, Motor ligado parado e Consumo meta por equipamento."""
    try:
        d = pd.read_excel(x, a_base, header=2)
    except Exception:
        return pd.DataFrame()
    d = d.rename(columns={c: c.strip() for c in d.columns if isinstance(c, str)})
    if "Processo" not in d.columns or "Equipamento" not in d.columns:
        return pd.DataFrame()
    cols = {}
    for c in d.columns:
        k = chave(c)
        if k.startswith("EFGERAL"):
            cols[c] = ("Eficiência geral", "MAIOR")
        elif k.startswith("MOTORLIGADOPARADO"):
            cols[c] = ("Motor ligado parado", "MENOR")
        elif k.startswith("CONSUMOMETA"):
            cols[c] = ("Consumo", "MENOR")
    regs = []
    for _, r in d.iterrows():
        for c, (ind, tp) in cols.items():
            v = limpar_numero(r.get(c))
            if not pd.isna(v) and str(r.get("Processo", "")).strip():
                regs.append({"Processo": r["Processo"], "Equipamento": r["Equipamento"], "Indicador": ind,
                             "Meta": v, "Unidade": "", "Tipo": tp})
    return pd.DataFrame(regs)


def carregar_metas(arquivo, log=print):
    """(metas_operacao, metas_equip) ou None se o arquivo não existe/está vazio."""
    arquivo = Path(arquivo)
    if not arquivo.exists():
        log(f"  AVISO: planilha de metas não encontrada ({arquivo.name}) — relatório sem comparação.")
        return None
    try:
        x = pd.ExcelFile(arquivo)
        a_op = _aba(x, ["Metas da Operação", "Metas da Operacao", "Metas Operação", "Metas"])
        a_eq = _aba(x, ["Metas por Equipamento", "Metas por Equipamentos", "Equipamentos"])
        vazio = pd.DataFrame(columns=["Processo", "Indicador", "Unidade", "Meta", "Tipo"])
        op = _padronizar(pd.read_excel(x, a_op)) if a_op else vazio.copy()
        if op.empty or "Processo" not in op.columns:
            op = vazio.copy()
        eq = _padronizar(pd.read_excel(x, a_eq)) if a_eq else pd.DataFrame()
        for d in (op, eq):
            if not d.empty:
                for c in ("Processo", "Indicador", "Unidade", "Tipo", "Equipamento"):
                    if c in d.columns:
                        d[c] = d[c].fillna("").astype(str).str.strip()
                if "Meta" in d.columns:
                    d["Meta"] = d["Meta"].apply(limpar_numero)
        if not eq.empty and "Indicador" not in eq.columns:
            eq = _largo_para_longo(eq)
            a_base = _aba(x, ["Base e Obs. - Equipamento", "Base e Obs - Equipamento"])
            if a_base:
                eq = pd.concat([eq, _parametros_base_equip(x, a_base)], ignore_index=True)
        for d in (op, eq):
            for c in ("Unidade", "Tipo"):
                if not d.empty and c not in d.columns:
                    d[c] = ""
        if op.empty and (eq is None or eq.empty):
            log(f"  AVISO: {arquivo.name} sem metas utilizáveis — relatório sem comparação.")
            return None
        log(f"  Metas: {arquivo.name} — {len(op)} meta(s) de operação, {0 if eq is None else len(eq)} por equipamento.")
        return op, (eq if eq is not None else pd.DataFrame())
    except Exception as e:  # noqa: BLE001
        log(f"  AVISO: falha ao ler {arquivo.name} ({type(e).__name__}: {e}) — relatório sem comparação.")
        return None


# ------------------------------------------------------------
# BUSCA DE METAS
# ------------------------------------------------------------

def _linhas_processo(df, cod, rotulo, desc):
    if df is None or df.empty or "Processo" not in df.columns:
        return pd.DataFrame()
    m = df["Processo"].map(lambda t: mesmo_item(t, cod, rotulo, desc))
    return df[m]


def meta_operacao(metas_op, cod, rotulo, desc, indicador):
    base = _linhas_processo(metas_op, cod, rotulo, desc)
    k = indicador_chave(indicador)
    for _, r in base.iterrows():
        if indicador_chave(r.get("Indicador", "")) == k:
            return numero(r.get("Meta")), str(r.get("Unidade", "")), str(r.get("Tipo", ""))
    return np.nan, "", ""


def meta_area_periodo(metas_op, cod, rotulo, desc, dias):
    base = _linhas_processo(metas_op, cod, rotulo, desc)
    duracao = area_total = np.nan
    for _, r in base.iterrows():
        k = indicador_chave(r.get("Indicador", ""))
        if k == "duracao_periodo":
            duracao = numero(r.get("Meta"))
        elif k == "area_total_planejada":
            area_total = numero(r.get("Meta"))
    dia = divisao(area_total, duracao)
    return duracao, area_total, (dia * dias if not pd.isna(dia) else np.nan)


def meta_equipamento(metas_eq, cod, rotulo, desc, cod_eq, rotulo_eq, indicador):
    if metas_eq is None or metas_eq.empty or "Equipamento" not in metas_eq.columns:
        return np.nan, "", ""
    base = _linhas_processo(metas_eq, cod, rotulo, desc)
    if base.empty:
        return np.nan, "", ""
    base = base[base["Equipamento"].map(lambda t: mesmo_item(t, cod_eq, rotulo_eq))]
    k = indicador_chave(indicador)
    for _, r in base.iterrows():
        if indicador_chave(r.get("Indicador", "")) == k:
            return numero(r.get("Meta")), str(r.get("Unidade", "")), str(r.get("Tipo", ""))
    return np.nan, "", ""


# ------------------------------------------------------------
# COMPARAÇÃO
# ------------------------------------------------------------

INDICADORES_PROCESSO = [
    ("Eficiência do motor", "eficiencia_motor", "Eficiência do motor (%)"),
    ("Eficiência geral", "eficiencia_geral", "Eficiência geral (%)"),
    ("Motor ligado parado", "motor_ligado_parado", "Motor ligado parado (%)"),
    ("Rendimento real", "rendimento_real", "Rendimento real (ha/h)"),
    ("Consumo", "consumo", "Consumo (L/ha)"),
    ("Velocidade média", "velocidade_media", "Velocidade trabalhando (km/h)"),
    ("Sem apontamento (meta)", "sem_apontamento", "Sem apontamento (%)"),
    ("Área por dia", "area_por_dia", "Área/dia/equip (ha/dia/equip)"),
    ("Jornada", "jornada", "Horas efetivas/dia/equip (h/dia/equip)"),
    ("RPM efetivo", "rpm_efetivo", "RPM efetivo"),
    ("RPM parada", "rpm_parada", "RPM parada"),
    ("RPM deslocamento", "rpm_deslocamento", "RPM deslocamento"),
]

INDICADORES_EQUIP = [  # (indicador da planilha, coluna do realizado)
    ("Área por dia", "Área operacional (ha)"), ("Rendimento real", "Rendimento real (ha/h)"),
    ("Eficiência do motor", "Eficiência do motor (%)"), ("Eficiência geral", "Eficiência geral (%)"),
    ("Consumo", "Consumo (L/ha)"), ("Velocidade média", "Velocidade trabalhando (km/h)"),
    ("Jornada", "Jornada média (h/dia)"), ("RPM efetivo", "RPM efetivo"), ("RPM parada", "RPM parada"),
    ("RPM deslocamento", "RPM deslocamento"), ("Motor ligado parado", "Motor ligado parado (%)"),
    ("Sem apontamento", "Sem apontamento (%)"),
]


def calcular_comparacao(secoes, metas, rotulo_proc, nome_proc, rotulo_equip, dias_analise=None, usar_op=True, usar_eq=True):
    """Devolve (comparacao_processo, comparacao_equipamentos) — DataFrames com Meta/Realizado/Variação/Status."""
    metas_op, metas_eq = metas
    # Cada página usa a sua aba: processo <- "Metas da Operação"; equipamentos <- "Metas por Equipamento".
    metas_op_p = metas_op if usar_op else metas_op.iloc[0:0]
    metas_eq_e = metas_eq if usar_eq else metas_eq.iloc[0:0]
    op, eq = secoes["Operacao"], secoes.get("Equipamentos_Op", pd.DataFrame())
    di = secoes.get("Diario_Op", pd.DataFrame())
    if dias_analise is None:
        if di is not None and not di.empty:
            datas = pd.to_datetime(di["Data"], errors="coerce").dropna()
            dias_analise = (datas.max() - datas.min()).days + 1 if len(datas) else DIAS_PADRAO
        else:
            dias_analise = DIAS_PADRAO

    reg_proc = []
    for _, p in op.iterrows():
        cod = int(p["Código processo"])
        rotulo, desc = rotulo_proc(cod), nome_proc(cod)
        eq_proc = eq[eq["Código processo"] == cod] if not eq.empty else pd.DataFrame()
        qtd_eq = len(eq_proc)
        dias_p = p.get("Dias com dados")
        jornada_real = divisao(p.get("Horas trabalhadas (h)"), numero(dias_p) * numero(p.get("Equipamentos")))

        for nome, k, col_real in INDICADORES_PROCESSO:
            real = jornada_real if k == "jornada" else p.get(col_real, np.nan)
            if k == "area_por_dia":
                duracao, area_total, area_p = meta_area_periodo(metas_op_p, cod, rotulo, desc, dias_analise)
                meta_dia_op = meta_operacao(metas_op_p, cod, rotulo, desc, nome)[0]
                if pd.isna(meta_dia_op):
                    meta_dia_op = divisao(area_total, duracao)
                if pd.isna(area_p) and not pd.isna(meta_dia_op):
                    area_p = meta_dia_op * dias_analise
                meta, origem = (meta_dia_op / qtd_eq, "Operação") if (not pd.isna(meta_dia_op) and qtd_eq > 0) else (np.nan, "Sem meta")
                var, st = comparar(numero(real), meta, "MAIOR")
                reg_proc.append({"Código processo": cod, "Processo": rotulo, "Indicador": nome, "Unidade": "ha/dia/equip",
                                 "Meta": meta, "Meta diária operação": meta_dia_op, "Meta 7 dias": area_p,
                                 "Duração planejamento (dias)": duracao, "Área total planejada": area_total,
                                 "Tipo": "MAIOR", "Origem": origem, "Realizado": real, "Variação": var, "Status": st})
                continue

            meta, un, tipo = meta_operacao(metas_op_p, cod, rotulo, desc, nome)
            origem = "Operação"
            tipo = tipo if norm(tipo) else tipo_padrao(nome)
            var, st = comparar(numero(real), meta, tipo)
            reg_proc.append({"Código processo": cod, "Processo": rotulo, "Indicador": nome, "Unidade": un, "Meta": meta,
                             "Meta diária operação": np.nan, "Meta 7 dias": np.nan, "Duração planejamento (dias)": np.nan,
                             "Área total planejada": np.nan, "Tipo": tipo,
                             "Origem": origem if not pd.isna(meta) else "Sem meta",
                             "Realizado": real, "Variação": var, "Status": st})

    reg_eq = []
    for _, e in eq.iterrows() if not eq.empty else []:
        cod = int(e["Código processo"])
        rotulo, desc = rotulo_proc(cod), nome_proc(cod)
        cod_eq = e["Equipamento (código)"]
        rot_eq = rotulo_equip(cod_eq)
        qtd_eq = int((eq["Código processo"] == cod).sum())
        for indicador, col_real in INDICADORES_EQUIP:
            meta, un, tipo = meta_equipamento(metas_eq_e, cod, rotulo, desc, cod_eq, rot_eq, indicador)
            origem = "Equipamento"
            if indicador_chave(indicador) == "area_por_dia" and not pd.isna(meta):
                dias_e = numero(e.get("Dias com dados"))
                if not pd.isna(dias_e) and dias_e > 0:
                    meta = meta * dias_e  # meta diária -> total do período do equipamento
                un = "ha"
            if pd.isna(meta):
                origem = "Sem meta"
            real = e.get(col_real, np.nan)
            tipo_f = tipo if norm(tipo) else tipo_padrao(indicador)
            var, st = (np.nan, "SEM DADO") if pd.isna(meta) else comparar(numero(real), numero(meta), tipo_f)
            reg_eq.append({"Código processo": cod, "Equipamento (código)": cod_eq, "Indicador": indicador, "Unidade": un,
                           "Meta": meta, "Realizado": real, "Variação": var, "Tipo": tipo_f, "Status": st, "Origem": origem})
    return pd.DataFrame(reg_proc), pd.DataFrame(reg_eq)


# ------------------------------------------------------------
# SIMULAÇÕES COM META (para a IA — calculadas em Python)
# ------------------------------------------------------------

def simulacoes_meta(comparacao, operacao, rotulo_proc):
    """Área potencial se a eficiência geral batesse a meta; litros a mais se o consumo passou da meta."""
    saida = []
    if comparacao is None or comparacao.empty:
        return saida
    for _, p in operacao.iterrows():
        cod = int(p["Código processo"])
        c = comparacao[comparacao["Código processo"] == cod]
        sim = {"processo": rotulo_proc(cod)}
        tem = False
        ef = c[c["Indicador"].map(indicador_chave) == "eficiencia_geral"]
        if not ef.empty:
            meta, real = ef.iloc[0]["Meta"], ef.iloc[0]["Realizado"]
            t, rr = p.get("Tempo total (h)"), p.get("Rendimento real (ha/h)")
            if not any(pd.isna(v) for v in (meta, real, t, rr)) and real < meta:
                area = (meta - real) / 100 * t * rr
                if area > 0:
                    sim["area_potencial_adicional_ha"] = round(float(area), 1)
                    tem = True
        co = c[c["Indicador"].map(indicador_chave) == "consumo"]
        if not co.empty:
            meta, real, area_t = co.iloc[0]["Meta"], co.iloc[0]["Realizado"], p.get("Área operacional (ha)")
            if not any(pd.isna(v) for v in (meta, real, area_t)) and real > meta:
                lit = (real - meta) * area_t
                if lit > 0:
                    sim["litros_consumidos_a_mais"] = round(float(lit), 1)
                    tem = True
        if tem:
            saida.append(sim)
    return saida
