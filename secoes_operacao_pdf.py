# ============================================================
# SECOES_OPERACAO_PDF.PY
#
# Páginas Operação / Paradas / Equipamentos / Operadores / Análise IA
# (layout do relatório SGPA, tema claro) — apenas o REALIZADO, sem metas.
#
# Por processo, 3 páginas:
#   1) Operação   : área realizada, gráfico por dia, indicadores, visão do processo
#   2) Detalhe    : paradas + equipamentos + operadores
#   3) Análise IA : 4 caixas + gargalos/causas-raiz
# Mais uma página inicial de resumo executivo da fazenda.
# ============================================================

import re
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, PageBreak, Image

from gerar_relatorio_qualidade import (
    header, numero, inteiro, percentual,
    STYLE_TITLE, STYLE_SUBTITLE, STYLE_SECTION, STYLE_SMALL, STYLE_BODY, STYLE_CENTER, STYLE_TABLE_HEAD,
    VERDE_HEADER, VERDE, VERDE_BG, BRANCO, FUNDO_CARD, FUNDO_CARD_ALT, FUNDO_TABELA_HEAD, BORDA,
    TEXTO_PRINCIPAL, TEXTO_SECUNDARIO, TEXTO_TERCIARIO, AMARELO,
)

AMARELO_BG = colors.HexColor("#FEF3E2")

# ---- comparação com metas (opcional) ----
COR_STATUS = {"OK": "#0E9F6E", "ABAIXO DA META": "#DC2626", "ATENÇÃO": "#D97706", "FORA DO ALVO": "#DC2626"}
ROTULO_STATUS = {"OK": "OK", "ABAIXO DA META": "ABAIXO", "ATENÇÃO": "ATENÇÃO", "FORA DO ALVO": "FORA DO ALVO"}


def _cor_status(st):
    return COR_STATUS.get(st, "#64748B")


def _seta(var, st):
    """Sem glifos especiais (Helvetica não tem ▲▼): a cor e o sinal já indicam a direção."""
    return ""


def _lookup_proc(comp_proc, cod):
    """{indicador: linha} do processo na comparação."""
    if comp_proc is None or comp_proc.empty:
        return {}
    d = comp_proc[comp_proc["Código processo"] == cod]
    return {r["Indicador"]: r for _, r in d.iterrows()}


def _lookup_equip(comp_eq, cod):
    if comp_eq is None or comp_eq.empty:
        return {}
    d = comp_eq[comp_eq["Código processo"] == cod]
    return {(str(r["Equipamento (código)"]), r["Indicador"]): r for _, r in d.iterrows()}
ABAS = ["Operacao", "Paradas", "Equipamentos_Op", "Operadores_Op", "Diario_Op"]


# ------------------------------------------------------------
# LEITURA DAS ABAS
# ------------------------------------------------------------

def carregar_secoes(arquivo_xlsx):
    """Lê as abas das seções do Excel de resultados. Devolve dict ou None se não existirem."""
    try:
        x = pd.ExcelFile(arquivo_xlsx)
    except Exception:
        return None
    if "Operacao" not in x.sheet_names:
        return None
    d = {a: (pd.read_excel(x, a, dtype={"Equipamento (código)": str}) if a in x.sheet_names else pd.DataFrame()) for a in ABAS}
    return d if not d["Operacao"].empty else None


def _norm(v):
    t = unicodedata.normalize("NFKD", str(v).strip().upper())
    return "".join(c for c in t if not unicodedata.combining(c))


# ------------------------------------------------------------
# COMPONENTES
# ------------------------------------------------------------

def _tabela_titulo(titulo, largura=278 * mm):
    tb = Table([[Paragraph(titulo, STYLE_SECTION)]], colWidths=[largura], rowHeights=[6.5 * mm])
    tb.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), FUNDO_CARD_ALT),
        ("LEFTPADDING", (0, 0), (-1, -1), 3 * mm),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return tb


def pill(texto, width=17, height=5.2, bold=False, font_size=7.8):
    st = ParagraphStyle("pill", parent=STYLE_SMALL, fontName="Helvetica-Bold" if bold else "Helvetica",
                        fontSize=font_size, leading=font_size + 1, textColor=TEXTO_PRINCIPAL, alignment=TA_CENTER)
    tb = Table([[Paragraph(texto, st)]], colWidths=[width * mm], rowHeights=[height * mm])
    tb.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), FUNDO_CARD_ALT),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1.5 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 1.5 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 0.6 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 0.6 * mm),
    ]))
    return tb


def caixa_ia(titulo, texto, largura, altura=None, fundo=FUNDO_CARD, cor_titulo=TEXTO_SECUNDARIO, tamanho=11.5):
    titulo_st = ParagraphStyle("iat", parent=STYLE_SMALL, fontName="Helvetica-Bold", fontSize=9.5, leading=11, textColor=cor_titulo)
    texto_st = ParagraphStyle("iab", parent=STYLE_BODY, fontSize=tamanho, leading=tamanho + 3, textColor=TEXTO_PRINCIPAL)
    conteudo = [Paragraph(str(titulo).upper(), titulo_st), Spacer(1, 1.3 * mm), Paragraph(texto_seguro(texto), texto_st)]
    kw = {"colWidths": [largura]}
    if altura:
        kw["rowHeights"] = [altura]
    tb = Table([[conteudo]], **kw)
    tb.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), fundo), ("BOX", (0, 0), (-1, -1), 0.6, BORDA),
        ("LEFTPADDING", (0, 0), (-1, -1), 4 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 4 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 3 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 3 * mm),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return tb


# ---- texto vindo da IA: só <b>, <font color='#RRGGBB'> ---------------------
_TAG_IA_RE = re.compile(r"<b>|</b>|<font\s+color=(['\"])#([0-9A-Fa-f]{6})\1\s*>|</font>", re.IGNORECASE)


def _escapar(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def sanitizar_texto_ia(texto):
    if not texto:
        return ""
    texto, partes, pilha, pos = str(texto), [], [], 0
    for m in _TAG_IA_RE.finditer(texto):
        partes.append(_escapar(texto[pos:m.start()]))
        tag, low = m.group(0), m.group(0).lower()
        if low == "<b>":
            partes.append("<b>"); pilha.append("b")
        elif low.startswith("<font"):
            partes.append(tag); pilha.append("font")
        elif low == "</b>" and pilha and pilha[-1] == "b":
            partes.append("</b>"); pilha.pop()
        elif low == "</font>" and pilha and pilha[-1] == "font":
            partes.append("</font>"); pilha.pop()
        pos = m.end()
    partes.append(_escapar(texto[pos:]))
    partes.extend(f"</{t}>" for t in reversed(pilha))
    return "".join(partes)


def texto_seguro(texto):
    return sanitizar_texto_ia(str(texto).strip()) if texto else ""


def processo_ia(analise, rotulo):
    if not analise:
        return {}
    alvo = _norm(rotulo)
    itens = analise.get("processos", [])
    for it in itens:
        if _norm(it.get("processo", "")) == alvo:
            return it
    cod = alvo.split(" - ")[0].strip()
    for it in itens:
        if _norm(it.get("processo", "")).split(" - ")[0].strip() == cod:
            return it
    return {}


# ------------------------------------------------------------
# GRÁFICO DIÁRIO (sem meta)
# ------------------------------------------------------------

def criar_grafico(diario_proc, caminho_png, meta_dia=None):
    if diario_proc is None or diario_proc.empty:
        return None, None
    d = diario_proc.copy()
    d["Data"] = pd.to_datetime(d["Data"], errors="coerce")
    d = d[d["Data"].notna()].sort_values("Data")
    if d.empty:
        return None, None
    datas = [x.date() for x in d["Data"]]
    valores = d["Área (ha)"].astype(float).values
    acum = np.cumsum(valores)
    ib, ip = int(np.argmax(valores)), int(np.argmin(valores))
    stats = {"n_dias": len(valores), "ini": datas[0], "fim": datas[-1], "media": float(np.mean(valores)),
             "melhor": (datas[ib], float(valores[ib])), "pior": (datas[ip], float(valores[ip]))}

    plt.close("all")
    fig, ax = plt.subplots(figsize=(8.9, 3.5), dpi=150)
    fig.patch.set_facecolor("#FFFFFF")
    x = np.arange(len(valores))
    ln, = ax.plot(x, acum, color="#0E9F6E", linewidth=2.6, zorder=3, label="Realizado diário (acumulado)")
    ax2 = ax.twinx()
    barras = ax2.bar(x, valores, width=0.58, color="#8FE0BE", alpha=0.9, zorder=1, edgecolor="#0E9F6E", linewidth=0.4,
                     label="Realizado por dia")
    ax2.set_zorder(ax.get_zorder() - 1)
    ax.patch.set_visible(False)
    for e in (ax, ax2):
        e.tick_params(colors="#475569", labelsize=9)
        for s in e.spines.values():
            s.set_color("#CBD5E1")
    ax.grid(axis="y", color="#E5E9F0", linewidth=0.7)
    ax.set_ylabel("ÁREA ACUMULADA (ha)", color="#475569", fontsize=9)
    ax2.set_ylabel("ÁREA POR DIA (ha)", color="#475569", fontsize=9)
    ax2.set_ylim(0, max(valores.max() * 2.2, 1))
    ax.set_ylim(0, max(acum.max() * 1.12, 1))
    ax.set_xticks(x)
    n = len(x)
    passo = 1 if n <= 10 else max(1, round(n / 8))
    ax.set_xticklabels([datas[i].strftime("%d/%m") if (i % passo == 0 or i == n - 1) else "" for i in range(n)],
                       fontsize=6.8, color="#475569")
    ax.set_xlabel("DIA", color="#475569", fontsize=9)
    ax.set_xlim(-0.8, n - 0.2)
    hs, ls = [ln, barras], ["Realizado diário (acumulado)", "Realizado por dia"]
    if meta_dia is not None and not pd.isna(meta_dia) and meta_dia > 0:
        meta_acum = meta_dia * (x + 1)
        lm, = ax.plot(x, meta_acum, color="#DC2626", linewidth=1.8, linestyle="--", zorder=4, label="Meta (acumulada)")
        ax.set_ylim(0, max(acum.max(), meta_acum.max()) * 1.12)
        hs.append(lm); ls.append("Meta (acumulada)")
    leg = ax.legend(hs, ls, loc="upper left", ncol=3 if len(hs) > 2 else 2, frameon=False, fontsize=8.5)
    for t in leg.get_texts():
        t.set_color("#334155")
    fig.tight_layout(pad=1.0)
    fig.savefig(caminho_png, facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)
    return caminho_png, stats


def faixa_resumo_grafico(s, meta_dia=None):
    if not s:
        return None

    com_meta = meta_dia is not None and not pd.isna(meta_dia)
    larg = 30 if com_meta else 37

    def caixa(rot, val):
        fs = 9.6 if com_meta else 11
        conteudo = [Paragraph(rot, ParagraphStyle("cxr", parent=STYLE_SMALL, fontName="Helvetica-Bold", fontSize=7.6 if com_meta else 9, leading=8.6 if com_meta else 10, textColor=TEXTO_SECUNDARIO)),
                    Paragraph(val, ParagraphStyle("cxv", parent=STYLE_BODY, fontName="Helvetica-Bold", fontSize=fs, leading=fs + 1.8))]
        tb = Table([[conteudo]], colWidths=[larg * mm], rowHeights=[15 * mm])
        tb.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), FUNDO_CARD_ALT),
            ("LEFTPADDING", (0, 0), (-1, -1), 2.5 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 2 * mm),
            ("TOPPADDING", (0, 0), (-1, -1), 1.6 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6 * mm),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        return tb

    cel = [
        caixa("PERÍODO", f"{s['ini'].strftime('%d/%m')} a {s['fim'].strftime('%d/%m')}"),
        caixa("REALIZADO DIÁRIO MÉDIO", f"{numero(s['media'], 1)} ha"),
        caixa("MELHOR DIA", f"{s['melhor'][0].strftime('%d/%m')} ({numero(s['melhor'][1], 0)} ha)"),
        caixa("PIOR DIA", f"{s['pior'][0].strftime('%d/%m')} ({numero(s['pior'][1], 0)} ha)"),
    ]
    if com_meta:
        cel.insert(2, caixa("META DIÁRIA MÉDIA", f"{numero(meta_dia, 1)} ha"))
    t = Table([cel], colWidths=[(larg + 0.5) * mm] * len(cel))
    t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0.6 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 0.6 * mm),
                           ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    return t


# ------------------------------------------------------------
# PÁGINA 1 DO PROCESSO — OPERAÇÃO
# ------------------------------------------------------------

KPIS = [
    ("Eficiência geral", "%", "Eficiência geral (%)", "pct"),
    ("Eficiência do motor", "%", "Eficiência do motor (%)", "pct"),
    ("Rendimento real", "ha/h", "Rendimento real (ha/h)", "n2"),
    ("Consumo", "L/ha", "Consumo (L/ha)", "n2"),
    ("Motor ligado parado", "%", "Motor ligado parado (%)", "pct"),
    ("Velocidade média", "km/h", "Velocidade trabalhando (km/h)", "n2"),
    ("Sem apontamento", "%", "Sem apontamento (%)", "pct"),
    ("Área/dia frota", "ha/dia/equip", "Área/dia/equip (ha/dia/equip)", "n1"),
]


def _fmt(v, tipo):
    if tipo == "pct":
        return percentual(v)
    return numero(v, 2 if tipo == "n2" else 1)


def card_area(area, dias, n_equip, meta=None):
    linha = Paragraph(f"<b>{numero(area, 0)}</b> <font size='11'>ha</font>",
                      ParagraphStyle("cv", parent=STYLE_BODY, fontSize=28, leading=31))
    sub = Paragraph(f"{inteiro(dias)} dia(s) com dados · {inteiro(n_equip)} equipamento(s)", STYLE_SMALL)
    extra = []
    if meta is not None:
        esperado = meta.get("Meta 7 dias")
        if esperado is not None and not pd.isna(esperado) and esperado > 0:
            pct = area / esperado * 100
            dif = area - esperado
            st = "OK" if dif >= 0 else "ABAIXO DA META"
            extra = [Spacer(1, 1.5 * mm), Paragraph(
                f"Meta do período: <b>{numero(esperado, 0)} ha</b> · <b>{numero(pct, 0)}%</b> concluído · "
                f"<font color='{_cor_status(st)}'><b>{_seta(dif, st)} {'+' if dif > 0 else ''}{numero(dif, 0)} ha</b></font>",
                ParagraphStyle("cm", parent=STYLE_SMALL, fontSize=9.5, leading=11))]
    conteudo = [Paragraph("ÁREA REALIZADA (TOTAL)", ParagraphStyle("ct", parent=STYLE_SMALL, fontName="Helvetica-Bold", fontSize=10.5, textColor=TEXTO_SECUNDARIO)),
                Spacer(1, 2 * mm), linha, Spacer(1, 2 * mm), sub] + extra
    tb = Table([[conteudo]], colWidths=[150 * mm], rowHeights=[34 * mm if extra else 30 * mm])
    tb.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), FUNDO_CARD), ("BOX", (0, 0), (-1, -1), 0.8, BORDA),
        ("LINEABOVE", (0, 0), (-1, 0), 1.6, VERDE),
        ("LEFTPADDING", (0, 0), (-1, -1), 5 * mm), ("TOPPADDING", (0, 0), (-1, -1), 3 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3 * mm), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return tb


KPI_META = {"Sem apontamento": "Sem apontamento (meta)", "Área/dia frota": "Área por dia"}


def kpi_table(r, lk=None):
    if lk:
        return kpi_table_metas(r, lk)
    head = [Paragraph("INDICADOR", STYLE_TABLE_HEAD), Paragraph("REALIZADO", STYLE_TABLE_HEAD)]
    rows = [head]
    for nome, un, col, tipo in KPIS:
        rows.append([
            Paragraph(f"<b>{nome}</b> <font color='#94A3B8'>({un})</font>", ParagraphStyle("kpin", parent=STYLE_BODY, fontSize=12, leading=14)),
            Paragraph(f"<b>{_fmt(r.get(col), tipo)}</b>", ParagraphStyle("kpiv", parent=STYLE_BODY, fontSize=14, leading=16, alignment=TA_CENTER)),
        ])
    tb = Table(rows, colWidths=[78 * mm, 42 * mm], repeatRows=1)
    st = [("BACKGROUND", (0, 0), (-1, -1), FUNDO_CARD), ("LINEBELOW", (0, 0), (-1, -1), 0.4, BORDA),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 2.8 * mm),
          ("TOPPADDING", (0, 0), (-1, -1), 2.9 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.9 * mm),
          ("ALIGN", (1, 0), (1, -1), "CENTER")]
    st += [("BACKGROUND", (0, i), (-1, i), FUNDO_CARD_ALT) for i in range(2, len(rows), 2)]
    tb.setStyle(TableStyle(st))
    return tb


def kpi_table_metas(r, lk):
    cab = ParagraphStyle("kh", parent=STYLE_TABLE_HEAD, alignment=TA_CENTER)
    rows = [[Paragraph("INDICADOR", STYLE_TABLE_HEAD), Paragraph("META", cab), Paragraph("REALIZADO", cab), Paragraph("VARIAÇÃO", cab)]]
    cores = []
    for i, (nome, un, col, tipo) in enumerate(KPIS, start=1):
        c = lk.get(KPI_META.get(nome, nome))
        real = r.get(col)
        meta_txt, var_txt, cor = "—", "sem meta", "#94A3B8"
        if c is not None and not pd.isna(c["Meta"]):
            meta_txt = _fmt(c["Meta"], tipo)
            st, var = c["Status"], c["Variação"]
            cor = _cor_status(st)
            if st == "INFORMATIVO":
                var_txt = "informativo"
            elif pd.isna(var):
                var_txt = "sem dado"
            else:
                ud = "%" if tipo == "pct" else ""
                var_txt = f"{_seta(var, st)} {'+' if var > 0 else ''}{numero(var, 2 if tipo == 'n2' else 1)}{ud}"
                if st in ROTULO_STATUS and st != "OK":
                    var_txt += f"<br/><font size='6.5'>{ROTULO_STATUS[st]}</font>"
        rows.append([
            Paragraph(f"<b>{nome}</b> <font color='#94A3B8'>({un})</font>", ParagraphStyle("kpin", parent=STYLE_BODY, fontSize=10.5, leading=12)),
            Paragraph(meta_txt, ParagraphStyle("kpm", parent=STYLE_BODY, fontSize=11, leading=13, alignment=TA_CENTER, textColor=TEXTO_SECUNDARIO)),
            Paragraph(f"<b>{_fmt(real, tipo)}</b>", ParagraphStyle("kpiv", parent=STYLE_BODY, fontSize=13, leading=15, alignment=TA_CENTER)),
            Paragraph(f"<font color='{cor}'><b>{var_txt}</b></font>", ParagraphStyle("kpd", parent=STYLE_BODY, fontSize=9.5, leading=10.5, alignment=TA_CENTER)),
        ])
    tb = Table(rows, colWidths=[52 * mm, 20 * mm, 24 * mm, 26 * mm], repeatRows=1)
    st_ = [("BACKGROUND", (0, 0), (-1, -1), FUNDO_CARD), ("LINEBELOW", (0, 0), (-1, -1), 0.4, BORDA),
           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 2.2 * mm),
           ("RIGHTPADDING", (0, 0), (-1, -1), 1.2 * mm),
           ("TOPPADDING", (0, 0), (-1, -1), 2.2 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2 * mm),
           ("ALIGN", (1, 0), (-1, -1), "CENTER")]
    st_ += [("BACKGROUND", (0, i), (-1, i), FUNDO_CARD_ALT) for i in range(2, len(rows), 2)]
    tb.setStyle(TableStyle(st_))
    return tb


def visao_processo(r):
    texto = (f"<b>VISÃO DO PROCESSO</b> · {numero(r.get('Área operacional (ha)'), 1)} ha · "
             f"{numero(r.get('Dias com dados'), 0)} dias · {numero(r.get('Equipamentos'), 0)} equipamento(s) · "
             f"{numero(r.get('Tempo total (h)'), 1)} h totais · {numero(r.get('Horas trabalhadas (h)'), 1)} h trabalhando · "
             f"{numero(r.get('Horas manobra (h)'), 1)} h de manobra · {numero(r.get('Horas deslocamento (h)'), 1)} h de deslocamento · "
             f"{numero(r.get('Motor ligado em parada (h)'), 1)} h de motor ligado em parada")
    tb = Table([[Paragraph(texto, ParagraphStyle("vp", parent=STYLE_SMALL, fontSize=11.5, leading=14, textColor=colors.HexColor("#0A6B45")))]],
               colWidths=[277 * mm])
    tb.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), VERDE_BG), ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#A9D9BF")),
        ("LEFTPADDING", (0, 0), (-1, -1), 3.5 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 3.5 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 2.4 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.4 * mm),
    ]))
    return tb


NOTA = ParagraphStyle("nota", parent=STYLE_SMALL, fontSize=8.5, leading=10.5, textColor=colors.HexColor("#7A8A82"))


def pagina_operacao(cliente, ini, fim, rotulo, r, diario_proc, caminho_png, lk=None):
    story = [header(cliente, f"Período: {ini} a {fim}", f"Processo analisado: {rotulo}"), Spacer(1, 3.5 * mm)]
    lk = lk or {}
    area_row = lk.get("Área por dia")
    meta_dia = area_row.get("Meta diária operação") if area_row is not None else None
    grafico, stats = criar_grafico(diario_proc, caminho_png, meta_dia)
    esq = [[card_area(r.get("Área operacional (ha)"), r.get("Dias com dados"), r.get("Equipamentos"), area_row)], [Spacer(1, 2 * mm)]]
    if grafico:
        esq.append([Image(str(grafico), width=150 * mm, height=60 * mm)])
        faixa = faixa_resumo_grafico(stats, meta_dia)
        if faixa is not None:
            esq += [[Spacer(1, 1.5 * mm)], [faixa]]
    left = Table(esq, colWidths=[155 * mm])
    left.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 1 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 1 * mm),
                              ("TOPPADDING", (0, 0), (-1, -1), 1.2 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.2 * mm),
                              ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    bloco = Table([[left, kpi_table(r, lk or None)]], colWidths=[156 * mm, 122 * mm])
    bloco.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                               ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    story += [bloco, Spacer(1, 2 * mm), visao_processo(r), Spacer(1, 1.5 * mm)]
    story.append(Paragraph(
        "Regra da análise: fim de turno, fim de expediente e troca de turno são desconsiderados dos indicadores "
        "operacionais; permanecem no motor ligado parado. Motor ligado é estimado pelo tempo com RPM &gt; 0. "
        + ("Metas desta página: aba Metas da Operação; as dos equipamentos aparecem na página seguinte (aba Metas por Equipamento). " if lk else "")
        + (
        "Consumo (L/ha) considera todo o litro gasto no processo (inclui manobra, deslocamento e paradas) sobre a área "
        "dos equipamentos que têm dado de consumo."), NOTA))
    return story


# ------------------------------------------------------------
# PÁGINA 2 DO PROCESSO — PARADAS, EQUIPAMENTOS, OPERADORES
# ------------------------------------------------------------

def tabela_paradas(par):
    if par is None or par.empty:
        return Table([[Paragraph("Nenhuma parada operacional registrada.", STYLE_SMALL)]], colWidths=[277 * mm])
    base = par.sort_values("Tempo (h)", ascending=False).head(7)
    head_st = ParagraphStyle("thp", parent=STYLE_TABLE_HEAD, fontSize=7.2, leading=8.2)
    body_st = ParagraphStyle("tbp", parent=STYLE_BODY, fontSize=8, leading=9.2)
    rows = [[Paragraph(t, head_st) for t in ("OPERAÇÃO", "HORAS", "MOTOR LIG.", "% PARADA", "QTD. PARADA", "TEMPO MÉDIO POR PARADA (MIN)")]]
    for _, r in base.iterrows():
        rows.append([Paragraph(str(r["Operação"]), body_st), pill(numero(r["Tempo (h)"], 1)),
                     pill(numero(r["Motor Ligado (h)"], 1)), pill(f"<b>{percentual(r['% do tempo de parada'])}</b>", width=22, bold=True),
                     pill(inteiro(r["Quantidade"])), pill(numero(r["Tempo médio por parada (min)"], 0))])
    tb = Table(rows, colWidths=[95 * mm, 26 * mm, 26 * mm, 28 * mm, 30 * mm, 45 * mm], repeatRows=1)
    tb.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), FUNDO_TABELA_HEAD), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BRANCO, FUNDO_CARD]),
        ("LINEBELOW", (0, 0), (-1, -1), 0.3, BORDA), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.2 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 1.4 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 0.8 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 0.8 * mm),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
    ]))
    return tb


COLUNAS_TABELA = [  # (rótulo, unidade, coluna, tipo)
    ("Área Total", "ha", "Área operacional (ha)", "n1"),
    ("Efic. Motor", "%", "Eficiência do motor (%)", "pct"),
    ("Rendimento Real", "ha/h", "Rendimento real (ha/h)", "n2"),
    ("Consumo", "L/ha", "Consumo (L/ha)", "n2"),
    ("RPM Efetivo", "RPM", "RPM efetivo", "int"),
    ("RPM Parada", "RPM", "RPM parada", "int"),
    ("RPM Deslocamento", "RPM", "RPM deslocamento", "int"),
    ("Motor Ligado Parado", "%", "Motor ligado parado (%)", "pct"),
    ("Veloc. Média", "km/h", "Velocidade trabalhando (km/h)", "n2"),
    ("Jornada", "h/dia", "Jornada média (h/dia)", "n1"),
]


def _v(v, tipo):
    if tipo == "int":
        return inteiro(v)
    return _fmt(v, tipo)


def _tabela_valores(titulo, rotulo_col, linhas, extra_sem_ap, largura_nome=30, lk_eq=None, chave_linha=None):
    cols = COLUNAS_TABELA + ([("Sem Apontamento", "%", "Sem apontamento (%)", "pct")] if extra_sem_ap else [])
    head_st = ParagraphStyle("theq", parent=STYLE_TABLE_HEAD, fontSize=6.0, leading=6.9)
    nome_st = ParagraphStyle("nmeq", parent=STYLE_BODY, fontSize=7.2, leading=8.3)
    val_st = ParagraphStyle("vleq", parent=STYLE_BODY, fontName="Helvetica-Bold", fontSize=7.9, leading=9.1, alignment=TA_CENTER)
    rows = [[Paragraph(rotulo_col, head_st)] + [Paragraph(f"{r} ({u})", head_st) for r, u, _, _ in cols]]
    meta_st = ParagraphStyle("mteq", parent=val_st, fontName="Helvetica", fontSize=6.2, leading=7, textColor=TEXTO_TERCIARIO)
    ind_de_coluna = {"Área Total": "Área por dia", "Efic. Motor": "Eficiência do motor", "Rendimento Real": "Rendimento real",
                     "Consumo": "Consumo", "RPM Efetivo": "RPM efetivo", "RPM Parada": "RPM parada",
                     "RPM Deslocamento": "RPM deslocamento", "Motor Ligado Parado": "Motor ligado parado",
                     "Veloc. Média": "Velocidade média", "Jornada": "Jornada"}
    for nome_fmt, r in linhas:
        cel = []
        for rot, _, c, t in cols:
            ref = lk_eq.get((chave_linha(r), ind_de_coluna.get(rot, ""))) if lk_eq else None
            if ref is not None and not pd.isna(ref["Meta"]):
                st, var = ref["Status"], ref["Variação"]
                cor = _cor_status(st) if st != "INFORMATIVO" else "#334155"
                cel.append([Paragraph(f"<font color='{cor}'>{_v(r.get(c), t)} {_seta(var, st) if st not in ('INFORMATIVO',) else ''}</font>", val_st),
                            Paragraph(f"Meta: {_v(ref['Meta'], t)}", meta_st)])
            else:
                cel.append(Paragraph(_v(r.get(c), t), val_st))
        rows.append([Paragraph(nome_fmt, nome_st)] + cel)
    util = 283 * mm
    larg = (util - largura_nome * mm) / len(cols)
    tb = Table(rows, colWidths=[largura_nome * mm] + [larg] * len(cols), repeatRows=1)
    tb.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), FUNDO_TABELA_HEAD), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BRANCO, FUNDO_CARD]),
        ("LINEBELOW", (0, 0), (-1, -1), 0.3, BORDA), ("GRID", (0, 0), (-1, -1), 0.25, BORDA),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1.2 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 0.8 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 1.0 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.0 * mm),
    ]))
    return [_tabela_titulo(titulo, 283 * mm), tb]


def _nome_equip(rotulo):
    prim, *resto = str(rotulo).split(" - ", 1)
    return f"<b>{prim}</b><br/><font size='6.0'>{resto[0]}</font>" if resto else f"<b>{rotulo}</b>"


def pagina_detalhes(cliente, ini, fim, rotulo, par, eq, op, rotulo_equip, lk_eq=None):
    story = [header(cliente, f"Período: {ini} a {fim}", f"Processo analisado: {rotulo}"), Spacer(1, 2 * mm),
             _tabela_titulo("PRINCIPAIS PARADAS · DETALHAMENTO"), tabela_paradas(par), Spacer(1, 2.5 * mm)]
    if eq is None or eq.empty:
        story.append(Paragraph("Nenhum equipamento identificado.", STYLE_SMALL))
    else:
        linhas = [(_nome_equip(rotulo_equip(r["Equipamento (código)"])), r) for _, r in eq.iterrows()]
        story += _tabela_valores("EQUIPAMENTOS · REALIZADO × META" if lk_eq else "EQUIPAMENTOS · REALIZADO", "EQUIPAMENTO", linhas, False,
                                 lk_eq=lk_eq, chave_linha=lambda r: str(r["Equipamento (código)"]))
    if op is not None and not op.empty:
        story.append(Spacer(1, 2.5 * mm))
        linhas = [(f"<b>{r['Operador']}</b>", r) for _, r in op.iterrows()]
        story += _tabela_valores("OPERADORES · REALIZADO", "OPERADOR", linhas, True)
    story += [Spacer(1, 1.5 * mm),
              Paragraph("Jornada é informativa e não recebe classificação de bom/ruim. Motor ligado estimado por RPM &gt; 0. "
                        + ("Setas: variação contra a meta (verde = dentro, vermelho = fora, laranja = atenção); operadores não têm meta. " if lk_eq else "")
                        + (
                        "Operador \"-1 - * não definido *\" = operação sem operador identificado (só aparece acima de 10 ha)."), NOTA)]
    return story


# ------------------------------------------------------------
# PÁGINA 3 DO PROCESSO — ANÁLISE IA
# ------------------------------------------------------------

def pagina_ia(cliente, ini, fim, rotulo, item, com_metas=False):
    story = [header(cliente, f"Período: {ini} a {fim}", f"Processo analisado: {rotulo}"), Spacer(1, 2.5 * mm)]
    if not item:
        story.append(caixa_ia(
            "Análise IA indisponível",
            "Não há análise de IA para este processo nesta execução (chave não configurada, serviço temporariamente "
            "indisponível ou cota do dia atingida). Os dados operacionais das páginas anteriores não são afetados.",
            278 * mm, altura=30 * mm, fundo=AMARELO_BG, cor_titulo=AMARELO, tamanho=12))
        return story

    def bloco(titulo, campo, vazio):
        return caixa_ia(titulo, item.get(campo) or vazio, 137 * mm, altura=47 * mm, tamanho=9.6)

    def linha(a, b):
        t = Table([[a, b]], colWidths=[139 * mm, 139 * mm])
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                               ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
        return t

    story += [linha(bloco("Operação", "analise_operacao", "Sem leitura disponível para a operação."),
                    bloco("Paradas", "analise_paradas", "Sem leitura disponível para as paradas.")),
              Spacer(1, 2.5 * mm),
              linha(bloco("Equipamentos", "analise_equipamentos", "Sem leitura disponível para os equipamentos."),
                    bloco("Operadores", "analise_operadores", "Sem leitura disponível para os operadores.")),
              Spacer(1, 2.5 * mm),
              caixa_ia("Principais gargalos e possíveis causas-raiz",
                       item.get("gargalos_causas_raiz") or "Sem gargalos relevantes identificados a partir dos dados disponíveis.",
                       278 * mm, altura=38 * mm, fundo=AMARELO_BG, cor_titulo=colors.HexColor("#8A5A00"), tamanho=10.8),
              Spacer(1, 2.5 * mm),
              Paragraph("A análise de IA é interpretativa e usa exclusivamente os dados deste relatório. "
                        + ("As comparações consideram as metas operacionais cadastradas para o cliente (operação e/ou equipamentos), além de equipamentos e operadores."
                           if com_metas else "Não há metas cadastradas; as comparações são entre equipamentos e entre operadores."), NOTA)]
    return story


# ------------------------------------------------------------
# PÁGINA INICIAL — RESUMO EXECUTIVO DA FAZENDA
# ------------------------------------------------------------

def pagina_resumo_executivo(cliente, ini, fim, secoes, analise, rotulo_proc, comparacao=None):
    op, eq = secoes["Operacao"], secoes["Equipamentos_Op"]
    story = [header(cliente, f"Período: {ini} a {fim}", "RESUMO EXECUTIVO DA OPERAÇÃO"), Spacer(1, 4 * mm)]
    geral = (analise or {}).get("resumo_geral") or ("Resumo dos processos realizados no período, comparados com as metas operacionais." if comparacao else "Resumo dos processos realizados no período, com dados realizados (sem metas).")
    story += [caixa_ia("Leitura executiva da fazenda", geral, 278 * mm, altura=27 * mm, tamanho=12), Spacer(1, 3 * mm)]

    n_eq = eq["Equipamento (código)"].nunique() if not eq.empty else op["Equipamentos"].sum()
    cards = [caixa_ia("PROCESSOS ANALISADOS", inteiro(len(op)), 66 * mm, altura=22 * mm, tamanho=15),
             caixa_ia("ÁREA TOTAL", f"{numero(op['Área operacional (ha)'].sum(), 1)} ha", 66 * mm, altura=22 * mm, tamanho=15),
             caixa_ia("EQUIPAMENTOS", inteiro(n_eq), 66 * mm, altura=22 * mm, tamanho=15),
             caixa_ia("HORAS TOTAIS", f"{numero(op['Tempo total (h)'].sum(), 1)} h", 66 * mm, altura=22 * mm, tamanho=15)]
    ct = Table([cards], colWidths=[69 * mm] * 4)
    ct.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 1 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 1 * mm),
                            ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                            ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story += [ct, Spacer(1, 3 * mm)]

    rows = [[Paragraph(t, STYLE_TABLE_HEAD) for t in ("PROCESSO", "ÁREA", "EQUIP.", "HORAS", "RESUMO")]]
    for _, r in op.iterrows():
        rot = rotulo_proc(int(r["Código processo"]))
        leitura = texto_seguro(processo_ia(analise, rot).get("resumo_executivo", ""))
        rows.append([Paragraph(f"<b>{rot}</b>", STYLE_BODY), Paragraph(f"{numero(r['Área operacional (ha)'], 1)} ha", STYLE_BODY),
                     Paragraph(inteiro(r["Equipamentos"]), STYLE_CENTER), Paragraph(f"{numero(r['Tempo total (h)'], 1)} h", STYLE_CENTER),
                     Paragraph(leitura or "Sem análise de IA disponível.", ParagraphStyle("fgrow", parent=STYLE_BODY, fontSize=9.5, leading=11))])
    tb = Table(rows, colWidths=[65 * mm, 27 * mm, 22 * mm, 24 * mm, 140 * mm], repeatRows=1)
    tb.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), FUNDO_TABELA_HEAD), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BRANCO, FUNDO_CARD]),
        ("LINEBELOW", (0, 0), (-1, -1), 0.35, BORDA), ("GRID", (0, 0), (-1, -1), 0.25, BORDA),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 2.5 * mm),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2 * mm), ("TOPPADDING", (0, 0), (-1, -1), 1.3 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.3 * mm), ("ALIGN", (1, 1), (3, -1), "CENTER"),
    ]))
    barra = Table([[Paragraph("VISÃO DOS PROCESSOS", STYLE_SECTION)], [Paragraph("Principais resultados e leitura executiva gerada pela IA.", STYLE_SMALL)]],
                  colWidths=[278 * mm])
    barra.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), FUNDO_CARD_ALT), ("LEFTPADDING", (0, 0), (-1, -1), 3.5 * mm),
                               ("TOPPADDING", (0, 0), (-1, -1), 2 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 2 * mm)]))
    story += [barra, Spacer(1, 1.5 * mm), tb]
    return story


# ------------------------------------------------------------
# API para o gerador do PDF
# ------------------------------------------------------------

def paginas_processo(cod, cliente, ini, fim, secoes, analise, rotulo_proc, rotulo_equip, pasta_tmp, usar_ia=True, comparacao=None):
    """3 páginas (ou 2 sem IA) do processo, SEM PageBreak final."""
    op = secoes["Operacao"]
    linha = op[op["Código processo"] == cod]
    if linha.empty:
        return []
    r = linha.iloc[0]
    rot = rotulo_proc(int(cod))
    pasta_tmp = Path(pasta_tmp)
    pasta_tmp.mkdir(parents=True, exist_ok=True)
    di = secoes["Diario_Op"]
    comp_proc, comp_eq = comparacao if comparacao else (None, None)
    story = pagina_operacao(cliente, ini, fim, rot, r, di[di["Código processo"] == cod] if not di.empty else di,
                            pasta_tmp / f"grafico_{int(cod)}.png", _lookup_proc(comp_proc, cod))
    story.append(PageBreak())
    par, eq, opr = secoes["Paradas"], secoes["Equipamentos_Op"], secoes["Operadores_Op"]
    story += pagina_detalhes(cliente, ini, fim, rot,
                             par[par["Código processo"] == cod] if not par.empty else par,
                             eq[eq["Código processo"] == cod] if not eq.empty else eq,
                             opr[opr["Código processo"] == cod] if not opr.empty else opr, rotulo_equip,
                             _lookup_equip(comp_eq, cod) or None)
    if usar_ia:
        story.append(PageBreak())
        story += pagina_ia(cliente, ini, fim, rot, processo_ia(analise, rot), bool(comparacao))
    return story
