# ============================================================
# GERAR_RELATORIO_QUALIDADE.PY
#
# PDF de visualização do novo pipeline (mapa analítico detalhado ->
# qualidade_operacional_processo_talhao.xlsx -> PDF).
#
# Mesmo tema visual do gerar_relatorio.py do projeto SGPA original
# (fundo claro, barra verde no cabeçalho, badges/pills de status),
# mas sem comparação com meta (esse arquivo não tem metas cadastradas
# ainda) — é um relatório DESCRITIVO de qualidade operacional por
# Processo x Talhão, para você validar o modelo de saída em PDF.
# ============================================================

from pathlib import Path
from datetime import datetime
import sys
import numpy as np
import pandas as pd
from PIL import Image as PILImage

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
)

from gerar_mapas_velocidade_talhao import gerar_mapas_velocidade, ColunaGeometriaAusente
from dados_cliente import carregar_cliente, _chave as _chave_dc, PASTA_CLIENTE as _PASTA_CLIENTE

ARQUIVO_ENTRADA = Path("resultados/qualidade_operacional_processo_talhao.xlsx")
PASTA_RELATORIOS = Path("relatorios")
PASTA_MAPAS = PASTA_RELATORIOS / "mapas_velocidade"
CLIENTE_NOME = "FAZENDA"  # título do PDF: só "FAZENDA" (sem o nome do cliente)
PASTA_RELATORIOS.mkdir(exist_ok=True)

# Imagens por página do slide de mapas: até 2 talhões -> 1 coluna, cada
# imagem maior; 3+ talhões -> grade 2 colunas (2-4 imagens por página),
# conforme pedido ("2 a 4 imagens por slide, dependendo da qualidade").
MAX_IMAGENS_GRADE = 4

# ------------------------------------------------------------
# PALETA (igual ao gerar_relatorio.py original)
# ------------------------------------------------------------

VERDE_HEADER = colors.HexColor("#075A3D")
VERDE = colors.HexColor("#0E9F6E")
VERDE_ESCURA = colors.HexColor("#0A8F5A")
VERDE_BG = colors.HexColor("#E6F7EF")

BRANCO = colors.white
FUNDO_CARD = colors.HexColor("#F8FAFC")
FUNDO_CARD_ALT = colors.HexColor("#EEF2F6")
FUNDO_TABELA_HEAD = colors.HexColor("#E9EEF3")
BORDA = colors.HexColor("#DCE3EA")

TEXTO_PRINCIPAL = colors.HexColor("#0F172A")
TEXTO_SECUNDARIO = colors.HexColor("#64748B")
TEXTO_TERCIARIO = colors.HexColor("#94A3B8")

VERMELHO = colors.HexColor("#DC2626")
AMARELO = colors.HexColor("#B45309")

# ------------------------------------------------------------
# FORMATACAO
# ------------------------------------------------------------

def numero(v, casas=1):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    try:
        return f"{float(v):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)


def inteiro(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    try:
        return f"{int(round(float(v))):,}".replace(",", ".")
    except Exception:
        return str(v)


def percentual(v, casas=1):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    return f"{float(v):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".") + "%"


# ------------------------------------------------------------
# ESTILOS
# ------------------------------------------------------------

styles = getSampleStyleSheet()
STYLE_TITLE = ParagraphStyle("title", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=22, leading=24, textColor=BRANCO)
STYLE_SUBTITLE = ParagraphStyle("subtitle", parent=styles["Normal"], fontName="Helvetica", fontSize=12, leading=14, textColor=colors.HexColor("#D7E8DF"))
STYLE_SECTION = ParagraphStyle("section", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=13.5, leading=15.5, textColor=TEXTO_PRINCIPAL)
STYLE_SMALL = ParagraphStyle("small", parent=styles["Normal"], fontName="Helvetica", fontSize=10.5, leading=12.5, textColor=TEXTO_SECUNDARIO)
STYLE_BODY = ParagraphStyle("body", parent=styles["Normal"], fontName="Helvetica", fontSize=12, leading=14, textColor=TEXTO_PRINCIPAL)
STYLE_CENTER = ParagraphStyle("center", parent=STYLE_BODY, alignment=TA_CENTER)
STYLE_TABLE_HEAD = ParagraphStyle("th", parent=STYLE_SMALL, fontName="Helvetica-Bold", fontSize=9.3, leading=10.6, textColor=TEXTO_SECUNDARIO)

# ------------------------------------------------------------
# NUMERACAO AUTOMATICA DE PAGINA (2 passadas — igual ao original)
# ------------------------------------------------------------

class NumberedCanvas(Canvas):
    def __init__(self, *args, **kwargs):
        Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            largura, _ = landscape(A4)
            self.setFont("Helvetica-Bold", 8.5)
            self.setFillColor(TEXTO_SECUNDARIO)
            self.drawRightString(largura - 7 * mm, 2.6 * mm, f"Pag. {self._pageNumber}/{total}")
            Canvas.showPage(self)
        Canvas.save(self)


# ------------------------------------------------------------
# COMPONENTES
# ------------------------------------------------------------

def header(titulo, subtitulo1, subtitulo2=""):
    linhas = [[Paragraph(f"<b>{titulo}</b>", STYLE_TITLE)], [Paragraph(subtitulo1, STYLE_SUBTITLE)]]
    alturas = [8.5 * mm, 6 * mm]
    if subtitulo2:
        linhas.append([Paragraph(subtitulo2, ParagraphStyle("s2", parent=STYLE_SUBTITLE, fontSize=11))])
        alturas.append(6 * mm)
    box = Table(linhas, colWidths=[172 * mm], rowHeights=alturas)
    box.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), VERDE_HEADER),
        ("LEFTPADDING", (0, 0), (-1, -1), 4 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 0.8 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0.3 * mm),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    t = Table([[box, ""]], colWidths=[172 * mm, 106 * mm], rowHeights=[sum(alturas)])
    t.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return t


def barra_titulo(titulo, subtitulo=""):
    conteudo = [[Paragraph(titulo, STYLE_SECTION)]]
    if subtitulo:
        conteudo.append([Paragraph(subtitulo, STYLE_SMALL)])
    tb = Table(conteudo, colWidths=[278 * mm])
    tb.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), FUNDO_CARD_ALT),
        ("LEFTPADDING", (0, 0), (-1, -1), 3.5 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 2.0 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.0 * mm),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return tb


def caixa(titulo, valor, largura, altura=22 * mm, tamanho=15):
    conteudo = [
        Paragraph(str(titulo).upper(), ParagraphStyle("cxt", parent=STYLE_SMALL, fontName="Helvetica-Bold", fontSize=9, textColor=TEXTO_SECUNDARIO)),
        Spacer(1, 1.3 * mm),
        Paragraph(str(valor), ParagraphStyle("cxv", parent=STYLE_BODY, fontName="Helvetica-Bold", fontSize=tamanho, textColor=TEXTO_PRINCIPAL)),
    ]
    tb = Table([[conteudo]], colWidths=[largura], rowHeights=[altura])
    tb.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), FUNDO_CARD),
        ("BOX", (0, 0), (-1, -1), 0.6, BORDA),
        ("LEFTPADDING", (0, 0), (-1, -1), 4 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 4 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 3 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 3 * mm),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return tb


def tabela_generica(titulo, subtitulo, cabecalhos, larguras, linhas):
    """Tabela compacta reaproveitável: cabecalhos = lista de strings,
    linhas = lista de listas de Paragraph/str já formatados."""
    elementos = [barra_titulo(titulo, subtitulo), Spacer(1, 1.2 * mm)]

    head = [Paragraph(c, STYLE_TABLE_HEAD) for c in cabecalhos]
    corpo = [head]
    for linha in linhas:
        corpo.append([
            v if hasattr(v, "getPlainText") else Paragraph(str(v), ParagraphStyle(
                "cell", parent=STYLE_BODY, fontSize=9.6, leading=11, alignment=TA_CENTER))
            for v in linha
        ])

    tb = Table(corpo, colWidths=larguras, repeatRows=1)
    tb.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), FUNDO_TABELA_HEAD),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BRANCO, FUNDO_CARD]),
        ("GRID", (0, 0), (-1, -1), 0.25, BORDA),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 2 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 1.8 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.8 * mm),
    ]))
    elementos.append(tb)
    return elementos


def celula_talhao(nome):
    return Paragraph(f"<b>{nome}</b>", ParagraphStyle("tal", parent=STYLE_BODY, fontSize=10.5, alignment=TA_CENTER))


# ------------------------------------------------------------
# PÁGINA 1 — RESUMO GERAL
# ------------------------------------------------------------

def pagina_resumo_geral(cliente, ini, fim, resumo_processo, resumo_talhao, dados_cliente=None):
    story = []
    story.append(header(
        cliente,
        f"Período: {ini} a {fim}",
        "QUALIDADE OPERACIONAL — RESUMO EXECUTIVO (Processo x Talhão)",
    ))
    story.append(Spacer(1, 4 * mm))

    total_area = resumo_processo["Área operacional (ha)"].sum()
    total_talhoes = resumo_talhao["Talhão"].nunique()
    total_equip = resumo_talhao["Equipamentos (códigos)"].astype(str).str.split(", ").explode().nunique()
    total_horas = resumo_processo["Horas efetivas (h)"].sum()

    cards = [
        caixa("Processos analisados", inteiro(len(resumo_processo)), 66 * mm),
        caixa("Área total (>20 ha/talhão)", f"{numero(total_area,1)} ha", 66 * mm),
        caixa("Talhões distintos", inteiro(total_talhoes), 66 * mm),
        caixa("Equipamentos distintos", inteiro(total_equip), 66 * mm),
    ]
    cards_tb = Table([cards], colWidths=[69 * mm] * 4)
    cards_tb.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 1 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 1 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(cards_tb)
    story.append(Spacer(1, 4 * mm))

    cabecalhos = ["PROCESSO", "TALHÕES (>20 ha)", "ÁREA TOTAL", "EQUIP.", "HORAS EFET.", "VELOC. MÉDIA", "RPM MÉDIO", "CONSUMO MÉDIO"]
    larguras = [55 * mm, 30 * mm, 30 * mm, 22 * mm, 30 * mm, 30 * mm, 30 * mm, 51 * mm]
    linhas = []
    for _, r in resumo_processo.iterrows():
        linhas.append([
            Paragraph(f"<b>{dados_cliente.rotulo_processo(r['Código processo']) if dados_cliente else r['Processo']}</b>", ParagraphStyle("proc", parent=STYLE_BODY, fontSize=10.5)),
            inteiro(r["Talhões (>20 ha)"]),
            f"{numero(r['Área operacional (ha)'],1)} ha",
            inteiro(r["Equipamentos"]),
            f"{numero(r['Horas efetivas (h)'],1)} h",
            f"{numero(r.get('Velocidade média'),2)} km/h",
            inteiro(r.get("RPM médio")),
            f"{numero(r.get('Consumo (L/ha)'),1)} L/ha",
        ])
    story.extend(tabela_generica(
        "VISÃO DOS PROCESSOS",
        "Consolidado de todos os talhões (>20 ha) de cada processo no período.",
        cabecalhos, larguras, linhas,
    ))
    story.append(Spacer(1, 3 * mm))
    story.append(Paragraph(
        "Os códigos de processo (cd_operacao) ainda não têm nome cadastrado — preencha o CATALOGO_OPERACOES "
        "no script de processamento para trocar \"Processo 1101\" por um nome descritivo (ex.: GRADAGEM, PLANTIO).",
        ParagraphStyle("obs", parent=STYLE_SMALL, fontSize=8.3, textColor=colors.HexColor("#7A8A82")),
    ))
    return story


# ------------------------------------------------------------
# PÁGINAS POR PROCESSO — 2 TALHÕES POR PÁGINA
#   tabela de indicadores + mapas de velocidade lado a lado + textos de análise
# ------------------------------------------------------------

STYLE_TXT = ParagraphStyle("txt", parent=STYLE_SMALL, fontSize=7.4, leading=9.0, textColor=colors.HexColor("#334155"))
STYLE_TXT_TIT = ParagraphStyle("txtt", parent=STYLE_TXT, fontName="Helvetica-Bold", fontSize=8.2, leading=10, textColor=TEXTO_PRINCIPAL)
STYLE_TXT_NOTA = ParagraphStyle("txtn", parent=STYLE_TXT, fontSize=6.6, leading=8, textColor=colors.HexColor("#7A8A82"))

ALTURA_MAPA_MM = 92
LARGURA_MAPA_MM = 136


def _imagem_proporcional(caminho_png, largura_max_mm, altura_max_mm):
    """Image flowable dimensionado dentro de (largura_max, altura_max),
    preservando a proporção real do PNG (nunca distorce)."""
    with PILImage.open(caminho_png) as img:
        w_px, h_px = img.size
    razao = w_px / h_px

    largura = largura_max_mm
    altura = largura / razao
    if altura > altura_max_mm:
        altura = altura_max_mm
        largura = altura * razao

    return Image(str(caminho_png), width=largura * mm, height=altura * mm)


def _un(u):
    return u if u == "%" else f" {u}"


def _classe_variacao(cv):
    if cv < 5:
        return "baixa", "#0A8F5A"
    if cv <= 10:
        return "moderada", "#EA7A0C"
    return "alta", "#DC2626"


def textos_talhao(cod_processo, talhao, rotulo, variacao, clima, mostrar_densidade=False):
    """Parágrafos de análise de um talhão: clima (% do tempo) + variação das médias."""
    itens = [Paragraph(f"Análises — Talhão {rotulo}", STYLE_TXT_TIT)]
    tem = False

    c = clima[(clima["Código processo"] == cod_processo) & (clima["Talhão"] == str(talhao))] if len(clima) else clima
    if len(c):
        r = c.iloc[0]
        partes = []
        for col, rotulo_col in (
            ("% tempo vento > 15", "vento acima de 15 km/h"),
            ("% tempo umidade < 45", "umidade relativa abaixo de 45%"),
            ("% tempo temp. > 35", "temperatura acima de 35°C"),
        ):
            if col in r and pd.notna(r[col]):
                partes.append(f"{rotulo_col}: <b>{percentual(r[col])}</b>")
        if partes:
            tem = True
            itens.append(Paragraph("<b>Clima</b> (% do tempo dos dados gerados): " + " · ".join(partes) + ".", STYLE_TXT))

    v = variacao[(variacao["Código processo"] == cod_processo) & (variacao["Talhão"] == str(talhao))] if len(variacao) else variacao
    for _, r in v.iterrows():
        if r["Indicador"] == "Densidade de plantio" and not mostrar_densidade:
            continue
        casas = int(r["Casas"])
        classe, cor = _classe_variacao(r["CV (%)"])
        itens.append(Paragraph(
            f"<b>{r['Indicador']}</b>: média {numero(r['Média'], casas)}{_un(r['Unidade'])}, variação típica de "
            f"±{numero(r['Desvio'], casas)}{_un(r['Unidade'])} (±{numero(r['CV (%)'], 1)}%) — "
            f"<font color='{cor}'><b>{classe}</b></font> "
            f"<font color='#94A3B8'>({numero(r['Dentro ±10% (%)'], 0)}% das leituras dentro de ±10% da média)</font>."
            + (" <font color='#DC2626'>Atenção: média fora da faixa usual (40 a 400 mil sem./ha) — conferir o dado/unidade na fonte.</font>"
               if r["Indicador"] == "Densidade de plantio" and not (40_000 <= r["Média"] <= 400_000) else ""),
            STYLE_TXT,
        ))
        tem = True

    if not tem:
        itens.append(Paragraph("Sem dados de clima ou de implemento para análise neste talhão.", STYLE_TXT))
    return itens


def texto_temp_motor(cod_processo, temp_equip, cliente):
    """Uma vez por processo: acionamentos de temperatura do motor, por equipamento."""
    if not len(temp_equip):
        return None
    t = temp_equip[temp_equip["Código processo"] == cod_processo]
    if not len(t):
        return None
    partes = []
    for _, r in t.iterrows():
        nome = cliente.rotulo_equip(r["Equipamento"])
        n105 = r.get("Acionamentos > 105°C")
        if pd.isna(n105):
            partes.append(f"<b>{nome}</b>: sem leitura de temperatura")
        elif n105 > 0:
            partes.append(f"<b>{nome}</b>: <font color='#DC2626'><b>{inteiro(n105)}</b></font>")
        else:
            partes.append(f"<b>{nome}</b>: 0")
    return Paragraph(
        "<b>Temperatura do motor</b> — vezes em que passou de 105°C (processo todo, estado Efetivo): "
        + " · ".join(partes) + ".",
        STYLE_TXT,
    )


def montar_colunas(base_talhao, mostrar_densidade=False):
    colunas_ctx = [
        ("Área operacional (ha)", "ÁREA", "ha", 1),
        ("Equipamentos (códigos)", "EQUIP. (FROTA)", "", 0),
        ("Horas efetivas (h)", "HORAS EFET.", "h", 1),
        ("Rendimento Real (%)", "RENDIMENTO REAL", "%", 1),
    ]
    # Temperatura do motor NÃO tem coluna (vira texto por equipamento).
    colunas_qual = [
        ("Velocidade média", "VELOCIDADE", "km/h", 2),
        ("RPM médio", "RPM", "RPM", 0),
        ("Consumo (L/ha)", "CONSUMO", "L/ha", 1),
        ("Torque médio", "TORQUE", "%", 1),
        ("Densidade de plantio", "DENSID. PLANTIO", "sem./ha", 0),
        ("Taxa de aplicação (sólido)", "TAXA APLIC.", "kg/ha", 1),
        ("Rendimento de colheita", "RENDIM. COLHEITA", "sc/ha", 1),
        ("Umidade dos grãos", "UMID. GRÃOS", "%", 1),
        ("Vazão de calda", "VAZÃO", "L/ha", 1),
        ("Pressão da bomba", "PRESSÃO BOMBA", "kPa", 1),
    ]
    # Densidade de plantio só aparece em processo de PLANTIO (identificado pela
    # descrição na planilha do cliente) — o dado vem errado em outras operações.
    colunas_qual = [c for c in colunas_qual if mostrar_densidade or c[0] != "Densidade de plantio"]
    colunas_qual = [c for c in colunas_qual if c[0] in base_talhao and base_talhao[c[0]].notna().any()]
    colunas_amb = [
        ("Temperatura ambiente média", "TEMP. AMBIENTE", "°C", 1),
        ("Umidade relativa média", "UMID. RELATIVA", "%", 1),
        ("Velocidade do vento média", "VENTO", "km/h", 1),
    ]
    colunas_amb = [c for c in colunas_amb if c[0] in base_talhao and base_talhao[c[0]].notna().any()]
    return colunas_ctx + colunas_qual + colunas_amb


def tabela_talhoes(grupo, todas_colunas, cliente):
    n = len(todas_colunas)
    largura_talhao = 24 * mm
    largura_equip = 30 * mm
    largura_resto = (278 * mm - largura_talhao - largura_equip) / (n - 1)

    cab = ["TALHÃO"] + [
        f"{label}<br/><font size='5.6' color='#94A3B8'>({un})</font>" if un else label
        for _, label, un, _ in todas_colunas
    ]
    head_st = ParagraphStyle("thq", parent=STYLE_TABLE_HEAD, fontSize=6.3, leading=7.3)
    corpo = [[Paragraph(c, head_st) for c in cab]]

    valor_st = ParagraphStyle("vlq", parent=STYLE_BODY, fontName="Helvetica-Bold", fontSize=8.3, leading=9.6, alignment=TA_CENTER)
    valor_st_frota = ParagraphStyle("vlf", parent=valor_st, fontSize=7.2, leading=8.4)
    for _, r in grupo.iterrows():
        nome = cliente.rotulo_talhao(r["Talhão"])
        cel = f"<b>{nome}</b>"
        if nome != cliente.codigo_curto(r["Talhão"]):
            cel += f"<br/><font size='5.6' color='#94A3B8'>cód. {cliente.codigo_curto(r['Talhão'])}</font>"
        linha = [Paragraph(cel, ParagraphStyle("tal", parent=STYLE_BODY, fontSize=10, leading=11, alignment=TA_CENTER))]
        for col, _, _, casas in todas_colunas:
            valor = r.get(col)
            if col == "Equipamentos (códigos)":
                frota = ", ".join(cliente.rotulo_equip(e) for e in str(valor).split(", ")) if pd.notna(valor) and valor else "-"
                linha.append(Paragraph(frota, valor_st_frota))
            elif col.endswith("(%)"):
                linha.append(Paragraph(percentual(valor), valor_st))
            else:
                linha.append(Paragraph(numero(valor, casas), valor_st))
        corpo.append(linha)

    larguras = [largura_talhao, largura_equip] + [largura_resto] * (n - 1)
    tb = Table(corpo, colWidths=larguras, repeatRows=1)
    tb.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), FUNDO_TABELA_HEAD),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BRANCO, FUNDO_CARD]),
        ("GRID", (0, 0), (-1, -1), 0.25, BORDA),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1.3 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 1 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 1.3 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.3 * mm),
    ]))
    return tb


def paginas_processo(cliente_nome, ini, fim, cod_processo, processo, base_talhao,
                     imagens_por_talhao, dados_cliente, variacao, clima, temp_equip):
    """Páginas de um processo: 2 talhões por página, cada uma com
    tabela + mapas de velocidade lado a lado + textos de análise."""
    plantio = dados_cliente.processo_eh_plantio(cod_processo)
    todas_colunas = montar_colunas(base_talhao, plantio)
    # imagens_por_talhao: {talhão: [png, ...]} (aceita também um caminho único)
    imagens_por_talhao = {t: (c if isinstance(c, (list, tuple)) else [c]) for t, c in imagens_por_talhao.items()}
    # 2 indicadores escolhidos no site -> 2 mapas por talhão: 1 talhão por página, mapas lado a lado
    dois_mapas = any(len(v) >= 2 for v in imagens_por_talhao.values())
    passo = 1 if dois_mapas else 2
    linhas = [base_talhao.iloc[i:i + passo] for i in range(0, len(base_talhao), passo)]
    story = []

    for idx, grupo in enumerate(linhas):
        rotulos = [dados_cliente.rotulo_talhao(t) for t in grupo["Talhão"]]
        story.append(header(
            cliente_nome, f"Período: {ini} a {fim}",
            f"PROCESSO {dados_cliente.rotulo_processo(cod_processo).upper()}:  TALHÃO {' E '.join(rotulos)}",
        ))
        story.append(Spacer(1, 2.5 * mm))
        story.append(tabela_talhoes(grupo, todas_colunas, dados_cliente))
        story.append(Spacer(1, 2 * mm))

        # ---- mapas lado a lado ----
        celulas_mapa = []
        for t in grupo["Talhão"]:
            caminhos = imagens_por_talhao.get(str(t), [])
            if dois_mapas:  # os 2 mapas (um por indicador) do mesmo talhão
                for c in caminhos[:2]:
                    celulas_mapa.append(_imagem_proporcional(c, LARGURA_MAPA_MM - 4, ALTURA_MAPA_MM))
                while len(celulas_mapa) < 2:
                    celulas_mapa.append(Paragraph("Mapa indisponível (sem dado do indicador escolhido)", STYLE_TXT))
            else:
                celulas_mapa.append(
                    _imagem_proporcional(caminhos[0], LARGURA_MAPA_MM - 4, ALTURA_MAPA_MM) if caminhos
                    else Paragraph("Mapa indisponível (CSV sem coordenadas ou sem dado do indicador)", STYLE_TXT)
                )
        unico = len(grupo) == 1  # textos: 1 talhão -> coluna única
        tb_mapas = Table([celulas_mapa], colWidths=[139 * mm, 139 * mm] if len(celulas_mapa) == 2 else [278 * mm])
        tb_mapas.setStyle(TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        story.append(tb_mapas)
        story.append(Spacer(1, 1.5 * mm))

        # ---- textos de análise (por talhão, embaixo do respectivo mapa) ----
        celulas_txt = [
            textos_talhao(cod_processo, t, dados_cliente.rotulo_talhao(t), variacao, clima, plantio)
            for t in grupo["Talhão"]
        ]
        tb_txt = Table([celulas_txt], colWidths=[278 * mm] if unico else [139 * mm, 139 * mm])
        tb_txt.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            *([] if unico else [("LINEBEFORE", (1, 0), (1, 0), 0.3, BORDA)]),
            ("LEFTPADDING", (0, 0), (-1, -1), 3 * mm), ("RIGHTPADDING", (0, 0), (-1, -1), 3 * mm),
            ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        story.append(tb_txt)

        # ---- 1ª página do processo: temp. do motor por equipamento (uma vez só) ----
        if idx == 0:
            tm = texto_temp_motor(cod_processo, temp_equip, dados_cliente)
            if tm is not None:
                story.append(Spacer(1, 1.5 * mm))
                story.append(tm)
        story.append(Spacer(1, 1 * mm))
        story.append(Paragraph(
            "Médias só no estado Efetivo e só entre equipamentos que têm o dado (0/-1 = sem dado); indicadores por ha ponderados pela área. "
            "Variação = dispersão típica (robusta, ignora picos isolados) em relação à média: baixa &lt;5% (verde), moderada 5–10% (laranja), alta &gt;10% (vermelho). "
            "Mapas (só estado Efetivo): indicador(es) escolhido(s) para o grupo de operação; sem escolha, velocidade ou, havendo, densidade de plantio, produtividade, taxa de sólidos ou vazão (cinza = equipamento sem dado). Azul = valor menor, vermelho = valor maior.",
            STYLE_TXT_NOTA,
        ))
        if idx < len(linhas) - 1:
            story.append(PageBreak())
    return story


# ------------------------------------------------------------
# 1º SLIDE DO PROCESSO — TABELA RESUMO POR TALHÃO (conforme o dado do processo)
#   vazão -> vazão/pressão/clima | taxa de sólidos | sementes (plantio) | colheita
# ------------------------------------------------------------

def _variacao_txt(variacao, cod, talhao, indicador):
    if not len(variacao):
        return "-"
    v = variacao[(variacao["Código processo"] == cod) & (variacao["Talhão"] == str(talhao)) & (variacao["Indicador"] == indicador)]
    if not len(v):
        return "-"
    cv = v.iloc[0]["CV (%)"]
    classe, cor = _classe_variacao(cv)
    return f"{numero(cv, 1)}% — <font color='{cor}'><b>{classe.capitalize()}</b></font>"


def _clima_pct(clima, cod, talhao, coluna):
    if not len(clima) or coluna not in clima.columns:
        return "-"
    c = clima[(clima["Código processo"] == cod) & (clima["Talhão"] == str(talhao))]
    if not len(c) or pd.isna(c.iloc[0][coluna]):
        return "-"
    return percentual(c.iloc[0][coluna])


def _tem(base, col):
    return col in base.columns and base[col].notna().any()


def tabelas_resumo_processo(cod, base, variacao, clima, plantio):
    """Lista de (título, [(cabeçalho, unidade, função(linha)->texto)]) conforme os
    dados que o processo tem."""
    area = ("ÁREA", "ha", lambda r: numero(r["Área operacional (ha)"], 1))
    vel = ("VELOCIDADE MÉDIA", "km/h", lambda r: numero(r.get("Velocidade média"), 2))
    tabelas = []

    if _tem(base, "Vazão de calda"):
        tabelas.append(("APLICAÇÃO — VAZÃO E PRESSÃO", [
            area,
            ("VAZÃO MÉDIA", "L/ha", lambda r: numero(r.get("Vazão de calda"), 1)),
            ("VARIAÇÃO DA VAZÃO", "", lambda r: _variacao_txt(variacao, cod, r["Talhão"], "Vazão de calda")),
            ("PRESSÃO MÉDIA", "kPa", lambda r: numero(r.get("Pressão da bomba"), 1)),
            ("VARIAÇÃO DA PRESSÃO", "", lambda r: _variacao_txt(variacao, cod, r["Talhão"], "Pressão da bomba")),
            ("VENTO > 15 km/h", "% do tempo", lambda r: _clima_pct(clima, cod, r["Talhão"], "% tempo vento > 15")),
            ("TEMPERATURA > 35°C", "% do tempo", lambda r: _clima_pct(clima, cod, r["Talhão"], "% tempo temp. > 35")),
            ("UMIDADE RELATIVA < 45%", "% do tempo", lambda r: _clima_pct(clima, cod, r["Talhão"], "% tempo umidade < 45")),
        ]))
    if _tem(base, "Taxa de aplicação (sólido)"):
        tabelas.append(("APLICAÇÃO DE SÓLIDOS — TAXA", [
            area, vel,
            ("TAXA DE SÓLIDOS", "kg/ha", lambda r: numero(r.get("Taxa de aplicação (sólido)"), 1)),
            ("VARIAÇÃO DA TAXA", "", lambda r: _variacao_txt(variacao, cod, r["Talhão"], "Taxa de aplicação")),
            ("VENTO MÉDIO", "km/h", lambda r: numero(r.get("Velocidade do vento média"), 1)),
            ("VENTO > 15 km/h", "% do tempo", lambda r: _clima_pct(clima, cod, r["Talhão"], "% tempo vento > 15")),
        ]))
    if plantio and _tem(base, "Densidade de plantio"):
        tabelas.append(("PLANTIO — SEMENTES", [
            area,
            ("SEMENTES MÉDIA", "sem./ha", lambda r: numero(r.get("Densidade de plantio"), 0)),
            ("VARIAÇÃO DAS SEMENTES", "", lambda r: _variacao_txt(variacao, cod, r["Talhão"], "Densidade de plantio")),
            vel,
        ]))
    if _tem(base, "Rendimento de colheita"):
        tabelas.append(("COLHEITA — PRODUTIVIDADE", [
            area,
            ("PRODUTIVIDADE MÉDIA", "sc/ha", lambda r: numero(r.get("Rendimento de colheita"), 1)),
            ("UMIDADE DOS GRÃOS", "%", lambda r: numero(r.get("Umidade dos grãos"), 1)),
        ]))
    return tabelas


def pagina_resumo_processo(cliente_nome, ini, fim, cod, base, variacao, clima, dados_cliente):
    tabelas = tabelas_resumo_processo(cod, base, variacao, clima, dados_cliente.processo_eh_plantio(cod))
    if not tabelas:
        return []
    story = [header(cliente_nome, f"Período: {ini} a {fim}",
                    f"PROCESSO {dados_cliente.rotulo_processo(cod).upper()}:  RESUMO POR TALHÃO"),
             Spacer(1, 3 * mm)]
    head_st = ParagraphStyle("thr", parent=STYLE_TABLE_HEAD, fontSize=7.4, leading=8.6)
    cel_st = ParagraphStyle("cr", parent=STYLE_BODY, fontName="Helvetica-Bold", fontSize=9.6, leading=11, alignment=TA_CENTER)
    for titulo, colunas in tabelas:
        corpo = [[Paragraph("TALHÃO", head_st)] + [
            Paragraph(f"{h}<br/><font size='6' color='#94A3B8'>({u})</font>" if u else h, head_st) for h, u, _ in colunas]]
        for _, r in base.iterrows():
            corpo.append([Paragraph(f"<b>{dados_cliente.rotulo_talhao(r['Talhão'])}</b>", cel_st)] +
                         [Paragraph(f(r), cel_st) for _, _, f in colunas])
        larg_t = 42 * mm
        tb = Table(corpo, colWidths=[larg_t] + [(278 * mm - larg_t) / len(colunas)] * len(colunas), repeatRows=1)
        tb.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), FUNDO_TABELA_HEAD),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BRANCO, FUNDO_CARD]),
            ("GRID", (0, 0), (-1, -1), 0.25, BORDA),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1.6 * mm), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6 * mm),
        ]))
        story += [barra_titulo(titulo), Spacer(1, 1.5 * mm), tb, Spacer(1, 4 * mm)]
    story.append(Paragraph(
        "Médias só no estado Efetivo e só entre equipamentos que têm o dado. Variação: baixa &lt;5%, moderada 5–10%, alta &gt;10% "
        "(dispersão típica em relação à média). Vento, temperatura e umidade: % do tempo dos dados gerados no talhão.", STYLE_TXT_NOTA))
    return story


# ------------------------------------------------------------
# CONSTRUÇÃO DO PDF
# ------------------------------------------------------------

def _aba(x, nome, colunas):
    """Lê uma aba (Talhão sempre como texto); vazia se não existir."""
    if nome not in x.sheet_names:
        return pd.DataFrame(columns=colunas)
    return pd.read_excel(x, nome, dtype={"Talhão": str, "Equipamento": str})


def carregar():
    x = pd.ExcelFile(ARQUIVO_ENTRADA)
    resumo_processo = pd.read_excel(x, "Resumo_Processo")
    resumo_talhao = pd.read_excel(x, "Resumo_Processo_Talhao", dtype={"Talhão": str, "Equipamentos (códigos)": str})
    validacao = pd.read_excel(x, "Validacao")
    variacao = _aba(x, "Variacao_Talhao", ["Código processo", "Talhão", "Indicador", "Unidade", "Casas", "Média", "Desvio", "CV (%)", "Dentro ±10% (%)"])
    clima = _aba(x, "Clima_Talhao", ["Código processo", "Talhão"])
    temp_equip = _aba(x, "TempMotor_Equip", ["Código processo", "Equipamento"])
    return resumo_processo, resumo_talhao, validacao, variacao, clima, temp_equip


def construir_pdf(caminho_csv=None):
    resumo_processo, resumo_talhao, validacao, variacao, clima, temp_equip = carregar()
    dados_cliente = carregar_cliente()
    dados_cliente.registrar_chaves(resumo_talhao["Talhão"].unique())

    m = dict(zip(validacao["Informação"], validacao["Valor"]))
    ini = pd.to_datetime(m.get("Data inicial")).strftime("%d/%m/%Y")
    fim = pd.to_datetime(m.get("Data final")).strftime("%d/%m/%Y")

    # ---- Mapas de velocidade (1 por Processo x Talhão > área mínima) ----
    imagens_por_chave = {}
    if caminho_csv:
        print("Gerando mapas de velocidade por talhão...")
        try:
            imagens_por_chave = gerar_mapas_velocidade(caminho_csv, PASTA_MAPAS)
        except ColunaGeometriaAusente as e:
            print(f"AVISO: {e}\n       O PDF será gerado normalmente, mas SEM os mapas.")
    else:
        print("Aviso: nenhum CSV informado — PDF sairá sem mapas.")

    nome = f"Relatorio_Qualidade_Operacional_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    saida = PASTA_RELATORIOS / nome

    doc = SimpleDocTemplate(
        str(saida), pagesize=landscape(A4),
        leftMargin=7 * mm, rightMargin=7 * mm, topMargin=7 * mm, bottomMargin=6 * mm,
        title="Relatório de Qualidade Operacional - Fazenda", author="Solinftec",
    )

    # ---- Seções Operação / Paradas / Equipamentos / Operadores / IA (sem metas) ----
    import analise_ia
    import secoes_operacao_pdf as sec
    secoes = sec.carregar_secoes(ARQUIVO_ENTRADA)
    rotulo_proc = lambda c: dados_cliente.rotulo_processo(int(c))

    def rotulo_equip(cod):
        nome = dados_cliente.rotulo_equip(cod)
        return str(cod) if nome == str(cod) else f"{cod} - {nome}"

    comparacao = None
    if secoes is not None:
        from opcoes_mapa import carregar_opcoes_relatorio
        _o = carregar_opcoes_relatorio()
        _usa_op = bool(_o.get("comparar_metas_operacao", _o.get("comparar_metas")))
        _usa_eq = bool(_o.get("comparar_metas_equipamentos", _o.get("comparar_metas")))
        if _usa_op or _usa_eq:
            import metas as _metas
            print("Comparação com metas operacionais ligada.")
            m = _metas.carregar_metas(_PASTA_CLIENTE / _metas.NOME_ARQUIVO)
            if m is not None:
                try:
                    comparacao = _metas.calcular_comparacao(
                        secoes, m, rotulo_proc, lambda c: dados_cliente.processos.get(_chave_dc(c), ""),
                        rotulo_equip, usar_op=_usa_op, usar_eq=_usa_eq)
                    cp, ce = comparacao
                    comparacao = (cp if _usa_op else None, ce if _usa_eq else None)
                    print(f"  Metas aplicadas: operação={'sim' if _usa_op else 'não'}, equipamentos={'sim' if _usa_eq else 'não'}.")
                except Exception as e:  # noqa: BLE001
                    print(f"  AVISO: falha ao comparar com metas ({type(e).__name__}: {e}) — relatório sem comparação.")
    analise, usar_ia = None, False
    if secoes is not None:
        usar_ia = bool(analise_ia._ler_chave())
        if usar_ia:
            print("Gerando análise de IA (Gemini)...")
            analise = analise_ia.gerar_analise(secoes, CLIENTE_NOME, rotulo_proc, rotulo_equip, ARQUIVO_ENTRADA.parent, comparacao=comparacao)
        else:
            print("Aviso: GEMINI_API_KEY não configurada — o PDF sairá sem as páginas de IA.")

    story = []
    if secoes is not None:
        story.extend(sec.pagina_resumo_executivo(CLIENTE_NOME, ini, fim, secoes, analise, rotulo_proc, comparacao))
    else:
        story.extend(pagina_resumo_geral(CLIENTE_NOME, ini, fim, resumo_processo, resumo_talhao, dados_cliente))
    story.append(PageBreak())

    codigos_talhao = set(resumo_processo["Código processo"].tolist())
    codigos = secoes["Operacao"]["Código processo"].tolist() if secoes is not None else resumo_processo["Código processo"].tolist()
    for i, cod in enumerate(codigos):
        if secoes is not None:
            story.extend(sec.paginas_processo(cod, CLIENTE_NOME, ini, fim, secoes, analise, rotulo_proc, rotulo_equip,
                                              PASTA_RELATORIOS / "_graficos_tmp", usar_ia, comparacao))
            if cod in codigos_talhao:
                story.append(PageBreak())
        if cod in codigos_talhao:
            base_talhao = resumo_talhao[resumo_talhao["Código processo"] == cod].copy()
            processo = base_talhao["Processo"].iloc[0]
            imagens_por_talhao = {t: c for (op, t), c in imagens_por_chave.items() if op == cod}
            resumo_pg = pagina_resumo_processo(CLIENTE_NOME, ini, fim, cod, base_talhao, variacao, clima, dados_cliente)
            if resumo_pg:  # tabela resumo por talhão, antes dos mapas
                story.extend(resumo_pg)
                story.append(PageBreak())
            story.extend(paginas_processo(
                CLIENTE_NOME, ini, fim, cod, processo, base_talhao, imagens_por_talhao,
                dados_cliente, variacao, clima, temp_equip,
            ))
        if i < len(codigos) - 1:
            story.append(PageBreak())

    doc.build(story)  # sem numeração de páginas
    return saida


def _encontrar_csv_mais_recente():
    """Usado quando o script roda sozinho (sem receber o caminho do
    CSV por argumento) — pega o mapa analítico mais recente em
    dados_brutos/, mesma convenção do orquestrador."""
    pasta = Path("dados_brutos")
    candidatos = sorted(pasta.glob("*.csv"), key=lambda p: p.stat().st_mtime, reverse=True) if pasta.exists() else []
    return candidatos[0] if candidatos else None


if __name__ == "__main__":
    caminho_csv = sys.argv[1] if len(sys.argv) > 1 else None
    if caminho_csv is None:
        encontrado = _encontrar_csv_mais_recente()
        caminho_csv = str(encontrado) if encontrado else None

    print("Gerando PDF de qualidade operacional...")
    pdf = construir_pdf(caminho_csv)
    print(f"PDF gerado: {pdf}")
    print(f"Tamanho: {pdf.stat().st_size/1024:.1f} KB")
