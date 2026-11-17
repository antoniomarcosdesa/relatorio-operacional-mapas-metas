# ============================================================
# AREA_TRABALHADA.PY — área trabalhada (sem sobreposição), opcional
#
# A "área operacional" do relatório inclui a sobreposição das passadas; a "área trabalhada"
# não. Ela vem de um Excel EXPORTADO À PARTE (colunas: Processo | Fazenda | Cód. Zona |
# Talhão | Data | Área Trabalhada (ha) | Área Operacional (ha) | Área Sobreposição (%)),
# em formato hierárquico com linhas "Total". Usamos:
#   - total do processo : linha com Fazenda = "Total" (soma do período);
#   - total do talhão   : linha com Processo+Fazenda+Zona+Talhão e Data = "Total".
# O Excel traz NOMES (processo "ROLO FACA", talhão "C 01"); o casamento com o relatório é
# por nome normalizado (sem acento/maiúsculas/símbolos) e zona. O que não casar é avisado
# no log e fica em branco. Sem o arquivo, nada muda no relatório.
# ============================================================

import re
import unicodedata
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

NOME_ARQUIVO = "area_trabalhada.xlsx"
COL_PROC = "Área trabalhada (ha)"
COL_SOBRE = "Sobreposição (%)"


def _n(t) -> str:
    if t is None or (not isinstance(t, str) and pd.isna(t)):
        return ""
    t = unicodedata.normalize("NFKD", str(t).upper())
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]+", "", t)


def _zona(z) -> str:
    s = str(z).strip()
    return re.sub(r"\.0$", "", s) if s.lower() != "nan" else ""


def _achar(df, *nomes):
    por = {_n(c): c for c in df.columns}
    for n in nomes:
        if _n(n) in por:
            return por[_n(n)]
    return None


def carregar(arquivo, log=print):
    """{'processos': {proc_norm: (trab, oper, sobre%)}, 'talhoes': {(proc_norm, zona, tal_norm): (...)}} ou None."""
    arquivo = Path(arquivo)
    if not arquivo.exists():
        log(f"  AVISO: arquivo de área trabalhada não encontrado ({arquivo.name}) — relatório sem esse dado.")
        return None
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            d = pd.read_excel(arquivo, dtype=object)
        c = {k: _achar(d, *v) for k, v in {
            "proc": ("Processo",), "faz": ("Fazenda",), "zona": ("Cód. Zona", "Cod Zona", "Zona"),
            "tal": ("Talhão", "Talhao"), "data": ("Data",), "trab": ("Área Trabalhada (ha)", "Area Trabalhada"),
            "oper": ("Área Operacional (ha)", "Area Operacional"), "sob": ("Área Sobreposição (%)", "Area Sobreposicao")}.items()}
        if not c["proc"] or not c["trab"]:
            log(f"  AVISO: {arquivo.name} sem colunas 'Processo' e 'Área Trabalhada (ha)' — ignorado.")
            return None
        num = lambda v: float(v) if v is not None and not pd.isna(v) else np.nan
        procs, tals = {}, {}
        for _, r in d.iterrows():
            proc = r[c["proc"]]
            if proc is None or pd.isna(proc) or _n(proc) in ("", "TOTAL") or str(proc).startswith("Filtros"):
                continue
            trab = num(r[c["trab"]])
            oper = num(r[c["oper"]]) if c["oper"] else np.nan
            sob = num(r[c["sob"]]) * 100 if c["sob"] and not pd.isna(r[c["sob"]]) else np.nan
            faz = str(r[c["faz"]]).strip() if c["faz"] else ""
            if _n(faz) == "TOTAL":
                procs[_n(proc)] = (trab, oper, sob)
                continue
            tal = r[c["tal"]] if c["tal"] else None
            data = str(r[c["data"]]).strip() if c["data"] else ""
            if tal is None or pd.isna(tal) or _n(tal) in ("", "TOTAL") or _n(data) != "TOTAL":
                continue
            zona = _zona(r[c["zona"]]) if c["zona"] else ""
            chave = (_n(proc), zona, _n(tal))
            antes = tals.get(chave)
            # mesma zona+talhão em fazendas diferentes: soma (a chave do relatório não traz o nome da fazenda)
            tals[chave] = (trab, oper, sob) if antes is None else (
                np.nansum([antes[0], trab]), np.nansum([antes[1], oper]),
                (1 - np.nansum([antes[0], trab]) / np.nansum([antes[1], oper])) * 100 if np.nansum([antes[1], oper]) else np.nan)
        if not procs and not tals:
            log(f"  AVISO: {arquivo.name} sem linhas de total utilizáveis — ignorado.")
            return None
        log(f"  Área trabalhada: {arquivo.name} — {len(procs)} processo(s), {len(tals)} talhão(ões).")
        return {"processos": procs, "talhoes": tals}
    except Exception as e:  # noqa: BLE001
        log(f"  AVISO: falha ao ler {arquivo.name} ({type(e).__name__}: {e}) — relatório sem área trabalhada.")
        return None


def _nome_proc(rotulo) -> str:
    """'1014 - GRADE INTERMEDIARIA' -> 'GRADEINTERMEDIARIA'."""
    return _n(re.sub(r"^\s*-?\d+\s*[-–:]?\s*", "", str(rotulo)))


def achar_processo(dados, rotulo):
    k = _nome_proc(rotulo)
    if not k:
        return None, ""
    if k in dados["processos"]:
        return dados["processos"][k], k
    cand = [p for p in dados["processos"] if p in k or k in p]
    return (dados["processos"][cand[0]], cand[0]) if len(cand) == 1 else (None, "")


def aplicar_operacao(secoes, dados, rotulo_proc, log=print):
    """Acrescenta 'Área trabalhada (ha)' e 'Sobreposição (%)' em secoes['Operacao']."""
    op = secoes["Operacao"]
    trab, sob, nao = [], [], []
    for cod in op["Código processo"]:
        v, _ = achar_processo(dados, rotulo_proc(int(cod)))
        trab.append(v[0] if v else np.nan)
        sob.append(v[2] if v else np.nan)
        if not v:
            nao.append(rotulo_proc(int(cod)))
    op[COL_PROC], op[COL_SOBRE] = trab, sob
    if nao:
        log("  AVISO área trabalhada: processo(s) sem correspondência no Excel (confira o nome no de-para): " + "; ".join(nao))


def aplicar_talhoes(base_talhao, cod, dados, dados_cliente, rotulo_proc, log=print):
    """Acrescenta 'Área trabalhada (ha)' ao resumo por talhão do processo `cod`."""
    _, kp = achar_processo(dados, rotulo_proc(int(cod)))
    valores = []
    for t in base_talhao["Talhão"]:
        f, z, tt = (str(t).split("-", 2) + ["", "", ""])[:3] if str(t).count("-") >= 2 else ("", "", str(t))
        nomes = {_n(dados_cliente.nome_talhao(t)), _n(tt)} - {""}
        v = None
        for zn in (_zona(z), ""):
            for nm in nomes:
                v = dados["talhoes"].get((kp, zn, nm))
                if v:
                    break
            if v:
                break
        if v is None and kp:  # sem zona no Excel: tenta qualquer zona com o mesmo nome
            cand = [x for (p, zz, n), x in dados["talhoes"].items() if p == kp and n in nomes]
            v = cand[0] if len(cand) == 1 else None
        valores.append(v[0] if v else np.nan)
    base_talhao[COL_PROC] = valores
    return base_talhao
