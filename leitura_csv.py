# Leitura padronizada do CSV bruto (mapa analítico), compartilhada pelos scripts.
import pandas as pd

COLUNAS_TEXTO = {"cd_fazenda": str, "cd_zona": str, "cd_talhao": str}


def _num(v) -> str:
    """'1', '1.0', ' 1 ' -> '1' (códigos numéricos comparáveis); texto fica como está."""
    t = str(v).strip()
    if t.lower() in ("nan", "none"):
        return ""
    if t.replace(".", "", 1).isdigit() and t.endswith(".0"):
        t = t[:-2]
    if t.isdigit():  # '00001' e '1' são o mesmo código
        return str(int(t))
    return t


def chave_talhao(fazenda, zona, talhao) -> str:
    """Identidade do talhão = FAZENDA + ZONA + TALHÃO (o mesmo código de talhão pode
    existir em fazendas/zonas diferentes). Ex.: fazenda 1, zona 1, talhão 47 -> '1-1-47'."""
    return f"{_num(fazenda)}-{_num(zona)}-{_num(talhao)}"


def partes_talhao(chave: str):
    """'1-1-47' -> ('1', '1', '47'); chave simples (sem fazenda/zona) -> ('', '', chave)."""
    p = str(chave).split("-", 2)
    return tuple(p) if len(p) == 3 else ("", "", str(chave))


def ler_csv_bruto(caminho, incluir_wkt=True) -> pd.DataFrame:
    """Lê o CSV com cd_fazenda/cd_zona/cd_talhao sempre como texto (evita
    tipos misturados, ex.: '1' vs 1 vs 1.0, e o DtypeWarning)."""
    # incluir_wkt=False: não carrega a coluna de coordenadas (a mais pesada do arquivo) —
    # a etapa de indicadores não precisa dela; só a etapa de mapas.
    usecols = None if incluir_wkt else (lambda c: str(c).strip().lower() != "wkt")
    df = pd.read_csv(caminho, sep=";", encoding="utf-8", dtype=COLUNAS_TEXTO, low_memory=False, usecols=usecols)
    df.columns = [str(c).strip() for c in df.columns]

    # Talhão = fazenda + zona + talhão. `cd_talhao` passa a ser essa chave composta
    # (o código original fica em `cd_talhao_orig`); talhão não definido sai aqui.
    if "cd_talhao" in df.columns:
        orig = df["cd_talhao"].fillna("").astype(str).str.strip()
        df = df[(orig != "") & (orig.str.upper() != "*NAO_DEFINIDO*")].copy()
        df["cd_talhao_orig"] = orig[df.index]
        if "cd_fazenda" in df.columns or "cd_zona" in df.columns:
            faz = df["cd_fazenda"] if "cd_fazenda" in df.columns else ""
            zona = df["cd_zona"] if "cd_zona" in df.columns else ""
            df["cd_talhao"] = [chave_talhao(f, z, t) for f, z, t in zip(
                faz if len(faz) else [""] * len(df), zona if len(zona) else [""] * len(df), df["cd_talhao_orig"])]
    return df


def parse_datas(serie: pd.Series) -> pd.Series:
    """Datas em ISO (2026-10-12 07:00:00) são lidas como ano-mês-dia; as demais
    (12/10/2026 07:00) como dia/mês/ano. Antes, dayfirst=True aplicado a datas
    ISO invertia dia e mês (12/10 virava 10/12)."""
    s = serie.astype(str).str.strip()
    iso = s.str.match(r"^\d{4}-\d{2}-\d{2}")
    em_iso = pd.to_datetime(s, format="ISO8601", errors="coerce")
    em_br = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    if (~iso).any():  # só as datas NÃO-ISO usam dia/mês/ano (evita o aviso do pandas)
        em_br[~iso] = pd.to_datetime(s[~iso], dayfirst=True, errors="coerce")
    return em_iso.where(iso, em_br)


def achar_coluna(df: pd.DataFrame, nome: str):
    """Procura a coluna ignorando maiúsculas/espaços (ex.: 'WKT', ' wkt ')."""
    for c in df.columns:
        if str(c).strip().lower() == nome.lower():
            return c
    return None
