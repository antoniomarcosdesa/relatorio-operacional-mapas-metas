# ============================================================
# DADOS_CLIENTE.PY
#
# Lê o que fica na pasta  cliente/  :
#   - um Excel (.xlsx) com a NOMENCLATURA DOS TALHÕES (código -> nome) e
#     o CÓDIGO/NOME DOS EQUIPAMENTOS (cd_equipamento -> frota/nome);
#   - um arquivo .kml com os POLÍGONOS DOS TALHÕES (mapa da fazenda), usado
#     para desenhar o contorno do talhão nos mapas do PDF.
#
# Tudo é opcional: sem a pasta (ou sem os arquivos), o relatório sai igual
# antes (códigos puros e mapas sem contorno). O que foi encontrado é
# impresso no log, para conferir se as colunas foram reconhecidas.
#
# Reconhecimento das colunas do Excel (por nome do cabeçalho, ignorando
# acentos/maiúsculas — qualquer aba serve):
#   talhão  : código = cd_talhao / codigo_talhao / talhao ...
#             nome   = coluna com "talh" + nome/desc/nomenclatura
#   equip.  : código = cd_equipamento / codigo_equipamento / equipamento ...
#             nome   = coluna "frota" ou equipamento + nome/desc
# Exemplo pronto: cliente/cliente_codigos.xlsx
# ============================================================

from pathlib import Path
import re
import unicodedata
import json
import struct
import xml.etree.ElementTree as ET

import pandas as pd

from leitura_csv import chave_talhao, partes_talhao

PASTA_CLIENTE = Path("cliente")


def _norm(txt) -> str:
    t = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "_", t.lower()).strip("_")


def _chave(valor) -> str:
    """Chave de comparação de códigos: '001', '1' e '1.0' viram '1'."""
    t = str(valor).strip()
    if re.fullmatch(r"\d+(\.0+)?", t):
        return str(int(float(t)))
    return _norm(t)


def _limpa(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    t = str(v).strip()
    return "" if t.lower() in ("nan", "none") else (t[:-2] if re.fullmatch(r"\d+\.0", t) else t)


class DadosCliente:
    def __init__(self):
        self.talhoes = {}     # 'fazenda-zona-talhão' -> nome do talhão
        self.talhoes_simples = {}  # só talhão (planilha sem fazenda/zona) -> nome
        self.multi = False    # há mais de uma fazenda/zona nos dados?
        self._ambiguos = set()
        self.equip = {}       # chave -> frota/nome
        self.processos = {}   # chave -> descrição do processo
        self.grupos = {}      # chave -> grupo de operação informado no de-para (texto livre)
        self.poligonos = []   # [(nomes_normalizados:set, [anéis (lon,lat)])]
        self._indice_kml = {}
        self.poligonos_chave = {}  # 'fazenda-zona-talhão' -> anéis (GeoJSON)

    # ---- consultas -------------------------------------------------
    def registrar_chaves(self, chaves):
        """Informa as chaves 'fazenda-zona-talhão' dos dados, para saber quando o
        código do talhão sozinho é ambíguo (mesmo talhão em fazendas/zonas diferentes)."""
        chaves = {str(c) for c in chaves}
        por_tal = {}
        for c in chaves:
            f, z, t = partes_talhao(c)
            por_tal.setdefault(t, set()).add((f, z))
        self._ambiguos = {t for t, fz in por_tal.items() if len(fz) > 1}
        self.multi = len({(partes_talhao(c)[0], partes_talhao(c)[1]) for c in chaves}) > 1

    def nome_talhao(self, chave) -> str:
        f, z, t = partes_talhao(chave)
        return self.talhoes.get(chave_talhao(f, z, t), "") or self.talhoes_simples.get(_chave(t), "")

    def rotulo_talhao(self, chave) -> str:
        """Nome do talhão pela nomenclatura do cliente (fazenda+zona+talhão); sem nome,
        o código (com fazenda/zona só se o código sozinho for ambíguo)."""
        nome = self.nome_talhao(chave)
        if nome:
            return nome
        f, z, t = partes_talhao(chave)
        return f"{t} (F{f}/Z{z})" if t in self._ambiguos else t

    def codigo_curto(self, chave) -> str:
        f, z, t = partes_talhao(chave)
        return f"{f}/{z}/{t}" if self.multi else t

    def rotulo_processo(self, cod) -> str:
        """'1101 - GRADAGEM' se o Excel do cliente tiver a descrição; senão só o código."""
        desc = self.processos.get(_chave(cod), "")
        return f"{cod} - {desc}" if desc else str(cod)

    def processo_eh_plantio(self, cod) -> bool:
        """True se a descrição do processo (planilha do cliente) tem 'plantio'."""
        return "plantio" in _norm(self.processos.get(_chave(cod), ""))

    def grupo_processo(self, cod) -> str:
        """Chave do grupo de operação (PLANTIO, COLHEITA, PULVERIZACAO, PREPARO_SOLO, LANCO) ou ''.
        Vem da coluna 'grupo' do de-para; sem ela, é deduzido da descrição do processo."""
        from opcoes_mapa import normalizar_grupo
        k = _chave(cod)
        return normalizar_grupo(self.grupos.get(k, "")) or normalizar_grupo(self.processos.get(k, ""))

    def rotulo_equip(self, cod) -> str:
        return self.equip.get(_chave(cod), "") or str(cod)

    def contorno_talhao(self, cod):
        """Anéis (lista de [(lon, lat), ...]) do talhão (GeoJSON pela chave fazenda-zona-talhão,
        ou KML pelo nome/código), ou []."""
        if str(cod) in self.poligonos_chave:
            return self.poligonos_chave[str(cod)]
        candidatos = [_chave(partes_talhao(cod)[2])]
        nome = self.nome_talhao(cod)
        if nome:
            candidatos += [_chave(nome), _norm(nome)]
        for c in candidatos:
            if c in self._indice_kml:
                return self._indice_kml[c]
        return []

    # ---- Excel -----------------------------------------------------
    def _ler_excel(self, caminho: Path):
        try:
            abas = pd.read_excel(caminho, sheet_name=None, dtype=str)
        except Exception as e:
            print(f"  AVISO: não consegui ler {caminho.name}: {e}")
            return
        for aba, df in abas.items():
            df.columns = [str(c).strip() for c in df.columns]
            cols = {c: _norm(c) for c in df.columns}

            # ---- talhão ----
            cod_t = next((c for c, n in cols.items() if n in (
                "cd_talhao", "codigo_talhao", "cod_talhao", "id_talhao", "codigo_do_talhao")), None)
            if cod_t is None:
                cod_t = next((c for c, n in cols.items() if n == "talhao"), None)
            if cod_t is not None:
                nome_t = next((c for c, n in cols.items() if c != cod_t and "talh" in n and any(
                    k in n for k in ("nome", "desc", "nomenclatura", "denomin"))), None)
                if nome_t is None:
                    nome_t = next((c for c, n in cols.items() if c != cod_t and "talh" in n), None)
                if nome_t is None:
                    nome_t = next((c for c, n in cols.items() if c != cod_t and n in (
                        "nome", "descricao", "nomenclatura")), None)
                if nome_t is not None:
                    col_f = next((c for c, n in cols.items() if n in ("cd_fazenda", "codigo_fazenda", "fazenda")), None)
                    col_z = next((c for c, n in cols.items() if n in ("cd_zona", "codigo_zona", "zona")), None)
                    for _, r in df.iterrows():
                        cod, nome = _limpa(r[cod_t]), _limpa(r[nome_t])
                        if cod and nome:
                            if col_f or col_z:  # identidade = fazenda + zona + talhão
                                k = chave_talhao(_limpa(r[col_f]) if col_f else "", _limpa(r[col_z]) if col_z else "", cod)
                                self.talhoes.setdefault(k, nome)
                            else:
                                self.talhoes_simples.setdefault(_chave(cod), nome)
                    ident = f"fazenda='{col_f}' + zona='{col_z}' + " if (col_f or col_z) else ""
                    print(f"  Talhões   (aba '{aba}'): {ident}código='{cod_t}' -> nome='{nome_t}'")

            # ---- processo (cd_operacao -> descrição) ----
            cod_p = next((c for c, n in cols.items() if n in (
                "cd_operacao", "codigo_operacao", "cod_operacao", "cd_processo", "codigo_processo",
                "cod_processo", "id_processo", "codigo_do_processo")), None)
            if cod_p is None:
                cod_p = next((c for c, n in cols.items() if n in ("processo", "operacao")), None)
            if cod_p is not None:
                nome_p = next((c for c, n in cols.items() if c != cod_p and ("process" in n or "operac" in n)), None)
                if nome_p is None:
                    nome_p = next((c for c, n in cols.items() if c != cod_p and n in ("descricao", "nome")), None)
                if nome_p is not None:
                    for _, r in df.iterrows():
                        cod, nome = _limpa(r[cod_p]), _limpa(r[nome_p])
                        if cod and nome:
                            self.processos.setdefault(_chave(cod), nome)
                    col_g = next((c for c, n in cols.items() if c not in (cod_p, nome_p) and n in (
                        "grupo", "grupo_operacao", "grupo_de_operacao", "grupo_processo", "grupo_do_processo",
                        "grupo_de_operacoes", "tipo_operacao", "tipo_de_operacao")), None)
                    if col_g is not None:
                        for _, r in df.iterrows():
                            cod, g = _limpa(r[cod_p]), _limpa(r[col_g])
                            if cod and g:
                                self.grupos.setdefault(_chave(cod), g)
                    extra = f" + grupo='{col_g}'" if col_g else " (sem coluna 'grupo': deduzido da descrição)"
                    print(f"  Processos (aba '{aba}'): código='{cod_p}' -> descrição='{nome_p}'{extra}")

            # ---- equipamento ----
            cod_e = next((c for c, n in cols.items() if n in (
                "cd_equipamento", "codigo_equipamento", "cod_equipamento", "id_equipamento",
                "cd_equip", "codigo_equip", "codigo_do_equipamento")), None)
            if cod_e is None:
                cod_e = next((c for c, n in cols.items() if n == "equipamento"), None)
            if cod_e is not None:
                nome_e = next((c for c, n in cols.items() if c != cod_e and "frota" in n), None)
                if nome_e is None:
                    nome_e = next((c for c, n in cols.items() if c != cod_e and "equip" in n and any(
                        k in n for k in ("nome", "desc"))), None)
                if nome_e is None:
                    nome_e = next((c for c, n in cols.items() if c != cod_e and n in (
                        "nome", "descricao")), None)
                if nome_e is not None:
                    for _, r in df.iterrows():
                        cod, nome = _limpa(r[cod_e]), _limpa(r[nome_e])
                        if cod and nome:
                            self.equip.setdefault(_chave(cod), nome)
                    print(f"  Equipam.  (aba '{aba}'): código='{cod_e}' -> frota/nome='{nome_e}'")

    # ---- GeoJSON (modelo de talhões: FAZENDA, ZONA, TALHAO, DESC_TALHA) -----------
    def _ler_geojson(self, caminho: Path):
        try:
            dados = json.loads(caminho.read_text(encoding="utf-8"))
        except Exception as e:
            print(f"  AVISO: não consegui ler {caminho.name}: {e}")
            return
        n = 0
        for feat in dados.get("features", []):
            props = {_norm(k): v for k, v in (feat.get("properties") or {}).items()}
            pega = lambda *nomes: next((props[x] for x in nomes if x in props and props[x] not in (None, "")), "")
            faz = pega("fazenda", "cd_fazenda")
            zona = pega("zona", "cd_zona")
            tal = pega("talhao", "cd_talhao")
            if tal == "":
                continue
            chave = chave_talhao(faz, zona, tal)
            desc = _limpa(pega("desc_talha", "desc_talhao", "descricao", "nome_talhao", "nome"))
            if desc:
                self.talhoes[chave] = desc  # a descrição do arquivo de talhões vale como nome
            geom = feat.get("geometry") or {}
            tipo, coords = geom.get("type"), geom.get("coordinates") or []
            poligonos = coords if tipo == "MultiPolygon" else [coords] if tipo == "Polygon" else []
            aneis = [[(float(pt[0]), float(pt[1])) for pt in anel] for pol in poligonos for anel in pol if len(anel) >= 3]
            if aneis:
                self.poligonos_chave.setdefault(chave, []).extend(aneis)
                n += 1
        print(f"  GeoJSON   ({caminho.name}): {n} talhão(ões) com fazenda+zona+talhão, descrição e contorno")

    # ---- mapa .yms + .ymx (formato próprio: geometria + atributos) -----------------
    def _ler_yms_ymx(self, yms: Path, ymx: Path):
        """Lê o par .yms (contornos) / .ymx (atributos FAZENDA, TALHAO, ZONA, DESC_TALHA...).
        Coordenadas em inteiros /1e7; a latitude vem sem sinal e é tratada como hemisfério
        sul (Brasil). Cada registro do .ymx guarda no início o deslocamento do seu contorno."""
        try:
            g = yms.read_bytes()
            a = ymx.read_bytes()
            nrec, nf = struct.unpack("<i", a[4:8])[0], a[8]
            tam = struct.unpack("<H", a[9:11])[0]
            campos, off = [], 16
            for _ in range(nf):
                d = a[off:off + 16]
                campos.append((d[:11].split(b"\0")[0].decode("latin-1"), d[12]))
                off += 16
            n = 0
            for r in range(nrec):
                rec = a[off + r * tam: off + (r + 1) * tam]
                desloc = struct.unpack("<i", rec[:4])[0]
                p, attrs = 4, {}
                for nome, w in campos:
                    attrs[_norm(nome)] = rec[p:p + w].decode("latin-1").strip("\0 ").strip()
                    p += w
                pos = 24 + desloc  # início dos dados no .yms = byte 24
                npart, npt = struct.unpack("<ii", g[pos + 16:pos + 24])
                inicios = list(struct.unpack(f"<{npart}i", g[pos + 24:pos + 24 + 4 * npart]))
                q = pos + 24 + 4 * npart
                v = struct.unpack(f"<{2 * npt}i", g[q:q + 8 * npt])
                pts = [(v[2 * k] / 1e7, -abs(v[2 * k + 1]) / 1e7) for k in range(npt)]
                fins = inicios[1:] + [npt]
                aneis = [pts[i:j] for i, j in zip(inicios, fins) if j - i >= 3]
                if attrs.get("talhao", "") == "":
                    continue
                chave = chave_talhao(attrs.get("fazenda", ""), attrs.get("zona", ""), attrs["talhao"])
                desc = attrs.get("desc_talha") or attrs.get("desc_talhao") or ""
                if desc:
                    self.talhoes[chave] = desc
                if aneis:
                    self.poligonos_chave.setdefault(chave, []).extend(aneis)
                    n += 1
            print(f"  Mapa      ({yms.name} + {ymx.name}): {n} talhão(ões) com fazenda+zona+talhão, descrição e contorno")
        except Exception as e:
            print(f"  AVISO: não consegui ler o par {yms.name}/{ymx.name}: {e}")

    # ---- MBTiles vetorial (.mbtiles, ex.: fields.mbtiles / info.mbtiles) --------------
    def _ler_mbtiles(self, caminho: Path):
        """Lê um .mbtiles VETORIAL (tiles .pbf) com os campos FAZENDA, ZONA, TALHAO, DESC_TALHA...
        Usa o zoom máximo, junta os pedaços de cada talhão (os tiles cortam os polígonos) e
        grava descrição + contorno. MBTiles raster (imagem) não tem atributos e é ignorado."""
        import gzip
        import math
        import sqlite3
        try:
            con = sqlite3.connect(f"file:{caminho}?mode=ro", uri=True)
            meta = dict(con.execute("select name, value from metadata"))
            if str(meta.get("format", "")).lower() != "pbf":
                print(f"  AVISO: {caminho.name} não é MBTiles vetorial (format={meta.get('format')}) — ignorado.")
                return
            zmax = con.execute("select max(zoom_level) from tiles").fetchone()[0]
            pedacos = {}  # chave -> [anéis (lon,lat) de cada pedaço]
            descs = {}
            for col, row, dado in con.execute("select tile_column, tile_row, tile_data from tiles where zoom_level=?", (zmax,)):
                try:
                    dado = gzip.decompress(dado)
                except Exception:
                    pass
                y_xyz = (1 << zmax) - 1 - row
                for attrs, aneis in _decodificar_mvt(dado, zmax, col, y_xyz):
                    a = {_norm(k): v for k, v in attrs.items()}
                    tal = a.get("talhao", "")
                    if tal in ("", None):
                        continue
                    chave = chave_talhao(a.get("fazenda", ""), a.get("zona", ""), tal)
                    d = _limpa(a.get("desc_talha") or a.get("desc_talhao") or "")
                    if d:
                        descs[chave] = d
                    pedacos.setdefault(chave, []).extend(aneis)
            con.close()
        except Exception as e:
            print(f"  AVISO: não consegui ler {caminho.name}: {e}")
            return
        n = 0
        for chave, aneis in pedacos.items():
            if chave in self.poligonos_chave:
                continue  # já veio de outro arquivo (ex.: fields + info idênticos)
            self.poligonos_chave[chave] = _unir_pedacos(aneis)
            if chave in descs:
                self.talhoes[chave] = descs[chave]
            n += 1
        print(f"  MBTiles   ({caminho.name}): {n} talhão(ões) com fazenda+zona+talhão, descrição e contorno")

    # ---- KML -------------------------------------------------------
    def _ler_kml(self, caminho: Path):
        try:
            raiz = ET.parse(caminho).getroot()
        except Exception as e:
            print(f"  AVISO: não consegui ler {caminho.name}: {e}")
            return

        def tag(el):
            return el.tag.split("}")[-1]

        n = 0
        for pm in (e for e in raiz.iter() if tag(e) == "Placemark"):
            nomes = set()
            for el in pm.iter():
                t = tag(el)
                if t == "name" and el.text:
                    nomes.add(el.text)
                elif t in ("SimpleData", "Data", "value") and el.text:
                    nomes.add(el.text)
            aneis = []
            fontes = [e for e in pm.iter() if tag(e) == "Polygon"] or [e for e in pm.iter() if tag(e) == "LinearRing"]
            for poly in fontes:
                for coords in (e for e in poly.iter() if tag(e) == "coordinates"):
                    pts = []
                    for trio in (coords.text or "").split():
                        partes = trio.split(",")
                        if len(partes) >= 2:
                            try:
                                pts.append((float(partes[0]), float(partes[1])))
                            except ValueError:
                                pass
                    if len(pts) >= 3:
                        aneis.append(pts)
            if not aneis:
                continue
            n += 1
            for nome in nomes:
                for chave in {_chave(nome), _norm(nome)} - {""}:
                    self._indice_kml.setdefault(chave, aneis)
                numeros = re.findall(r"\d+", nome)
                if numeros:  # "Talhão 104" -> 104
                    self._indice_kml.setdefault(_chave(numeros[-1]), aneis)
        print(f"  KML       ({caminho.name}): {n} polígono(s) de talhão")



# ---- decodificador mínimo de Mapbox Vector Tile (sem dependências) ----------------
def _varint(b, i):
    r = sh = 0
    while True:
        x = b[i]; i += 1
        r |= (x & 0x7F) << sh
        if x < 0x80:
            return r, i
        sh += 7


def _campos_pb(b):
    i, n = 0, len(b)
    while i < n:
        k, i = _varint(b, i)
        f, w = k >> 3, k & 7
        if w == 0:
            v, i = _varint(b, i)
        elif w == 2:
            ln, i = _varint(b, i)
            v = b[i:i + ln]; i += ln
        elif w == 1:
            v = b[i:i + 8]; i += 8
        elif w == 5:
            v = b[i:i + 4]; i += 4
        else:
            raise ValueError("wire type inválido")
        yield f, w, v


def _pacote(b):
    i, out = 0, []
    while i < len(b):
        v, i = _varint(b, i)
        out.append(v)
    return out


def _decodificar_mvt(dado, z, tx, ty):
    """Gera (atributos, [anéis exteriores em (lon,lat)]) de cada polígono do tile."""
    import math
    import struct
    for f, w, camada in _campos_pb(dado):
        if f != 3:
            continue
        chaves, valores, feats, extent = [], [], [], 4096
        for ff, ww, v in _campos_pb(camada):
            if ff == 2:
                feats.append(v)
            elif ff == 3:
                chaves.append(v.decode("utf-8"))
            elif ff == 4:
                val = None
                for f3, w3, x in _campos_pb(v):
                    if f3 == 1: val = x.decode("utf-8")
                    elif f3 == 2: val = struct.unpack("<f", x)[0]
                    elif f3 == 3: val = struct.unpack("<d", x)[0]
                    elif f3 in (4, 5): val = x
                    elif f3 == 6: val = (x >> 1) ^ -(x & 1)
                    elif f3 == 7: val = bool(x)
                valores.append(val)
            elif ff == 5:
                extent = v
        for fb in feats:
            tags, tipo, geom = [], 0, []
            for f2, w2, v in _campos_pb(fb):
                if f2 == 2: tags = _pacote(v)
                elif f2 == 3: tipo = v
                elif f2 == 4: geom = _pacote(v)
            if tipo != 3:
                continue
            attrs = {chaves[tags[k]]: valores[tags[k + 1]] for k in range(0, len(tags) - 1, 2)}
            aneis, atual, x = [], [], 0
            y, i = 0, 0
            while i < len(geom):
                cmd, cnt = geom[i] & 7, geom[i] >> 3; i += 1
                if cmd in (1, 2):
                    for _ in range(cnt):
                        dx, dy = geom[i], geom[i + 1]; i += 2
                        x += (dx >> 1) ^ -(dx & 1); y += (dy >> 1) ^ -(dy & 1)
                        atual.append((x, y))
                elif cmd == 7:
                    if len(atual) >= 3:
                        aneis.append(atual)
                    atual = []
            saida = []
            for a in aneis:
                area = sum(a[k][0] * a[(k + 1) % len(a)][1] - a[(k + 1) % len(a)][0] * a[k][1] for k in range(len(a)))
                if area <= 0:
                    continue  # buracos (sentido oposto) são ignorados
                ring = []
                for px, py in a:
                    xx = (tx + px / extent) / (1 << z)
                    yy = (ty + py / extent) / (1 << z)
                    ring.append((xx * 360.0 - 180.0, math.degrees(math.atan(math.sinh(math.pi * (1 - 2 * yy))))))
                saida.append(ring)
            if saida:
                yield attrs, saida


def _unir_pedacos(aneis):
    """Une os pedaços de um talhão cortados pelos tiles (shapely, se disponível)."""
    try:
        from shapely.geometry import Polygon
        from shapely.ops import unary_union
        u = unary_union([Polygon(a).buffer(0) for a in aneis if len(a) >= 3])
        polys = [u] if u.geom_type == "Polygon" else list(getattr(u, "geoms", []))
        return [list(p.exterior.coords) for p in polys if p.geom_type == "Polygon"]
    except Exception:
        return aneis


def carregar_cliente(pasta=PASTA_CLIENTE) -> DadosCliente:
    dados = DadosCliente()
    pasta = Path(pasta)
    if not pasta.exists():
        print("Pasta 'cliente' não encontrada — usando códigos puros e mapas sem contorno.")
        return dados
    print("Dados do cliente (pasta 'cliente'):")
    excels = [p for p in sorted(pasta.glob("*.xlsx"))
              if not p.name.startswith("~$") and not p.name.lower().startswith("metas")]
    kmls = sorted(pasta.glob("*.kml"))
    geojsons = sorted(list(pasta.glob("*.geojson")) + list(pasta.glob("*.json")))
    for yms in sorted(pasta.glob("*.yms")):
        ymx = yms.with_suffix(".ymx")
        if ymx.exists():
            dados._ler_yms_ymx(yms, ymx)
    temmapa = bool(list(pasta.glob("*.yms")))
    mbt = sorted(pasta.glob("*.mbtiles"))
    for m in mbt:
        dados._ler_mbtiles(m)
    temmapa = temmapa or bool(mbt)
    for gj in geojsons:  # primeiro: descrição/contorno dos talhões; o Excel complementa (frota, processos)
        dados._ler_geojson(gj)
    for x in excels:
        dados._ler_excel(x)
    for k in kmls:
        dados._ler_kml(k)
    if not excels and not kmls and not geojsons and not temmapa:
        print("  (nenhum .xlsx/.kml encontrado — usando códigos puros e mapas sem contorno)")
    return dados
