#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Painel Presidencial 2026 — V1.7 | Streamlit
pip install -U streamlit plotly pandas
streamlit run PAINEL_COMPARATIVO_PRESIDENCIAL_2026_V1_7.py

Lê o XLSX municipal do 1º turno (também dentro do ZIP fornecido)
e o SQLite ou CSV produzido pelo coletor municipal 2T V1.1.
Nunca grava ou modifica as bases originais.
"""
import csv
import io
import json
import re
import unicodedata
import sqlite3
import zipfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

BASE = Path(__file__).resolve().parent
DEFAULT_1T = BASE.parent / "Resultados_1turno" / "VOTOS_PRESIDENTE_MUNICIPIOS_1T.zip"
DEFAULT_2T = BASE.parent / "Resultados_2turno"
if not DEFAULT_1T.is_file():
    DEFAULT_1T = BASE / "VOTOS_PRESIDENTE_MUNICIPIOS_1T.zip"
if not DEFAULT_2T.exists():
    DEFAULT_2T = BASE / "RESULTADOS_MUNICIPAIS_2T_V1_1"
NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
HEAD_1T = {"UF", "CD_MUNICIPIO", "NM_MUNICIPIO", "SQ_CANDIDATO",
           "NR_CANDIDATO", "NM_CANDIDATO", "QT_VOTOS_1T"}
HEAD_2T = {"uf", "cd_municipio", "nm_municipio", "sq_candidato",
           "nr_candidato", "nm_candidato", "votos_2t", "concluido"}


def nint(x):
    try:
        return int(str(x).strip())
    except (ValueError, TypeError):
        return None


def celula(c, shared):
    kind = c.get("t")
    v = c.find("x:v", NS)
    if kind == "inlineStr":
        inline = c.find("x:is", NS)
        return "".join(t.text or "" for t in inline.iter()
                       if t.tag.endswith("}t")) if inline is not None else ""
    if v is None:
        return ""
    if kind == "s":
        return shared[int(v.text)]
    return v.text or ""


def xlsx_rows(content):
    """Leitura OOXML em streaming, sem exigir pandas/openpyxl."""
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared = ["".join(t.text or "" for t in item.iter()
                              if t.tag.endswith("}t")) for item in root]
        sheet = "xl/worksheets/sheet1.xml"
        with archive.open(sheet) as fp:
            header = None
            for _, element in ET.iterparse(fp, events=("end",)):
                if element.tag != f"{{{NS['x']}}}row":
                    continue
                vals = [celula(c, shared) for c in element.findall("x:c", NS)]
                if header is None:
                    header = [str(v).strip().upper() for v in vals]
                    if not HEAD_1T.issubset(set(header)):
                        raise ValueError("Cabeçalhos da primeira planilha não correspondem ao coletor 1T.")
                else:
                    yield dict(zip(header, vals))
                element.clear()


def conteudo_primeiro(path):
    if path.suffix.lower() == ".zip":
        with zipfile.ZipFile(path) as archive:
            nomes = [x for x in archive.namelist() if x.lower().endswith(".xlsx")]
            if len(nomes) != 1:
                raise ValueError("O ZIP deve conter exatamente um XLSX municipal.")
            return archive.read(nomes[0])
    if path.suffix.lower() == ".xlsx":
        return path.read_bytes()
    raise ValueError("Informe o ZIP ou XLSX municipal do primeiro turno.")


@st.cache_data(show_spinner="Lendo a base municipal do primeiro turno...")
def carregar_1t(caminho, tamanho, modificado):
    path = Path(caminho)
    votos = {}
    nomes = {}
    candidatos = {}
    municipios = set()
    for r in xlsx_rows(conteudo_primeiro(path)):
        uf = str(r["UF"]).strip().upper()
        cd = str(r["CD_MUNICIPIO"]).strip().zfill(5)
        nr = str(r["NR_CANDIDATO"]).strip()
        sq = str(r["SQ_CANDIDATO"]).strip()
        qt = nint(r["QT_VOTOS_1T"])
        if qt is None or qt < 0 or not nr or not sq:
            raise ValueError(f"Linha inválida no 1T: {uf}/{cd}/{nr}")
        chave = (uf, cd, nr)
        if chave in votos:
            raise ValueError(f"Duplicidade no 1T: {chave}")
        votos[chave] = qt
        nomes[(uf, cd)] = str(r["NM_MUNICIPIO"])
        candidatos[nr] = str(r["NM_CANDIDATO"])
        municipios.add((uf, cd))
    return votos, nomes, candidatos, municipios


def carregar_2t_sqlite(path):
    uri = path.resolve().as_uri() + "?mode=ro"
    con = sqlite3.connect(uri, uri=True, timeout=5)
    try:
        rows = con.execute("""SELECT uf,cd_municipio,nm_municipio,sq_candidato,
                             nr_candidato,nm_candidato,votos_2t,concluido,coletado_em
                             FROM votos WHERE concluido=1""").fetchall()
        # Confirma que o município também está marcado como concluído no coletor.
        validos = set(con.execute("SELECT upper(uf),cd FROM municipios WHERE concluido=1").fetchall())
    finally:
        con.close()
    return rows, validos


def carregar_2t_csv(path):
    with path.open("r", encoding="utf-8-sig", newline="") as fp:
        reader = csv.DictReader(fp, delimiter=";")
        if not HEAD_2T.issubset({c.lower() for c in reader.fieldnames or []}):
            raise ValueError("CSV do 2T sem as colunas esperadas.")
        rows = []
        for raw in reader:
            r = {k.lower(): v for k, v in raw.items()}
            rows.append((r["uf"], r["cd_municipio"], r["nm_municipio"],
                         r["sq_candidato"], r["nr_candidato"], r["nm_candidato"],
                         r["votos_2t"], r["concluido"], r.get("coletado_em", "")))
    validos = {(str(r[0]).upper(), str(r[1]).zfill(5)) for r in rows if nint(r[7]) == 1}
    return rows, validos


@st.cache_data(show_spinner=False, ttl=60)
def carregar_2t(caminho, tamanho, modificado):
    path = Path(caminho)
    rows, validos = (carregar_2t_sqlite(path) if path.suffix == ".sqlite3"
                     else carregar_2t_csv(path))
    grupos = defaultdict(dict)
    nomes = {}
    atualizacoes = []
    candidatos = {}
    for uf, cd, nome, sq, nr, nm, votos, final, coletado in rows:
        uf, cd, nr = str(uf).upper(), str(cd).zfill(5), str(nr).strip()
        chave = (uf, cd)
        if chave not in validos or nint(final) != 1:
            continue
        qt = nint(votos)
        if qt is None or qt < 0 or not nr:
            raise ValueError(f"2T inválido: {chave}/{nr}")
        if nr in grupos[chave]:
            raise ValueError(f"2T duplicado: {chave}/{nr}")
        grupos[chave][nr] = qt
        nomes[chave] = str(nome)
        candidatos[nr] = str(nm)
        if coletado:
            atualizacoes.append(str(coletado))
    # Município só é comparável se houver resultados de candidatos.
    return dict(grupos), nomes, candidatos, max(atualizacoes, default="")


def simular(votos1, municipios, candidatos1):
    """Dados artificiais apenas para validar a interface; NUNCA representam resultados."""
    selecionados = sorted(municipios)[:min(120, len(municipios))]
    numeros = sorted(candidatos1, key=lambda nr: -sum(
        votos1.get((uf, cd, nr), 0) for uf, cd in selecionados))[:2]
    return ({mun: {nr: round(votos1.get((*mun, nr), 0) * (1.06 if i == 0 else 0.94))
                   for i, nr in enumerate(numeros)} for mun in selecionados},
            {nr: candidatos1[nr] for nr in numeros}, "DADOS SIMULADOS")


def csv_bytes(linhas):
    f = io.StringIO()
    w = csv.writer(f, delimiter=";")
    w.writerow(["UF", "CD_MUNICIPIO", "NM_MUNICIPIO", "NR_CANDIDATO", "NM_CANDIDATO",
                "VOTOS_1T", "VOTOS_2T", "DIFERENCA", "VARIACAO_PERCENTUAL"])
    for r in linhas:
        w.writerow([r["uf"], r["cd"], r["municipio"], r["nr"], r["candidato"],
                    r["v1"], r["v2"], r["delta"],
                    "" if r["pct"] is None else f'{r["pct"]:.4f}'.replace(".", ",")])
    return f.getvalue().encode("utf-8-sig")



def csv_generico(registros):
    """Exporta tabelas sem dependências adicionais, com cabeçalhos estáveis."""
    f = io.StringIO()
    if registros:
        escritor = csv.DictWriter(f, fieldnames=list(registros[0]), delimiter=";")
        escritor.writeheader()
        escritor.writerows(registros)
    return f.getvalue().encode("utf-8-sig")



# MAPA V1.4: a geometria e os votos são lidos sem alterar os arquivos de origem.
IBGE_UF = {"11":"RO","12":"AC","13":"AM","14":"RR","15":"PA","16":"AP",
           "17":"TO","21":"MA","22":"PI","23":"CE","24":"RN","25":"PB",
           "26":"PE","27":"AL","28":"SE","29":"BA","31":"MG","32":"ES",
           "33":"RJ","35":"SP","41":"PR","42":"SC","43":"RS","50":"MS",
           "51":"MT","52":"GO","53":"DF"}


def normalizar_nome(nome):
    t = unicodedata.normalize("NFKD", str(nome or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]+", " ", t.upper()).strip()


def codigo_ibge_feature(feat):
    p = feat.get("properties") or {}
    for k in ("CD_MUN", "CD_MUN_2022", "CD_GEOCMU", "codarea", "codigo_ibge",
              "cod_ibge", "id", "code", "geocodigo", "codigo"):
        v = p.get(k)
        if v is not None and re.fullmatch(r"\d{7}", str(v).strip()):
            return str(v).strip()
    v = feat.get("id")
    return str(v).strip() if v is not None and re.fullmatch(r"\d{7}", str(v).strip()) else None


def nome_feature(feat):
    p = feat.get("properties") or {}
    for k in ("NM_MUN", "NM_MUNICIP", "name", "nome", "NOME", "municipio", "description"):
        if p.get(k):
            return str(p[k])
    return ""


@st.cache_data(show_spinner="Carregando malha geográfica municipal...")
def carregar_malha(caminho, tamanho, mtime):
    with Path(caminho).open("r", encoding="utf-8-sig") as f:
        geo = json.load(f)
    if geo.get("type") != "FeatureCollection" or not isinstance(geo.get("features"), list):
        raise ValueError("O arquivo deve ser um GeoJSON FeatureCollection de municípios.")
    por_nome = defaultdict(list)
    ids = set()
    features = []
    for feat in geo["features"]:
        cod = codigo_ibge_feature(feat)
        if not cod or cod[:2] not in IBGE_UF or cod in ids:
            continue
        ids.add(cod)
        nome = nome_feature(feat)
        uf = IBGE_UF[cod[:2]]
        if nome:
            por_nome[(uf, normalizar_nome(nome))].append(cod)
        feat["id"] = cod
        features.append(feat)
    if not features:
        raise ValueError("Nenhum polígono com código IBGE de 7 dígitos encontrado no GeoJSON.")
    return {"type": "FeatureCollection", "features": features}, dict(por_nome), ids


@st.cache_data(show_spinner=False)
def carregar_depara(caminho, tamanho, mtime):
    # CSV: UF;CD_MUNICIPIO_TSE;CD_MUNICIPIO_IBGE
    with Path(caminho).open("r", encoding="utf-8-sig", newline="") as f:
        amostra = f.read(4096)
        f.seek(0)
        delimitador = ";" if amostra.count(";") >= amostra.count(",") else ","
        reader = csv.DictReader(f, delimiter=delimitador)
        cols = {str(c).strip().upper(): c for c in (reader.fieldnames or [])}
        necessarias = ("UF", "CD_MUNICIPIO_TSE", "CD_MUNICIPIO_IBGE")
        if any(k not in cols for k in necessarias):
            raise ValueError("De-para requer UF;CD_MUNICIPIO_TSE;CD_MUNICIPIO_IBGE")
        resultado = {}
        for row in reader:
            uf = str(row[cols["UF"]]).strip().upper()
            tse = str(row[cols["CD_MUNICIPIO_TSE"]]).strip().zfill(5)
            ibge = str(row[cols["CD_MUNICIPIO_IBGE"]]).strip()
            chave = (uf, tse)
            if not re.fullmatch(r"\d{7}", ibge) or IBGE_UF.get(ibge[:2]) != uf:
                raise ValueError(f"Código IBGE incompatível com UF: {uf}/{tse}/{ibge}")
            if chave in resultado and resultado[chave] != ibge:
                raise ValueError(f"Correspondência conflitante: {chave}")
            resultado[chave] = ibge
    return resultado


def resolver_geografia(municipios, nomes, por_nome, ids, depara):
    associados, pendentes, ambiguos = {}, [], []
    usados = {}
    for m in sorted(municipios):
        ibge = depara.get(m)
        if ibge is None:
            possiveis = por_nome.get((m[0], normalizar_nome(nomes.get(m, ""))), [])
            if len(possiveis) == 1:
                ibge = possiveis[0]
            elif len(possiveis) > 1:
                ambiguos.append((m, nomes.get(m, ""), "Nome geográfico ambíguo"))
                continue
        if ibge not in ids:
            pendentes.append((m, nomes.get(m, ""), "Sem polígono correspondente"))
            continue
        if ibge in usados and usados[ibge] != m:
            ambiguos.append((m, nomes.get(m, ""), "Polígono atribuído a outro município"))
            continue
        usados[ibge] = m
        associados[m] = ibge
    return associados, pendentes, ambiguos


REGIOES = {
    "Norte": ["AC", "AM", "AP", "PA", "RO", "RR", "TO"],
    "Nordeste": ["AL", "BA", "CE", "MA", "PB", "PE", "PI", "RN", "SE"],
    "Centro-Oeste": ["DF", "GO", "MS", "MT"],
    "Sudeste": ["ES", "MG", "RJ", "SP"],
    "Sul": ["PR", "RS", "SC"],
}


def recortar_geojson(geo, codigos):
    # Envia ao navegador apenas os polígonos utilizados no recorte.
    return {"type": "FeatureCollection", "features": [
        feat for feat in geo["features"] if feat.get("id") in codigos
    ]}


def enquadramento_geojson(geo):
    # BBox calculado dos vértices do recorte; evita coordenadas centrais fixas.
    lons, lats = [], []
    def percorrer(x):
        if not isinstance(x, list) or not x:
            return
        if isinstance(x[0], (int, float)) and len(x) >= 2:
            lons.append(float(x[0])); lats.append(float(x[1]))
        else:
            for item in x:
                percorrer(item)
    for feat in geo["features"]:
        percorrer((feat.get("geometry") or {}).get("coordinates", []))
    if not lons:
        return {"lat": -14.5, "lon": -52.5}, 3
    x0, x1, y0, y1 = min(lons), max(lons), min(lats), max(lats)
    centro = {"lat": (y0 + y1) / 2, "lon": (x0 + x1) / 2}
    # Aproximação conservadora para viewport; o usuário pode ajustar o zoom.
    amplitude = max((x1-x0)*0.8, (y1-y0)*1.1, 0.3)
    import math
    zoom = max(2.3, min(10.5, math.log2(300 / amplitude)))
    return centro, zoom


def escala_robusta(dados, metrica):
    import pandas as pd
    chave = "Diferença" if metrica == "Votos absolutos" else "Variação (%)"
    valores = pd.to_numeric(pd.Series([
        abs(r[chave]) for rr in dados.values() for r in rr if r[chave] is not None
    ], dtype="float64"), errors="coerce").dropna()
    if valores.empty:
        return 1.0
    # A escala robusta satura extremos apenas na cor; hover e CSV preservam valores.
    return max(float(valores.quantile(0.95)), 1.0 if metrica == "Votos absolutos" else 0.01)


def figura_municipal(geo, registros, titulo, metrica, escala, altura=620):
    if not registros:
        return None
    import pandas as pd
    df = pd.DataFrame(registros)
    coluna = "Diferença" if metrica == "Votos absolutos" else "Variação (%)"
    df[coluna] = pd.to_numeric(df[coluna], errors="coerce")
    codigos = set(df["IBGE"])
    geo_local = recortar_geojson(geo, codigos)
    centro, zoom = enquadramento_geojson(geo_local)
    fig = px.choropleth_map(
        df, geojson=geo_local, locations="IBGE", featureidkey="id", color=coluna,
        color_continuous_scale=["#b91c1c", "#f87171", "#f3f4f6", "#4ade80", "#166534"],
        range_color=(-escala, escala),
        hover_name="Município", custom_data=["UF", "1º turno", "2º turno", "Diferença", "Variação (%)"],
        center=centro, zoom=zoom, map_style="carto-positron", opacity=0.94,
        labels={coluna: "Votos" if metrica == "Votos absolutos" else "%"},
    )
    fig.update_traces(marker_line_width=0.15, marker_line_color="rgba(65,65,65,0.55)",
        hovertemplate=("<b>%{hovertext}</b> (%{customdata[0]})<br>"
                       "1º turno: %{customdata[1]:,.0f}<br>"
                       "2º turno: %{customdata[2]:,.0f}<br>"
                       "Diferença: %{customdata[3]:+,.0f}<br>"
                       "Variação: %{customdata[4]:+.2f}%<extra></extra>"))
    fig.update_layout(title=titulo, height=altura, margin=dict(l=0,r=0,t=42,b=0),
                      coloraxis_colorbar=dict(title="Votos" if metrica == "Votos absolutos" else "%"),
                      uirevision=None)
    return fig


st.set_page_config(page_title="Presidencial 2026 | Comparativo Municipal", page_icon="🗳️",
                   layout="wide")
st.title("🗳️ Eleições 2026 — Comparativo municipal presidencial")
st.caption("1º turno × 2º turno · Mesmos municípios concluídos · Fonte: bases eleitorais publicadas para consulta")

# Publicação pública: fontes fixas no repositório; sem acesso arbitrário a arquivos do servidor.
with st.sidebar:
    st.header("Fontes públicas")
    caminho1 = str(DEFAULT_1T)
    pasta2 = str(DEFAULT_2T)
    demonstracao = False
    caminho_geo = str(BASE.parent / "Mapas_IBGE" / "municipios_br_leve.geojson")
    caminho_depara = str(BASE.parent / "Mapas_IBGE" / "depara_tse_ibge_v15.csv")
    st.caption("Resultados eleitorais e malha geográfica publicados com esta aplicação.")
    st.markdown("**Atualização:** use o botão abaixo ou recarregue a página.")
    if st.button("🔄 Atualizar agora"):
        st.cache_data.clear()
        st.rerun()
    st.caption("Painel de consulta. A coleta dos dados é feita separadamente.")

primeiro = Path(caminho1).expanduser()
if not primeiro.is_file():
    st.error(f"Base do 1º turno não encontrada: {primeiro}")
    st.stop()

try:
    stat1 = primeiro.stat()
    v1, nomes1, cand1, mun1 = carregar_1t(str(primeiro.resolve()),
                                          stat1.st_size, stat1.st_mtime_ns)
except Exception as exc:
    st.error(f"Não foi possível ler a base do primeiro turno: {exc}")
    st.stop()

@st.fragment(run_every="300s")
def mostrar_painel():
    if demonstracao:
        grupos2, cand2, atualizado = simular(v1, mun1, cand1)
        st.error("⚠️ MODO DEMONSTRAÇÃO — VOTOS FICTÍCIOS, NÃO PUBLICAR COMO RESULTADOS.")
    else:
        p2 = Path(pasta2).expanduser()
        arquivo = (p2 / "RESULTADOS_MUNICIPAIS_2T_V1_1.sqlite3" if p2.is_dir()
                   else p2)
        if not arquivo.is_file() and p2.is_dir():
            candidatos_db = list(p2.rglob("RESULTADOS_MUNICIPAIS_2T_V1_1.sqlite3"))
            if len(candidatos_db) == 1:
                arquivo = candidatos_db[0]
            else:
                arquivo = p2 / "VOTOS_PRESIDENTE_MUNICIPIOS_2T.csv"
        if not arquivo.is_file():
            st.info("Aguardando a primeira coleta oficial do segundo turno. "
                    "Pode ativar o modo demonstração para testar o painel.")
            st.metric("Municípios concluídos e comparáveis", 0)
            st.stop()
        try:
            stat2 = arquivo.stat()
            grupos2, _, cand2, atualizado = carregar_2t(
                str(arquivo.resolve()), stat2.st_size, stat2.st_mtime_ns)
        except Exception as exc:
            st.error(f"Não foi possível ler a base do segundo turno: {exc}")
            st.stop()

    inexistentes = set(grupos2) - mun1
    if inexistentes:
        st.warning(f"{len(inexistentes)} localidades do 2º turno não constam da base nacional "
                   "do 1º turno e foram excluídas do comparativo.")
    grupos2 = {m: r for m, r in grupos2.items() if m in mun1}
    if not grupos2:
        st.info("Ainda não existem municípios concluídos com votos válidos para comparar.")
        st.stop()

    uf_opts = sorted({m[0] for m in grupos2})
    st.caption("Filtros gerais das abas Desempenho individual e Tabela municipal completa. A aba Dois candidatos possui filtros próprios.")
    uf_sel = st.selectbox("Abrangência (abas 1 e 3)", ["Brasil"] + uf_opts)
    municipios_uf = {m: votos for m, votos in grupos2.items()
                     if uf_sel == "Brasil" or m[0] == uf_sel}
    municipios_opcoes = sorted(municipios_uf, key=lambda m: (nomes1.get(m, ""), m[1]))
    rotulos = {m: f"{nomes1.get(m, m[1])} ({m[0]} / {m[1]})" for m in municipios_opcoes}
    opcao = st.selectbox("Município (abas 1 e 3)", ["Todos os municípios concluídos"] +
                         [rotulos[m] for m in municipios_opcoes])
    selecionados = (municipios_uf if opcao == "Todos os municípios concluídos"
                    else {next(m for m in municipios_opcoes if rotulos[m] == opcao):
                          municipios_uf[next(m for m in municipios_opcoes if rotulos[m] == opcao)]})

    nums = sorted({nr for votos in selecionados.values() for nr in votos},
                  key=lambda n: nint(n) if nint(n) is not None else 9999)
    if not nums:
        st.warning("Nenhum candidato do segundo turno encontrado.")
        st.stop()

    linhas = []
    agregado = defaultdict(lambda: [0, 0])
    faltantes = set()
    for (uf, cd), votos2 in selecionados.items():
        for nr in nums:
            if nr not in votos2:
                faltantes.add((uf, cd, nr))
                continue
            chave = (uf, cd, nr)
            if chave not in v1:
                faltantes.add(chave)
                continue
            a, b = v1[chave], votos2[nr]
            pct = (b - a) / a * 100 if a else None
            linhas.append(dict(uf=uf, cd=cd, municipio=nomes1.get((uf, cd), cd),
                               nr=nr, candidato=cand2.get(nr, cand1.get(nr, nr)),
                               v1=a, v2=b, delta=b - a, pct=pct))
            agregado[nr][0] += a
            agregado[nr][1] += b

    if faltantes:
        st.warning(f"{len(faltantes)} combinações município/candidato sem correspondência "
                   "entre os dois turnos. Elas foram excluídas dos cálculos. "
                   "Verifique se os números dos candidatos permanecem consistentes.")
    if not linhas:
        st.error("Nenhuma correspondência válida entre as bases.")
        st.stop()

    st.caption(f"Última coleta do 2º turno: {atualizado or 'não informada'} · "
               f"Localidades selecionadas: {len(selecionados):,} · "
               f"Localidades nacionais no 1º turno: {len(mun1):,}")
    if atualizado and not demonstracao:
        try:
            idade = (datetime.now().astimezone() -
                     datetime.fromisoformat(atualizado)).total_seconds()
            if idade > 900:
                st.warning(f"⚠️ Dados possivelmente desatualizados: última coleta há {idade/60:.0f} minutos.")
        except ValueError:
            pass

    c1, c2, c3 = st.columns(3)
    c1.metric("Municípios concluídos no recorte", f"{len(selecionados):,}".replace(",", "."))
    c2.metric("Cobertura nacional", f"{len(grupos2)/len(mun1)*100:.2f}%".replace(".", ","))
    c3.metric("Candidatos com correspondência", len(agregado))

    def fmt(n):
        return f"{n:,}".replace(",", ".")

    def candidato_nome(n):
        return f"{n} — {cand2.get(n, cand1.get(n, n))}"

    def resumir(n):
        a, b = agregado.get(n, (0, 0))
        return {"Candidato": cand2.get(n, cand1.get(n, n)),
                "Número": n, "1º turno": a, "2º turno": b,
                "Diferença": b-a,
                "Variação (%)": round((b-a)/a*100, 2) if a else None}

    aba1, aba2, aba3, aba4 = st.tabs([
        "📊 Desempenho individual",
        "⚖️ Comparativo — dois candidatos",
        "📋 Tabela municipal completa",
        "🗺️ Mapa de evolução dos votos"
    ])

    with aba1:
        st.subheader("Votos por candidato — mesmos municípios")
        st.dataframe([resumir(n) for n in nums if n in agregado],
                     hide_index=True, use_container_width=True)
        escolha = st.selectbox("Candidato", nums, format_func=candidato_nome,
                               key="candidato_individual")
        a, b = agregado.get(escolha, (0, 0))
        x, y, z = st.columns(3)
        x.metric("1º turno", fmt(a))
        y.metric("2º turno", fmt(b))
        z.metric("Variação", f"{b-a:+,}".replace(",", "."),
                 delta=f"{(b-a)/a*100:+.2f}%" if a else "Base zero")
        st.bar_chart({"1º turno": [a], "2º turno": [b]})
        st.subheader("Detalhamento municipal")
        detalhe = [r for r in linhas if r["nr"] == escolha]
        st.dataframe([{"UF": r["uf"], "Município": r["municipio"],
                       "Código TSE": r["cd"], "1º turno": r["v1"],
                       "2º turno": r["v2"], "Diferença": r["delta"],
                       "Variação (%)": round(r["pct"], 2) if r["pct"] is not None else None}
                      for r in sorted(detalhe, key=lambda r: (r["uf"], r["municipio"]))],
                     hide_index=True, use_container_width=True)
        st.download_button("⬇️ Exportar candidato (CSV)",
                           data=csv_bytes(detalhe),
                           file_name=f"comparativo_2026_{escolha}_{uf_sel.lower()}.csv",
                           mime="text/csv", key="exportar_individual")

    with aba2:
        st.subheader("Dois candidatos — comparação simples")
        st.caption("Escolha dois candidatos e consulte os votos lado a lado. "
                   "Os mesmos municípios são usados nos dois turnos.")
        # Esta aba trabalha sobre a base nacional completa, independentemente
        # dos filtros gerais aplicados às abas 1 e 3.
        nums_dupla = sorted(
            {nr for votos in grupos2.values() for nr in votos},
            key=lambda n: nint(n) if nint(n) is not None else 9999)
        if len(nums_dupla) < 2:
            st.info("São necessários dois candidatos do segundo turno para esta análise.")
        else:
            col_a, col_b = st.columns(2)
            with col_a:
                candidato_a = st.selectbox("Candidato A", nums_dupla, index=0,
                                           format_func=candidato_nome, key="comp_a")
            with col_b:
                candidato_b = st.selectbox("Candidato B", nums_dupla, index=1,
                                           format_func=candidato_nome, key="comp_b")
            if candidato_a == candidato_b:
                st.warning("Escolha dois candidatos diferentes.")
            else:
                st.markdown("#### Escolha a abrangência")
                ufs_dupla = sorted({m[0] for m in grupos2})
                col_uf, col_mun = st.columns(2)
                with col_uf:
                    uf_dupla = st.selectbox(
                        "UF para comparação", ["Brasil"] + ufs_dupla,
                        key="uf_comparativo")
                # A chave do município depende da UF, de modo que a escolha
                # anterior seja descartada automaticamente ao trocar de estado.
                municipios_disponiveis = sorted(
                    (m for m in grupos2 if uf_dupla == "Brasil" or m[0] == uf_dupla),
                    key=lambda m: (m[0], nomes1.get(m, ""), m[1]))
                labels = {
                    m: f"{nomes1.get(m, m[1])} ({m[0]} · {m[1]})"
                    for m in municipios_disponiveis
                }
                with col_mun:
                    escolha_mun = st.selectbox(
                        "Município (opcional)",
                        [None] + municipios_disponiveis,
                        format_func=lambda m: (
                            "Todos — agregado do Brasil" if m is None and uf_dupla == "Brasil"
                            else "Todos — agregado da UF" if m is None
                            else labels[m]),
                        key=f"municipio_comparativo_{uf_dupla}")
                chaves_recorte = (
                    municipios_disponiveis if escolha_mun is None else [escolha_mun])
                pares = []
                excluidos = 0
                for m in chaves_recorte:
                    votos2 = grupos2[m]
                    chave_a = (m[0], m[1], candidato_a)
                    chave_b = (m[0], m[1], candidato_b)
                    if (candidato_a not in votos2 or candidato_b not in votos2
                            or chave_a not in v1 or chave_b not in v1):
                        excluidos += 1
                        continue
                    pares.append((
                        m,
                        {"v1": v1[chave_a], "v2": votos2[candidato_a],
                         "municipio": nomes1.get(m, m[1])},
                        {"v1": v1[chave_b], "v2": votos2[candidato_b],
                         "municipio": nomes1.get(m, m[1])},
                    ))
                if excluidos:
                    st.warning(
                        f"{excluidos} municípios do recorte não têm os quatro valores "
                        "necessários e foram excluídos da comparação.")
                if not pares:
                    st.info("Não há municípios comparáveis para ambos os candidatos neste recorte.")
                else:
                    a1 = sum(a["v1"] for _, a, _ in pares)
                    a2 = sum(a["v2"] for _, a, _ in pares)
                    b1 = sum(b["v1"] for _, _, b in pares)
                    b2 = sum(b["v2"] for _, _, b in pares)
                    titulo = (
                        f"Município: {labels[escolha_mun]}" if escolha_mun is not None
                        else "Agregado nacional — municípios concluídos" if uf_dupla == "Brasil"
                        else f"Agregado da UF {uf_dupla} — municípios concluídos")
                    st.markdown(f"**{titulo}**")
                    st.caption(
                        f"Municípios comparados: {fmt(len(pares))} "
                        f"de {fmt(len(chaves_recorte))} no recorte"
                        + (" · MODO DEMONSTRAÇÃO: dados fictícios"
                           if demonstracao else ""))
                    st.markdown("#### Votos nos dois turnos")
                    st.dataframe([
                        {"Candidato": candidato_nome(candidato_a),
                         "1º turno": a1, "2º turno": a2,
                         "Variação de votos": a2 - a1},
                        {"Candidato": candidato_nome(candidato_b),
                         "1º turno": b1, "2º turno": b2,
                         "Variação de votos": b2 - b1},
                    ], hide_index=True, use_container_width=True)
                    st.bar_chart(
                        [
                            {"Candidato": f"A · {candidato_a}",
                             "1º turno": a1, "2º turno": a2},
                            {"Candidato": f"B · {candidato_b}",
                             "1º turno": b1, "2º turno": b2},
                        ],
                        x="Candidato", y=["1º turno", "2º turno"],
                        stack=False, use_container_width=True)
                    st.caption(
                        "Barras lado a lado: primeiro e segundo turnos, "
                        "no mesmo conjunto de municípios.")
                    st.markdown("#### Municípios — visão resumida")
                    resumo_municipal = [{
                        "UF": m[0], "Código TSE": m[1],
                        "Município": a["municipio"],
                        "A · 1º turno": a["v1"], "A · 2º turno": a["v2"],
                        "B · 1º turno": b["v1"], "B · 2º turno": b["v2"]
                    } for m, a, b in pares]
                    st.dataframe(
                        resumo_municipal, hide_index=True,
                        use_container_width=True)
                    st.download_button(
                        "⬇️ Exportar comparação municipal (CSV)",
                        data=csv_generico(resumo_municipal),
                        file_name=(
                            f"comparacao_2026_{candidato_a}_{candidato_b}_"
                            f"{uf_dupla.lower()}.csv"),
                        mime="text/csv", key="exportar_dupla")
                    with st.expander("Ver indicadores adicionais"):
                        def variacao(v1_, v2_):
                            return (
                                f"{(v2_-v1_)/v1_*100:+.2f}%"
                                if v1_ else "Indefinida")
                        st.write(f"Variação percentual de A: {variacao(a1, a2)}")
                        st.write(f"Variação percentual de B: {variacao(b1, b2)}")
                        st.write(f"Diferença de votos A − B no 1º turno: {fmt(a1-b1)}")
                        st.write(f"Diferença de votos A − B no 2º turno: {fmt(a2-b2)}")
                        st.caption(
                            "Indicadores descritivos. Não representam "
                            "transferência individual de votos.")

    with aba3:
        st.subheader("Tabela municipal completa")
        st.caption("Uma linha por município e candidato, limitada aos municípios "
                   "concluídos no segundo turno e selecionados nos filtros.")
        tabela = [{"UF": r["uf"], "Código TSE": r["cd"],
                   "Município": r["municipio"], "Número": r["nr"],
                   "Candidato": r["candidato"], "Votos 1T": r["v1"],
                   "Votos 2T": r["v2"], "Diferença": r["delta"],
                   "Variação (%)": round(r["pct"], 2) if r["pct"] is not None else None}
                  for r in sorted(linhas, key=lambda r:
                                  (r["uf"], r["municipio"], r["nr"]))]
        st.dataframe(tabela, hide_index=True, use_container_width=True)
        st.download_button("⬇️ Exportar tabela municipal completa (CSV)",
                           data=csv_generico(tabela),
                           file_name=f"municipios_2026_{uf_sel.lower()}.csv",
                           mime="text/csv", key="exportar_tabela")


    with aba4:
        st.subheader("Mapa municipal — variação entre 1º e 2º turnos")
        st.caption("Verde: aumento de votos; vermelho: redução; cinza: variação próxima de zero. "
                   "O mapa representa a diferença de votos de cada candidato, não o vencedor do município.")
        arquivo_geo = Path(caminho_geo).expanduser()
        if not arquivo_geo.is_file():
            st.warning("Para habilitar o mapa, indique na barra lateral um GeoJSON municipal "
                       "com códigos IBGE de 7 dígitos e nomes dos municípios. "
                       "As três abas anteriores continuam funcionando normalmente.")
        else:
            try:
                stat_geo = arquivo_geo.stat()
                geo, indice_nomes, ids_geo = carregar_malha(str(arquivo_geo.resolve()),
                                                              stat_geo.st_size, stat_geo.st_mtime_ns)
                arquivo_depara = Path(caminho_depara).expanduser()
                depara = {}
                if arquivo_depara.is_file():
                    stat_dp = arquivo_depara.stat()
                    depara = carregar_depara(str(arquivo_depara.resolve()), stat_dp.st_size, stat_dp.st_mtime_ns)
                associacoes, pendentes, ambiguos = resolver_geografia(
                    set(grupos2), nomes1, indice_nomes, ids_geo, depara)
                if pendentes or ambiguos:
                    st.warning(f"Municípios sem correspondência geográfica segura: "
                               f"{len(pendentes)} não encontrados e {len(ambiguos)} ambíguos. "
                               "Não são desenhados até que a correspondência seja validada.")
                    erros = [{"UF": m[0], "CD_MUNICIPIO_TSE": m[1], "NM_MUNICIPIO": nm,
                              "MOTIVO": motivo} for m, nm, motivo in pendentes + ambiguos]
                    st.download_button("Exportar pendências de correspondência (CSV)",
                                       csv_generico(erros), file_name="pendencias_malha_v14.csv",
                                       mime="text/csv")
                if depara:
                    aplicados = sum(1 for chave, cod in depara.items() if associacoes.get(chave) == cod)
                    st.caption(f"De-para carregado: {len(depara)} registros; {aplicados} aplicados ao conjunto eleitoral.")
                    if aplicados != len(depara):
                        st.warning("Alguns vínculos do de-para não foram aplicados. Confira UF/código TSE e arquivo CSV.")
                    st.caption("Os vínculos do de-para são propostos e devem ser validados contra tabela oficial TSE–IBGE.")
                else:
                    st.warning("De-para TSE–IBGE não carregado. Confira o caminho do CSV na barra lateral.")
                st.caption(f"Polígonos municipais na malha: {len(ids_geo):,} · "
                           f"Municípios vinculados aos votos: {len(associacoes):,} de {len(grupos2):,}")
                st.markdown("**Filtros exclusivos do mapa**")
                cmodo, cmetrica, creg, cuf = st.columns(4)
                with cmodo:
                    modo_mapa = st.radio("Visualização", ["Individual", "Lado a lado"], horizontal=True)
                with cmetrica:
                    metrica_mapa = st.selectbox("Indicador", ["Votos absolutos", "Variação percentual"])
                with creg:
                    regiao_mapa = st.selectbox("Região", ["Brasil"] + list(REGIOES), key="regiao_mapa_v17")
                with cuf:
                    ufs_regiao = sorted({m[0] for m in grupos2} if regiao_mapa == "Brasil"
                                        else REGIOES[regiao_mapa])
                    uf_mapa = st.selectbox("UF", ["Todas"] + ufs_regiao, key="uf_mapa_v17")
                ufs_filtro = ({uf_mapa} if uf_mapa != "Todas" else
                              set(ufs_regiao))
                st.caption("Recorte: " + (uf_mapa if uf_mapa != "Todas" else regiao_mapa) +
                           " · O mapa é enquadrado automaticamente e permite zoom e consulta ao passar o mouse.")
                numeros_mapa = sorted({nr for grupo in grupos2.values() for nr in grupo},
                                      key=lambda x: nint(x) or 9999)
                if not numeros_mapa:
                    st.info("Nenhum candidato encontrado para o mapa.")
                else:
                    candidatos_escolhidos = []
                    if modo_mapa == "Individual":
                        candidatos_escolhidos = [st.selectbox("Candidato", numeros_mapa,
                                                             format_func=candidato_nome, key="mapa_individual")]
                    else:
                        ca, cb = st.columns(2)
                        with ca:
                            a = st.selectbox("Candidato A", numeros_mapa, index=0,
                                             format_func=candidato_nome, key="mapa_a")
                        with cb:
                            b = st.selectbox("Candidato B", numeros_mapa,
                                             index=min(1, len(numeros_mapa)-1),
                                             format_func=candidato_nome, key="mapa_b")
                        candidatos_escolhidos = [a, b]
                    if len(set(candidatos_escolhidos)) != len(candidatos_escolhidos):
                        st.warning("Escolha dois candidatos diferentes.")
                    else:
                        dados_mapa = {}
                        for nr in candidatos_escolhidos:
                            registros = []
                            for m, ibge in associacoes.items():
                                if m[0] not in ufs_filtro:
                                    continue
                                if nr not in grupos2[m] or (*m, nr) not in v1:
                                    continue
                                a, b = v1[(*m, nr)], grupos2[m][nr]
                                registros.append({"IBGE": ibge, "UF": m[0],
                                                  "Município": nomes1.get(m, m[1]),
                                                  "1º turno": a, "2º turno": b,
                                                  "Diferença": b-a,
                                                  "Variação (%)": (b-a)/a*100 if a else None})
                            dados_mapa[nr] = registros
                        escala = escala_robusta(dados_mapa, metrica_mapa)
                        st.caption(f"Escala cromática simétrica: ±{escala:,.2f} " +
                                   ("votos" if metrica_mapa == "Votos absolutos" else "%") +
                                   " (percentil 95 das magnitudes no recorte; valores extremos preservados no hover e CSV).")
                        colunas = st.columns(len(candidatos_escolhidos))
                        for col, nr in zip(colunas, candidatos_escolhidos):
                            with col:
                                rr = dados_mapa[nr]
                                ganhos = sum(r["Diferença"] > 0 for r in rr)
                                perdas = sum(r["Diferença"] < 0 for r in rr)
                                iguais = sum(r["Diferença"] == 0 for r in rr)
                                st.markdown(f"**{candidato_nome(nr)}**")
                                st.caption(f"{len(rr):,} municípios · {ganhos:,} com aumento · "
                                           f"{perdas:,} com redução · {iguais:,} estáveis")
                                if rr:
                                    fig = figura_municipal(geo, rr, "Variação de votos por município",
                                                          metrica_mapa, escala,
                                                          altura=640 if modo_mapa == "Individual" else 540)
                                    st.plotly_chart(fig, use_container_width=True,
                                                    key=f"fig_mapa_{nr}_{modo_mapa}")
                                    st.download_button("Exportar dados do mapa (CSV)",
                                                       csv_generico(rr),
                                                       file_name=f"mapa_2026_{nr}_{regiao_mapa}_{uf_mapa}.csv",
                                                       mime="text/csv", key=f"csv_mapa_{nr}")
                                else:
                                    st.info("Sem municípios comparáveis com geometria neste recorte.")
                        st.caption("Percentuais com base zero são indefinidos e ficam sem cor na "
                                   "escala percentual. Ambos os mapas usam a mesma escala de cores. "
                                   "Correspondência automática apenas por nome normalizado único na UF; "
                                   "casos não resolvidos exigem de-para validado.")
            except Exception as exc:
                st.error(f"Falha ao preparar mapa municipal: {exc}")

    st.caption("Metodologia: diferença = votos 2T − votos 1T; variação = diferença / votos 1T. "
               "Se votos 1T = 0, a variação percentual é indefinida. "
               "Candidatos identificados pelo número de urna e municípios por UF/código TSE. "
               "Os números de urna devem ser conferidos com os registros oficiais antes "
               "de publicar resultados. Não se inferem transferências individuais de votos.")


mostrar_painel()
