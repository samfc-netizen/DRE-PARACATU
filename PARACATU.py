# PARACATU.py
# Dashboard Streamlit (DRE + DFC + Indicador de Compras + Indicadores Comerciais)
#
# Abas esperadas no Excel:
# - RECEITA E CMV   (colunas mínimas: DATA, VR.TOTAL, CUSTO; demais colunas usadas nos indicadores comerciais: CLIENTE, SEGMENTO, MARCA, LINHA, etc.)
# - DRE             (colunas mínimas: DTA.PAG, CONTA DE RESULTADO, VAL.PAG; e para drill: DESPESA, FAVORECIDO, HISTÓRICO, DUPLICATA)
# - RECEBIMENTOS    (colunas mínimas: MÊS, ANO, VALOR)
# - Compras fornecedor (colunas mínimas: DATA, FORNECEDOR, VR. CONTÁBIL)
#
# Requisitos (requirements.txt):
# streamlit
# pandas
# openpyxl
# plotly

import os
import re
import glob
import unicodedata
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px


# =========================
# Constantes / Mês
# =========================
MESES_PT = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
MES_NUM_TO_PT = {1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN",
                 7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"}
MES_PT_TO_NUM = {v: k for k, v in MES_NUM_TO_PT.items()}
MES_LONG_TO_NUM = {
    "JANEIRO": 1, "JAN": 1,
    "FEVEREIRO": 2, "FEV": 2,
    "MARCO": 3, "MARÇO": 3, "MAR": 3,
    "ABRIL": 4, "ABR": 4,
    "MAIO": 5, "MAI": 5,
    "JUNHO": 6, "JUN": 6,
    "JULHO": 7, "JUL": 7,
    "AGOSTO": 8, "AGO": 8,
    "SETEMBRO": 9, "SET": 9,
    "OUTUBRO": 10, "OUT": 10,
    "NOVEMBRO": 11, "NOV": 11,
    "DEZEMBRO": 12, "DEZ": 12,
}
MESES_FULL = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
    7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}
MESES_FULL_INV = {v: k for k, v in MESES_FULL.items()}


# =========================
# Helpers
# =========================
def to_num(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip()
    if s == "":
        return 0.0
    s = s.replace("\u00a0", " ").replace("R$", "").strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def format_brl(x) -> str:
    try:
        return f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00"


def fmt_pct(x) -> str:
    try:
        return f"{float(x):,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00%"


def parse_mes(v) -> Optional[int]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        m = int(v)
        return m if 1 <= m <= 12 else None
    s = str(v).strip().upper()
    if s.isdigit():
        m = int(s)
        return m if 1 <= m <= 12 else None
    s_norm = unicodedata.normalize("NFKD", s)
    s_norm = "".join(ch for ch in s_norm if not unicodedata.combining(ch))
    s_norm = s_norm.replace(".", "").strip()
    return MES_LONG_TO_NUM.get(s_norm)


def sintetizar_despesa(nome: str) -> str:
    if nome is None or (isinstance(nome, float) and pd.isna(nome)):
        return "—"
    s = str(nome).strip()
    s = re.sub(r"\s*\(\s*\d+\s*-\s*DESPESAS\s*\)\s*$", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s*\([^)]*\)\s*$", "", s).strip()
    s = re.sub(r"\s{2,}", " ", s)
    return s if s else "—"


def excel_signature(path: str) -> Tuple[int, int]:
    stt = os.stat(path)
    return (stt.st_mtime_ns, stt.st_size)


def _auto_find_excel() -> Optional[str]:
    preferred = ["projeto Paracatu.xlsx", "PROJETO PARACATU.xlsx", "Paracatu.xlsx", "PARACATU.xlsx"]
    for fn in preferred:
        if os.path.exists(fn):
            return fn
    files = []
    for pat in ["*.xlsx", "*.xlsm", "*.xls"]:
        files.extend(glob.glob(pat))
    files = [f for f in files if os.path.isfile(f)]
    if not files:
        return None
    files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return files[0]


@st.cache_data(show_spinner=False)
def read_sheet(excel_path: str, sheet_name: str, sig: Tuple[int, int]) -> Optional[pd.DataFrame]:
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    except Exception:
        return None
    df.columns = [str(c).strip() for c in df.columns]
    return df


@st.cache_data(show_spinner=False)
def prep_receita_cmv(excel_path: str, sig: Tuple[int, int]) -> Optional[pd.DataFrame]:
    df = read_sheet(excel_path, "RECEITA E CMV", sig)
    if df is None:
        return None
    r = df.copy()
    r["_dt"] = pd.to_datetime(r.get("DATA"), errors="coerce", dayfirst=True)
    r = r[r["_dt"].notna()].copy()
    r["_ano"] = r["_dt"].dt.year
    r["_mes"] = r["_dt"].dt.month
    r["_receita"] = r.get("VR.TOTAL").apply(to_num) if "VR.TOTAL" in r.columns else 0.0
    r["_cmv"] = r.get("CUSTO").apply(to_num) if "CUSTO" in r.columns else 0.0

    # normaliza texto das colunas comerciais se existirem
    for c in ["CLIENTE", "SEGMENTO", "MARCA", "LINHA", "VENDEDOR", "CIDADE"]:
        if c in r.columns:
            r[c] = r[c].astype(str).str.strip()
    return r


@st.cache_data(show_spinner=False)
def prep_dre_lancamentos(excel_path: str, sig: Tuple[int, int]) -> Optional[pd.DataFrame]:
    df = read_sheet(excel_path, "DRE", sig)
    if df is None:
        return None
    d = df.copy()
    d["_dt"] = pd.to_datetime(d.get("DTA.PAG"), errors="coerce", dayfirst=True)
    d = d[d["_dt"].notna()].copy()
    d["_ano"] = d["_dt"].dt.year
    d["_mes"] = d["_dt"].dt.month
    d["_v"] = d.get("VAL.PAG").apply(to_num) if "VAL.PAG" in d.columns else 0.0

    for c in ["CONTA DE RESULTADO", "DESPESA", "FAVORECIDO", "HISTÓRICO", "DUPLICATA"]:
        if c in d.columns:
            d[c] = d[c].astype(str).str.strip()
    return d


@st.cache_data(show_spinner=False)
def prep_recebimentos(excel_path: str, sig: Tuple[int, int]) -> Optional[pd.DataFrame]:
    df = read_sheet(excel_path, "RECEBIMENTOS", sig)
    if df is None:
        return None
    r = df.copy()
    r["_ano"] = pd.to_numeric(r.get("ANO"), errors="coerce").astype("Int64")
    r["_mes"] = r.get("MÊS").apply(parse_mes)
    r["_v"] = r.get("VALOR").apply(to_num) if "VALOR" in r.columns else 0.0
    r = r[r["_ano"].notna() & r["_mes"].notna()].copy()
    r["_ano"] = r["_ano"].astype(int)
    r["_mes"] = r["_mes"].astype(int)
    return r


@st.cache_data(show_spinner=False)
def prep_compras_fornecedor(excel_path: str, sig: Tuple[int, int]) -> Optional[pd.DataFrame]:
    df = read_sheet(excel_path, "Compras fornecedor", sig)
    if df is None:
        return None
    c = df.copy()
    c["_dt"] = pd.to_datetime(c.get("DATA"), errors="coerce", dayfirst=True)
    c = c[c["_dt"].notna()].copy()
    c["_ano"] = c["_dt"].dt.year
    c["_mes"] = c["_dt"].dt.month
    if "VR. CONTÁBIL" in c.columns:
        c["_v"] = c["VR. CONTÁBIL"].apply(to_num)
    elif "VR.CONTÁBIL" in c.columns:
        c["_v"] = c["VR.CONTÁBIL"].apply(to_num)
    else:
        c["_v"] = 0.0
    if "FORNECEDOR" in c.columns:
        c["FORNECEDOR"] = c["FORNECEDOR"].astype(str).str.strip()
    else:
        c["FORNECEDOR"] = "—"
    return c


def month_series(df: pd.DataFrame, val_col: str, ano: int, meses: List[int]) -> Dict[int, float]:
    tmp = df[df["_ano"] == int(ano)].copy()
    if meses:
        tmp = tmp[tmp["_mes"].isin(meses)]
    grp = tmp.groupby("_mes")[val_col].sum()
    return {m: float(grp.get(m, 0.0)) for m in range(1, 13)}


def sum_by_account(df_dre: pd.DataFrame, conta_prefix: str, ano: int, meses: List[int]) -> Dict[int, float]:
    tmp = df_dre[df_dre["_ano"] == int(ano)].copy()
    if meses:
        tmp = tmp[tmp["_mes"].isin(meses)]
    mask = tmp["CONTA DE RESULTADO"].astype(str).str.startswith(conta_prefix)
    grp = tmp[mask].groupby("_mes")["_v"].sum()
    return {m: float(grp.get(m, 0.0)) for m in range(1, 13)}


def sum_by_account_shift_next_month(df_all: pd.DataFrame, conta_prefix: str, ano_ref: int) -> Dict[int, float]:
    """
    Regra: para exibir o mês m do ano_ref, usa os lançamentos do mês (m+1).
    Dezembro usa Janeiro do ano_ref+1.
    """
    cur = df_all[df_all["_ano"] == int(ano_ref)].copy()
    nxt = df_all[df_all["_ano"] == int(ano_ref) + 1].copy()

    def _sum_for(df: pd.DataFrame, mes: int) -> float:
        if df.empty:
            return 0.0
        d = df[(df["_mes"] == mes) & (df["CONTA DE RESULTADO"].astype(str).str.startswith(conta_prefix))]
        return float(d["_v"].sum())

    out = {}
    for m in range(1, 13):
        if m < 12:
            out[m] = _sum_for(cur, m + 1)
        else:
            out[m] = _sum_for(nxt, 1)
    return out


def _date_filter_to_months(date_ini: Optional[pd.Timestamp], date_fim: Optional[pd.Timestamp], ano_ref: int) -> List[int]:
    if date_ini is None or date_fim is None:
        return list(range(1, 13))
    di = pd.Timestamp(date_ini)
    df = pd.Timestamp(date_fim)
    if df < di:
        di, df = df, di
    if df.year < ano_ref or di.year > ano_ref:
        return list(range(1, 13))
    start_m = 1 if di.year < ano_ref else int(di.month)
    end_m = 12 if df.year > ano_ref else int(df.month)
    start_m = max(1, min(12, start_m))
    end_m = max(1, min(12, end_m))
    return list(range(start_m, end_m + 1))


def make_dre_table(
    receita_by_month: Dict[int, float],
    cmv_by_month: Dict[int, float],
    despesas_map_by_month: List[Tuple[str, Dict[int, float], str]],
    meses_exib: List[int],
) -> pd.DataFrame:
    margem_by_month = {m: float(receita_by_month.get(m, 0.0)) - float(cmv_by_month.get(m, 0.0)) for m in range(1, 13)}
    markup_by_month = {m: (float(receita_by_month.get(m, 0.0)) / float(cmv_by_month.get(m, 0.0)) if float(cmv_by_month.get(m, 0.0)) != 0 else 0.0) for m in range(1, 13)}

    # despesas total e componentes
    desp_total_by_month = {m: 0.0 for m in range(1, 13)}
    inv_by_month = {m: 0.0 for m in range(1, 13)}
    fin_by_month = {m: 0.0 for m in range(1, 13)}
    for nome, by_m, prefix in despesas_map_by_month:
        for m in range(1, 13):
            desp_total_by_month[m] += float(by_m.get(m, 0.0))
            if prefix == "00022":
                inv_by_month[m] = float(by_m.get(m, 0.0))
            if prefix == "00023":
                fin_by_month[m] = float(by_m.get(m, 0.0))

    resultado_oper_by_month = {m: float(margem_by_month.get(m, 0.0)) - float(desp_total_by_month.get(m, 0.0)) for m in range(1, 13)}
    resultado_antes_fin_ret_by_month = {m: float(resultado_oper_by_month[m]) + float(inv_by_month[m]) + float(fin_by_month[m]) for m in range(1, 13)}

    linhas: List[Tuple[str, Dict[int, float], str]] = []
    linhas.append(("RECEITA", receita_by_month, "currency"))
    linhas.append(("CMV", cmv_by_month, "currency"))
    linhas.append(("MARGEM BRUTA", margem_by_month, "currency"))
    linhas.append(("MARKUP", markup_by_month, "ratio"))
    for nome, by_m, _p in despesas_map_by_month:
        linhas.append((nome, by_m, "currency"))

    # ORDEM solicitada: "antes..." primeiro, depois "operacional"
    linhas.append(("RESULTADO antes das Desp financeiras e RETIRADAS", resultado_antes_fin_ret_by_month, "currency"))
    linhas.append(("RESULTADO OPERACIONAL", resultado_oper_by_month, "currency"))

    rows = []
    receita_total = float(sum(receita_by_month.get(m, 0.0) for m in meses_exib))
    cmv_total = float(sum(cmv_by_month.get(m, 0.0) for m in meses_exib))
    markup_acum = (receita_total / cmv_total) if cmv_total != 0 else 0.0

    for nome, by_m, typ in linhas:
        row = {"LINHA": nome, "_type": typ}
        acum_val = 0.0

        for m in meses_exib:
            mes_pt = MES_NUM_TO_PT[m]
            v = float(by_m.get(m, 0.0))
            rec = float(receita_by_month.get(m, 0.0))
            pct = 100.0 if nome == "RECEITA" else ((v / rec * 100.0) if (rec != 0 and typ != "ratio") else 0.0)

            row[mes_pt] = v
            row[f"{mes_pt}%"] = pct

            if typ != "ratio":
                acum_val += v

        if typ == "ratio":
            row["ACUM"] = markup_acum
            row["ACUM%"] = np.nan
        else:
            row["ACUM"] = acum_val
            row["ACUM%"] = 100.0 if nome == "RECEITA" else ((acum_val / receita_total * 100.0) if receita_total != 0 else 0.0)

        rows.append(row)

    return pd.DataFrame(rows)


def make_dfc_table(
    receb_by_month: Dict[int, float],
    saidas_map_by_month: List[Tuple[str, Dict[int, float], str]],
    meses_exib: List[int],
) -> pd.DataFrame:
    saidas_total = {m: 0.0 for m in range(1, 13)}
    inv_by_month = {m: 0.0 for m in range(1, 13)}
    fin_by_month = {m: 0.0 for m in range(1, 13)}
    for nome, by_m, prefix in saidas_map_by_month:
        for m in range(1, 13):
            saidas_total[m] += float(by_m.get(m, 0.0))
            if prefix == "00022":
                inv_by_month[m] = float(by_m.get(m, 0.0))
            if prefix == "00023":
                fin_by_month[m] = float(by_m.get(m, 0.0))

    saldo_by_month = {m: float(receb_by_month.get(m, 0.0)) - float(saidas_total.get(m, 0.0)) for m in range(1, 13)}
    saldo_antes_fin_ret_by_month = {m: float(saldo_by_month[m]) + float(fin_by_month[m]) + float(inv_by_month[m]) for m in range(1, 13)}

    linhas: List[Tuple[str, Dict[int, float], str]] = []
    linhas.append(("RECEBIMENTOS", receb_by_month, "currency"))
    for nome, by_m, _p in saidas_map_by_month:
        linhas.append((nome, by_m, "currency"))

    # ORDEM solicitada: "antes..." primeiro, depois "saldo"
    linhas.append(("SALDO OPERACIONAL antes das Desp financeiras e RETIRADAS", saldo_antes_fin_ret_by_month, "currency"))
    linhas.append(("SALDO OPERACIONAL", saldo_by_month, "currency"))

    rows = []
    receb_total = float(sum(receb_by_month.get(m, 0.0) for m in meses_exib))

    for nome, by_m, typ in linhas:
        row = {"LINHA": nome, "_type": typ}
        acum_val = 0.0

        for m in meses_exib:
            mes_pt = MES_NUM_TO_PT[m]
            v = float(by_m.get(m, 0.0))
            rec = float(receb_by_month.get(m, 0.0))
            pct = 100.0 if nome == "RECEBIMENTOS" else ((v / rec * 100.0) if rec != 0 else 0.0)

            row[mes_pt] = v
            row[f"{mes_pt}%"] = pct
            acum_val += v

        row["ACUM"] = acum_val
        row["ACUM%"] = 100.0 if nome == "RECEBIMENTOS" else ((acum_val / receb_total * 100.0) if receb_total != 0 else 0.0)
        rows.append(row)

    return pd.DataFrame(rows)


def style_table(df: pd.DataFrame, meses_exib: List[int], highlight_rows: List[str]) -> "pd.io.formats.style.Styler":
    """
    Compatível com pandas antigos.
    - Formata valores em R$ (exceto MARKUP, que é numérico).
    - Formata percentuais em % (MARKUP% vira "—").
    - Colore linhas destacadas: azul se positivo, vermelho se negativo.
    """
    cols_value = [MES_NUM_TO_PT[m] for m in meses_exib] + ["ACUM"]
    cols_pct = [f"{MES_NUM_TO_PT[m]}%" for m in meses_exib] + ["ACUM%"]
    cols_all = ["LINHA"] + sum([[MES_NUM_TO_PT[m], f"{MES_NUM_TO_PT[m]}%"] for m in meses_exib], []) + ["ACUM", "ACUM%"]

    base = df[cols_all + ["_type"]].copy()
    num_base = base.copy()

    show = base[cols_all].copy()

    for i in show.index:
        typ = str(base.loc[i, "_type"])
        for c in cols_value:
            v = num_base.loc[i, c]
            if typ == "ratio":
                try:
                    show.loc[i, c] = f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                except Exception:
                    show.loc[i, c] = "0,00"
            else:
                show.loc[i, c] = f"R$ {format_brl(v)}"

        for c in cols_pct:
            v = num_base.loc[i, c]
            if typ == "ratio":
                show.loc[i, c] = "—"
            else:
                show.loc[i, c] = "—" if pd.isna(v) else fmt_pct(v)

    sty = show.style

    def _row_style(row):
        styles = [""] * len(row.index)
        if str(row["LINHA"]) not in highlight_rows:
            return styles
        for j, col in enumerate(row.index):
            if col in cols_value:
                try:
                    v = float(num_base.loc[row.name, col])
                except Exception:
                    v = 0.0
                if v > 0:
                    styles[j] = "color: blue; font-weight: 700;"
                elif v < 0:
                    styles[j] = "color: red; font-weight: 700;"
        return styles

    return sty.apply(_row_style, axis=1)


# =========================
# UI
# =========================
st.set_page_config(page_title="Indicadores Paracatu (DRE/DFC)", layout="wide")

excel_path = _auto_find_excel()
if not excel_path:
    st.error("Não encontrei nenhum Excel (.xlsx/.xlsm/.xls) na pasta do app. Coloque o Excel junto do .py (ex.: 'projeto Paracatu.xlsx').")
    st.stop()

sig = excel_signature(excel_path)

df_rcm = prep_receita_cmv(excel_path, sig)
df_dre = prep_dre_lancamentos(excel_path, sig)
df_rec = prep_recebimentos(excel_path, sig)
df_comp = prep_compras_fornecedor(excel_path, sig)

if df_rcm is None or df_dre is None or df_rec is None:
    faltas = []
    if df_rcm is None: faltas.append("RECEITA E CMV")
    if df_dre is None: faltas.append("DRE")
    if df_rec is None: faltas.append("RECEBIMENTOS")
    st.error(f"Falha ao ler abas obrigatórias: {', '.join(faltas)}")
    st.stop()

# Sidebar
st.sidebar.title("Filtros")

anos = sorted(set(df_rcm["_ano"].dropna().astype(int).unique().tolist()) |
              set(df_dre["_ano"].dropna().astype(int).unique().tolist()) |
              set(df_rec["_ano"].dropna().astype(int).unique().tolist()))
if not anos:
    st.sidebar.error("Não encontrei anos válidos no arquivo.")
    st.stop()

ano_ref = st.sidebar.selectbox("Ano", options=anos, index=len(anos) - 1)

min_dt = pd.Timestamp(year=int(ano_ref), month=1, day=1)
max_dt = pd.Timestamp(year=int(ano_ref), month=12, day=31)
date_ini, date_fim = st.sidebar.date_input(
    "Período (opcional)",
    value=(min_dt.date(), max_dt.date()),
    min_value=min_dt.date(),
    max_value=max_dt.date(),
)
date_ini = pd.Timestamp(date_ini) if date_ini else None
date_fim = pd.Timestamp(date_fim) if date_fim else None

# Filtro adicional por mês (mantém o calendário)
meses_sel = st.sidebar.multiselect(
    "Meses (opcional)",
    options=[MES_NUM_TO_PT[m] for m in range(1, 13)],
    default=[],
)

# Meses exibidos: interseção do período do calendário com os meses selecionados (se houver)
meses_exib = _date_filter_to_months(date_ini, date_fim, int(ano_ref))
if meses_sel:
    meses_exib = [m for m in meses_exib if MES_NUM_TO_PT[m] in set(meses_sel)]
    if not meses_exib:
        # se o usuário selecionar meses fora do período do calendário, respeita os meses selecionados
        meses_exib = [MES_PT_TO_NUM[x] for x in meses_sel]


st.sidebar.caption(f"Meses exibidos: **{', '.join(MES_NUM_TO_PT[m] for m in meses_exib)}**")

pagina = st.sidebar.radio("Página", ["DRE", "DFC", "INDICADOR DE COMPRAS", "INDICADORES COMERCIAIS"], index=0)


# =========================
# DRE
# =========================
if pagina == "DRE":
    st.title("DRE — Indicadores Paracatu")

    receita_by_month = month_series(df_rcm, "_receita", int(ano_ref), meses_exib)
    cmv_by_month = month_series(df_rcm, "_cmv", int(ano_ref), meses_exib)

    contas_dre = [
        ("DESPESAS COM PESSOAL", "00024"),
        ("DESPESAS ADMINISTRATIVAS", "00019"),
        ("DESPESAS OPERACIONAIS", "00018"),
        ("DESPESAS COMERCIAIS", "00020"),
        ("DEDUÇÕES (IMPOSTOS SOBRE VENDAS)", "00021"),
        ("INVESTIMENTOS / RETIRADAS", "00022"),
        ("DESPESAS FINANCEIRAS", "00023"),
    ]

    despesas_map_by_month: List[Tuple[str, Dict[int, float], str]] = []
    for nome, prefix in contas_dre:
        if prefix in ("00024", "00021"):
            by_m = sum_by_account_shift_next_month(df_dre, prefix, int(ano_ref))
        else:
            by_m = sum_by_account(df_dre, prefix, int(ano_ref), meses_exib)
        despesas_map_by_month.append((nome, by_m, prefix))

    dre_tbl = make_dre_table(receita_by_month, cmv_by_month, despesas_map_by_month, meses_exib)

    highlight_rows = ["RESULTADO OPERACIONAL", "RESULTADO antes das Desp financeiras e RETIRADAS"]

    st.subheader("Tabela DRE")
    st.dataframe(style_table(dre_tbl, meses_exib, highlight_rows), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Drill — Por linha (todas as linhas)")

    linhas_opts = dre_tbl["LINHA"].tolist()
    linha_sel = st.selectbox("Selecione a linha", options=linhas_opts, index=0)

    row = dre_tbl[dre_tbl["LINHA"] == linha_sel].iloc[0]
    typ = row["_type"]

    if typ == "ratio":
        receita_total = float(sum(receita_by_month.get(m, 0.0) for m in meses_exib))
        cmv_total = float(sum(cmv_by_month.get(m, 0.0) for m in meses_exib))
        total = (receita_total / cmv_total) if cmv_total != 0 else 0.0
        media = total
        c1, c2, c3 = st.columns(3)
        c1.metric("Markup (período)", f"{total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        c2.metric("Markup (média)", f"{media:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        c3.metric("% sobre Receita", "—")
    else:
        vals = np.array([float(row[MES_NUM_TO_PT[m]]) for m in meses_exib], dtype=float)
        total = float(np.nansum(vals))
        media = float(total / max(len(meses_exib), 1))
        receita_total = float(sum(receita_by_month.get(m, 0.0) for m in meses_exib))
        pct_receita = 100.0 if linha_sel == "RECEITA" else ((total / receita_total * 100.0) if receita_total != 0 else 0.0)

        c1, c2, c3 = st.columns(3)
        c1.metric("Total (período)", f"R$ {format_brl(total)}")
        c2.metric("Média mensal (período)", f"R$ {format_brl(media)}")
        c3.metric("% sobre Receita (período)", fmt_pct(pct_receita))

    st.divider()
    st.subheader("Drill — Detalhamento (DESPESA / HISTÓRICO / FAVORECIDO)")

    conta_map = {n: p for n, p in contas_dre}
    if linha_sel in conta_map:
        prefix = conta_map[linha_sel]
        mes_opt = ["TODOS"] + [MES_NUM_TO_PT[m] for m in meses_exib]
        mes_sel = st.selectbox("Mês (para detalhar)", options=mes_opt, index=0)

        meses_drill = meses_exib if mes_sel == "TODOS" else [MES_PT_TO_NUM[mes_sel]]

        # aplica shift no drill
        if prefix in ("00024", "00021"):
            src_months_cur = [m + 1 for m in meses_drill if m < 12]
            need_jan_next = any(m == 12 for m in meses_drill)
            cur = df_dre[df_dre["_ano"] == int(ano_ref)].copy()
            nxt = df_dre[df_dre["_ano"] == int(ano_ref) + 1].copy()

            b1 = cur[(cur["_mes"].isin(src_months_cur)) & (cur["CONTA DE RESULTADO"].astype(str).str.startswith(prefix))]
            if need_jan_next:
                b2 = nxt[(nxt["_mes"] == 1) & (nxt["CONTA DE RESULTADO"].astype(str).str.startswith(prefix))]
                base_raw = pd.concat([b1, b2], ignore_index=True)
            else:
                base_raw = b1.copy()
        else:
            base_raw = df_dre[(df_dre["_ano"] == int(ano_ref)) & (df_dre["_mes"].isin(meses_drill))].copy()
            base_raw = base_raw[base_raw["CONTA DE RESULTADO"].astype(str).str.startswith(prefix)]

        if base_raw.empty:
            st.info("Sem lançamentos para o filtro selecionado.")
        else:
            for c in ["DESPESA", "FAVORECIDO", "HISTÓRICO"]:
                if c not in base_raw.columns:
                    base_raw[c] = "—"

            base_raw["DESPESA_SINT"] = base_raw["DESPESA"].apply(sintetizar_despesa)

            receita_base = float(sum(receita_by_month.get(m, 0.0) for m in (meses_drill if mes_sel != "TODOS" else meses_exib)))

            agg = (base_raw.groupby("DESPESA_SINT", dropna=False)["_v"].sum()
                   .reset_index().rename(columns={"_v": "Valor"}))
            agg["% Receita"] = (agg["Valor"] / receita_base * 100.0) if receita_base != 0 else 0.0
            agg = agg.sort_values("Valor", ascending=False)

            st.markdown("#### Despesas (por DESPESA sintetizada)")
            show = agg.copy()
            show["Valor"] = show["Valor"].apply(lambda x: f"R$ {format_brl(x)}")
            show["% Receita"] = show["% Receita"].apply(fmt_pct)
            st.dataframe(show, use_container_width=True, hide_index=True)

            desp_sel = st.selectbox("Selecione uma despesa", options=agg["DESPESA_SINT"].tolist(), index=0)
            raw_sel = base_raw[base_raw["DESPESA_SINT"] == desp_sel].copy()
            total_sel = float(raw_sel["_v"].sum())
            pct_sel = (total_sel / receita_base * 100.0) if receita_base != 0 else 0.0
            st.metric("Total da despesa selecionada", f"R$ {format_brl(total_sel)}", fmt_pct(pct_sel))

            tab1, tab2, tab3 = st.tabs(["Histórico detalhado", "Histórico sintetizado", "Sintetizado por favorecido"])

            with tab1:
                cols = [c for c in ["DTA.PAG", "CONTA DE RESULTADO", "DESPESA", "FAVORECIDO", "DUPLICATA", "HISTÓRICO", "VAL.PAG"] if c in raw_sel.columns]
                det = raw_sel.sort_values("_dt", ascending=False)[cols].copy() if cols else raw_sel.copy()
                if "VAL.PAG" in det.columns:
                    det["VAL.PAG"] = det["VAL.PAG"].apply(to_num).apply(lambda x: f"R$ {format_brl(x)}")
                st.dataframe(det, use_container_width=True, hide_index=True)

            with tab2:
                tmp = raw_sel.copy()
                tmp["HISTÓRICO"] = tmp["HISTÓRICO"].astype(str).str.strip().replace({"": "—"})
                hist = (tmp.groupby("HISTÓRICO", dropna=False)["_v"].sum().reset_index().rename(columns={"_v": "Valor"}))
                hist["% Receita"] = (hist["Valor"] / receita_base * 100.0) if receita_base != 0 else 0.0
                hist = hist.sort_values("Valor", ascending=False)
                show2 = hist.copy()
                show2["Valor"] = show2["Valor"].apply(lambda x: f"R$ {format_brl(x)}")
                show2["% Receita"] = show2["% Receita"].apply(fmt_pct)
                st.dataframe(show2, use_container_width=True, hide_index=True)

            with tab3:
                tmp = raw_sel.copy()
                tmp["FAVORECIDO"] = tmp["FAVORECIDO"].astype(str).str.strip().replace({"": "—"})
                fav = (tmp.groupby("FAVORECIDO", dropna=False)["_v"].sum().reset_index().rename(columns={"_v": "Valor"}))
                fav["% Receita"] = (fav["Valor"] / receita_base * 100.0) if receita_base != 0 else 0.0
                fav = fav.sort_values("Valor", ascending=False)
                show3 = fav.copy()
                show3["Valor"] = show3["Valor"].apply(lambda x: f"R$ {format_brl(x)}")
                show3["% Receita"] = show3["% Receita"].apply(fmt_pct)
                st.dataframe(show3, use_container_width=True, hide_index=True)
    else:
        st.info("Para detalhar (DESPESA/HISTÓRICO/FAVORECIDO), selecione uma linha de despesa (00018..00024).")


# =========================
# DFC
# =========================
elif pagina == "DFC":
    st.title("DFC — Indicadores Paracatu")

    receb_by_month = month_series(df_rec, "_v", int(ano_ref), meses_exib)

    contas_dfc = [
        ("FORNECEDORES (COMPRAS P/ REVENDA)", "00025"),
        ("DEDUÇÕES (IMPOSTOS SOBRE VENDAS)", "00021"),
        ("DESPESAS COM PESSOAL", "00024"),
        ("DESPESAS ADMINISTRATIVAS", "00019"),
        ("DESPESAS OPERACIONAIS", "00018"),
        ("DESPESAS COMERCIAIS", "00020"),
        ("INVESTIMENTOS / RETIRADAS", "00022"),
        ("DESPESAS FINANCEIRAS", "00023"),
    ]

    saidas_map_by_month: List[Tuple[str, Dict[int, float], str]] = []
    for nome, prefix in contas_dfc:
        by_m = sum_by_account(df_dre, prefix, int(ano_ref), meses_exib)
        saidas_map_by_month.append((nome, by_m, prefix))

    dfc_tbl = make_dfc_table(receb_by_month, saidas_map_by_month, meses_exib)

    highlight_rows = ["SALDO OPERACIONAL", "SALDO OPERACIONAL antes das Desp financeiras e RETIRADAS"]

    st.subheader("Tabela DFC")
    st.dataframe(style_table(dfc_tbl, meses_exib, highlight_rows), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Drill — Por linha (todas as linhas)")

    linhas_opts = dfc_tbl["LINHA"].tolist()
    linha_sel = st.selectbox("Selecione a linha", options=linhas_opts, index=0, key="dfc_line")

    row = dfc_tbl[dfc_tbl["LINHA"] == linha_sel].iloc[0]
    vals = np.array([float(row[MES_NUM_TO_PT[m]]) for m in meses_exib], dtype=float)
    total = float(np.nansum(vals))
    media = float(total / max(len(meses_exib), 1))
    receb_total = float(sum(receb_by_month.get(m, 0.0) for m in meses_exib))
    pct_receb = 100.0 if linha_sel == "RECEBIMENTOS" else ((total / receb_total * 100.0) if receb_total != 0 else 0.0)

    c1, c2, c3 = st.columns(3)
    c1.metric("Total (período)", f"R$ {format_brl(total)}")
    c2.metric("Média mensal (período)", f"R$ {format_brl(media)}")
    c3.metric("% sobre Recebimentos (período)", fmt_pct(pct_receb))

    st.divider()
    st.subheader("Drill — Detalhamento (DESPESA / HISTÓRICO / FAVORECIDO)")

    conta_map = {n: p for n, p in contas_dfc}
    if linha_sel in conta_map:
        prefix = conta_map[linha_sel]
        mes_opt = ["TODOS"] + [MES_NUM_TO_PT[m] for m in meses_exib]
        mes_sel = st.selectbox("Mês (para detalhar)", options=mes_opt, index=0, key="dfc_mes_sel")
        meses_drill = meses_exib if mes_sel == "TODOS" else [MES_PT_TO_NUM[mes_sel]]

        base_raw = df_dre[(df_dre["_ano"] == int(ano_ref)) & (df_dre["_mes"].isin(meses_drill))].copy()
        base_raw = base_raw[base_raw["CONTA DE RESULTADO"].astype(str).str.startswith(prefix)]

        if base_raw.empty:
            st.info("Sem lançamentos para o filtro selecionado.")
        else:
            for c in ["DESPESA", "FAVORECIDO", "HISTÓRICO"]:
                if c not in base_raw.columns:
                    base_raw[c] = "—"
            base_raw["DESPESA_SINT"] = base_raw["DESPESA"].apply(sintetizar_despesa)

            receb_base = float(sum(receb_by_month.get(m, 0.0) for m in (meses_drill if mes_sel != "TODOS" else meses_exib)))

            agg = (base_raw.groupby("DESPESA_SINT", dropna=False)["_v"].sum()
                   .reset_index().rename(columns={"_v": "Valor"}))
            agg["% Recebimentos"] = (agg["Valor"] / receb_base * 100.0) if receb_base != 0 else 0.0
            agg = agg.sort_values("Valor", ascending=False)

            st.markdown("#### Saídas (por DESPESA sintetizada)")
            show = agg.copy()
            show["Valor"] = show["Valor"].apply(lambda x: f"R$ {format_brl(x)}")
            show["% Recebimentos"] = show["% Recebimentos"].apply(fmt_pct)
            st.dataframe(show, use_container_width=True, hide_index=True)

            desp_sel = st.selectbox("Selecione uma despesa", options=agg["DESPESA_SINT"].tolist(), index=0, key="dfc_desp")
            raw_sel = base_raw[base_raw["DESPESA_SINT"] == desp_sel].copy()
            total_sel = float(raw_sel["_v"].sum())
            pct_sel = (total_sel / receb_base * 100.0) if receb_base != 0 else 0.0
            st.metric("Total da despesa selecionada", f"R$ {format_brl(total_sel)}", fmt_pct(pct_sel))

            tab1, tab2, tab3 = st.tabs(["Histórico detalhado", "Histórico sintetizado", "Sintetizado por favorecido"])

            with tab1:
                cols = [c for c in ["DTA.PAG", "CONTA DE RESULTADO", "DESPESA", "FAVORECIDO", "DUPLICATA", "HISTÓRICO", "VAL.PAG"] if c in raw_sel.columns]
                det = raw_sel.sort_values("_dt", ascending=False)[cols].copy() if cols else raw_sel.copy()
                if "VAL.PAG" in det.columns:
                    det["VAL.PAG"] = det["VAL.PAG"].apply(to_num).apply(lambda x: f"R$ {format_brl(x)}")
                st.dataframe(det, use_container_width=True, hide_index=True)

            with tab2:
                tmp = raw_sel.copy()
                tmp["HISTÓRICO"] = tmp["HISTÓRICO"].astype(str).str.strip().replace({"": "—"})
                hist = (tmp.groupby("HISTÓRICO", dropna=False)["_v"].sum().reset_index().rename(columns={"_v": "Valor"}))
                hist["% Recebimentos"] = (hist["Valor"] / receb_base * 100.0) if receb_base != 0 else 0.0
                hist = hist.sort_values("Valor", ascending=False)
                show2 = hist.copy()
                show2["Valor"] = show2["Valor"].apply(lambda x: f"R$ {format_brl(x)}")
                show2["% Recebimentos"] = show2["% Recebimentos"].apply(fmt_pct)
                st.dataframe(show2, use_container_width=True, hide_index=True)

            with tab3:
                tmp = raw_sel.copy()
                tmp["FAVORECIDO"] = tmp["FAVORECIDO"].astype(str).str.strip().replace({"": "—"})
                fav = (tmp.groupby("FAVORECIDO", dropna=False)["_v"].sum().reset_index().rename(columns={"_v": "Valor"}))
                fav["% Recebimentos"] = (fav["Valor"] / receb_base * 100.0) if receb_base != 0 else 0.0
                fav = fav.sort_values("Valor", ascending=False)
                show3 = fav.copy()
                show3["Valor"] = show3["Valor"].apply(lambda x: f"R$ {format_brl(x)}")
                show3["% Recebimentos"] = show3["% Recebimentos"].apply(fmt_pct)
                st.dataframe(show3, use_container_width=True, hide_index=True)
    else:
        st.info("Para detalhar (DESPESA/HISTÓRICO/FAVORECIDO), selecione uma linha de saída (00018..00025).")


# =========================
# INDICADOR DE COMPRAS
# =========================
elif pagina == "INDICADOR DE COMPRAS":
    st.title("INDICADOR DE COMPRAS")

    if df_comp is None:
        st.error("A aba 'Compras fornecedor' não foi encontrada ou não pôde ser lida. Verifique o nome da aba no Excel.")
        st.stop()

    cmv_by_month = month_series(df_rcm, "_cmv", int(ano_ref), meses_exib)
    compras_by_month = month_series(df_comp, "_v", int(ano_ref), meses_exib)
    diff_by_month = {m: float(cmv_by_month.get(m, 0.0)) - float(compras_by_month.get(m, 0.0)) for m in range(1, 13)}

    cols = ["LINHA"] + [MESES_FULL[m] for m in meses_exib] + ["ACUM"]

    def _row(nome, by_month):
        r = {"LINHA": nome}
        acum = 0.0
        for m in meses_exib:
            v = float(by_month.get(m, 0.0))
            r[MESES_FULL[m]] = v
            acum += v
        r["ACUM"] = acum
        return r

    tbl = pd.DataFrame([
        _row("CMV NECESSIDADE", cmv_by_month),
        _row("COMPRAS REALIZADO", compras_by_month),
        _row("DIFERENÇA (CMV - COMPRAS)", diff_by_month),
    ])[cols]

    num_tbl = tbl.copy()
    show = tbl.copy()
    for c in cols[1:]:
        show[c] = show[c].apply(lambda x: f"R$ {format_brl(x)}")

    sty = show.style

    def _row_style(row):
        styles = [""] * len(row.index)
        if row["LINHA"] != "DIFERENÇA (CMV - COMPRAS)":
            return styles
        for j, col in enumerate(row.index):
            if col == "LINHA":
                continue
            try:
                v = float(num_tbl.loc[row.name, col])
            except Exception:
                v = 0.0
            if v > 0:
                styles[j] = "color: blue; font-weight: 700;"
            elif v < 0:
                styles[j] = "color: red; font-weight: 700;"
        return styles

    st.subheader("Tabela — CMV x Compras")
    st.dataframe(sty.apply(_row_style, axis=1), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Drill — Compras por fornecedor")

    mes_opt = ["TODOS"] + [MESES_FULL[m] for m in meses_exib]
    mes_sel = st.selectbox("Mês (opcional)", options=mes_opt, index=0, key="compras_mes_sel")
    meses_drill = meses_exib if mes_sel == "TODOS" else [MESES_FULL_INV[mes_sel]]

    base = df_comp[(df_comp["_ano"] == int(ano_ref)) & (df_comp["_mes"].isin(meses_drill))].copy()
    total = float(base["_v"].sum())

    if base.empty or total == 0:
        st.info("Sem compras no período selecionado.")
    else:
        agg = (base.groupby("FORNECEDOR", dropna=False)["_v"].sum()
               .reset_index().rename(columns={"_v": "Valor"}))
        agg["% Participação"] = (agg["Valor"] / total * 100.0) if total != 0 else 0.0
        agg = agg.sort_values("Valor", ascending=False)

        show2 = agg.copy()
        show2["Valor"] = show2["Valor"].apply(lambda x: f"R$ {format_brl(x)}")
        show2["% Participação"] = show2["% Participação"].apply(fmt_pct)
        st.dataframe(show2, use_container_width=True, hide_index=True)


# =========================
# INDICADORES COMERCIAIS
# =========================
else:
    st.title("INDICADORES COMERCIAIS")

    start = pd.Timestamp(date_ini) if date_ini is not None else pd.Timestamp(year=int(ano_ref), month=1, day=1)
    end = pd.Timestamp(date_fim) if date_fim is not None else pd.Timestamp(year=int(ano_ref), month=12, day=31)
    if end < start:
        start, end = end, start

    base_cur = df_rcm[(df_rcm["_dt"] >= start) & (df_rcm["_dt"] <= end)].copy()

    start_y1 = start - pd.DateOffset(years=1)
    end_y1 = end - pd.DateOffset(years=1)
    base_y1 = df_rcm[(df_rcm["_dt"] >= start_y1) & (df_rcm["_dt"] <= end_y1)].copy()

    fat_cur = float(base_cur["_receita"].sum())
    fat_y1 = float(base_y1["_receita"].sum())
    crescimento_pct = ((fat_cur / fat_y1 - 1) * 100.0) if fat_y1 != 0 else (100.0 if fat_cur != 0 else 0.0)

    c1, c2, c3 = st.columns(3)
    c1.metric("Faturamento (período)", f"R$ {format_brl(fat_cur)}")
    c2.metric("Faturamento Ano-1 (mesmo período)", f"R$ {format_brl(fat_y1)}")
    c3.metric("Crescimento", fmt_pct(crescimento_pct))

    st.divider()

    with st.expander("Comparativo Ano-1 por mês (abrir/fechar)", expanded=False):
        meses_periodo = list(meses_exib)
        cur_m = base_cur.groupby(base_cur["_dt"].dt.month)["_receita"].sum()
        y1_m = base_y1.groupby(base_y1["_dt"].dt.month)["_receita"].sum()

        rows = []
        for m in meses_periodo:
            v_cur = float(cur_m.get(m, 0.0))
            v_y1 = float(y1_m.get(m, 0.0))
            dif_r = v_cur - v_y1
            dif_p = ((v_cur / v_y1 - 1) * 100.0) if v_y1 != 0 else (100.0 if v_cur != 0 else 0.0)
            rows.append({
                "MÊS": MESES_FULL.get(int(m), str(m)),
                f"{int(ano_ref)-1}": v_y1,
                f"{int(ano_ref)}": v_cur,
                "DIF (R$)": dif_r,
                "DIF (%)": dif_p,
            })

        t = pd.DataFrame(rows)
        if t.empty:
            st.info("Sem dados no período selecionado.")
        else:
            show = t.copy()
            for col in [str(int(ano_ref)-1), str(int(ano_ref)), "DIF (R$)"]:
                show[col] = show[col].apply(lambda x: f"R$ {format_brl(x)}")
            show["DIF (%)"] = show["DIF (%)"].apply(fmt_pct)
            st.dataframe(show, use_container_width=True, hide_index=True)

    st.subheader("Faturamento por mês")
    if base_cur.empty:
        st.info("Sem dados para o período selecionado.")
    else:
        bar = (base_cur.assign(MES=base_cur["_dt"].dt.month)
               .groupby("MES")["_receita"].sum()
               .reindex(range(1, 13), fill_value=0.0)
               .reset_index())
        bar["MÊS"] = bar["MES"].map(MESES_FULL)
        fig_bar = px.bar(bar, x="MÊS", y="_receita")
        fig_bar.update_layout(yaxis_title="Faturamento (R$)", xaxis_title=None)
        st.plotly_chart(fig_bar, use_container_width=True)

    st.divider()
    st.subheader("Participação por Segmento")
    if base_cur.empty or "SEGMENTO" not in base_cur.columns:
        st.info("Coluna SEGMENTO não encontrada ou sem dados no período.")
    else:
        seg = (base_cur.groupby("SEGMENTO", dropna=False)["_receita"].sum()
               .reset_index().rename(columns={"_receita": "Faturamento"}))
        seg["SEGMENTO"] = seg["SEGMENTO"].fillna("—").astype(str).str.strip().replace({"": "—"})
        total_seg = float(seg["Faturamento"].sum())
        seg["%"] = (seg["Faturamento"] / total_seg * 100.0) if total_seg != 0 else 0.0
        seg = seg.sort_values("Faturamento", ascending=False)

        fig_pie = px.pie(seg, names="SEGMENTO", values="Faturamento")
        st.plotly_chart(fig_pie, use_container_width=True)

        show = seg.copy()
        show["Faturamento"] = show["Faturamento"].apply(lambda x: f"R$ {format_brl(x)}")
        show["%"] = show["%"].apply(fmt_pct)
        st.dataframe(show, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Marcas — Top 10")
    if base_cur.empty or "MARCA" not in base_cur.columns:
        st.info("Coluna MARCA não encontrada ou sem dados.")
    else:
        mdf = (base_cur.groupby("MARCA", dropna=False)["_receita"].sum()
               .reset_index().rename(columns={"_receita": "Faturamento"}))
        mdf["MARCA"] = mdf["MARCA"].fillna("—").astype(str).str.strip().replace({"": "—"})
        tot = float(mdf["Faturamento"].sum())
        mdf["%"] = (mdf["Faturamento"] / tot * 100.0) if tot != 0 else 0.0
        mdf = mdf.sort_values("Faturamento", ascending=False)

        top10 = mdf.head(10).copy()
        show10 = top10.copy()
        show10["Faturamento"] = show10["Faturamento"].apply(lambda x: f"R$ {format_brl(x)}")
        show10["%"] = show10["%"].apply(fmt_pct)
        st.dataframe(show10, use_container_width=True, hide_index=True)

        with st.expander("Ver todas as marcas", expanded=False):
            show_all = mdf.copy()
            show_all["Faturamento"] = show_all["Faturamento"].apply(lambda x: f"R$ {format_brl(x)}")
            show_all["%"] = show_all["%"].apply(fmt_pct)
            st.dataframe(show_all, use_container_width=True, hide_index=True)

        st.markdown("#### Drill — Linhas dentro da Marca")
        marca_sel = st.selectbox("Selecione a marca", options=mdf["MARCA"].tolist(), index=0, key="marca_sel")
        base_m = base_cur[base_cur["MARCA"].fillna("—").astype(str).str.strip().replace({"": "—"}) == marca_sel].copy()
        if base_m.empty or "LINHA" not in base_m.columns:
            st.info("Sem dados para a marca selecionada ou coluna LINHA ausente.")
        else:
            ldf = (base_m.groupby("LINHA", dropna=False)["_receita"].sum()
                   .reset_index().rename(columns={"_receita": "Faturamento"}))
            ldf["LINHA"] = ldf["LINHA"].fillna("—").astype(str).str.strip().replace({"": "—"})
            totm = float(ldf["Faturamento"].sum())
            ldf["% (sobre a marca)"] = (ldf["Faturamento"] / totm * 100.0) if totm != 0 else 0.0
            ldf = ldf.sort_values("Faturamento", ascending=False)

            showl = ldf.copy()
            showl["Faturamento"] = showl["Faturamento"].apply(lambda x: f"R$ {format_brl(x)}")
            showl["% (sobre a marca)"] = showl["% (sobre a marca)"].apply(fmt_pct)
            st.dataframe(showl, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Linhas — Top 10")
    if base_cur.empty or "LINHA" not in base_cur.columns:
        st.info("Coluna LINHA não encontrada ou sem dados.")
    else:
        lall = (base_cur.groupby("LINHA", dropna=False)["_receita"].sum()
                .reset_index().rename(columns={"_receita": "Faturamento"}))
        lall["LINHA"] = lall["LINHA"].fillna("—").astype(str).str.strip().replace({"": "—"})
        totl = float(lall["Faturamento"].sum())
        lall["%"] = (lall["Faturamento"] / totl * 100.0) if totl != 0 else 0.0
        lall = lall.sort_values("Faturamento", ascending=False)

        top10l = lall.head(10).copy()
        show10l = top10l.copy()
        show10l["Faturamento"] = show10l["Faturamento"].apply(lambda x: f"R$ {format_brl(x)}")
        show10l["%"] = show10l["%"].apply(fmt_pct)
        st.dataframe(show10l, use_container_width=True, hide_index=True)

        with st.expander("Ver todas as linhas", expanded=False):
            show_all_l = lall.copy()
            show_all_l["Faturamento"] = show_all_l["Faturamento"].apply(lambda x: f"R$ {format_brl(x)}")
            show_all_l["%"] = show_all_l["%"].apply(fmt_pct)
            st.dataframe(show_all_l, use_container_width=True, hide_index=True)

        st.markdown("#### Drill — Marcas dentro da Linha")
        linha_sel = st.selectbox("Selecione a linha", options=lall["LINHA"].tolist(), index=0, key="linha_sel")
        base_l = base_cur[base_cur["LINHA"].fillna("—").astype(str).str.strip().replace({"": "—"}) == linha_sel].copy()
        if base_l.empty or "MARCA" not in base_l.columns:
            st.info("Sem dados para a linha selecionada ou coluna MARCA ausente.")
        else:
            bdf = (base_l.groupby("MARCA", dropna=False)["_receita"].sum()
                   .reset_index().rename(columns={"_receita": "Faturamento"}))
            bdf["MARCA"] = bdf["MARCA"].fillna("—").astype(str).str.strip().replace({"": "—"})
            totline = float(bdf["Faturamento"].sum())
            bdf["% (sobre a linha)"] = (bdf["Faturamento"] / totline * 100.0) if totline != 0 else 0.0
            bdf = bdf.sort_values("Faturamento", ascending=False)

            showb = bdf.copy()
            showb["Faturamento"] = showb["Faturamento"].apply(lambda x: f"R$ {format_brl(x)}")
            showb["% (sobre a linha)"] = showb["% (sobre a linha)"].apply(fmt_pct)
            st.dataframe(showb, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Clientes")
    if base_cur.empty or "CLIENTE" not in base_cur.columns:
        st.info("Coluna CLIENTE não encontrada ou sem dados.")
    else:
        cdf = (base_cur.groupby("CLIENTE", dropna=False)["_receita"].sum()
               .reset_index().rename(columns={"_receita": "Faturamento"}))
        cdf["CLIENTE"] = cdf["CLIENTE"].fillna("—").astype(str).str.strip().replace({"": "—"})
        totc = float(cdf["Faturamento"].sum())
        cdf["%"] = (cdf["Faturamento"] / totc * 100.0) if totc != 0 else 0.0
        cdf = cdf.sort_values("Faturamento", ascending=False)

        topc = cdf.head(20).copy()
        showc = topc.copy()
        showc["Faturamento"] = showc["Faturamento"].apply(lambda x: f"R$ {format_brl(x)}")
        showc["%"] = showc["%"].apply(fmt_pct)
        st.dataframe(showc, use_container_width=True, hide_index=True)

        with st.expander("Ver todos os clientes", expanded=False):
            show_all_c = cdf.copy()
            show_all_c["Faturamento"] = show_all_c["Faturamento"].apply(lambda x: f"R$ {format_brl(x)}")
            show_all_c["%"] = show_all_c["%"].apply(fmt_pct)
            st.dataframe(show_all_c, use_container_width=True, hide_index=True)

        st.markdown("#### Drill — Linhas e Marcas do Cliente")
        cliente_sel = st.selectbox("Selecione o cliente", options=cdf["CLIENTE"].tolist(), index=0, key="cliente_sel")
        base_c = base_cur[base_cur["CLIENTE"].fillna("—").astype(str).str.strip().replace({"": "—"}) == cliente_sel].copy()
        if base_c.empty:
            st.info("Sem dados para o cliente selecionado.")
        else:
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("**Linhas compradas**")
                if "LINHA" in base_c.columns:
                    lc = (base_c.groupby("LINHA", dropna=False)["_receita"].sum()
                          .reset_index().rename(columns={"_receita": "Faturamento"}))
                    lc["LINHA"] = lc["LINHA"].fillna("—").astype(str).str.strip().replace({"": "—"})
                    totlc = float(lc["Faturamento"].sum())
                    lc["%"] = (lc["Faturamento"] / totlc * 100.0) if totlc != 0 else 0.0
                    lc = lc.sort_values("Faturamento", ascending=False)
                    show_lc = lc.copy()
                    show_lc["Faturamento"] = show_lc["Faturamento"].apply(lambda x: f"R$ {format_brl(x)}")
                    show_lc["%"] = show_lc["%"].apply(fmt_pct)
                    st.dataframe(show_lc, use_container_width=True, hide_index=True)
                else:
                    st.info("Coluna LINHA não encontrada.")

            with col2:
                st.markdown("**Marcas compradas**")
                if "MARCA" in base_c.columns:
                    mc = (base_c.groupby("MARCA", dropna=False)["_receita"].sum()
                          .reset_index().rename(columns={"_receita": "Faturamento"}))
                    mc["MARCA"] = mc["MARCA"].fillna("—").astype(str).str.strip().replace({"": "—"})
                    totmc = float(mc["Faturamento"].sum())
                    mc["%"] = (mc["Faturamento"] / totmc * 100.0) if totmc != 0 else 0.0
                    mc = mc.sort_values("Faturamento", ascending=False)
                    show_mc = mc.copy()
                    show_mc["Faturamento"] = show_mc["Faturamento"].apply(lambda x: f"R$ {format_brl(x)}")
                    show_mc["%"] = show_mc["%"].apply(fmt_pct)
                    st.dataframe(show_mc, use_container_width=True, hide_index=True)
                else:
                    st.info("Coluna MARCA não encontrada.")


    st.divider()
    st.subheader("Evolução de clientes (por mês)")

    if base_cur.empty or "CLIENTE" not in base_cur.columns:
        st.info("Sem dados no período selecionado ou coluna CLIENTE não encontrada.")
    else:
        meses_cols = list(meses_exib) if meses_exib else list(range(1, 13))
        meses_names = [MESES_FULL[m] for m in meses_cols]

        tmp = base_cur.copy()
        tmp["MES"] = tmp["_dt"].dt.month

        piv = (tmp.pivot_table(index="CLIENTE", columns="MES", values="_receita", aggfunc="sum", fill_value=0.0)
               .reindex(columns=meses_cols, fill_value=0.0))

        piv.columns = [MESES_FULL[int(c)] for c in piv.columns]
        piv = piv.reset_index()

        piv["_TOTAL"] = piv[meses_names].sum(axis=1)
        piv = piv.sort_values("_TOTAL", ascending=False).drop(columns=["_TOTAL"])

        topn = 30
        top = piv.head(topn).copy()

        def _format_tbl(df_show: pd.DataFrame) -> pd.DataFrame:
            out = df_show.copy()
            for c in meses_names:
                if c in out.columns:
                    out[c] = out[c].apply(lambda x: f"R$ {format_brl(x)}")
            return out

        st.caption(f"Mostrando Top {topn} clientes por faturamento no período. (Zerado quando não há vendas no mês.)")
        st.dataframe(_format_tbl(top), use_container_width=True, hide_index=True)

        if len(piv) > topn:
            with st.expander("Ver todos os clientes (tabela completa)", expanded=False):
                st.dataframe(_format_tbl(piv), use_container_width=True, hide_index=True)
