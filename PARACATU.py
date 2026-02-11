# PARACATU.py
# Streamlit dashboard (DRE + DFC) para "projeto Paracatu.xlsx"
#
# Implementações:
# - 1 tabela por página
# - Colunas JAN | JAN% | ... | DEZ | DEZ% + ACUM | ACUM%
# - Shift (mês seguinte) para 00024 (Pessoal) e 00021 (Deduções) na DRE
# - Drill por linha (todas as linhas) + detalhamento de despesas (DESPESA/HISTÓRICO/FAVORECIDO)
# - Linhas extras:
#   * DRE: RESULTADO antes das Desp financeiras e RETIRADAS = RESULTADO OPERACIONAL + INVESTIMENTOS/RETIRADAS + DESPESAS FINANCEIRAS
#   * DFC: SALDO OPERACIONAL antes das Desp financeiras e RETIRADAS = SALDO OPERACIONAL + DESPESAS FINANCEIRAS + INVESTIMENTOS/RETIRADAS
# - Coloração: valores de linhas de RESULTADO/SALDO em azul quando positivo e vermelho quando negativo
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

import pandas as pd
import numpy as np
import streamlit as st


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
    preferred = ["projeto Paracatu.xlsx", "PROJETO PARACATU.xlsx", "Paracatu.xlsx"]
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

    for c in ["CONTA DE RESULTADO", "DESPESA", "FAVORECIDO", "HISTÓRICO"]:
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
    # valor
    if "VR. CONTÁBIL" in c.columns:
        c["_v"] = c["VR. CONTÁBIL"].apply(to_num)
    elif "VR.CONTÁBIL" in c.columns:
        c["_v"] = c["VR.CONTÁBIL"].apply(to_num)
    else:
        c["_v"] = 0.0
    # fornecedor
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
            if prefix == "00022":  # investimentos/retiradas
                inv_by_month[m] = float(by_m.get(m, 0.0))
            if prefix == "00023":  # despesas financeiras
                fin_by_month[m] = float(by_m.get(m, 0.0))

    resultado_oper_by_month = {m: float(margem_by_month.get(m, 0.0)) - float(desp_total_by_month.get(m, 0.0)) for m in range(1, 13)}

    # NOVA LINHA solicitada: resultado antes de fin/retiradas (soma de volta)
    resultado_antes_fin_ret_by_month = {m: float(resultado_oper_by_month[m]) + float(inv_by_month[m]) + float(fin_by_month[m]) for m in range(1, 13)}

    # Ordem das linhas
    linhas: List[Tuple[str, Dict[int, float], str]] = []
    linhas.append(("RECEITA", receita_by_month, "currency"))
    linhas.append(("CMV", cmv_by_month, "currency"))
    linhas.append(("MARGEM BRUTA", margem_by_month, "currency"))
    linhas.append(("MARKUP", markup_by_month, "ratio"))
    for nome, by_m, _p in despesas_map_by_month:
        linhas.append((nome, by_m, "currency"))
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

    # NOVA LINHA solicitada (add back fin/retiradas)
    saldo_antes_fin_ret_by_month = {m: float(saldo_by_month[m]) + float(fin_by_month[m]) + float(inv_by_month[m]) for m in range(1, 13)}

    linhas: List[Tuple[str, Dict[int, float], str]] = []
    linhas.append(("RECEBIMENTOS", receb_by_month, "currency"))
    for nome, by_m, _p in saidas_map_by_month:
        linhas.append((nome, by_m, "currency"))
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
    Compatível com pandas antigos (sem .hide_columns).
    - Formata valores em R$ (exceto MARKUP, que é numérico).
    - Formata percentuais em % (MARKUP% vira "—").
    - Colore linhas de resultado/saldo: azul se positivo, vermelho se negativo (por mês e ACUM).
    """
    cols_value = [MES_NUM_TO_PT[m] for m in meses_exib] + ["ACUM"]
    cols_pct = [f"{MES_NUM_TO_PT[m]}%" for m in meses_exib] + ["ACUM%"]
    cols_all = ["LINHA"] + sum([[MES_NUM_TO_PT[m], f"{MES_NUM_TO_PT[m]}%"] for m in meses_exib], []) + ["ACUM", "ACUM%"]

    base = df[cols_all + ["_type"]].copy()
    num_base = base.copy()  # para regras de cor

    # Constrói DataFrame já formatado (strings) para exibição
    show = base[cols_all].copy()

    for i in show.index:
        linha = str(show.loc[i, "LINHA"])
        typ = str(base.loc[i, "_type"])

        for c in cols_value:
            v = num_base.loc[i, c]
            if typ == "ratio":  # MARKUP
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

    # Coloração nas linhas destacadas
    def _row_style(row):
        linha = str(row["LINHA"])
        styles = [""] * len(row.index)
        if linha not in highlight_rows:
            return styles

        # pinta apenas colunas de valores (meses + ACUM)
        for j, col in enumerate(row.index):
            if col in cols_value:
                # valor numérico correspondente
                try:
                    v = float(num_base.loc[row.name, col])
                except Exception:
                    v = 0.0
                if v > 0:
                    styles[j] = "color: blue; font-weight: 700;"
                elif v < 0:
                    styles[j] = "color: red; font-weight: 700;"
        return styles

    sty = sty.apply(_row_style, axis=1)
    return sty


# =========================
# UI
# =========================
st.set_page_config(page_title="Indicadores Paracatu (DRE/DFC)", layout="wide")

excel_path = _auto_find_excel()
if not excel_path:
    st.error("Não encontrei nenhum Excel (.xlsx/.xlsm/.xls) na pasta do app. Coloque o 'projeto Paracatu.xlsx' junto do .py.")
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

# Compras fornecedor é opcional (só usado na página INDICADOR DE COMPRAS)

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
meses_exib = _date_filter_to_months(date_ini, date_fim, int(ano_ref))

st.sidebar.caption(f"Meses exibidos: **{', '.join(MES_NUM_TO_PT[m] for m in meses_exib)}**")

pagina = st.sidebar.radio("Página", ["DRE", "DFC", "INDICADOR DE COMPRAS"], index=0)


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
    st.dataframe(style_table(dre_tbl, meses_exib, highlight_rows), use_container_width=True)

    # =========================
    # Drill (todas as linhas)
    # =========================
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

    # =========================
    # Drill despesas (somente se for linha de despesa mapeada)
    # =========================
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
    st.dataframe(style_table(dfc_tbl, meses_exib, highlight_rows), use_container_width=True)

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
else:
    st.title("INDICADOR DE COMPRAS")

    if df_comp is None:
        st.error("A aba 'Compras fornecedor' não foi encontrada ou não pôde ser lida. Verifique o nome da aba no Excel.")
        st.stop()

    # séries mensais
    cmv_by_month = month_series(df_rcm, "_cmv", int(ano_ref), meses_exib)
    compras_by_month = month_series(df_comp, "_v", int(ano_ref), meses_exib)
    diff_by_month = {m: float(cmv_by_month.get(m, 0.0)) - float(compras_by_month.get(m, 0.0)) for m in range(1, 13)}

    # Tabela com meses por extenso + ACUM
    meses_full = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }

    cols = ["LINHA"] + [meses_full[m] for m in meses_exib] + ["ACUM"]
    rows = []

    def _row(nome, by_month):
        r = {"LINHA": nome}
        acum = 0.0
        for m in meses_exib:
            v = float(by_month.get(m, 0.0))
            r[meses_full[m]] = v
            acum += v
        r["ACUM"] = acum
        return r

    rows.append(_row("CMV NECESSIDADE", cmv_by_month))
    rows.append(_row("COMPRAS REALIZADO", compras_by_month))
    rows.append(_row("DIFERENÇA (CMV - COMPRAS)", diff_by_month))

    tbl = pd.DataFrame(rows)[cols]

    # Styler com R$ e cor na linha Diferença
    num_tbl = tbl.copy()

    def _fmt_brl(v):
        return f"R$ {format_brl(v)}"

    show = tbl.copy()
    for c in cols[1:]:
        show[c] = show[c].apply(lambda x: _fmt_brl(x))

    sty = show.style

    diff_idx = show.index[show["LINHA"] == "DIFERENÇA (CMV - COMPRAS)"].tolist()

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

    sty = sty.apply(_row_style, axis=1)

    st.subheader("Tabela — CMV x Compras")
    st.dataframe(sty, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Drill — Compras por fornecedor")

    mes_opt = ["TODOS"] + [meses_full[m] for m in meses_exib]
    mes_sel = st.selectbox("Mês (opcional)", options=mes_opt, index=0, key="compras_mes_sel")

    if mes_sel == "TODOS":
        meses_drill = meses_exib
    else:
        inv = {v: k for k, v in meses_full.items()}
        meses_drill = [inv[mes_sel]]

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

