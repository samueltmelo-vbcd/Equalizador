import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, timedelta
from pyxlsb import open_workbook
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

st.set_page_config(
    page_title="Equalização de Estoques",
    page_icon="📦",
    layout="wide",
)

# ── Estilos ───────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 1.4rem; }
.stAlert { border-radius: 8px; }
.block-container { padding-top: 2rem; }
</style>
""", unsafe_allow_html=True)

# ── Helpers de leitura ────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def read_xlsb_sheet(file_bytes, sheet_name):
    import tempfile, os
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsb") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        rows = []
        with open_workbook(tmp_path) as wb:
            with wb.get_sheet(sheet_name) as ws:
                for row in ws.rows():
                    rows.append([c.v for c in row])
        if not rows:
            return pd.DataFrame()
        return pd.DataFrame(rows[1:], columns=rows[0])
    finally:
        os.unlink(tmp_path)

@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes):
    import tempfile, os
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsb") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        with open_workbook(tmp_path) as wb:
            return wb.sheets
    finally:
        os.unlink(tmp_path)

def excel_serial_to_date(s):
    if pd.isna(s) or s <= 0:
        return None
    try:
        return date(1899, 12, 30) + timedelta(days=int(s))
    except:
        return None

# ── Lógica de análise ─────────────────────────────────────────────────
def calcular_ideal(row, dias):
    meses = dias / 30
    if row["GRADE"] > 0:
        return row["GRADE"], True
    elif row["CMM"] > 0:
        return row["CMM"] * meses, False
    else:
        return 0, False

def processar_base_estudos(df_raw, linhas_selecionadas, politica_dias):
    num_cols = ["FÍSICO", "CMM", "GRADE", "VALOR UNIT", "VALOR TOTAL ",
                "01/09/2025", "01/10/2025", "01/11/2025", "01/12/2025",
                "01/01/2026", "01/02/2026", "01/03/2026", "01/04/2026", "TOTAL CONSUMO"]
    df = df_raw[
        df_raw["LINHA"].isin(linhas_selecionadas) &
        (df_raw["CONSIDERAR NO ESTUDO?"] == "SIM")
    ].copy()
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    ideal_90  = df.apply(lambda r: calcular_ideal(r, 90),  axis=1)
    ideal_alt = df.apply(lambda r: calcular_ideal(r, politica_dias), axis=1)

    df["IDEAL_90"]     = ideal_90.apply(lambda x: x[0])
    df["IDEAL_ALT"]    = ideal_alt.apply(lambda x: x[0])
    df["ORIGEM_GRADE"] = ideal_90.apply(lambda x: x[1])
    df["EXCESSO_90"]   = df["FÍSICO"] - df["IDEAL_90"]
    df["EXCESSO_ALT"]  = df["FÍSICO"] - df["IDEAL_ALT"]
    df["VALOR_EXCESSO_90"]  = df["EXCESSO_90"].clip(lower=0)  * df["VALOR UNIT"]
    df["VALOR_EXCESSO_ALT"] = df["EXCESSO_ALT"].clip(lower=0) * df["VALOR UNIT"]

    cmm_rede = df.groupby("COD VALORES")["CMM"].sum().reset_index()
    cmm_rede.columns = ["COD VALORES", "CMM_REDE"]
    df = df.merge(cmm_rede, on="COD VALORES", how="left")

    def classificar(row):
        if row["FÍSICO"] == 0:
            return "SEM ESTOQUE"
        sufixo = " - GRADE" if row["ORIGEM_GRADE"] else ""
        if row["EXCESSO_90"] <= 0:
            return f"ADEQUADO/DEFICITARIO{sufixo}"
        if row["CMM_REDE"] > 0:
            return f"TRANSFERIVEL{sufixo}"
        return f"RETORNO_CD{sufixo}"

    df["CLASSIFICACAO"] = df.apply(classificar, axis=1)
    return df

def processar_lotes(df_lotes_raw, df_base):
    df_lotes = df_lotes_raw.copy()
    # Aceitar tanto base de lotes externa quanto CMV_2
    # Colunas esperadas: COD_PRODUTO / COD VALORES, HOSPITAL/Filial, LOTE/Lote fabricante, VALIDADE/Validade, QUANTIDADE/Quantidade
    col_map = {}
    for possivel, destino in [
        (["COD_PRODUTO","COD VALORES","Codigo do item","COD ARRUMAR"], "COD_STR"),
        (["HOSPITAL","Filial","SIGLA"], "FILIAL"),
        (["LOTE","Lote fabricante","Lote interno","LOTE_FABRICANTE"], "LOTE"),
        (["VALIDADE","Validade","DATA_VALIDADE"], "VALIDADE_RAW"),
        (["QUANTIDADE","Quantidade","QTDE","QTD"], "QUANTIDADE"),
    ]:
        for p in possivel:
            if p in df_lotes.columns:
                col_map[p] = destino
                break

    df_lotes = df_lotes.rename(columns=col_map)
    df_lotes["COD_STR"] = df_lotes["COD_STR"].astype(str).str.strip()
    df_lotes["QUANTIDADE"] = pd.to_numeric(df_lotes["QUANTIDADE"], errors="coerce").fillna(0)
    val_raw = pd.to_numeric(df_lotes["VALIDADE_RAW"], errors="coerce")
    df_lotes["VALIDADE_DT"] = val_raw.apply(excel_serial_to_date)

    skus = set(df_base["COD VALORES"].astype(str).str.strip())
    df_lotes = df_lotes[df_lotes["COD_STR"].isin(skus) & df_lotes["VALIDADE_DT"].notna()].copy()

    cmm_map = df_base.groupby(df_base["COD VALORES"].astype(str).str.strip())["CMM_REDE"].first()
    desc_map = df_base.groupby(df_base["COD VALORES"].astype(str).str.strip())["DESCRICAO"].first()
    df_lotes["CMM_REDE"] = df_lotes["COD_STR"].map(cmm_map).fillna(0)
    df_lotes["DESCRICAO"] = df_lotes["COD_STR"].map(desc_map)

    hoje = date.today()
    df_lotes["MESES_ATE_VENCER"] = df_lotes["VALIDADE_DT"].apply(
        lambda v: max(0, (v - hoje).days / 30))
    df_lotes["MESES_PARA_CONSUMIR"] = df_lotes.apply(
        lambda r: r["QUANTIDADE"] / r["CMM_REDE"] if r["CMM_REDE"] > 0 else np.inf, axis=1)
    df_lotes["SALDO_MESES"] = df_lotes["MESES_ATE_VENCER"] - df_lotes["MESES_PARA_CONSUMIR"]

    def status_ven(row):
        if row["CMM_REDE"] == 0:
            return f"SEM CONSUMO — VENCE EM {row['MESES_ATE_VENCER']:.0f}m"
        if row["SALDO_MESES"] < 0:
            return "VENCE ANTES DE CONSUMIR"
        elif row["SALDO_MESES"] <= 1:
            return "MARGEM APERTADA"
        return "OK"

    df_lotes["STATUS_VENCIMENTO"] = df_lotes.apply(status_ven, axis=1)
    return df_lotes

# ── Geração do Excel ──────────────────────────────────────────────────
CORES = {
    "TRANSFERIVEL":                  "D5F5E3",
    "TRANSFERIVEL - GRADE":          "A9DFBF",
    "RETORNO_CD":                    "FADBD8",
    "RETORNO_CD - GRADE":            "F1948A",
    "ADEQUADO/DEFICITARIO":          "EBF3FB",
    "ADEQUADO/DEFICITARIO - GRADE":  "D6EAF8",
    "SEM ESTOQUE":                   "F9F9F9",
}

def make_excel(df, df_lotes, politica_dias):
    thin = Side(style="thin", color="CCCCCC")
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr(ws, row, col, val, bg="1F3864", fg="FFFFFF", wrap=False, size=10):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(bold=True, color=fg, size=size, name="Arial")
        c.fill = PatternFill("solid", start_color=bg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
        c.border = brd
        return c

    def cel(ws, row, col, v, fmt=None, bg=None, bold=False, halign="left"):
        c = ws.cell(row=row, column=col, value=v)
        c.font = Font(bold=bold, color="000000", size=9, name="Arial")
        c.alignment = Alignment(horizontal=halign, vertical="center")
        if fmt: c.number_format = fmt
        if bg:  c.fill = PatternFill("solid", start_color=bg)
        c.border = brd
        return c

    def widths(ws, w):
        for i, v in enumerate(w, 1):
            ws.column_dimensions[get_column_letter(i)].width = v

    wb = Workbook()

    # ── Aba 1: Resumo ─────────────────────────────────────────────────
    ws1 = wb.active; ws1.title = "1. Resumo Executivo"
    hdr(ws1,1,1,f"EQUALIZAÇÃO DE ESTOQUES  |  {date.today().strftime('%d/%m/%Y')}  |  Política alternativa: {politica_dias}d",bg="1F3864",size=12)
    ws1.merge_cells("A1:G1"); ws1.row_dimensions[1].height=28

    kpis = [
        ("Valor total físico",                  df["VALOR TOTAL "].sum(),                                                    "R$ #,##0.00"),
        ("Excesso total (90d)",                  df["VALOR_EXCESSO_90"].sum(),                                                "R$ #,##0.00"),
        (f"Excesso total ({politica_dias}d)",    df["VALOR_EXCESSO_ALT"].sum(),                                               "R$ #,##0.00"),
        ("→ Transferível — consumo real",        df[df["CLASSIFICACAO"]=="TRANSFERIVEL"]["VALOR_EXCESSO_90"].sum(),           "R$ #,##0.00"),
        ("→ Transferível — excesso de grade",    df[df["CLASSIFICACAO"]=="TRANSFERIVEL - GRADE"]["VALOR_EXCESSO_90"].sum(),   "R$ #,##0.00"),
        ("→ Retorno ao CD — sem consumo",        df[df["CLASSIFICACAO"]=="RETORNO_CD"]["VALOR_EXCESSO_90"].sum(),             "R$ #,##0.00"),
        ("→ Retorno ao CD — grade sem consumo",  df[df["CLASSIFICACAO"]=="RETORNO_CD - GRADE"]["VALOR_EXCESSO_90"].sum(),     "R$ #,##0.00"),
        ("Hospitais analisados",                 df["HOSPITAL AJUSTADO"].nunique(),                                           "#,##0"),
        ("SKUs únicos",                          df["COD VALORES"].nunique(),                                                 "#,##0"),
    ]
    if df_lotes is not None:
        kpis += [
            ("Lotes com risco de vencer",  df_lotes[df_lotes["STATUS_VENCIMENTO"]=="VENCE ANTES DE CONSUMIR"].shape[0], "#,##0"),
            ("SKUs com risco de vencer",   df_lotes[df_lotes["STATUS_VENCIMENTO"]=="VENCE ANTES DE CONSUMIR"]["COD_STR"].nunique(), "#,##0"),
        ]

    hdr(ws1,3,1,"Indicador",bg="2E75B6"); hdr(ws1,3,2,"Valor",bg="2E75B6")
    for i,(label,value,fmt) in enumerate(kpis,4):
        bg="EBF3FB" if i%2==0 else "FFFFFF"
        bold = "→" not in label
        cel(ws1,i,1,label,bg=bg,bold=bold); cel(ws1,i,2,value,fmt=fmt,bg=bg,bold=bold,halign="right")

    reg = df.groupby("REGIONAL").agg(VF=("VALOR TOTAL ","sum"),EX=("VALOR_EXCESSO_90","sum")).reset_index().sort_values("EX",ascending=False)
    hdr(ws1,3,4,"Regional",bg="2E75B6"); hdr(ws1,3,5,"Valor Físico",bg="2E75B6")
    hdr(ws1,3,6,"Excesso 90d",bg="2E75B6"); hdr(ws1,3,7,"% Excesso",bg="2E75B6")
    for i,(_,r) in enumerate(reg.iterrows(),4):
        bg="EBF3FB" if i%2==0 else "FFFFFF"
        cel(ws1,i,4,r["REGIONAL"],bg=bg); cel(ws1,i,5,r["VF"],fmt="R$ #,##0.00",bg=bg,halign="right")
        cel(ws1,i,6,r["EX"],fmt="R$ #,##0.00",bg=bg,halign="right")
        cel(ws1,i,7,r["EX"]/r["VF"] if r["VF"] else 0,fmt="0.0%",bg=bg,halign="right")

    widths(ws1,[38,18,4,14,18,18,12])

    # ── Aba 2: Transferências ─────────────────────────────────────────
    ws2 = wb.create_sheet("2. Transferências")
    hdr(ws2,1,1,"ESTOQUE PARA TRANSFERÊNCIA  |  verde escuro = excesso de grade (sem consumo real)",bg="1F3864",size=11)
    ws2.merge_cells("A1:N1")
    cols2 = ["Regional","Hospital","Sigla","Cód. Produto","Descrição","Classificação","Origem Ideal",
             "Físico",f"Ideal 90d",f"Ideal {politica_dias}d","Excesso 90d",f"Excesso {politica_dias}d","Valor Unit.","Valor Excesso 90d"]
    fmts2 = [None,None,None,None,None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","R$ #,##0.00","R$ #,##0.00"]
    for j,h in enumerate(cols2,1): hdr(ws2,2,j,h,bg="2E75B6",wrap=True)
    transf = df[df["CLASSIFICACAO"].str.startswith("TRANSFERIVEL")].sort_values("VALOR_EXCESSO_90",ascending=False)
    for i,(_,r) in enumerate(transf.iterrows(),3):
        bg=CORES.get(r["CLASSIFICACAO"],"FFFFFF")
        orig="Grade" if r["ORIGEM_GRADE"] else "CMM (consumo)"
        vals=[r["REGIONAL"],r["HOSPITAL AJUSTADO"],r["SIGLA"],str(r["COD VALORES"]),r["DESCRICAO"],
              r["CLASSIFICACAO"],orig,r["FÍSICO"],r["IDEAL_90"],r["IDEAL_ALT"],
              r["EXCESSO_90"],r["EXCESSO_ALT"],r["VALOR UNIT"],r["VALOR_EXCESSO_90"]]
        for j,(v,f) in enumerate(zip(vals,fmts2),1): cel(ws2,i,j,v,fmt=f,bg=bg,halign="right" if f else "left")
    widths(ws2,[10,22,7,16,38,26,14,10,10,10,10,10,12,16])
    ws2.freeze_panes="A3"

    # ── Aba 3: Retorno ao CD ──────────────────────────────────────────
    ws3 = wb.create_sheet("3. Retorno ao CD")
    hdr(ws3,1,1,"RETORNO AO CD  |  vermelho escuro = sem consumo nem grade que justifique",bg="7B2C2C",size=11)
    ws3.merge_cells("A1:N1")
    cols3=["Regional","Hospital","Sigla","Cód. Produto","Descrição","Classificação","Origem Ideal",
           "Físico","Ideal 90d","Excesso 90d","CMM Hosp.","CMM Rede","Valor Unit.","Valor a Retornar"]
    fmts3=[None,None,None,None,None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","R$ #,##0.00","R$ #,##0.00"]
    for j,h in enumerate(cols3,1): hdr(ws3,2,j,h,bg="C0392B",wrap=True)
    retorno=df[df["CLASSIFICACAO"].str.startswith("RETORNO_CD")].sort_values("VALOR_EXCESSO_90",ascending=False)
    for i,(_,r) in enumerate(retorno.iterrows(),3):
        bg=CORES.get(r["CLASSIFICACAO"],"FFFFFF")
        orig="Grade" if r["ORIGEM_GRADE"] else "CMM (consumo)"
        vals=[r["REGIONAL"],r["HOSPITAL AJUSTADO"],r["SIGLA"],str(r["COD VALORES"]),r["DESCRICAO"],
              r["CLASSIFICACAO"],orig,r["FÍSICO"],r["IDEAL_90"],r["EXCESSO_90"],
              r["CMM"],r["CMM_REDE"],r["VALOR UNIT"],r["VALOR_EXCESSO_90"]]
        for j,(v,f) in enumerate(zip(vals,fmts3),1): cel(ws3,i,j,v,fmt=f,bg=bg,halign="right" if f else "left")
    widths(ws3,[10,22,7,16,38,26,14,10,10,10,10,10,12,16])
    ws3.freeze_panes="A3"

    # ── Aba 4: Vencimento (opcional) ──────────────────────────────────
    if df_lotes is not None:
        ws4 = wb.create_sheet("4. Risco de Vencimento")
        hdr(ws4,1,1,"ANÁLISE DE VENCIMENTO POR LOTE  |  risco = vence antes de conseguir consumir",bg="6E2D2D",size=11)
        ws4.merge_cells("A1:L1")
        cols4=["CD/Filial","Cód. Produto","Descrição","Lote","Validade","Dias até Vencer",
               "Meses até Vencer","Qtd. Lote","CMM Rede","Meses p/ Consumir","Saldo (meses)","Status"]
        fmts4=[None,None,None,None,"DD/MM/YYYY","#,##0","#,##0.1","#,##0.00","#,##0.00","#,##0.1","#,##0.1",None]
        for j,h in enumerate(cols4,1): hdr(ws4,2,j,h,bg="922B21",wrap=True)
        df_r=df_lotes[~df_lotes["STATUS_VENCIMENTO"].str.startswith("OK")].sort_values(["STATUS_VENCIMENTO","MESES_ATE_VENCER"])
        hoje=date.today()
        for i,(_,r) in enumerate(df_r.iterrows(),3):
            st=str(r["STATUS_VENCIMENTO"])
            bg="F1948A" if st=="VENCE ANTES DE CONSUMIR" else ("FAD7A0" if st=="MARGEM APERTADA" else "FDEDEC")
            dias=(r["VALIDADE_DT"]-hoje).days if r["VALIDADE_DT"] else None
            mc=r["MESES_PARA_CONSUMIR"] if r["MESES_PARA_CONSUMIR"]!=np.inf else None
            saldo=r["SALDO_MESES"] if r["MESES_PARA_CONSUMIR"]!=np.inf else None
            lote_val=r.get("LOTE", r.get("Lote fabricante",""))
            filial_val=r.get("FILIAL", r.get("Filial",""))
            vals=[filial_val,r["COD_STR"],r.get("DESCRICAO",""),lote_val,
                  r["VALIDADE_DT"],dias,r["MESES_ATE_VENCER"],r["QUANTIDADE"],r["CMM_REDE"],mc,saldo,st]
            for j,(v,f) in enumerate(zip(vals,fmts4),1):
                c=cel(ws4,i,j,v,fmt=f,bg=bg,halign="right" if f else "left")
                if j==12 and st=="VENCE ANTES DE CONSUMIR": c.font=Font(bold=True,color="000000",size=9,name="Arial")
        widths(ws4,[10,16,38,18,13,12,14,12,10,14,12,24])
        ws4.freeze_panes="A3"

    # ── Aba 5: Por Hospital ───────────────────────────────────────────
    wsn = wb.create_sheet("5. Por Hospital")
    hdr(wsn,1,1,"RESUMO POR HOSPITAL",bg="1F3864",size=11)
    ws1.merge_cells("A1:I1")
    cols5=["Regional","Hospital","Sigla","Valor Físico","Excesso 90d","% Excesso","SKUs Total","SKUs Sem Consumo","SKUs Retorno CD"]
    fmts5=[None,None,None,"R$ #,##0.00","R$ #,##0.00","0.0%","#,##0","#,##0","#,##0"]
    for j,h in enumerate(cols5,1): hdr(wsn,2,j,h,bg="2E75B6",wrap=True)
    ret_hosp=df[df["CLASSIFICACAO"].str.startswith("RETORNO_CD")].groupby("HOSPITAL AJUSTADO").size()
    resumo=df.groupby(["HOSPITAL AJUSTADO","SIGLA","REGIONAL"]).agg(
        VF=("VALOR TOTAL ","sum"),EX=("VALOR_EXCESSO_90","sum"),
        SKUS=("COD VALORES","count"),SEM=("CMM",lambda x:(x==0).sum())
    ).reset_index().sort_values("EX",ascending=False)
    resumo["RET"]=resumo["HOSPITAL AJUSTADO"].map(ret_hosp).fillna(0).astype(int)
    for i,(_,r) in enumerate(resumo.iterrows(),3):
        bg="EBF3FB" if i%2==0 else "FFFFFF"
        pct=r["EX"]/r["VF"] if r["VF"] else 0
        vals=[r["REGIONAL"],r["HOSPITAL AJUSTADO"],r["SIGLA"],r["VF"],r["EX"],pct,r["SKUS"],r["SEM"],r["RET"]]
        for j,(v,f) in enumerate(zip(vals,fmts5),1):
            c=cel(wsn,i,j,v,fmt=f,bg=bg,halign="right" if f else "left")
            if j==6 and pct>0.6: c.fill=PatternFill("solid",start_color="FADBD8")
    widths(wsn,[10,26,7,18,16,12,10,14,14])
    wsn.freeze_panes="A3"

    # ── Aba 6: Base completa ──────────────────────────────────────────
    ws6=wb.create_sheet("6. Base Completa")
    cols_b=["HOSPITAL AJUSTADO","SIGLA","REGIONAL","COD VALORES","DESCRICAO","LINHA",
            "CLASSIFICACAO","ORIGEM_GRADE","FÍSICO","IDEAL_90","IDEAL_ALT",
            "EXCESSO_90","EXCESSO_ALT","VALOR UNIT","VALOR TOTAL ","VALOR_EXCESSO_90",
            "CMM","CMM_REDE","GRADE","STATUS GRADE","STATUS"]
    hdrs_b=["Hospital","Sigla","Regional","Cód. Produto","Descrição","Linha",
            "Classificação","Ideal veio da Grade?","Físico","Ideal 90d",f"Ideal {politica_dias}d",
            "Excesso 90d",f"Excesso {politica_dias}d","Valor Unit.","Valor Total","Valor Excesso 90d",
            "CMM Hosp.","CMM Rede","Grade","Status Grade","Status Consumo"]
    fmts_b=[None,None,None,None,None,None,None,None,
            "#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00",
            "R$ #,##0.00","R$ #,##0.00","R$ #,##0.00","#,##0.00","#,##0.00","#,##0.00",None,None]
    for j,h in enumerate(hdrs_b,1): hdr(ws6,1,j,h,bg="2E75B6",wrap=True)
    df_b=df[cols_b].copy()
    df_b["ORIGEM_GRADE"]=df_b["ORIGEM_GRADE"].map({True:"SIM",False:"NÃO"})
    df_b=df_b.sort_values(["CLASSIFICACAO","VALOR_EXCESSO_90"],ascending=[True,False])
    for i,(_,r) in enumerate(df_b.iterrows(),2):
        bg=CORES.get(r["CLASSIFICACAO"],"FFFFFF")
        for j,(col,f) in enumerate(zip(cols_b,fmts_b),1):
            cel(ws6,i,j,r[col],fmt=f,bg=bg,halign="right" if f else "left")
    widths(ws6,[22,7,10,16,38,16,26,16,10,10,10,10,10,12,14,16,10,10,8,16,18])
    ws6.freeze_panes="A2"

    buf=io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Interface ─────────────────────────────────────────────────────────
st.title("📦 Equalização de Estoques")
st.caption("Análise de redistribuição e retorno de estoques consignados")

with st.sidebar:
    st.header("⚙️ Configurações")
    politica_dias = st.selectbox("Política alternativa de cobertura", [120, 150, 180], index=1,
                                  help="90 dias é sempre calculado. Aqui você define a política adicional.")
    st.divider()
    st.markdown("**Legenda de cores**")
    for clf, cor, desc in [
        ("Transferível","#D5F5E3","Excesso real — consumo justifica"),
        ("Transferível - Grade","#A9DFBF","Excesso de grade — sem consumo real"),
        ("Retorno CD","#FADBD8","Sem consumo na rede"),
        ("Retorno CD - Grade","#F1948A","Sem consumo nem grade"),
        ("Adequado/Deficitário","#EBF3FB","Estoque dentro do ideal"),
    ]:
        st.markdown(f'<span style="background:{cor};padding:2px 8px;border-radius:4px;font-size:0.8rem">{clf}</span> {desc}', unsafe_allow_html=True)

# Etapa 1: Upload da base principal
st.subheader("1. Base de estoques (.xlsb)")
file_base = st.file_uploader("Faça upload do arquivo de estudo (BASE_ESTUDOS)", type=["xlsb"],
                              help="Arquivo com aba BASE_ESTUDOS contendo estoque físico, CMM, grade e flags de qualificação")

if file_base:
    file_bytes = file_base.read()
    with st.spinner("Lendo abas do arquivo..."):
        sheets = get_sheet_names(file_bytes)

    if "BASE_ESTUDOS" not in sheets:
        st.error(f"Aba BASE_ESTUDOS não encontrada. Abas disponíveis: {sheets}")
        st.stop()

    with st.spinner("Carregando BASE_ESTUDOS..."):
        df_raw = read_xlsb_sheet(file_bytes, "BASE_ESTUDOS")

    num_cols_raw = ["FÍSICO","CMM","GRADE","VALOR UNIT","VALOR TOTAL "]
    for c in num_cols_raw:
        if c in df_raw.columns:
            df_raw[c] = pd.to_numeric(df_raw[c], errors="coerce").fillna(0)

    linhas_disponiveis = sorted(df_raw["LINHA"].dropna().unique().tolist())
    col1, col2 = st.columns([2,1])
    with col1:
        linhas_sel = st.multiselect("Linhas de produto para analisar", linhas_disponiveis,
                                     default=["HEMODINAMICA"] if "HEMODINAMICA" in linhas_disponiveis else linhas_disponiveis[:1])
    with col2:
        regionais_disp = sorted(df_raw["REGIONAL"].dropna().unique().tolist())
        regionais_sel  = st.multiselect("Filtrar regionais (opcional)", regionais_disp, default=[])

    if not linhas_sel:
        st.warning("Selecione ao menos uma linha de produto.")
        st.stop()

    df_filtrado = df_raw if not regionais_sel else df_raw[df_raw["REGIONAL"].isin(regionais_sel)]

    # Etapa 2: Upload de lotes (opcional)
    st.subheader("2. Base de lotes e validades (opcional)")
    file_lotes = st.file_uploader("Upload da base de lotes (.xlsx ou .csv)", type=["xlsx","csv"],
                                   help="Colunas esperadas: COD_PRODUTO, HOSPITAL, LOTE, VALIDADE, QUANTIDADE")
    df_lotes = None
    usar_cmv2 = False
    if file_lotes is None and "CMV_2" in sheets:
        usar_cmv2 = st.checkbox("Usar aba CMV_2 do próprio arquivo (lotes do consignado)", value=True)

    # Processar
    if st.button("▶ Rodar análise", type="primary", use_container_width=True):
        with st.spinner("Processando..."):
            df = processar_base_estudos(df_filtrado, linhas_sel, politica_dias)

            # Lotes
            if file_lotes:
                if file_lotes.name.endswith(".csv"):
                    df_lotes_raw = pd.read_csv(file_lotes)
                else:
                    df_lotes_raw = pd.read_excel(file_lotes)
                df_lotes = processar_lotes(df_lotes_raw, df)
            elif usar_cmv2:
                with st.spinner("Carregando CMV_2..."):
                    df_cmv2 = read_xlsb_sheet(file_bytes, "CMV_2")
                df_lotes = processar_lotes(df_cmv2, df)

        # ── Métricas ──────────────────────────────────────────────────
        st.divider()
        st.subheader("📊 Resultados")
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Valor físico total",    f"R$ {df['VALOR TOTAL '].sum():,.0f}")
        c2.metric("Excesso (90d)",         f"R$ {df['VALOR_EXCESSO_90'].sum():,.0f}")
        c3.metric("Transferível",          f"R$ {df[df['CLASSIFICACAO'].str.startswith('TRANSFERIVEL')]['VALOR_EXCESSO_90'].sum():,.0f}")
        c4.metric("Retorno ao CD",         f"R$ {df[df['CLASSIFICACAO'].str.startswith('RETORNO_CD')]['VALOR_EXCESSO_90'].sum():,.0f}")

        if df_lotes is not None:
            c5,c6,_,_ = st.columns(4)
            n_risco = df_lotes[df_lotes["STATUS_VENCIMENTO"]=="VENCE ANTES DE CONSUMIR"].shape[0]
            s_risco = df_lotes[df_lotes["STATUS_VENCIMENTO"]=="VENCE ANTES DE CONSUMIR"]["COD_STR"].nunique()
            c5.metric("Lotes em risco de vencer", f"{n_risco:,}")
            c6.metric("SKUs em risco",             f"{s_risco:,}")

        # ── Tabela preview ────────────────────────────────────────────
        st.subheader("Prévia — classificação por item")
        clf_counts = df["CLASSIFICACAO"].value_counts().reset_index()
        clf_counts.columns=["Classificação","Qtd. itens"]
        st.dataframe(clf_counts, use_container_width=True, hide_index=True)

        # ── Download ──────────────────────────────────────────────────
        st.divider()
        excel_buf = make_excel(df, df_lotes, politica_dias)
        nome_arquivo = f"Equalizacao_{'-'.join(linhas_sel)}_{date.today().strftime('%Y%m%d')}.xlsx"
        st.download_button(
            label="⬇️ Baixar Excel completo",
            data=excel_buf,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
