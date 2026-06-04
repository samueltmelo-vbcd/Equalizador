import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, timedelta, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from python_calamine import CalamineWorkbook
import io, tempfile, os

st.set_page_config(
    page_title="Equalização de Estoques",
    page_icon="📦",
    layout="wide",
)

st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 1.4rem; }
.block-container { padding-top: 2rem; }
</style>
""", unsafe_allow_html=True)

# ── Helpers de leitura ────────────────────────────────────────────────
def save_tmp(file_bytes, suffix=".xlsb"):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(file_bytes)
    tmp.close()
    return tmp.name

@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes):
    path = save_tmp(file_bytes)
    try:
        return CalamineWorkbook.from_path(path).sheet_names
    finally:
        os.unlink(path)

@st.cache_data(show_spinner=False)
def read_sheet(file_bytes, sheet_name):
    path = save_tmp(file_bytes)
    try:
        return pd.read_excel(path, sheet_name=sheet_name, engine="calamine", dtype=str)
    finally:
        os.unlink(path)

def to_float(s):
    try:
        return float(str(s).replace(",", ".").strip())
    except:
        return 0.0

def parse_date(v):
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return None
    if isinstance(v, (datetime, date)):
        return v.date() if isinstance(v, datetime) else v
    # tenta string DD/MM/YYYY ou YYYY-MM-DD
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%Y"):
        try:
            return datetime.strptime(str(v).strip()[:10], fmt[:len(fmt)]).date()
        except:
            continue
    return None

# ── Lógica de análise ─────────────────────────────────────────────────
def processar_base_estudos(df_raw, linhas_sel, politica_dias):
    df = df_raw[
        df_raw["LINHA"].isin(linhas_sel) &
        (df_raw["CONSIDERAR NO ESTUDO?"].str.strip() == "SIM")
    ].copy()

    for c in ["FÍSICO", "CMM", "GRADE", "VALOR UNIT", "VALOR TOTAL "]:
        if c in df.columns:
            df[c] = df[c].apply(to_float)

    def ideal(row, dias):
        m = dias / 30
        if row["GRADE"] > 0:
            return row["GRADE"], True
        elif row["CMM"] > 0:
            return row["CMM"] * m, False
        return 0.0, False

    r90  = df.apply(lambda r: ideal(r, 90),          axis=1)
    ralt = df.apply(lambda r: ideal(r, politica_dias), axis=1)

    df["IDEAL_90"]        = r90.apply(lambda x: x[0])
    df["IDEAL_ALT"]       = ralt.apply(lambda x: x[0])
    df["ORIGEM_GRADE"]    = r90.apply(lambda x: x[1])
    df["EXCESSO_90"]      = df["FÍSICO"] - df["IDEAL_90"]
    df["EXCESSO_ALT"]     = df["FÍSICO"] - df["IDEAL_ALT"]
    df["VALOR_EXCESSO_90"]  = df["EXCESSO_90"].clip(lower=0)  * df["VALOR UNIT"]
    df["VALOR_EXCESSO_ALT"] = df["EXCESSO_ALT"].clip(lower=0) * df["VALOR UNIT"]

    cmm_rede = df.groupby("COD VALORES")["CMM"].sum().rename("CMM_REDE")
    df = df.join(cmm_rede, on="COD VALORES")

    def clf(row, excesso_col, politica_label=""):
        if row["FÍSICO"] == 0:
            return "SEM ESTOQUE"
        sufixo = " - GRADE" if row["ORIGEM_GRADE"] else ""
        if row[excesso_col] <= 0:
            return f"ADEQUADO/DEFICITARIO{sufixo}"
        if row["CMM_REDE"] > 0:
            return f"TRANSFERIVEL{sufixo}"
        return f"RETORNO_CD{sufixo}"

    df["CLASSIFICACAO"]     = df.apply(lambda r: clf(r, "EXCESSO_90"),  axis=1)
    df["CLASSIFICACAO_ALT"] = df.apply(lambda r: clf(r, "EXCESSO_ALT"), axis=1)
    return df

def processar_lotes(df_raw, df_base):
    df = df_raw.copy()
    col_map = {}
    for candidates, dest in [
        (["COD_PRODUTO","COD VALORES","Codigo do item","COD ARRUMAR"], "COD_STR"),
        (["HOSPITAL","Filial","SIGLA"],                                 "FILIAL"),
        (["LOTE","Lote fabricante","Lote interno"],                     "LOTE"),
        (["VALIDADE","Validade","DATA_VALIDADE"],                       "VALIDADE_RAW"),
        (["QUANTIDADE","Quantidade","QTDE","QTD"],                      "QUANTIDADE"),
    ]:
        for c in candidates:
            if c in df.columns:
                col_map[c] = dest
                break
    df = df.rename(columns=col_map)
    df["COD_STR"]    = df["COD_STR"].astype(str).str.strip()
    df["QUANTIDADE"] = df["QUANTIDADE"].apply(to_float)
    df["VALIDADE_DT"] = df["VALIDADE_RAW"].apply(parse_date)

    skus = set(df_base["COD VALORES"].astype(str).str.strip())
    df = df[df["COD_STR"].isin(skus) & df["VALIDADE_DT"].notna()].copy()

    cmm_map  = df_base.groupby(df_base["COD VALORES"].astype(str).str.strip())["CMM_REDE"].first()
    desc_map = df_base.groupby(df_base["COD VALORES"].astype(str).str.strip())["DESCRICAO"].first()
    df["CMM_REDE"] = df["COD_STR"].map(cmm_map).fillna(0).apply(to_float)
    df["DESCRICAO"] = df["COD_STR"].map(desc_map)

    hoje = date.today()
    df["MESES_ATE_VENCER"]    = df["VALIDADE_DT"].apply(lambda v: max(0, (v - hoje).days / 30))
    df["MESES_PARA_CONSUMIR"] = df.apply(
        lambda r: r["QUANTIDADE"] / r["CMM_REDE"] if r["CMM_REDE"] > 0 else np.inf, axis=1)
    df["SALDO_MESES"] = df["MESES_ATE_VENCER"] - df["MESES_PARA_CONSUMIR"]

    def status(row):
        if row["CMM_REDE"] == 0:
            return f"SEM CONSUMO — VENCE EM {row['MESES_ATE_VENCER']:.0f}m"
        if row["SALDO_MESES"] < 0:
            return "VENCE ANTES DE CONSUMIR"
        if row["SALDO_MESES"] <= 1:
            return "MARGEM APERTADA"
        return "OK"

    df["STATUS_VENCIMENTO"] = df.apply(status, axis=1)
    return df

# ── Excel output ──────────────────────────────────────────────────────
CORES = {
    "TRANSFERIVEL":                 "D5F5E3",
    "TRANSFERIVEL - GRADE":         "A9DFBF",
    "RETORNO_CD":                   "FADBD8",
    "RETORNO_CD - GRADE":           "F1948A",
    "ADEQUADO/DEFICITARIO":         "EBF3FB",
    "ADEQUADO/DEFICITARIO - GRADE": "D6EAF8",
    "SEM ESTOQUE":                  "F9F9F9",
}

def make_excel(df, df_lotes, politica_dias):
    thin = Side(style="thin", color="CCCCCC")
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr(ws, row, col, val, bg="1F3864", fg="FFFFFF", size=10, wrap=False):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(bold=True, color=fg, size=size, name="Arial")
        c.fill = PatternFill("solid", start_color=bg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
        c.border = brd

    def cel(ws, row, col, v, fmt=None, bg=None, bold=False, halign="left"):
        c = ws.cell(row=row, column=col, value=v)
        c.font = Font(bold=bold, color="000000", size=9, name="Arial")
        c.alignment = Alignment(horizontal=halign, vertical="center")
        if fmt: c.number_format = fmt
        if bg:  c.fill = PatternFill("solid", start_color=bg)
        c.border = brd
        return c

    def W(ws, widths):
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    wb = Workbook()

    # ── Aba 1: Resumo ─────────────────────────────────────────────────
    ws1 = wb.active; ws1.title = "1. Resumo Executivo"
    ws1.row_dimensions[1].height = 28
    hdr(ws1,1,1,f"EQUALIZAÇÃO DE ESTOQUES  |  {date.today().strftime('%d/%m/%Y')}  |  Política alt.: {politica_dias}d",size=12)
    ws1.merge_cells("A1:I1")

    def kpi_bloco(col_ini, label_dias, exc_col, clf_col):
        hdr(ws1,3,col_ini,  f"Indicador — {label_dias}", bg="2E75B6", size=9)
        hdr(ws1,3,col_ini+1,"Valor",                     bg="2E75B6", size=9)
        kpis = [
            ("Valor físico total",               df["VALOR TOTAL "].sum(),                                                    "R$ #,##0.00"),
            (f"Excesso total ({label_dias})",     df[exc_col].sum(),                                                           "R$ #,##0.00"),
            ("→ Transferível — consumo real",     df[df[clf_col]=="TRANSFERIVEL"][exc_col].sum(),                              "R$ #,##0.00"),
            ("→ Transferível — grade",            df[df[clf_col]=="TRANSFERIVEL - GRADE"][exc_col].sum(),                      "R$ #,##0.00"),
            ("→ Retorno CD — sem consumo",        df[df[clf_col]=="RETORNO_CD"][exc_col].sum(),                                "R$ #,##0.00"),
            ("→ Retorno CD — grade sem consumo",  df[df[clf_col]=="RETORNO_CD - GRADE"][exc_col].sum(),                       "R$ #,##0.00"),
            ("Hospitais analisados",              df["HOSPITAL AJUSTADO"].nunique(),                                           "#,##0"),
            ("SKUs únicos",                       df["COD VALORES"].nunique(),                                                 "#,##0"),
        ]
        if df_lotes is not None:
            kpis += [
                ("Lotes com risco de vencer", df_lotes[df_lotes["STATUS_VENCIMENTO"]=="VENCE ANTES DE CONSUMIR"].shape[0], "#,##0"),
                ("SKUs com risco de vencer",  df_lotes[df_lotes["STATUS_VENCIMENTO"]=="VENCE ANTES DE CONSUMIR"]["COD_STR"].nunique(), "#,##0"),
            ]
        for i,(lbl,val,fmt) in enumerate(kpis,4):
            bg = "EBF3FB" if i%2==0 else "FFFFFF"
            bold = "→" not in lbl
            cel(ws1,i,col_ini,  lbl, bg=bg, bold=bold, halign="left")
            cel(ws1,i,col_ini+1,val, fmt=fmt, bg=bg, bold=bold, halign="right")

    kpi_bloco(1, "90d",  "VALOR_EXCESSO_90",  "CLASSIFICACAO")
    kpi_bloco(4, f"{politica_dias}d", "VALOR_EXCESSO_ALT", "CLASSIFICACAO_ALT")

    hdr(ws1,3,7,"Regional",bg="2E75B6",size=9)
    hdr(ws1,3,8,"Excesso 90d",bg="2E75B6",size=9)
    hdr(ws1,3,9,f"Excesso {politica_dias}d",bg="2E75B6",size=9)
    reg = df.groupby("REGIONAL").agg(E90=("VALOR_EXCESSO_90","sum"),EALT=("VALOR_EXCESSO_ALT","sum")).reset_index().sort_values("E90",ascending=False)
    for i,(_,r) in enumerate(reg.iterrows(),4):
        bg="EBF3FB" if i%2==0 else "FFFFFF"
        cel(ws1,i,7,r["REGIONAL"],bg=bg)
        cel(ws1,i,8,r["E90"], fmt="R$ #,##0.00",bg=bg,halign="right")
        cel(ws1,i,9,r["EALT"],fmt="R$ #,##0.00",bg=bg,halign="right")

    row_leg = 16
    hdr(ws1,row_leg,1,"Legenda",bg="2E75B6",size=9); ws1.merge_cells(f"A{row_leg}:I{row_leg}")
    for j,(clf,cor,desc) in enumerate([
        ("TRANSFERIVEL","D5F5E3","Excesso real — consumo justifica; redirecionar para hospital com déficit"),
        ("TRANSFERIVEL - GRADE","A9DFBF","Excesso de grade — acima do contrato, sem consumo real"),
        ("RETORNO_CD","FADBD8","Sem consumo na rede — retornar ao CD para tratativa com fornecedor"),
        ("RETORNO_CD - GRADE","F1948A","Sem consumo nem grade — retornar com prioridade máxima"),
        ("ADEQUADO/DEFICITARIO","EBF3FB","Estoque dentro ou abaixo do ideal"),
        ("ADEQUADO/DEFICITARIO - GRADE","D6EAF8","Adequado conforme grade contratual"),
    ],row_leg+1):
        cel(ws1,j,1,clf,bg=cor,bold=True); cel(ws1,j,2,desc,bg="FFFFFF",halign="left")
        ws1.merge_cells(f"B{j}:I{j}")
    W(ws1,[36,18,3,36,18,3,12,18,18])

    # ── Aba 2: Transferências 90d ─────────────────────────────────────
    def aba_transf(title, clf_col, exc_col, exc_label, bg_hdr, sheet_name):
        ws = wb.create_sheet(sheet_name)
        hdr(ws,1,1,title,bg="1F3864",size=11); ws.merge_cells("A1:N1")
        cols=["Regional","Hospital","Sigla","Cód. Produto","Descrição","Classificação","Origem Ideal",
              "Físico","Ideal 90d",f"Ideal {politica_dias}d","Excesso 90d",f"Excesso {politica_dias}d","Valor Unit.",exc_label]
        fmts=[None,None,None,None,None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","R$ #,##0.00","R$ #,##0.00"]
        for j,h in enumerate(cols,1): hdr(ws,2,j,h,bg=bg_hdr,wrap=True)
        data = df[df[clf_col].str.startswith("TRANSFERIVEL")].sort_values(exc_col,ascending=False)
        for i,(_,r) in enumerate(data.iterrows(),3):
            bg=CORES.get(r[clf_col],"FFFFFF")
            orig="Grade" if r["ORIGEM_GRADE"] else "CMM (consumo)"
            vals=[r["REGIONAL"],r["HOSPITAL AJUSTADO"],r["SIGLA"],str(r["COD VALORES"]),r["DESCRICAO"],
                  r[clf_col],orig,r["FÍSICO"],r["IDEAL_90"],r["IDEAL_ALT"],
                  r["EXCESSO_90"],r["EXCESSO_ALT"],r["VALOR UNIT"],r[exc_col]]
            for j,(v,f) in enumerate(zip(vals,fmts),1): cel(ws,i,j,v,fmt=f,bg=bg,halign="right" if f else "left")
        W(ws,[10,22,7,16,38,26,14,10,10,10,10,10,12,16])
        ws.freeze_panes="A3"

    aba_transf("TRANSFERÊNCIAS — política 90 dias  |  verde escuro = excesso de grade",
               "CLASSIFICACAO","VALOR_EXCESSO_90","Valor Excesso 90d","2E75B6","2. Transferências 90d")
    aba_transf(f"TRANSFERÊNCIAS — política {politica_dias} dias",
               "CLASSIFICACAO_ALT","VALOR_EXCESSO_ALT",f"Valor Excesso {politica_dias}d","155A8A",f"3. Transferências {politica_dias}d")

    # ── Aba 4: Retorno ao CD ──────────────────────────────────────────
    ws4 = wb.create_sheet("4. Retorno ao CD")
    hdr(ws4,1,1,"RETORNO AO CD  |  igual nas duas políticas — CMM zero na rede",bg="7B2C2C",size=11); ws4.merge_cells("A1:O1")
    cols4=["Regional","Hospital","Sigla","Cód. Produto","Descrição","Classificação","Origem Ideal",
           "Físico","Ideal 90d",f"Ideal {politica_dias}d","Excesso 90d",f"Excesso {politica_dias}d","CMM Rede","Valor Unit.","Valor a Retornar"]
    fmts4=[None,None,None,None,None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","R$ #,##0.00","R$ #,##0.00"]
    for j,h in enumerate(cols4,1): hdr(ws4,2,j,h,bg="C0392B",wrap=True)
    retorno=df[df["CLASSIFICACAO"].str.startswith("RETORNO_CD")].sort_values("VALOR_EXCESSO_90",ascending=False)
    for i,(_,r) in enumerate(retorno.iterrows(),3):
        bg=CORES.get(r["CLASSIFICACAO"],"FFFFFF")
        orig="Grade" if r["ORIGEM_GRADE"] else "CMM (consumo)"
        vals=[r["REGIONAL"],r["HOSPITAL AJUSTADO"],r["SIGLA"],str(r["COD VALORES"]),r["DESCRICAO"],
              r["CLASSIFICACAO"],orig,r["FÍSICO"],r["IDEAL_90"],r["IDEAL_ALT"],
              r["EXCESSO_90"],r["EXCESSO_ALT"],r["CMM_REDE"],r["VALOR UNIT"],r["VALOR_EXCESSO_90"]]
        for j,(v,f) in enumerate(zip(vals,fmts4),1): cel(ws4,i,j,v,fmt=f,bg=bg,halign="right" if f else "left")
    W(ws4,[10,22,7,16,38,26,14,10,10,10,10,10,10,12,16]); ws4.freeze_panes="A3"

    # ── Aba 5: Vencimento ─────────────────────────────────────────────
    if df_lotes is not None:
        ws5 = wb.create_sheet("5. Risco de Vencimento")
        hdr(ws5,1,1,"ANÁLISE DE VENCIMENTO POR LOTE  |  risco = vence antes de conseguir consumir",bg="6E2D2D",size=11)
        ws5.merge_cells("A1:L1")
        cols5=["CD/Filial","Cód. Produto","Descrição","Lote","Validade","Dias até Vencer",
               "Meses até Vencer","Qtd. Lote","CMM Rede","Meses p/ Consumir","Saldo (meses)","Status"]
        fmts5=[None,None,None,None,"DD/MM/YYYY","#,##0","#,##0.1","#,##0.00","#,##0.00","#,##0.1","#,##0.1",None]
        for j,h in enumerate(cols5,1): hdr(ws5,2,j,h,bg="922B21",wrap=True)
        df_r=df_lotes[~df_lotes["STATUS_VENCIMENTO"].str.startswith("OK")].sort_values(["STATUS_VENCIMENTO","MESES_ATE_VENCER"])
        hoje=date.today()
        for i,(_,r) in enumerate(df_r.iterrows(),3):
            st=str(r["STATUS_VENCIMENTO"])
            bg="F1948A" if st=="VENCE ANTES DE CONSUMIR" else ("FAD7A0" if st=="MARGEM APERTADA" else "FDEDEC")
            dias=(r["VALIDADE_DT"]-hoje).days if r["VALIDADE_DT"] else None
            mc=r["MESES_PARA_CONSUMIR"] if r["MESES_PARA_CONSUMIR"]!=np.inf else None
            saldo=r["SALDO_MESES"] if r["MESES_PARA_CONSUMIR"]!=np.inf else None
            lote=r.get("LOTE", r.get("Lote fabricante",""))
            filial=r.get("FILIAL", r.get("Filial",""))
            vals=[filial,r["COD_STR"],r.get("DESCRICAO",""),lote,r["VALIDADE_DT"],dias,
                  r["MESES_ATE_VENCER"],r["QUANTIDADE"],r["CMM_REDE"],mc,saldo,st]
            for j,(v,f) in enumerate(zip(vals,fmts5),1):
                c=cel(ws5,i,j,v,fmt=f,bg=bg,halign="right" if f else "left")
                if j==12 and st=="VENCE ANTES DE CONSUMIR": c.font=Font(bold=True,color="000000",size=9,name="Arial")
        W(ws5,[10,16,38,18,13,12,14,12,10,14,12,24]); ws5.freeze_panes="A3"

    # ── Aba 6: Por Hospital ───────────────────────────────────────────
    ws6=wb.create_sheet("6. Por Hospital")
    hdr(ws6,1,1,"RESUMO POR HOSPITAL",bg="1F3864",size=11); ws6.merge_cells("A1:J1")
    cols6=["Regional","Hospital","Sigla","Valor Físico","Excesso 90d","% Excesso 90d",
           f"Excesso {politica_dias}d",f"% Excesso {politica_dias}d","SKUs Total","SKUs Sem Consumo"]
    fmts6=[None,None,None,"R$ #,##0.00","R$ #,##0.00","0.0%","R$ #,##0.00","0.0%","#,##0","#,##0"]
    for j,h in enumerate(cols6,1): hdr(ws6,2,j,h,bg="2E75B6",wrap=True)
    resumo=df.groupby(["HOSPITAL AJUSTADO","SIGLA","REGIONAL"]).agg(
        VF=("VALOR TOTAL ","sum"),E90=("VALOR_EXCESSO_90","sum"),EALT=("VALOR_EXCESSO_ALT","sum"),
        SKUS=("COD VALORES","count"),SEM=("CMM",lambda x:(x==0).sum())
    ).reset_index().sort_values("E90",ascending=False)
    for i,(_,r) in enumerate(resumo.iterrows(),3):
        bg="EBF3FB" if i%2==0 else "FFFFFF"
        p90=r["E90"]/r["VF"] if r["VF"] else 0
        palt=r["EALT"]/r["VF"] if r["VF"] else 0
        vals=[r["REGIONAL"],r["HOSPITAL AJUSTADO"],r["SIGLA"],r["VF"],r["E90"],p90,r["EALT"],palt,r["SKUS"],r["SEM"]]
        for j,(v,f) in enumerate(zip(vals,fmts6),1):
            c=cel(ws6,i,j,v,fmt=f,bg=bg,halign="right" if f else "left")
            if j in (6,8) and v>0.6: c.fill=PatternFill("solid",start_color="FADBD8")
    W(ws6,[10,26,7,18,16,13,16,13,10,14]); ws6.freeze_panes="A3"

    # ── Aba 7: Base completa ──────────────────────────────────────────
    ws7=wb.create_sheet("7. Base Completa")
    cols_b=["HOSPITAL AJUSTADO","SIGLA","REGIONAL","COD VALORES","DESCRICAO","LINHA",
            "CLASSIFICACAO","CLASSIFICACAO_ALT","ORIGEM_GRADE",
            "FÍSICO","IDEAL_90","IDEAL_ALT","EXCESSO_90","EXCESSO_ALT",
            "VALOR UNIT","VALOR TOTAL ","VALOR_EXCESSO_90","VALOR_EXCESSO_ALT",
            "CMM","CMM_REDE","GRADE","STATUS GRADE","STATUS"]
    hdrs_b=["Hospital","Sigla","Regional","Cód. Produto","Descrição","Linha",
            "Classif. 90d",f"Classif. {politica_dias}d","Ideal da Grade?",
            "Físico","Ideal 90d",f"Ideal {politica_dias}d","Excesso 90d",f"Excesso {politica_dias}d",
            "Valor Unit.","Valor Total","Valor Exc. 90d",f"Valor Exc. {politica_dias}d",
            "CMM Hosp.","CMM Rede","Grade","Status Grade","Status Consumo"]
    fmts_b=[None,None,None,None,None,None,None,None,None,
            "#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00",
            "R$ #,##0.00","R$ #,##0.00","R$ #,##0.00","R$ #,##0.00",
            "#,##0.00","#,##0.00","#,##0.00",None,None]
    for j,h in enumerate(hdrs_b,1): hdr(ws7,1,j,h,bg="2E75B6",wrap=True)
    df_b=df[cols_b].copy()
    df_b["ORIGEM_GRADE"]=df_b["ORIGEM_GRADE"].map({True:"SIM",False:"NÃO"})
    df_b=df_b.sort_values(["CLASSIFICACAO","VALOR_EXCESSO_90"],ascending=[True,False])
    for i,(_,r) in enumerate(df_b.iterrows(),2):
        bg=CORES.get(r["CLASSIFICACAO"],"FFFFFF")
        for j,(col,f) in enumerate(zip(cols_b,fmts_b),1):
            cel(ws7,i,j,r[col],fmt=f,bg=bg,halign="right" if f else "left")
    W(ws7,[22,7,10,16,38,16,26,26,14,10,10,10,10,10,12,14,16,16,10,10,8,16,18])
    ws7.freeze_panes="A2"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ── Interface ─────────────────────────────────────────────────────────
st.title("📦 Equalização de Estoques")
st.caption("Análise de redistribuição e retorno de estoques consignados")

with st.sidebar:
    st.header("⚙️ Configurações")
    politica_dias = st.selectbox("Política alternativa de cobertura", [120,150,180], index=1,
                                  help="90 dias é sempre calculado. Esta é a política adicional.")
    st.divider()
    st.markdown("**Legenda de cores**")
    for clf,cor,desc in [
        ("Transferível","#D5F5E3","Excesso real — consumo justifica"),
        ("Transferível - Grade","#A9DFBF","Excesso de grade — sem consumo real"),
        ("Retorno CD","#FADBD8","Sem consumo na rede"),
        ("Retorno CD - Grade","#F1948A","Sem consumo nem grade"),
        ("Adequado/Deficitário","#EBF3FB","Estoque dentro do ideal"),
    ]:
        st.markdown(f'<span style="background:{cor};padding:2px 8px;border-radius:4px;font-size:0.8rem">{clf}</span> {desc}',unsafe_allow_html=True)

st.subheader("1. Base de estoques (.xlsb)")
file_base = st.file_uploader("Upload do arquivo de estudo",type=["xlsb"],
                              help="Arquivo com aba BASE_ESTUDOS")

if file_base:
    file_bytes = file_base.read()
    with st.spinner("Lendo arquivo..."):
        sheets = get_sheet_names(file_bytes)

    if "BASE_ESTUDOS" not in sheets:
        st.error(f"Aba BASE_ESTUDOS não encontrada. Abas: {sheets}")
        st.stop()

    with st.spinner("Carregando BASE_ESTUDOS..."):
        df_raw = read_sheet(file_bytes, "BASE_ESTUDOS")

    linhas_disp = sorted(df_raw["LINHA"].dropna().unique().tolist())
    c1,c2 = st.columns([2,1])
    with c1:
        linhas_sel = st.multiselect("Linhas de produto",linhas_disp,
                                     default=["HEMODINAMICA"] if "HEMODINAMICA" in linhas_disp else linhas_disp[:1])
    with c2:
        regs_disp = sorted(df_raw["REGIONAL"].dropna().unique().tolist())
        regs_sel  = st.multiselect("Regionais (opcional)",regs_disp)

    if not linhas_sel:
        st.warning("Selecione ao menos uma linha de produto.")
        st.stop()

    df_filtrado = df_raw if not regs_sel else df_raw[df_raw["REGIONAL"].isin(regs_sel)]

    st.subheader("2. Base de lotes e validades (opcional)")
    file_lotes = st.file_uploader("Upload de lotes (.xlsx ou .csv)",type=["xlsx","csv"],
                                   help="Colunas: COD_PRODUTO, HOSPITAL, LOTE, VALIDADE, QUANTIDADE")
    df_lotes = None
    usar_cmv2 = False
    if file_lotes is None and "CMV_2" in sheets:
        usar_cmv2 = st.checkbox("Usar aba CMV_2 do próprio arquivo (lotes do consignado)",value=True)

    if st.button("▶ Rodar análise",type="primary",use_container_width=True):
        with st.spinner("Processando..."):
            df = processar_base_estudos(df_filtrado, linhas_sel, politica_dias)

            if file_lotes:
                raw_lotes = pd.read_csv(file_lotes) if file_lotes.name.endswith(".csv") else pd.read_excel(file_lotes)
                df_lotes = processar_lotes(raw_lotes, df)
            elif usar_cmv2:
                with st.spinner("Carregando CMV_2..."):
                    df_cmv2 = read_sheet(file_bytes,"CMV_2")
                df_lotes = processar_lotes(df_cmv2, df)

        st.divider()
        st.subheader("📊 Resultados")
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Valor físico",    f"R$ {df['VALOR TOTAL '].sum():,.0f}")
        c2.metric("Excesso 90d",     f"R$ {df['VALOR_EXCESSO_90'].sum():,.0f}")
        c3.metric("Transferível 90d",f"R$ {df[df['CLASSIFICACAO'].str.startswith('TRANSFERIVEL')]['VALOR_EXCESSO_90'].sum():,.0f}")
        c4.metric("Retorno ao CD",   f"R$ {df[df['CLASSIFICACAO'].str.startswith('RETORNO_CD')]['VALOR_EXCESSO_90'].sum():,.0f}")

        if df_lotes is not None:
            c5,c6,_,_ = st.columns(4)
            c5.metric("Lotes em risco",df_lotes[df_lotes["STATUS_VENCIMENTO"]=="VENCE ANTES DE CONSUMIR"].shape[0])
            c6.metric("SKUs em risco", df_lotes[df_lotes["STATUS_VENCIMENTO"]=="VENCE ANTES DE CONSUMIR"]["COD_STR"].nunique())

        st.dataframe(df["CLASSIFICACAO"].value_counts().reset_index().rename(
            columns={"CLASSIFICACAO":"Classificação 90d","count":"Qtd. itens"}),
            use_container_width=True, hide_index=True)

        st.divider()
        buf = make_excel(df, df_lotes, politica_dias)
        nome = f"Equalizacao_{'-'.join(linhas_sel)}_{date.today().strftime('%Y%m%d')}.xlsx"
        st.download_button("⬇️ Baixar Excel completo",data=buf,file_name=nome,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True,type="primary")
