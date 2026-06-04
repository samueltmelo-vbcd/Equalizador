import streamlit as st
import pandas as pd
import numpy as np
import xlsxwriter
import io
import tempfile
import os
from datetime import date, timedelta, datetime
from python_calamine import CalamineWorkbook

st.set_page_config(page_title="Equalização de Estoques", page_icon="📦", layout="wide")
st.markdown("""
<style>
[data-testid="stMetricValue"]{font-size:1.4rem}
.block-container{padding-top:2rem}
</style>""", unsafe_allow_html=True)

# ── Leitura de arquivo ────────────────────────────────────────────────
def save_tmp(b, suffix=".xlsb"):
    f = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    f.write(b); f.close()
    return f.name

@st.cache_data(show_spinner=False)
def get_sheets(b):
    p = save_tmp(b)
    try:    return CalamineWorkbook.from_path(p).sheet_names
    finally: os.unlink(p)

@st.cache_data(show_spinner=False)
def read_sheet(b, sheet):
    p = save_tmp(b)
    try:    return pd.read_excel(p, sheet_name=sheet, engine="calamine", dtype=str)
    finally: os.unlink(p)

def to_float(v):
    try:    return float(str(v).replace(",", ".").strip())
    except: return 0.0

def parse_date(v):
    if v is None or (isinstance(v, str) and not v.strip()): return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date):     return v
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%Y"):
        try:    return datetime.strptime(str(v).strip()[:len(fmt)], fmt).date()
        except: pass
    return None

# ── Análise ───────────────────────────────────────────────────────────
def processar(df_raw, linhas, pol_dias):
    df = df_raw[df_raw["LINHA"].isin(linhas) &
                (df_raw["CONSIDERAR NO ESTUDO?"].str.strip() == "SIM")].copy()
    for c in ["FÍSICO","CMM","GRADE","VALOR UNIT","VALOR TOTAL "]:
        if c in df.columns: df[c] = df[c].apply(to_float)

    def ideal(row, d):
        m = d / 30
        if row["GRADE"] > 0:   return row["GRADE"], True
        if row["CMM"]   > 0:   return row["CMM"] * m, False
        return 0.0, False

    r90  = df.apply(lambda r: ideal(r, 90),       axis=1)
    ralt = df.apply(lambda r: ideal(r, pol_dias),  axis=1)
    df["IDEAL_90"]       = r90.apply(lambda x: x[0])
    df["IDEAL_ALT"]      = ralt.apply(lambda x: x[0])
    df["ORIGEM_GRADE"]   = r90.apply(lambda x: x[1])
    df["EXCESSO_90"]     = df["FÍSICO"] - df["IDEAL_90"]
    df["EXCESSO_ALT"]    = df["FÍSICO"] - df["IDEAL_ALT"]
    df["VEX_90"]         = df["EXCESSO_90"].clip(lower=0)  * df["VALOR UNIT"]
    df["VEX_ALT"]        = df["EXCESSO_ALT"].clip(lower=0) * df["VALOR UNIT"]
    cmm_r = df.groupby("COD VALORES")["CMM"].sum().rename("CMM_REDE")
    df = df.join(cmm_r, on="COD VALORES")

    def clf(row, exc):
        if row["FÍSICO"] == 0:  return "SEM ESTOQUE"
        sfx = " - GRADE" if row["ORIGEM_GRADE"] else ""
        if row[exc] <= 0:       return f"ADEQUADO/DEFICITARIO{sfx}"
        if row["CMM_REDE"] > 0: return f"TRANSFERIVEL{sfx}"
        return f"RETORNO_CD{sfx}"

    df["CLF"]     = df.apply(lambda r: clf(r, "EXCESSO_90"),  axis=1)
    df["CLF_ALT"] = df.apply(lambda r: clf(r, "EXCESSO_ALT"), axis=1)
    return df

def processar_lotes(df_raw, df_base):
    df = df_raw.copy()
    for cands, dest in [
        (["COD_PRODUTO","COD VALORES","Codigo do item","COD ARRUMAR"], "COD_STR"),
        (["HOSPITAL","Filial","SIGLA"],                                 "FILIAL"),
        (["LOTE","Lote fabricante","Lote interno"],                     "LOTE"),
        (["VALIDADE","Validade","DATA_VALIDADE"],                       "VALIDADE_RAW"),
        (["QUANTIDADE","Quantidade","QTDE","QTD"],                      "QUANTIDADE"),
    ]:
        for c in cands:
            if c in df.columns: df = df.rename(columns={c: dest}); break
    df["COD_STR"]   = df["COD_STR"].astype(str).str.strip()
    df["QUANTIDADE"] = df["QUANTIDADE"].apply(to_float)
    df["VALIDADE_DT"] = df["VALIDADE_RAW"].apply(parse_date)
    skus = set(df_base["COD VALORES"].astype(str).str.strip())
    df = df[df["COD_STR"].isin(skus) & df["VALIDADE_DT"].notna()].copy()
    cmm_m  = df_base.groupby(df_base["COD VALORES"].astype(str).str.strip())["CMM_REDE"].first()
    desc_m = df_base.groupby(df_base["COD VALORES"].astype(str).str.strip())["DESCRICAO"].first()
    df["CMM_REDE"] = df["COD_STR"].map(cmm_m).fillna(0).apply(to_float)
    df["DESCRICAO"] = df["COD_STR"].map(desc_m)
    hoje = date.today()
    df["MAV"] = df["VALIDADE_DT"].apply(lambda v: max(0,(v-hoje).days/30))
    df["MPC"] = df.apply(lambda r: r["QUANTIDADE"]/r["CMM_REDE"] if r["CMM_REDE"]>0 else np.inf, axis=1)
    df["SALDO"] = df["MAV"] - df["MPC"]
    def sv(r):
        if r["CMM_REDE"]==0:  return f"SEM CONSUMO — VENCE EM {r['MAV']:.0f}m"
        if r["SALDO"]<0:      return "VENCE ANTES DE CONSUMIR"
        if r["SALDO"]<=1:     return "MARGEM APERTADA"
        return "OK"
    df["STATUS_VEN"] = df.apply(sv, axis=1)
    return df

# ── Excel com xlsxwriter ──────────────────────────────────────────────
CORES = {
    "TRANSFERIVEL":                 "#D5F5E3",
    "TRANSFERIVEL - GRADE":         "#A9DFBF",
    "RETORNO_CD":                   "#FADBD8",
    "RETORNO_CD - GRADE":           "#F1948A",
    "ADEQUADO/DEFICITARIO":         "#EBF3FB",
    "ADEQUADO/DEFICITARIO - GRADE": "#D6EAF8",
    "SEM ESTOQUE":                  "#F9F9F9",
}

def make_excel(df, df_lotes, pol_dias):
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True})

    # formatos base
    def F(**kw):
        base = {"font_name":"Arial","font_size":9,"border":1,"valign":"vcenter"}
        base.update(kw); return wb.add_format(base)

    fH   = F(bold=True, bg_color="#1F3864", font_color="white", font_size=10, align="center", text_wrap=True)
    fH2  = F(bold=True, bg_color="#2E75B6", font_color="white", font_size=9,  align="center", text_wrap=True)
    fH3  = F(bold=True, bg_color="#C0392B", font_color="white", font_size=9,  align="center", text_wrap=True)
    fH4  = F(bold=True, bg_color="#922B21", font_color="white", font_size=9,  align="center", text_wrap=True)
    fH5  = F(bold=True, bg_color="#155A8A", font_color="white", font_size=9,  align="center", text_wrap=True)
    fTxt = F(align="left")
    fNum = F(align="right", num_format="#,##0.00")
    fBRL = F(align="right", num_format='R$ #,##0.00')
    fPct = F(align="right", num_format="0.0%")
    fDat = F(align="center",num_format="dd/mm/yyyy")
    fBld = F(bold=True, align="left")

    def bg_txt(cor): return F(align="left",  bg_color=cor)
    def bg_num(cor): return F(align="right", bg_color=cor, num_format="#,##0.00")
    def bg_brl(cor): return F(align="right", bg_color=cor, num_format="R$ #,##0.00")
    def bg_pct(cor): return F(align="right", bg_color=cor, num_format="0.0%")
    def bg_dat(cor): return F(align="center",bg_color=cor, num_format="dd/mm/yyyy")

    def write_row(ws, row, vals, fmts):
        for col,(v,f) in enumerate(zip(vals,fmts)):
            if isinstance(v, (datetime,)): ws.write_datetime(row, col, v, f)
            elif isinstance(v, date):      ws.write_datetime(row, col, datetime(v.year,v.month,v.day), f)
            elif v is None or (isinstance(v,float) and np.isnan(v)): ws.write_blank(row, col, None, f)
            else: ws.write(row, col, v, f)

    hoje = date.today()

    # ── ABA 1: Resumo ─────────────────────────────────────────────────
    ws1 = wb.add_worksheet("1. Resumo Executivo")
    ws1.set_row(0, 28); ws1.set_row(1, 6)
    ws1.merge_range(0,0,0,8, f"EQUALIZAÇÃO DE ESTOQUES  |  {hoje.strftime('%d/%m/%Y')}  |  Política alt.: {pol_dias}d", fH)

    def kpi_bloco(col, dias_lbl, exc_col, clf_col):
        ws1.write(2, col,   f"Indicador — {dias_lbl}", fH2)
        ws1.write(2, col+1, "Valor",                   fH2)
        kpis = [
            ("Valor físico total",              df["VALOR TOTAL "].sum()),
            (f"Excesso total ({dias_lbl})",     df[exc_col].sum()),
            ("→ Transferível — consumo real",   df[df[clf_col]=="TRANSFERIVEL"][exc_col].sum()),
            ("→ Transferível — grade",          df[df[clf_col]=="TRANSFERIVEL - GRADE"][exc_col].sum()),
            ("→ Retorno CD — sem consumo",      df[df[clf_col]=="RETORNO_CD"][exc_col].sum()),
            ("→ Retorno CD — grade s/ consumo", df[df[clf_col]=="RETORNO_CD - GRADE"][exc_col].sum()),
            ("Hospitais analisados",            df["HOSPITAL AJUSTADO"].nunique()),
            ("SKUs únicos",                     df["COD VALORES"].nunique()),
        ]
        if df_lotes is not None:
            kpis += [
                ("Lotes com risco de vencer", int(df_lotes[df_lotes["STATUS_VEN"]=="VENCE ANTES DE CONSUMIR"].shape[0])),
                ("SKUs com risco de vencer",  int(df_lotes[df_lotes["STATUS_VEN"]=="VENCE ANTES DE CONSUMIR"]["COD_STR"].nunique())),
            ]
        ev="#EBF3FB"; od="#FFFFFF"
        for i,(lbl,val) in enumerate(kpis):
            bg = ev if i%2==0 else od
            bold = "→" not in lbl
            ws1.write(3+i, col,   lbl, F(align="left",  bg_color=bg, bold=bold))
            ws1.write(3+i, col+1, val, F(align="right", bg_color=bg, bold=bold,
                                          num_format="R$ #,##0.00" if isinstance(val,float) else "#,##0"))

    kpi_bloco(0, "90d",       "VEX_90",  "CLF")
    kpi_bloco(3, f"{pol_dias}d", "VEX_ALT", "CLF_ALT")

    ws1.write(2,6,"Regional",fH2); ws1.write(2,7,"Excesso 90d",fH2); ws1.write(2,8,f"Excesso {pol_dias}d",fH2)
    reg = df.groupby("REGIONAL").agg(E90=("VEX_90","sum"),EALT=("VEX_ALT","sum")).reset_index().sort_values("E90",ascending=False)
    for i,(_,r) in enumerate(reg.iterrows()):
        bg="#EBF3FB" if i%2==0 else "#FFFFFF"
        ws1.write(3+i,6,r["REGIONAL"],bg_txt(bg))
        ws1.write(3+i,7,r["E90"], bg_brl(bg))
        ws1.write(3+i,8,r["EALT"],bg_brl(bg))

    leg_row = 15
    ws1.merge_range(leg_row,0,leg_row,8,"Legenda de classificações",fH2)
    for j,(clf,cor,desc) in enumerate([
        ("TRANSFERIVEL","#D5F5E3","Excesso real — consumo justifica; redirecionar para hospital com déficit"),
        ("TRANSFERIVEL - GRADE","#A9DFBF","Excesso de grade — acima do contrato, sem consumo real"),
        ("RETORNO_CD","#FADBD8","Sem consumo na rede — retornar ao CD para tratativa com fornecedor"),
        ("RETORNO_CD - GRADE","#F1948A","Sem consumo nem grade — retornar com prioridade máxima"),
        ("ADEQUADO/DEFICITARIO","#EBF3FB","Estoque dentro ou abaixo do ideal"),
        ("ADEQUADO/DEFICITARIO - GRADE","#D6EAF8","Adequado conforme grade contratual"),
    ]):
        ws1.write(leg_row+1+j, 0, clf,  F(bold=True, bg_color=cor, align="left"))
        ws1.merge_range(leg_row+1+j,1,leg_row+1+j,8, desc, F(bg_color="#FFFFFF", align="left"))

    for i,w in enumerate([36,18,3,36,18,3,12,18,18]): ws1.set_column(i,i,w)
    ws1.freeze_panes(3,0)

    # ── Gerador de abas de transferência ─────────────────────────────
    def aba_transf(name, clf_col, exc_col, exc_lbl, hdr_fmt):
        ws = wb.add_worksheet(name)
        ws.merge_range(0,0,0,13, name.split(". ",1)[1], fH)
        cols=["Regional","Hospital","Sigla","Cód. Produto","Descrição","Classificação","Origem Ideal",
              "Físico","Ideal 90d",f"Ideal {pol_dias}d","Excesso 90d",f"Excesso {pol_dias}d","Valor Unit.",exc_lbl]
        for j,h in enumerate(cols): ws.write(1,j,h,hdr_fmt)
        data=df[df[clf_col].str.startswith("TRANSFERIVEL")].sort_values(exc_col,ascending=False)
        for i,(_,r) in enumerate(data.iterrows()):
            bg=CORES.get(r[clf_col],"#FFFFFF")
            orig="Grade" if r["ORIGEM_GRADE"] else "CMM (consumo)"
            vals=[r["REGIONAL"],r["HOSPITAL AJUSTADO"],r["SIGLA"],str(r["COD VALORES"]),r["DESCRICAO"],
                  r[clf_col],orig,r["FÍSICO"],r["IDEAL_90"],r["IDEAL_ALT"],
                  r["EXCESSO_90"],r["EXCESSO_ALT"],r["VALOR UNIT"],r[exc_col]]
            fmts=[bg_txt(bg),bg_txt(bg),bg_txt(bg),bg_txt(bg),bg_txt(bg),bg_txt(bg),bg_txt(bg),
                  bg_num(bg),bg_num(bg),bg_num(bg),bg_num(bg),bg_num(bg),bg_brl(bg),bg_brl(bg)]
            write_row(ws, 2+i, vals, fmts)
        for i,w in enumerate([10,22,7,16,38,26,14,10,10,10,10,10,12,16]): ws.set_column(i,i,w)
        ws.freeze_panes(2,0)

    aba_transf("2. Transferências 90d",   "CLF",     "VEX_90",  "Valor Excesso 90d",    fH2)
    aba_transf(f"3. Transferências {pol_dias}d","CLF_ALT","VEX_ALT",f"Valor Excesso {pol_dias}d",fH5)

    # ── Aba 4: Retorno ao CD ──────────────────────────────────────────
    ws4 = wb.add_worksheet("4. Retorno ao CD")
    ws4.merge_range(0,0,0,14,"RETORNO AO CD  |  igual nas duas políticas — CMM zero na rede",fH)
    cols4=["Regional","Hospital","Sigla","Cód. Produto","Descrição","Classificação","Origem Ideal",
           "Físico","Ideal 90d",f"Ideal {pol_dias}d","Excesso 90d",f"Excesso {pol_dias}d","CMM Rede","Valor Unit.","Valor a Retornar"]
    for j,h in enumerate(cols4): ws4.write(1,j,h,fH3)
    retorno=df[df["CLF"].str.startswith("RETORNO_CD")].sort_values("VEX_90",ascending=False)
    for i,(_,r) in enumerate(retorno.iterrows()):
        bg=CORES.get(r["CLF"],"#FFFFFF")
        orig="Grade" if r["ORIGEM_GRADE"] else "CMM (consumo)"
        vals=[r["REGIONAL"],r["HOSPITAL AJUSTADO"],r["SIGLA"],str(r["COD VALORES"]),r["DESCRICAO"],
              r["CLF"],orig,r["FÍSICO"],r["IDEAL_90"],r["IDEAL_ALT"],
              r["EXCESSO_90"],r["EXCESSO_ALT"],r["CMM_REDE"],r["VALOR UNIT"],r["VEX_90"]]
        fmts=[bg_txt(bg),bg_txt(bg),bg_txt(bg),bg_txt(bg),bg_txt(bg),bg_txt(bg),bg_txt(bg),
              bg_num(bg),bg_num(bg),bg_num(bg),bg_num(bg),bg_num(bg),bg_num(bg),bg_brl(bg),bg_brl(bg)]
        write_row(ws4, 2+i, vals, fmts)
    for i,w in enumerate([10,22,7,16,38,26,14,10,10,10,10,10,10,12,16]): ws4.set_column(i,i,w)
    ws4.freeze_panes(2,0)

    # ── Aba 5: Vencimento ─────────────────────────────────────────────
    if df_lotes is not None:
        ws5 = wb.add_worksheet("5. Risco de Vencimento")
        ws5.merge_range(0,0,0,11,"ANÁLISE DE VENCIMENTO POR LOTE  |  risco = vence antes de consumir",fH)
        cols5=["CD/Filial","Cód. Produto","Descrição","Lote","Validade","Dias até Vencer",
               "Meses até Vencer","Qtd. Lote","CMM Rede","Meses p/ Consumir","Saldo (meses)","Status"]
        for j,h in enumerate(cols5): ws5.write(1,j,h,fH4)
        df_r=df_lotes[~df_lotes["STATUS_VEN"].str.startswith("OK")].sort_values(["STATUS_VEN","MAV"])
        for i,(_,r) in enumerate(df_r.iterrows()):
            st=str(r["STATUS_VEN"])
            bg="#F1948A" if st=="VENCE ANTES DE CONSUMIR" else ("#FAD7A0" if st=="MARGEM APERTADA" else "#FDEDEC")
            dias=(r["VALIDADE_DT"]-hoje).days if r["VALIDADE_DT"] else None
            mc=r["MPC"]   if r["MPC"]!=np.inf else None
            sd=r["SALDO"] if r["MPC"]!=np.inf else None
            lote  = r.get("LOTE",   r.get("Lote fabricante",""))
            filial= r.get("FILIAL", r.get("Filial",""))
            vdt   = datetime(r["VALIDADE_DT"].year,r["VALIDADE_DT"].month,r["VALIDADE_DT"].day) if r["VALIDADE_DT"] else None
            vals=[filial,r["COD_STR"],r.get("DESCRICAO",""),lote,vdt,dias,
                  r["MAV"],r["QUANTIDADE"],r["CMM_REDE"],mc,sd,st]
            fmts=[bg_txt(bg),bg_txt(bg),bg_txt(bg),bg_txt(bg),bg_dat(bg),
                  F(align="right",bg_color=bg,num_format="#,##0"),
                  bg_num(bg),bg_num(bg),bg_num(bg),bg_num(bg),bg_num(bg),
                  F(bold=(st=="VENCE ANTES DE CONSUMIR"),align="left",bg_color=bg)]
            write_row(ws5, 2+i, vals, fmts)
        for i,w in enumerate([10,16,38,18,13,12,14,12,10,14,12,24]): ws5.set_column(i,i,w)
        ws5.freeze_panes(2,0)

    # ── Aba 6: Por Hospital ───────────────────────────────────────────
    ws6=wb.add_worksheet("6. Por Hospital")
    ws6.merge_range(0,0,0,9,"RESUMO POR HOSPITAL",fH)
    cols6=["Regional","Hospital","Sigla","Valor Físico","Excesso 90d","% Excesso 90d",
           f"Excesso {pol_dias}d",f"% Excesso {pol_dias}d","SKUs Total","SKUs Sem Consumo"]
    for j,h in enumerate(cols6): ws6.write(1,j,h,fH2)
    resumo=df.groupby(["HOSPITAL AJUSTADO","SIGLA","REGIONAL"]).agg(
        VF=("VALOR TOTAL ","sum"),E90=("VEX_90","sum"),EALT=("VEX_ALT","sum"),
        SKUS=("COD VALORES","count"),SEM=("CMM",lambda x:(x==0).sum())
    ).reset_index().sort_values("E90",ascending=False)
    for i,(_,r) in enumerate(resumo.iterrows()):
        bg="#EBF3FB" if i%2==0 else "#FFFFFF"
        p90 =r["E90"]/r["VF"]  if r["VF"] else 0
        palt=r["EALT"]/r["VF"] if r["VF"] else 0
        vals=[r["REGIONAL"],r["HOSPITAL AJUSTADO"],r["SIGLA"],r["VF"],r["E90"],p90,r["EALT"],palt,int(r["SKUS"]),int(r["SEM"])]
        fmts=[bg_txt(bg),bg_txt(bg),bg_txt(bg),bg_brl(bg),bg_brl(bg),
              bg_pct("#FADBD8" if p90>0.6 else bg),bg_brl(bg),
              bg_pct("#FADBD8" if palt>0.6 else bg),
              F(align="right",bg_color=bg,num_format="#,##0"),F(align="right",bg_color=bg,num_format="#,##0")]
        write_row(ws6, 2+i, vals, fmts)
    for i,w in enumerate([10,26,7,18,16,13,16,13,10,14]): ws6.set_column(i,i,w)
    ws6.freeze_panes(2,0)

    # ── Aba 7: Base completa ──────────────────────────────────────────
    ws7=wb.add_worksheet("7. Base Completa")
    cols_b=["HOSPITAL AJUSTADO","SIGLA","REGIONAL","COD VALORES","DESCRICAO","LINHA",
            "CLF","CLF_ALT","ORIGEM_GRADE","FÍSICO","IDEAL_90","IDEAL_ALT",
            "EXCESSO_90","EXCESSO_ALT","VALOR UNIT","VALOR TOTAL ","VEX_90","VEX_ALT",
            "CMM","CMM_REDE","GRADE","STATUS GRADE","STATUS"]
    hdrs_b=["Hospital","Sigla","Regional","Cód. Produto","Descrição","Linha",
            "Classif. 90d",f"Classif. {pol_dias}d","Ideal da Grade?",
            "Físico","Ideal 90d",f"Ideal {pol_dias}d","Excesso 90d",f"Excesso {pol_dias}d",
            "Valor Unit.","Valor Total","Valor Exc. 90d",f"Valor Exc. {pol_dias}d",
            "CMM Hosp.","CMM Rede","Grade","Status Grade","Status Consumo"]
    for j,h in enumerate(hdrs_b): ws7.write(0,j,h,fH2)
    df_b=df[cols_b].copy()
    df_b["ORIGEM_GRADE"]=df_b["ORIGEM_GRADE"].map({True:"SIM",False:"NÃO"})
    df_b=df_b.sort_values(["CLF","VEX_90"],ascending=[True,False])
    num_set={"FÍSICO","IDEAL_90","IDEAL_ALT","EXCESSO_90","EXCESSO_ALT","CMM","CMM_REDE","GRADE"}
    brl_set={"VALOR UNIT","VALOR TOTAL ","VEX_90","VEX_ALT"}
    for i,(_,r) in enumerate(df_b.iterrows()):
        bg=CORES.get(r["CLF"],"#FFFFFF")
        for j,col in enumerate(cols_b):
            v=r[col]
            if col in brl_set:   ws7.write(1+i,j,v,bg_brl(bg))
            elif col in num_set: ws7.write(1+i,j,v,bg_num(bg))
            else:                ws7.write(1+i,j,v,bg_txt(bg))
    for i,w in enumerate([22,7,10,16,38,16,26,26,12,10,10,10,10,10,12,14,14,14,10,10,8,16,18]):
        ws7.set_column(i,i,w)
    ws7.freeze_panes(1,0)

    wb.close(); buf.seek(0)
    return buf

# ── Interface ─────────────────────────────────────────────────────────
st.title("📦 Equalização de Estoques")
st.caption("Análise de redistribuição e retorno de estoques consignados")

with st.sidebar:
    st.header("⚙️ Configurações")
    pol_dias = st.selectbox("Política alternativa de cobertura", [120,150,180], index=1)
    st.divider()
    st.markdown("**Legenda**")
    for clf,cor,desc in [
        ("Transferível","#D5F5E3","Excesso real"),
        ("Transferível - Grade","#A9DFBF","Excesso de grade"),
        ("Retorno CD","#FADBD8","Sem consumo na rede"),
        ("Retorno CD - Grade","#F1948A","Sem consumo nem grade"),
        ("Adequado","#EBF3FB","Dentro do ideal"),
    ]:
        st.markdown(f'<span style="background:{cor};padding:2px 8px;border-radius:4px;font-size:.8rem">{clf}</span> {desc}',unsafe_allow_html=True)

st.subheader("1. Base de estoques (.xlsb)")
file_base = st.file_uploader("Upload do arquivo de estudo", type=["xlsb"])

if file_base:
    fb = file_base.read()
    with st.spinner("Lendo arquivo..."):
        sheets = get_sheets(fb)

    if "BASE_ESTUDOS" not in sheets:
        st.error(f"Aba BASE_ESTUDOS não encontrada. Abas: {sheets}"); st.stop()

    with st.spinner("Carregando dados..."):
        df_raw = read_sheet(fb, "BASE_ESTUDOS")

    c1,c2 = st.columns([2,1])
    with c1:
        linhas_disp = sorted(df_raw["LINHA"].dropna().unique())
        linhas_sel  = st.multiselect("Linhas de produto", linhas_disp,
                      default=["HEMODINAMICA"] if "HEMODINAMICA" in linhas_disp else linhas_disp[:1])
    with c2:
        regs_disp = sorted(df_raw["REGIONAL"].dropna().unique())
        regs_sel  = st.multiselect("Regionais (opcional)", regs_disp)

    if not linhas_sel:
        st.warning("Selecione ao menos uma linha."); st.stop()

    df_filtrado = df_raw if not regs_sel else df_raw[df_raw["REGIONAL"].isin(regs_sel)]

    st.subheader("2. Base de lotes e validades (opcional)")
    file_lotes = st.file_uploader("Upload de lotes (.xlsx ou .csv)", type=["xlsx","csv"])
    df_lotes   = None
    usar_cmv2  = False
    if file_lotes is None and "CMV_2" in sheets:
        usar_cmv2 = st.checkbox("Usar aba CMV_2 do próprio arquivo", value=True)

    if st.button("▶ Rodar análise", type="primary", use_container_width=True):
        with st.spinner("Processando..."):
            df = processar(df_filtrado, linhas_sel, pol_dias)
            if file_lotes:
                rl = pd.read_csv(file_lotes) if file_lotes.name.endswith(".csv") else pd.read_excel(file_lotes)
                df_lotes = processar_lotes(rl, df)
            elif usar_cmv2:
                df_lotes = processar_lotes(read_sheet(fb,"CMV_2"), df)

        st.divider()
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Valor físico",     f"R$ {df['VALOR TOTAL '].sum():,.0f}")
        c2.metric("Excesso 90d",      f"R$ {df['VEX_90'].sum():,.0f}")
        c3.metric("Transferível 90d", f"R$ {df[df['CLF'].str.startswith('TRANSFERIVEL')]['VEX_90'].sum():,.0f}")
        c4.metric("Retorno ao CD",    f"R$ {df[df['CLF'].str.startswith('RETORNO_CD')]['VEX_90'].sum():,.0f}")

        if df_lotes is not None:
            c5,c6,_,_ = st.columns(4)
            c5.metric("Lotes em risco", int(df_lotes[df_lotes["STATUS_VEN"]=="VENCE ANTES DE CONSUMIR"].shape[0]))
            c6.metric("SKUs em risco",  int(df_lotes[df_lotes["STATUS_VEN"]=="VENCE ANTES DE CONSUMIR"]["COD_STR"].nunique()))

        st.dataframe(df["CLF"].value_counts().reset_index().rename(
            columns={"CLF":"Classificação 90d","count":"Qtd. itens"}),
            use_container_width=True, hide_index=True)

        st.divider()
        buf  = make_excel(df, df_lotes, pol_dias)
        nome = f"Equalizacao_{'_'.join(linhas_sel)}_{date.today().strftime('%Y%m%d')}.xlsx"
        st.download_button("⬇️ Baixar Excel completo", data=buf, file_name=nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary")
