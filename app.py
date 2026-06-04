import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import io
from datetime import date, timedelta, datetime

st.set_page_config(page_title="Equalização de Estoques", page_icon="📦", layout="wide")
st.markdown("""
<style>
[data-testid="stMetricValue"]{font-size:1.5rem;font-weight:600}
[data-testid="stMetricLabel"]{font-size:.8rem;color:#666}
.block-container{padding-top:1.5rem;padding-bottom:2rem}
div[data-testid="stTabs"] button{font-size:.9rem}
</style>""", unsafe_allow_html=True)

# ── Leitura ───────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def get_sheets(b):
    wb = openpyxl.load_workbook(io.BytesIO(b), read_only=True, data_only=True)
    names = wb.sheetnames; wb.close(); return names

@st.cache_data(show_spinner=False)
def read_sheet(b, sheet):
    return pd.read_excel(io.BytesIO(b), sheet_name=sheet, engine="openpyxl", dtype=str)

def to_float(v):
    try:    return float(str(v).replace(",",".").strip())
    except: return 0.0

def parse_date(v):
    if v is None or (isinstance(v, str) and not v.strip()): return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date):     return v
    try:
        n = float(str(v).strip())
        return (date(1899,12,30)+timedelta(days=int(n))) if n>0 else None
    except: pass
    for fmt in ("%Y-%m-%d","%d/%m/%Y","%m/%Y"):
        try: return datetime.strptime(str(v).strip()[:10],fmt).date()
        except: pass
    return None

# ── Análise ───────────────────────────────────────────────────────────
def processar(df_raw, linhas, pol_dias):
    df = df_raw[df_raw["LINHA"].isin(linhas) &
                (df_raw["CONSIDERAR NO ESTUDO?"].astype(str).str.strip()=="SIM")].copy()
    for c in ["FÍSICO","CMM","GRADE","VALOR UNIT","VALOR TOTAL "]:
        if c in df.columns: df[c] = df[c].apply(to_float)

    def ideal(row, d):
        m = d/30
        if row["GRADE"]>0: return row["GRADE"], True
        if row["CMM"]>0:   return row["CMM"]*m,  False
        return 0.0, False

    r90  = df.apply(lambda r: ideal(r,90),      axis=1)
    ralt = df.apply(lambda r: ideal(r,pol_dias), axis=1)
    df["IDEAL_90"]     = r90.apply(lambda x: x[0])
    df["IDEAL_ALT"]    = ralt.apply(lambda x: x[0])
    df["ORIGEM_GRADE"] = r90.apply(lambda x: x[1])
    df["EXCESSO_90"]   = df["FÍSICO"] - df["IDEAL_90"]
    df["EXCESSO_ALT"]  = df["FÍSICO"] - df["IDEAL_ALT"]
    df["VEX_90"]       = df["EXCESSO_90"].clip(lower=0) * df["VALOR UNIT"]
    df["VEX_ALT"]      = df["EXCESSO_ALT"].clip(lower=0)* df["VALOR UNIT"]
    cmm_r = df.groupby("COD VALORES")["CMM"].sum().rename("CMM_REDE")
    df = df.join(cmm_r, on="COD VALORES")

    def clf(row, exc):
        if row["FÍSICO"]==0: return "SEM ESTOQUE"
        sfx = " - GRADE" if row["ORIGEM_GRADE"] else ""
        if row[exc]<=0:        return f"ADEQUADO{sfx}"
        if row["CMM_REDE"]>0:  return f"TRANSFERIVEL{sfx}"
        return f"RETORNO_CD{sfx}"

    df["CLF"]     = df.apply(lambda r: clf(r,"EXCESSO_90"),  axis=1)
    df["CLF_ALT"] = df.apply(lambda r: clf(r,"EXCESSO_ALT"), axis=1)
    return df

def processar_lotes(df_raw, df_base):
    df = df_raw.copy()
    for cands,dest in [
        (["COD_PRODUTO","COD VALORES","Codigo do item","COD ARRUMAR"],"COD_STR"),
        (["HOSPITAL","Filial","SIGLA"],"FILIAL"),
        (["LOTE","Lote fabricante","Lote interno"],"LOTE"),
        (["VALIDADE","Validade","DATA_VALIDADE"],"VALIDADE_RAW"),
        (["QUANTIDADE","Quantidade","QTDE","QTD"],"QUANTIDADE"),
    ]:
        for c in cands:
            if c in df.columns: df=df.rename(columns={c:dest}); break
    df["COD_STR"]     = df["COD_STR"].astype(str).str.strip()
    df["QUANTIDADE"]  = df["QUANTIDADE"].apply(to_float)
    df["VALIDADE_DT"] = df["VALIDADE_RAW"].apply(parse_date)
    skus = set(df_base["COD VALORES"].astype(str).str.strip())
    df = df[df["COD_STR"].isin(skus) & df["VALIDADE_DT"].notna()].copy()
    cmm_m  = df_base.groupby(df_base["COD VALORES"].astype(str).str.strip())["CMM_REDE"].first()
    desc_m = df_base.groupby(df_base["COD VALORES"].astype(str).str.strip())["DESCRICAO"].first()
    df["CMM_REDE"] = df["COD_STR"].map(cmm_m).fillna(0).apply(to_float)
    df["DESCRICAO"]= df["COD_STR"].map(desc_m)
    hoje = date.today()
    df["MAV"]   = df["VALIDADE_DT"].apply(lambda v: max(0,(v-hoje).days/30))
    df["MPC"]   = df.apply(lambda r: r["QUANTIDADE"]/r["CMM_REDE"] if r["CMM_REDE"]>0 else np.inf, axis=1)
    df["SALDO"] = df["MAV"] - df["MPC"]
    def sv(r):
        if r["CMM_REDE"]==0: return f"Sem consumo — vence em {r['MAV']:.0f}m"
        if r["SALDO"]<0:     return "Vence antes de consumir"
        if r["SALDO"]<=1:    return "Margem apertada"
        return "OK"
    df["STATUS_VEN"] = df.apply(sv, axis=1)
    return df

# ── Excel download ────────────────────────────────────────────────────
def make_excel(df, df_lotes, pol_dias):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    thin = Side(style="thin", color="CCCCCC")
    brd  = Border(left=thin,right=thin,top=thin,bottom=thin)
    CORES={"TRANSFERIVEL":"D5F5E3","TRANSFERIVEL - GRADE":"A9DFBF",
           "RETORNO_CD":"FADBD8","RETORNO_CD - GRADE":"F1948A",
           "ADEQUADO":"EBF3FB","ADEQUADO - GRADE":"D6EAF8","SEM ESTOQUE":"F9F9F9"}
    def hdr(ws,r,c,v,bg="1F3864",fg="FFFFFF",sz=9,wrap=False):
        x=ws.cell(r,c,v); x.font=Font(bold=True,color=fg,size=sz,name="Arial")
        x.fill=PatternFill("solid",start_color=bg)
        x.alignment=Alignment(horizontal="center",vertical="center",wrap_text=wrap)
        x.border=brd
    def cel(ws,r,c,v,fmt=None,bg=None,bold=False,ha="left"):
        x=ws.cell(r,c,v); x.font=Font(bold=bold,color="000000",size=9,name="Arial")
        x.alignment=Alignment(horizontal=ha,vertical="center")
        if fmt: x.number_format=fmt
        if bg:  x.fill=PatternFill("solid",start_color=bg)
        x.border=brd; return x
    def W(ws,ws_):
        for i,w in enumerate(ws_,1): ws.column_dimensions[get_column_letter(i)].width=w
    wb=Workbook(); hoje=date.today()
    # Resumo
    ws1=wb.active; ws1.title="1. Resumo"
    hdr(ws1,1,1,f"EQUALIZAÇÃO — {hoje.strftime('%d/%m/%Y')} — pol. alt.: {pol_dias}d",sz=11)
    ws1.merge_cells("A1:F1")
    kpis=[("Valor físico total",df["VALOR TOTAL "].sum(),"R$ #,##0.00"),
          ("Excesso 90d",df["VEX_90"].sum(),"R$ #,##0.00"),
          (f"Excesso {pol_dias}d",df["VEX_ALT"].sum(),"R$ #,##0.00"),
          ("→ Transferível 90d",df[df["CLF"].str.startswith("TRANSFERIVEL")]["VEX_90"].sum(),"R$ #,##0.00"),
          ("→ Retorno ao CD 90d",df[df["CLF"].str.startswith("RETORNO_CD")]["VEX_90"].sum(),"R$ #,##0.00"),
          ("Hospitais",df["HOSPITAL AJUSTADO"].nunique(),"#,##0"),
          ("SKUs",df["COD VALORES"].nunique(),"#,##0")]
    if df_lotes is not None:
        kpis.append(("Lotes em risco",df_lotes[df_lotes["STATUS_VEN"]=="Vence antes de consumir"].shape[0],"#,##0"))
    hdr(ws1,3,1,"Indicador",bg="2E75B6"); hdr(ws1,3,2,"Valor",bg="2E75B6")
    for i,(l,v,f) in enumerate(kpis,4):
        bg="EBF3FB" if i%2==0 else "FFFFFF"
        cel(ws1,i,1,l,bg=bg); cel(ws1,i,2,v,fmt=f,bg=bg,ha="right")
    W(ws1,[38,20])
    # Transferências
    ws2=wb.create_sheet("2. Transferências 90d")
    hdr(ws2,1,1,"TRANSFERÊNCIAS — 90d",sz=11); ws2.merge_cells("A1:L1")
    cols=["Regional","Hospital","Sigla","Código","Descrição","Classificação","Origem",
          "Físico","Ideal 90d",f"Ideal {pol_dias}d","Excesso 90d","Valor Excesso 90d"]
    [hdr(ws2,2,j+1,h,bg="2E75B6",wrap=True) for j,h in enumerate(cols)]
    for i,(_,r) in enumerate(df[df["CLF"].str.startswith("TRANSFERIVEL")].sort_values("VEX_90",ascending=False).iterrows(),3):
        bg=CORES.get(r["CLF"],"FFFFFF")
        vals=[r["REGIONAL"],r["HOSPITAL AJUSTADO"],r["SIGLA"],str(r["COD VALORES"]),r["DESCRICAO"],
              r["CLF"],"Grade" if r["ORIGEM_GRADE"] else "CMM",
              r["FÍSICO"],r["IDEAL_90"],r["IDEAL_ALT"],r["EXCESSO_90"],r["VEX_90"]]
        fmts=[None,None,None,None,None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","R$ #,##0.00"]
        [cel(ws2,i,j+1,v,fmt=f,bg=bg,ha="right" if f else "left") for j,(v,f) in enumerate(zip(vals,fmts))]
    W(ws2,[10,22,7,16,38,24,10,10,10,10,10,16]); ws2.freeze_panes="A3"
    # Retorno CD
    ws3=wb.create_sheet("3. Retorno ao CD")
    hdr(ws3,1,1,"RETORNO AO CD",sz=11); ws3.merge_cells("A1:K1")
    cols3=["Regional","Hospital","Sigla","Código","Descrição","Classificação","Origem",
           "Físico","Excesso 90d","CMM Rede","Valor a Retornar"]
    [hdr(ws3,2,j+1,h,bg="C0392B",wrap=True) for j,h in enumerate(cols3)]
    for i,(_,r) in enumerate(df[df["CLF"].str.startswith("RETORNO_CD")].sort_values("VEX_90",ascending=False).iterrows(),3):
        bg=CORES.get(r["CLF"],"FFFFFF")
        vals=[r["REGIONAL"],r["HOSPITAL AJUSTADO"],r["SIGLA"],str(r["COD VALORES"]),r["DESCRICAO"],
              r["CLF"],"Grade" if r["ORIGEM_GRADE"] else "CMM",
              r["FÍSICO"],r["EXCESSO_90"],r["CMM_REDE"],r["VEX_90"]]
        fmts=[None,None,None,None,None,None,None,"#,##0.00","#,##0.00","#,##0.00","R$ #,##0.00"]
        [cel(ws3,i,j+1,v,fmt=f,bg=bg,ha="right" if f else "left") for j,(v,f) in enumerate(zip(vals,fmts))]
    W(ws3,[10,22,7,16,38,24,10,10,10,10,16]); ws3.freeze_panes="A3"
    # Vencimento
    if df_lotes is not None:
        ws4=wb.create_sheet("4. Risco Vencimento")
        hdr(ws4,1,1,"RISCO DE VENCIMENTO POR LOTE",sz=11); ws4.merge_cells("A1:K1")
        cols4=["CD","Código","Descrição","Lote","Validade","Meses até Vencer",
               "Qtd. Lote","CMM Rede","Meses p/ Consumir","Saldo (meses)","Status"]
        [hdr(ws4,2,j+1,h,bg="922B21",wrap=True) for j,h in enumerate(cols4)]
        for i,(_,r) in enumerate(df_lotes[~df_lotes["STATUS_VEN"].str.startswith("OK")].sort_values(["STATUS_VEN","MAV"]).iterrows(),3):
            st=str(r["STATUS_VEN"]); bg="F1948A" if st=="Vence antes de consumir" else ("FAD7A0" if st=="Margem apertada" else "FDEDEC")
            mc=r["MPC"] if r["MPC"]!=np.inf else None; sd=r["SALDO"] if r["MPC"]!=np.inf else None
            vdt=datetime(r["VALIDADE_DT"].year,r["VALIDADE_DT"].month,r["VALIDADE_DT"].day) if r["VALIDADE_DT"] else None
            vals=[r.get("FILIAL",""),r["COD_STR"],r.get("DESCRICAO",""),r.get("LOTE",""),
                  vdt,r["MAV"],r["QUANTIDADE"],r["CMM_REDE"],mc,sd,st]
            fmts=[None,None,None,None,"DD/MM/YYYY","#,##0.1","#,##0.00","#,##0.00","#,##0.1","#,##0.1",None]
            [cel(ws4,i,j+1,v,fmt=f,bg=bg,ha="right" if f else "left") for j,(v,f) in enumerate(zip(vals,fmts))]
        W(ws4,[10,16,36,18,13,14,12,10,14,12,24]); ws4.freeze_panes="A3"
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ── Interface ─────────────────────────────────────────────────────────
st.title("📦 Equalização de Estoques")

with st.sidebar:
    st.header("⚙️ Configurações")
    pol_dias = st.selectbox("Política alternativa", [120,150,180], index=1)
    st.divider()
    st.markdown("**Legenda**")
    for clf,cor,desc in [
        ("Transferível","#D5F5E3","Excesso real"),
        ("Transf. - Grade","#A9DFBF","Excesso de grade"),
        ("Retorno CD","#FADBD8","Sem consumo"),
        ("Ret. CD - Grade","#F1948A","Sem consumo nem grade"),
        ("Adequado","#EBF3FB","Dentro do ideal"),
    ]:
        st.markdown(f'<span style="background:{cor};padding:2px 8px;border-radius:4px;font-size:.8rem">{clf}</span> {desc}',unsafe_allow_html=True)

st.subheader("1. Upload da base")
st.info("⚠️ O arquivo deve estar no formato **.xlsx** (Excel padrão). Para converter: abra o .xlsb no Excel → Salvar Como → Excel (.xlsx)", icon="ℹ️")
file_base = st.file_uploader("Arquivo de estudo (.xlsx)", type=["xlsx"])

if file_base:
    fb = file_base.read()
    with st.spinner("Lendo arquivo..."):
        sheets = get_sheets(fb)
    if "BASE_ESTUDOS" not in sheets:
        st.error(f"Aba BASE_ESTUDOS não encontrada. Abas disponíveis: {sheets}"); st.stop()
    with st.spinner("Carregando dados..."):
        df_raw = read_sheet(fb, "BASE_ESTUDOS")

    c1,c2 = st.columns([2,1])
    with c1:
        linhas_disp = sorted(df_raw["LINHA"].dropna().astype(str).unique())
        linhas_sel  = st.multiselect("Linhas de produto", linhas_disp,
                      default=["HEMODINAMICA"] if "HEMODINAMICA" in linhas_disp else linhas_disp[:1])
    with c2:
        regs_disp = sorted(df_raw["REGIONAL"].dropna().astype(str).unique())
        regs_sel  = st.multiselect("Regionais (opcional)", regs_disp)

    if not linhas_sel: st.warning("Selecione ao menos uma linha."); st.stop()
    df_filtrado = df_raw if not regs_sel else df_raw[df_raw["REGIONAL"].isin(regs_sel)]

    st.subheader("2. Base de lotes (opcional)")
    file_lotes = st.file_uploader("Lotes e validades (.xlsx ou .csv)", type=["xlsx","csv"])
    df_lotes = None
    usar_cmv2 = False
    if file_lotes is None and "CMV_2" in sheets:
        usar_cmv2 = st.checkbox("Usar aba CMV_2 do próprio arquivo", value=True)

    if st.button("▶ Rodar análise", type="primary", use_container_width=True):
        with st.spinner("Processando..."):
            df = processar(df_filtrado, linhas_sel, pol_dias)
            if usar_cmv2:
                df_lotes = processar_lotes(read_sheet(fb,"CMV_2"), df)
            elif file_lotes:
                rl = pd.read_csv(file_lotes) if file_lotes.name.endswith(".csv") else pd.read_excel(file_lotes)
                df_lotes = processar_lotes(rl, df)
        st.session_state.update({"df":df,"df_lotes":df_lotes,"pol_dias":pol_dias,"linhas_sel":linhas_sel})

if "df" in st.session_state:
    df       = st.session_state["df"]
    df_lotes = st.session_state["df_lotes"]
    pol_dias = st.session_state["pol_dias"]
    linhas_sel = st.session_state["linhas_sel"]

    st.divider()
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("💰 Valor físico",     f"R$ {df['VALOR TOTAL '].sum():,.0f}")
    c2.metric("📉 Excesso 90d",      f"R$ {df['VEX_90'].sum():,.0f}",
              delta=f"-R$ {df['VEX_90'].sum()-df['VEX_ALT'].sum():,.0f} em {pol_dias}d",delta_color="inverse")
    c3.metric("🔀 Transferível 90d", f"R$ {df[df['CLF'].str.startswith('TRANSFERIVEL')]['VEX_90'].sum():,.0f}")
    c4.metric("🏭 Retorno ao CD",    f"R$ {df[df['CLF'].str.startswith('RETORNO_CD')]['VEX_90'].sum():,.0f}")

    if df_lotes is not None:
        c5,c6,_,_ = st.columns(4)
        c5.metric("⚠️ Lotes em risco", df_lotes[df_lotes["STATUS_VEN"]=="Vence antes de consumir"].shape[0])
        c6.metric("🧬 SKUs em risco",  df_lotes[df_lotes["STATUS_VEN"]=="Vence antes de consumir"]["COD_STR"].nunique())

    st.divider()

    tabs = ["📊 Por Regional","🏥 Por Hospital","🔀 Transferências","🏭 Retorno ao CD"]
    if df_lotes is not None: tabs.append("⚠️ Risco de Vencimento")
    tabs.append("⬇️ Download Excel")
    tab_list = st.tabs(tabs)
    t_reg,t_hosp,t_transf,t_ret = tab_list[0],tab_list[1],tab_list[2],tab_list[3]
    t_ven = tab_list[4] if df_lotes is not None else None
    t_dl  = tab_list[-1]

    with t_reg:
        st.markdown("#### Excesso por regional")
        reg = df.groupby("REGIONAL").agg(
            Físico=("VALOR TOTAL ","sum"), Excesso_90d=("VEX_90","sum"),
            Excesso_alt=("VEX_ALT","sum"), Hospitais=("HOSPITAL AJUSTADO","nunique"),
            SKUs=("COD VALORES","count")
        ).reset_index().sort_values("Excesso_90d",ascending=False)
        reg["%_Excesso"] = (reg["Excesso_90d"]/reg["Físico"]*100).round(1)
        st.bar_chart(reg.set_index("REGIONAL")[["Físico","Excesso_90d"]])
        reg_disp = reg.copy()
        reg_disp["Físico"]      = reg_disp["Físico"].apply(lambda x: f"R$ {x:,.0f}")
        reg_disp["Excesso_90d"] = reg_disp["Excesso_90d"].apply(lambda x: f"R$ {x:,.0f}")
        reg_disp["Excesso_alt"] = reg_disp["Excesso_alt"].apply(lambda x: f"R$ {x:,.0f}")
        reg_disp["%_Excesso"]   = reg_disp["%_Excesso"].apply(lambda x: f"{x:.1f}%")
        reg_disp.columns = ["Regional","Valor Físico","Excesso 90d",f"Excesso {pol_dias}d","Hospitais","SKUs","% Excesso"]
        st.dataframe(reg_disp, use_container_width=True, hide_index=True)
        st.markdown("#### Distribuição por classificação")
        st.dataframe(df["CLF"].value_counts().reset_index().rename(columns={"CLF":"Classificação","count":"Itens"}),
                     use_container_width=True, hide_index=True)

    with t_hosp:
        st.markdown("#### Top 20 hospitais por excesso (90d)")
        hosp = df.groupby(["HOSPITAL AJUSTADO","SIGLA","REGIONAL"]).agg(
            Físico=("VALOR TOTAL ","sum"), Excesso_90d=("VEX_90","sum"),
            Excesso_alt=("VEX_ALT","sum"), SKUs=("COD VALORES","count"),
            Sem_consumo=("CMM",lambda x:(x==0).sum())
        ).reset_index().sort_values("Excesso_90d",ascending=False)
        hosp["%_Excesso"] = (hosp["Excesso_90d"]/hosp["Físico"]*100).round(1)
        st.bar_chart(hosp.head(20).set_index("HOSPITAL AJUSTADO")[["Físico","Excesso_90d"]])
        hosp_disp = hosp.copy()
        hosp_disp["Físico"]      = hosp_disp["Físico"].apply(lambda x: f"R$ {x:,.0f}")
        hosp_disp["Excesso_90d"] = hosp_disp["Excesso_90d"].apply(lambda x: f"R$ {x:,.0f}")
        hosp_disp["Excesso_alt"] = hosp_disp["Excesso_alt"].apply(lambda x: f"R$ {x:,.0f}")
        hosp_disp["%_Excesso"]   = hosp_disp["%_Excesso"].apply(lambda x: f"{x:.1f}%")
        hosp_disp.columns = ["Hospital","Sigla","Regional","Valor Físico","Excesso 90d",
                              f"Excesso {pol_dias}d","SKUs","Sem Consumo","% Excesso"]
        st.dataframe(hosp_disp, use_container_width=True, hide_index=True)

    with t_transf:
        st.markdown("#### Itens para redistribuição entre hospitais")
        col_f1,col_f2 = st.columns(2)
        with col_f1:
            reg_f = st.selectbox("Filtrar regional",["Todas"]+sorted(df["REGIONAL"].unique().tolist()),key="reg_transf")
        with col_f2:
            tipo_f = st.selectbox("Tipo de excesso",["Todos","Consumo real","Excesso de grade"],key="tipo_transf")
        transf = df[df["CLF"].str.startswith("TRANSFERIVEL")].copy()
        if reg_f!="Todas":           transf=transf[transf["REGIONAL"]==reg_f]
        if tipo_f=="Consumo real":   transf=transf[transf["CLF"]=="TRANSFERIVEL"]
        if tipo_f=="Excesso de grade": transf=transf[transf["CLF"]=="TRANSFERIVEL - GRADE"]
        transf_disp = transf[["REGIONAL","HOSPITAL AJUSTADO","SIGLA","COD VALORES","DESCRICAO",
                               "CLF","FÍSICO","IDEAL_90","EXCESSO_90","VALOR UNIT","VEX_90"]].copy()
        transf_disp.columns=["Regional","Hospital","Sigla","Código","Descrição",
                              "Classificação","Físico","Ideal 90d","Excesso 90d","Valor Unit.","Valor Excesso"]
        transf_disp=transf_disp.sort_values("Valor Excesso",ascending=False)
        st.caption(f"{len(transf_disp):,} itens  |  R$ {transf_disp['Valor Excesso'].sum():,.0f} em excesso")
        st.dataframe(transf_disp.style.format({
            "Físico":"{:.2f}","Ideal 90d":"{:.2f}","Excesso 90d":"{:.2f}",
            "Valor Unit.":"R$ {:,.2f}","Valor Excesso":"R$ {:,.2f}"}),
            use_container_width=True, hide_index=True, height=500)

    with t_ret:
        st.markdown("#### Itens para retorno ao CD")
        col_f3,col_f4 = st.columns(2)
        with col_f3:
            reg_r = st.selectbox("Filtrar regional",["Todas"]+sorted(df["REGIONAL"].unique().tolist()),key="reg_ret")
        with col_f4:
            tipo_r = st.selectbox("Tipo",["Todos","Sem consumo","Grade sem consumo"],key="tipo_ret")
        ret = df[df["CLF"].str.startswith("RETORNO_CD")].copy()
        if reg_r!="Todas":              ret=ret[ret["REGIONAL"]==reg_r]
        if tipo_r=="Sem consumo":       ret=ret[ret["CLF"]=="RETORNO_CD"]
        if tipo_r=="Grade sem consumo": ret=ret[ret["CLF"]=="RETORNO_CD - GRADE"]
        ret_disp=ret[["REGIONAL","HOSPITAL AJUSTADO","SIGLA","COD VALORES","DESCRICAO",
                       "CLF","FÍSICO","IDEAL_90","EXCESSO_90","CMM_REDE","VALOR UNIT","VEX_90"]].copy()
        ret_disp.columns=["Regional","Hospital","Sigla","Código","Descrição",
                           "Classificação","Físico","Ideal 90d","Excesso 90d","CMM Rede","Valor Unit.","Valor a Retornar"]
        ret_disp=ret_disp.sort_values("Valor a Retornar",ascending=False)
        st.caption(f"{len(ret_disp):,} itens  |  R$ {ret_disp['Valor a Retornar'].sum():,.0f} a retornar")
        st.dataframe(ret_disp.style.format({
            "Físico":"{:.2f}","Ideal 90d":"{:.2f}","Excesso 90d":"{:.2f}",
            "CMM Rede":"{:.2f}","Valor Unit.":"R$ {:,.2f}","Valor a Retornar":"R$ {:,.2f}"}),
            use_container_width=True, hide_index=True, height=500)

    if t_ven is not None:
        with t_ven:
            st.markdown("#### Lotes com risco de vencer antes de ser consumidos")
            status_f = st.selectbox("Filtrar status",
                ["Todos com risco","Vence antes de consumir","Margem apertada","Sem consumo"],key="status_ven")
            ven = df_lotes[~df_lotes["STATUS_VEN"].str.startswith("OK")].copy()
            if status_f=="Vence antes de consumir": ven=ven[ven["STATUS_VEN"]=="Vence antes de consumir"]
            elif status_f=="Margem apertada":        ven=ven[ven["STATUS_VEN"]=="Margem apertada"]
            elif status_f=="Sem consumo":            ven=ven[ven["STATUS_VEN"].str.startswith("Sem consumo")]
            ca,cb,cc = st.columns(3)
            ca.metric("Vence antes de consumir", ven[ven["STATUS_VEN"]=="Vence antes de consumir"].shape[0])
            cb.metric("Margem apertada",          ven[ven["STATUS_VEN"]=="Margem apertada"].shape[0])
            cc.metric("Sem consumo",              ven[ven["STATUS_VEN"].str.startswith("Sem consumo")].shape[0])
            ven_disp = ven[["COD_STR","DESCRICAO","VALIDADE_DT","MAV","QUANTIDADE","CMM_REDE","SALDO","STATUS_VEN"]].copy()
            ven_disp["MPC_txt"] = ven["MPC"].apply(lambda x: "∞" if x==np.inf else f"{x:.1f}")
            ven_disp.columns=["Código","Descrição","Validade","Meses até Vencer","Qtd. Lote","CMM Rede","Saldo (meses)","Status","Meses p/ Consumir"]
            st.dataframe(ven_disp.sort_values(["Status","Meses até Vencer"]),
                         use_container_width=True, hide_index=True, height=500)

    with t_dl:
        st.markdown("#### Download completo em Excel")
        st.markdown("O arquivo contém todas as análises com formatação de cores.")
        buf  = make_excel(df, df_lotes, pol_dias)
        nome = f"Equalizacao_{'_'.join(linhas_sel)}_{date.today().strftime('%Y%m%d')}.xlsx"
        st.download_button("⬇️ Baixar Excel", data=buf, file_name=nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary")
