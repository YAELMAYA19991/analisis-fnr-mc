
import io, re
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Control FNR & MC", page_icon="📊", layout="wide")
FNR_OBJ, MC_OBJ = 1.50, 1.00

def norm(x):
    x = str(x).strip().lower().translate(str.maketrans("áéíóúüñ","aeiouun"))
    return re.sub(r"[^a-z0-9]+","_",x).strip("_")

def clean(df):
    x=df.copy(); x.columns=[norm(c) for c in x.columns]; return x

def col(df,names):
    d={norm(c):c for c in df.columns}
    for n in names:
        if norm(n) in d: return d[norm(n)]
    for n in names:
        for c in df.columns:
            if norm(n) in norm(c): return c
    return None

def sheets(upload):
    xls=pd.ExcelFile(upload); out={}
    for sh in xls.sheet_names:
        df=pd.read_excel(upload,sheet_name=sh)
        if not df.empty: out[sh]=clean(df)
    return out

def choose(ss,words):
    for name,df in ss.items():
        if any(norm(w) in norm(name) for w in words): return df
    return next(iter(ss.values()))

def parse_base(df):
    p=col(df,["picker","picker_nombre","nombre","email_picker"])
    l=col(df,["total_lineas","lineas_totales","lineas","total_lineas_pickeadas"])
    ped=col(df,["total_pedidos","pedidos_totales","pedidos"])
    sh=col(df,["turno","shift"]); ar=col(df,["area","departamento","department"])
    if not p or not l: raise ValueError("La base necesita PICKER y Total lineas.")
    x=pd.DataFrame({
        "PICKER":df[p].astype(str).str.strip(),
        "LINEAS":pd.to_numeric(df[l],errors="coerce").fillna(0),
        "PEDIDOS":pd.to_numeric(df[ped],errors="coerce").fillna(0) if ped else 0,
        "TURNO":df[sh].astype(str).str.strip() if sh else "No especificado",
        "AREA_BASE":df[ar].astype(str).str.strip() if ar else "No especificada"})
    return x.sort_values("LINEAS").drop_duplicates("PICKER",keep="last")


def parse_roster(df, turno_fijo=None):
    """Maestro de personal.
    Requiere PICKER y opcionalmente SUPERVISOR/AREA_BASE.
    Si turno_fijo viene de la pantalla de carga, ese turno manda.
    """
    p=col(df,["picker","picker_nombre","email_picker","nombre","persona"])
    sh=col(df,["turno","shift"])
    sup=col(df,["supervisor","supervisor_nombre","jefe","responsable"])
    ar=col(df,["area_base","area","departamento","department"])
    if not p:
        raise ValueError("La plantilla de personal necesita la columna PICKER.")
    x=pd.DataFrame({
        "PICKER":df[p].astype(str).str.strip(),
        "TURNO_MAESTRO":(df[sh].astype(str).str.strip() if sh and not turno_fijo else turno_fijo or ""),
        "SUPERVISOR":df[sup].astype(str).str.strip() if sup else "No asignado",
        "AREA_MAESTRO":df[ar].astype(str).str.strip() if ar else ""
    })
    x=x[x["PICKER"].str.strip().ne("")].copy()
    x=x.drop_duplicates("PICKER",keep="last")
    return x

def apply_roster(base, roster):
    if roster is None or roster.empty: return base
    r=roster[["PICKER","TURNO_MAESTRO","SUPERVISOR","AREA_MAESTRO"]].copy()
    x=base.merge(r,on="PICKER",how="left")
    tm=x["TURNO_MAESTRO"].fillna("").astype(str).str.strip()
    tb=x["TURNO"].fillna("").astype(str).str.strip()
    x["TURNO"]=tm.where(tm.ne(""),tb)
    am=x["AREA_MAESTRO"].fillna("").astype(str).str.strip()
    ab=x["AREA_BASE"].fillna("").astype(str).str.strip()
    x["AREA_BASE"]=am.where(am.ne(""),ab)
    x["SUPERVISOR"]=x["SUPERVISOR"].fillna("No asignado").astype(str).str.strip()
    return x.drop(columns=["TURNO_MAESTRO","AREA_MAESTRO"])

def roster_from_uploads(uploads):
    frames=[]
    for turno,up in uploads.items():
        if up:
            frames.append(parse_roster(choose(sheets(up),["turno","personal","plantilla","maestro"]),turno))
    if not frames: return None
    r=pd.concat(frames,ignore_index=True)
    return r.drop_duplicates("PICKER",keep="last")

def parse_inc(df,tipo):
    p=col(df,["picker","picker_nombre","email_picker","nombre"])
    prod=col(df,["product","producto","item","articulo","artículo"])
    order=col(df,["order_number","order","pedido","numero_pedido","order_id"])
    ar=col(df,["department","departamento","area","depto"])
    sh=col(df,["turno","shift"]); fe=col(df,["date","fecha","created_at"])
    qty=col(df,["cantidad","qty","quantity"])
    if not p: raise ValueError(f"El detalle {tipo} necesita PICKER.")
    x=pd.DataFrame({
        "PICKER":df[p].astype(str).str.strip(),
        "PRODUCTO":df[prod].astype(str).str.strip() if prod else "Sin producto",
        "ORDER_NUMBER":df[order].astype(str).str.strip() if order else "",
        "AREA":df[ar].astype(str).str.strip() if ar else "No especificada",
        "TURNO":df[sh].astype(str).str.strip() if sh else "No especificado",
        "FECHA":df[fe] if fe else "",
        "INCIDENCIAS":pd.to_numeric(df[qty],errors="coerce").fillna(1) if qty else 1,
        "TIPO":tipo})
    x["INCIDENCIAS"]=x["INCIDENCIAS"].where(x["INCIDENCIAS"]>0,1)
    return x

def attach(inc,base):
    ref=base[["PICKER","LINEAS","PEDIDOS","TURNO"]].rename(columns={"TURNO":"TURNO_REF"})
    return inc.merge(ref,on="PICKER",how="left")

def summary(base,fnr,mc):
    s=base.copy()
    f=fnr.groupby("PICKER",as_index=False).INCIDENCIAS.sum().rename(columns={"INCIDENCIAS":"FNR"})
    m=mc.groupby("PICKER",as_index=False).INCIDENCIAS.sum().rename(columns={"INCIDENCIAS":"MC"})
    s=s.merge(f,on="PICKER",how="left").merge(m,on="PICKER",how="left")
    s[["FNR","MC"]]=s[["FNR","MC"]].fillna(0)
    # Asegurar columnas numéricas y evitar pd.NA + round() incompatibles con algunas versiones de pandas
    s["LINEAS"]=pd.to_numeric(s["LINEAS"],errors="coerce").fillna(0.0)
    s["FNR"]=pd.to_numeric(s["FNR"],errors="coerce").fillna(0.0)
    s["MC"]=pd.to_numeric(s["MC"],errors="coerce").fillna(0.0)
    den=s["LINEAS"].where(s["LINEAS"].ne(0), float("nan"))
    s["FNR_%"]=pd.to_numeric(s["FNR"].div(den)*100,errors="coerce").round(2)
    s["MC_%"]=pd.to_numeric(s["MC"].div(den)*100,errors="coerce").round(2)
    def sem(r):
        if (pd.notna(r["FNR_%"]) and r["FNR_%"]>=FNR_OBJ) or (pd.notna(r["MC_%"]) and r["MC_%"]>=MC_OBJ):
            return "🔴 FUERA DE OBJETIVO"
        if (pd.notna(r["FNR_%"]) and r["FNR_%"]>=1.20) or (pd.notna(r["MC_%"]) and r["MC_%"]>=0.80):
            return "🟡 PREVENTIVO"
        return "🟢 EN OBJETIVO"
    s["ESTADO"]=s.apply(sem,axis=1)
    return s

def monthly_kpis(base,fnr,mc):
    total_pedidos=pd.to_numeric(base["PEDIDOS"],errors="coerce").fillna(0).sum()
    fnr_pedidos=fnr.loc[fnr["ORDER_NUMBER"].astype(str).str.strip().ne(""),"ORDER_NUMBER"].nunique()
    mc_pedidos=mc.loc[mc["ORDER_NUMBER"].astype(str).str.strip().ne(""),"ORDER_NUMBER"].nunique()
    # Si no hay números de pedido en el detalle, usamos incidencias como respaldo y lo indicamos en la UI.
    fnr_rate=fnr_pedidos/total_pedidos*100 if total_pedidos else None
    mc_rate=mc_pedidos/total_pedidos*100 if total_pedidos else None
    return total_pedidos,fnr_pedidos,mc_pedidos,fnr_rate,mc_rate

def feedback_export(rows):
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine="openpyxl") as w:
        pd.DataFrame(rows).to_excel(w,"Retroalimentacion",index=False)
    return b.getvalue()

def products(inc,picker=None):
    x=inc if not picker or picker=="Todos" else inc[inc.PICKER==picker]
    return x.groupby("PRODUCTO",as_index=False).INCIDENCIAS.sum().rename(columns={"INCIDENCIAS":"CANTIDAD"}).sort_values("CANTIDAD",ascending=False)

def orders(inc,picker=None):
    x=inc if not picker or picker=="Todos" else inc[inc.PICKER==picker]
    return x.groupby("ORDER_NUMBER",as_index=False).agg(INCIDENCIAS=("INCIDENCIAS","sum"),PICKERS=("PICKER","nunique"),PRODUCTOS=("PRODUCTO","nunique")).sort_values(["INCIDENCIAS","PICKERS"],ascending=False)

def groups(inc,base,key):
    x=inc.copy()
    if key=="TURNO":
        x["GRUPO"]=x.TURNO_REF
        den=base.groupby("TURNO",as_index=False).LINEAS.rename(columns={"TURNO":"GRUPO"})
    else:
        x["GRUPO"]=x.AREA
        if (base.AREA_BASE!="No especificada").any():
            den=base.groupby("AREA_BASE",as_index=False).LINEAS.rename(columns={"AREA_BASE":"GRUPO"})
        else: den=None
    g=x.groupby("GRUPO",as_index=False).INCIDENCIAS.sum()
    g["% DEL TOTAL"]=(g.INCIDENCIAS/g.INCIDENCIAS.sum()*100).round(2)
    if den is not None:
        g=g.merge(den,on="GRUPO",how="left")
        g["LINEAS"]=pd.to_numeric(g["LINEAS"],errors="coerce").fillna(0.0)
        den=g["LINEAS"].where(g["LINEAS"].ne(0), float("nan"))
        g["% / LINEAS"]=pd.to_numeric(g.INCIDENCIAS.div(den)*100,errors="coerce").round(2)
    return g.sort_values("INCIDENCIAS",ascending=False)

def export(summary,fnr,mc,picker,roster=None):
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine="openpyxl") as w:
        summary.to_excel(w,"Resumen_Pickers",index=False)
        fnr.to_excel(w,"Detalle_FNR",index=False); mc.to_excel(w,"Detalle_MC",index=False)
        products(fnr,picker).to_excel(w,"Productos_FNR",index=False)
        products(mc,picker).to_excel(w,"Productos_MC",index=False)
        orders(fnr,picker).to_excel(w,"Pedidos_FNR",index=False)
        orders(mc,picker).to_excel(w,"Pedidos_MC",index=False)
        if roster is not None: roster.to_excel(w,"Maestro_Turnos",index=False)
    return b.getvalue()

st.title("📊 Control FNR & Mala Calidad")
st.caption("Dashboard automático por picker, turno, área, producto y pedido")

with st.sidebar:
    st.header("Carga")
    ub=st.file_uploader("① Base de líneas / pickers",type=["xlsx","xls"])
    uf=st.file_uploader("② Detalle FNR",type=["xlsx","xls"])
    um=st.file_uploader("③ Detalle Mala Calidad",type=["xlsx","xls"])
    st.subheader("Personal por turno")
    st.caption("Carga una plantilla por turno. La app asigna automáticamente el turno del archivo.")
    u_mat=st.file_uploader("① Matutino",type=["xlsx","xls"],key="roster_mat")
    u_int=st.file_uploader("② Intermedio",type=["xlsx","xls"],key="roster_int")
    u_ves=st.file_uploader("③ Vespertino",type=["xlsx","xls"],key="roster_ves")
    u_noc=st.file_uploader("④ Nocturno",type=["xlsx","xls"],key="roster_noc")
    st.caption("Cada plantilla requiere PICKER. SUPERVISOR y AREA_BASE son opcionales.")
    st.divider(); periodo=st.text_input("Periodo",datetime.now().strftime("%Y-%m"))
    ex=st.text_area("Pedidos operativos a excluir (uno por línea)")
    excluded={x.strip() for x in ex.splitlines() if x.strip()}

if not (ub and uf and um):
    st.info("Carga los 3 Excel para comenzar.")
    st.markdown("**La V4 incluye:** FNR/MC sobre líneas, seguimiento individual, artículos, pedidos repetidos, turnos, áreas, exclusiones operativas, exportación Excel y maestro de personas por turno.")
    st.stop()

try:
    base=parse_base(choose(sheets(ub),["picker","lineas","resumen"]))
    roster = roster_from_uploads({"Matutino":u_mat,"Intermedio":u_int,"Vespertino":u_ves,"Nocturno":u_noc})
    base = apply_roster(base, roster)
    fnr=parse_inc(choose(sheets(uf),["fnr","detalle"]),"FNR")
    mc=parse_inc(choose(sheets(um),["mc","mala"]),"MC")
except Exception as e:
    st.error(str(e)); st.stop()

if excluded:
    fnr=fnr[~fnr.ORDER_NUMBER.isin(excluded)].copy()
    mc=mc[~mc.ORDER_NUMBER.isin(excluded)].copy()
fnr=attach(fnr,base); mc=attach(mc,base); s=summary(base,fnr,mc)

if roster is not None:
    matched=base["PICKER"].isin(roster["PICKER"]).sum()
    total=len(base); unmatched=total-matched
    with st.sidebar:
        st.success(f"Personal cruzado: {matched}/{total} pickers.")
        if unmatched: st.warning(f"{unmatched} pickers de la base no aparecen en las plantillas.")

with st.sidebar:
    pickers = ["Todos"]
    pickers.extend(sorted({str(x) for x in s["PICKER"].tolist() if str(x).strip()}))
    turns = ["Todos"]
    turns.extend(sorted({str(x) for x in s["TURNO"].tolist() if str(x).strip()}))
    sp=st.selectbox("Picker",pickers); stn=st.selectbox("Turno",turns)

a,b,c,d,e,g,f=st.tabs(["🏠 Bodega","👤 Picker","🏷️ Artículos","📦 Pedidos","🌙 Turnos / Áreas","👥 Supervisores","📥 Exportar"])

with a:
    lines=base.LINEAS.sum(); F=fnr.INCIDENCIAS.sum(); M=mc.INCIDENCIAS.sum()
    total_pedidos,fnr_pedidos,mc_pedidos,fnr_rate,mc_rate=monthly_kpis(base,fnr,mc)
    st.subheader(f"Resumen de bodega — {periodo}")
    q=st.columns(6)
    q[0].metric("Líneas",f"{lines:,.0f}")
    q[1].metric("Pedidos",f"{total_pedidos:,.0f}")
    q[2].metric("FNR mensual",f"{fnr_rate:.2f}%" if fnr_rate is not None else "N/D", "Objetivo < 1.50%")
    q[3].metric("MC mensual",f"{mc_rate:.2f}%" if mc_rate is not None else "N/D", "Objetivo < 1.00%")
    q[4].metric("Pickers",len(s))
    q[5].metric("Fuera objetivo",(s.ESTADO=="🔴 FUERA DE OBJETIVO").sum())
    st.caption(f"KPI mensual: {fnr_pedidos:,} pedidos con FNR / {total_pedidos:,} pedidos = {fnr_rate:.2f}% · {mc_pedidos:,} pedidos con MC / {total_pedidos:,} pedidos = {mc_rate:.2f}%" if fnr_rate is not None and mc_rate is not None else "No hay suficientes pedidos para calcular el KPI mensual.")
    st.divider()
    st.subheader("Indicador operativo por líneas")
    k=st.columns(2); k[0].metric("FNR / líneas",f"{F/lines*100:.2f}%" if lines else "N/D"); k[1].metric("MC / líneas",f"{M/lines*100:.2f}%" if lines else "N/D")
    st.dataframe(s,use_container_width=True,hide_index=True)

with b:
    if sp=="Todos": st.info("Selecciona un picker.")
    else:
        r=s[s.PICKER==sp].iloc[0]
        q=st.columns(4); q[0].metric("Líneas",f"{r.LINEAS:,.0f}"); q[1].metric("FNR",f"{r.FNR:,.0f}",f"{r['FNR_%']:.2f}%")
        q[2].metric("MC",f"{r.MC:,.0f}",f"{r['MC_%']:.2f}%"); q[3].metric("Estado",r.ESTADO)
        st.subheader("Artículos FNR"); st.dataframe(products(fnr,sp),use_container_width=True,hide_index=True)
        st.subheader("Artículos MC"); st.dataframe(products(mc,sp),use_container_width=True,hide_index=True)
        st.subheader("Pedidos FNR"); st.dataframe(orders(fnr,sp),use_container_width=True,hide_index=True)

with c:
    tipo=st.radio("Tipo",["FNR","MC"],horizontal=True); data=fnr if tipo=="FNR" else mc
    if sp!="Todos": data=data[data.PICKER==sp]
    if stn!="Todos": data=data[data.TURNO_REF.astype(str)==stn]
    st.dataframe(data.groupby(["PRODUCTO","AREA"],as_index=False).INCIDENCIAS.sum().rename(columns={"INCIDENCIAS":"CANTIDAD"}).sort_values("CANTIDAD",ascending=False),use_container_width=True,hide_index=True)

with d:
    tipo=st.radio("Incidencia",["FNR","MC"],horizontal=True); data=fnr if tipo=="FNR" else mc
    if sp!="Todos": data=data[data.PICKER==sp]
    st.dataframe(orders(data),use_container_width=True,hide_index=True)
    st.caption("PICKERS indica cuántos pickers aparecen en el mismo pedido.")

with e:
    if ut:
        st.success("Los turnos mostrados aquí provienen del maestro de personas cargado en ④.")
    else:
        st.info("Puedes cargar un maestro de personas por turno en ④ para asignar el turno de cada picker y cruzarlo con FNR/MC.")
    st.subheader("FNR por turno"); st.dataframe(groups(fnr,base,"TURNO"),use_container_width=True,hide_index=True)
    st.subheader("MC por turno"); st.dataframe(groups(mc,base,"TURNO"),use_container_width=True,hide_index=True)
    st.subheader("FNR por área"); st.dataframe(groups(fnr,base,"AREA"),use_container_width=True,hide_index=True)
    st.subheader("MC por área"); st.dataframe(groups(mc,base,"AREA"),use_container_width=True,hide_index=True)
    st.warning("El % / líneas por área solo aparece si existe un denominador real de líneas por área.")

with g:
    st.subheader("Retroalimentación por turno")
    if roster is None or roster.empty:
        st.info("Carga las plantillas de Matutino, Intermedio, Vespertino y/o Nocturno en el panel lateral.")
    else:
        turnos=[x for x in ["Matutino","Intermedio","Vespertino","Nocturno"] if x in set(roster["TURNO_MAESTRO"])]
        turno_sel=st.selectbox("Turno",turnos if turnos else ["Sin turno"])
        rt=roster[roster["TURNO_MAESTRO"]==turno_sel].copy()
        supervisores=sorted([str(x) for x in rt["SUPERVISOR"].dropna().unique() if str(x).strip() and str(x)!="No asignado"])
        sup_sel=st.selectbox("Supervisor",["Todos"]+supervisores)
        if sup_sel!="Todos": rt=rt[rt["SUPERVISOR"]==sup_sel]
        disponibles=[x for x in rt["PICKER"].tolist() if x in set(s["PICKER"])]
        picker_sel=st.selectbox("Picker para retroalimentación",["Selecciona..."]+sorted(disponibles))
        if picker_sel!="Selecciona...":
            r=s[s["PICKER"]==picker_sel].iloc[0]
            st.markdown(f"### {picker_sel}")
            q=st.columns(4); q[0].metric("Líneas",f"{r.LINEAS:,.0f}"); q[1].metric("FNR / líneas",f"{r['FNR_%']:.2f}%"); q[2].metric("MC / líneas",f"{r['MC_%']:.2f}%"); q[3].metric("Estado",r.ESTADO)
            st.subheader("Incidencias del picker")
            cc=st.columns(2)
            with cc[0]: st.dataframe(products(fnr,picker_sel).head(15),use_container_width=True,hide_index=True)
            with cc[1]: st.dataframe(products(mc,picker_sel).head(15),use_container_width=True,hide_index=True)
            st.subheader("Retroalimentación")
            supervisor_actual=sup_sel if sup_sel!="Todos" else (rt["SUPERVISOR"].iloc[0] if len(rt) else "No asignado")
            fortalezas=st.text_area("Fortalezas observadas",key=f"fort_{picker_sel}")
            mejora=st.text_area("Punto de mejora",key=f"mej_{picker_sel}")
            acuerdo=st.text_area("Acuerdo / acción de mejora",key=f"acuerdo_{picker_sel}")
            fecha=st.date_input("Fecha de retroalimentación",key=f"fecha_{picker_sel}")
            if st.button("Guardar retroalimentación",key=f"save_{picker_sel}"):
                row={"FECHA":str(fecha),"SUPERVISOR":supervisor_actual,"TURNO":turno_sel,"PICKER":picker_sel,"FNR_LINEAS_%":float(r["FNR_%"]),"MC_LINEAS_%":float(r["MC_%"]),"ESTADO":r["ESTADO"],"FORTALEZAS":fortalezas,"PUNTO_MEJORA":mejora,"ACUERDO":acuerdo}
                st.session_state.setdefault("feedback_rows",[]).append(row)
                st.success("Retroalimentación guardada en esta sesión.")
        if st.session_state.get("feedback_rows"):
            st.subheader("Retroalimentaciones capturadas en esta sesión")
            fb=pd.DataFrame(st.session_state["feedback_rows"])
            st.dataframe(fb,use_container_width=True,hide_index=True)
            st.download_button("📥 Descargar retroalimentaciones",feedback_export(st.session_state["feedback_rows"]),f"Retroalimentacion_{periodo}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with f:
    st.download_button("📥 Descargar Excel completo",export(s,fnr,mc,None if sp=="Todos" else sp,roster),f"Analisis_FNR_MC_{periodo}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.info("La V4 permite cargar personal por cuatro turnos y capturar retroalimentación por supervisor.")
