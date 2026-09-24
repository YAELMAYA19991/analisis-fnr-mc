import io, re, json, os
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Control FNR & MC", page_icon="📊", layout="wide")
FNR_OBJ, MC_OBJ = 1.50, 1.00
STORE_FILE = "picker_seguimiento.json"

def load_store():
    if os.path.exists(STORE_FILE):
        try:
            with open(STORE_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except Exception: pass
    return {"pickers": {}}

def save_store(store):
    tmp = STORE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f: json.dump(store, f, ensure_ascii=False, indent=2)
    os.replace(tmp, STORE_FILE)

def picker_record(store, picker):
    rec = store.setdefault("pickers", {}).setdefault(picker, {"estado":"ACTIVO", "comentarios":[], "acciones":[]})
    rec.setdefault("estado", "ACTIVO"); rec.setdefault("comentarios", []); rec.setdefault("acciones", [])
    return rec

def active_picker_names(store):
    return {p for p,r in store.get("pickers",{}).items() if r.get("estado", "ACTIVO") == "ACTIVO"}

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
    """Maestro consolidado de personal: PICKER, TURNO, CORREO, SUPERVISOR y AREA."""
    p=col(df,["picker","picker_nombre","email_picker","nombre","persona"])
    sh=col(df,["turno","shift"])
    em=col(df,["codigo_correo","codigo + correo","correo","email","usuario"])
    sup=col(df,["supervisor","supervisor_nombre","jefe","responsable"])
    ar=col(df,["area_base","area","departamento","department"])
    if not p:
        raise ValueError("La hoja de personal necesita la columna PICKER.")
    x=pd.DataFrame({
        "PICKER":df[p].astype(str).str.strip(),
        "TURNO_MAESTRO":(df[sh].astype(str).str.strip() if sh and not turno_fijo else turno_fijo or ""),
        "CORREO":df[em].astype(str).str.strip() if em else "",
        "SUPERVISOR":df[sup].astype(str).str.strip() if sup else "No asignado",
        "AREA_MAESTRO":df[ar].astype(str).str.strip() if ar else ""
    })
    x=x[x["PICKER"].str.strip().ne("")].copy()
    x=x.drop_duplicates("PICKER",keep="last")
    return x

def apply_roster(base, roster):
    if roster is None or roster.empty: return base
    cols=["PICKER","TURNO_MAESTRO","CORREO","SUPERVISOR","AREA_MAESTRO"]
    r=roster[[c for c in cols if c in roster.columns]].copy()
    x=base.merge(r,on="PICKER",how="left")
    tm=x["TURNO_MAESTRO"].fillna("").astype(str).str.strip()
    tb=x["TURNO"].fillna("").astype(str).str.strip()
    x["TURNO"]=tm.where(tm.ne(""),tb)
    am=x["AREA_MAESTRO"].fillna("").astype(str).str.strip()
    ab=x["AREA_BASE"].fillna("").astype(str).str.strip()
    x["AREA_BASE"]=am.where(am.ne(""),ab)
    x["SUPERVISOR"]=x["SUPERVISOR"].fillna("No asignado").astype(str).str.strip()
    x["CORREO"]=x["CORREO"].fillna("").astype(str).str.strip()
    return x.drop(columns=["TURNO_MAESTRO","AREA_MAESTRO"],errors="ignore")

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
    uc=st.file_uploader("📤 Excel consolidado",type=["xlsx","xls"])
    st.caption("Un solo archivo con las hojas Base_Pickers, Detalle_FNR, Detalle_MC y Maestro_Personal.")
    st.divider(); periodo=st.text_input("Periodo",datetime.now().strftime("%Y-%m"))
    ex=st.text_area("Pedidos operativos a excluir (uno por línea)")
    excluded={x.strip() for x in ex.splitlines() if x.strip()}

if not uc:
    st.info("Carga un solo Excel consolidado para comenzar.")
    st.markdown("**Formato:** Base_Pickers + Detalle_FNR + Detalle_MC + Maestro_Personal. El maestro permite filtrar por picker, turno, correo, supervisor y área.")
    st.stop()

try:
    ss=sheets(uc)
    base_df=choose(ss,["base_pickers","picker","lineas","resumen"])
    fnr_df=choose({k:v for k,v in ss.items() if "fnr" in norm(k)},["fnr","detalle"])
    mc_df=choose({k:v for k,v in ss.items() if "mc" in norm(k) or "mala" in norm(k)},["mc","mala"])
    roster_df=choose({k:v for k,v in ss.items() if "maestro" in norm(k) or "personal" in norm(k)},["maestro_personal","personal","maestro"])
    base=parse_base(base_df)
    roster=parse_roster(roster_df)
    base=apply_roster(base,roster)
    fnr=parse_inc(fnr_df,"FNR")
    mc=parse_inc(mc_df,"MC")
except Exception as e:
    st.error(f"Error en el Excel consolidado: {e}"); st.stop()

if excluded:
    fnr=fnr[~fnr.ORDER_NUMBER.isin(excluded)].copy()
    mc=mc[~mc.ORDER_NUMBER.isin(excluded)].copy()
fnr=attach(fnr,base); mc=attach(mc,base); s=summary(base,fnr,mc)
store = load_store()
# Registrar automáticamente pickers vistos en la operación, sin alterar su estado histórico.
for _p in base["PICKER"].astype(str).str.strip().unique(): picker_record(store, _p)
save_store(store)

if roster is not None:
    matched=base["PICKER"].isin(roster["PICKER"]).sum()
    total=len(base); unmatched=total-matched
    with st.sidebar:
        st.success(f"Personal cruzado: {matched}/{total} pickers desde el Excel consolidado.")
        if unmatched: st.warning(f"{unmatched} pickers de la base no aparecen en Maestro_Personal.")

with st.sidebar:
    st.subheader("Filtros de personal")
    pickers=["Todos"]+sorted({str(x) for x in s["PICKER"].tolist() if str(x).strip()})
    turns=["Todos"]+sorted({str(x) for x in s["TURNO"].tolist() if str(x).strip()})
    sups=["Todos"]+sorted({str(x) for x in s["SUPERVISOR"].tolist() if str(x).strip() and str(x)!="No asignado"})
    areas=["Todos"]+sorted({str(x) for x in s["AREA_BASE"].tolist() if str(x).strip() and str(x)!="No especificada"})
    sp=st.selectbox("Picker",pickers)
    stn=st.selectbox("Turno",turns)
    ssup=st.selectbox("Supervisor",sups)
    sar=st.selectbox("Área",areas)

# Filtros combinados de personal. Los KPI se recalculan sobre la selección.
s_view=s.copy()
if sp!="Todos": s_view=s_view[s_view.PICKER==sp]
if stn!="Todos": s_view=s_view[s_view.TURNO==stn]
if ssup!="Todos": s_view=s_view[s_view.SUPERVISOR==ssup]
if sar!="Todos": s_view=s_view[s_view.AREA_BASE==sar]
base_view=base[base.PICKER.isin(s_view.PICKER)]
fnr_view=fnr[fnr.PICKER.isin(s_view.PICKER)]
mc_view=mc[mc.PICKER.isin(s_view.PICKER)]

a,b,c,d,e,g,h,f=st.tabs(["🏠 Bodega","👤 Picker","🏷️ Artículos","📦 Pedidos","🌙 Turnos / Áreas","👥 Supervisores","🛡️ Seguimiento","📥 Exportar"])

with a:
    lines=base_view.LINEAS.sum(); F=fnr_view.INCIDENCIAS.sum(); M=mc_view.INCIDENCIAS.sum()
    total_pedidos,fnr_pedidos,mc_pedidos,fnr_rate,mc_rate=monthly_kpis(base_view,fnr_view,mc_view)
    st.subheader(f"Resumen de bodega — {periodo}")
    q=st.columns(6)
    q[0].metric("Líneas",f"{lines:,.0f}")
    q[1].metric("Pedidos",f"{total_pedidos:,.0f}")
    q[2].metric("FNR mensual",f"{fnr_rate:.2f}%" if fnr_rate is not None else "N/D", "Objetivo < 1.50%")
    q[3].metric("MC mensual",f"{mc_rate:.2f}%" if mc_rate is not None else "N/D", "Objetivo < 1.00%")
    q[4].metric("Pickers",len(s_view))
    q[5].metric("Fuera objetivo",(s_view.ESTADO=="🔴 FUERA DE OBJETIVO").sum())
    st.caption(f"KPI mensual: {fnr_pedidos:,} pedidos con FNR / {total_pedidos:,} pedidos = {fnr_rate:.2f}% · {mc_pedidos:,} pedidos con MC / {total_pedidos:,} pedidos = {mc_rate:.2f}%" if fnr_rate is not None and mc_rate is not None else "No hay suficientes pedidos para calcular el KPI mensual.")
    st.divider()
    st.subheader("Indicador operativo por líneas")
    k=st.columns(2); k[0].metric("FNR / líneas",f"{F/lines*100:.2f}%" if lines else "N/D"); k[1].metric("MC / líneas",f"{M/lines*100:.2f}%" if lines else "N/D")
    st.dataframe(s_view,use_container_width=True,hide_index=True)

with b:
    if sp=="Todos": st.info("Selecciona un picker.")
    else:
        r=s[s.PICKER==sp].iloc[0]
        q=st.columns(4); q[0].metric("Líneas",f"{r.LINEAS:,.0f}"); q[1].metric("FNR",f"{r.FNR:,.0f}",f"{r['FNR_%']:.2f}%")
        st.caption(f"Turno: {r.TURNO} · Correo/usuario: {r.CORREO} · Supervisor: {r.SUPERVISOR} · Área: {r.AREA_BASE}")
        q[2].metric("MC",f"{r.MC:,.0f}",f"{r['MC_%']:.2f}%"); q[3].metric("Estado",r.ESTADO)
        st.subheader("Artículos FNR"); st.dataframe(products(fnr,sp),use_container_width=True,hide_index=True)
        st.subheader("Artículos MC"); st.dataframe(products(mc,sp),use_container_width=True,hide_index=True)
        st.subheader("Pedidos FNR"); st.dataframe(orders(fnr,sp),use_container_width=True,hide_index=True)
        rec = picker_record(store, sp)
        st.subheader("📝 Comentarios del picker")
        st.caption("Los comentarios quedan guardados para que cualquier supervisor pueda consultar antecedentes del picker.")
        nuevo_com = st.text_area("Nuevo comentario", key=f"comentario_picker_{sp}", placeholder="Escribe una observación relevante del picker...")
        comentario_sup = st.text_input("Supervisor que registra", key=f"comentario_sup_{sp}")
        if st.button("Guardar comentario", key=f"guardar_com_{sp}"):
            if nuevo_com.strip():
                rec["comentarios"].append({"fecha":datetime.now().strftime("%Y-%m-%d %H:%M"),"supervisor":comentario_sup.strip() or "No especificado","texto":nuevo_com.strip()})
                save_store(store); st.success("Comentario guardado."); st.rerun()
        if rec["comentarios"]:
            st.dataframe(pd.DataFrame(rec["comentarios"]),use_container_width=True,hide_index=True)
        else:
            st.info("Sin comentarios registrados.")

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
    if roster is not None:
        st.success("Los turnos, correos, supervisores y áreas provienen del Maestro_Personal del único Excel cargado.")
    else:
        st.info("El Maestro_Personal del Excel consolidado asigna turno, correo, supervisor y área a cada picker.")
    st.subheader("FNR por turno"); st.dataframe(groups(fnr_view,base_view,"TURNO"),use_container_width=True,hide_index=True)
    st.subheader("MC por turno"); st.dataframe(groups(mc_view,base_view,"TURNO"),use_container_width=True,hide_index=True)
    st.subheader("FNR por área"); st.dataframe(groups(fnr_view,base_view,"AREA"),use_container_width=True,hide_index=True)
    st.subheader("MC por área"); st.dataframe(groups(mc_view,base_view,"AREA"),use_container_width=True,hide_index=True)
    st.warning("El % / líneas por área solo aparece si existe un denominador real de líneas por área.")

with g:
    st.subheader("Retroalimentación por turno")
    if roster is None or roster.empty:
        st.info("Elige el turno directamente desde el filtro del Excel consolidado.")
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

with h:
    st.subheader("🛡️ Seguimiento de Pickers")
    st.caption("La baja saca al picker del seguimiento activo, pero conserva todo su historial. Sus incidencias siguen contando en las estadísticas del periodo en que ocurrieron.")
    # Lista de seguimiento se separa de las estadísticas: s sigue conteniendo todos los pickers del periodo.
    all_names = sorted(set(base["PICKER"].astype(str).str.strip()) | set(store.get("pickers",{}).keys()))
    vista = st.radio("Mostrar", ["Activos", "Dados de baja", "Todos"], horizontal=True)
    rows=[]
    for p in all_names:
        rec=picker_record(store,p)
        acciones=rec.get("acciones",[])
        counts={a:sum(1 for x in acciones if x.get("accion")==a) for a in ["Advertencia verbal 1","Advertencia verbal 2","Acta 1","Acta 2","Acta 3","Cero tolerancia"]}
        rows.append({"PICKER":p,"ESTADO":rec.get("estado","ACTIVO"),**counts,"ÚLTIMA ACCIÓN":acciones[-1].get("fecha","") if acciones else ""})
    track=pd.DataFrame(rows)
    if vista=="Activos": track=track[track.ESTADO=="ACTIVO"]
    elif vista=="Dados de baja": track=track[track.ESTADO=="BAJA"]
    st.dataframe(track,use_container_width=True,hide_index=True)
    st.divider()
    opciones = track["PICKER"].tolist() if not track.empty else all_names
    if opciones:
        picker_track=st.selectbox("Picker", opciones, key="picker_tracking")
        rec=picker_record(store,picker_track)
        c1,c2,c3=st.columns(3)
        c1.metric("Estado",rec.get("estado","ACTIVO"))
        c2.metric("Comentarios",len(rec.get("comentarios",[])))
        c3.metric("Acciones",len(rec.get("acciones",[])))
        st.subheader("Registrar acción")
        accion=st.selectbox("Tipo de seguimiento",["Advertencia verbal 1","Advertencia verbal 2","Acta 1","Acta 2","Acta 3","Cero tolerancia"],key="accion_tipo")
        sup_acc=st.text_input("Supervisor",key="accion_supervisor")
        motivo=st.text_area("Motivo / detalle",key="accion_motivo")
        if st.button("Guardar acción",key="guardar_accion"):
            rec["acciones"].append({"fecha":datetime.now().strftime("%Y-%m-%d %H:%M"),"accion":accion,"supervisor":sup_acc.strip() or "No especificado","motivo":motivo.strip()})
            save_store(store); st.success("Acción registrada."); st.rerun()
        if rec.get("acciones"):
            st.subheader("Historial de acciones")
            st.dataframe(pd.DataFrame(rec["acciones"]),use_container_width=True,hide_index=True)
        st.subheader("Estado del picker")
        if rec.get("estado","ACTIVO")=="ACTIVO":
            st.warning("Dar de baja NO borra al picker ni sus estadísticas históricas. Solo lo retira del seguimiento activo.")
            if st.button("⛔ Dar de baja de la base de seguimiento",key="baja_picker"):
                rec["estado"]="BAJA"; rec["baja_fecha"]=datetime.now().strftime("%Y-%m-%d %H:%M")
                save_store(store); st.success("Picker dado de baja del seguimiento activo. Su histórico permanece."); st.rerun()
        else:
            if st.button("↩️ Reactivar en seguimiento",key="reactivar_picker"):
                rec["estado"]="ACTIVO"; rec["reactivacion_fecha"]=datetime.now().strftime("%Y-%m-%d %H:%M")
                save_store(store); st.success("Picker reactivado."); st.rerun()
        st.info("Las estadísticas de FNR/MC de este periodo NO se filtran por estado. Un picker dado de baja sigue contando porque sus incidencias ocurrieron durante su periodo de operación.")

with f:
    st.download_button("📥 Descargar Excel completo",export(s,fnr,mc,None if sp=="Todos" else sp,roster),f"Analisis_FNR_MC_{periodo}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.info("La V4 permite cargar personal por cuatro turnos y capturar retroalimentación por supervisor.")
