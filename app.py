import io, re, json, os
from difflib import SequenceMatcher
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Control FNR & MC", page_icon="📊", layout="wide")
FNR_OBJ, MC_OBJ = 1.50, 1.00
STORE_FILE = "picker_seguimiento.json"
PERSIST_DIR = "app_data"
PERSIST_FILES = {
    "base_picker": os.path.join(PERSIST_DIR, "base_picker.xlsx"),
    "detalle_fnr": os.path.join(PERSIST_DIR, "detalle_fnr.xlsx"),
    "detalle_mc": os.path.join(PERSIST_DIR, "detalle_mc.xlsx"),
    "plantilla_personal": os.path.join(PERSIST_DIR, "master_pickers.xlsx"),
}

def ensure_persist_dir():
    os.makedirs(PERSIST_DIR, exist_ok=True)

def persist_upload(upload, key):
    """Guarda físicamente el Excel cargado para que sobreviva a un rerun/refresh."""
    if upload is None:
        return None
    ensure_persist_dir()
    path=PERSIST_FILES[key]
    with open(path, "wb") as f:
        f.write(upload.getvalue())
    return io.BytesIO(upload.getvalue())

def load_persisted_upload(key):
    """Recupera el último Excel guardado cuando el uploader está vacío."""
    path=PERSIST_FILES[key]
    if not os.path.exists(path):
        return None
    try:
        with open(path, "rb") as f:
            data=f.read()
        if not data:
            return None
        return io.BytesIO(data)
    except Exception:
        return None

def persisted_status():
    return {k: os.path.exists(v) and os.path.getsize(v) > 0 for k,v in PERSIST_FILES.items()}

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

def person_key(x):
    """Clave robusta para nombres: sin acentos, signos ni diferencias de espacios."""
    return re.sub(r"[^a-z0-9]", "", norm(x))

def name_tokens(x):
    """Tokens de identidad de una persona."""
    return {t for t in norm(x).split("_") if t and len(t) > 1}

def token_key(x):
    return "_".join(sorted(name_tokens(x)))

def extract_code(value):
    """Extrae el código del formato 1188502.nombre@justo.mx o JT1180680.nombre@..."""
    s=str(value or "").strip().lower()
    if not s or s in {"nan", "none"}: return ""
    m=re.match(r"^(jt)?(\d+)(?:\.|\s|$)", s)
    return (m.group(2) if m else "")

def email_name_tokens(value):
    """Obtiene los nombres del correo ignorando el código inicial."""
    s=str(value or "").strip().lower()
    if "@" in s: s=s.split("@",1)[0]
    s=re.sub(r"^(?:jt)?\d+[._-]*", "", s)
    return name_tokens(s.replace(".", "_"))

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
    p=col(df,["picker","picker_nombre","nombre","email_picker","persona"])
    l=col(df,["total_lineas","lineas_totales","lineas","total_lineas_pickeadas"])
    ped=col(df,["total_pedidos","pedidos_totales","pedidos"])
    sh=col(df,["turno","shift"]); ar=col(df,["area","departamento","department"])
    em=col(df,["codigo_correo","codigo + correo","correo","email","usuario"])
    if not p or not l: raise ValueError("La base necesita PICKER y Total lineas.")
    x=pd.DataFrame({
        "PICKER":df[p].astype(str).str.strip(),
        "LINEAS":pd.to_numeric(df[l],errors="coerce").fillna(0),
        "PEDIDOS":pd.to_numeric(df[ped],errors="coerce").fillna(0) if ped else 0,
        "TURNO":df[sh].astype(str).str.strip() if sh else "No especificado",
        "AREA_BASE":df[ar].astype(str).str.strip() if ar else "No especificada",
        "CORREO":df[em].astype(str).str.strip() if em else ""
    })
    x["_KEY"]=x["PICKER"].map(person_key)
    x["_TOKEN_KEY"]=x["PICKER"].map(token_key)
    x["_CODE_KEY"]=x["CORREO"].map(extract_code)
    x["_EMAIL_TOKENS"]=x["CORREO"].map(email_name_tokens)
    return x.sort_values("LINEAS").drop_duplicates("PICKER",keep="last")

def parse_roster(df, turno_fijo=None):
    """Maestro consolidado: PICKER, TURNO, CODIGO + CORREO, SUPERVISOR y AREA_BASE."""
    p=col(df,["picker","picker_nombre","email_picker","nombre","persona"])
    sh=col(df,["turno","shift"])
    em=col(df,["codigo_correo","codigo + correo","correo","email","usuario"])
    sup=col(df,["supervisor","supervisor_nombre","jefe","responsable"])
    ar=col(df,["area_base","area","departamento","department"])
    if not p: raise ValueError("La hoja de personal necesita la columna PICKER.")
    x=pd.DataFrame({
        "PICKER":df[p].astype(str).str.strip(),
        "TURNO_MAESTRO":(df[sh].astype(str).str.strip() if sh and not turno_fijo else turno_fijo or ""),
        "CORREO":df[em].astype(str).str.strip() if em else "",
        "SUPERVISOR":df[sup].astype(str).str.strip() if sup else "No asignado",
        "AREA_MAESTRO":df[ar].astype(str).str.strip() if ar else ""
    })
    x["TURNO_MAESTRO"]=x["TURNO_MAESTRO"].fillna("").astype(str).str.strip()
    x=x[x["PICKER"].str.strip().ne("")].copy()
    x["_KEY"]=x["PICKER"].map(person_key)
    x["_TOKEN_KEY"]=x["PICKER"].map(token_key)
    x["_CODE_KEY"]=x["CORREO"].map(extract_code)
    x["_EMAIL_TOKENS"]=x["CORREO"].map(email_name_tokens)
    x=x.drop_duplicates("_KEY",keep="last")
    return x

def _match_master_row(value, roster, code_value=""):
    """Resuelve una persona contra el Master.

    Prioridad de identidad:
    1) nombre exacto normalizado;
    2) mismas palabras aunque cambie el orden;
    3) código de empleado;
    4) nombre parcial/subconjunto (ej. "AMARO PAULINA ANDREA" contra
       "AMARO OROPEZA PAULINA ANDREA");
    5) nombre del correo;
    6) similitud conservadora.

    Nunca asigna por una sola palabra genérica.
    """
    if roster is None or roster.empty: return None
    value=str(value or "").strip()
    key=person_key(value); toks=name_tokens(value)
    code=extract_code(code_value) or extract_code(value)
    email_toks=email_name_tokens(code_value)

    # 1. Nombre exacto
    hits=roster[roster["_KEY"].astype(str)==key] if key else roster.iloc[0:0]
    if len(hits)==1: return hits.iloc[0]

    # 2. Todas las palabras coinciden, sin importar orden
    if toks:
        hits=roster[roster["_TOKEN_KEY"].astype(str)==token_key(value)]
        if len(hits)==1: return hits.iloc[0]

    # 3. Código único
    if code:
        hits=roster[roster["_CODE_KEY"].astype(str)==code]
        if len(hits)==1: return hits.iloc[0]

    # 4/5. Puntaje por identidad: favorece que TODOS los tokens del nombre
    # corto estén contenidos en el nombre del Master, o que el correo aporte
    # nombre+apellido.
    candidates=[]
    for _,rr in roster.iterrows():
        rtoks=name_tokens(rr.get("PICKER",""))
        if not rtoks: continue
        overlap=len(toks & rtoks)
        union=len(toks | rtoks) or 1
        # Cobertura del nombre de entrada y del Master.
        cov_in=overlap/(len(toks) or 1)
        cov_master=overlap/(len(rtoks) or 1)
        j=overlap/union
        seq=SequenceMatcher(None,key,str(rr.get("_KEY","") or "")).ratio()
        etoks=rr.get("_EMAIL_TOKENS",set())
        email_overlap=len(email_toks & etoks) if email_toks else 0
        email_cov=email_overlap/(len(email_toks) or 1) if email_toks else 0
        # Código ya fue evaluado; aquí el correo ayuda cuando el Excel
        # operativo trae nombre y el Master tiene nombre+correo.
        score=(0.50*cov_in)+(0.18*j)+(0.17*seq)+(0.15*email_cov)
        candidates.append((score,cov_in,j,seq,email_cov,rr))

    candidates.sort(key=lambda z:z[0], reverse=True)
    if not candidates: return None
    best=candidates[0]; second=candidates[1][0] if len(candidates)>1 else 0
    score,cov_in,j,seq,email_cov,rr=best

    # Caso fuerte: al menos 2 palabras del nombre de entrada están en el Master
    # y todas las palabras de entrada están cubiertas por ese Master.
    if len(toks)>=2 and cov_in>=1.0 and len(toks & name_tokens(rr["PICKER"]))>=2:
        if score>=0.60 and (score-second>=0.05 or score>=0.82): return rr
    # Si el correo identifica al menos dos tokens, es una señal fuerte.
    if len(email_toks)>=2 and email_cov>=1.0 and len(email_toks & name_tokens(rr["PICKER"]))>=2:
        if score>=0.58 and (score-second>=0.04 or score>=0.82): return rr
    # Coincidencia muy cercana para errores de escritura.
    if score>=0.90 and (score-second>=0.04 or score>=0.96): return rr
    return None

def apply_roster(base, roster):
    if roster is None or roster.empty: return base
    x=base.copy()
    matches=[]
    for _,row in x.iterrows():
        rr=_match_master_row(row.get("PICKER",""),roster,row.get("CORREO",""))
        matches.append(rr)

    out=[]
    for row,rr in zip(x.to_dict("records"),matches):
        if rr is None:
            row["_MASTER_MATCH"]=False
            row["_MASTER_PICKER"]=""
            row["_MASTER_TURNO"]=""
            row["_MASTER_CORREO"]=""
            row["_MASTER_SUPERVISOR"]="No asignado"
            row["_MASTER_AREA"]=""
        else:
            row["_MASTER_MATCH"]=True
            row["_MASTER_PICKER"]=str(rr.get("PICKER","")).strip()
            row["_MASTER_TURNO"]=str(rr.get("TURNO_MAESTRO","")).strip()
            row["_MASTER_CORREO"]=str(rr.get("CORREO","")).strip()
            row["_MASTER_SUPERVISOR"]=str(rr.get("SUPERVISOR","No asignado")).strip()
            row["_MASTER_AREA"]=str(rr.get("AREA_MAESTRO","")).strip()
        out.append(row)
    x=pd.DataFrame(out)

    ok=x["_MASTER_MATCH"]
    x.loc[ok,"PICKER"]=x.loc[ok,"_MASTER_PICKER"]
    x.loc[ok,"TURNO"]=x.loc[ok,"_MASTER_TURNO"]
    x.loc[ok,"CORREO"]=x.loc[ok,"_MASTER_CORREO"]
    x.loc[ok,"SUPERVISOR"]=x.loc[ok,"_MASTER_SUPERVISOR"]
    x.loc[ok,"AREA_BASE"]=x.loc[ok,"_MASTER_AREA"]
    return x.drop(columns=["_KEY","_TOKEN_KEY","_CODE_KEY","_EMAIL_TOKENS","_MASTER_PICKER","_MASTER_TURNO","_MASTER_CORREO","_MASTER_SUPERVISOR","_MASTER_AREA"],errors="ignore")

def apply_manual_personnel(base, store):
    """Aplica asignaciones/exclusiones capturadas desde la interfaz.

    Las asignaciones se guardan por una clave normalizada del picker original,
    de modo que no es necesario editar el Excel operativo. Una exclusión solo
    retira a la persona de los filtros/segmentaciones de personal; no borra
    sus líneas ni sus incidencias del cálculo global.
    """
    x=base.copy()
    overrides=store.get("master_overrides",{}) if isinstance(store,dict) else {}
    excluded=set(store.get("master_excluded",[])) if isinstance(store,dict) else set()
    if "_MANUAL_MATCH" not in x.columns: x["_MANUAL_MATCH"]=False
    if "_EXCLUDED_PERSONNEL" not in x.columns: x["_EXCLUDED_PERSONNEL"]=False
    for i,row in x.iterrows():
        key=person_key(row.get("PICKER",""))
        if not key: continue
        if key in excluded:
            x.at[i,"_EXCLUDED_PERSONNEL"]=True
            x.at[i,"_MASTER_MATCH"]=True
            x.at[i,"TURNO"]="__EXCLUIDO__"
            x.at[i,"SUPERVISOR"]="Excluido"
            x.at[i,"AREA_BASE"]="Excluido"
            continue
        ov=overrides.get(key)
        if not ov: continue
        x.at[i,"_MANUAL_MATCH"]=True
        x.at[i,"_MASTER_MATCH"]=True
        if ov.get("PICKER"): x.at[i,"PICKER"]=str(ov.get("PICKER")).strip()
        x.at[i,"TURNO"]=str(ov.get("TURNO","")).strip() or "No especificado"
        x.at[i,"CORREO"]=str(ov.get("CORREO","")).strip()
        x.at[i,"SUPERVISOR"]=str(ov.get("SUPERVISOR","")).strip() or "No asignado"
        x.at[i,"AREA_BASE"]=str(ov.get("AREA_BASE","")).strip() or "No especificada"
    return x

def effective_roster(roster, store):
    """Construye el maestro efectivo: Excel + asignaciones manuales - exclusiones."""
    r=roster.copy() if roster is not None else pd.DataFrame()
    overrides=store.get("master_overrides",{}) if isinstance(store,dict) else {}
    excluded=set(store.get("master_excluded",[])) if isinstance(store,dict) else set()
    if overrides:
        rows=[]
        for _,ov in overrides.items():
            rows.append({
                "PICKER":str(ov.get("PICKER","")).strip(),
                "TURNO_MAESTRO":str(ov.get("TURNO","")).strip(),
                "CORREO":str(ov.get("CORREO","")).strip(),
                "SUPERVISOR":str(ov.get("SUPERVISOR","No asignado")).strip() or "No asignado",
                "AREA_MAESTRO":str(ov.get("AREA_BASE","")).strip(),
            })
        if rows:
            manual=pd.DataFrame(rows)
            manual["_KEY"]=manual["PICKER"].map(person_key)
            manual["_TOKEN_KEY"]=manual["PICKER"].map(token_key)
            manual["_CODE_KEY"]=manual["CORREO"].map(extract_code)
            manual["_EMAIL_TOKENS"]=manual["CORREO"].map(email_name_tokens)
            r=pd.concat([r,manual],ignore_index=True)
    if not r.empty:
        r["_KEY"]=r["PICKER"].map(person_key)
        r=r[~r["_KEY"].isin(excluded)].copy()
        r=r.drop_duplicates("_KEY",keep="last")
    return r

def canonicalize_incidents(inc, base):
    """Lleva FNR/MC al nombre canónico de la base, usando nombre/código/correo."""
    if inc is None or inc.empty or base is None or base.empty: return inc
    out=inc.copy()
    records=base.to_dict("records")
    # Índices directos
    by_key={person_key(r.get("PICKER","")):r for r in records if person_key(r.get("PICKER",""))}
    by_token={token_key(r.get("PICKER","")):r for r in records if token_key(r.get("PICKER",""))}
    by_code={extract_code(r.get("CORREO","")):r for r in records if extract_code(r.get("CORREO",""))}

    result=[]
    for _,row in out.iterrows():
        value=str(row.get("PICKER","")).strip()
        code=extract_code(row.get("CORREO","") or value)
        key=person_key(value); tk=token_key(value)
        rr=by_key.get(key) or by_token.get(tk) or (by_code.get(code) if code else None)
        if rr is not None:
            result.append(str(rr.get("PICKER",value)).strip())
            continue
        # Resolver contra la lista de la base con la misma lógica de identidad.
        toks=name_tokens(value); best=None; best_tuple=None
        for cand in records:
            ctoks=name_tokens(cand.get("PICKER","")); overlap=len(toks & ctoks)
            cov=overlap/(len(toks) or 1); j=overlap/(len(toks|ctoks) or 1)
            seq=SequenceMatcher(None,person_key(value),person_key(cand.get("PICKER",""))).ratio()
            score=.55*cov+.20*j+.25*seq
            tup=(score,cov,seq,cand)
            if best_tuple is None or tup[:3]>best_tuple[:3]: best_tuple=tup; best=cand
        if best_tuple and len(toks)>=2 and best_tuple[1]>=1 and best_tuple[0]>=.70:
            result.append(str(best.get("PICKER",value)).strip())
        else:
            result.append(value)
    out["PICKER"]=result
    return out

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
        "CORREO":df[col(df,["codigo_correo","codigo + correo","correo","email","usuario"])].astype(str).str.strip() if col(df,["codigo_correo","codigo + correo","correo","email","usuario"]) else "",
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
    ref=base[[c for c in ["PICKER","LINEAS","PEDIDOS","TURNO","CORREO","_EXCLUDED_PERSONNEL"] if c in base.columns]].copy()
    ref["_KEY"]=ref["PICKER"].map(person_key)
    ref["_TOKEN_KEY"]=ref["PICKER"].map(token_key)
    ref["_CODE_KEY"]=ref.get("CORREO",pd.Series([""]*len(ref),index=ref.index)).map(extract_code)
    y=inc.copy()
    y["_KEY"]=y["PICKER"].map(person_key)
    y["_TOKEN_KEY"]=y["PICKER"].map(token_key)
    y["_CODE_KEY"]=y.get("CORREO",pd.Series([""]*len(y),index=y.index)).map(extract_code)
    ref_exact=ref.drop_duplicates("_KEY").set_index("_KEY")
    ref_token=ref.drop_duplicates("_TOKEN_KEY").set_index("_TOKEN_KEY")
    ref_code=ref[ref["_CODE_KEY"].astype(str).str.strip().ne("")].drop_duplicates("_CODE_KEY").set_index("_CODE_KEY")
    def get_ref(row):
        for key,mp in [(row["_KEY"],ref_exact),(row["_TOKEN_KEY"],ref_token),(row["_CODE_KEY"],ref_code)]:
            if key and key in mp.index:
                rr=mp.loc[key]
                if isinstance(rr,pd.DataFrame): rr=rr.iloc[0]
                return pd.Series([rr["PICKER"],rr["LINEAS"],rr["PEDIDOS"],rr["TURNO"],bool(rr.get("_EXCLUDED_PERSONNEL",False))])
        return pd.Series([row["PICKER"],0,0,"No especificado"])
    vals=y.apply(get_ref,axis=1)
    vals.columns=["PICKER_REF","LINEAS","PEDIDOS","TURNO_REF","_EXCLUDED_PERSONNEL"]
    y=pd.concat([y.reset_index(drop=True),vals.reset_index(drop=True)],axis=1)
    y.loc[y["PICKER_REF"].astype(str).str.strip().ne(""),"PICKER"]=y.loc[y["PICKER_REF"].astype(str).str.strip().ne(""),"PICKER_REF"]
    return y.drop(columns=["_KEY","_TOKEN_KEY","_CODE_KEY"],errors="ignore")

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
    """Agrupa FNR/MC por turno o área usando acceso por nombre de columna.
    Evita el acceso por atributo de pandas (p.ej. .LINEAS), que puede fallar
    según la versión de pandas/Streamlit Cloud.
    """
    x=inc.copy()
    if "_EXCLUDED_PERSONNEL" in x.columns:
        x=x[~x["_EXCLUDED_PERSONNEL"].fillna(False)].copy()
    if x.empty:
        return pd.DataFrame(columns=["GRUPO","INCIDENCIAS","% DEL TOTAL","LINEAS","% / LINEAS"])

    if key=="TURNO":
        x["GRUPO"]=x["TURNO_REF"].astype(str).str.strip()
        den=(base.groupby("TURNO",as_index=False)["LINEAS"].sum()
             .rename(columns={"TURNO":"GRUPO"}))
    else:
        x["GRUPO"]=x["AREA"].astype(str).str.strip()
        if "AREA_BASE" in base.columns and (base["AREA_BASE"].astype(str).str.strip()!="No especificada").any():
            den=(base.groupby("AREA_BASE",as_index=False)["LINEAS"].sum()
                 .rename(columns={"AREA_BASE":"GRUPO"}))
        else:
            den=None

    g=x.groupby("GRUPO",as_index=False)["INCIDENCIAS"].sum()
    total_inc=float(g["INCIDENCIAS"].sum())
    g["% DEL TOTAL"]=(g["INCIDENCIAS"]/total_inc*100).round(2) if total_inc else 0.0

    if den is not None:
        g=g.merge(den,on="GRUPO",how="left")
        g["LINEAS"]=pd.to_numeric(g["LINEAS"],errors="coerce").fillna(0.0)
        den_lineas=g["LINEAS"].where(g["LINEAS"].ne(0), float("nan"))
        g["% / LINEAS"]=pd.to_numeric(g["INCIDENCIAS"].div(den_lineas)*100,errors="coerce").round(2)
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

# Estado persistente: se carga antes de construir los widgets.
store=load_store()
store.setdefault("excluded_orders", [])
store.setdefault("feedback_rows", [])
store.setdefault("master_overrides", {})
store.setdefault("master_excluded", [])

with st.sidebar:
    st.header("Carga")
    ub_upload=st.file_uploader("① Base de Pickers / Líneas",type=["xlsx","xls"],key="base_picker")
    uf_upload=st.file_uploader("② Detalle FNR",type=["xlsx","xls"],key="detalle_fnr")
    um_upload=st.file_uploader("③ Detalle Mala Calidad",type=["xlsx","xls"],key="detalle_mc")
    up_upload=st.file_uploader("④ Plantilla consolidada de personal",type=["xlsx","xls"],key="plantilla_personal")

    # Cada archivo nuevo reemplaza automáticamente al guardado. Si solo se recarga
    # la página, la app recupera la última versión guardada sin pedir volver a subirla.
    ub=persist_upload(ub_upload,"base_picker") if ub_upload else load_persisted_upload("base_picker")
    uf=persist_upload(uf_upload,"detalle_fnr") if uf_upload else load_persisted_upload("detalle_fnr")
    um=persist_upload(um_upload,"detalle_mc") if um_upload else load_persisted_upload("detalle_mc")
    up=persist_upload(up_upload,"plantilla_personal") if up_upload else load_persisted_upload("plantilla_personal")

    status=persisted_status()
    labels={"base_picker":"Pickers/Líneas","detalle_fnr":"FNR","detalle_mc":"Mala Calidad","plantilla_personal":"Master Pickers"}
    guardados=[labels[k] for k,v in status.items() if v]
    if guardados:
        st.success("Archivos guardados: " + ", ".join(guardados))
    st.caption("Los 3 archivos operativos siguen separados. El Excel maestro concentra PICKER, TURNO, CODIGO + CORREO, SUPERVISOR y AREA_BASE. Los archivos quedan guardados para los siguientes recargados de la página.")
    st.divider()
    periodo=st.text_input("Periodo",value=str(store.get("periodo",datetime.now().strftime("%Y-%m"))),key="periodo_persistente")
    saved_excluded="\n".join(str(x) for x in store.get("excluded_orders",[]) if str(x).strip())
    ex=st.text_area("Pedidos operativos a excluir (uno por línea)",value=saved_excluded,key="pedidos_excluidos_persistentes")
    excluded={x.strip() for x in ex.splitlines() if x.strip()}
    # Guardar automáticamente preferencias y exclusiones.
    store["periodo"]=periodo.strip() or datetime.now().strftime("%Y-%m")
    store["excluded_orders"]=sorted(excluded)
    save_store(store)

if not (ub and uf and um and up):
    st.info("Carga los 3 Excel operativos y la plantilla consolidada de personal para comenzar.")
    st.markdown("**Archivos:** ① Pickers/Líneas · ② FNR · ③ Mala Calidad · ④ Master Pickers (PICKER, TURNO, CODIGO + CORREO, SUPERVISOR, AREA_BASE).")
    st.stop()

try:
    base=parse_base(choose(sheets(ub),["picker","lineas","resumen"]))
    fnr=parse_inc(choose(sheets(uf),["fnr","detalle"]),"FNR")
    mc=parse_inc(choose(sheets(um),["mc","mala"]),"MC")
    # La plantilla maestra actual tiene una sola hoja: Base_Pickers.
    # Columnas: PICKER, TURNO, CODIGO + CORREO, SUPERVISOR, AREA_BASE.
    master_sheets=sheets(up)
    if "Base_Pickers" in master_sheets:
        master_df=master_sheets["Base_Pickers"]
    else:
        master_df=choose(master_sheets,["base_pickers","maestro_personal","personal","maestro","turno"])
    roster=parse_roster(master_df)
    if roster.empty:
        raise ValueError("El Excel maestro no contiene pickers válidos.")
    base=apply_roster(base,roster)
    base=apply_manual_personnel(base,store)
    roster=effective_roster(roster,store)
    # Normaliza FNR/MC contra los pickers canónicos de la base.
    fnr=canonicalize_incidents(fnr,base)
    mc=canonicalize_incidents(mc,base)
except Exception as e:
    st.error(f"Error en los archivos cargados: {e}"); st.stop()

if excluded:
    fnr=fnr[~fnr.ORDER_NUMBER.isin(excluded)].copy()
    mc=mc[~mc.ORDER_NUMBER.isin(excluded)].copy()
fnr=attach(fnr,base); mc=attach(mc,base); s=summary(base,fnr,mc)
# Columna técnica de cruce: se usa para diagnóstico, no para mostrarla en el dashboard.
master_match=base.get("_MASTER_MATCH",pd.Series(False,index=base.index)).copy()
base_display=base.drop(columns=["_MASTER_MATCH"],errors="ignore")
s=s.drop(columns=["_MASTER_MATCH"],errors="ignore")
# Registrar automáticamente pickers vistos en la operación, sin alterar su estado histórico.
for _p in base["PICKER"].astype(str).str.strip().unique(): picker_record(store, _p)
save_store(store)

if roster is not None:
    match_series=base.get("_MASTER_MATCH",pd.Series(False,index=base.index)).fillna(False).astype(bool)
    excluded_series=base.get("_EXCLUDED_PERSONNEL",pd.Series(False,index=base.index)).fillna(False).astype(bool)
    matched=int((match_series & ~excluded_series).sum())
    total=len(base); unmatched=total-matched-int(excluded_series.sum())
    matched_turn=int((match_series & ~excluded_series & base["TURNO"].astype(str).str.strip().ne("") & base["TURNO"].astype(str).str.strip().ne("No especificado") & base["TURNO"].astype(str).str.strip().ne("__EXCLUIDO__")).sum())
    with st.sidebar:
        st.success(f"Personal identificado: {matched}/{total} pickers.")
        st.caption(f"Turnos asignados: {matched_turn}/{total} · Excluidos: {int(excluded_series.sum())}")
        if unmatched: st.warning(f"{unmatched} pickers necesitan asignación o exclusión.")
        with st.expander("🛠️ Resolver pickers sin coincidencia", expanded=bool(unmatched)):
            diag=base.loc[(~match_series) & (~excluded_series),["PICKER","CORREO"]].copy() if unmatched else pd.DataFrame()
            if diag.empty:
                st.success("Todos los pickers están identificados o fueron excluidos.")
            else:
                unresolved=diag["PICKER"].astype(str).tolist()
                selected_unresolved=st.selectbox("Picker pendiente", unresolved, key="resolver_picker")
                current_key=person_key(selected_unresolved)
                st.caption("Puedes resolverlo aquí sin modificar el Excel operativo.")
                ov=store.get("master_overrides",{}).get(current_key,{})
                turn_options=sorted(set(["Matutino","Intermedio","Tarde","Vespertino","Nocturno","Sin turno"] + [str(x).strip() for x in roster["TURNO_MAESTRO"].dropna().tolist() if str(x).strip()]))
                default_turn=ov.get("TURNO") if ov.get("TURNO") in turn_options else (turn_options[0] if turn_options else "Sin turno")
                turno_manual=st.selectbox("Turno",turn_options,index=turn_options.index(default_turn),key=f"manual_turn_{current_key}")
                correo_manual=st.text_input("Código + correo",value=str(ov.get("CORREO","") or ""),key=f"manual_email_{current_key}")
                sup_manual=st.text_input("Supervisor",value=str(ov.get("SUPERVISOR","") or ""),key=f"manual_sup_{current_key}")
                area_manual=st.text_input("Área",value=str(ov.get("AREA_BASE","") or ""),key=f"manual_area_{current_key}")
                c1,c2=st.columns(2)
                with c1:
                    if st.button("✅ Asignar y guardar",key=f"assign_{current_key}"):
                        store.setdefault("master_overrides",{})[current_key]={"PICKER":selected_unresolved,"TURNO":turno_manual,"CORREO":correo_manual.strip(),"SUPERVISOR":sup_manual.strip() or "No asignado","AREA_BASE":area_manual.strip() or "No especificada"}
                        store.setdefault("master_excluded",[])
                        store["master_excluded"]=[k for k in store["master_excluded"] if k!=current_key]
                        save_store(store); st.success("Asignación guardada."); st.rerun()
                with c2:
                    if st.button("🗑️ Excluir de personal",key=f"exclude_{current_key}"):
                        store.setdefault("master_excluded",[])
                        if current_key not in store["master_excluded"]: store["master_excluded"].append(current_key)
                        save_store(store); st.success("Picker excluido del maestro de personal. Sus líneas/FNR/MC globales se conservan."); st.rerun()
                st.dataframe(diag.head(100),use_container_width=True,hide_index=True)
            # Administración de asignaciones ya hechas
            saved=store.get("master_overrides",{})
            if saved:
                st.caption(f"Asignaciones manuales guardadas: {len(saved)}")
            excluded_saved=store.get("master_excluded",[])
            if excluded_saved:
                st.caption(f"Exclusiones guardadas: {len(excluded_saved)}")
                excluded_labels=[next((str(v) for v in base["PICKER"].tolist() if person_key(v)==k),k) for k in excluded_saved]
                restore_label=st.selectbox("Reactivar picker excluido", ["Selecciona..."]+excluded_labels, key="restore_excluded")
                if restore_label!="Selecciona..." and st.button("↩️ Reactivar en personal",key="restore_excluded_btn"):
                    rk=person_key(restore_label)
                    store["master_excluded"]=[k for k in store.get("master_excluded",[]) if k!=rk]
                    save_store(store); st.success("Picker reactivado para el análisis de personal."); st.rerun()

# Retira columnas técnicas antes de mostrar/exportar.
base=base.drop(columns=["_MASTER_MATCH"],errors="ignore")

# Filtros combinados de personal. Las personas excluidas no aparecen en los filtros,
# pero sus líneas/incidencias siguen disponibles para el KPI global.
s_view=s.copy()
if "_EXCLUDED_PERSONNEL" in s_view.columns:
    s_view=s_view[~s_view["_EXCLUDED_PERSONNEL"].fillna(False)]
with st.sidebar:
    st.subheader("Filtros de personal")
    pickers=["Todos"]+sorted({str(x) for x in s_view["PICKER"].tolist() if str(x).strip()})
    turns=["Todos"]+sorted({str(x) for x in s_view["TURNO"].tolist() if str(x).strip() and str(x)!="__EXCLUIDO__"})
    sups=["Todos"]+sorted({str(x) for x in s_view["SUPERVISOR"].tolist() if str(x).strip() and str(x)!="No asignado"})
    areas=["Todos"]+sorted({str(x) for x in s_view["AREA_BASE"].tolist() if str(x).strip() and str(x)!="No especificada"})
    sp=st.selectbox("Picker",pickers)
    stn=st.selectbox("Turno",turns)
    ssup=st.selectbox("Supervisor",sups)
    sar=st.selectbox("Área",areas)

# Filtros combinados de personal. Los KPI se recalculan sobre la selección.
if sp!="Todos": s_view=s_view[s_view.PICKER==sp]
if stn!="Todos": s_view=s_view[s_view.TURNO==stn]
if ssup!="Todos": s_view=s_view[s_view.SUPERVISOR==ssup]
if sar!="Todos": s_view=s_view[s_view.AREA_BASE==sar]
base_view=base[base.PICKER.isin(s_view.PICKER)]
fnr_view=fnr[fnr.PICKER.isin(s_view.PICKER)]
mc_view=mc[mc.PICKER.isin(s_view.PICKER)]

a,b,c,d,e,g,h,i,f=st.tabs(["🏠 Bodega","👤 Picker","🏷️ Artículos","📦 Pedidos","🌙 Turnos / Áreas","👥 Supervisores","🛡️ Seguimiento","🛠️ Resolver personal","📥 Exportar"])

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
        turnos=sorted({str(x).strip() for x in roster["TURNO_MAESTRO"].dropna() if str(x).strip()})
        turno_sel=st.selectbox("Turno",turnos if turnos else ["Sin turno"])
        rt=roster[roster["TURNO_MAESTRO"]==turno_sel].copy()
        supervisores=sorted([str(x) for x in rt["SUPERVISOR"].dropna().unique() if str(x).strip() and str(x)!="No asignado"])
        sup_sel=st.selectbox("Supervisor",["Todos"]+supervisores)
        if sup_sel!="Todos": rt=rt[rt["SUPERVISOR"]==sup_sel]
        disponibles=[str(x).strip() for x in rt["PICKER"].tolist() if str(x).strip() in set(s["PICKER"].astype(str).str.strip())]
        picker_sel=st.selectbox("Picker para retroalimentación",["Selecciona..."]+sorted(disponibles,key=lambda z:z.upper()))
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
                store.setdefault("feedback_rows",[]).append(row)
                save_store(store)
                st.success("Retroalimentación guardada permanentemente.")
                st.rerun()
        if store.get("feedback_rows"):
            st.subheader("Retroalimentaciones guardadas")
            fb=pd.DataFrame(store["feedback_rows"])
            st.dataframe(fb,use_container_width=True,hide_index=True)
            st.download_button("📥 Descargar retroalimentaciones",feedback_export(store["feedback_rows"]),f"Retroalimentacion_{periodo}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with h:
    st.subheader("🛡️ Seguimiento de Pickers")
    st.caption("La baja saca al picker del seguimiento activo, pero conserva todo su historial. Sus incidencias siguen contando en las estadísticas del periodo en que ocurrieron.")
    # Lista de seguimiento se separa de las estadísticas: s sigue conteniendo todos los pickers del periodo.
    all_names = sorted({str(x).strip() for x in base["PICKER"].tolist()} | {str(x).strip() for x in store.get("pickers",{}).keys() if str(x).strip()}, key=lambda z:z.upper())
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

with i:
    st.subheader("🛠️ Resolver personal pendiente")
    st.caption("Selecciona varios pickers sin coincidencia para asignarles el mismo turno, supervisor y área, o excluirlos del personal. Las líneas y los KPI históricos no se borran.")

    pending_mask=(~match_series) & (~excluded_series) if "match_series" in globals() else pd.Series(False,index=base.index)
    pending=base.loc[pending_mask,[c for c in ["PICKER","CORREO"] if c in base.columns]].copy()
    pending["ETIQUETA"]=pending.apply(lambda r: f"{r.get('PICKER','')} | {r.get('CORREO','')}".strip(" |"),axis=1)
    pending_options=pending["ETIQUETA"].astype(str).tolist()

    if not pending_options:
        st.success("No hay pickers pendientes de resolver.")
    else:
        selected_labels=st.multiselect("Pickers pendientes (selección múltiple)", pending_options, key="bulk_pending_select")
        selected_rows=pending[pending["ETIQUETA"].isin(selected_labels)].copy()

        roster_turns=sorted({str(x).strip() for x in roster.get("TURNO_MAESTRO",pd.Series(dtype=str)).dropna() if str(x).strip()})
        roster_sups=sorted({str(x).strip() for x in roster.get("SUPERVISOR",pd.Series(dtype=str)).dropna() if str(x).strip() and str(x)!="No asignado"})
        roster_areas=sorted({str(x).strip() for x in roster.get("AREA_MAESTRO",pd.Series(dtype=str)).dropna() if str(x).strip()})
        turn_bulk=st.selectbox("Turno para los seleccionados", roster_turns or ["Matutino","Intermedio","Tarde"], key="bulk_turn")
        sup_bulk=st.selectbox("Supervisor para los seleccionados", ["No asignado"]+roster_sups, key="bulk_sup")
        area_bulk=st.selectbox("Área para los seleccionados", ["No especificada"]+roster_areas, key="bulk_area")

        c_bulk1,c_bulk2=st.columns(2)
        with c_bulk1:
            if st.button("✅ Asignar seleccionados", key="bulk_assign"):
                if not selected_labels:
                    st.warning("Selecciona al menos un picker.")
                else:
                    store.setdefault("master_overrides",{})
                    for _,rr in selected_rows.iterrows():
                        original=str(rr.get("PICKER","")).strip()
                        k=person_key(original)
                        store["master_overrides"][k]={
                            "PICKER":original,
                            "TURNO":turn_bulk,
                            "CORREO":str(rr.get("CORREO","") or "").strip(),
                            "SUPERVISOR":sup_bulk,
                            "AREA_BASE":area_bulk
                        }
                        store.setdefault("master_excluded",[])
                        store["master_excluded"]=[x for x in store["master_excluded"] if x!=k]
                    save_store(store)
                    st.success(f"Se asignaron {len(selected_rows)} pickers.")
                    st.rerun()
        with c_bulk2:
            if st.button("🗑️ Excluir seleccionados", key="bulk_exclude"):
                if not selected_labels:
                    st.warning("Selecciona al menos un picker.")
                else:
                    store.setdefault("master_excluded",[])
                    for _,rr in selected_rows.iterrows():
                        k=person_key(str(rr.get("PICKER","")).strip())
                        if k and k not in store["master_excluded"]:
                            store["master_excluded"].append(k)
                    save_store(store)
                    st.success(f"Se excluyeron {len(selected_rows)} pickers del personal.")
                    st.rerun()

        st.dataframe(selected_rows.drop(columns=["ETIQUETA"],errors="ignore") if not selected_rows.empty else pending.drop(columns=["ETIQUETA"],errors="ignore"), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Administrar asignaciones y exclusiones")
    saved_overrides=store.get("master_overrides",{})
    saved_excluded=store.get("master_excluded",[])
    st.write(f"Asignaciones manuales guardadas: **{len(saved_overrides)}** · Exclusiones guardadas: **{len(saved_excluded)}**")
    if saved_excluded:
        excluded_labels=[next((str(v) for v in base["PICKER"].tolist() if person_key(v)==k),k) for k in saved_excluded]
        restore=st.multiselect("Pickers excluidos para reactivar", excluded_labels, key="bulk_restore_select")
        if st.button("↩️ Reactivar seleccionados", key="bulk_restore"):
            selected_keys={person_key(x) for x in restore}
            store["master_excluded"]=[k for k in saved_excluded if k not in selected_keys]
            save_store(store); st.success(f"Se reactivaron {len(selected_keys)} pickers."); st.rerun()

with f:
    st.download_button("📥 Descargar Excel completo",export(s,fnr,mc,None if sp=="Todos" else sp,roster),f"Analisis_FNR_MC_{periodo}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.info("La app utiliza 3 Excel operativos separados (Pickers, FNR y MC) y una sola plantilla consolidada para turno, correo, supervisor y área.")
    st.success(f"Persistencia activa: {len(store.get('excluded_orders',[]))} pedidos excluidos · {len(store.get('master_overrides',{}))} asignaciones manuales · {len(store.get('master_excluded',[]))} exclusiones de personal · {len(store.get('feedback_rows',[]))} retroalimentaciones guardadas.")
