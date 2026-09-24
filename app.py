import io, re, json, os
from difflib import SequenceMatcher
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Control FNR & Mala Calidad", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

# Estilo visual inspirado en la interfaz limpia de Jüsto: blanco, rojo de marca y tarjetas suaves.
st.markdown("""
<style>
:root { --justo-red:#BD2426; --ink:#272936; --muted:#6f7480; --soft:#f7f7f5; --line:#e7e7e3; --success:#15803d; }
.main .block-container { max-width: 1500px; padding-top: 1.1rem; padding-bottom: 2.5rem; }
[data-testid="stSidebar"] { border-right: 1px solid var(--line); }
section[data-testid="stSidebar"] { width: 320px !important; }
section[data-testid="stSidebar"] > div { width: 320px !important; }
.page-filter { background:#fafaf8; border:1px solid var(--line); border-radius:16px; padding:12px 14px 2px; margin:0 0 18px; }
[data-testid="stMetric"] { background:#fff; border:1px solid var(--line); border-radius:16px; padding:.75rem .9rem; box-shadow:0 2px 10px rgba(30,30,30,.035); }
.kpi-soft-blue { background:#f3f8ff; border:1px solid #dcecff; }
.kpi-soft-red { background:#fff6f6; border:1px solid #f4d9da; }
.kpi-soft-green { background:#f3faf5; border:1px solid #dcefe1; }
.kpi-soft-amber { background:#fffbf1; border:1px solid #f2e7c5; }
.context-banner { background:linear-gradient(90deg,#f7f9fc,#fffafa); border:1px solid #e7e9ee; border-radius:14px; padding:12px 15px; margin:10px 0 16px; }
.context-title { color:var(--ink); font-weight:700; font-size:.92rem; }
.context-detail { color:var(--muted); font-size:.82rem; margin-top:3px; }
.compare-strip { display:flex; gap:10px; flex-wrap:wrap; margin:6px 0 16px; }
.compare-chip { background:#fafaf8; border:1px solid var(--line); border-radius:12px; padding:8px 12px; min-width:150px; }
.compare-chip strong { color:var(--ink); }
.compare-chip span { color:var(--muted); font-size:.78rem; }
div[data-testid="stExpander"] { border:1px solid var(--line); border-radius:14px; overflow:hidden; }
button[kind="primary"] { background:var(--justo-red); border-color:var(--justo-red); }
button[kind="primary"]:hover { background:#a91f21; border-color:#a91f21; }
[data-testid="stTabs"] button[aria-selected="true"] { color:var(--justo-red); border-bottom-color:var(--justo-red); }
.justo-card { background:#fff; border:1px solid var(--line); border-radius:18px; padding:18px 20px; box-shadow:0 3px 14px rgba(30,30,30,.045); margin-bottom:12px; }
.justo-kicker { color:var(--justo-red); font-size:.78rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase; }
.justo-title { color:var(--ink); font-size:1.55rem; font-weight:750; margin:.1rem 0 .35rem; }
.justo-muted { color:var(--muted); font-size:.92rem; }
.status-dot { display:inline-block; width:9px; height:9px; border-radius:50%; margin-right:6px; }
</style>
""", unsafe_allow_html=True)

@st.cache_data(show_spinner=False, max_entries=12)
def sheets_from_bytes(data):
    """Lee un Excel una sola vez por contenido; evita releer los mismos archivos en cada rerun."""
    bio=io.BytesIO(data)
    xls=pd.ExcelFile(bio)
    out={}
    for sh in xls.sheet_names:
        df=pd.read_excel(xls, sheet_name=sh)
        if not df.empty: out[sh]=clean(df)
    return out
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

def _safe_filename(name):
    base=os.path.basename(str(name or "imagen"))
    base=re.sub(r"[^A-Za-z0-9._-]+","_",base)
    return base[:120] or "imagen"

def save_process_image(data, process_id, filename):
    ensure_persist_dir()
    folder=os.path.join(PERSIST_DIR,"process_images",str(process_id))
    os.makedirs(folder, exist_ok=True)
    path=os.path.join(folder,_safe_filename(filename))
    with open(path,"wb") as f: f.write(data)
    return path

def load_process_images(process):
    result=[]
    for path in process.get("imagenes",[]):
        if os.path.exists(path):
            try:
                with open(path,"rb") as f: result.append((path, f.read()))
            except Exception: pass
    return result

def new_process_id():
    return datetime.now().strftime("%Y%m%d%H%M%S%f")

def save_followup_pdf(data, picker, filename):
    """Guarda un PDF de acta/seguimiento asociado a un picker."""
    ensure_persist_dir()
    doc_id=datetime.now().strftime("%Y%m%d%H%M%S%f")
    folder=os.path.join(PERSIST_DIR, "seguimiento_pdfs", person_key(picker) or "picker", doc_id)
    os.makedirs(folder, exist_ok=True)
    safe=_safe_filename(filename)
    path=os.path.join(folder, safe)
    with open(path, "wb") as f: f.write(data)
    return path, doc_id

def load_followup_pdf(doc):
    path=str(doc.get("path", ""))
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, "rb") as f: return f.read()
    except Exception:
        return None

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
    rec = store.setdefault("pickers", {}).setdefault(picker, {"estado":"ACTIVO", "comentarios":[], "acciones":[], "documentos":[]})
    rec.setdefault("estado", "ACTIVO"); rec.setdefault("comentarios", []); rec.setdefault("acciones", []); rec.setdefault("documentos", [])
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

def _excel_safe_df(df):
    """Prepara DataFrames para exportación robusta a Excel/OpenPyXL."""
    if df is None:
        return pd.DataFrame()
    x=df.copy()
    # Evita errores por nombres de columnas duplicados.
    cols=[]; seen={}
    for c in x.columns:
        base=str(c) if str(c).strip() else "Columna"
        n=seen.get(base,0)
        cols.append(base if n==0 else f"{base}_{n}")
        seen[base]=n+1
    x.columns=cols
    # Excel/OpenPyXL no admite NaN/inf ni algunos objetos pandas directamente.
    x=x.replace([float("inf"),float("-inf")],pd.NA)
    for c in x.columns:
        if pd.api.types.is_object_dtype(x[c]) or pd.api.types.is_string_dtype(x[c]):
            x[c]=x[c].map(lambda v: "" if pd.isna(v) else (str(v) if isinstance(v,(dict,list,set,tuple)) else v))
    return x

def _write_sheet(writer, df, sheet_name):
    x=_excel_safe_df(df)
    # Mantener nombres de hoja válidos para Excel.
    name=re.sub(r"[\\/*?:\[\]]", "_", str(sheet_name))[:31] or "Hoja"
    x.to_excel(writer, sheet_name=name, index=False)

def export(summary,fnr,mc,picker,roster=None):
    """Genera el Excel de salida sin romper el dashboard si algún dato viene irregular."""
    b=io.BytesIO()
    writer=pd.ExcelWriter(b,engine="openpyxl")
    try:
        _write_sheet(writer,summary,"Resumen_Pickers")
        _write_sheet(writer,fnr,"Detalle_FNR")
        _write_sheet(writer,mc,"Detalle_MC")
        _write_sheet(writer,products(fnr,picker),"Productos_FNR")
        _write_sheet(writer,products(mc,picker),"Productos_MC")
        _write_sheet(writer,orders(fnr,picker),"Pedidos_FNR")
        _write_sheet(writer,orders(mc,picker),"Pedidos_MC")
        if roster is not None and not roster.empty:
            _write_sheet(writer,roster,"Maestro_Turnos")
        writer.close()
    except Exception:
        try:
            writer.close()
        except Exception:
            pass
        raise
    return b.getvalue()

st.markdown('<div class="justo-kicker">Operación · Coyoacán</div>', unsafe_allow_html=True)
st.title("📊 Control FNR & Mala Calidad")
st.caption("Control operativo de pickers, calidad, seguimiento y procesos")

# Estado persistente: se carga antes de construir los widgets.
store=load_store()
store.setdefault("excluded_orders", [])
store.setdefault("feedback_rows", [])
store.setdefault("master_overrides", {})
store.setdefault("master_excluded", [])
store.setdefault("procesos", [])
store.setdefault("recursos_formatos", [])
store.setdefault("seguimientos_documentos", [])

with st.sidebar:
    st.header("Control operativo")
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
    base=parse_base(choose(sheets_from_bytes(ub.getvalue()),["picker","lineas","resumen"]))
    fnr=parse_inc(choose(sheets_from_bytes(uf.getvalue()),["fnr","detalle"]),"FNR")
    mc=parse_inc(choose(sheets_from_bytes(um.getvalue()),["mc","mala"]),"MC")
    # La plantilla maestra actual tiene una sola hoja: Base_Pickers.
    # Columnas: PICKER, TURNO, CODIGO + CORREO, SUPERVISOR, AREA_BASE.
    master_sheets=sheets_from_bytes(up.getvalue())
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
        st.success(f"Personal identificado: {matched}/{total}")
        st.caption(f"Turnos asignados: {matched_turn}/{total} · Excluidos: {int(excluded_series.sum())}")
        if unmatched:
            st.caption(f"{unmatched} picker(s) sin coincidencia en el maestro.")

# Retira columnas técnicas antes de mostrar/exportar.
base=base.drop(columns=["_MASTER_MATCH"],errors="ignore")

def person_options(df):
    return sorted({str(x).strip() for x in df["PICKER"].tolist() if str(x).strip()}, key=lambda z:z.upper())

def render_person_filters(df, key_prefix, include_picker=True, include_area=True):
    """Filtros locales de cada página; evita depender de un filtro global en el sidebar."""
    turns=["Todos"]+sorted({str(x).strip() for x in df["TURNO"].tolist() if str(x).strip() and str(x)!="__EXCLUIDO__"}, key=lambda z:z.upper())
    sups=["Todos"]+sorted({str(x).strip() for x in df["SUPERVISOR"].tolist() if str(x).strip() and str(x)!="No asignado"}, key=lambda z:z.upper())
    areas=["Todos"]+sorted({str(x).strip() for x in df["AREA_BASE"].tolist() if str(x).strip() and str(x)!="No especificada"}, key=lambda z:z.upper())
    ncols=4 if include_picker and include_area else 3
    c=st.columns(ncols)
    pos=0
    picker_sel="Todos"
    if include_picker:
        with c[pos]:
            search=st.text_input("Buscar picker", placeholder="Escribe un nombre…", key=f"{key_prefix}_search")
            names=person_options(df)
            if search.strip():
                term=norm(search)
                names=[n for n in names if term in norm(n)]
            picker_sel=st.selectbox("Picker",["Todos"]+names, key=f"{key_prefix}_picker")
        pos+=1
    with c[pos]:
        turn_sel=st.selectbox("Turno",turns,key=f"{key_prefix}_turn")
    pos+=1
    with c[pos]:
        sup_sel=st.selectbox("Supervisor",sups,key=f"{key_prefix}_sup")
    pos+=1
    area_sel="Todos"
    if include_area:
        with c[pos]:
            area_sel=st.selectbox("Área",areas,key=f"{key_prefix}_area")
    return picker_sel,turn_sel,sup_sel,area_sel

def apply_person_filters(df, picker_sel="Todos", turn_sel="Todos", sup_sel="Todos", area_sel="Todos"):
    out=df.copy()
    if picker_sel!="Todos": out=out[out.PICKER==picker_sel]
    if turn_sel!="Todos": out=out[out.TURNO.astype(str)==turn_sel]
    if sup_sel!="Todos": out=out[out.SUPERVISOR.astype(str)==sup_sel]
    if area_sel!="Todos": out=out[out.AREA_BASE.astype(str)==area_sel]
    return out

def context_values(df):
    """Opciones del contexto global; se comparten entre todas las pestañas."""
    turns=["Todos"]+sorted({str(x).strip() for x in df["TURNO"].tolist() if str(x).strip() and str(x)!="__EXCLUIDO__"}, key=lambda z:z.upper())
    sups=["Todos"]+sorted({str(x).strip() for x in df["SUPERVISOR"].tolist() if str(x).strip() and str(x)!="No asignado"}, key=lambda z:z.upper())
    areas=["Todos"]+sorted({str(x).strip() for x in df["AREA_BASE"].tolist() if str(x).strip() and str(x)!="No especificada"}, key=lambda z:z.upper())
    return turns,sups,areas

def global_context(df):
    """Lee el contexto elegido en Bodega. Las demás páginas lo heredan automáticamente."""
    turns,sups,areas=context_values(df)
    ctx=st.session_state.setdefault("global_context", {"turno":"Todos","supervisor":"Todos","area":"Todos"})
    if ctx.get("turno") not in turns: ctx["turno"]="Todos"
    if ctx.get("supervisor") not in sups: ctx["supervisor"]="Todos"
    if ctx.get("area") not in areas: ctx["area"]="Todos"
    return ctx,turns,sups,areas

def apply_context(df, ctx, include_turn=True):
    return apply_person_filters(
        df,
        "Todos",
        ctx.get("turno","Todos") if include_turn else "Todos",
        ctx.get("supervisor","Todos"),
        ctx.get("area","Todos"),
    )

def context_reference(df, ctx):
    """Base de comparación: mantiene supervisor/área, pero abre todos los turnos."""
    return apply_person_filters(df,"Todos","Todos",ctx.get("supervisor","Todos"),ctx.get("area","Todos"))

def render_context_banner(ctx, selected_df, reference_df, label="Contexto de análisis"):
    parts=[]
    for k,lab in [("turno","Turno"),("supervisor","Supervisor"),("area","Área")]:
        v=ctx.get(k,"Todos")
        if v!="Todos": parts.append(f"{lab}: {v}")
    scope=" · ".join(parts) if parts else "Todos los turnos"
    sel_p=len(selected_df)
    ref_p=len(reference_df)
    pct=(sel_p/ref_p*100) if ref_p else 0
    st.markdown(
        f"<div class='context-banner'><div class='context-title'>🎯 {label}: {scope}</div>"
        f"<div class='context-detail'>{sel_p:,} pickers · {pct:.1f}% del universo de comparación · Las demás pestañas utilizan este mismo contexto.</div></div>",
        unsafe_allow_html=True)

def render_soft_kpis(cards):
    """Tarjetas KPI con colores muy suaves para facilitar lectura sin saturar la vista."""
    cols=st.columns(len(cards))
    for i,(label,value,detail,tone) in enumerate(cards):
        with cols[i]:
            st.markdown(
                f"<div class='justo-card kpi-soft-{tone}'><div class='justo-muted'>{label}</div>"
                f"<div class='justo-title' style='font-size:1.55rem;margin:.18rem 0'>{value}</div>"
                f"<div class='justo-muted'>{detail}</div></div>",
                unsafe_allow_html=True)

def render_turn_comparison(selected_base, reference_base, selected_fnr, reference_fnr, selected_mc, reference_mc):
    """Muestra cantidad y porcentaje del contexto contra el universo de comparación."""
    def pct(v,d): return (v/d*100) if d else 0
    sl=selected_base.LINEAS.sum(); rl=reference_base.LINEAS.sum()
    sp=selected_base.PEDIDOS.sum(); rp=reference_base.PEDIDOS.sum()
    sf=selected_fnr.INCIDENCIAS.sum(); rf=reference_fnr.INCIDENCIAS.sum()
    sm=selected_mc.INCIDENCIAS.sum(); rm=reference_mc.INCIDENCIAS.sum()
    sfp,_,_,sfr,s_mr=monthly_kpis(selected_base,selected_fnr,selected_mc)
    _,_,_,rfr,r_mr=monthly_kpis(reference_base,reference_fnr,reference_mc)
    cards=[
        ("Líneas",sl, pct(sl,rl), "del universo"),
        ("Pedidos",sp, pct(sp,rp), "del universo"),
        ("FNR",sf, pct(sf,rf), "de incidencias"),
        ("MC",sm, pct(sm,rm), "de incidencias"),
    ]
    chips="".join(f"<div class='compare-chip'><strong>{lab}: {val:,.0f}</strong><br><span>{share:.1f}% {sub}</span></div>" for lab,val,share,sub in cards)
    st.markdown(f"<div class='compare-strip'>{chips}</div>",unsafe_allow_html=True)
    c1,c2=st.columns(2)
    c1.metric("FNR del contexto", f"{sfr:.2f}%" if sfr is not None else "N/D", f"vs {rfr:.2f}% del universo" if rfr is not None else "")
    c2.metric("MC del contexto", f"{s_mr:.2f}%" if s_mr is not None else "N/D", f"vs {r_mr:.2f}% del universo" if r_mr is not None else "")

s_view=s.copy()
if "_EXCLUDED_PERSONNEL" in s_view.columns:
    s_view=s_view[~s_view["_EXCLUDED_PERSONNEL"].fillna(False)]

# Defaults para pestañas que no necesitan filtros globales.
sp="Todos"
stn="Todos"
ssup="Todos"
sar="Todos"
base_view=base[base.PICKER.isin(s_view.PICKER)]
fnr_view=fnr[fnr.PICKER.isin(s_view.PICKER)]
mc_view=mc[mc.PICKER.isin(s_view.PICKER)]

a,b,c,d,e,g,h,f,j=st.tabs(["🏠 Bodega","👤 Picker","🏷️ Artículos","📦 Pedidos","🌙 Turnos / Áreas","👥 Supervisores","🛡️ Seguimiento","📥 Exportar","📌 Pendientes & Procesos"])

with a:
    st.subheader(f"Resumen de bodega — {periodo}")
    st.caption("El contexto elegido aquí se comparte automáticamente con todas las pestañas.")
    ctx,turns,sups,areas=global_context(s_view)
    with st.container(border=True):
        st.markdown("**🎯 Contexto global de análisis**")
        c1,c2,c3=st.columns(3)
        with c1:
            st.selectbox("Turno",turns,key="global_turno")
        with c2:
            st.selectbox("Supervisor",sups,key="global_supervisor")
        with c3:
            st.selectbox("Área",areas,key="global_area")
    ctx["turno"]=st.session_state.get("global_turno","Todos")
    ctx["supervisor"]=st.session_state.get("global_supervisor","Todos")
    ctx["area"]=st.session_state.get("global_area","Todos")

    s_bodega=apply_context(s_view,ctx)
    s_reference=context_reference(s_view,ctx)
    base_bodega=base[base.PICKER.isin(s_bodega.PICKER)]
    fnr_bodega=fnr[fnr.PICKER.isin(s_bodega.PICKER)]
    mc_bodega=mc[mc.PICKER.isin(s_bodega.PICKER)]
    base_reference=base[base.PICKER.isin(s_reference.PICKER)]
    fnr_reference=fnr[fnr.PICKER.isin(s_reference.PICKER)]
    mc_reference=mc[mc.PICKER.isin(s_reference.PICKER)]
    render_context_banner(ctx,s_bodega,s_reference)

    lines=base_bodega.LINEAS.sum(); F=fnr_bodega.INCIDENCIAS.sum(); M=mc_bodega.INCIDENCIAS.sum()
    total_pedidos,fnr_pedidos,mc_pedidos,fnr_rate,mc_rate=monthly_kpis(base_bodega,fnr_bodega,mc_bodega)
    render_soft_kpis([
        ("Líneas",f"{lines:,.0f}",f"{lines/base_reference.LINEAS.sum()*100:.1f}% del universo" if base_reference.LINEAS.sum() else "Sin referencia","blue"),
        ("Pedidos",f"{total_pedidos:,.0f}",f"{total_pedidos/base_reference.PEDIDOS.sum()*100:.1f}% del universo" if base_reference.PEDIDOS.sum() else "Sin referencia","green"),
        ("Pickers",f"{len(s_bodega):,}",f"{len(s_bodega)/len(s_reference)*100:.1f}% del universo" if len(s_reference) else "Sin referencia","blue"),
    ])
    render_soft_kpis([
        ("FNR mensual",f"{fnr_rate:.2f}%" if fnr_rate is not None else "N/D","Objetivo < 1.50%","red"),
        ("MC mensual",f"{mc_rate:.2f}%" if mc_rate is not None else "N/D","Objetivo < 1.00%","amber"),
        ("Fuera objetivo",f"{int((s_bodega.ESTADO=="🔴 FUERA DE OBJETIVO").sum()):,}",f"de {len(s_bodega):,} pickers","red"),
    ])
    if fnr_rate is not None and mc_rate is not None:
        st.caption(f"KPI mensual del contexto: {fnr_pedidos:,} pedidos con FNR / {total_pedidos:,} pedidos = {fnr_rate:.2f}% · {mc_pedidos:,} pedidos con MC / {total_pedidos:,} pedidos = {mc_rate:.2f}%")
    else:
        st.caption("No hay suficientes pedidos para calcular el KPI mensual.")

    st.subheader("Comparativo del contexto")
    render_turn_comparison(base_bodega,base_reference,fnr_bodega,fnr_reference,mc_bodega,mc_reference)

    st.subheader("Indicador operativo por líneas")
    k=st.columns(2)
    k[0].metric("FNR / líneas",f"{F/lines*100:.2f}%" if lines else "N/D")
    k[1].metric("MC / líneas",f"{M/lines*100:.2f}%" if lines else "N/D")
    st.subheader("Detalle por picker")
    st.dataframe(s_bodega,use_container_width=True,hide_index=True)

with b:
    ctx,_,_,_=global_context(s_view)
    selected_context=apply_context(s_view,ctx)
    st.subheader("👤 Ficha de picker")
    st.caption("El turno, supervisor y área se heredan del contexto elegido en Bodega. Aquí solo buscas el picker.")
    render_context_banner(ctx,selected_context,context_reference(s_view,ctx),"Contexto heredado")
    names=person_options(selected_context)
    with st.container(border=True):
        search=st.text_input("🔎 Buscar picker",placeholder="Escribe parte del nombre…",key="picker_page_search")
        filtered_names=names
        if search.strip():
            term=norm(search)
            filtered_names=[n for n in names if term in norm(n)]
        picker_choice=st.selectbox("Selecciona un picker",["Todos"]+filtered_names,key="picker_page_picker")
    sp=picker_choice
    if sp=="Todos":
        st.info("Escribe parte del nombre o selecciona un picker para consultar su ficha completa.")
    else:
        r=s[s.PICKER==sp].iloc[0]
        st.markdown(f"<div class='justo-card'><div class='justo-kicker'>Ficha de picker</div><div class='justo-title'>{sp}</div><div class='justo-muted'>Turno: {r.TURNO} · Supervisor: {r.SUPERVISOR} · Área: {r.AREA_BASE} · Usuario: {r.CORREO}</div></div>", unsafe_allow_html=True)
        q=st.columns(4)
        q[0].metric("Líneas",f"{r.LINEAS:,.0f}")
        q[1].metric("FNR",f"{r.FNR:,.0f}",f"{r['FNR_%']:.2f}%")
        q[2].metric("MC",f"{r.MC:,.0f}",f"{r['MC_%']:.2f}%")
        q[3].metric("Estado",r.ESTADO)

        st.subheader("Incidencias del picker")
        cc=st.columns(2)
        with cc[0]: st.dataframe(products(fnr,sp).head(15),use_container_width=True,hide_index=True)
        with cc[1]: st.dataframe(products(mc,sp).head(15),use_container_width=True,hide_index=True)
        st.subheader("Pedidos FNR")
        st.warning(
            "⚠️ **Antes de tomar en cuenta los FNR, revisa si la factura/orden tiene algún reporte o incidencia registrada.** "
            "Valida primero la orden en Backoffice para confirmar si el faltante corresponde realmente a un FNR."
        )
        st.markdown(
            "🔎 **Revisar factura / orden en Backoffice:** "
            "[Abrir órdenes en Backoffice](https://orders.backoffice.justo.cloud/es/orders)"
        )
        st.dataframe(orders(fnr,sp).head(25),use_container_width=True,hide_index=True)

        rec=picker_record(store,sp)
        st.subheader("📝 Retroalimentación y seguimiento")
        st.caption("Todo lo que registres aquí queda guardado en el historial del picker.")
        fb_col, act_col=st.columns(2)
        with fb_col:
            with st.form(f"feedback_picker_form_{sp}", clear_on_submit=True):
                fb_sup=st.text_input("Supervisor que registra",value=str(r.get("SUPERVISOR","")))
                fortalezas=st.text_area("Fortalezas observadas")
                mejora=st.text_area("Punto de mejora")
                acuerdo=st.text_area("Acuerdo / acción de mejora")
                if st.form_submit_button("Guardar retroalimentación", type="primary"):
                    store.setdefault("feedback_rows",[]).append({
                        "FECHA":datetime.now().strftime("%Y-%m-%d %H:%M"),"SUPERVISOR":fb_sup.strip() or "No especificado",
                        "TURNO":str(r.get("TURNO","")),"PICKER":sp,"FNR_LINEAS_%":float(r["FNR_%"]),
                        "MC_LINEAS_%":float(r["MC_%"]),"ESTADO":str(r["ESTADO"]),"FORTALEZAS":fortalezas,
                        "PUNTO_MEJORA":mejora,"ACUERDO":acuerdo})
                    save_store(store); st.success("Retroalimentación guardada."); st.rerun()
        with act_col:
            with st.form(f"accion_picker_form_{sp}", clear_on_submit=True):
                act_type=st.selectbox("Tipo de seguimiento",["Llamada de atención","Advertencia verbal 1","Advertencia verbal 2","Acta 1","Acta 2","Acta 3","Cero tolerancia"])
                act_sup=st.text_input("Supervisor",value=str(r.get("SUPERVISOR","")))
                act_motivo=st.text_area("Motivo / detalle")
                if st.form_submit_button("Guardar llamada / acta", type="primary"):
                    rec["acciones"].append({"fecha":datetime.now().strftime("%Y-%m-%d %H:%M"),"accion":act_type,"supervisor":act_sup.strip() or "No especificado","motivo":act_motivo.strip()})
                    save_store(store); st.success("Seguimiento guardado."); st.rerun()

        st.markdown("### 🔗 Formatos y recursos rápidos")
        st.caption("Accede desde aquí a tus formatos de retroalimentación, llamadas de atención, actas y seguimiento.")
        recursos=store.get("recursos_formatos",[])
        if recursos:
            cols=st.columns(min(3,max(1,len(recursos))))
            for idx,recurso in enumerate(recursos):
                with cols[idx % len(cols)]:
                    titulo=str(recurso.get("titulo","Formato")); tipo=str(recurso.get("tipo","Recurso")); url=str(recurso.get("url",""))
                    st.markdown(f"<div class='justo-card'><div class='justo-kicker'>{tipo}</div><div class='justo-title' style='font-size:1rem'>{titulo}</div></div>",unsafe_allow_html=True)
                    if url: st.link_button("🔗 Abrir formato",url,use_container_width=True)
                    if st.button("🗑️ Quitar",key=f"delete_recurso_{idx}",use_container_width=True):
                        store["recursos_formatos"].pop(idx); save_store(store); st.rerun()
        else:
            st.info("Todavía no hay formatos configurados. Puedes agregar aquí enlaces de Google Forms, Drive, Excel, Word u otra plataforma.")

        with st.expander("⚙️ Administrar formatos / enlaces", expanded=False):
            with st.form(f"recurso_form_{sp}", clear_on_submit=True):
                rr1,rr2=st.columns([1,2])
                with rr1: recurso_tipo=st.selectbox("Tipo",["Retroalimentación","Seguimiento","Llamada de atención","Acta","Otro"])
                with rr2: recurso_titulo=st.text_input("Nombre del formato",placeholder="Ej. Formato de retroalimentación semanal")
                recurso_url=st.text_input("Enlace",placeholder="https://...")
                if st.form_submit_button("➕ Guardar enlace",type="primary"):
                    url=recurso_url.strip()
                    if not recurso_titulo.strip() or not url.startswith(("http://","https://")):
                        st.error("Captura un nombre y un enlace válido que comience con http:// o https://.")
                    else:
                        store.setdefault("recursos_formatos",[]).append({"tipo":recurso_tipo,"titulo":recurso_titulo.strip(),"url":url,"fecha":datetime.now().strftime("%Y-%m-%d %H:%M")}); save_store(store); st.success("Enlace guardado."); st.rerun()

        picker_fb=[x for x in store.get("feedback_rows",[]) if str(x.get("PICKER",""))==sp]
        if picker_fb:
            st.markdown("**Historial de retroalimentaciones**")
            st.dataframe(pd.DataFrame(picker_fb).sort_values("FECHA",ascending=False),use_container_width=True,hide_index=True)
        if rec.get("comentarios"):
            st.markdown("**Comentarios**")
            st.dataframe(pd.DataFrame(rec["comentarios"]).sort_values("fecha",ascending=False),use_container_width=True,hide_index=True)
        if rec.get("acciones"):
            st.markdown("**Llamadas de atención / actas**")
            st.dataframe(pd.DataFrame(rec["acciones"]).sort_values("fecha",ascending=False),use_container_width=True,hide_index=True)

        # El expediente documental es el mismo que usa la pestaña Seguimiento.
        # Así, los PDFs/enlaces cargados en Seguimiento también quedan visibles en Picker.
        picker_docs=sorted(rec.get("documentos",[]) or [], key=lambda x:str(x.get("fecha","")), reverse=True)
        st.markdown("### 📄 Actas y documentos vinculados")
        if picker_docs:
            for i,doc in enumerate(picker_docs):
                tipo=str(doc.get("tipo","Documento")); titulo=str(doc.get("titulo",doc.get("archivo",tipo)))
                fecha=str(doc.get("fecha","")); supervisor=str(doc.get("supervisor",r.get("SUPERVISOR","")))
                archivo=load_followup_pdf(doc)
                with st.container(border=True):
                    d1,d2,d3=st.columns([1.1,3.1,1.4])
                    with d1:
                        st.markdown(f"**{tipo}**")
                        st.caption(fecha)
                    with d2:
                        st.markdown(f"**{titulo}**")
                        if doc.get("detalle"): st.caption(doc.get("detalle"))
                        st.caption(f"Registró: {supervisor}")
                    with d3:
                        if archivo is not None:
                            st.download_button("📥 Ver PDF",archivo,file_name=str(doc.get("archivo") or f"{tipo}.pdf"),mime="application/pdf",key=f"picker_doc_download_{sp}_{doc.get('id',i)}",use_container_width=True)
                        if doc.get("url"):
                            st.link_button("🔗 Abrir enlace",str(doc.get("url")),use_container_width=True)
        else:
            st.info("No hay actas o documentos vinculados. Puedes agregarlos desde la pestaña 🛡️ Seguimiento; quedarán visibles aquí automáticamente.")

        if not picker_fb and not rec.get("comentarios") and not rec.get("acciones") and not picker_docs:
            st.info("Este picker todavía no tiene retroalimentaciones, seguimientos ni documentos registrados.")

with c:
    ctx,_,_,_=global_context(s_view)
    st.subheader("🏷️ Artículos")
    st.caption("Esta página respeta el mismo turno, supervisor y área seleccionados en Bodega.")
    selected=apply_context(s_view,ctx)
    render_context_banner(ctx,selected,context_reference(s_view,ctx),"Contexto heredado")
    tipo=st.radio("Tipo",["FNR","MC"],horizontal=True,key="articulos_tipo")
    data=fnr if tipo=="FNR" else mc
    data=data[data.PICKER.isin(selected.PICKER)]
    art=data.groupby(["PRODUCTO","AREA"],as_index=False).INCIDENCIAS.sum().rename(columns={"INCIDENCIAS":"CANTIDAD"}).sort_values("CANTIDAD",ascending=False)
    st.dataframe(art,use_container_width=True,hide_index=True)
    st.caption(f"{len(art):,} artículos con incidencia · {int(art.CANTIDAD.sum()) if not art.empty else 0:,} incidencias dentro del contexto seleccionado.")

with d:
    ctx,_,_,_=global_context(s_view)
    st.subheader("📦 Pedidos")
    st.caption("Esta página respeta el mismo turno, supervisor y área seleccionados en Bodega.")
    selected=apply_context(s_view,ctx)
    render_context_banner(ctx,selected,context_reference(s_view,ctx),"Contexto heredado")
    tipo=st.radio("Incidencia",["FNR","MC"],horizontal=True,key="pedidos_tipo")
    data=fnr if tipo=="FNR" else mc
    data=data[data.PICKER.isin(selected.PICKER)]
    pedidos_view=orders(data)
    st.dataframe(pedidos_view,use_container_width=True,hide_index=True)
    st.caption("PICKERS indica cuántos pickers aparecen en el mismo pedido.")

with e:
    ctx,_,_,_=global_context(s_view)
    st.subheader("🌙 Turnos / Áreas")
    st.caption("El contexto seleccionado se mantiene, pero el comparativo conserva todos los turnos para no perder proporciones.")
    selected=apply_context(s_view,ctx)
    reference=context_reference(s_view,ctx)
    render_context_banner(ctx,selected,reference,"Contexto heredado")
    bsel=base[base.PICKER.isin(selected.PICKER)]
    fsel=fnr[fnr.PICKER.isin(selected.PICKER)]
    msel=mc[mc.PICKER.isin(selected.PICKER)]
    bref=base[base.PICKER.isin(reference.PICKER)]
    fref=fnr[fnr.PICKER.isin(reference.PICKER)]
    mref=mc[mc.PICKER.isin(reference.PICKER)]
    render_turn_comparison(bsel,bref,fsel,fref,msel,mref)
    st.subheader("FNR por turno")
    st.dataframe(groups(fsel,bsel,"TURNO"),use_container_width=True,hide_index=True)
    st.subheader("MC por turno")
    st.dataframe(groups(msel,bsel,"TURNO"),use_container_width=True,hide_index=True)
    st.subheader("FNR por área")
    st.dataframe(groups(fsel,bsel,"AREA"),use_container_width=True,hide_index=True)
    st.subheader("MC por área")
    st.dataframe(groups(msel,bsel,"AREA"),use_container_width=True,hide_index=True)
    st.warning("El % / líneas por área solo aparece si existe un denominador real de líneas por área.")

with g:
    ctx,_,_,_=global_context(s_view)
    st.subheader("👥 Supervisores")
    st.caption("Vista operativa bajo el mismo contexto global. Las cantidades se mantienen y los KPIs conservan su porcentaje.")
    sup_data=apply_context(s_view,ctx)
    reference=context_reference(s_view,ctx)
    render_context_banner(ctx,sup_data,reference,"Contexto heredado")
    if not sup_data.empty:
        sup_summary=(sup_data.groupby("SUPERVISOR",as_index=False)
            .agg(PICKERS=("PICKER","nunique"),LINEAS=("LINEAS","sum"),FNR=("FNR","sum"),MC=("MC","sum"))
            .sort_values("FNR",ascending=False))
        sup_summary["FNR %"]=(sup_summary["FNR"]/sup_summary["LINEAS"].replace(0,pd.NA)*100).round(2)
        sup_summary["MC %"]=(sup_summary["MC"]/sup_summary["LINEAS"].replace(0,pd.NA)*100).round(2)
        st.dataframe(sup_summary,use_container_width=True,hide_index=True)
        sup_focus=st.selectbox("Supervisor",["Todos"]+sorted([str(x) for x in sup_summary["SUPERVISOR"] if str(x).strip()]))
        if sup_focus!="Todos":
            st.dataframe(sup_data[sup_data["SUPERVISOR"]==sup_focus],use_container_width=True,hide_index=True)
    else:
        st.info("No hay datos para el contexto actual.")

with h:
    st.subheader("🛡️ Seguimiento")
    st.caption("Consulta el historial por picker, revisa llamadas, advertencias y actas, y adjunta los PDFs para conservar evidencia.")

    ctx,_,_,_=global_context(s_view)
    selected_context=apply_context(s_view,ctx)
    render_context_banner(ctx,selected_context,context_reference(s_view,ctx),"Contexto heredado")

    nombres=person_options(selected_context)
    with st.container(border=True):
        st.markdown("**🔎 Buscar picker**")
        buscar=st.text_input("Nombre del picker",placeholder="Escribe parte del nombre…",key="seguimiento_picker_search")
        opciones=nombres
        if buscar.strip():
            term=norm(buscar)
            opciones=[n for n in nombres if term in norm(n)]
        seguimiento_picker=st.selectbox("Picker", ["Selecciona un picker"]+opciones, key="seguimiento_picker_select")

    if seguimiento_picker == "Selecciona un picker":
        st.info("Selecciona un picker para consultar su historial de seguimiento.")
    else:
        r=s[s.PICKER==seguimiento_picker].iloc[0]
        rec=picker_record(store,seguimiento_picker)
        acciones=rec.get("acciones",[]) or []
        feedback=[x for x in store.get("feedback_rows",[]) if str(x.get("PICKER",""))==seguimiento_picker]
        documentos=rec.get("documentos",[]) or []

        st.markdown(
            f"<div class='justo-card'><div class='justo-kicker'>Expediente de seguimiento</div>"
            f"<div class='justo-title'>{seguimiento_picker}</div>"
            f"<div class='justo-muted'>Turno: {r.TURNO} · Supervisor: {r.SUPERVISOR} · Área: {r.AREA_BASE}</div></div>",
            unsafe_allow_html=True
        )

        k1,k2,k3,k4=st.columns(4)
        k1.metric("Retroalimentaciones", f"{len(feedback):,}")
        k2.metric("Seguimientos", f"{len(acciones):,}")
        k3.metric("Actas / llamadas", f"{sum(1 for x in acciones if 'Acta' in str(x.get('accion','')) or 'Llamada' in str(x.get('accion',''))):,}")
        k4.metric("PDF / documentos", f"{len(documentos):,}")

        st.markdown("### 📋 Historial de acciones")
        if acciones:
            hist=pd.DataFrame(acciones).copy()
            hist=hist.rename(columns={"fecha":"Fecha","accion":"Tipo","supervisor":"Supervisor","motivo":"Motivo"})
            cols=[c for c in ["Fecha","Tipo","Supervisor","Motivo"] if c in hist.columns]
            st.dataframe(hist[cols].sort_values("Fecha",ascending=False),use_container_width=True,hide_index=True)
        else:
            st.info("Este picker todavía no tiene llamadas, advertencias o actas registradas.")

        st.markdown("### 📄 Actas y documentos de seguimiento")
        st.caption("Arrastra un PDF aquí para asociarlo al picker. También puedes guardar un enlace si el documento está en Drive, SharePoint u otra plataforma.")
        with st.form(f"seguimiento_documento_form_{seguimiento_picker}", clear_on_submit=True):
            dc1,dc2=st.columns([1,2])
            with dc1:
                doc_tipo=st.selectbox("Tipo de documento", ["Acta 1","Acta 2","Acta 3","Llamada de atención","Advertencia verbal 1","Advertencia verbal 2","Otro"])
            with dc2:
                doc_titulo=st.text_input("Nombre / referencia", placeholder="Ej. Acta por FNR — septiembre 2026")
            doc_detalle=st.text_area("Detalle / motivo", placeholder="Qué originó el seguimiento y cualquier dato importante…")
            doc_pdf=st.file_uploader("📎 Adjuntar PDF", type=["pdf"], accept_multiple_files=False, key=f"seguimiento_pdf_{seguimiento_picker}")
            doc_url=st.text_input("🔗 O vincular documento externo", placeholder="https://...")
            if st.form_submit_button("💾 Guardar documento", type="primary"):
                titulo=doc_titulo.strip() or (doc_pdf.name if doc_pdf is not None else doc_tipo)
                url=doc_url.strip()
                if doc_pdf is None and not url:
                    st.error("Adjunta un PDF o captura un enlace externo.")
                elif url and not url.startswith(("http://","https://")):
                    st.error("El enlace debe comenzar con http:// o https://")
                else:
                    registro={
                        "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "tipo": doc_tipo,
                        "titulo": titulo,
                        "detalle": doc_detalle.strip(),
                        "supervisor": str(r.get("SUPERVISOR","")),
                        "url": url,
                        "path":"",
                        "archivo":""
                    }
                    if doc_pdf is not None:
                        path,doc_id=save_followup_pdf(doc_pdf.getvalue(),seguimiento_picker,doc_pdf.name)
                        registro["path"]=path
                        registro["archivo"]=_safe_filename(doc_pdf.name)
                    rec.setdefault("documentos",[]).append(registro)
                    save_store(store)
                    st.success("Documento de seguimiento guardado y asociado al picker.")
                    st.rerun()

        documentos=sorted(rec.get("documentos",[]) or [], key=lambda x:str(x.get("fecha","")), reverse=True)
        if documentos:
            for i,doc in enumerate(documentos):
                tipo=str(doc.get("tipo","Documento"))
                titulo=str(doc.get("titulo",doc.get("archivo",tipo)))
                fecha=str(doc.get("fecha",""))
                supervisor=str(doc.get("supervisor",r.get("SUPERVISOR","")))
                archivo=load_followup_pdf(doc)
                with st.container(border=True):
                    a1,a2,a3=st.columns([1.2,2.8,1.2])
                    with a1:
                        st.markdown(f"**{tipo}**")
                        st.caption(fecha)
                    with a2:
                        st.markdown(f"**{titulo}**")
                        if doc.get("detalle"): st.caption(doc.get("detalle"))
                        st.caption(f"Registró: {supervisor}")
                    with a3:
                        if archivo is not None:
                            st.download_button("📥 Ver / descargar PDF", archivo, file_name=str(doc.get("archivo") or f"{tipo}.pdf"), mime="application/pdf", key=f"download_doc_{seguimiento_picker}_{doc.get('id',i)}", use_container_width=True)
                        if doc.get("url"):
                            st.link_button("🔗 Abrir enlace", str(doc.get("url")), use_container_width=True)
                        if st.button("🗑️ Eliminar", key=f"delete_doc_{seguimiento_picker}_{doc.get('id',i)}", use_container_width=True):
                            rec["documentos"]=[x for x in rec.get("documentos",[]) if x.get("id")!=doc.get("id")]
                            save_store(store); st.rerun()
        else:
            st.info("No hay PDFs o documentos vinculados para este picker.")

        st.markdown("### 📝 Retroalimentación registrada")
        if feedback:
            fbd=pd.DataFrame(feedback).copy()
            columnas=[c for c in ["FECHA","SUPERVISOR","FORTALEZAS","PUNTO_MEJORA","ACUERDO"] if c in fbd.columns]
            st.dataframe(fbd.sort_values("FECHA",ascending=False)[columnas],use_container_width=True,hide_index=True)
        else:
            st.info("No hay retroalimentaciones registradas para este picker.")

with f:
    st.download_button("📥 Descargar Excel completo",export(s,fnr,mc,None if sp=="Todos" else sp,roster),f"Analisis_FNR_MC_{periodo}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.info("La app utiliza 3 Excel operativos separados (Pickers, FNR y MC) y un Master consolidado para turno, correo, supervisor y área.")
    st.success(f"Persistencia activa: {len(store.get('excluded_orders',[]))} pedidos excluidos · {len(store.get('master_overrides',{}))} asignaciones manuales · {len(store.get('master_excluded',[]))} exclusiones de personal · {len(store.get('feedback_rows',[]))} retroalimentaciones guardadas.")

with j:
    st.subheader("📌 Tablero de pendientes y procesos")
    st.caption("Aquí el responsable puede publicar procesos nuevos, pendientes operativos y material visual de referencia. Todo queda guardado.")

    with st.expander("➕ Crear nuevo proceso / pendiente", expanded=False):
        with st.form("nuevo_proceso_form", clear_on_submit=True):
            p1,p2=st.columns([2,1])
            with p1:
                proc_titulo=st.text_input("Nombre del proceso / pendiente",placeholder="Ej. Validación de pedidos 07:00–12:00")
                proc_obj=st.text_area("Objetivo / qué se debe hacer",height=90)
                proc_pasos=st.text_area("Pasos o instrucciones",height=130,placeholder="1. ...\n2. ...\n3. ...")
            with p2:
                proc_resp=st.text_input("Responsable")
                proc_prior=st.selectbox("Prioridad",["Alta","Media","Baja"])
                proc_status=st.selectbox("Estado",["Pendiente","En proceso","Completado"])
                proc_fecha=st.date_input("Fecha objetivo",value=datetime.now().date())
                proc_imgs=st.file_uploader("Imágenes referenciales",type=["png","jpg","jpeg","webp"],accept_multiple_files=True,key="proc_imgs_new")
            if st.form_submit_button("Guardar proceso",type="primary"):
                if not proc_titulo.strip():
                    st.warning("Escribe un nombre para el proceso.")
                else:
                    pid=new_process_id()
                    paths=[]
                    for img in proc_imgs or []:
                        paths.append(save_process_image(img.getvalue(),pid,img.name))
                    store.setdefault("procesos",[]).append({
                        "id":pid,"titulo":proc_titulo.strip(),"objetivo":proc_obj.strip(),"pasos":proc_pasos.strip(),
                        "responsable":proc_resp.strip(),"prioridad":proc_prior,"estado":proc_status,
                        "fecha_objetivo":str(proc_fecha),"creado":datetime.now().strftime("%Y-%m-%d %H:%M"),"imagenes":paths})
                    save_store(store); st.success("Proceso guardado."); st.rerun()

    procesos=store.get("procesos",[])
    if not procesos:
        st.info("Todavía no hay procesos o pendientes publicados.")
    else:
        status_order=["Pendiente","En proceso","Completado"]
        for status_name in status_order:
            items=[p for p in procesos if p.get("estado")==status_name]
            st.markdown(f"### {status_name} · {len(items)}")
            cols=st.columns(3)
            if not items:
                st.caption("Sin elementos en esta columna.")
            for idx,proc in enumerate(items):
                with cols[idx%3]:
                    priority=proc.get("prioridad","Media")
                    st.markdown(f"<div class='justo-card'><div class='justo-kicker'>{priority}</div><div class='justo-title'>{proc.get('titulo','Sin título')}</div><div class='justo-muted'>Responsable: {proc.get('responsable') or 'Sin asignar'} · Fecha: {proc.get('fecha_objetivo') or 'Sin fecha'}</div></div>",unsafe_allow_html=True)
                    if proc.get("objetivo"): st.write(proc["objetivo"])
                    if proc.get("pasos"):
                        with st.expander("Ver instrucciones"):
                            st.write(proc["pasos"])
                    imgs=load_process_images(proc)
                    if imgs:
                        st.image([data for _,data in imgs],caption=[os.path.basename(path) for path,_ in imgs],use_container_width=True)
                    new_status=st.selectbox("Cambiar estado",status_order,index=status_order.index(proc.get("estado","Pendiente")) if proc.get("estado") in status_order else 0,key=f"proc_status_{proc['id']}")
                    ec1,ec2=st.columns(2)
                    with ec1:
                        if st.button("Guardar estado",key=f"proc_save_{proc['id']}"):
                            for pp in store["procesos"]:
                                if pp.get("id")==proc.get("id"): pp["estado"]=new_status
                            save_store(store); st.rerun()
                    with ec2:
                        if st.button("Eliminar",key=f"proc_del_{proc['id']}"):
                            store["procesos"]=[pp for pp in store["procesos"] if pp.get("id")!=proc.get("id")]
                            save_store(store); st.rerun()

    st.divider()
    st.subheader("💾 Respaldo de configuración")
    st.caption("Para evitar perder asignaciones, seguimientos y procesos si Streamlit Cloud reinicia el contenedor, puedes descargar un respaldo y conservarlo.")
    backup_json=json.dumps(store,ensure_ascii=False,indent=2).encode("utf-8")
    st.download_button("Descargar respaldo de configuración",backup_json,f"Respaldo_Control_FNR_{periodo}.json","application/json")
