# ================================================================
# PROPIEDAD / FIRMA DEL DESARROLLO
# Desarrollado por: Yael Maya Ruíz
# Control FNR & Mala Calidad · Operación Coyoacán
# ================================================================


import io, re, json, os, smtplib, ssl, zipfile, urllib.request, urllib.error, urllib.parse
from difflib import SequenceMatcher
from datetime import datetime
from email.message import EmailMessage
import numpy as np
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
PERSIST_DIR = "app_data"
UPLOAD_HISTORY_DIR = os.path.join(PERSIST_DIR, "upload_history")
MAX_UPLOAD_HISTORY = 20
STORE_FILE = os.path.join(PERSIST_DIR, "picker_seguimiento.json")
LEGACY_STORE_FILE = "picker_seguimiento.json"
PERSIST_FILES = {
    "base_picker": os.path.join(PERSIST_DIR, "base_picker.xlsx"),
    "detalle_fnr": os.path.join(PERSIST_DIR, "detalle_fnr.xlsx"),
    "detalle_mc": os.path.join(PERSIST_DIR, "detalle_mc.xlsx"),
    "plantilla_personal": os.path.join(PERSIST_DIR, "master_pickers.xlsx"),
}
CLOUD_STORE_KEY="state/picker_seguimiento.json"
_CLOUD_BUCKET_CHECKED=False

def cloud_config():
    """Credenciales de Supabase Storage desde secrets o variables de entorno."""
    url=_secret("SUPABASE_URL").strip().rstrip("/")
    key=(_secret("SUPABASE_SECRET_KEY") or _secret("SUPABASE_SERVICE_ROLE_KEY")).strip()
    bucket=_secret("SUPABASE_STORAGE_BUCKET","control-fnr-mc").strip() or "control-fnr-mc"
    return url,key,bucket

def cloud_enabled():
    url,key,_=cloud_config()
    return bool(url and key)

def _cloud_request(method,path,data=None,content_type="application/json",missing_ok=False,extra_headers=None):
    url,key,_=cloud_config()
    if not url or not key:
        raise RuntimeError("Falta configurar SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY.")
    request=urllib.request.Request(url+path,data=data,method=method)
    request.add_header("apikey",key)
    # Las claves nuevas sb_secret_ no son JWT: van en apikey y no en Bearer.
    # Mantener Authorization para la clave legacy service_role (JWT).
    if not key.startswith("sb_secret_"):
        request.add_header("Authorization",f"Bearer {key}")
    if content_type: request.add_header("Content-Type",content_type)
    for name,value in (extra_headers or {}).items(): request.add_header(name,value)
    try:
        with urllib.request.urlopen(request,timeout=30) as response:
            return response.read()
    except urllib.error.HTTPError as exc:
        detail=exc.read().decode("utf-8","replace")[:500]
        # Supabase Storage may return HTTP 400 while its JSON body reports a
        # missing object as statusCode 404 / code NoSuchKey. Treat that as a
        # normal empty result for optional reads (first cloud startup, absent
        # current upload, or missing old asset), just like a direct HTTP 404.
        missing_resource=exc.code==404
        if missing_ok and not missing_resource:
            try:
                error_body=json.loads(detail)
                missing_resource=(
                    str(error_body.get("statusCode",""))=="404"
                    or error_body.get("code")=="NoSuchKey"
                )
            except Exception:
                pass
        if missing_ok and missing_resource: return None
        raise RuntimeError(f"Almacenamiento en nube respondió HTTP {exc.code}: {detail}") from exc
    except Exception as exc:
        raise RuntimeError(f"No se pudo conectar con el almacenamiento en nube: {exc}") from exc

def _cloud_ensure_bucket():
    global _CLOUD_BUCKET_CHECKED
    if _CLOUD_BUCKET_CHECKED or not cloud_enabled(): return
    _,_,bucket=cloud_config()
    bucket_path=urllib.parse.quote(bucket,safe="")
    found=_cloud_request("GET",f"/storage/v1/bucket/{bucket_path}",missing_ok=True)
    if found is None:
        _cloud_request("POST","/storage/v1/bucket",json.dumps({"id":bucket,"name":bucket,"public":False}).encode("utf-8"))
    _CLOUD_BUCKET_CHECKED=True

def cloud_upload(key,data,content_type="application/octet-stream"):
    """Guarda un objeto en el bucket privado con reemplazo idempotente."""
    _cloud_ensure_bucket()
    _,_,bucket=cloud_config()
    bucket_path=urllib.parse.quote(bucket,safe="")
    object_path=urllib.parse.quote(str(key).lstrip("/"),safe="/")
    return _cloud_request("POST",f"/storage/v1/object/{bucket_path}/{object_path}",bytes(data),content_type,
                          extra_headers={"x-upsert":"true","cache-control":"no-store"})

def cloud_download(key,missing_ok=True):
    if not cloud_enabled(): return None
    _cloud_ensure_bucket()
    _,_,bucket=cloud_config()
    bucket_path=urllib.parse.quote(bucket,safe="")
    object_path=urllib.parse.quote(str(key).lstrip("/"),safe="/")
    return _cloud_request("GET",f"/storage/v1/object/authenticated/{bucket_path}/{object_path}",missing_ok=missing_ok)

def cloud_list(prefix,limit=1000):
    if not cloud_enabled(): return []
    _cloud_ensure_bucket()
    _,_,bucket=cloud_config()
    body=json.dumps({"prefix":str(prefix).strip("/"),"limit":int(limit),"offset":0,"sortBy":{"column":"name","order":"desc"}}).encode("utf-8")
    raw=_cloud_request("POST",f"/storage/v1/object/list/{urllib.parse.quote(bucket,safe='')}",body)
    try: return json.loads(raw.decode("utf-8"))
    except Exception as exc: raise RuntimeError("La nube devolvió una respuesta de historial inválida.") from exc

def ensure_persist_dir():
    os.makedirs(PERSIST_DIR, exist_ok=True)

def persist_upload(upload, key):
    """Guarda el Excel actual y una copia versionada local y en la nube."""
    if upload is None:
        return None
    ensure_persist_dir()
    path=PERSIST_FILES[key]
    data=upload.getvalue()
    history_dir=os.path.join(UPLOAD_HISTORY_DIR,key)
    os.makedirs(history_dir,exist_ok=True)
    current_key=f"uploads/current/{key}.xlsx"
    if cloud_enabled():
        cloud_prior=cloud_download(current_key,missing_ok=True)
        changed=cloud_prior!=data
    else:
        prior=sorted([os.path.join(history_dir,n) for n in os.listdir(history_dir) if os.path.isfile(os.path.join(history_dir,n))])
        same_as_latest=False
        if prior:
            try:
                with open(prior[-1],"rb") as f: same_as_latest=(f.read()==data)
            except Exception: pass
        changed=not same_as_latest
    if changed:
        stamp=datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        safe=re.sub(r"[^A-Za-z0-9._-]+","_",os.path.basename(str(getattr(upload,"name","archivo.xlsx"))))[:100]
        hist_path=os.path.join(history_dir,f"{stamp}_{safe or 'archivo.xlsx'}")
        with open(hist_path,"wb") as f: f.write(data)
        if cloud_enabled():
            cloud_upload(f"upload_history/{key}/{os.path.basename(hist_path)}",data,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        saved=sorted([os.path.join(history_dir,n) for n in os.listdir(history_dir) if os.path.isfile(os.path.join(history_dir,n))])
        for old in saved[:-MAX_UPLOAD_HISTORY]:
            try: os.remove(old)
            except OSError: pass
    with open(path, "wb") as f:
        f.write(data)
    if cloud_enabled():
        cloud_upload(current_key,data,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if changed:
        try: upload_history_zip.clear()
        except Exception: pass
    return io.BytesIO(data)

def upload_history_rows():
    rows=[]
    labels={"base_picker":"Base de Pickers / Líneas","detalle_fnr":"Detalle FNR","detalle_mc":"Detalle Mala Calidad","plantilla_personal":"Plantilla consolidada"}
    if cloud_enabled():
        for key,label in labels.items():
            for item in cloud_list(f"upload_history/{key}",MAX_UPLOAD_HISTORY+5):
                name=str(item.get("name","")).strip()
                if not name or item.get("id") is None: continue
                meta=item.get("metadata") or {}
                rows.append({"Archivo":label,"Versión":name,"Fecha de carga":str(item.get("updated_at") or item.get("created_at") or ""),"Tamaño (MB)":round(float(meta.get("size",0) or 0)/1048576,2)})
        return rows
    if not os.path.isdir(UPLOAD_HISTORY_DIR): return rows
    for key,label in labels.items():
        folder=os.path.join(UPLOAD_HISTORY_DIR,key)
        if not os.path.isdir(folder): continue
        for name in sorted(os.listdir(folder),reverse=True):
            path=os.path.join(folder,name)
            if not os.path.isfile(path): continue
            try:
                rows.append({"Archivo":label,"Versión":name,"Fecha de carga":datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S"),"Tamaño (MB)":round(os.path.getsize(path)/1048576,2)})
            except OSError: pass
    return rows

@st.cache_data(show_spinner=False,max_entries=1,ttl=300)
def upload_history_zip():
    out=io.BytesIO()
    with zipfile.ZipFile(out,"w",compression=zipfile.ZIP_DEFLATED) as zf:
        if cloud_enabled():
            for key in ["base_picker","detalle_fnr","detalle_mc","plantilla_personal"]:
                prefix=f"upload_history/{key}"
                for item in cloud_list(prefix,MAX_UPLOAD_HISTORY+5):
                    name=str(item.get("name","")).strip()
                    if not name or item.get("id") is None: continue
                    object_key=f"{prefix}/{name}"
                    data=cloud_download(object_key,missing_ok=True)
                    if data is not None: zf.writestr(f"{key}/{name}",data)
        elif os.path.isdir(UPLOAD_HISTORY_DIR):
            for root,_,files in os.walk(UPLOAD_HISTORY_DIR):
                for name in files:
                    path=os.path.join(root,name)
                    zf.write(path,os.path.relpath(path,UPLOAD_HISTORY_DIR))
    return out.getvalue()

def load_persisted_upload(key):
    """Recupera el último Excel guardado cuando el uploader está vacío."""
    path=PERSIST_FILES[key]
    if cloud_enabled():
        data=cloud_download(f"uploads/current/{key}.xlsx",missing_ok=True)
        if data is not None:
            ensure_persist_dir()
            with open(path,"wb") as f: f.write(data)
            return io.BytesIO(data)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "rb") as f:
            data=f.read()
        if not data:
            return None
        if cloud_enabled():
            cloud_upload(f"uploads/current/{key}.xlsx",data,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        return io.BytesIO(data)
    except Exception:
        return None

def persisted_status():
    status={k: os.path.exists(v) and os.path.getsize(v)>0 for k,v in PERSIST_FILES.items()}
    if cloud_enabled():
        remote={str(item.get("name","")) for item in cloud_list("uploads/current",20) if item.get("id") is not None}
        status={k:(f"{k}.xlsx" in remote) or status[k] for k in status}
    return status

def migrate_local_upload_history_to_cloud():
    """Copia al bucket versiones de Excel que solo existan en disco local."""
    if not cloud_enabled() or not os.path.isdir(UPLOAD_HISTORY_DIR): return 0
    moved=0
    for key in ["base_picker","detalle_fnr","detalle_mc","plantilla_personal"]:
        folder=os.path.join(UPLOAD_HISTORY_DIR,key)
        if not os.path.isdir(folder): continue
        prefix=f"upload_history/{key}"
        existing={str(x.get("name","")) for x in cloud_list(prefix,MAX_UPLOAD_HISTORY+10) if x.get("id") is not None}
        for name in os.listdir(folder):
            path=os.path.join(folder,name)
            if not os.path.isfile(path) or name in existing: continue
            with open(path,"rb") as f: data=f.read()
            cloud_upload(f"{prefix}/{name}",data,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            moved+=1
    if moved:
        try: upload_history_zip.clear()
        except Exception: pass
    return moved

def _safe_filename(name):
    base=os.path.basename(str(name or "imagen"))
    base=re.sub(r"[^A-Za-z0-9._-]+","_",base)
    return base[:120] or "imagen"

def save_process_image(data, process_id, filename):
    ensure_persist_dir()
    folder=os.path.join(PERSIST_DIR,"process_images",str(process_id))
    os.makedirs(folder, exist_ok=True)
    safe=_safe_filename(filename)
    path=os.path.join(folder,safe)
    with open(path,"wb") as f: f.write(data)
    if cloud_enabled():
        key=f"process_images/{_safe_filename(process_id)}/{safe}"
        cloud_upload(key,data,"image/"+(safe.rsplit(".",1)[-1].lower() if "." in safe else "jpeg"))
        return "cloud://"+key
    return path

def load_process_images(process):
    result=[]
    for path in process.get("imagenes",[]):
        if str(path).startswith("cloud://"):
            key=str(path)[len("cloud://"):]
            try:
                data=cloud_download(key,missing_ok=True)
                if data is not None: result.append((os.path.basename(key),data))
            except Exception: pass
            continue
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
    if cloud_enabled():
        key=f"seguimiento_pdfs/{person_key(picker) or 'picker'}/{doc_id}/{safe}"
        cloud_upload(key,data,"application/pdf")
        return "cloud://"+key, doc_id
    return path, doc_id

def load_followup_pdf(doc):
    path=str(doc.get("path", ""))
    if path.startswith("cloud://"):
        try: return cloud_download(path[len("cloud://"):],missing_ok=True)
        except Exception: return None
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, "rb") as f: return f.read()
    except Exception:
        return None

def migrate_local_assets_to_cloud(store):
    """Migra PDFs e imágenes existentes al bucket y actualiza sus referencias."""
    if not cloud_enabled(): return 0,0
    moved=0; missing=0; changed=False
    def migrate(path,kind):
        nonlocal moved,missing,changed
        path=str(path or "")
        if not path or path.startswith("cloud://"): return path
        if not os.path.isfile(path):
            missing+=1
            return path
        try:
            relative=os.path.relpath(path,PERSIST_DIR).replace(os.sep,"/")
            if relative.startswith("../") or relative=="..": relative=f"legacy/{kind}/{_safe_filename(os.path.basename(path))}"
            with open(path,"rb") as f: data=f.read()
            mime="application/pdf" if path.lower().endswith(".pdf") else "application/octet-stream"
            cloud_upload(relative,data,mime)
            moved+=1; changed=True
            return "cloud://"+relative
        except Exception:
            raise
    for process in store.get("procesos",[]) or []:
        updated=[]
        for path in process.get("imagenes",[]) or []:
            updated.append(migrate(path,"process_images"))
        process["imagenes"]=updated
    for record in (store.get("pickers",{}) or {}).values():
        for doc in record.get("documentos",[]) or []:
            old=doc.get("path",""); new=migrate(old,"seguimiento_pdfs")
            if new!=old: doc["path"]=new
    for doc in store.get("seguimientos_documentos",[]) or []:
        old=doc.get("path",""); new=migrate(old,"seguimiento_pdfs")
        if new!=old: doc["path"]=new
    if changed: save_store(store)
    return moved,missing

def load_store():
    """Carga el expediente persistente y migra el JSON antiguo si existe."""
    ensure_persist_dir()
    cloud_data=cloud_download(CLOUD_STORE_KEY,missing_ok=True) if cloud_enabled() else None
    if cloud_data is not None:
        data=json.loads(cloud_data.decode("utf-8"))
        if not isinstance(data,dict): raise RuntimeError("El expediente guardado en la nube no contiene un objeto JSON válido.")
        for key,value in {"pickers":{},"feedback_rows":[],"recursos_formatos":[],"procesos":[],"excluded_orders":[],"master_overrides":{},"master_excluded":[],"seguimientos_documentos":[],"supervisores":[],"upload_meta":{}}.items(): data.setdefault(key,value)
        return data
    candidates=[STORE_FILE, LEGACY_STORE_FILE]
    for path in candidates:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data=json.load(f)
                if not isinstance(data, dict):
                    continue
                data.setdefault("pickers", {})
                data.setdefault("feedback_rows", [])
                data.setdefault("recursos_formatos", [])
                data.setdefault("procesos", [])
                data.setdefault("excluded_orders", [])
                data.setdefault("master_overrides", {})
                data.setdefault("master_excluded", [])
                data.setdefault("seguimientos_documentos", [])
                data.setdefault("supervisores", [])
                data.setdefault("upload_meta", {})
                if path != STORE_FILE or cloud_enabled(): save_store(data)
                return data
            except Exception:
                continue
    data={"pickers": {}, "feedback_rows": [], "recursos_formatos": [], "procesos": [], "excluded_orders": [], "master_overrides": {}, "master_excluded": [], "seguimientos_documentos": [], "supervisores": [], "upload_meta": {}}
    if cloud_enabled(): save_store(data)
    return data

def save_store(store):
    ensure_persist_dir()
    tmp = STORE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f: json.dump(store, f, ensure_ascii=False, indent=2)
    os.replace(tmp, STORE_FILE)
    if cloud_enabled(): cloud_upload(CLOUD_STORE_KEY,json.dumps(store,ensure_ascii=False,indent=2).encode("utf-8"),"application/json")

def picker_record(store, picker, aliases=None):
    """Devuelve un único expediente por persona, aunque cambie el orden/formato del nombre.

    También migra automáticamente expedientes antiguos guardados con el nombre
    del Excel operativo hacia el nombre canónico del Maestro.
    """
    store.setdefault("pickers", {})
    target=str(picker or "").strip()
    aliases=[str(a).strip() for a in (aliases or []) if str(a).strip()]
    candidates=[]
    for name in [target]+aliases:
        if name and name not in candidates: candidates.append(name)
    found_key=None
    # 1) coincidencia exacta / normalizada / por tokens
    target_key=person_key(target); target_tokens=token_key(target)
    for existing in list(store["pickers"].keys()):
        if existing in candidates or (target_key and person_key(existing)==target_key) or (target_tokens and token_key(existing)==target_tokens):
            found_key=existing; break
    # 2) aliases
    if found_key is None:
        for alias in aliases:
            ak=person_key(alias); at=token_key(alias)
            for existing in list(store["pickers"].keys()):
                if (ak and person_key(existing)==ak) or (at and token_key(existing)==at):
                    found_key=existing; break
            if found_key is not None: break
    if found_key is None:
        rec={"estado":"ACTIVO","comentarios":[],"acciones":[],"documentos":[]}
        store["pickers"][target]=rec
        return rec
    rec=store["pickers"].pop(found_key)
    rec.setdefault("estado","ACTIVO"); rec.setdefault("comentarios",[]); rec.setdefault("acciones",[]); rec.setdefault("documentos",[])
    # Si ya existía un registro con el nombre canónico, fusionar sin perder historial.
    if target in store["pickers"] and target != found_key:
        current=store["pickers"][target]
        current.setdefault("comentarios",[]); current.setdefault("acciones",[]); current.setdefault("documentos",[])
        current["comentarios"]=current["comentarios"]+rec.get("comentarios",[])
        current["acciones"]=current["acciones"]+rec.get("acciones",[])
        current["documentos"]=current["documentos"]+rec.get("documentos",[])
        return current
    store["pickers"][target]=rec
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

def identity_name_tokens(value):
    """Quita códigos del identificador y devuelve los tokens del nombre."""
    s=str(value or "").strip()
    if not s or s.lower() in {"nan","none","nat"}: return set()
    if "@" in s: return email_name_tokens(s)
    s=re.sub(r"^(?:jt)?\d+[._\-\s]*", "", s, flags=re.I)
    s=re.sub(r"(?<![a-z0-9])(?:jt)?\d+(?![a-z0-9])", " ", s, flags=re.I)
    return name_tokens(s.replace(".", "_"))

def extract_email_address(value):
    """Extrae y normaliza el correo real aunque Excel traiga CODIGO + CORREO."""
    s=str(value or "").strip().lower()
    if s in {"", "nan", "none", "nat"}: return ""
    m=re.search(r"[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}", s, flags=re.I)
    if not m: return ""
    email=m.group(0).strip().lower()
    local,domain=email.split("@",1)
    local=re.sub(r"^(?:jt)?\d+[._-]+", "", local)
    return f"{local}@{domain}"

def email_key(value):
    """Llave principal de identidad: el correo real normalizado."""
    email=extract_email_address(value)
    if email: return email
    s=str(value or "").strip().lower().replace(" ","")
    if s in {"nan","none","nat"}: return ""
    return s

def person_context_key(row):
    """Llave de filtros: correo para personal registrado, nombre para Sin registrar."""
    category=str(row.get("CATEGORIA","Picker") or "Picker").strip()
    if category=="Sin registrar":
        return "sin-registrar:"+(person_key(row.get("PICKER","")) or "sin-nombre")
    ek=email_key(row.get("CORREO",""))
    return "correo:"+ek if ek else "persona:"+(person_key(row.get("PICKER","")) or "sin-nombre")

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
    """Elige una hoja por nombre, ignorando mayúsculas, espacios y guiones."""
    if not ss:
        raise ValueError("El Excel no contiene hojas con datos.")
    for name,df in ss.items():
        n=norm(name)
        if any(norm(w) in n for w in words):
            return df
    return next(iter(ss.values()))

def parse_base(df):
    """Lee la base operativa aunque NO traiga correo.

    El correo se incorpora después desde Maestro_Personal/plantilla.
    Así la plantilla sí funciona como fuente de identidad y los Excel
    operativos no están obligados a repetir el correo en cada archivo.
    """
    p=col(df,["picker","picker_nombre","nombre","email_picker","persona"])
    l=col(df,["total_lineas","lineas_totales","lineas","total_lineas_pickeadas","lineas_totales_pickeadas"])
    ped=col(df,["total_pedidos","pedidos_totales","pedidos","ordenes","total_ordenes"])
    sh=col(df,["turno","shift"]); ar=col(df,["area","departamento","department"])
    em=col(df,["codigo_correo","codigo + correo","correo","email","usuario"])
    if not l:
        raise ValueError("La base necesita una columna de Total líneas / Líneas totales.")
    # Conservar lo que viene en PICKER como descriptor/origen. En algunas
    # bases operativas esta columna trae código + correo, aunque CORREO no exista.
    if p:
        picker_values=df[p].fillna("").astype(str).str.strip()
    elif em:
        picker_values=df[em].fillna("").astype(str).str.strip().map(email_name_tokens).map(lambda z:" ".join(sorted(z)))
    else:
        raise ValueError("La base necesita PICKER/nombre o una columna de CORREO para identificar al personal.")
    correo_values=df[em].fillna("").astype(str).str.strip() if em else pd.Series("",index=df.index,dtype=str)
    correo_values=correo_values.mask(
        correo_values.eq("") | correo_values.str.lower().isin({"nan","none","nat"}),
        picker_values.where(
            picker_values.map(lambda v: bool(extract_email_address(v) or extract_code(v))),
            ""
        )
    )
    x=pd.DataFrame({
        "PICKER":picker_values,
        "LINEAS":pd.to_numeric(df[l],errors="coerce").fillna(0),
        "PEDIDOS":pd.to_numeric(df[ped],errors="coerce").fillna(0) if ped else 0,
        "TURNO":df[sh].fillna("").astype(str).str.strip() if sh else "No especificado",
        "AREA_BASE":df[ar].fillna("").astype(str).str.strip() if ar else "No especificada",
        "CORREO":correo_values
    })
    x["_KEY"]=x["PICKER"].map(person_key)
    x["_TOKEN_KEY"]=x["PICKER"].map(token_key)
    x["_CODE_KEY"]=x["CORREO"].map(extract_code)
    x["_EMAIL_TOKENS"]=x["CORREO"].map(email_name_tokens)
    x["_EMAIL_KEY"]=x["CORREO"].map(email_key)
    # Sumar todas las filas de una persona: conservar solo la de más líneas
    # descartaba producción cuando el Excel separaba turnos o áreas.
    x["_IDENTITY_KEY"]=x.apply(lambda r: ("email:"+r["_EMAIL_KEY"]) if str(r["_EMAIL_KEY"]).strip() else (("name:"+r["_KEY"]) if str(r["_KEY"]).strip() else ("row:"+str(r.name))),axis=1)
    rows=[]
    for _,g in x.groupby("_IDENTITY_KEY",sort=False):
        row=g.iloc[g["LINEAS"].astype(float).argmax()].copy()
        row["LINEAS"]=pd.to_numeric(g["LINEAS"],errors="coerce").fillna(0).sum()
        row["PEDIDOS"]=pd.to_numeric(g["PEDIDOS"],errors="coerce").fillna(0).sum()
        for field,plural in [("TURNO","Varios turnos"),("AREA_BASE","Varias áreas")]:
            vals=[str(v).strip() for v in g[field].tolist() if str(v).strip() and str(v).strip().lower() not in {"nan","none"}]
            unique=list(dict.fromkeys(vals))
            row[field]=unique[0] if len(unique)==1 else (plural if unique else ("No especificado" if field=="TURNO" else "No especificada"))
        rows.append(row)
    out=pd.DataFrame(rows).drop(columns=["_IDENTITY_KEY"],errors="ignore")
    return out.reset_index(drop=True)

def parse_roster(df, turno_fijo=None):
    """Maestro consolidado: PICKER, TURNO, CODIGO + CORREO, SUPERVISOR y AREA_BASE."""
    p=col(df,["picker","picker_nombre","email_picker","nombre","persona"])
    sh=col(df,["turno","shift"])
    em=col(df,["codigo_correo","codigo + correo","correo","email","usuario"])
    sup=col(df,["supervisor","supervisor_nombre","jefe","responsable"])
    ar=col(df,["area_base","area","departamento","department"])
    if not p and not em: raise ValueError("La hoja de personal necesita PICKER o CORREO.")
    picker_series=df[p].astype(str).str.strip() if p else df[em].astype(str).str.strip().map(email_name_tokens).map(lambda z:" ".join(sorted(z)))
    x=pd.DataFrame({
        "PICKER":picker_series,
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
    x["_EMAIL_KEY"]=x["CORREO"].map(email_key)
    has_email=x["_EMAIL_KEY"].astype(str).str.strip().ne("")
    x=pd.concat([x[has_email].drop_duplicates("_EMAIL_KEY",keep="last"), x[~has_email].drop_duplicates("_KEY",keep="last")],ignore_index=True)
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
    toks=identity_name_tokens(value)
    key=person_key(" ".join(sorted(toks)))
    code=extract_code(code_value) or extract_code(value)
    email_toks=email_name_tokens(code_value) or toks

    # 1. Nombre exacto
    hits=roster[roster["_KEY"].astype(str)==key] if key else roster.iloc[0:0]
    if len(hits)==1: return hits.iloc[0]

    # 2. Todas las palabras coinciden, sin importar orden
    if toks:
        hits=roster[roster["_TOKEN_KEY"].astype(str)=="_".join(sorted(toks))]
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
    # Fallback controlado para nombres con segundo apellido, iniciales o pequeñas
    # diferencias: exige al menos 2 tokens distintivos y que no haya empate cercano.
    distinctive={t for t in toks if len(t)>=4}
    best_dist={t for t in name_tokens(rr.get("PICKER","")) if len(t)>=4}
    dist_overlap=len(distinctive & best_dist)
    if len(distinctive)>=2 and dist_overlap>=2 and cov_in>=0.80 and (score-second>=0.08 or score>=0.76):
        return rr
    return None

def apply_roster(base, roster):
    """Asocia los 3 Excel operativos con la PLANTILLA CONSOLIDADA.

    La única llave de identidad entre archivos es CORREO.
    El nombre canónico, turno, supervisor y área se toman de la plantilla.
    Si un registro operativo no trae correo o el correo no existe en la
    plantilla, queda como NO ASIGNADO; no se hace cruce aproximado por nombre.
    """
    if roster is None or roster.empty: return base
    roster=roster.copy()
    # template_roster_from_upload debe exponer esta llave, pero calcularla aquí
    # también protege el cruce si el maestro llega desde otra ruta.
    if "_EMAIL_KEY" not in roster.columns:
        roster["_EMAIL_KEY"]=roster.get("CORREO",pd.Series("",index=roster.index)).map(email_key)
    x=base.copy()
    x["_SOURCE_PICKER"]=x["PICKER"].astype(str).str.strip()
    roster_email={}
    if "_EMAIL_KEY" in roster.columns:
        for _,rr in roster.iterrows():
            ek=email_key(rr.get("CORREO",""))
            if ek: roster_email[ek]=rr
    matches=[]
    for _,row in x.iterrows():
        ek=email_key(row.get("CORREO",""))
        if ek:
            rr=roster_email.get(ek)
        else:
            rr=None
        # Recuperar la ruta que funcionaba en la versión anterior: el correo
        # tiene prioridad; si falta, usar la resolución conservadora por código/nombre.
        if rr is None:
            rr=_match_master_row(row.get("_SOURCE_PICKER",row.get("PICKER","")),roster,row.get("CORREO",""))
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
    # Contextos explícitos para FNR/MC y segmentaciones posteriores.
    x["TURNO"]=x["TURNO"].fillna("").astype(str).str.strip()
    x["AREA_BASE"]=x["AREA_BASE"].fillna("").astype(str).str.strip()
    x.loc[x["TURNO"].eq(""),"TURNO"]="No especificado"
    x.loc[x["AREA_BASE"].eq(""),"AREA_BASE"]="No especificada"
    # CORREO_KEY queda visible para que TODOS los cruces posteriores usen la misma llave.
    x["CORREO_KEY"]=x["CORREO"].map(email_key)
    # Un mismo picker puede venir en varias filas con correo/código/nombre en
    # formatos distintos. Agrupar después de resolverlo evita repetir FNR/MC
    # en el resumen final.
    x["_IDENTITY_KEY"]=x.apply(
        lambda r: ("email:"+str(r.get("CORREO_KEY",""))) if str(r.get("CORREO_KEY"," ")).strip()
        else ("name:"+person_key(r.get("PICKER","")) if person_key(r.get("PICKER","")) else "row:"+str(r.name)),
        axis=1,
    )
    consolidated=[]
    for _,g in x.groupby("_IDENTITY_KEY",sort=False):
        matched=g["_MASTER_MATCH"].fillna(False).astype(bool) if "_MASTER_MATCH" in g.columns else pd.Series(False,index=g.index)
        row=g.loc[matched].iloc[0].copy() if matched.any() else g.iloc[0].copy()
        for field in ["LINEAS","PEDIDOS"]:
            if field in g.columns:
                row[field]=pd.to_numeric(g[field],errors="coerce").fillna(0).sum()
        for field in ["_MASTER_MATCH","_MANUAL_MATCH","_EXCLUDED_PERSONNEL"]:
            if field in g.columns:
                row[field]=g[field].fillna(False).astype(bool).any()
        consolidated.append(row)
    x=pd.DataFrame(consolidated).drop(columns=["_IDENTITY_KEY"],errors="ignore")
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
        ek=email_key(row.get("CORREO",""))
        override_key=ek or key
        if not override_key: continue
        if key in excluded or ek in excluded:
            x.at[i,"_EXCLUDED_PERSONNEL"]=True
            x.at[i,"_MASTER_MATCH"]=True
            x.at[i,"TURNO"]="__EXCLUIDO__"
            x.at[i,"SUPERVISOR"]="Excluido"
            x.at[i,"AREA_BASE"]="Excluido"
            continue
        ov=overrides.get(override_key) or overrides.get(key)
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
            manual["_EMAIL_KEY"]=manual["CORREO"].map(email_key)
            r=pd.concat([r,manual],ignore_index=True)
    if not r.empty:
        r["_KEY"]=r["PICKER"].map(person_key)
        r["_EMAIL_KEY"]=r["CORREO"].map(email_key)
        r=r[~r["_KEY"].isin(excluded) & ~r["_EMAIL_KEY"].isin(excluded)].copy()
        has_email=r["_EMAIL_KEY"].astype(str).str.strip().ne("")
        r=pd.concat([r[has_email].drop_duplicates("_EMAIL_KEY",keep="last"), r[~has_email].drop_duplicates("_KEY",keep="last")],ignore_index=True)
    return r

def canonicalize_incidents(inc, base):
    """Asocia FNR/MC a la persona canónica usando exclusivamente CORREO.

    El nombre que viene en el archivo operativo es descriptivo; el nombre
    definitivo se toma de la plantilla consolidada mediante el correo.
    """
    if inc is None or inc.empty:
        return inc
    out=inc.copy()
    base_email={}
    for _,r in base.iterrows():
        ek=email_key(r.get("CORREO",""))
        if ek:
            base_email[ek]=r
    result=[]
    matched=[]
    canonical_email=[]
    for _,row in out.iterrows():
        ek=email_key(row.get("CORREO",""))
        rr=base_email.get(ek) if ek else None
        if rr is None:
            rr=_match_master_row(row.get("PICKER",""),base,row.get("CORREO",""))
        if rr is not None:
            result.append(str(rr.get("PICKER", row.get("PICKER",""))).strip())
            canonical_email.append(str(rr.get("CORREO",row.get("CORREO",""))).strip())
            matched.append(True)
        else:
            result.append(str(row.get("PICKER","")).strip())
            canonical_email.append(str(row.get("CORREO","")).strip())
            matched.append(False)
    out["PICKER"]=result
    out["CORREO"]=canonical_email
    out["_EMAIL_MATCH"]=matched
    return out

def template_roster_from_upload(data):
    """Construye la identidad desde la PLANTILLA CONSOLIDADA.

    CORREO_KEY es la llave única. La plantilla puede tener información
    repartida entre Maestro_Personal y Base_Pickers (por ejemplo, un sheet
    tiene correo/nombre y otro tiene turno/supervisor/área). Por eso NO se
    elige una sola fila por correo: se consolidan los campos no vacíos de
    todas las hojas de la plantilla para que el turno nunca se pierda.
    """
    ss=sheets_from_bytes(data) if isinstance(data,(bytes,bytearray)) else sheets(data)
    frames=[]
    for name,df in ss.items():
        n=norm(name)
        if n in {"instrucciones","diagnostico_cruce","diagnostico"}:
            continue
        em=col(df,["codigo_correo","codigo + correo","codigo correo","correo","email","usuario"])
        if not em:
            continue
        p=col(df,["picker","picker_nombre","email_picker","nombre","persona"])
        sh=col(df,["turno","shift"])
        sup=col(df,["supervisor","supervisor_nombre","jefe","responsable"])
        ar=col(df,["area_base","area","departamento","department"])
        estado=col(df,["estado","estatus","status","estatus_cruce"])
        tmp=pd.DataFrame({
            "PICKER":df[p].fillna("").astype(str).str.strip() if p else "",
            "TURNO_MAESTRO":df[sh].fillna("").astype(str).str.strip() if sh else "",
            "CORREO":df[em].fillna("").astype(str).str.strip(),
            "SUPERVISOR":df[sup].fillna("").astype(str).str.strip() if sup else "",
            "AREA_MAESTRO":df[ar].fillna("").astype(str).str.strip() if ar else "",
            "ESTADO_PLANTILLA":df[estado].fillna("").astype(str).str.strip() if estado else "",
        })
        tmp["_EMAIL_KEY"]=tmp["CORREO"].map(email_key)
        tmp=tmp[tmp["_EMAIL_KEY"].ne("")].copy()
        if not tmp.empty:
            tmp["_FUENTE_HOJA"]=str(name)
            frames.append(tmp)

    if not frames:
        raise ValueError("La plantilla consolidada necesita una columna CORREO / CODIGO + CORREO.")

    raw=pd.concat(frames,ignore_index=True)

    def useful(v):
        z=str(v or "").strip()
        return z not in {"", "nan", "None", "NaT", "No asignado", "No especificada", "__EXCLUIDO__"}

    # Consolidar campo por campo por correo. Esto permite que una hoja aporte
    # el turno y otra aporte supervisor/área sin perder ninguno.
    rows=[]
    for ek,g in raw.groupby("_EMAIL_KEY",sort=False):
        row={"_EMAIL_KEY":ek}
        for field in ["PICKER","TURNO_MAESTRO","CORREO","SUPERVISOR","AREA_MAESTRO","ESTADO_PLANTILLA"]:
            vals=[v for v in g[field].tolist() if useful(v)]
            row[field]=str(vals[0]).strip() if vals else ""
        # Preferir el valor que aparezca más veces cuando existan varias fuentes.
        for field in ["PICKER","TURNO_MAESTRO","SUPERVISOR","AREA_MAESTRO","ESTADO_PLANTILLA"]:
            vals=[str(v).strip() for v in g[field].tolist() if useful(v)]
            if vals:
                counts=pd.Series(vals).value_counts()
                row[field]=str(counts.index[0]).strip()
        # El correo canónico siempre sale del valor con @.
        emails=[str(v).strip() for v in g["CORREO"].tolist() if extract_email_address(v)]
        row["CORREO"]=emails[0] if emails else str(g["CORREO"].iloc[0]).strip()
        rows.append(row)

    r=pd.DataFrame(rows)
    # Si la plantilla tiene nombre vacío, usar el nombre del correo solo como
    # descriptor; no se usa para cruzar identidad.
    empty_name=r["PICKER"].astype(str).str.strip().eq("")
    if empty_name.any():
        r.loc[empty_name,"PICKER"]=r.loc[empty_name,"CORREO"].map(email_name_tokens).map(lambda z:" ".join(sorted(z)))

    r["_KEY"]=r["PICKER"].map(person_key)
    r["_TOKEN_KEY"]=r["PICKER"].map(token_key)
    r["_EMAIL_KEY"]=r["CORREO"].map(email_key)
    r["_CODE_KEY"]=r["CORREO"].map(extract_code)
    r["_EMAIL_TOKENS"]=r["CORREO"].map(email_name_tokens)
    return r

def roster_from_uploads(uploads):
    frames=[]
    for turno,up in uploads.items():
        if up:
            frames.append(parse_roster(choose(sheets(up),["turno","personal","plantilla","maestro"]),turno))
    if not frames: return None
    r=pd.concat(frames,ignore_index=True)
    return r.drop_duplicates("PICKER",keep="last")

def parse_inc(df,tipo):
    """Lee FNR/MC con o sin correo.

    Si el detalle no trae correo, se cruza después contra la base/plantilla
    por el nombre y, una vez resuelto, se conserva el correo canónico del Master.
    """
    p=col(df,["picker","picker_nombre","email_picker","nombre","persona"])
    prod=col(df,["product","producto","item","articulo","artículo"])
    order=col(df,["order_number","order","pedido","numero_pedido","order_id"])
    ar=col(df,["department","departamento","area","depto"])
    sh=col(df,["turno","shift"]); fe=col(df,["date","fecha","created_at"])
    qty=col(df,["cantidad","qty","quantity"])
    em=col(df,["codigo_correo","codigo + correo","correo","email","usuario"])
    if not p and not em:
        raise ValueError(f"El detalle {tipo} necesita PICKER/nombre o una columna de CORREO.")
    picker_values=df[p].fillna("").astype(str).str.strip() if p else df[em].fillna("").astype(str).str.strip().map(email_name_tokens).map(lambda z:" ".join(sorted(z)))
    x=pd.DataFrame({
        "PICKER":picker_values,
        "CORREO":df[em].fillna("").astype(str).str.strip() if em else "",
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
    """Adjunta contexto operativo ya resuelto por CORREO.

    El turno y el área que llegan a FNR/MC se copian desde la base, la cual
    previamente fue asociada a la plantilla por correo.
    """
    if inc is None or inc.empty:
        return inc
    base=_dedupe_columns(base)
    ref_cols=[c for c in ["PICKER","LINEAS","PEDIDOS","TURNO","AREA_BASE","SUPERVISOR","CORREO","_EXCLUDED_PERSONNEL"] if c in base.columns]
    ref=base[ref_cols].copy()
    ref["_EMAIL_KEY"]=ref.get("CORREO",pd.Series([""]*len(ref),index=ref.index)).map(email_key)
    ref=ref[ref["_EMAIL_KEY"].astype(str).str.strip().ne("")].drop_duplicates("_EMAIL_KEY").set_index("_EMAIL_KEY")
    y=inc.copy()
    y["_EMAIL_KEY"]=y.get("CORREO",pd.Series([""]*len(y),index=y.index)).map(email_key)
    y["_TEMPLATE_MATCH"]=False
    y["TURNO_REF"]=y.get("TURNO",pd.Series(["No especificado"]*len(y),index=y.index)).astype(str).str.strip()
    y["AREA_REF"]=y.get("AREA",pd.Series(["No especificada"]*len(y),index=y.index)).astype(str).str.strip()
    for i,row in y.iterrows():
        ek=row.get("_EMAIL_KEY","")
        if ek and ek in ref.index:
            rr=ref.loc[ek]
            y.at[i,"PICKER"]=str(rr.get("PICKER",row.get("PICKER",""))).strip()
            y.at[i,"CORREO"]=str(rr.get("CORREO",row.get("CORREO",""))).strip()
            turno=str(rr.get("TURNO","")).strip()
            area=str(rr.get("AREA_BASE","")).strip()
            sup=str(rr.get("SUPERVISOR","")).strip()
            if turno and turno not in {"nan","None","No especificado"}:
                y.at[i,"TURNO"]=turno
                y.at[i,"TURNO_REF"]=turno
            else:
                y.at[i,"TURNO_REF"]=str(row.get("TURNO","No especificado")).strip() or "No especificado"
            if area and area not in {"nan","None","No especificada"}:
                y.at[i,"AREA"]=area
                y.at[i,"AREA_REF"]=area
            if sup:
                y.at[i,"SUPERVISOR_REF"]=sup
            y.at[i,"_TEMPLATE_MATCH"]=True
            if "_EXCLUDED_PERSONNEL" in ref.columns:
                y.at[i,"_EXCLUDED_PERSONNEL"]=bool(rr.get("_EXCLUDED_PERSONNEL",False))
    matched=(y.get("_EMAIL_MATCH",y.get("_TEMPLATE_MATCH",pd.Series(False,index=y.index)))
             .fillna(False).astype(bool))
    y["CATEGORIA"]=np.where(matched,"Picker","Sin registrar")
    y.drop(columns=["_EMAIL_KEY"],errors="ignore",inplace=True)
    return y

def summary(base,fnr,mc):
    keys=["CATEGORIA","PICKER"]
    b=base.copy()
    if "CATEGORIA" not in b:
        matched=b.get("_MASTER_MATCH",pd.Series(False,index=b.index)).fillna(False).astype(bool)
        manual=b.get("_MANUAL_MATCH",pd.Series(False,index=b.index)).fillna(False).astype(bool)
        b["CATEGORIA"]=np.where(matched|manual,"Picker","Sin registrar")
    b["PICKER"]=b.get("PICKER",pd.Series("Sin registrar",index=b.index)).fillna("").astype(str).str.strip()
    b.loc[b["PICKER"].eq(""),"PICKER"]="Sin registrar"
    for field in ["LINEAS","PEDIDOS"]:
        if field not in b: b[field]=0

    # Incluye personal que aparece en FNR/MC aunque no tenga registro en la base
    # de líneas. Su tasa queda N/D cuando no existe denominador real.
    meta=[b]
    for inc in [fnr,mc]:
        if inc is None or inc.empty: continue
        q=inc.copy()
        if "CATEGORIA" not in q:
            matched=q.get("_EMAIL_MATCH",q.get("_TEMPLATE_MATCH",pd.Series(False,index=q.index))).fillna(False).astype(bool)
            q["CATEGORIA"]=np.where(matched,"Picker","Sin registrar")
        q["PICKER"]=q.get("PICKER",pd.Series("Sin registrar",index=q.index)).fillna("").astype(str).str.strip()
        q.loc[q["PICKER"].eq(""),"PICKER"]="Sin registrar"
        q["LINEAS"]=0.0; q["PEDIDOS"]=0.0
        q["AREA_BASE"]=q.get("AREA_BASE",q.get("AREA",pd.Series("No especificada",index=q.index)))
        q["SUPERVISOR"]=q.get("SUPERVISOR",q.get("SUPERVISOR_REF",pd.Series("No asignado",index=q.index)))
        meta.append(q)
    meta_df=pd.concat(meta,ignore_index=True,sort=False)
    for field in ["LINEAS","PEDIDOS"]:
        meta_df[field]=pd.to_numeric(meta_df.get(field,0),errors="coerce").fillna(0)
    metadata_fields=[c for c in ["TURNO","SUPERVISOR","AREA_BASE","CORREO"] if c in meta_df.columns]
    agg={"LINEAS":"sum","PEDIDOS":"sum"}
    agg.update({c:lambda values: next((str(v).strip() for v in values if str(v).strip() and str(v).strip().lower() not in {"nan","none"}),"") for c in metadata_fields})
    for flag in ["_EXCLUDED_PERSONNEL","_MASTER_MATCH","_MANUAL_MATCH"]:
        if flag in meta_df.columns:
            meta_df[flag]=meta_df[flag].astype("boolean").fillna(False).astype(bool)
            agg[flag]="any"
    if "_SOURCE_PICKER" in meta_df.columns:
        agg["_SOURCE_PICKER"]=lambda values: next((str(v).strip() for v in values if str(v).strip() and str(v).strip().lower() not in {"nan","none"}),"")
    s=meta_df.groupby(keys,as_index=False,dropna=False).agg(agg)
    f=(fnr.groupby(keys,as_index=False,dropna=False).INCIDENCIAS.sum().rename(columns={"INCIDENCIAS":"FNR"})
       if fnr is not None and not fnr.empty else pd.DataFrame(columns=keys+["FNR"]))
    m=(mc.groupby(keys,as_index=False,dropna=False).INCIDENCIAS.sum().rename(columns={"INCIDENCIAS":"MC"})
       if mc is not None and not mc.empty else pd.DataFrame(columns=keys+["MC"]))
    s=s.merge(f,on=keys,how="outer").merge(m,on=keys,how="outer")
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



def _dedupe_columns(df):
    """Elimina columnas duplicadas conservando la primera aparición.
    Evita que pandas convierta df["columna"] en DataFrame y rompa filtros/booleanos.
    """
    if df is None:
        return df
    x=df.copy()
    if getattr(x.columns, "duplicated", None) is not None and x.columns.duplicated().any():
        x=x.loc[:, ~x.columns.duplicated(keep="first")].copy()
    return x

def _excluded_mask(df, column="_EXCLUDED_PERSONNEL"):
    """Devuelve una Serie booleana segura aun cuando la columna exista duplicada."""
    if df is None or len(df)==0:
        return pd.Series(dtype=bool, index=getattr(df, "index", None))
    if column not in df.columns:
        return pd.Series(False, index=df.index, dtype=bool)
    raw=df.loc[:, df.columns == column]
    if isinstance(raw, pd.DataFrame):
        raw=raw.apply(lambda c: c.fillna(False).astype(bool))
        return raw.any(axis=1)
    return raw.fillna(False).astype(bool)

def safe_pct(num, den, decimals=2):
    """Porcentaje robusto para pandas/Streamlit Cloud."""
    n=pd.to_numeric(num,errors="coerce")
    d=pd.to_numeric(den,errors="coerce")
    out=n.div(d.where(d.ne(0),float("nan"))).mul(100)
    return out.astype("float64").round(decimals)

def groups(inc,base,key):
    """Agrupa FNR/MC por turno o área usando el contexto ya cruzado por correo.
    Acepta TURNO_REF/TURNO y AREA_REF/AREA para evitar que una diferencia de
    nombre interno deje vacía la segmentación.
    """
    x=_dedupe_columns(inc)
    base=_dedupe_columns(base)
    excluded_mask=_excluded_mask(x)
    if len(excluded_mask)==len(x):
        x=x.loc[~excluded_mask].copy()
    if x.empty:
        return pd.DataFrame(columns=["GRUPO","INCIDENCIAS","% DEL TOTAL","LINEAS","% / LINEAS"])

    if key=="TURNO":
        turn_col=next((c for c in ["TURNO_REF","TURNO","_MASTER_TURNO","TURNO_MAESTRO"] if c in x.columns),None)
        if turn_col is None:
            x["GRUPO"]="No especificado"
        else:
            x["GRUPO"]=x[turn_col].fillna("No especificado").astype(str).str.strip()
            x.loc[x["GRUPO"].eq(""),"GRUPO"]="No especificado"
        den_col="TURNO" if "TURNO" in base.columns else ("_MASTER_TURNO" if "_MASTER_TURNO" in base.columns else None)
        if den_col:
            den=(base.groupby(den_col,as_index=False)["LINEAS"].sum()
                 .rename(columns={den_col:"GRUPO"}))
        else:
            den=None
    else:
        area_col=next((c for c in ["AREA_REF","AREA","AREA_BASE","_MASTER_AREA","AREA_MAESTRO"] if c in x.columns),None)
        if area_col is None:
            x["GRUPO"]="No especificada"
        else:
            x["GRUPO"]=x[area_col].fillna("No especificada").astype(str).str.strip()
            x.loc[x["GRUPO"].eq(""),"GRUPO"]="No especificada"
        den_col="AREA_BASE" if "AREA_BASE" in base.columns else None
        if den_col and (base[den_col].astype(str).str.strip()!="No especificada").any():
            den=(base.groupby(den_col,as_index=False)["LINEAS"].sum()
                 .rename(columns={den_col:"GRUPO"}))
        else:
            den=None

    g=x.groupby("GRUPO",as_index=False)["INCIDENCIAS"].sum()
    total_inc=float(pd.to_numeric(g["INCIDENCIAS"],errors="coerce").fillna(0).sum())
    g["% DEL TOTAL"]=(pd.to_numeric(g["INCIDENCIAS"],errors="coerce").fillna(0)/total_inc*100).round(2) if total_inc else 0.0

    if den is not None:
        g=g.merge(den,on="GRUPO",how="left")
        g["LINEAS"]=pd.to_numeric(g["LINEAS"],errors="coerce").fillna(0.0)
        den_lineas=g["LINEAS"].where(g["LINEAS"].ne(0),float("nan"))
        g["% / LINEAS"]=pd.to_numeric(g["INCIDENCIAS"],errors="coerce").fillna(0).div(den_lineas).mul(100).round(2)
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

def supervisor_email_keys(store):
    keys=set(); names=set()
    for item in store.get("supervisores",[]) or []:
        if not isinstance(item,dict) or not item.get("activo",True): continue
        ek=email_key(item.get("correo","")); nk=person_key(item.get("nombre",""))
        if ek: keys.add(ek)
        if nk: names.add(nk)
    return keys,names

def exclude_registered_supervisors(base, fnr, mc, store):
    email_keys,name_keys=supervisor_email_keys(store)
    if not email_keys and not name_keys: return base,fnr,mc
    def mask(df):
        if df is None or df.empty: return pd.Series(False,index=getattr(df,"index",[]))
        emails=df.get("CORREO",pd.Series("",index=df.index)).map(email_key)
        names=df.get("PICKER",pd.Series("",index=df.index)).map(person_key)
        return emails.isin(email_keys) | (emails.eq("") & names.isin(name_keys))
    return base.loc[~mask(base)].copy(), fnr.loc[~mask(fnr)].copy(), mc.loc[~mask(mc)].copy()

def _secret(name, default=""):
    try:
        return str(st.secrets.get(name, default) or default)
    except Exception:
        return str(os.getenv(name, default) or default)

def send_followup_email(subject, body, recipients, attachment_bytes=None, attachment_name="seguimiento.pdf"):
    recipients=[x.strip() for x in re.split(r"[,;]",str(recipients or "")) if x.strip()]
    host=_secret("SMTP_HOST").strip(); user=_secret("SMTP_USER").strip(); password=_secret("SMTP_PASSWORD")
    sender=_secret("SMTP_FROM",user).strip() or user
    if not recipients: return False,"No se capturaron destinatarios. Escribe al menos un correo."
    invalid=[x for x in recipients if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+",x)]
    if invalid: return False,"Revisa los correos: "+", ".join(invalid)
    if not host or not user or not password:
        return False,"Correo no enviado: configura SMTP_HOST, SMTP_USER, SMTP_PASSWORD y SMTP_PORT en los Secrets del servidor. SMTP_FROM es opcional."
    try:
        port=int(_secret("SMTP_PORT","587") or 587)
        msg=EmailMessage(); msg["Subject"]=subject; msg["From"]=sender; msg["To"] = ", ".join(recipients); msg.set_content(body)
        if attachment_bytes is not None:
            msg.add_attachment(attachment_bytes,maintype="application",subtype="pdf",filename=attachment_name)
        context=ssl.create_default_context()
        if port==465:
            with smtplib.SMTP_SSL(host,port,timeout=30,context=context) as smtp:
                smtp.login(user,password); smtp.send_message(msg)
        else:
            with smtplib.SMTP(host,port,timeout=30) as smtp:
                smtp.ehlo()
                if smtp.has_extn("starttls"):
                    smtp.starttls(context=context); smtp.ehlo()
                smtp.login(user,password); smtp.send_message(msg)
        return True,"Copia enviada correctamente a: "+", ".join(recipients)
    except Exception as exc:
        return False,f"No fue posible enviar la copia. Revisa el servidor y los datos SMTP. Detalle: {exc}"

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
cloud_migrated,cloud_missing=migrate_local_assets_to_cloud(store)
cloud_history_migrated=migrate_local_upload_history_to_cloud()
store.setdefault("excluded_orders", [])
store.setdefault("feedback_rows", [])
store.setdefault("master_overrides", {})
store.setdefault("master_excluded", [])
store.setdefault("procesos", [])
store.setdefault("recursos_formatos", [])
store.setdefault("seguimientos_documentos", [])
store.setdefault("supervisores", [])
store.setdefault("upload_meta", {})

with st.sidebar:
    st.header("Control operativo")
    if cloud_enabled():
        _,_,cloud_bucket=cloud_config()
        st.success(f"☁️ Respaldo en nube activo · {cloud_bucket}")
        if cloud_migrated:
            st.caption(f"Se migraron {cloud_migrated} documentos e imágenes al almacenamiento privado.")
        if cloud_history_migrated:
            st.caption(f"Se migraron {cloud_history_migrated} versiones previas de Excel al historial en nube.")
        if cloud_missing:
            st.warning(f"{cloud_missing} archivos antiguos no estaban disponibles en el servidor para migrarlos.")
    else:
        st.warning("Respaldo en nube pendiente: configura SUPABASE_URL y SUPABASE_SECRET_KEY en los secretos del servidor.")
        st.caption('Agrega también SUPABASE_STORAGE_BUCKET="control-fnr-mc". El bucket privado se crea automáticamente al conectar.')
    ub_upload=st.file_uploader("① Base de Pickers / Líneas",type=["xlsx","xls"],key="base_picker")
    uf_upload=st.file_uploader("② Detalle FNR",type=["xlsx","xls"],key="detalle_fnr")
    um_upload=st.file_uploader("③ Detalle Mala Calidad",type=["xlsx","xls"],key="detalle_mc")
    up_upload=st.file_uploader("④ Plantilla consolidada",type=["xlsx","xls"],key="plantilla_personal")

    # Cada archivo nuevo reemplaza automáticamente al guardado. Si solo se recarga
    # la página, la app recupera la última versión guardada sin pedir volver a subirla.
    ub=persist_upload(ub_upload,"base_picker") if ub_upload else load_persisted_upload("base_picker")
    uf=persist_upload(uf_upload,"detalle_fnr") if uf_upload else load_persisted_upload("detalle_fnr")
    um=persist_upload(um_upload,"detalle_mc") if um_upload else load_persisted_upload("detalle_mc")
    up=persist_upload(up_upload,"plantilla_personal") if up_upload else load_persisted_upload("plantilla_personal")
    for _key,_upload,_label in [("base_picker",ub_upload,"Pickers/Líneas"),("detalle_fnr",uf_upload,"FNR"),("detalle_mc",um_upload,"Mala Calidad"),("plantilla_personal",up_upload,"Plantilla consolidada")]:
        if _upload is not None:
            store.setdefault("upload_meta",{})[_key]={"fecha":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"archivo":str(getattr(_upload,"name",_label))}
    save_store(store)

    status=persisted_status()
    labels={"base_picker":"Pickers/Líneas","detalle_fnr":"FNR","detalle_mc":"Mala Calidad","plantilla_personal":"Master Pickers"}
    guardados=[labels[k] for k,v in status.items() if v]
    if guardados:
        st.success("Archivos guardados: " + ", ".join(guardados))
    with st.expander("🗂️ Historial de archivos Excel",expanded=False):
        _history=upload_history_rows()
        if _history:
            st.dataframe(pd.DataFrame(_history),use_container_width=True,hide_index=True)
            st.download_button("Descargar respaldo del historial (.zip)",upload_history_zip(),"Historial_Excel_FNR_MC.zip","application/zip",key="download_excel_history")
            st.caption("Se conservan hasta 20 versiones por archivo en el almacenamiento local de la app. Descarga el ZIP para guardar una copia fuera de Streamlit.")
        else:
            st.info("Aún no hay versiones en el historial. Se registra una versión al cargar cada Excel.")
    st.caption("Los 3 Excel operativos se cruzan con la PLANTILLA CONSOLIDADA usando CORREO como única llave. La plantilla aporta el nombre asociado, turno, supervisor y área. El nombre no se usa para hacer coincidencias entre archivos.")
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
    st.info("Carga los 3 Excel operativos y la plantilla consolidada para comenzar. Para que el cruce sea por correo, los registros que deban asociarse deben traer CORREO / CODIGO + CORREO.")
    st.markdown("**Fuentes:** ① Pickers/Líneas · ② FNR · ③ Mala Calidad · ④ Plantilla consolidada (correo → nombre asociado, turno, supervisor y área).")
    st.stop()

try:
    base=parse_base(choose(sheets_from_bytes(ub.getvalue()),["picker","lineas","resumen"]))
    fnr=parse_inc(choose(sheets_from_bytes(uf.getvalue()),["fnr","detalle"]),"FNR")
    mc=parse_inc(choose(sheets_from_bytes(um.getvalue()),["mc","mala"]),"MC")
    # LA PLANTILLA CONSOLIDADA es la única fuente de identidad.
    # No se usa un Excel maestro externo.
    roster=template_roster_from_upload(up.getvalue())
    if roster.empty:
        raise ValueError("La plantilla consolidada no contiene correos válidos.")
    base=apply_roster(base,roster)
    base=apply_manual_personnel(base,store)
    _base_match=(base.get("_MASTER_MATCH",pd.Series(False,index=base.index)).fillna(False).astype(bool)
                 | base.get("_MANUAL_MATCH",pd.Series(False,index=base.index)).fillna(False).astype(bool))
    base["CATEGORIA"]=np.where(_base_match,"Picker","Sin registrar")
    base=_dedupe_columns(base)
    roster=effective_roster(roster,store)
    # FNR/MC se identifican directamente contra la PLANTILLA por correo.
    # NO dependen de que el picker exista primero en la base de líneas.
    fnr=canonicalize_incidents(fnr,roster)
    mc=canonicalize_incidents(mc,roster)
except Exception as e:
    st.error(f"Error en los archivos cargados: {e}"); st.stop()

if excluded:
    fnr=fnr[~fnr.ORDER_NUMBER.isin(excluded)].copy()
    mc=mc[~mc.ORDER_NUMBER.isin(excluded)].copy()
# Para FNR/MC, el contexto de turno/área/supervisor viene de la plantilla por correo.
identity_for_attach=roster.rename(columns={"TURNO_MAESTRO":"TURNO","AREA_MAESTRO":"AREA_BASE"}).copy()
fnr=attach(fnr,identity_for_attach); mc=attach(mc,identity_for_attach)
fnr["CORREO_KEY"]=fnr.get("CORREO",pd.Series("",index=fnr.index)).map(email_key)
mc["CORREO_KEY"]=mc.get("CORREO",pd.Series("",index=mc.index)).map(email_key)
fnr=_dedupe_columns(fnr); mc=_dedupe_columns(mc); base=_dedupe_columns(base)
base,fnr,mc=exclude_registered_supervisors(base,fnr,mc,store)
s=summary(base,fnr,mc)
# Estado de cruce de TODOS los archivos: Pickers/Líneas + FNR + MC.
# La llave es CORREO_KEY; cualquier registro sin coincidencia con la plantilla
# se concentra en la pestaña 🧩 Correos / Cruce para asignación manual.
def _cross_frame(df, fuente):
    if df is None or df.empty:
        return pd.DataFrame(columns=["FUENTE","CATEGORIA","PICKER","CORREO","CORREO_KEY","TURNO","SUPERVISOR","AREA_BASE","IDENTIFICADO"])
    y=df.copy()
    y["CORREO_KEY"]=y.get("CORREO",pd.Series("",index=y.index)).map(email_key)
    if fuente=="Pickers / Líneas":
        ident=y.get("_MASTER_MATCH",False)
        area=y.get("AREA_BASE",pd.Series("",index=y.index))
    else:
        ident=y.get("_TEMPLATE_MATCH",y.get("_EMAIL_MATCH",False))
        area=y.get("AREA_REF",y.get("AREA",pd.Series("",index=y.index)))
    out=pd.DataFrame({
        "FUENTE":fuente,
        "CATEGORIA":y.get("CATEGORIA",pd.Series("Picker",index=y.index)).astype(str).str.strip(),
        "PICKER":y.get("PICKER",pd.Series("",index=y.index)).astype(str).str.strip(),
        "CORREO":y.get("CORREO",pd.Series("",index=y.index)).astype(str).str.strip(),
        "CORREO_KEY":y["CORREO_KEY"],
        "TURNO":y.get("TURNO",y.get("TURNO_REF",pd.Series("No especificado",index=y.index))).astype(str).str.strip(),
        "SUPERVISOR":y.get("SUPERVISOR",y.get("SUPERVISOR_REF",pd.Series("No asignado",index=y.index))).astype(str).str.strip(),
        "AREA_BASE":area.astype(str).str.strip(),
        "IDENTIFICADO":pd.Series(ident,index=y.index).fillna(False).astype(bool),
    })
    return out

cross_status=pd.concat([
    _cross_frame(base,"Pickers / Líneas"),
    _cross_frame(fnr,"FNR"),
    _cross_frame(mc,"Mala Calidad"),
],ignore_index=True)
master_match=base.get("_MASTER_MATCH",pd.Series(False,index=base.index)).copy()
base_display=base.drop(columns=["_MASTER_MATCH"],errors="ignore")
s=s.drop(columns=["_MASTER_MATCH"],errors="ignore")
# Registrar automáticamente pickers vistos en la operación, sin alterar su estado histórico.
for _idx,_row in base.iterrows():
    _p=str(_row.get("PICKER","")).strip()
    _src=str(_row.get("_SOURCE_PICKER","")).strip()
    if _p: picker_record(store, _p, aliases=[_src] if _src else [])
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
            st.caption(f"{unmatched} registro(s) en la categoría Sin registrar.")
        email_match=int(base.get("_MASTER_MATCH",pd.Series(False,index=base.index)).fillna(False).astype(bool).sum())
        st.caption(f"🔑 Cruce de identidad: CORREO primero · {email_match} registros identificados")

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

def _context_identity_view(roster,base=None,fnr=None,mc=None):
    """Universo de contexto: plantilla consolidada más personal Sin registrar.

    El correo identifica al personal registrado; el nombre normalizado identifica
    por separado a quienes no están en la plantilla.
    """
    if roster is None or roster.empty:
        r=pd.DataFrame(columns=["PICKER","TURNO","SUPERVISOR","AREA_BASE","CORREO","CORREO_KEY","_CONTEXT_KEY","CATEGORIA"])
    else:
        r=roster.copy()
    r["PICKER"]=r.get("PICKER",pd.Series("",index=r.index)).fillna("").astype(str).str.strip()
    r["TURNO"]=r.get("TURNO_MAESTRO",pd.Series("",index=r.index)).fillna("").astype(str).str.strip()
    r["SUPERVISOR"]=r.get("SUPERVISOR",pd.Series("",index=r.index)).fillna("").astype(str).str.strip()
    r["AREA_BASE"]=r.get("AREA_MAESTRO",pd.Series("",index=r.index)).fillna("").astype(str).str.strip()
    r["CORREO"]=r.get("CORREO",pd.Series("",index=r.index)).fillna("").astype(str).str.strip()
    r["CORREO_KEY"]=r["CORREO"].map(email_key)
    r=r[r["CORREO_KEY"].ne("")].copy()
    r.loc[r["TURNO"].eq(""),"TURNO"]="No especificado"
    r.loc[r["SUPERVISOR"].eq(""),"SUPERVISOR"]="No asignado"
    r.loc[r["AREA_BASE"].eq(""),"AREA_BASE"]="No especificada"
    r["CATEGORIA"]="Picker"
    r["_CONTEXT_KEY"]=r["CORREO_KEY"].map(lambda z:"correo:"+str(z))
    r=r[r["CORREO_KEY"].ne("")].copy()
    extra=[]
    for source in [base,fnr,mc]:
        if source is None or source.empty: continue
        q=source.copy()
        if "CATEGORIA" not in q: continue
        q=q[q["CATEGORIA"].astype(str).eq("Sin registrar")]
        for _,row in q.iterrows():
            name=str(row.get("PICKER","") or "").strip() or "Sin registrar"
            extra.append({
                "PICKER":name,
                "TURNO":str(row.get("TURNO",row.get("TURNO_REF","No especificado")) or "No especificado").strip(),
                "SUPERVISOR":str(row.get("SUPERVISOR",row.get("SUPERVISOR_REF","No asignado")) or "No asignado").strip(),
                "AREA_BASE":str(row.get("AREA_BASE",row.get("AREA_REF",row.get("AREA","No especificada"))) or "No especificada").strip(),
                "CORREO":"","CORREO_KEY":"","CATEGORIA":"Sin registrar",
                "_CONTEXT_KEY":"sin-registrar:"+(person_key(name) or "sin-nombre"),
            })
    if extra:
        r=pd.concat([r,pd.DataFrame(extra)],ignore_index=True,sort=False)
    return r.drop_duplicates("_CONTEXT_KEY",keep="last")

def _records_for_context(df, ctx, identity_df, include_turn=True):
    """Filtra cualquier dataset operativo usando los CORREO_KEY seleccionados en plantilla."""
    if df is None or df.empty:
        return df.copy() if isinstance(df,pd.DataFrame) else pd.DataFrame()
    ids=apply_person_filters(
        identity_df,
        "Todos",
        ctx.get("turno","Todos") if include_turn else "Todos",
        ctx.get("supervisor","Todos"),
        ctx.get("area","Todos"),
    )
    out=df.copy()
    ids=ids.copy()
    if "_CONTEXT_KEY" not in ids:
        ids["_CONTEXT_KEY"]=ids.apply(person_context_key,axis=1)
    keys=set(ids["_CONTEXT_KEY"].astype(str))
    out["_CONTEXT_KEY"]=out.apply(person_context_key,axis=1)
    return out[out["_CONTEXT_KEY"].isin(keys)].copy()

def apply_context(df, ctx, include_turn=True, identity_df=None):
    if identity_df is not None:
        return _records_for_context(df,ctx,identity_df,include_turn)
    return apply_person_filters(
        df,"Todos",ctx.get("turno","Todos") if include_turn else "Todos",
        ctx.get("supervisor","Todos"),ctx.get("area","Todos"))

def select_summary_people(df,summary_df):
    """Filtra por categoría y nombre para conservar aparte Sin registrar."""
    if df is None or df.empty or summary_df is None or summary_df.empty:
        return df.iloc[0:0].copy() if isinstance(df,pd.DataFrame) else pd.DataFrame()
    if "CATEGORIA" not in df.columns or "CATEGORIA" not in summary_df.columns:
        return df[df["PICKER"].isin(summary_df["PICKER"])].copy()
    allowed=set(zip(summary_df["CATEGORIA"].astype(str),summary_df["PICKER"].astype(str)))
    mask=df.apply(lambda r:(str(r.get("CATEGORIA","Picker")),str(r.get("PICKER",""))) in allowed,axis=1)
    return df.loc[mask].copy()

def context_reference(df, ctx, identity_df=None):
    """Base de comparación: supervisor/área de la plantilla, todos los turnos."""
    if identity_df is not None:
        return _records_for_context(df,ctx,identity_df,False)
    return apply_person_filters(df,"Todos","Todos",ctx.get("supervisor","Todos"),ctx.get("area","Todos"))

def render_context_banner(ctx, selected_df, reference_df, label="Contexto de análisis"):
    parts=[]
    for k,lab in [("turno","Turno"),("supervisor","Supervisor"),("area","Área")]:
        v=ctx.get(k,"Todos")
        if v!="Todos": parts.append(f"{lab}: {v}")
    scope=" · ".join(parts) if parts else "Todos los turnos"
    sel_p=int(selected_df.get("CATEGORIA",pd.Series("Picker",index=selected_df.index)).astype(str).eq("Picker").sum())
    sel_u=int(selected_df.get("CATEGORIA",pd.Series("Picker",index=selected_df.index)).astype(str).eq("Sin registrar").sum())
    ref_p=int(reference_df.get("CATEGORIA",pd.Series("Picker",index=reference_df.index)).astype(str).eq("Picker").sum())
    pct=(sel_p/ref_p*100) if ref_p else 0
    count_text=f"{sel_p:,} pickers · {sel_u:,} sin registrar" if sel_u else f"{sel_p:,} pickers"
    st.markdown(
        f"<div class='context-banner'><div class='context-title'>🎯 {label}: {scope}</div>"
        f"<div class='context-detail'>{count_text} · {pct:.1f}% del universo de comparación · Las demás pestañas utilizan este mismo contexto.</div></div>",
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

s_view=_dedupe_columns(s)
excluded_mask=_excluded_mask(s_view)
if len(excluded_mask)==len(s_view):
    s_view=s_view.loc[~excluded_mask].copy()

# El CONTEXTO GLOBAL sale de la PLANTILLA CONSOLIDADA, no del nombre del Excel operativo.
# Cada pestaña después filtra sus propios registros por CORREO_KEY.
context_identity=_context_identity_view(roster,base,fnr,mc)

# Defaults para pestañas que no necesitan filtros globales.
sp="Todos"
stn="Todos"
ssup="Todos"
sar="Todos"
base_view=select_summary_people(base,s_view)
fnr_view=select_summary_people(fnr,s_view)
mc_view=select_summary_people(mc,s_view)

a,b,e,x,g,h,f,j=st.tabs(["🏠 Bodega","👤 Picker","🌙 Turnos / Áreas","🧩 Correos / Cruce","👥 Supervisores","🛡️ Seguimiento","📥 Exportar","📌 Pendientes & Procesos"])

with a:
    st.subheader(f"Resumen de bodega — {periodo}")
    st.caption("El contexto elegido aquí se comparte automáticamente con todas las pestañas.")
    _meta=store.get("upload_meta",{}) or {}
    _upload_rows=[]
    for _k,_label in [("base_picker","Pickers / Líneas"),("detalle_fnr","FNR"),("detalle_mc","Mala Calidad"),("plantilla_personal","Plantilla consolidada")]:
        _m=_meta.get(_k,{}) or {}
        _upload_rows.append({"Archivo":_label,"Última carga":_m.get("fecha","Sin registro"),"Nombre":_m.get("archivo","")})
    with st.expander("🕒 Última carga de Excel",expanded=True):
        st.dataframe(pd.DataFrame(_upload_rows),use_container_width=True,hide_index=True)
    ctx,turns,sups,areas=global_context(context_identity)
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

    # Mantener el flujo de la versión anterior que sí cargaba el contexto:
    # primero filtra el resumen canónico desde plantilla y después aplica
    # esos pickers a base, FNR y MC.
    s_bodega=apply_context(s_view,ctx,identity_df=context_identity)
    s_reference=context_reference(s_view,ctx,identity_df=context_identity)
    base_bodega=select_summary_people(base,s_bodega)
    fnr_bodega=select_summary_people(fnr,s_bodega)
    mc_bodega=select_summary_people(mc,s_bodega)
    base_reference=select_summary_people(base,s_reference)
    fnr_reference=select_summary_people(fnr,s_reference)
    mc_reference=select_summary_people(mc,s_reference)
    render_context_banner(ctx,s_bodega,s_reference)

    lines=base_bodega.LINEAS.sum(); F=fnr_bodega.INCIDENCIAS.sum(); M=mc_bodega.INCIDENCIAS.sum()
    total_pedidos,fnr_pedidos,mc_pedidos,fnr_rate,mc_rate=monthly_kpis(base_bodega,fnr_bodega,mc_bodega)
    _picker_count=int(s_bodega.get("CATEGORIA",pd.Series("Picker",index=s_bodega.index)).astype(str).eq("Picker").sum())
    _unregistered_count=int(s_bodega.get("CATEGORIA",pd.Series("Picker",index=s_bodega.index)).astype(str).eq("Sin registrar").sum())
    _reference_picker_count=int(s_reference.get("CATEGORIA",pd.Series("Picker",index=s_reference.index)).astype(str).eq("Picker").sum())
    render_soft_kpis([
        ("Líneas",f"{lines:,.0f}",f"{lines/base_reference.LINEAS.sum()*100:.1f}% del universo" if base_reference.LINEAS.sum() else "Sin referencia","blue"),
        ("Pedidos",f"{total_pedidos:,.0f}",f"{total_pedidos/base_reference.PEDIDOS.sum()*100:.1f}% del universo" if base_reference.PEDIDOS.sum() else "Sin referencia","green"),
        ("Pickers",f"{_picker_count:,}",f"{_picker_count/_reference_picker_count*100:.1f}% del universo · {_unregistered_count:,} sin registrar" if _reference_picker_count else f"{_unregistered_count:,} sin registrar","blue"),
    ])
    render_soft_kpis([
        ("FNR mensual",f"{fnr_rate:.2f}%" if fnr_rate is not None else "N/D","Objetivo < 1.50%","red"),
        ("MC mensual",f"{mc_rate:.2f}%" if mc_rate is not None else "N/D","Objetivo < 1.00%","amber"),
        ("Fuera objetivo",f"{int(((s_bodega.ESTADO=="🔴 FUERA DE OBJETIVO") & s_bodega.CATEGORIA.eq("Picker")).sum()):,}",f"de {_picker_count:,} pickers · {_unregistered_count:,} sin registrar","red"),
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

    st.divider()
    st.subheader("🏷️ Artículos")
    st.caption("Artículos con incidencia dentro del mismo turno, supervisor y área seleccionados arriba.")
    tipo_articulos=st.radio("Tipo de incidencia",["FNR","MC"],horizontal=True,key="articulos_tipo")
    datos_articulos=fnr_bodega if tipo_articulos=="FNR" else mc_bodega
    articulos_view=(datos_articulos.groupby(["PRODUCTO","AREA"],as_index=False).INCIDENCIAS.sum()
                    .rename(columns={"INCIDENCIAS":"CANTIDAD"}).sort_values("CANTIDAD",ascending=False))
    st.dataframe(articulos_view,use_container_width=True,hide_index=True)
    st.caption(f"{len(articulos_view):,} artículos con incidencia · {int(articulos_view.CANTIDAD.sum()) if not articulos_view.empty else 0:,} incidencias dentro del contexto seleccionado.")

    st.divider()
    st.subheader("📦 Pedidos")
    st.caption("Pedidos con incidencia dentro del mismo turno, supervisor y área seleccionados arriba.")
    tipo_pedidos=st.radio("Tipo de incidencia",["FNR","MC"],horizontal=True,key="pedidos_tipo")
    datos_pedidos=fnr_bodega if tipo_pedidos=="FNR" else mc_bodega
    pedidos_bodega=orders(datos_pedidos)
    st.dataframe(pedidos_bodega,use_container_width=True,hide_index=True)
    st.caption("PICKERS indica cuántos pickers aparecen en el mismo pedido.")

with b:
    ctx,_,_,_=global_context(context_identity)
    selected_context=apply_context(s_view,ctx,identity_df=context_identity)
    st.subheader("👤 Ficha de picker")
    st.caption("El turno, supervisor y área se heredan del contexto elegido en Bodega. Aquí solo buscas el picker.")
    render_context_banner(ctx,selected_context,context_reference(s_view,ctx,identity_df=context_identity),"Contexto heredado")
    names=person_options(selected_context[selected_context.get("CATEGORIA",pd.Series("Picker",index=selected_context.index)).astype(str).eq("Picker")])
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

        st.subheader("⚠️ Incidencias del picker")
        st.caption("Cada incidencia está identificada explícitamente como FNR o Mala Calidad (MC).")
        inc_fnr = fnr[fnr.PICKER == sp].copy()
        inc_mc = mc[mc.PICKER == sp].copy()
        total_fnr_inc = int(pd.to_numeric(inc_fnr["INCIDENCIAS"], errors="coerce").fillna(0).sum()) if not inc_fnr.empty else 0
        total_mc_inc = int(pd.to_numeric(inc_mc["INCIDENCIAS"], errors="coerce").fillna(0).sum()) if not inc_mc.empty else 0
        ic1, ic2 = st.columns(2)
        with ic1:
            st.markdown("<div class='justo-card' style='border-left:4px solid #d64545'><div class='justo-kicker'>FNR · Faltante no reportado</div><div class='justo-title' style='font-size:1.25rem'>%s incidencias</div></div>" % f"{total_fnr_inc:,}", unsafe_allow_html=True)
        with ic2:
            st.markdown("<div class='justo-card' style='border-left:4px solid #d6a72c'><div class='justo-kicker'>MC · Mala Calidad</div><div class='justo-title' style='font-size:1.25rem'>%s incidencias</div></div>" % f"{total_mc_inc:,}", unsafe_allow_html=True)
        cc=st.columns(2)
        with cc[0]:
            st.markdown("**🔴 FNR · Faltantes no reportados**")
            fprod=products(fnr,sp).head(15).copy()
            if not fprod.empty:
                fprod.insert(0,"TIPO","FNR")
            st.dataframe(fprod,use_container_width=True,hide_index=True)
        with cc[1]:
            st.markdown("**🟡 MC · Mala Calidad**")
            mprod=products(mc,sp).head(15).copy()
            if not mprod.empty:
                mprod.insert(0,"TIPO","MC")
            st.dataframe(mprod,use_container_width=True,hide_index=True)
        st.subheader("📦 Pedidos con incidencia")
        st.caption("El tipo de incidencia se muestra para distinguir FNR de MC.")
        of=orders(fnr,sp).head(25).copy()
        om=orders(mc,sp).head(25).copy()
        if not of.empty: of.insert(0,"TIPO","FNR")
        if not om.empty: om.insert(0,"TIPO","MC")
        pedidos_incidencias=pd.concat([of,om],ignore_index=True)
        if not pedidos_incidencias.empty:
            pedidos_incidencias=pedidos_incidencias.sort_values(["TIPO","INCIDENCIAS"],ascending=[True,False])
        st.dataframe(pedidos_incidencias,use_container_width=True,hide_index=True)
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

        _src_alias=[]
        try:
            _src_alias=[str(s[s.PICKER==sp].iloc[0].get("_SOURCE_PICKER","")).strip()] if not s[s.PICKER==sp].empty else []
        except Exception:
            _src_alias=[]
        rec=picker_record(store,sp,aliases=_src_alias)
        st.subheader("📝 Retroalimentación y seguimiento")
        st.caption("Solo se solicita el tipo de seguimiento. La retroalimentación opcional queda dentro del mismo registro.")
        with st.form(f"accion_picker_form_{sp}", clear_on_submit=True):
            act_type=st.selectbox("Tipo de seguimiento",["Llamada de atención","Advertencia verbal 1","Advertencia verbal 2","Acta 1","Acta 2","Acta 3","Cero tolerancia"])
            act_sup=st.text_input("Supervisor",value=str(r.get("SUPERVISOR","")))
            act_motivo=st.text_area("Motivo / detalle")
            act_retro=st.text_area("Retroalimentación (opcional)",placeholder="Observación de retroalimentación, si aplica…")
            if st.form_submit_button("Guardar seguimiento", type="primary"):
                rec.setdefault("acciones",[])
                rec["acciones"].append({"id":datetime.now().strftime("%Y%m%d%H%M%S%f"),"fecha":datetime.now().strftime("%Y-%m-%d %H:%M"),"accion":act_type,"supervisor":act_sup.strip() or "No especificado","motivo":act_motivo.strip(),"retroalimentacion":act_retro.strip()})
                save_store(store); st.success("Seguimiento guardado."); st.rerun()

        st.markdown("### 🔗 Formatos y recursos")
        recursos=store.get("recursos_formatos",[])
        with st.expander(f"Ver formatos y enlaces ({len(recursos)})",expanded=False):
            if recursos:
                for idx,recurso in enumerate(recursos):
                    tipo=str(recurso.get("tipo","Recurso")); titulo=str(recurso.get("titulo","Formato")); url=str(recurso.get("url",""))
                    rr1,rr2,rr3=st.columns([1,4,1])
                    rr1.caption(tipo); rr2.markdown(f"**{titulo}**")
                    with rr3:
                        if url: st.link_button("Abrir",url,use_container_width=True)
                    if st.button("Quitar",key=f"delete_recurso_{sp}_{idx}"): store["recursos_formatos"].pop(idx); save_store(store); st.rerun()
            else: st.info("No hay formatos configurados.")
            with st.form(f"recurso_form_{sp}", clear_on_submit=True):
                rr1,rr2=st.columns([1,2])
                with rr1: recurso_tipo=st.selectbox("Tipo",["Retroalimentación","Seguimiento","Llamada de atención","Acta","Otro"])
                with rr2: recurso_titulo=st.text_input("Nombre del formato",placeholder="Ej. Formato de seguimiento")
                recurso_url=st.text_input("Enlace",placeholder="https://...")
                if st.form_submit_button("➕ Guardar enlace",type="primary"):
                    url=recurso_url.strip()
                    if not recurso_titulo.strip() or not url.startswith(("http://","https://")): st.error("Captura un nombre y un enlace válido.")
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
            st.markdown("**Seguimientos / actas**")
            _ah=pd.DataFrame(rec["acciones"]).sort_values("fecha",ascending=False).rename(columns={"fecha":"Fecha","accion":"Tipo","supervisor":"Supervisor","motivo":"Motivo","retroalimentacion":"Retroalimentación"})
            _cols=[c for c in ["Fecha","Tipo","Supervisor","Motivo","Retroalimentación"] if c in _ah.columns]
            st.dataframe(_ah[_cols],use_container_width=True,hide_index=True)

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

with e:
    ctx,_,_,_=global_context(context_identity)
    st.subheader("🌙 Turnos / Áreas")
    st.caption("El contexto seleccionado se mantiene, pero el comparativo conserva todos los turnos para no perder proporciones.")
    selected=apply_context(s_view,ctx,identity_df=context_identity)
    reference=context_reference(s_view,ctx,identity_df=context_identity)
    render_context_banner(ctx,selected,reference,"Contexto heredado")
    bsel=select_summary_people(base,selected)
    fsel=select_summary_people(fnr,selected)
    msel=select_summary_people(mc,selected)
    bref=select_summary_people(base,reference)
    fref=select_summary_people(fnr,reference)
    mref=select_summary_people(mc,reference)
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

with x:
    st.subheader("🧩 Cruce por correo y personal no asignado")
    st.caption("El cruce intenta primero CORREO y, si falta o no coincide, usa una coincidencia conservadora por nombre. El personal que no aparece en la plantilla se conserva en resultados como Sin registrar.")
    cross=cross_status.copy()
    if not cross.empty:
        cross["MANUAL"]=cross["CORREO_KEY"].map(lambda z: bool(store.get("master_overrides",{}).get(str(z))))
        cross["ID_PERSONA"]=cross.apply(lambda r: str(r.get("CORREO_KEY","")).strip() or ("nombre:"+person_key(r.get("PICKER",""))),axis=1)
    unmatched_cross=cross[~cross["IDENTIFICADO"]].copy() if not cross.empty else pd.DataFrame()
    k1,k2,k3=st.columns(3)
    k1.metric("Personas en los 3 archivos",f"{cross['ID_PERSONA'].nunique():,}" if not cross.empty else "0")
    k2.metric("Registros sin empatar",f"{len(unmatched_cross):,}" if not unmatched_cross.empty else "0")
    k3.metric("Asignaciones manuales",f"{len(store.get('master_overrides',{})):,}")
    if unmatched_cross.empty:
        st.success("✅ Todos los registros de Pickers, FNR y Mala Calidad están asociados a la plantilla.")
    else:
        view_cols=[c for c in ["FUENTE","PICKER","CORREO","TURNO","SUPERVISOR","AREA_BASE"] if c in unmatched_cross.columns]
        st.dataframe(unmatched_cross[view_cols].drop_duplicates(["FUENTE","CORREO"]),use_container_width=True,hide_index=True)
        emails=sorted([str(v) for v in unmatched_cross["CORREO_KEY"].dropna().unique() if str(v).strip()])
        if emails:
          with st.expander("➕ Asignar correo manualmente",expanded=True):
            with st.form("manual_email_assignment_form",clear_on_submit=True):
                correo_sel=st.selectbox("Correo sin asignar",emails)
                row_opts=unmatched_cross[unmatched_cross["CORREO_KEY"]==correo_sel]
                suggested_name=str(row_opts.iloc[0].get("PICKER","")) if not row_opts.empty else ""
                picker_manual=st.text_input("Nombre canónico del picker",value=suggested_name)
                mc1,mc2,mc3=st.columns(3)
                with mc1: turno_manual=st.selectbox("Turno",["Matutino","Intermedio","Tarde","Nocturno","Turno de avance","Otro"])
                with mc2:
                    sup_options=[str(v) for v in sorted(s_view["SUPERVISOR"].dropna().unique()) if str(v).strip() and str(v)!="No asignado"]
                    sup_manual=st.selectbox("Supervisor",["No asignado"]+sup_options)
                with mc3:
                    area_options=[str(v) for v in sorted(s_view["AREA_BASE"].dropna().unique()) if str(v).strip() and str(v)!="No especificada"]
                    area_manual=st.selectbox("Área",["No especificada"]+area_options)
                if st.form_submit_button("💾 Guardar asignación",type="primary"):
                    if not correo_sel or not picker_manual.strip(): st.error("Correo y nombre son obligatorios.")
                    else:
                        store.setdefault("master_overrides",{})[correo_sel]={"PICKER":picker_manual.strip(),"CORREO":correo_sel,"TURNO":turno_manual,"SUPERVISOR":sup_manual,"AREA_BASE":area_manual,"FECHA":datetime.now().strftime("%Y-%m-%d %H:%M"),"ORIGEN":"Manual por correo"}
                        save_store(store); st.success("Asignación guardada por correo. El cruce la utilizará en la siguiente carga."); st.rerun()
        else:
            st.info("Estos registros no traen correo. Revisa que sus nombres coincidan con la plantilla; el cruce automático por nombre ya se intentó.")
    with st.expander("📋 Asignaciones manuales guardadas",expanded=False):
        manual_rows=[]
        for ek,ov in store.get("master_overrides",{}).items(): manual_rows.append({"CORREO":ov.get("CORREO",ek),"PICKER":ov.get("PICKER",""),"TURNO":ov.get("TURNO",""),"SUPERVISOR":ov.get("SUPERVISOR",""),"AREA":ov.get("AREA_BASE",""),"FECHA":ov.get("FECHA","")})
        if manual_rows: st.dataframe(pd.DataFrame(manual_rows),use_container_width=True,hide_index=True)
        else: st.info("Todavía no hay asignaciones manuales.")

with g:
    ctx,_,_,_=global_context(context_identity)
    st.subheader("👥 Supervisores")
    st.caption("Vista operativa bajo el mismo contexto global. Los supervisores registrados por correo se excluyen automáticamente de incidencias y filtros de pickers.")
    with st.expander("➕ Dar de alta supervisor",expanded=False):
        with st.form("alta_supervisor_form",clear_on_submit=True):
            sc1,sc2=st.columns(2)
            with sc1: sup_nombre_nuevo=st.text_input("Nombre del supervisor")
            with sc2: sup_correo_nuevo=st.text_input("Correo del supervisor")
            sup_activo=st.checkbox("Excluirlo de incidencias y filtros",value=True)
            if st.form_submit_button("💾 Guardar supervisor",type="primary"):
                ek=email_key(sup_correo_nuevo)
                if not sup_nombre_nuevo.strip() or not ek: st.error("Captura nombre y un correo válido.")
                else:
                    current=[z for z in store.get("supervisores",[]) if email_key(z.get("correo",""))!=ek]
                    current.append({"nombre":sup_nombre_nuevo.strip(),"correo":ek,"activo":sup_activo,"fecha":datetime.now().strftime("%Y-%m-%d %H:%M")})
                    store["supervisores"]=current; save_store(store); st.success("Supervisor guardado."); st.rerun()
    _supreg=store.get("supervisores",[]) or []
    if _supreg:
        st.dataframe(pd.DataFrame([{"Supervisor":z.get("nombre",""),"Correo":z.get("correo",""),"Excluido":"Sí" if z.get("activo",True) else "No","Alta":z.get("fecha","")} for z in _supreg]),use_container_width=True,hide_index=True)
    sup_data=apply_context(s_view,ctx,identity_df=context_identity)
    reference=context_reference(s_view,ctx,identity_df=context_identity)
    render_context_banner(ctx,sup_data,reference,"Contexto heredado")
    if not sup_data.empty:
        sup_summary=(sup_data.groupby("SUPERVISOR",as_index=False)
            .agg(PICKERS=("PICKER","nunique"),LINEAS=("LINEAS","sum"),FNR=("FNR","sum"),MC=("MC","sum"))
            .sort_values("FNR",ascending=False))
        sup_summary["FNR %"]=safe_pct(sup_summary["FNR"],sup_summary["LINEAS"])
        sup_summary["MC %"]=safe_pct(sup_summary["MC"],sup_summary["LINEAS"])
        st.dataframe(sup_summary,use_container_width=True,hide_index=True)
        sup_focus=st.selectbox("Supervisor",["Todos"]+sorted([str(x) for x in sup_summary["SUPERVISOR"] if str(x).strip()]))
        if sup_focus!="Todos":
            st.dataframe(sup_data[sup_data["SUPERVISOR"]==sup_focus],use_container_width=True,hide_index=True)
    else:
        st.info("No hay datos para el contexto actual.")

with h:
    st.subheader("🛡️ Seguimiento")
    st.caption("Vista consolidada de todos los pickers: retroalimentaciones, seguimientos, actas y tolerancias. Filtra y ordena para ver rápidamente dónde hay más seguimiento.")
    ctx,_,_,_=global_context(context_identity)
    selected_context=apply_context(s_view,ctx,identity_df=context_identity)
    selected_context=selected_context[selected_context.get("CATEGORIA",pd.Series("Picker",index=selected_context.index)).astype(str).eq("Picker")].copy()
    render_context_banner(ctx,selected_context,context_reference(s_view,ctx,identity_df=context_identity),"Contexto heredado")

    all_rows=[]
    for _,rr in selected_context.iterrows():
        picker=str(rr.get("PICKER","")); rec0=picker_record(store,picker,aliases=[str(rr.get("_SOURCE_PICKER",""))])
        acciones=rec0.get("acciones",[]) or []
        retro_legacy=[z for z in store.get("feedback_rows",[]) if str(z.get("PICKER",""))==picker]
        retro_new=[z for z in acciones if str(z.get("retroalimentacion","")).strip()]
        tipos=[str(z.get("accion","")) for z in acciones]
        actas=sum(1 for t in tipos if "Acta" in t); llamadas=sum(1 for t in tipos if "Llamada" in t or "Advertencia" in t); cero=sum(1 for t in tipos if "Cero tolerancia" in t)
        all_rows.append({"PICKER":picker,"TURNO":rr.get("TURNO",""),"SUPERVISOR":rr.get("SUPERVISOR",""),"AREA":rr.get("AREA_BASE",""),"RETROALIMENTACIONES":len(retro_legacy)+len(retro_new),"SEGUIMIENTOS":len(acciones),"ACTAS":actas,"LLAMADAS / ADVERTENCIAS":llamadas,"CERO TOLERANCIA":cero,"TOTAL":len(acciones)+len(retro_legacy),"ÚLTIMO SEGUIMIENTO":max([str(z.get("fecha","")) for z in acciones],default="")})
    seguimiento_df=pd.DataFrame(all_rows)
    if seguimiento_df.empty: st.info("No hay pickers disponibles para seguimiento.")
    else:
        with st.container(border=True):
            fc1,fc2,fc3,fc4=st.columns(4)
            search_seg=fc1.text_input("Buscar picker",placeholder="Nombre…",key="seguimiento_global_search")
            seg_filter=fc2.selectbox("Estado",["Todos","Con actas","Sin actas","Con cero tolerancia","Sin seguimiento"],key="seguimiento_global_estado")
            seg_sort=fc3.selectbox("Ordenar por",["ACTAS","SEGUIMIENTOS","CERO TOLERANCIA","RETROALIMENTACIONES","TOTAL"],key="seguimiento_global_sort")
            seg_dir=fc4.selectbox("Orden",["Mayor a menor","Menor a mayor"],key="seguimiento_global_dir")
            view=seguimiento_df.copy()
            if search_seg.strip(): view=view[view["PICKER"].map(norm).str.contains(norm(search_seg),regex=False)]
            if seg_filter=="Con actas": view=view[view["ACTAS"]>0]
            elif seg_filter=="Sin actas": view=view[view["ACTAS"]==0]
            elif seg_filter=="Con cero tolerancia": view=view[view["CERO TOLERANCIA"]>0]
            elif seg_filter=="Sin seguimiento": view=view[view["SEGUIMIENTOS"]==0]
            view=view.sort_values(seg_sort,ascending=(seg_dir=="Menor a mayor"))
            st.dataframe(view,use_container_width=True,hide_index=True)
        st.caption(f"{len(view):,} pickers visibles.")

    st.divider()
    nombres=person_options(selected_context)
    with st.container(border=True):
        st.markdown("**🔎 Abrir expediente individual**")
        buscar=st.text_input("Nombre del picker",placeholder="Escribe parte del nombre…",key="seguimiento_picker_search")
        opciones=nombres
        if buscar.strip():
            term=norm(buscar); opciones=[n for n in nombres if term in norm(n)]
        seguimiento_picker=st.selectbox("Picker",["Selecciona un picker"]+opciones,key="seguimiento_picker_select")

    if seguimiento_picker == "Selecciona un picker":
        st.info("Selecciona un picker para consultar y documentar su expediente.")
    else:
        r=s[s.PICKER==seguimiento_picker].iloc[0]
        rec=picker_record(store,seguimiento_picker,aliases=[str(r.get("_SOURCE_PICKER",""))])
        acciones=rec.get("acciones",[]) or []; documentos=rec.get("documentos",[]) or []
        st.markdown(f"<div class='justo-card'><div class='justo-kicker'>Expediente</div><div class='justo-title'>{seguimiento_picker}</div><div class='justo-muted'>Turno: {r.TURNO} · Supervisor: {r.SUPERVISOR} · Área: {r.AREA_BASE}</div></div>",unsafe_allow_html=True)
        k1,k2,k3,k4=st.columns(4)
        k1.metric("Retroalimentaciones",f"{sum(1 for z in acciones if str(z.get('retroalimentacion','')).strip()):,}")
        k2.metric("Seguimientos",f"{len(acciones):,}")
        k3.metric("Actas / tolerancias",f"{sum(1 for z in acciones if 'Acta' in str(z.get('accion','')) or 'Cero tolerancia' in str(z.get('accion',''))):,}")
        k4.metric("Documentos",f"{len(documentos):,}")
        with st.container(border=True):
            st.markdown("### 📋 Historial de seguimiento")
            if acciones:
                hist=pd.DataFrame(acciones).rename(columns={"fecha":"Fecha","accion":"Tipo","supervisor":"Supervisor","motivo":"Motivo","retroalimentacion":"Retroalimentación"})
                cols=[c for c in ["Fecha","Tipo","Supervisor","Motivo","Retroalimentación"] if c in hist.columns]
                st.dataframe(hist.sort_values("Fecha",ascending=False)[cols],use_container_width=True,hide_index=True)
            else: st.info("Este picker todavía no tiene seguimientos registrados.")
        with st.container(border=True):
            st.markdown("### 📄 Subir seguimiento / acta")
            st.caption("Adjunta un PDF o elige uno de los enlaces de seguimiento que ya guardaste.")
            recursos=store.get("recursos_formatos",[]) or []
            recurso_lookup={}
            for i,z in enumerate(recursos):
                if str(z.get("url","")).strip():
                    recurso_lookup[f"{i+1}. {z.get('tipo','Recurso')} · {z.get('titulo','Enlace')}"]=z
            recurso_opciones=["Sin enlace guardado"]+list(recurso_lookup)
            with st.container(border=True):
                st.markdown("**🔗 Elegir enlace guardado (opcional)**")
                recurso_seleccionado=st.selectbox("Enlaces de seguimiento",recurso_opciones,key=f"seguimiento_enlace_guardado_{person_key(seguimiento_picker)}")
                recurso_actual=recurso_lookup.get(recurso_seleccionado)
                if recurso_actual:
                    st.info(f"Seleccionado: {recurso_actual.get('tipo','Seguimiento')} · {recurso_actual.get('titulo','Enlace de seguimiento')}")
                    st.link_button("Abrir enlace para revisar",str(recurso_actual.get("url","")))
            enviar_copia=st.checkbox("Enviar copia por correo al guardar",value=False,key=f"seguimiento_email_{person_key(seguimiento_picker)}")
            destinatarios=st.text_input("Correo(s) destinatario(s)",placeholder="persona@correo.com (separa varios con coma)",key=f"seguimiento_destinatarios_{person_key(seguimiento_picker)}") if enviar_copia else ""
            with st.form(f"seguimiento_documento_form_{seguimiento_picker}",clear_on_submit=True):
                dc1,dc2=st.columns([1,2])
                with dc1: doc_tipo=st.selectbox("Tipo de documento",["Acta 1","Acta 2","Acta 3","Llamada de atención","Advertencia verbal 1","Advertencia verbal 2","Cero tolerancia","Otro"])
                with dc2: doc_titulo=st.text_input("Nombre / referencia",placeholder="Ej. Acta por FNR — septiembre 2026")
                doc_detalle=st.text_area("Detalle / motivo",placeholder="Qué originó el seguimiento y cualquier dato importante…")
                doc_pdf=st.file_uploader("📎 Adjuntar PDF",type=["pdf"],accept_multiple_files=False,key=f"seguimiento_pdf_{seguimiento_picker}")
                if st.form_submit_button("💾 Guardar seguimiento / documento",type="primary"):
                    url=str(recurso_actual.get("url","")).strip() if recurso_actual else ""
                    titulo=doc_titulo.strip() or (doc_pdf.name if doc_pdf is not None else (str(recurso_actual.get("titulo",doc_tipo)) if recurso_actual else doc_tipo))
                    if doc_pdf is None and not url: st.error("Adjunta un PDF o selecciona un enlace guardado.")
                    else:
                        registro={"id":datetime.now().strftime("%Y%m%d%H%M%S%f"),"fecha":datetime.now().strftime("%Y-%m-%d %H:%M"),"tipo":doc_tipo,"titulo":titulo,"detalle":doc_detalle.strip(),"supervisor":str(r.get("SUPERVISOR","")),"url":url,"url_titulo":(str(recurso_actual.get("titulo","")) if recurso_actual else ""),"path":"","archivo":"","destinatarios":destinatarios}
                        pdf_bytes=None
                        if doc_pdf is not None:
                            path,_=save_followup_pdf(doc_pdf.getvalue(),seguimiento_picker,doc_pdf.name); registro["path"]=path; registro["archivo"]=_safe_filename(doc_pdf.name); pdf_bytes=doc_pdf.getvalue()
                        rec.setdefault("documentos",[]).append(registro); save_store(store)
                        if enviar_copia:
                            email_body=f"Se registró un {doc_tipo} para {seguimiento_picker}.\n\nDetalle: {doc_detalle.strip() or 'Sin detalle.'}"
                            if url: email_body+=f"\n\nEnlace de seguimiento: {url}"
                            ok,msg=send_followup_email(f"Seguimiento {doc_tipo} · {seguimiento_picker}",email_body,destinatarios,pdf_bytes,registro.get("archivo") or "seguimiento.pdf")
                            if ok: st.success(msg)
                            else: st.error(msg)
                        else: st.success("Seguimiento / documento guardado.")
        st.markdown("### 🔗 Enlaces de seguimiento")
        st.caption("Aquí puedes consultar y administrar los enlaces guardados.")
        with st.expander(f"Administrar enlaces de seguimiento ({len(recursos)})",expanded=False):
            st.info("Los enlaces guardados aparecen aquí para abrirlos o quitarlos.")
            for idx,recurso in enumerate(recursos):
                with st.container(border=True):
                    rr1,rr2,rr3=st.columns([4,1.3,1])
                    with rr1:
                        st.markdown(f"**🔗 {recurso.get('titulo','Formato')}**")
                        st.caption(f"🟦 {recurso.get('tipo','Recurso')} · Agregado {recurso.get('fecha','')}")
                    with rr2:
                        if recurso.get("url"): st.link_button("Abrir enlace",str(recurso.get("url")),use_container_width=True)
                    with rr3:
                        if st.button("Quitar",key=f"seg_recurso_del_{idx}",use_container_width=True): store["recursos_formatos"].pop(idx); save_store(store); st.rerun()
            with st.form("seg_recurso_form",clear_on_submit=True):
                q1,q2=st.columns([1,2])
                with q1: rt=st.selectbox("Tipo",["Retroalimentación","Seguimiento","Llamada de atención","Acta","Otro"])
                with q2: rn=st.text_input("Nombre")
                ru=st.text_input("Enlace",placeholder="https://...")
                if st.form_submit_button("Guardar enlace reutilizable"):
                    if rn.strip() and ru.startswith(("http://","https://")):
                        store.setdefault("recursos_formatos",[]).append({"tipo":rt,"titulo":rn.strip(),"url":ru.strip(),"fecha":datetime.now().strftime("%Y-%m-%d %H:%M")}); save_store(store); st.rerun()
                    else: st.error("Nombre y enlace válido son obligatorios.")
        documentos=sorted(rec.get("documentos",[]) or [],key=lambda z:str(z.get("fecha","")),reverse=True)
        st.markdown("### 📎 Documentos del expediente")
        if documentos:
            for i,doc in enumerate(documentos):
                archivo=load_followup_pdf(doc)
                rr1,rr2,rr3=st.columns([1.2,5,1.4])
                rr1.markdown(f"**{doc.get('tipo','Documento')}**"); rr1.caption(str(doc.get('fecha','')))
                rr2.markdown(f"**{doc.get('titulo',doc.get('archivo','Documento'))}**")
                if doc.get("detalle"): rr2.caption(str(doc.get("detalle")))
                with rr3:
                    if archivo is not None: st.download_button("Ver PDF",archivo,file_name=str(doc.get("archivo") or "seguimiento.pdf"),mime="application/pdf",key=f"seg_download_{seguimiento_picker}_{doc.get('id',i)}",use_container_width=True)
                    if doc.get("url"): st.link_button(f"Abrir: {doc.get('url_titulo') or 'enlace'}",str(doc.get("url")),use_container_width=True)
        else: st.info("No hay documentos vinculados a este expediente.")

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
