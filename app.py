# ================================================================
# PROPIEDAD / FIRMA DEL DESARROLLO
# Desarrollado por: Yael Maya Ruíz
# Control FNR & Mala Calidad · Operación Coyoacán
# ================================================================


import io, re, json, os, smtplib, ssl, zipfile, hashlib, urllib.request, urllib.error, urllib.parse, html, uuid
from difflib import SequenceMatcher
from datetime import datetime, timedelta
from email.message import EmailMessage
import numpy as np
import pandas as pd
import streamlit as st
import qrcode
from zoneinfo import ZoneInfo
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

st.set_page_config(page_title="Control FNR & Mala Calidad", page_icon="🥑", layout="wide", initial_sidebar_state="expanded")

# Estilo visual inspirado en la interfaz limpia de Jüsto: blanco, rojo de marca y tarjetas suaves.
st.markdown("""
<style>
:root { --justo-red:#BD2426; --ink:#272936; --muted:#6f7480; --soft:#f7f7f5; --line:#e7e7e3; --success:#15803d; }
.main .block-container { max-width: 1500px; padding-top: 1.1rem; padding-bottom: 2.5rem; }
[data-testid="stSidebar"] { border-right: 1px solid var(--line); }
section[data-testid="stSidebar"] { width: min(320px, 28vw) !important; min-width: 260px !important; }
section[data-testid="stSidebar"] > div { width: inherit !important; }
@media (max-width: 700px) {
  section[data-testid="stSidebar"] { width: min(88vw, 320px) !important; min-width: 0 !important; }
  section[data-testid="stSidebar"] > div { width: 100% !important; }
}
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

@st.cache_data(show_spinner=False, max_entries=4)
def parse_oneoff_area_fr(data):
    """Lee un Excel puntual de FR por artículo, sin guardar el archivo en la nube."""
    try:
        raw=pd.read_excel(io.BytesIO(data))
    except Exception as exc:
        raise ValueError(f"No pude abrir el Excel: {exc}") from exc
    columns={str(c).strip().casefold().replace("_"," ").replace("-"," "):c for c in raw.columns}
    def pick(*names):
        for name in names:
            key=name.casefold().replace("_"," ").replace("-"," ")
            if key in columns: return columns[key]
        return None
    date_col=pick("date","fecha")
    product_col=pick("product","producto","article","articulo")
    department_col=pick("department","departamento","area","área")
    order_col=pick("order number","order_number","pedido","numero de pedido")
    if date_col is None or product_col is None or department_col is None:
        raise ValueError("El Excel debe incluir columnas de fecha, artículo/producto y departamento.")
    out=pd.DataFrame({
        "Fecha":pd.to_datetime(raw[date_col],errors="coerce"),
        "Artículo":raw[product_col].astype("string").str.strip(),
        "Departamento":raw[department_col].astype("string").str.strip(),
    })
    out["Pedido"]=raw[order_col].astype("string").str.strip() if order_col is not None else pd.NA
    out=out.dropna(subset=["Fecha","Artículo","Departamento"])
    out=out[out["Artículo"].ne("") & out["Departamento"].ne("")]
    return out.reset_index(drop=True)

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

def _attendance_find_sheet(data, required_columns, label):
    """Encuentra la hoja de asistencia que contiene las columnas necesarias."""
    parsed=sheets_from_bytes(data)
    wanted=set(required_columns)
    for frame in parsed.values():
        if wanted.issubset(set(frame.columns)):
            return frame
    actual=[", ".join(map(str,frame.columns)) for frame in parsed.values()]
    raise ValueError(f"No encontré en el Excel de {label} las columnas necesarias: {', '.join(required_columns)}. Hojas encontradas: {'; '.join(actual)}")

def _attendance_minutes(value):
    """Convierte el atraso a minutos, admitiendo los formatos comunes de Excel."""
    if pd.isna(value): return 0
    if isinstance(value,(int,float,np.integer,np.floating)):
        return max(0,int(round(float(value))))
    try:
        td=pd.to_timedelta(value)
        return max(0,int(round(td.total_seconds()/60)))
    except Exception:
        try: return max(0,int(round(float(str(value).strip()))))
        except Exception: return 0

def _attendance_start_hour(value):
    """Devuelve la hora de inicio del turno como HH:MM."""
    if pd.isna(value): return ""
    if hasattr(value,"hour") and hasattr(value,"minute"):
        return f"{int(value.hour):02d}:{int(value.minute):02d}"
    if isinstance(value,(int,float,np.integer,np.floating)):
        number=float(value)
        if 0 <= number < 1: number*=24
        hour=int(number); minute=int(round((number-hour)*60))
        if minute==60: hour+=1; minute=0
        return f"{hour%24:02d}:{minute:02d}"
    text=str(value).strip()
    try:
        td=pd.to_timedelta(text)
        total=int(round(td.total_seconds()/60))
        return f"{(total//60)%24:02d}:{total%60:02d}"
    except Exception:
        return text[:5]

def _attendance_shift(start,config=None):
    """Clasifica horario usando reglas configurables."""
    if not start: return "Sin dato"
    try: hour=int(str(start).split(":",1)[0])
    except Exception: return "Sin dato"
    rules=config or {
        "Mañana":[6,7,8,9,10],
        "Intermedio":[11],
        "Tarde":[13,14],
        "Nocturno":[22],
    }
    for label,hours in rules.items():
        try:
            if hour in {int(x) for x in hours}: return str(label)
        except Exception:
            continue
    return "Otro horario"

def _attendance_roster_shift(value,config=None):
    text=str(value or "").strip()
    folded=norm(text)
    if "manana" in folded: return "Mañana"
    if "inter" in folded: return "Intermedio"
    if "tarde" in folded: return "Tarde"
    if "noct" in folded: return "Nocturno"
    start=_attendance_start_hour(value)
    return _attendance_shift(start,config) if start else "Sin turno"
def _attendance_person_key(identifier, lastname, firstname):
    """Usa el CURP solo durante el cruce; el CURP no se guarda en el resumen."""
    value=str(identifier or "").strip().upper()
    if value and value not in {"NAN","NONE","NAT"}: return "id:"+value
    return "nombre:"+(token_key(f"{lastname} {firstname}") or person_key(f"{lastname} {firstname}"))

def build_attendance_summary(absence_bytes,tardy_bytes,absence_name="Faltas",tardy_name="Retardos",roster_rows=None,shift_config=None):
    """Cruza faltas/retardos y completa turno/supervisor/área desde plantilla."""
    absent=_attendance_find_sheet(absence_bytes,["apellidos","nombre","fecha"],"faltas")
    tardy=_attendance_find_sheet(tardy_bytes,["apellidos","nombre","fecha","hora_inicio_turno","minutos_de_atraso"],"retardos")
    people={}
    event_rows=[]
    date_values=[]
    roster_df=roster_rows.copy() if isinstance(roster_rows,pd.DataFrame) else pd.DataFrame(roster_rows or [])

    def row_person(row):
        _last=row.get("apellidos",""); _first=row.get("nombre","")
        lastname="" if pd.isna(_last) else str(_last).strip()
        firstname="" if pd.isna(_first) else str(_first).strip()
        display=f"{lastname}, {firstname}".strip(" ,")
        _identifier=row.get("curp","")
        if pd.isna(_identifier): _identifier=""
        return _attendance_person_key(_identifier,lastname,firstname),display

    def ensure_person(key,name):
        item=people.setdefault(key,{
            "nombre":name,"faltas":0,"retardos":0,"minutos_atraso":0,
            "turnos":{},"cargos":set(),"turno_plantilla":"","supervisor":"","area":"",
        })
        if not item.get("nombre"): item["nombre"]=name
        return item

    for _,row in absent.iterrows():
        key,name=row_person(row)
        if not name: continue
        item=ensure_person(key,name)
        item["faltas"]+=1
        _cargo=row.get("cargo",""); cargo="" if pd.isna(_cargo) else str(_cargo).strip()
        if cargo and cargo.lower() not in {"nan","none"}: item["cargos"].add(cargo)
        date=pd.to_datetime(row.get("fecha"),errors="coerce")
        if not pd.isna(date): date_values.append(date.strftime("%Y-%m-%d"))

    for _,row in tardy.iterrows():
        key,name=row_person(row)
        if not name: continue
        item=ensure_person(key,name)
        item["retardos"]+=1
        minutes=_attendance_minutes(row.get("minutos_de_atraso",0))
        item["minutos_atraso"]+=minutes
        start=_attendance_start_hour(row.get("hora_inicio_turno",""))
        shift=_attendance_shift(start,shift_config)
        item["turnos"][shift]=item["turnos"].get(shift,0)+1
        event_rows.append({"turno":shift,"persona_key":key,"minutos":minutes})
        _cargo=row.get("cargo",""); cargo="" if pd.isna(_cargo) else str(_cargo).strip()
        if cargo and cargo.lower() not in {"nan","none"}: item["cargos"].add(cargo)
        date=pd.to_datetime(row.get("fecha"),errors="coerce")
        if not pd.isna(date): date_values.append(date.strftime("%Y-%m-%d"))

    # Completar contexto desde la plantilla usando el mismo resolvedor conservador.
    if not roster_df.empty:
        for item in people.values():
            rr=_match_master_row(str(item.get("nombre","")).replace(","," "),roster_df,"")
            if rr is None: continue
            raw_shift=rr.get("TURNO_MAESTRO",rr.get("TURNO",""))
            item["turno_plantilla"]=_attendance_roster_shift(raw_shift,shift_config)
            item["supervisor"]=str(rr.get("SUPERVISOR","")).strip()
            item["area"]=str(rr.get("AREA_MAESTRO",rr.get("AREA_BASE",""))).strip()

    person_rows=[]
    absent_by_shift={}
    for key,item in people.items():
        turns=item["turnos"]
        template_shift=str(item.get("turno_plantilla","")).strip()
        if template_shift and template_shift not in {"Sin turno","Sin dato","Otro horario"}:
            usual=template_shift
        elif turns:
            max_count=max(turns.values())
            leading=sorted([shift for shift,count in turns.items() if count==max_count])
            usual=leading[0] if len(leading)==1 else "Varios turnos"
        else:
            usual="Sin turno"
        observed={z for z in turns if z not in {"Sin dato","Otro horario"}}
        varies=len(observed)>1 or (bool(template_shift) and bool(observed) and any(z!=template_shift for z in observed))
        person_rows.append({
            "Persona":item["nombre"],
            "Turno habitual":usual,
            "Turno plantilla":template_shift or "Sin dato",
            "Turnos observados":", ".join(sorted(turns)) if turns else "Sin dato",
            "Turno variable":bool(varies),
            "Supervisor":item.get("supervisor",""),
            "Área":item.get("area",""),
            "Faltas":int(item["faltas"]),
            "Retardos":int(item["retardos"]),
            "Minutos acumulados":int(item["minutos_atraso"]),
            "Cargo":", ".join(sorted(item["cargos"])) if item["cargos"] else "",
        })
        shift=usual if usual!="Varios turnos" else "Revisar turno"
        if item["faltas"]:
            absent_by_shift.setdefault(shift,{"personas":set(),"faltas":0})
            absent_by_shift[shift]["personas"].add(key)
            absent_by_shift[shift]["faltas"]+=int(item["faltas"])

    shift_summary={}
    for event in event_rows:
        shift=event["turno"]
        shift_summary.setdefault(shift,{"personas":set(),"retardos":0,"minutos":0})
        shift_summary[shift]["personas"].add(event["persona_key"])
        shift_summary[shift]["retardos"]+=1
        shift_summary[shift]["minutos"]+=int(event["minutos"])
    return {
        "personas":person_rows,
        "faltas_por_turno":[{"Turno":k,"Personas":len(v["personas"]),"Faltas":int(v["faltas"])} for k,v in sorted(absent_by_shift.items())],
        "retardos_por_turno":[{"Turno":k,"Personas":len(v["personas"]),"Retardos":int(v["retardos"]),"Minutos acumulados":int(v["minutos"])} for k,v in sorted(shift_summary.items())],
        "total_personas":len(person_rows),
        "total_faltas":int(sum(x["Faltas"] for x in person_rows)),
        "personas_con_faltas":int(sum(x["Faltas"]>0 for x in person_rows)),
        "total_retardos":int(sum(x["Retardos"] for x in person_rows)),
        "personas_con_retardos":int(sum(x["Retardos"]>0 for x in person_rows)),
        "minutos_acumulados":int(sum(x["Minutos acumulados"] for x in person_rows)),
        "fecha_desde":min(date_values) if date_values else "",
        "fecha_hasta":max(date_values) if date_values else "",
        "archivos":{"faltas":_safe_filename(absence_name),"retardos":_safe_filename(tardy_name)},
        "fecha_carga":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

FNR_OBJ, MC_OBJ = 1.50, 1.00
FOLLOWUP_CATEGORIES = ["1. Seguimiento", "2. Acta", "3. 0 Tolerancia"]
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
STORE_SHARDS={
    "pickers":("state/pickers.json",os.path.join(PERSIST_DIR,"pickers.json"),{}),
    "order_audits":("state/order_audits.json",os.path.join(PERSIST_DIR,"order_audits.json"),[]),
    "monthly_history":("state/monthly_history.json",os.path.join(PERSIST_DIR,"monthly_history.json"),{}),
    "powerbi_krs_history":("state/powerbi_krs_history.json",os.path.join(PERSIST_DIR,"powerbi_krs_history.json"),[]),
}
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

def cloud_list(prefix,limit=1000,offset=0):
    if not cloud_enabled(): return []
    _cloud_ensure_bucket()
    _,_,bucket=cloud_config()
    body=json.dumps({"prefix":str(prefix).strip("/"),"limit":int(limit),"offset":int(offset),"sortBy":{"column":"name","order":"desc"}}).encode("utf-8")
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
    if cloud_enabled() and changed:
        cloud_upload(current_key,data,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if changed:
        st.session_state.pop("_excel_history_zip_ready",None)
        for _cached in (upload_history_zip,upload_history_rows,persisted_status,_load_persisted_upload_bytes):
            try: _cached.clear()
            except Exception: pass
    return io.BytesIO(data)

def persist_uploaded_once(upload,key):
    """Evita volver a guardar el mismo archivo tras cada interacción de Streamlit."""
    if upload is None: return load_persisted_upload(key)
    file_id=getattr(upload,"file_id",None) or getattr(upload,"id",None)
    identity=(file_id,str(getattr(upload,"name","")),int(getattr(upload,"size",0) or 0))
    if not file_id:
        identity+=(hashlib.sha256(upload.getvalue()).hexdigest(),)
    seen=st.session_state.setdefault("_processed_excel_uploads",{})
    if seen.get(key)==identity:
        return io.BytesIO(upload.getvalue())
    result=persist_upload(upload,key)
    seen[key]=identity
    return result

@st.cache_data(show_spinner=False,max_entries=4,ttl=30)
def _load_persisted_upload_bytes(key):
    """Cache corto para que los reruns no vuelvan a descargar los cuatro Excel."""
    path=PERSIST_FILES[key]
    if cloud_enabled():
        data=cloud_download(f"uploads/current/{key}.xlsx",missing_ok=True)
        if data is not None:
            ensure_persist_dir()
            with open(path,"wb") as f: f.write(data)
            return data
    if not os.path.exists(path): return None
    try:
        with open(path,"rb") as f: data=f.read()
        if not data: return None
        if cloud_enabled():
            cloud_upload(f"uploads/current/{key}.xlsx",data,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        return data
    except Exception:
        return None

@st.cache_data(show_spinner=False,max_entries=1,ttl=60)
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
    """Recupera el Excel guardado; los bytes se cachean durante 30 segundos."""
    data=_load_persisted_upload_bytes(key)
    return io.BytesIO(data) if data is not None else None

@st.cache_data(show_spinner=False,max_entries=1,ttl=60)
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

@st.cache_data(show_spinner=False,max_entries=300,ttl=300)
def _load_cloud_asset_bytes(key):
    return cloud_download(key,missing_ok=True)

def load_process_images(process):
    result=[]
    for path in process.get("imagenes",[]):
        if str(path).startswith("cloud://"):
            key=str(path)[len("cloud://"):]
            try:
                data=_load_cloud_asset_bytes(key)
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

def save_followup_pdf(data, picker, filename, identity_email=""):
    """Guarda un PDF de acta/seguimiento asociado a un picker."""
    ensure_persist_dir()
    doc_id=datetime.now().strftime("%Y%m%d%H%M%S%f")
    _email_key=email_key(identity_email)
    _owner_folder=("email-"+hashlib.sha256(_email_key.encode("utf-8")).hexdigest()[:20]) if _email_key else (person_key(picker) or "picker")
    folder=os.path.join(PERSIST_DIR, "seguimiento_pdfs", _owner_folder, doc_id)
    os.makedirs(folder, exist_ok=True)
    safe=_safe_filename(filename)
    path=os.path.join(folder, safe)
    with open(path, "wb") as f: f.write(data)
    if cloud_enabled():
        key=f"seguimiento_pdfs/{_owner_folder}/{doc_id}/{safe}"
        cloud_upload(key,data,"application/pdf")
        return "cloud://"+key, doc_id
    return path, doc_id

def load_followup_pdf(doc):
    path=str(doc.get("path", ""))
    if path.startswith("cloud://"):
        try: return _load_cloud_asset_bytes(path[len("cloud://"):])
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

def _json_atomic_write(path,value):
    os.makedirs(os.path.dirname(path) or ".",exist_ok=True)
    payload=json.dumps(value,ensure_ascii=False,indent=2).encode("utf-8")
    tmp=path+".tmp"
    with open(tmp,"wb") as f: f.write(payload)
    os.replace(tmp,path)
    return payload

def _json_local_read(path,default):
    try:
        with open(path,"r",encoding="utf-8") as f:
            value=json.load(f)
        return value
    except Exception:
        return default.copy() if isinstance(default,dict) else list(default) if isinstance(default,list) else default

def _cloud_json_read(key,default):
    if not cloud_enabled():
        return default.copy() if isinstance(default,dict) else list(default) if isinstance(default,list) else default
    raw=cloud_download(key,missing_ok=True)
    if raw is None:
        return default.copy() if isinstance(default,dict) else list(default) if isinstance(default,list) else default
    try:
        return json.loads(raw.decode("utf-8"))
    except (ValueError,UnicodeDecodeError) as exc:
        raise RuntimeError(f"El archivo {key} de Supabase está dañado o no es JSON válido. No se guardará encima.") from exc

def _unique_records(records):
    """Fusión por ID sin duplicar seguimientos cuando cambia su estado de correo."""
    out=[]; positions={}
    for item in records or []:
        if isinstance(item,dict) and str(item.get("id","")).strip():
            marker="id:"+str(item["id"]).strip()
        else:
            marker="data:"+json.dumps(item,ensure_ascii=False,sort_keys=True,default=str)
        if marker in positions:
            # La versión posterior prevalece para metadatos, nunca elimina campos anteriores.
            idx=positions[marker]
            if isinstance(out[idx],dict) and isinstance(item,dict):
                updated=dict(out[idx]); updated.update(item);out[idx]=updated
            continue
        positions[marker]=len(out);out.append(item)
    return out

def _merge_picker_store(remote,local):
    result=dict(remote or {})
    for person,lrec in (local or {}).items():
        if person not in result or not isinstance(result.get(person),dict) or not isinstance(lrec,dict):
            result[person]=lrec
            continue
        merged=dict(result[person])
        for key,value in lrec.items():
            if key in {"comentarios","acciones","documentos"}:
                merged[key]=_unique_records(list(merged.get(key,[]) or [])+list(value or []))
            else:
                merged[key]=value
        result[person]=merged
    return result

def _merge_store_states(remote,local):
    """Fusión conservadora si dos sesiones guardan casi al mismo tiempo."""
    merged=dict(remote or {})
    merged.update({k:v for k,v in (local or {}).items() if k not in STORE_SHARDS})
    merged["pickers"]=_merge_picker_store((remote or {}).get("pickers",{}),(local or {}).get("pickers",{}))

    remote_audits=(remote or {}).get("order_audits",[]) or []
    local_audits=(local or {}).get("order_audits",[]) or []
    by_id={}
    for item in remote_audits+local_audits:
        if not isinstance(item,dict): continue
        key=str(item.get("id") or (str(item.get("pedido",""))+"|"+str(item.get("fecha_auditoria",""))))
        by_id[key]=item
    merged["order_audits"]=list(by_id.values())

    hist=dict((remote or {}).get("monthly_history",{}) or {})
    hist.update((local or {}).get("monthly_history",{}) or {})
    merged["monthly_history"]=hist

    pbi={}
    for item in list((remote or {}).get("powerbi_krs_history",[]) or [])+list((local or {}).get("powerbi_krs_history",[]) or []):
        if not isinstance(item,dict): continue
        key=(str(item.get("fecha","")),str(item.get("periodo","")),str(item.get("tienda","")))
        old=pbi.get(key)
        if old is None or str(item.get("cargado",""))>=str(old.get("cargado","")):
            pbi[key]=item
    merged["powerbi_krs_history"]=sorted(pbi.values(),key=lambda item:(str(item.get("fecha","")),str(item.get("cargado",""))))
    return merged

FOLLOWUP_JOURNAL_PREFIX="seguimientos_registros/"

def _followup_event_objects():
    """Lista TODOS los eventos paginando el bucket. Un fallo bloquea la carga."""
    if not cloud_enabled():
        return []
    objects=[]
    offset=0
    while True:
        batch=cloud_list(FOLLOWUP_JOURNAL_PREFIX,limit=500,offset=offset)
        if not isinstance(batch,list):
            raise RuntimeError("La lista de respaldos de seguimiento no se pudo consultar.")
        objects.extend(item for item in batch if isinstance(item,dict) and str(item.get("name","")).endswith(".json"))
        if len(batch)<500:
            return objects
        offset+=len(batch)

def _load_followup_journal(store):
    """Reincorpora registros individuales que faltan en el índice de expedientes."""
    if not cloud_enabled():
        return 0
    existing_ids={
        str(rec.get("id"))
        for p in (store.get("pickers",{}) or {}).values()
        if isinstance(p,dict)
        for typ in ("acciones","documentos")
        for rec in p.get(typ,[]) or []
        if isinstance(rec,dict) and rec.get("id")
    }
    restored=0
    for obj in _followup_event_objects():
        name=str(obj.get("name",""))
        record_id=name.rsplit("/",1)[-1].removesuffix(".json")
        if record_id in existing_ids:
            continue
        cloud_path=name if name.startswith(FOLLOWUP_JOURNAL_PREFIX) else FOLLOWUP_JOURNAL_PREFIX+name
        payload=_cloud_json_read(cloud_path,None)
        if not isinstance(payload,dict) or payload.get("kind") not in {"acciones","documentos"}:
            raise RuntimeError(f"Respaldo individual inválido: {cloud_path}")
        picker=str(payload.get("picker","")).strip()
        record=payload.get("registro")
        if not picker or not isinstance(record,dict):
            raise RuntimeError(f"Respaldo individual incompleto: {cloud_path}")
        record_id=str(record.get("id") or record_id)
        record["id"]=record_id
        _correo_key=email_key(payload.get("correo_key",payload.get("correo","")))
        if _correo_key:
            rec=followup_email_record(
                store,picker,_correo_key,
                aliases=payload.get("aliases",[]) or [],
                migrate_legacy=False,create=True,
            )
        else:
            rec=picker_record(store,picker,aliases=payload.get("aliases",[]) or [])
        rec.setdefault(payload["kind"],[]).append(record)
        existing_ids.add(record_id)
        restored+=1
    return restored

def save_followup_entry(store,picker,kind,registro,aliases=None,identity_email="",identity_fields=None):
    """Respalda cada seguimiento por separado ANTES de modificar el índice global.

    Si falla Supabase, no se informa un falso éxito. El archivo de evento
    es independiente de pickers.json y no puede desaparecer por una escritura
    concurrente de otro supervisor.
    """
    if kind not in {"acciones","documentos"}:
        raise ValueError("Tipo de seguimiento inválido.")
    if not cloud_enabled():
        raise RuntimeError("No se guardó el seguimiento: falta configurar la nube. No se admiten guardados temporales.")
    picker=str(picker or "").strip()
    if not picker:
        raise ValueError("Selecciona un picker.")
    record=dict(registro)
    record["id"]=str(record.get("id") or uuid.uuid4().hex)
    path=FOLLOWUP_JOURNAL_PREFIX+record["id"]+".json"
    _correo_key=email_key(identity_email)
    _identity_fields=dict(identity_fields or {})
    if _correo_key:
        record["picker"]=picker
        record["correo"]=str(_identity_fields.get("CORREO",identity_email) or identity_email).strip()
        record["correo_key"]=_correo_key
        for _src,_dst in (("PERSONAL_TIPO","tipo_personal"),("TURNO","turno"),("SUPERVISOR","supervisor"),("AREA_BASE","area")):
            if str(_identity_fields.get(_src,"")).strip(): record[_dst]=str(_identity_fields[_src]).strip()
    journal={"v":2 if _correo_key else 1,"picker":picker,"aliases":list(aliases or []),"kind":kind,"registro":record}
    if _correo_key:
        journal["correo_key"]=_correo_key
        journal["identidad_plantilla"]={
            "PICKER":picker,"CORREO_KEY":_correo_key,
            "CORREO":str(_identity_fields.get("CORREO",identity_email) or identity_email).strip(),
            "PERSONAL_TIPO":str(_identity_fields.get("PERSONAL_TIPO","")).strip(),
            "TURNO":str(_identity_fields.get("TURNO","")).strip(),
            "SUPERVISOR":str(_identity_fields.get("SUPERVISOR","")).strip(),
            "AREA_BASE":str(_identity_fields.get("AREA_BASE","")).strip(),
        }
    payload=json.dumps(journal,ensure_ascii=False,sort_keys=True).encode("utf-8")
    cloud_upload(path,payload,"application/json")
    rec=(followup_email_record(store,picker,_correo_key,aliases=aliases,migrate_legacy=True,create=True)
         if _correo_key else picker_record(store,picker,aliases=aliases))
    rec[kind]=_unique_records(list(rec.get(kind,[]) or [])+[record])
    try:
        save_store(store)
    except Exception as exc:
        raise RuntimeError("El seguimiento SÍ quedó respaldado individualmente en la nube, pero falló la actualización del índice; al volver a entrar se reconstruirá desde el respaldo. "+str(exc)) from exc
    return record

def _load_store_from_sources():
    ensure_persist_dir()
    cloud_main=_cloud_json_read(CLOUD_STORE_KEY,{}) if cloud_enabled() else {}
    if cloud_main:
        data=cloud_main if isinstance(cloud_main,dict) else {}
    else:
        data={}
        for path in [STORE_FILE,LEGACY_STORE_FILE]:
            if os.path.exists(path):
                candidate=_json_local_read(path,{})
                if isinstance(candidate,dict):
                    data=candidate
                    break

    defaults={
        "pickers":{},"feedback_rows":[],"recursos_formatos":[],"procesos":[],
        "excluded_orders":[],"master_overrides":{},"master_excluded":[],
        "seguimientos_documentos":[],"supervisores":[],"upload_meta":{},
        "attendance_summary":None,"attendance_meta":{},"attendance_links":{},
        "attendance_shift_config":{"Mañana":[6,7,8,9,10],"Intermedio":[11],"Tarde":[13,14],"Nocturno":[22]},
        "order_audits":[],"monthly_history":{},"powerbi_krs_history":[],
        "_revision":0,
    }
    for key,value in defaults.items():
        data.setdefault(key,value.copy() if isinstance(value,dict) else list(value) if isinstance(value,list) else value)

    # Los shards nuevos prevalecen; si aún no existen se conserva el dato legacy
    # que venía dentro del JSON principal.
    for field,(cloud_key,local_path,default) in STORE_SHARDS.items():
        shard=None
        if cloud_enabled():
            raw=cloud_download(cloud_key,missing_ok=True)
            if raw is not None:
                try:
                    shard=json.loads(raw.decode("utf-8"))
                except (ValueError,UnicodeDecodeError) as exc:
                    raise RuntimeError(f"El archivo {cloud_key} no contiene JSON válido; se bloqueó la carga para evitar sobrescribir los datos.") from exc
        if shard is None and os.path.exists(local_path):
            shard=_json_local_read(local_path,default)
        if shard is not None:
            if field=="pickers":
                data[field]=_merge_picker_store(data.get(field,{}) or {},shard if isinstance(shard,dict) else {})
            elif field=="order_audits":
                data[field]=_unique_records(list(data.get(field,[]) or [])+list(shard or []))
            elif field=="monthly_history":
                data[field]={**(data.get(field,{}) or {}),**(shard or {})}
            else:
                data[field]=shard

    _load_followup_journal(data)
    # Mantener copia local útil para recuperación.
    main_local={k:v for k,v in data.items() if k not in STORE_SHARDS}
    try: _json_atomic_write(STORE_FILE,main_local)
    except OSError: pass
    for field,(_,local_path,_) in STORE_SHARDS.items():
        try: _json_atomic_write(local_path,data.get(field))
        except OSError: pass
    return data

def load_store():
    """Carga el estado una sola vez por sesión para evitar GETs a nube en cada rerun."""
    if "_persistent_store" not in st.session_state:
        st.session_state["_persistent_store"]=_load_store_from_sources()
    return st.session_state["_persistent_store"]

def save_store(store):
    """Guarda solo cuando hay cambios y usa shards para los bloques que más crecen."""
    ensure_persist_dir()
    store.setdefault("_revision",0)

    # Fusiona siempre los registros remotos, aunque otra sesión solo haya
    # actualizado el shard pickers.json y no la revisión del archivo principal.
    missing_shards=set()
    remote_shard_snapshot={}
    if cloud_enabled():
        remote_main=_cloud_json_read(CLOUD_STORE_KEY,{})
        if not isinstance(remote_main,dict):
            raise RuntimeError("El índice principal no tiene formato válido.")
        remote=dict(remote_main)
        for field,(cloud_key,_,default) in STORE_SHARDS.items():
            raw=cloud_download(cloud_key,missing_ok=True)
            if raw is None:
                missing_shards.add(field)
                remote[field]=remote_main.get(field,default)
            else:
                try: shard=json.loads(raw.decode("utf-8"))
                except (ValueError,UnicodeDecodeError) as exc:
                    raise RuntimeError(f"Se bloqueó la escritura porque {cloud_key} está dañado.") from exc
                if field=="pickers":
                    remote[field]=_merge_picker_store(remote_main.get(field,{}) or {},shard or {})
                elif field=="order_audits":
                    remote[field]=_unique_records(list(remote_main.get(field,[]) or [])+list(shard or []))
                elif field=="monthly_history":
                    remote[field]={**(remote_main.get(field,{}) or {}),**(shard or {})}
                else:
                    remote[field]=shard
                remote_shard_snapshot[field]=shard
        local_revision=pd.to_numeric(store.get("_revision",0),errors="coerce")
        remote_revision=pd.to_numeric(remote_main.get("_revision",0),errors="coerce")
        local_revision=0 if pd.isna(local_revision) else int(local_revision)
        remote_revision=0 if pd.isna(remote_revision) else int(remote_revision)
        merged=_merge_store_states(remote,store)
        store.clear(); store.update(merged)
        store["_revision"]=max(local_revision,remote_revision)+1

    main_data={k:v for k,v in store.items() if k not in STORE_SHARDS}
    main_payload=json.dumps(main_data,ensure_ascii=False,indent=2).encode("utf-8")

    # No escribir si ni el principal ni ningún shard cambió.
    changed_main=True
    try:
        with open(STORE_FILE,"rb") as f: changed_main=(f.read()!=main_payload)
    except OSError:
        pass
    changed_shards=[]
    for field,(_,local_path,_) in STORE_SHARDS.items():
        payload=json.dumps(store.get(field),ensure_ascii=False,indent=2).encode("utf-8")
        same=False
        try:
            with open(local_path,"rb") as f: same=(f.read()==payload)
        except OSError:
            pass
        # También comparar con la nube: un fallo de red no debe dejar
        # un archivo local actualizado pero el respaldo remoto desfasado.
        remote_same=True
        if cloud_enabled() and field not in missing_shards:
            remote_payload=json.dumps(remote_shard_snapshot.get(field),ensure_ascii=False,indent=2).encode("utf-8")
            remote_same=(payload==remote_payload)
        if not same or field in missing_shards or not remote_same:
            changed_shards.append((field,payload,local_path))

    if not cloud_enabled():
        if not changed_main and not changed_shards:
            return
        if changed_main:
            _json_atomic_write(STORE_FILE,main_data)
        for field,payload,local_path in changed_shards:
            _json_atomic_write(local_path,store.get(field))
    else:
        previous_backup=store.get("cloud_backup_at")
        backup_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        store["cloud_backup_at"]=backup_at
        main_data={k:v for k,v in store.items() if k not in STORE_SHARDS}
        main_payload=json.dumps(main_data,ensure_ascii=False,indent=2).encode("utf-8")
        try:
            for field,payload,_ in changed_shards:
                cloud_upload(STORE_SHARDS[field][0],payload,"application/json")
            cloud_upload(CLOUD_STORE_KEY,main_payload,"application/json")
        except Exception:
            if previous_backup is None: store.pop("cloud_backup_at",None)
            else: store["cloud_backup_at"]=previous_backup
            raise
        # La nube es la fuente persistente. En Streamlit Cloud el directorio
        # local puede ser efímero o no permitir reemplazos; el espejo local es
        # best-effort y nunca debe impedir guardar correctamente en Supabase.
        try:
            _json_atomic_write(STORE_FILE,{k:v for k,v in store.items() if k not in STORE_SHARDS})
        except OSError:
            pass
        for field,(_,local_path,_) in STORE_SHARDS.items():
            try: _json_atomic_write(local_path,store.get(field))
            except OSError: pass

def parse_powerbi_krs_excel(data, filename=""):
    """Extrae On Time, FNR y Mala Calidad de la hoja exportada desde Power BI."""
    try:
        sheets=pd.read_excel(io.BytesIO(data),sheet_name=None)
    except Exception as exc:
        raise ValueError(f"No pude abrir el Excel: {exc}") from exc
    selected=None
    for sheet_name,frame in sheets.items():
        frame=frame.copy()
        frame.columns=[str(c).strip() for c in frame.columns]
        lower={str(c).strip().casefold():c for c in frame.columns}
        kr_col=next((v for k,v in lower.items() if k in {"kr","indicador","nombre kr"}),None)
        value_col=next((v for k,v in lower.items() if k in {"valor","value"}),None)
        if kr_col is not None and value_col is not None:
            selected=(sheet_name,frame,kr_col,value_col,lower)
            break
    if selected is None:
        raise ValueError("No encontré una tabla con las columnas KR y Valor. En Power BI descarga los datos de la tabla de indicadores.")

    sheet_name,frame,kr_col,value_col,lower=selected
    def find_col(*names):
        for name in names:
            if name.casefold() in lower: return lower[name.casefold()]
        return None
    target_col=find_col("Meta MTD","Meta")
    def as_percent(value):
        if pd.isna(value): return None
        if isinstance(value,str):
            cleaned=value.strip().replace(",","")
            has_percent="%" in cleaned
            cleaned=cleaned.replace("%","").strip()
            try: number=float(cleaned)
            except (TypeError,ValueError): return None
            if not has_percent and abs(number)<=1: number*=100
            return round(number,4)
        try: number=float(value)
        except (TypeError,ValueError): return None
        if abs(number)<=1: number*=100
        return round(number,4)

    definitions={
        "On Time":lambda text:"on time" in text,
        "FNR":lambda text:"fnr" in text and ("no report" in text or "no reportad" in text or "faltantes no" in text),
        "Mala Calidad":lambda text:"mala calidad" in text,
    }
    indicators={}
    for _,row in frame.iterrows():
        label=str(row.get(kr_col,"")).strip()
        normalized=label.casefold()
        for name,matcher in definitions.items():
            if name not in indicators and matcher(normalized):
                indicators[name]={"valor":as_percent(row.get(value_col)),"meta":as_percent(row.get(target_col)) if target_col else None,"nombre_fuente":label}
    missing=[name for name in definitions if name not in indicators or indicators[name]["valor"] is None]
    if missing:
        raise ValueError("Faltan estos indicadores o no pude leer sus valores: " + ", ".join(missing) + ". Revisa que el Excel incluya On Time, FNR y Mala calidad.")

    all_text="\n".join(str(value) for sheet in sheets.values() for value in sheet.to_numpy().ravel() if pd.notna(value))
    period_match=re.search(r"MesAño\s+es\s+([^\r\n]+)",all_text,re.IGNORECASE)
    store_match=re.search(r"Tienda\s+es\s+([^\r\n]+)",all_text,re.IGNORECASE)
    period=period_match.group(1).strip() if period_match else "No indicado"
    store_name=store_match.group(1).strip() if store_match else "No indicada"
    if "coyoac" in store_name.casefold(): store_name="Coyoacán"
    local_now=datetime.now(ZoneInfo("America/Mexico_City"))
    fingerprint=hashlib.sha256(data).hexdigest()
    return {"fecha":local_now.strftime("%Y-%m-%d"),"cargado":local_now.strftime("%Y-%m-%d %H:%M:%S"),
            "periodo":period,"tienda":store_name,"archivo":_safe_filename(filename),
            "huella":fingerprint,"indicadores":indicators}

def normalize_followup_category(value):
    """Reduce los nombres históricos al catálogo operativo de tres categorías."""
    text=str(value or "").strip().casefold()
    if "tolerancia" in text or text in {"3", "0"}: return FOLLOWUP_CATEGORIES[2]
    if "acta" in text or text == "2": return FOLLOWUP_CATEGORIES[1]
    return FOLLOWUP_CATEGORIES[0]

def normalize_saved_followup_categories(store):
    """Migra etiquetas anteriores sin cambiar fechas, motivos, PDFs ni enlaces."""
    changed=False
    for rec in (store.get("pickers",{}) or {}).values():
        for item in (rec.get("acciones",[]) or []):
            old=item.get("accion",""); new=normalize_followup_category(old)
            if old!=new: item["accion"]=new; changed=True
        for item in (rec.get("documentos",[]) or []):
            old=item.get("tipo",""); new=normalize_followup_category(old)
            if old!=new: item["tipo"]=new; changed=True
    for item in (store.get("seguimientos_documentos",[]) or []):
        old=item.get("tipo",""); new=normalize_followup_category(old)
        if old!=new: item["tipo"]=new; changed=True
    for item in (store.get("recursos_formatos",[]) or []):
        old=item.get("tipo",""); new=normalize_followup_category(old)
        if old!=new: item["tipo"]=new; changed=True
    return changed

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
        current["comentarios"]=_unique_records(current["comentarios"]+rec.get("comentarios",[]))
        current["acciones"]=_unique_records(current["acciones"]+rec.get("acciones",[]))
        current["documentos"]=_unique_records(current["documentos"]+rec.get("documentos",[]))
        return current
    store["pickers"][target]=rec
    return rec

def followup_email_record(store,picker,email,aliases=None,migrate_legacy=False,create=False):
    """Expediente de Seguimiento por correo maestro, sin resolver por parecido de nombre.

    Los historiales antiguos se muestran/migran sólo cuando la vista confirma que
    el nombre del Maestro pertenece a una sola dirección de correo.
    """
    store.setdefault("pickers",{})
    target=str(picker or "").strip()
    ek=email_key(email)
    if not ek:
        return picker_record(store,target,aliases=aliases)
    identity_key="email:"+ek
    records=store["pickers"]
    rec=records.get(identity_key)
    if rec is None:
        rec=None
        if migrate_legacy:
            # Sólo exactos: nunca person_key/token_key ni coincidencia parcial.
            for candidate in [target]+[str(a).strip() for a in (aliases or []) if str(a).strip()]:
                if candidate in records and candidate!=identity_key:
                    rec=records.pop(candidate)
                    break
        if rec is None and create:
            rec={"estado":"ACTIVO","comentarios":[],"acciones":[],"documentos":[]}
        if rec is not None:
            records[identity_key]=rec
    if rec is None:
        return {"estado":"ACTIVO","comentarios":[],"acciones":[],"documentos":[]}
    rec.setdefault("estado","ACTIVO")
    rec.setdefault("comentarios",[])
    rec.setdefault("acciones",[])
    rec.setdefault("documentos",[])
    rec["identidad_plantilla"]={
        **(rec.get("identidad_plantilla",{}) or {}),
        "PICKER":target,"CORREO_KEY":ek,
    }
    return rec

def master_name_is_unique(identity_df,picker,email):
    """Indica si un historial histórico por nombre puede vincularse sin ambigüedad."""
    if identity_df is None or identity_df.empty:
        return False
    name_key=person_key(picker)
    ek=email_key(email)
    if not name_key or not ek or "PICKER" not in identity_df:
        return False
    matches=identity_df[identity_df["PICKER"].map(person_key)==name_key]
    return matches.get("CORREO_KEY",pd.Series(dtype=str)).map(email_key).nunique()==1 and bool(
        (matches.get("CORREO_KEY",pd.Series(dtype=str)).map(email_key)==ek).any()
    )

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
    # Algunos reportes de operación incluyen filas de cierre dentro de la
    # misma tabla (Total/Total general) y filas vacías rotuladas como
    # "Sin registrar". No son personas: si pasan al resumen se muestran como
    # pickers y duplican todas las líneas/pedidos del reporte.
    def _is_non_person_row(value):
        if extract_email_address(value):
            return False
        label=norm(value).replace("_"," ").strip()
        return label in {"", "total", "total general", "gran total", "grand total", "sin registrar", "picker"}
    x=x.loc[~x["PICKER"].map(_is_non_person_row)].copy()
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
        _existing_role_by_email={
            email_key(row.get("CORREO","")):str(row.get("PERSONAL_TIPO",""))
            for _,row in r.iterrows()
        } if not r.empty else {}
        for _,ov in overrides.items():
            _override_email=email_key(ov.get("CORREO",""))
            rows.append({
                "PICKER":str(ov.get("PICKER","")).strip(),
                "TURNO_MAESTRO":str(ov.get("TURNO","")).strip(),
                "CORREO":str(ov.get("CORREO","")).strip(),
                "SUPERVISOR":str(ov.get("SUPERVISOR","No asignado")).strip() or "No asignado",
                "AREA_MAESTRO":str(ov.get("AREA_BASE","")).strip(),
                "PERSONAL_TIPO":str(ov.get("PERSONAL_TIPO",_existing_role_by_email.get(_override_email,""))).strip(),
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

def canonicalize_incidents(inc,base):
    """Asocia FNR/MC por correo de forma vectorizada; fallback conservador solo en pendientes."""
    if inc is None or inc.empty:
        return inc
    out=inc.copy()
    roster=base.copy()
    roster["_EMAIL_KEY"]=roster.get("CORREO",pd.Series("",index=roster.index)).map(email_key)
    roster=roster[roster["_EMAIL_KEY"].astype(str).str.strip().ne("")].drop_duplicates("_EMAIL_KEY")
    name_map=dict(zip(roster["_EMAIL_KEY"],roster["PICKER"].astype(str).str.strip()))
    email_map=dict(zip(roster["_EMAIL_KEY"],roster["CORREO"].astype(str).str.strip()))

    out["_EMAIL_KEY"]=out.get("CORREO",pd.Series("",index=out.index)).map(email_key)
    mapped_name=out["_EMAIL_KEY"].map(name_map)
    mapped_email=out["_EMAIL_KEY"].map(email_map)
    matched=mapped_name.notna()

    out.loc[matched,"PICKER"]=mapped_name.loc[matched]
    out.loc[matched,"CORREO"]=mapped_email.loc[matched]

    unresolved=out.index[~matched]
    for idx in unresolved:
        row=out.loc[idx]
        rr=_match_master_row(row.get("PICKER",""),base,row.get("CORREO",""))
        if rr is None: continue
        out.at[idx,"PICKER"]=str(rr.get("PICKER",row.get("PICKER",""))).strip()
        out.at[idx,"CORREO"]=str(rr.get("CORREO",row.get("CORREO",""))).strip()
        matched.at[idx]=True

    out["_EMAIL_MATCH"]=matched.astype(bool)
    out.drop(columns=["_EMAIL_KEY"],errors="ignore",inplace=True)
    return out

def template_roster_from_upload(data):
    """Construye la identidad desde la PLANTILLA CONSOLIDADA.

    CORREO_KEY es la llave única. La plantilla puede tener información
    repartida entre Maestro_Personal y Base_Pickers (por ejemplo, un sheet
    tiene correo/nombre y otro tiene turno/supervisor/área). Por eso NO se
    elige una sola fila por correo: se consolidan los campos no vacíos de
    todas las hojas de la plantilla para que el turno nunca se pierda.
    """
    # Si el libro contiene Outbound OKI, usar únicamente esa pestaña como maestro.
    # Su encabezado se repite más abajo para el roster; se detecta por sus columnas.
    if isinstance(data,(bytes,bytearray)):
        try:
            _xls=pd.ExcelFile(io.BytesIO(data))
        except Exception:
            _xls=None
        _outbound_sheet=next((name for name in (_xls.sheet_names if _xls else []) if norm(name)=="outbound_oki"),None)
        if _outbound_sheet:
            _raw=pd.read_excel(io.BytesIO(data),sheet_name=_outbound_sheet,header=None,dtype=object)
            _header_idx=None
            _header_map={}
            _required={"puesto","turno","nombre","cargo","correo"}
            for _i in range(8,len(_raw)):
                _candidate={norm(v):j for j,v in enumerate(_raw.iloc[_i].tolist()) if pd.notna(v) and str(v).strip()}
                if _required.issubset(_candidate):
                    _header_idx=_i
                    _header_map=_candidate
                    break
            if _header_idx is None:
                raise ValueError("La hoja Outbound OKI no tiene los encabezados PUESTO, TURNO, NOMBRE, CARGO y CORREO.")
            _rows=[]
            def _shift_label(value):
                _key=norm(value)
                return {
                    "manana":"Mañana","matutino":"Mañana",
                    "intermedio":"Intermedio","tarde":"Tarde",
                    "noche":"Nocturno","nocturno":"Nocturno","mixto":"Mixto",
                }.get(_key,str(value or "").strip())

            # La misma pestaña incluye arriba el bloque TURNOS LÍDERES.
            # Incorporar sus personas nombradas, sin vacantes, usando USUARIO como correo.
            _leader_header_idx=None
            _leader_header_map={}
            _leader_required={"puesto","turno","nombre","cargo","usuario"}
            for _i in range(0,_header_idx):
                _candidate={norm(v):j for j,v in enumerate(_raw.iloc[_i].tolist()) if pd.notna(v) and str(v).strip()}
                if _leader_required.issubset(_candidate):
                    _leader_header_idx=_i
                    _leader_header_map=_candidate
                    break
            if _leader_header_idx is not None:
                for _i in range(_leader_header_idx+1,_header_idx):
                    def _leader_value(_field):
                        _v=_raw.iat[_i,_leader_header_map[_field]]
                        return "" if pd.isna(_v) else str(_v).strip()
                    _nombre=_leader_value("nombre")
                    if not _nombre or norm(_nombre) in {"vacante","nombre","nan","none"}:
                        continue
                    _role=_leader_value("cargo") or _leader_value("puesto")
                    _role_key=norm(_role).replace("_","")
                    _role_label={"supervisor":"Supervisor","sup":"Supervisor","subgerente":"Subgerente","jefe":"Jefe"}.get(_role_key)
                    if not _role_label:
                        continue
                    _rows.append({
                        "PICKER":_nombre,
                        "TURNO_MAESTRO":_shift_label(_leader_value("turno")),
                        "CORREO":_leader_value("usuario"),
                        "SUPERVISOR":"",
                        "AREA_MAESTRO":"Outbound OKI",
                        "PERSONAL_TIPO":_role_label,
                        "ESTADO_PLANTILLA":"",
                    })
            for _i in range(_header_idx+1,len(_raw)):
                def _value(_field):
                    _v=_raw.iat[_i,_header_map[_field]]
                    return "" if pd.isna(_v) else str(_v).strip()
                _cargo=_value("cargo")
                _cargo_key=norm(_cargo).replace("_","")
                if _cargo_key=="picker":
                    _tipo="Picker"
                elif _cargo_key=="layout":
                    _tipo="Layout"
                else:
                    continue
                _nombre=_value("nombre")
                if not _nombre or norm(_nombre) in {"nombre","nan","none"}:
                    continue
                _rows.append({
                    "PICKER":_nombre,
                    "TURNO_MAESTRO":_shift_label(_value("turno")),
                    # Algunas filas de Outbound OKI dejan CORREO vacío, pero
                    # sí tienen el correo corporativo en USUARIO. Usarlo como
                    # respaldo mantiene la llave de identidad y el expediente.
                    "CORREO":_value("correo") or _value("usuario"),
                    "SUPERVISOR":"",
                    "AREA_MAESTRO":"Outbound OKI",
                    "PERSONAL_TIPO":_tipo,
                    "ESTADO_PLANTILLA":"",
                })
            if not _rows:
                raise ValueError("No encontré filas con CARGO Picker o Lay Out en Outbound OKI.")
            r=pd.DataFrame(_rows)
            r["_EMAIL_KEY"]=r["CORREO"].map(email_key)
            # Con correo, la identidad se deduplica por email; cuando falta,
            # se conserva el registro para el cruce conservador por nombre.
            r["_ROSTER_KEY"]=r["_EMAIL_KEY"]
            _missing=r["_ROSTER_KEY"].eq("")
            r.loc[_missing,"_ROSTER_KEY"]="nombre:"+r.loc[_missing,"PICKER"].map(person_key)
            r=r[r["_ROSTER_KEY"].ne("nombre:")].drop_duplicates("_ROSTER_KEY",keep="first").drop(columns=["_ROSTER_KEY"]).reset_index(drop=True)
            r["_KEY"]=r["PICKER"].map(person_key)
            r["_TOKEN_KEY"]=r["PICKER"].map(token_key)
            r["_EMAIL_KEY"]=r["CORREO"].map(email_key)
            r["_CODE_KEY"]=r["CORREO"].map(extract_code)
            r["_EMAIL_TOKENS"]=r["CORREO"].map(email_name_tokens)
            return r

    ss=sheets_from_bytes(data) if isinstance(data,(bytes,bytearray)) else sheets(data)
    frames=[]
    for name,df in ss.items():
        n=norm(name)
        if n in {"instrucciones","diagnostico_cruce","diagnostico"}:
            continue
        em=col(df,["codigo_correo","codigo + correo","codigo correo","correo","email","usuario"])
        if not em:
            continue
        p=col(df,["picker","picker_nombre","email_picker","nombre","nombre completo","persona","colaborador","empleado"])
        sh=col(df,["turno","shift"])
        sup=col(df,["supervisor","supervisor_nombre","jefe","responsable"])
        ar=col(df,["area_base","area","departamento","department"])
        role=col(df,["tipo de personal","tipo personal","tipo de colaborador","tipo colaborador","puesto","cargo","rol","perfil","funcion"])
        estado=col(df,["estado","estatus","status","estatus_cruce"])
        _role_values=[]
        for _idx in df.index:
            _role_value=str(df.at[_idx,role]).strip() if role and pd.notna(df.at[_idx,role]) else ""
            _area_value=str(df.at[_idx,ar]).strip() if ar and pd.notna(df.at[_idx,ar]) else ""
            if any("layout" in re.sub(r"[^a-z0-9]", "", norm(value)) for value in (n,_area_value,_role_value)):
                _role_values.append("Layout")
            else:
                _role_values.append(_role_value)
        tmp=pd.DataFrame({
            "PICKER":df[p].fillna("").astype(str).str.strip() if p else "",
            "TURNO_MAESTRO":df[sh].fillna("").astype(str).str.strip() if sh else "",
            "CORREO":df[em].fillna("").astype(str).str.strip(),
            "SUPERVISOR":df[sup].fillna("").astype(str).str.strip() if sup else "",
            "AREA_MAESTRO":df[ar].fillna("").astype(str).str.strip() if ar else "",
            "PERSONAL_TIPO":_role_values,
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
        for field in ["PICKER","TURNO_MAESTRO","CORREO","SUPERVISOR","AREA_MAESTRO","PERSONAL_TIPO","ESTADO_PLANTILLA"]:
            vals=[v for v in g[field].tolist() if useful(v)]
            row[field]=str(vals[0]).strip() if vals else ""
        # Preferir el valor que aparezca más veces cuando existan varias fuentes.
        for field in ["PICKER","TURNO_MAESTRO","SUPERVISOR","AREA_MAESTRO","PERSONAL_TIPO","ESTADO_PLANTILLA"]:
            vals=[str(v).strip() for v in g[field].tolist() if useful(v)]
            if vals:
                if field=="PERSONAL_TIPO" and any("layout" in re.sub(r"[^a-z0-9]", "", norm(v)) for v in vals):
                    row[field]="Layout"
                    continue
                counts=pd.Series(vals).value_counts()
                row[field]=str(counts.index[0]).strip()
        # El correo canónico siempre sale del valor con @.
        emails=[str(v).strip() for v in g["CORREO"].tolist() if extract_email_address(v)]
        row["CORREO"]=emails[0] if emails else str(g["CORREO"].iloc[0]).strip()
        if not row.get("PERSONAL_TIPO"):
            row["PERSONAL_TIPO"]="Picker"
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
    sku=col(df,["sku","seller_sku","product_sku","codigo_sku","código sku","item_sku"])
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
        "SKU":df[sku].fillna("").astype(str).str.strip() if sku else "",
        "ORDER_NUMBER":df[order].astype(str).str.strip() if order else "",
        "AREA":df[ar].astype(str).str.strip() if ar else "No especificada",
        "TURNO":df[sh].astype(str).str.strip() if sh else "No especificado",
        "FECHA":df[fe] if fe else "",
        "INCIDENCIAS":pd.to_numeric(df[qty],errors="coerce").fillna(1) if qty else 1,
        "TIPO":tipo})
    # Los reportes operativos pueden repetir la fila de encabezados dentro de
    # la hoja. Sin este filtro, "PICKER" termina sumándose como persona/área.
    _header_or_total=x["PICKER"].map(lambda v:norm(v).replace("_"," ").strip() in {
        "picker", "picker nombre", "total", "total general", "gran total", "grand total"
    })
    x=x.loc[~_header_or_total].copy()
    x["INCIDENCIAS"]=x["INCIDENCIAS"].where(x["INCIDENCIAS"]>0,1)
    return x

def attach(inc,base):
    """Adjunta turno/área/supervisor por correo usando merge vectorizado."""
    if inc is None or inc.empty:
        return inc
    base=_dedupe_columns(base)
    ref_cols=[col for col in ["PICKER","TURNO","AREA_BASE","SUPERVISOR","CORREO","_EXCLUDED_PERSONNEL"] if col in base.columns]
    ref=base[ref_cols].copy()
    ref["_EMAIL_KEY"]=ref.get("CORREO",pd.Series("",index=ref.index)).map(email_key)
    ref=ref[ref["_EMAIL_KEY"].astype(str).str.strip().ne("")].drop_duplicates("_EMAIL_KEY")
    rename={
        "PICKER":"_REF_PICKER","TURNO":"_REF_TURNO","AREA_BASE":"_REF_AREA",
        "SUPERVISOR":"_REF_SUPERVISOR","CORREO":"_REF_CORREO",
        "_EXCLUDED_PERSONNEL":"_REF_EXCLUDED",
    }
    ref=ref.rename(columns=rename)

    y=inc.copy()
    y["_EMAIL_KEY"]=y.get("CORREO",pd.Series("",index=y.index)).map(email_key)
    y["_ORIG_INDEX"]=np.arange(len(y))
    y["TURNO_REF"]=y.get("TURNO",pd.Series("No especificado",index=y.index)).fillna("No especificado").astype(str).str.strip()
    y["AREA_REF"]=y.get("AREA",pd.Series("No especificada",index=y.index)).fillna("No especificada").astype(str).str.strip()
    y=y.merge(ref,on="_EMAIL_KEY",how="left",sort=False)
    y=y.sort_values("_ORIG_INDEX").reset_index(drop=True)

    matched=y.get("_REF_PICKER",pd.Series(pd.NA,index=y.index)).notna()
    y["_TEMPLATE_MATCH"]=matched
    if matched.any():
        y.loc[matched,"PICKER"]=y.loc[matched,"_REF_PICKER"].astype(str).str.strip()
        y.loc[matched,"CORREO"]=y.loc[matched,"_REF_CORREO"].astype(str).str.strip()
        good_turn=matched & y["_REF_TURNO"].fillna("").astype(str).str.strip().ne("") & ~y["_REF_TURNO"].fillna("").astype(str).isin(["No especificado","nan","None"])
        y.loc[good_turn,"TURNO"]=y.loc[good_turn,"_REF_TURNO"].astype(str).str.strip()
        y.loc[good_turn,"TURNO_REF"]=y.loc[good_turn,"_REF_TURNO"].astype(str).str.strip()
        good_area=matched & y["_REF_AREA"].fillna("").astype(str).str.strip().ne("") & ~y["_REF_AREA"].fillna("").astype(str).isin(["No especificada","nan","None"])
        y.loc[good_area,"AREA"]=y.loc[good_area,"_REF_AREA"].astype(str).str.strip()
        y.loc[good_area,"AREA_REF"]=y.loc[good_area,"_REF_AREA"].astype(str).str.strip()
        good_sup=matched & y["_REF_SUPERVISOR"].fillna("").astype(str).str.strip().ne("")
        y.loc[good_sup,"SUPERVISOR_REF"]=y.loc[good_sup,"_REF_SUPERVISOR"].astype(str).str.strip()
        if "_REF_EXCLUDED" in y.columns:
            y["_EXCLUDED_PERSONNEL"]=y["_REF_EXCLUDED"].fillna(False).astype(bool)

    source_match=y.get("_EMAIL_MATCH",pd.Series(False,index=y.index)).fillna(False).astype(bool)
    y["CATEGORIA"]=np.where(source_match|matched,"Picker","Sin registrar")
    return y.drop(columns=[
        "_EMAIL_KEY","_ORIG_INDEX","_REF_PICKER","_REF_TURNO","_REF_AREA",
        "_REF_SUPERVISOR","_REF_CORREO","_REF_EXCLUDED",
    ],errors="ignore")

@st.cache_data(show_spinner="Procesando los Excel por primera vez…",max_entries=2,ttl=1800)
def prepare_source_frames(base_data,fnr_data,mc_data,roster_data,personnel_config_json):
    """Procesa y cruza las fuentes una vez por versión de archivos/configuración."""
    config=json.loads(personnel_config_json)
    base=parse_base(choose(sheets_from_bytes(base_data),["picker","lineas","resumen"]))
    fnr=parse_inc(choose(sheets_from_bytes(fnr_data),["fnr","detalle"]),"FNR")
    mc=parse_inc(choose(sheets_from_bytes(mc_data),["mc","mala"]),"MC")
    roster=template_roster_from_upload(roster_data)
    if roster.empty:
        raise ValueError("La plantilla consolidada no contiene correos válidos.")
    base=apply_roster(base,roster)
    base=apply_manual_personnel(base,config)
    base_match=(base.get("_MASTER_MATCH",pd.Series(False,index=base.index)).fillna(False).astype(bool)
                | base.get("_MANUAL_MATCH",pd.Series(False,index=base.index)).fillna(False).astype(bool))
    base["CATEGORIA"]=np.where(base_match,"Picker","Sin registrar")
    base=_dedupe_columns(base)
    roster=effective_roster(roster,config)
    fnr=canonicalize_incidents(fnr,roster)
    mc=canonicalize_incidents(mc,roster)
    identity=roster.rename(columns={"TURNO_MAESTRO":"TURNO","AREA_MAESTRO":"AREA_BASE"}).copy()
    fnr=attach(fnr,identity); mc=attach(mc,identity)
    fnr["CORREO_KEY"]=fnr.get("CORREO",pd.Series("",index=fnr.index)).map(email_key)
    mc["CORREO_KEY"]=mc.get("CORREO",pd.Series("",index=mc.index)).map(email_key)
    fnr=_dedupe_columns(fnr); mc=_dedupe_columns(mc)
    return base,fnr,mc,roster

@st.cache_data(show_spinner="Procesando histórico del mes anterior…",max_entries=4,ttl=1800)
def prepare_history_frames(base_data,fnr_data,mc_data,roster,personnel_config_json):
    """Procesa los mismos 3 Excel operativos de un periodo anterior.

    Se usa la plantilla consolidada actual como identidad canónica para poder
    comparar al mismo picker entre periodos sin pedir un cuarto archivo.
    """
    config=json.loads(personnel_config_json)
    hist_base=parse_base(choose(sheets_from_bytes(base_data),["picker","lineas","resumen"]))
    hist_fnr=parse_inc(choose(sheets_from_bytes(fnr_data),["fnr","detalle"]),"FNR")
    hist_mc=parse_inc(choose(sheets_from_bytes(mc_data),["mc","mala"]),"MC")
    hist_base=apply_roster(hist_base,roster)
    hist_base=apply_manual_personnel(hist_base,config)
    matched=(hist_base.get("_MASTER_MATCH",pd.Series(False,index=hist_base.index)).fillna(False).astype(bool)
             | hist_base.get("_MANUAL_MATCH",pd.Series(False,index=hist_base.index)).fillna(False).astype(bool))
    hist_base["CATEGORIA"]=np.where(matched,"Picker","Sin registrar")
    hist_base=_dedupe_columns(hist_base)

    hist_fnr=canonicalize_incidents(hist_fnr,roster)
    hist_mc=canonicalize_incidents(hist_mc,roster)
    identity=roster.rename(columns={"TURNO_MAESTRO":"TURNO","AREA_MAESTRO":"AREA_BASE"}).copy()
    hist_fnr=attach(hist_fnr,identity)
    hist_mc=attach(hist_mc,identity)
    hist_fnr["CORREO_KEY"]=hist_fnr.get("CORREO",pd.Series("",index=hist_fnr.index)).map(email_key)
    hist_mc["CORREO_KEY"]=hist_mc.get("CORREO",pd.Series("",index=hist_mc.index)).map(email_key)
    return _dedupe_columns(hist_base),_dedupe_columns(hist_fnr),_dedupe_columns(hist_mc)

@st.cache_data(show_spinner=False,max_entries=8,ttl=1800)
def summary(base,fnr,mc):
    """Resumen por identidad canónica: correo cuando existe, nombre solo como fallback."""
    def prepare(frame,is_incident=False):
        if frame is None or frame.empty:
            return pd.DataFrame()
        q=frame.copy()
        if "CATEGORIA" not in q:
            if is_incident:
                matched=q.get("_EMAIL_MATCH",q.get("_TEMPLATE_MATCH",pd.Series(False,index=q.index))).fillna(False).astype(bool)
            else:
                matched=q.get("_MASTER_MATCH",pd.Series(False,index=q.index)).fillna(False).astype(bool) | q.get("_MANUAL_MATCH",pd.Series(False,index=q.index)).fillna(False).astype(bool)
            q["CATEGORIA"]=np.where(matched,"Picker","Sin registrar")
        q["PICKER"]=q.get("PICKER",pd.Series("Sin registrar",index=q.index)).fillna("").astype(str).str.strip()
        q.loc[q["PICKER"].eq(""),"PICKER"]="Sin registrar"
        q["CORREO"]=q.get("CORREO",pd.Series("",index=q.index)).fillna("").astype(str).str.strip()
        q["CORREO_KEY"]=q["CORREO"].map(email_key)
        q["_PERSON_KEY"]=q.apply(person_context_key,axis=1)
        return q

    b=prepare(base,False)
    for field in ["LINEAS","PEDIDOS"]:
        if field not in b: b[field]=0.0
        b[field]=pd.to_numeric(b[field],errors="coerce").fillna(0.0)

    fi=prepare(fnr,True)
    mi=prepare(mc,True)
    meta=[b]
    for q in [fi,mi]:
        if q.empty: continue
        q=q.copy()
        q["LINEAS"]=0.0; q["PEDIDOS"]=0.0
        q["AREA_BASE"]=q.get("AREA_BASE",q.get("AREA",pd.Series("No especificada",index=q.index)))
        q["SUPERVISOR"]=q.get("SUPERVISOR",q.get("SUPERVISOR_REF",pd.Series("No asignado",index=q.index)))
        meta.append(q)

    meta_df=pd.concat(meta,ignore_index=True,sort=False)
    keys=["CATEGORIA","_PERSON_KEY"]
    metadata_fields=[col for col in ["PICKER","CORREO","CORREO_KEY","TURNO","SUPERVISOR","AREA_BASE"] if col in meta_df.columns]
    agg={"LINEAS":"sum","PEDIDOS":"sum"}
    def _first_value(values):
        return next((str(v).strip() for v in values if str(v).strip() and str(v).strip().lower() not in {"nan","none"}),"")
    agg.update({field:_first_value for field in metadata_fields})
    for flag in ["_EXCLUDED_PERSONNEL","_MASTER_MATCH","_MANUAL_MATCH"]:
        if flag in meta_df.columns:
            meta_df[flag]=meta_df[flag].astype("boolean").fillna(False).astype(bool)
            agg[flag]="any"
    if "_SOURCE_PICKER" in meta_df.columns:
        agg["_SOURCE_PICKER"]=_first_value
    s=meta_df.groupby(keys,as_index=False,dropna=False).agg(agg)

    if not fi.empty:
        f=fi.groupby(keys,as_index=False,dropna=False).INCIDENCIAS.sum().rename(columns={"INCIDENCIAS":"FNR"})
    else:
        f=pd.DataFrame(columns=keys+["FNR"])
    if not mi.empty:
        m=mi.groupby(keys,as_index=False,dropna=False).INCIDENCIAS.sum().rename(columns={"INCIDENCIAS":"MC"})
    else:
        m=pd.DataFrame(columns=keys+["MC"])

    s=s.merge(f,on=keys,how="outer").merge(m,on=keys,how="outer")
    s[["FNR","MC"]]=s[["FNR","MC"]].fillna(0)
    for field in ["LINEAS","PEDIDOS","FNR","MC"]:
        s[field]=pd.to_numeric(s.get(field,0),errors="coerce").fillna(0.0)
    den=s["LINEAS"].where(s["LINEAS"].ne(0),float("nan"))
    s["FNR_%"]=pd.to_numeric(s["FNR"].div(den)*100,errors="coerce").round(2)
    s["MC_%"]=pd.to_numeric(s["MC"].div(den)*100,errors="coerce").round(2)
    def sem(row):
        if (pd.notna(row["FNR_%"]) and row["FNR_%"]>=FNR_OBJ) or (pd.notna(row["MC_%"]) and row["MC_%"]>=MC_OBJ):
            return "🔴 FUERA DE OBJETIVO"
        if (pd.notna(row["FNR_%"]) and row["FNR_%"]>=1.20) or (pd.notna(row["MC_%"]) and row["MC_%"]>=0.80):
            return "🟡 PREVENTIVO"
        return "🟢 EN OBJETIVO"
    s["ESTADO"]=s.apply(sem,axis=1)
    return s.drop(columns=["_PERSON_KEY"],errors="ignore")

def monthly_history_snapshot(label,summary_df,source_files=None):
    """Convierte un resumen mensual en un snapshot JSON pequeño y persistente."""
    x=summary_df.copy()
    keep=[col for col in ["CATEGORIA","PICKER","CORREO","TURNO","SUPERVISOR","AREA_BASE","LINEAS","PEDIDOS","FNR","MC"] if col in x.columns]
    x=x[keep].copy()
    for colname in ["LINEAS","PEDIDOS","FNR","MC"]:
        if colname not in x.columns: x[colname]=0.0
        x[colname]=pd.to_numeric(x[colname],errors="coerce").fillna(0.0)
    lines=float(x["LINEAS"].sum())
    fnr_total=float(x["FNR"].sum())
    mc_total=float(x["MC"].sum())
    pedidos=float(x["PEDIDOS"].sum())
    snapshot={
        "periodo":str(label).strip() or "Periodo anterior",
        "guardado":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "archivos":source_files or {},
        "totales":{
            "lineas":lines,
            "pedidos":pedidos,
            "fnr":fnr_total,
            "mc":mc_total,
            "fnr_pct":(fnr_total/lines*100) if lines else None,
            "mc_pct":(mc_total/lines*100) if lines else None,
            "incidencias":fnr_total+mc_total,
            "incidencias_1000":((fnr_total+mc_total)/lines*1000) if lines else None,
        },
        "pickers":json.loads(x.to_json(orient="records",force_ascii=False)),
    }
    return snapshot

@st.cache_data(show_spinner=False,max_entries=12,ttl=1800)
def monthly_picker_comparison(current_summary,previous_rows):
    """Compara producción e incidencias por identidad canónica entre dos periodos."""
    def rollup(frame,suffix):
        x=frame.copy() if isinstance(frame,pd.DataFrame) else pd.DataFrame(frame or [])
        if x.empty:
            return pd.DataFrame(columns=["_CMP_KEY",f"Picker {suffix}",f"Líneas {suffix}",f"FNR {suffix}",f"MC {suffix}"])
        for colname in ["PICKER","CORREO","LINEAS","FNR","MC"]:
            if colname not in x.columns:
                x[colname]="" if colname in {"PICKER","CORREO"} else 0.0
        x["PICKER"]=x["PICKER"].fillna("").astype(str).str.strip()
        x["CORREO"]=x["CORREO"].fillna("").astype(str).str.strip()
        _email=x["CORREO"].map(email_key)
        _name=x["PICKER"].map(person_key)
        x["_CMP_KEY"]=[
            ("email:"+ek) if ek else (("name:"+nk) if nk else ("row:"+str(idx)))
            for idx,(ek,nk) in enumerate(zip(_email,_name))
        ]
        for colname in ["LINEAS","FNR","MC"]:
            x[colname]=pd.to_numeric(x[colname],errors="coerce").fillna(0.0)
        rows=[]
        for key,g in x.groupby("_CMP_KEY",sort=False):
            name=next((str(v).strip() for v in g["PICKER"] if str(v).strip()),"Sin registrar")
            rows.append({
                "_CMP_KEY":key,
                f"Picker {suffix}":name,
                f"Líneas {suffix}":float(g["LINEAS"].sum()),
                f"FNR {suffix}":float(g["FNR"].sum()),
                f"MC {suffix}":float(g["MC"].sum()),
            })
        return pd.DataFrame(rows)

    cur=rollup(current_summary,"actual")
    prev=rollup(pd.DataFrame(previous_rows or []),"anterior")
    out=prev.merge(cur,on="_CMP_KEY",how="outer")
    for colname in ["Picker anterior","Picker actual"]:
        if colname not in out.columns: out[colname]=""
        out[colname]=out[colname].fillna("").astype(str)
    out["Picker"]=out["Picker actual"].where(out["Picker actual"].str.strip().ne(""),out["Picker anterior"])
    for colname in ["Líneas anterior","Líneas actual","FNR anterior","FNR actual","MC anterior","MC actual"]:
        if colname not in out.columns: out[colname]=0.0
        out[colname]=pd.to_numeric(out[colname],errors="coerce").fillna(0.0)

    out["Δ líneas"]=out["Líneas actual"]-out["Líneas anterior"]
    out["% Δ líneas"]=np.where(out["Líneas anterior"]>0,out["Δ líneas"]/out["Líneas anterior"]*100,np.nan)
    out["Δ FNR"]=out["FNR actual"]-out["FNR anterior"]
    out["Δ MC"]=out["MC actual"]-out["MC anterior"]
    out["Incidencias anterior"]=out["FNR anterior"]+out["MC anterior"]
    out["Incidencias actual"]=out["FNR actual"]+out["MC actual"]
    out["Δ incidencias"]=out["Incidencias actual"]-out["Incidencias anterior"]

    out["FNR % anterior"]=np.where(out["Líneas anterior"]>0,out["FNR anterior"]/out["Líneas anterior"]*100,np.nan)
    out["FNR % actual"]=np.where(out["Líneas actual"]>0,out["FNR actual"]/out["Líneas actual"]*100,np.nan)
    out["Δ FNR pp"]=out["FNR % actual"]-out["FNR % anterior"]
    out["MC % anterior"]=np.where(out["Líneas anterior"]>0,out["MC anterior"]/out["Líneas anterior"]*100,np.nan)
    out["MC % actual"]=np.where(out["Líneas actual"]>0,out["MC actual"]/out["Líneas actual"]*100,np.nan)
    out["Δ MC pp"]=out["MC % actual"]-out["MC % anterior"]
    out["Inc./1000 anterior"]=np.where(out["Líneas anterior"]>0,out["Incidencias anterior"]/out["Líneas anterior"]*1000,np.nan)
    out["Inc./1000 actual"]=np.where(out["Líneas actual"]>0,out["Incidencias actual"]/out["Líneas actual"]*1000,np.nan)
    out["Δ Inc./1000"]=out["Inc./1000 actual"]-out["Inc./1000 anterior"]

    def trend(row):
        inc_delta=float(row.get("Δ incidencias",0))
        line_delta=float(row.get("Δ líneas",0))
        rate_delta=row.get("Δ Inc./1000",np.nan)
        if inc_delta<0 and line_delta>=0:
            return "🟢 Más/igual líneas y menos incidencias"
        if inc_delta>0 and line_delta<=0:
            return "🔴 Menos/igual líneas y más incidencias"
        if pd.notna(rate_delta) and rate_delta<0:
            return "🟢 Menor tasa de incidencias"
        if pd.notna(rate_delta) and rate_delta>0:
            return "🟠 Mayor tasa de incidencias"
        if line_delta>0:
            return "🔵 Más líneas, incidencia estable"
        if line_delta<0:
            return "⚪ Menos líneas, incidencia estable"
        return "⚪ Sin cambio"
    out["Tendencia"]=out.apply(trend,axis=1)

    numeric_round=[
        "% Δ líneas","FNR % anterior","FNR % actual","Δ FNR pp",
        "MC % anterior","MC % actual","Δ MC pp",
        "Inc./1000 anterior","Inc./1000 actual","Δ Inc./1000",
    ]
    for colname in numeric_round:
        out[colname]=pd.to_numeric(out[colname],errors="coerce").round(2)
    return out.sort_values(["Δ Inc./1000","Δ incidencias"],ascending=[False,False],na_position="last").reset_index(drop=True)

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


def _audit_item_key(sku="", product=""):
    sku_text="" if pd.isna(sku) else str(sku).strip()
    if sku_text and sku_text.casefold() not in {"nan","none","<na>"}:
        return "SKU:"+re.sub(r"[^a-z0-9]", "", sku_text.casefold())
    product_key=norm(product)
    return "PRODUCTO:"+product_key if product_key else ""



def _audit_incident_date(value):
    """Fecha real del evento FNR/MC, no fecha de subida del archivo."""
    if value is None or (not isinstance(value,(list,dict)) and pd.isna(value)):
        return None
    if isinstance(value,(datetime,pd.Timestamp)):
        return value.date()
    if isinstance(value,(int,float,np.integer,np.floating)):
        number=float(value)
        if 20000<=number<=100000:
            try: return pd.to_datetime(number,unit="D",origin="1899-12-30").date()
            except (ValueError,OverflowError): return None
        text=str(int(number)) if number.is_integer() else str(number)
    else:
        text=str(value).strip()
    if not text or text.lower() in {"nan","nat","none","sin fecha","0"}:
        return None
    # Excel exporta ISO y también fechas de México dd/mm/aaaa.
    if re.fullmatch(r"\d{8}",text):
        try: return datetime.strptime(text,"%Y%m%d").date()
        except ValueError: return None
    # ISO AAAA-MM-DD siempre se interpreta como año-mes-día, incluso
    # cuando día y mes son ambos <=12 (pandas dayfirst invierte ambiguos).
    if re.match(r"^\d{4}-\d{2}-\d{2}(?:[T\s]|$)",text):
        try: return datetime.strptime(text[:10],"%Y-%m-%d").date()
        except ValueError: return None
    try:
        ts=pd.to_datetime(text,dayfirst=True,errors="coerce")
        return ts.date() if pd.notna(ts) else None
    except (ValueError,TypeError,OverflowError):
        return None


def _audit_incident_fingerprint(row):
    """Identidad operativa de la incidencia para no contar cada recarga de nuevo."""
    email=email_key(row.get("CORREO",""))
    person=("e:"+email) if email else ("n:"+token_key(row.get("PICKER","")))
    order=norm(row.get("ORDER_NUMBER",""))
    product=_audit_item_key(row.get("SKU",""),row.get("PRODUCTO",""))
    amount=pd.to_numeric(row.get("INCIDENCIAS",1),errors="coerce")
    amount=1.0 if pd.isna(amount) or amount<=0 else float(amount)
    return (
        str(row.get("_AUDIT_DATE","")),person,order,product,
        norm(row.get("AREA","")),amount,
    )


def audit_dedupe_30day_versions(versions,today):
    """Unión por máxima multiplicidad por archivo: evita duplicar Excel acumulativos.

    Cada exportación puede repetirse total o parcialmente; para una misma
    fecha/pedido/persona/artículo se conserva el máximo de filas presentes
    en una versión, sin eliminar dos filas idénticas dentro de un mismo Excel.
    Sin ID único de incidencia no se puede distinguir al 100% dos eventos
    independientes con todos sus campos iguales.
    """
    start=today-timedelta(days=30)
    records={}
    max_occurrences={}
    skipped_dates=0
    skipped_outside=0
    source_rows=0
    dated_rows=0
    for version in versions:
        if version is None or version.empty:
            continue
        occurrences={}
        for _,row in version.iterrows():
            source_rows+=1
            incident_day=_audit_incident_date(row.get("FECHA"))
            if incident_day is None:
                skipped_dates+=1
                continue
            if not (start<=incident_day<today):
                skipped_outside+=1
                continue
            dated_rows+=1
            saved=dict(row)
            saved["_AUDIT_DATE"]=incident_day.isoformat()
            key=_audit_incident_fingerprint(saved)
            occurrences[key]=occurrences.get(key,0)+1
            occurrence_no=occurrences[key]
            if occurrence_no>max_occurrences.get(key,0):
                max_occurrences[key]=occurrence_no
                records[(key,occurrence_no)]=saved
    fields=["PICKER","CORREO","PRODUCTO","SKU","ORDER_NUMBER","AREA","TURNO","FECHA","INCIDENCIAS","TIPO","_AUDIT_DATE"]
    result=pd.DataFrame(list(records.values()),columns=fields)
    info={"versiones":len(versions),"filas_leidas":source_rows,"fechas_ausentes":skipped_dates,
          "fuera_de_periodo":skipped_outside,"filas_con_fecha_valida":dated_rows,
          "registros_unicos":len(result)}
    return result,info


@st.cache_data(show_spinner=False,ttl=120,max_entries=4)
def _audit_history_archive_names(key,today_iso):
    """Versiones que pudieron contener incidentes de los últimos 30 días."""
    cutoff=datetime.fromisoformat(today_iso)-timedelta(days=31)
    names=[]
    prefix=f"upload_history/{key}"
    if cloud_enabled():
        offset=0
        while True:
            batch=cloud_list(prefix,limit=500,offset=offset)
            if not isinstance(batch,list):
                raise RuntimeError("Supabase no pudo listar las versiones históricas.")
            for entry in batch:
                name=str(entry.get("name","")).strip()
                if not name: continue
                m=re.match(r"^(\d{8})_\d{6}",name)
                if m:
                    stamp=datetime.strptime(m.group(1),"%Y%m%d")
                    if stamp.date()<cutoff.date():
                        continue
                names.append(name)
            if len(batch)<500: break
            offset+=len(batch)
    else:
        folder=os.path.join(UPLOAD_HISTORY_DIR,key)
        if os.path.isdir(folder):
            for name in os.listdir(folder):
                m=re.match(r"^(\d{8})_\d{6}",name)
                if m and datetime.strptime(m.group(1),"%Y%m%d").date()<cutoff.date():
                    continue
                names.append(name)
    return sorted(set(names),reverse=True)


@st.cache_data(show_spinner=False,ttl=1800,max_entries=100)
def _audit_parse_history_bytes(data,tipo):
    words=["fnr","detalle"] if tipo=="FNR" else ["mc","mala"]
    return parse_inc(choose(sheets_from_bytes(data),words),tipo)


@st.cache_data(show_spinner=False,ttl=1800,max_entries=120)
def _audit_archive_parsed_version(key,filename,tipo):
    """Una sola descarga por versión durante el cacheo, no cada vez que cambia un filtro."""
    if cloud_enabled():
        raw=cloud_download(f"upload_history/{key}/{filename}",missing_ok=False)
    else:
        with open(os.path.join(UPLOAD_HISTORY_DIR,key,filename),"rb") as f:
            raw=f.read()
    if raw is None:
        raise RuntimeError(f"No se pudo recuperar la versión {filename}")
    return hashlib.sha256(raw).hexdigest(),_audit_parse_history_bytes(raw,tipo)


def audit_history_30days(current_fnr,current_mc,roster,base,store,today=None):
    """Fuente exclusiva para el riesgo histórico de auditoría: 30 días completos previos."""
    today=today or datetime.now(ZoneInfo("America/Mexico_City")).date()
    result={}
    metadata={}
    for key,tipo,now_bytes in [
        ("detalle_fnr","FNR",current_fnr),
        ("detalle_mc","MC",current_mc),
    ]:
        versions=[]
        seen_hash=set()
        # Incluir la versión actual incluso si todavía no aparece en el índice remoto.
        current_hash=hashlib.sha256(now_bytes).hexdigest()
        versions.append(canonicalize_incidents(_audit_parse_history_bytes(now_bytes,tipo),roster))
        seen_hash.add(current_hash)
        names=_audit_history_archive_names(key,today.isoformat())
        for name in names:
            # Cachea el archivo y omite copias exactas de una versión anterior.
            fingerprint,parsed=_audit_archive_parsed_version(key,name,tipo)
            if fingerprint in seen_hash:
                continue
            seen_hash.add(fingerprint)
            versions.append(canonicalize_incidents(parsed,roster))
        frame,stats=audit_dedupe_30day_versions(versions,today)
        if not frame.empty:
            identity=roster.rename(columns={"TURNO_MAESTRO":"TURNO","AREA_MAESTRO":"AREA_BASE"}).copy()
            frame=attach(frame,identity)
            frame["CORREO_KEY"]=frame.get("CORREO",pd.Series("",index=frame.index)).map(email_key)
        result[tipo]=frame
        metadata[tipo]=stats
    _,result["FNR"],result["MC"]=exclude_registered_supervisors(
        base,result["FNR"],result["MC"],store
    )
    return result["FNR"],result["MC"],metadata


def audit_record_is_complete(record):
    """Solo una revisión confirmada puede retirar un pedido de la lista pendiente."""
    if not isinstance(record,dict):
        return False
    result=str(record.get("resultado","")).strip().casefold()
    if result in {"no se pudo validar","pendiente","sin validar",""}:
        return False
    checked=record.get("pedido_validado",None)
    if checked is None:
        # Compatibilidad con las revisiones antiguas que carecían del indicador.
        return result in {"pedido correcto","con diferencias"}
    if isinstance(checked,str):
        return checked.strip().casefold() in {"true","1","yes","si","sí","verdadero"}
    return bool(checked)


def order_risk_analysis(audit_lines,picker_summary,fnr_inc,mc_inc,prior_audits=None):
    """Prioriza pedidos por coincidencias históricas concretas de picker y artículo.

    Reglas principales:
    - Picker con más de 5 incidencias históricas FNR+MC.
    - Artículo con al menos 5 FNR o al menos 5 MC.
    - La coincidencia picker + artículo en el mismo renglón recibe el mayor peso.
    El puntaje es una regla operativa de foco, no una probabilidad estadística.
    """
    lines=audit_lines.copy()
    prior_audits=prior_audits or []
    PICKER_THRESHOLD=5
    PICKER_EXTRA_THRESHOLD=10
    ITEM_THRESHOLD=5

    def _incident_amount(row):
        value=pd.to_numeric(row.get("INCIDENCIAS",1),errors="coerce")
        return 1.0 if pd.isna(value) or value<=0 else float(value)

    def picker_counts(inc):
        counts={}
        if inc is None or inc.empty:
            return counts
        for _,row in inc.iterrows():
            key=person_key(row.get("PICKER",""))
            if key:
                counts[key]=counts.get(key,0.0)+_incident_amount(row)
        return counts

    def product_counts(inc):
        counts={}
        if inc is None or inc.empty:
            return counts
        for _,row in inc.iterrows():
            keys={
                _audit_item_key(row.get("SKU",""),row.get("PRODUCTO","")),
                _audit_item_key("",row.get("PRODUCTO","")),
            }
            keys.discard("")
            amount=_incident_amount(row)
            for key in keys:
                counts[key]=counts.get(key,0.0)+amount
        return counts

    def order_counts(inc):
        counts={}
        if inc is None or inc.empty or "ORDER_NUMBER" not in inc.columns:
            return counts
        for _,row in inc.iterrows():
            key=norm(row.get("ORDER_NUMBER",""))
            if key:
                counts[key]=counts.get(key,0.0)+_incident_amount(row)
        return counts

    def picker_product_counts(inc):
        """Antecedentes donde el mismo picker y el mismo artículo coincidieron."""
        counts={}
        if inc is None or inc.empty:
            return counts
        for _,row in inc.iterrows():
            picker_key=person_key(row.get("PICKER",""))
            product_key=_audit_item_key("",row.get("PRODUCTO",""))
            if picker_key and product_key:
                pair=(picker_key,product_key)
                counts[pair]=counts.get(pair,0.0)+_incident_amount(row)
        return counts

    fnr_picker=picker_counts(fnr_inc)
    mc_picker=picker_counts(mc_inc)
    fnr_products=product_counts(fnr_inc)
    mc_products=product_counts(mc_inc)
    fnr_orders=order_counts(fnr_inc)
    mc_orders=order_counts(mc_inc)
    fnr_picker_product=picker_product_counts(fnr_inc)
    mc_picker_product=picker_product_counts(mc_inc)

    prior_by_order={}
    for audit in prior_audits:
        if not audit_record_is_complete(audit):
            continue
        key=norm(audit.get("pedido",""))
        if not key:
            continue
        previous=prior_by_order.get(key)
        if previous is None or str(audit.get("fecha_auditoria",""))>=str(previous.get("fecha_auditoria","")):
            prior_by_order[key]=audit

    # Aprendizaje descriptivo: solo se activa con una muestra mínima.
    # Usa incidencias posteriores a la auditoría como señal observada; no implica causalidad.
    learning=[]
    for order_key,audit in prior_by_order.items():
        if not isinstance(audit,dict): continue
        saved_fnr=pd.to_numeric(audit.get("fnr_al_guardar",0),errors="coerce")
        saved_mc=pd.to_numeric(audit.get("mc_al_guardar",0),errors="coerce")
        saved=(0.0 if pd.isna(saved_fnr) else float(saved_fnr))+(0.0 if pd.isna(saved_mc) else float(saved_mc))
        current=float(fnr_orders.get(order_key,0.0))+float(mc_orders.get(order_key,0.0))
        learning.append({
            "hit":bool(current>saved),
            "picker":bool(audit.get("picker_riesgo",False)),
            "item":bool(audit.get("articulo_fnr_riesgo",False) or audit.get("articulo_mc_riesgo",False)),
            "combo":bool(audit.get("coincidencia_picker_articulo",False)),
        })

    baseline=(sum(1 for z in learning if z["hit"])/len(learning)) if learning else None
    def _learned_bonus(field):
        subset=[z for z in learning if z.get(field)]
        if len(subset)<10 or baseline is None:
            return 0, len(subset), None
        rate=sum(1 for z in subset if z["hit"])/len(subset)
        uplift=rate-baseline
        if rate>=baseline*1.8 and uplift>=0.15:
            return 2,len(subset),rate
        if rate>=baseline*1.3 and uplift>=0.08:
            return 1,len(subset),rate
        return 0,len(subset),rate

    learned_picker,learn_n_picker,learn_rate_picker=_learned_bonus("picker")
    learned_item,learn_n_item,learn_rate_item=_learned_bonus("item")
    learned_combo,learn_n_combo,learn_rate_combo=_learned_bonus("combo")

    rows=[]
    for _,row in lines.iterrows():
        _picker_value=row.get("Picker","")
        picker_name="" if pd.isna(_picker_value) else str(_picker_value).strip()
        picker_key=person_key(picker_name)
        picker_fnr=float(fnr_picker.get(picker_key,0.0))
        picker_mc=float(mc_picker.get(picker_key,0.0))
        picker_total=picker_fnr+picker_mc
        picker_signal=picker_total>PICKER_THRESHOLD

        item_key=_audit_item_key(row.get("SKU",""),row.get("Artículo",""))
        product_key=_audit_item_key("",row.get("Artículo",""))
        fnr_count=float(fnr_products.get(product_key,fnr_products.get(item_key,0.0)))
        mc_count=float(mc_products.get(product_key,mc_products.get(item_key,0.0)))
        fnr_item_signal=fnr_count>=ITEM_THRESHOLD
        mc_item_signal=mc_count>=ITEM_THRESHOLD
        item_signal=fnr_item_signal or mc_item_signal
        coincidence=picker_signal and item_signal

        pair_fnr=float(fnr_picker_product.get((picker_key,product_key),0.0)) if picker_key and product_key else 0.0
        pair_mc=float(mc_picker_product.get((picker_key,product_key),0.0)) if picker_key and product_key else 0.0
        pair_total=pair_fnr+pair_mc
        pair_signal=pair_total>=2

        ordered=pd.to_numeric(str(row.get("Cantidad pedida","")).replace(",",""),errors="coerce")
        picked=pd.to_numeric(str(row.get("Cantidad pickeada","")).replace(",",""),errors="coerce")
        quantity_mismatch=bool(pd.notna(ordered) and pd.notna(picked) and abs(float(ordered)-float(picked))>1e-9)
        quantity_delta=(float(picked)-float(ordered)) if quantity_mismatch else 0.0

        signal_parts=[]
        if picker_signal:
            signal_parts.append(f"Picker {picker_total:g} inc.")
        if fnr_item_signal:
            signal_parts.append(f"Artículo FNR {fnr_count:g}")
        if mc_item_signal:
            signal_parts.append(f"Artículo MC {mc_count:g}")
        if pair_signal:
            signal_parts.append(f"Mismo picker+artículo {pair_total:g} veces")
        if quantity_mismatch:
            signal_parts.append(f"Cantidad distinta ({quantity_delta:+g})")
        rows.append({
            "Picker relacionado":picker_name or "Sin asignar",
            "Incidencias picker FNR":picker_fnr,
            "Incidencias picker MC":picker_mc,
            "Incidencias picker total":picker_total,
            "Historial artículo FNR":fnr_count,
            "Historial artículo MC":mc_count,
            "Señal picker >5":bool(picker_signal),
            "Señal artículo FNR ≥5":bool(fnr_item_signal),
            "Señal artículo MC ≥5":bool(mc_item_signal),
            "Coincidencia picker+artículo":bool(coincidence),
            "Historial mismo picker+artículo FNR":pair_fnr,
            "Historial mismo picker+artículo MC":pair_mc,
            "Historial mismo picker+artículo":pair_total,
            "Reincidencia picker+artículo":bool(pair_signal),
            "Diferencia cantidad actual":bool(quantity_mismatch),
            "Delta cantidad":quantity_delta,
            "Señales":" · ".join(signal_parts) if signal_parts else "Sin señal",
        })
    detail=pd.concat([lines.reset_index(drop=True),pd.DataFrame(rows)],axis=1)

    summary_rows=[]
    for order_id,group in detail.groupby("Número de pedido",sort=False,dropna=False):
        pickers=sorted({
            str(v).strip() for v in group["Picker relacionado"]
            if str(v).strip() and str(v).strip()!="Sin asignar"
        })
        picker_signal=bool(group["Señal picker >5"].any())
        picker_extra=bool((pd.to_numeric(group["Incidencias picker total"],errors="coerce").fillna(0)>PICKER_EXTRA_THRESHOLD).any())
        fnr_item_signal=bool(group["Señal artículo FNR ≥5"].any())
        mc_item_signal=bool(group["Señal artículo MC ≥5"].any())
        coincidence=bool(group["Coincidencia picker+artículo"].any())

        risky_items=set()
        for _,g_row in group.iterrows():
            if bool(g_row.get("Señal artículo FNR ≥5")) or bool(g_row.get("Señal artículo MC ≥5")):
                risky_items.add(_audit_item_key(g_row.get("SKU",""),g_row.get("Artículo","")) or str(g_row.get("Artículo","")))
        risky_item_count=len(risky_items)
        combo_lines=int(group["Coincidencia picker+artículo"].fillna(False).astype(bool).sum())
        pair_lines=int(group["Reincidencia picker+artículo"].fillna(False).astype(bool).sum())
        pair_max=float(pd.to_numeric(group["Historial mismo picker+artículo"],errors="coerce").fillna(0).max())
        quantity_mismatch_lines=int(group["Diferencia cantidad actual"].fillna(False).astype(bool).sum())
        max_picker_inc=float(pd.to_numeric(group["Incidencias picker total"],errors="coerce").fillna(0).max())

        score=0
        factors=[]
        if picker_signal:
            score+=2; factors.append("Picker >5 incidencias (+2)")
        if picker_extra:
            score+=1; factors.append("Picker >10 incidencias (+1)")
        if fnr_item_signal:
            score+=2; factors.append("Artículo ≥5 FNR (+2)")
        if mc_item_signal:
            score+=2; factors.append("Artículo ≥5 MC (+2)")
        if coincidence:
            score+=3; factors.append("Coincidencia picker + artículo (+3)")
        if risky_item_count>=2:
            score+=2; factors.append("2+ artículos de riesgo (+2)")
        if fnr_item_signal and mc_item_signal:
            score+=1; factors.append("Señales FNR y MC en el pedido (+1)")

        learned_total=0
        if picker_signal and learned_picker:
            score+=learned_picker; learned_total+=learned_picker
            factors.append(f"Ajuste aprendido picker (+{learned_picker})")
        if (fnr_item_signal or mc_item_signal) and learned_item:
            score+=learned_item; learned_total+=learned_item
            factors.append(f"Ajuste aprendido artículo (+{learned_item})")
        if coincidence and learned_combo:
            score+=learned_combo; learned_total+=learned_combo
            factors.append(f"Ajuste aprendido coincidencia (+{learned_combo})")

        order_key=norm(order_id)
        prior=prior_by_order.get(order_key)
        prior_bonus=False
        if prior is not None:
            current_total=float(fnr_orders.get(order_key,0.0))+float(mc_orders.get(order_key,0.0))
            saved_fnr=pd.to_numeric(prior.get("fnr_al_guardar",0),errors="coerce")
            saved_mc=pd.to_numeric(prior.get("mc_al_guardar",0),errors="coerce")
            saved_total=(0.0 if pd.isna(saved_fnr) else float(saved_fnr))+(0.0 if pd.isna(saved_mc) else float(saved_mc))
            if current_total>saved_total:
                score+=3
                prior_bonus=True
                factors.append("Auditado antes y luego sumó FNR/MC (+3)")

        # Puntaje de alarma: exige señales cruzadas más fuertes que el foco normal.
        alarm_score=int(score)
        alarm_factors=[]
        if pair_max>=2:
            alarm_score+=5
            alarm_factors.append(f"Mismo picker + artículo reincidente {pair_max:g} veces (+5)")
        if pair_max>=5:
            alarm_score+=2
            alarm_factors.append("Reincidencia picker + artículo ≥5 (+2)")
        if quantity_mismatch_lines>0 and (picker_signal or fnr_item_signal or mc_item_signal):
            alarm_score+=6
            alarm_factors.append(f"Diferencia de cantidad + historial de riesgo en {quantity_mismatch_lines} renglón(es) (+6)")
        if combo_lines>=2:
            alarm_score+=3
            alarm_factors.append("2+ coincidencias picker riesgoso + artículo riesgoso (+3)")
        if risky_item_count>=3:
            alarm_score+=2
            alarm_factors.append("3+ artículos históricos de riesgo (+2)")
        if max_picker_inc>=15:
            alarm_score+=2
            alarm_factors.append("Picker con 15+ incidencias (+2)")

        critical_candidate=bool(
            (pair_max>=2 and (picker_signal or fnr_item_signal or mc_item_signal))
            or (quantity_mismatch_lines>0 and (picker_signal or fnr_item_signal or mc_item_signal))
            or (coincidence and (risky_item_count>=2 or combo_lines>=2))
            or (coincidence and max_picker_inc>=10 and (fnr_item_signal or mc_item_signal))
        )
        independent_signals=sum([
            bool(picker_signal),
            bool(fnr_item_signal or mc_item_signal),
            bool(pair_max>=2),
            bool(quantity_mismatch_lines>0),
            bool(risky_item_count>=2),
        ])
        if critical_candidate and independent_signals>=2:
            alarm_level="🚨 Muy alarmante"
        elif coincidence or pair_max>=1 or quantity_mismatch_lines>0:
            alarm_level="⚠️ Alerta"
        else:
            alarm_level="Sin alarma crítica"

        if score>=7:
            priority="Foco alto"
        elif score>=4:
            priority="Revisar"
        else:
            priority="Sin foco"

        # Muestra de control determinística: cerca de 5% de pedidos sin foco.
        control_sample=False
        if priority=="Sin foco":
            digest=int(hashlib.sha256(str(order_id).encode("utf-8")).hexdigest()[:8],16)
            control_sample=(digest%20==0)
            if control_sample:
                priority="Muestra control"
                factors.append("Muestra aleatoria de control (~5%)")

        summary_rows.append({
            "Número de pedido":str(order_id),
            "Slot":str(group["Slot"].iloc[0]),
            "Estatus":", ".join(sorted({str(v).strip() for v in group["Estatus"] if str(v).strip()})) if "Estatus" in group else "",
            "Prioridad":priority,
            "Puntaje foco":int(score),
            "Picker >5":picker_signal,
            "Artículo FNR ≥5":fnr_item_signal,
            "Artículo MC ≥5":mc_item_signal,
            "Coincidencia picker+artículo":coincidence,
            "Artículos de riesgo":int(risky_item_count),
            "Auditoría previa con incidencia nueva":prior_bonus,
            "Ajuste aprendido":int(learned_total),
            "Alarma":alarm_level,
            "Puntaje alarma":int(alarm_score),
            "Candidato crítico":bool(critical_candidate and independent_signals>=2),
            "Coincidencias críticas":int(combo_lines),
            "Reincidencias picker+artículo":int(pair_lines),
            "Máx. historial picker+artículo":float(pair_max),
            "Diferencias cantidad":int(quantity_mismatch_lines),
            "Máx. incidencias picker":float(max_picker_inc),
            "Factores alarma":" · ".join(alarm_factors) if alarm_factors else "Sin factor crítico adicional",
            "Pickers asignados":", ".join(pickers) if pickers else "Sin asignar",
            "Renglones":int(len(group)),
            "Renglones sin picker":int(group["Picker relacionado"].eq("Sin asignar").sum()),
            "Factores":" · ".join(factors) if factors else "Sin señales de foco",
        })
    order_summary=pd.DataFrame(summary_rows)
    if not order_summary.empty:
        def _slot_hour(value):
            match=re.search(r"(?<!\\d)(\\d{1,2}):(\\d{2})(?!\\d)",str(value or ""))
            if not match: return "Sin hora"
            hour=int(match.group(1))
            return f"{hour:02d}:00–{hour:02d}:59"

        order_summary["Hora auditoría"]=order_summary["Slot"].map(_slot_hour)
        order_summary["Ranking crítico hora"]=pd.Series(pd.NA,index=order_summary.index,dtype="Int64")
        critical=order_summary[order_summary["Candidato crítico"]].copy()
        if not critical.empty:
            critical=critical.sort_values(
                ["Hora auditoría","Puntaje alarma","Máx. historial picker+artículo","Diferencias cantidad","Puntaje foco","Número de pedido"],
                ascending=[True,False,False,False,False,True],
            )
            critical["_hour_rank"]=critical.groupby("Hora auditoría",sort=False).cumcount()+1
            rank_map=dict(zip(critical["Número de pedido"].astype(str),critical["_hour_rank"].astype(int)))
            order_summary["Ranking crítico hora"]=order_summary["Número de pedido"].astype(str).map(rank_map).astype("Int64")

        level_order={"Foco alto":3,"Revisar":2,"Muestra control":1,"Sin foco":0}
        order_summary["_critical_sort"]=order_summary["Candidato crítico"].astype(int)
        order_summary["_priority_sort"]=order_summary["Prioridad"].map(level_order).fillna(0)
        order_summary=order_summary.sort_values(
            ["_critical_sort","Hora auditoría","Ranking crítico hora","Puntaje alarma","_priority_sort","Puntaje foco"],
            ascending=[False,True,True,False,False,False],
            na_position="last",
        ).drop(columns=["_critical_sort","_priority_sort"]).reset_index(drop=True)
    return detail,order_summary

def audit_coverage_filter(risks,mode):
    """Separa responsabilidades por el slot del Excel, no por hora de picking."""
    if mode=="early":
        return risks[pd.to_numeric(risks["_slot_max"],errors="coerce")<600].copy()
    if mode=="day":
        return risks[pd.to_numeric(risks["_slot_min"],errors="coerce")>=600].copy()
    return risks.copy()

def audit_hourly_quota_table(coverage,percentage=20):
    """Cupo por hora de slot: porcentaje del TOTAL de pedidos, redondeado arriba."""
    percentage=int(percentage)
    if not 1<=percentage<=100:
        raise ValueError("El porcentaje de auditoría debe estar entre 1 y 100.")
    columns=["Hora auditoría","Pedidos detectados","Porcentaje","Cupo máximo"]
    if coverage is None or coverage.empty:
        return pd.DataFrame(columns=columns)
    table=(
        coverage.groupby("Hora auditoría",dropna=False)["Número de pedido"]
        .nunique().reset_index(name="Pedidos detectados")
    )
    table["Porcentaje"]=percentage
    table["Cupo máximo"]=(
        (table["Pedidos detectados"].astype(int)*percentage+99)//100
    ).astype(int)
    return table.sort_values("Hora auditoría").reset_index(drop=True)

def audit_hourly_critical_queue(available,percentage=20,population=None):
    """Selecciona solo pedidos críticos hasta el % calculado por hora de slot.

    El denominador es toda la cobertura, incluso pedidos ya auditados: al
    ocultarlos no se amplía el cupo con casos menos urgentes. El orden conserva
    las señales FNR/MC y usa el turno nocturno para desempatar.
    """
    full_coverage=population if population is not None else available
    quotas=audit_hourly_quota_table(full_coverage,percentage)
    critical=available[available["Candidato crítico"].fillna(False).astype(bool)].copy()
    critical=critical.sort_values(
        ["Hora auditoría","Puntaje selección","Puntaje alarma",
         "Máx. historial picker+artículo","Diferencias cantidad","Número de pedido"],
        ascending=[True,False,False,False,False,True],
    )
    critical["Ranking crítico hora"]=critical.groupby("Hora auditoría",sort=False).cumcount()+1
    quota_map=quotas.set_index("Hora auditoría")["Cupo máximo"].to_dict()
    critical["Cupo máximo hora"]=critical["Hora auditoría"].map(quota_map).fillna(0).astype(int)
    return critical[critical["Ranking crítico hora"]<=critical["Cupo máximo hora"]].copy()


@st.cache_data(show_spinner="Preparando paquete de auditoría…",max_entries=6,ttl=900)
def build_bulk_audit_pdf(audit_detail,risk_summary,selected_orders,source_name=""):
    """Genera un solo PDF imprimible con una sección por pedido seleccionado."""
    selected_orders=[str(x) for x in selected_orders if str(x).strip()]
    if not selected_orders:
        return b""

    detail=audit_detail.copy()
    summary=risk_summary.copy()
    detail["Número de pedido"]=detail["Número de pedido"].astype(str)
    summary["Número de pedido"]=summary["Número de pedido"].astype(str)
    summary=summary[summary["Número de pedido"].isin(selected_orders)].copy()
    order_rank={order:i for i,order in enumerate(selected_orders)}
    summary["_packet_order"]=summary["Número de pedido"].map(order_rank).fillna(999999)
    summary=summary.sort_values(["_packet_order","Puntaje foco"],ascending=[True,False])

    out=io.BytesIO()
    doc=SimpleDocTemplate(
        out,
        pagesize=landscape(A4),
        leftMargin=9*mm,rightMargin=9*mm,topMargin=8*mm,bottomMargin=8*mm,
        title="Paquete de auditoría de pedidos",
        author="Control FNR & Mala Calidad - Coyoacán",
    )
    styles=getSampleStyleSheet()
    title_style=ParagraphStyle(
        "AuditTitle",parent=styles["Heading1"],fontName="Helvetica-Bold",
        fontSize=15,leading=17,spaceAfter=4,textColor=colors.HexColor("#272936"),
    )
    small=ParagraphStyle(
        "AuditSmall",parent=styles["BodyText"],fontName="Helvetica",
        fontSize=7.6,leading=9.2,spaceAfter=0,
    )
    small_bold=ParagraphStyle(
        "AuditSmallBold",parent=small,fontName="Helvetica-Bold",
    )
    tiny=ParagraphStyle(
        "AuditTiny",parent=styles["BodyText"],fontName="Helvetica",
        fontSize=6.8,leading=8.1,spaceAfter=0,
    )
    center=ParagraphStyle(
        "AuditCenter",parent=small,alignment=TA_CENTER,fontName="Helvetica-Bold",
    )

    story=[]
    total_orders=len(summary)
    generated=datetime.now(ZoneInfo("America/Mexico_City")).strftime("%Y-%m-%d %H:%M")

    for pos,(_,order_row) in enumerate(summary.iterrows(),start=1):
        order=str(order_row.get("Número de pedido",""))
        lines=detail[detail["Número de pedido"].eq(order)].copy()
        slot=str(order_row.get("Slot","") or "")
        priority=str(order_row.get("Alarma",order_row.get("Prioridad","")) or "")
        score=int(pd.to_numeric(order_row.get("Puntaje alarma",order_row.get("Puntaje foco",0)),errors="coerce") or 0)
        pickers=str(order_row.get("Pickers asignados","Sin asignar") or "Sin asignar")
        factors=str(order_row.get("Factores","Sin señales de foco") or "Sin señales de foco")

        story.append(Paragraph(f"AUDITORÍA DE PEDIDO - {html.escape(order)}",title_style))
        story.append(Paragraph(
            f"Paquete {pos} de {total_orders} | Generado: {generated} | Archivo: {html.escape(str(source_name or 'Sin nombre'))}",
            tiny,
        ))
        story.append(Spacer(1,2*mm))

        header_data=[
            [Paragraph("<b>Pedido</b>",small),Paragraph(html.escape(order),small_bold),
             Paragraph("<b>Slot</b>",small),Paragraph(html.escape(slot),small_bold),
             Paragraph("<b>Alarma</b>",small),Paragraph(html.escape(priority),small_bold),
             Paragraph("<b>Puntaje alarma</b>",small),Paragraph(str(score),center)],
            [Paragraph("<b>Picker(s)</b>",small),Paragraph(html.escape(pickers),small),
             Paragraph("<b>Revisión completa</b>",small),Paragraph("[ ] Sí",center),
             Paragraph("<b>Resultado</b>",small),Paragraph("[ ] Correcto   [ ] Diferencia",small),
             Paragraph("<b>Auditor</b>",small),Paragraph("________________",small)],
        ]
        header=Table(header_data,colWidths=[18*mm,42*mm,14*mm,28*mm,18*mm,32*mm,15*mm,28*mm])
        header.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.45,colors.HexColor("#b8bcc4")),
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#f3f4f6")),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LEFTPADDING",(0,0),(-1,-1),3),
            ("RIGHTPADDING",(0,0),(-1,-1),3),
            ("TOPPADDING",(0,0),(-1,-1),3),
            ("BOTTOMPADDING",(0,0),(-1,-1),3),
        ]))
        story.append(header)
        story.append(Spacer(1,1.8*mm))
        story.append(Paragraph(
            f"<b>Cobertura:</b> {html.escape(str(order_row.get('Cobertura equipo','Sin determinar')))}"
            f"  |  <b>Turno del picker:</b> {html.escape(str(order_row.get('Turno picker','Sin identificar')))}",
            tiny,
        ))
        story.append(Spacer(1,1.8*mm))
        story.append(Paragraph(f"<b>Factores de foco:</b> {html.escape(factors)}",tiny))
        story.append(Spacer(1,1.8*mm))

        table_rows=[[
            Paragraph("<b>SKU</b>",tiny),
            Paragraph("<b>Artículo</b>",tiny),
            Paragraph("<b>Picker</b>",tiny),
            Paragraph("<b>Pedida</b>",tiny),
            Paragraph("<b>Pickeada</b>",tiny),
            Paragraph("<b>Diferencia</b>",tiny),
            Paragraph("<b>Corrección / observación</b>",tiny),
        ]]
        for _,line in lines.iterrows():
            table_rows.append([
                Paragraph(html.escape(str(line.get("SKU",""))),tiny),
                Paragraph(html.escape(str(line.get("Artículo",""))),tiny),
                Paragraph(html.escape(str(line.get("Picker relacionado",line.get("Picker","")))),tiny),
                Paragraph(html.escape(str(line.get("Cantidad pedida",""))),center),
                Paragraph(html.escape(str(line.get("Cantidad pickeada",""))),center),
                Paragraph("[ ]",center),
                Paragraph("________________________________",tiny),
            ])

        item_table=Table(
            table_rows,
            colWidths=[28*mm,73*mm,45*mm,18*mm,20*mm,20*mm,62*mm],
            repeatRows=1,
            splitByRow=1,
        )
        item_table.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#c7c9ce")),
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#e9ecef")),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),2.5),
            ("RIGHTPADDING",(0,0),(-1,-1),2.5),
            ("TOPPADDING",(0,0),(-1,-1),2.4),
            ("BOTTOMPADDING",(0,0),(-1,-1),2.4),
        ]))
        story.append(item_table)
        story.append(Spacer(1,2*mm))
        story.append(Paragraph(
            "<b>Observaciones generales:</b> ________________________________________________________________________________________________",
            small,
        ))
        if pos<total_orders:
            story.append(PageBreak())

    def _page_number(canvas,doc_obj):
        canvas.saveState()
        canvas.setFont("Helvetica",6.5)
        canvas.setFillColor(colors.HexColor("#6f7480"))
        canvas.drawRightString(landscape(A4)[0]-9*mm,4.5*mm,f"Página {doc_obj.page}")
        canvas.restoreState()

    doc.build(story,onFirstPage=_page_number,onLaterPages=_page_number)
    return out.getvalue()

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

def _h(value):
    """Texto seguro para interpolar dentro de HTML renderizado por Streamlit."""
    return html.escape(str(value if value is not None else ""),quote=True)


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
                if not smtp.has_extn("starttls"):
                    return False,"Correo no enviado: el servidor SMTP no ofrece STARTTLS; por seguridad no se enviaron credenciales sin cifrado."
                smtp.starttls(context=context); smtp.ehlo()
                smtp.login(user,password); smtp.send_message(msg)
        return True,"Copia enviada correctamente a: "+", ".join(recipients)
    except Exception as exc:
        return False,f"No fue posible enviar la copia. Revisa el servidor y los datos SMTP. Detalle: {exc}"

@st.cache_data(show_spinner="Preparando Excel…",max_entries=4,ttl=1800)
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
st.markdown("""
<div style="display:flex;align-items:center;gap:14px;margin:.5rem 0 .15rem;">
  <svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 80 80" role="img" aria-label="Dibujo de un aguacate">
    <path d="M39 9c2-5 8-7 13-5-1 6-6 10-12 10" fill="#72a848"/>
    <path d="M42 14c-5-7-14-8-20-3-6 5-8 14-11 23-4 11-6 19-2 28 5 12 18 16 31 14 14-2 26-10 28-23 2-10-4-22-10-31-4-6-9-10-15-8z" fill="#3f7d3b"/>
    <path d="M41 19c-5-5-12-5-17-1-5 5-7 13-10 21-3 9-5 16-1 23 4 9 15 12 26 10 12-2 21-8 23-19 2-8-3-19-8-27-4-5-8-8-13-7z" fill="#a8d66d"/>
    <path d="M40 32c-8 0-14 7-14 15 0 9 7 15 15 15s15-6 15-15c0-8-7-15-16-15z" fill="#87522f"/>
    <path d="M40 37c-5 0-9 5-9 10s4 10 10 10 10-5 10-10-5-10-11-10z" fill="#a9683b"/>
    <path d="M19 22c-1-7 3-13 10-15 3 7 1 13-5 17" fill="#78a944"/>
  </svg>
  <h1 style="margin:0;color:#272936;font-size:2.45rem;line-height:1.15;font-weight:750;">Control FNR &amp; Mala Calidad</h1>
</div>
""", unsafe_allow_html=True)
st.caption("Control operativo de pickers, calidad, seguimiento y procesos")

# Estado persistente: se carga antes de construir los widgets.
store=load_store()
if "_startup_migrations_done" not in st.session_state:
    _migration_result=migrate_local_assets_to_cloud(store)
    _history_migrated=migrate_local_upload_history_to_cloud()
    st.session_state["_startup_migrations_done"]=(_migration_result[0],_migration_result[1],_history_migrated)
cloud_migrated,cloud_missing,cloud_history_migrated=st.session_state["_startup_migrations_done"]
store.setdefault("excluded_orders", [])
store.setdefault("feedback_rows", [])
store.setdefault("master_overrides", {})
store.setdefault("master_excluded", [])
store.setdefault("procesos", [])
store.setdefault("recursos_formatos", [])
store.setdefault("seguimientos_documentos", [])
store.setdefault("supervisores", [])
store.setdefault("upload_meta", {})
store.setdefault("attendance_summary", None)
store.setdefault("attendance_meta", {})
store.setdefault("attendance_links", {})
store.setdefault("attendance_shift_config",{"Mañana":[6,7,8,9,10],"Intermedio":[11],"Tarde":[13,14],"Nocturno":[22]})
store.setdefault("order_audits", [])
store.setdefault("monthly_history", {})
store.setdefault("powerbi_krs_history", [])
if normalize_saved_followup_categories(store):
    save_store(store)

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
    with st.expander("📁 Actualizar archivos operativos",expanded=False):
        st.caption("Carga los Excel que alimentan el análisis. Se conserva cada versión en el historial.")
        ub_upload=st.file_uploader("① Base de Pickers / Líneas",type=["xlsx","xls"],key="base_picker")
        uf_upload=st.file_uploader("② Detalle FNR",type=["xlsx","xls"],key="detalle_fnr")
        um_upload=st.file_uploader("③ Detalle Mala Calidad",type=["xlsx","xls"],key="detalle_mc")
        up_upload=st.file_uploader("④ Plantilla consolidada",type=["xlsx","xls"],key="plantilla_personal")

    # Cada archivo nuevo reemplaza automáticamente al guardado. Si solo se recarga
    # la página, la app recupera la última versión guardada sin pedir volver a subirla.
    ub=persist_uploaded_once(ub_upload,"base_picker")
    uf=persist_uploaded_once(uf_upload,"detalle_fnr")
    um=persist_uploaded_once(um_upload,"detalle_mc")
    up=persist_uploaded_once(up_upload,"plantilla_personal")
    _upload_meta_changed=False
    for _key,_upload,_label in [("base_picker",ub_upload,"Pickers/Líneas"),("detalle_fnr",uf_upload,"FNR"),("detalle_mc",um_upload,"Mala Calidad"),("plantilla_personal",up_upload,"Plantilla consolidada")]:
        if _upload is None: continue
        _bytes=_upload.getvalue()
        _fingerprint=hashlib.sha256(_bytes).hexdigest()
        _old_meta=(store.get("upload_meta",{}) or {}).get(_key,{}) or {}
        if _old_meta.get("huella")!=_fingerprint:
            store.setdefault("upload_meta",{})[_key]={
                "fecha":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "archivo":str(getattr(_upload,"name",_label)),
                "huella":_fingerprint,
            }
            _upload_meta_changed=True
    if _upload_meta_changed:
        save_store(store)
    if cloud_enabled():
        _backup_at=str(store.get("cloud_backup_at","")).strip()
        if _backup_at:
            st.caption(f"Último respaldo confirmado: {_backup_at}")
        else:
            st.caption("El respaldo se confirmará al guardar el primer cambio.")

    status=persisted_status()
    labels={"base_picker":"Pickers/Líneas","detalle_fnr":"FNR","detalle_mc":"Mala Calidad","plantilla_personal":"Master Pickers"}
    guardados=[labels[k] for k,v in status.items() if v]
    if guardados:
        st.success("Archivos guardados: " + ", ".join(guardados))
    with st.expander("🗂️ Historial de archivos Excel",expanded=False):
        _history=upload_history_rows()
        if _history:
            # Importación local defensiva: esta tabla es la línea señalada por
            # Streamlit cuando una copia de app.py quedó sin el alias global.
            import pandas as pd
            st.dataframe(pd.DataFrame(_history),use_container_width=True,hide_index=True)
            if st.button("Preparar ZIP de respaldo",key="prepare_excel_history_zip"):
                with st.spinner("Preparando el respaldo de archivos…"):
                    st.session_state["_excel_history_zip_ready"]=upload_history_zip()
            if st.session_state.get("_excel_history_zip_ready"):
                st.download_button("Descargar respaldo del historial (.zip)",st.session_state["_excel_history_zip_ready"],"Historial_Excel_FNR_MC.zip","application/zip",key="download_excel_history")
            st.caption("El ZIP se prepara solo cuando solicitas el respaldo.")
        else:
            st.info("Aún no hay versiones en el historial. Se registra una versión al cargar cada Excel.")
    st.caption("Los 3 Excel operativos se cruzan con la PLANTILLA CONSOLIDADA usando CORREO como llave principal. Si falta el correo, la app intenta una resolución conservadora por código/nombre y deja pendientes los casos ambiguos.")
    st.divider()
    periodo=st.text_input("Periodo",value=str(store.get("periodo",datetime.now().strftime("%Y-%m"))),key="periodo_persistente")
    saved_excluded="\n".join(str(x) for x in store.get("excluded_orders",[]) if str(x).strip())
    ex=st.text_area("Pedidos operativos a excluir (uno por línea)",value=saved_excluded,key="pedidos_excluidos_persistentes")
    excluded={x.strip() for x in ex.splitlines() if x.strip()}
    # Guardar preferencias solo cuando cambien; evita escrituras a nube en cada interacción.
    _new_periodo=periodo.strip() or datetime.now().strftime("%Y-%m")
    _new_excluded=sorted(excluded)
    if store.get("periodo")!=_new_periodo or list(store.get("excluded_orders",[]))!=_new_excluded:
        store["periodo"]=_new_periodo
        store["excluded_orders"]=_new_excluded
        save_store(store)

# Indicadores compactos de Power BI: siempre visibles en portada, fuera de las pestañas.
with st.container(border=True):
    st.markdown("### 📊 Indicadores Power BI · Coyoacán")
    st.caption("Valores acumulados del periodo, actualizados al cargar el Excel de Power BI. Se muestran aparte del cálculo operativo basado en los Excel.")
    _krs_history=store.get("powerbi_krs_history",[]) or []
    _home_latest=_krs_history[-1] if _krs_history else None

    # El cargador queda plegado para dejar los resultados como foco de la portada.
    with st.expander("⬆️ Actualizar indicadores desde Power BI",expanded=False):
        st.caption("Sube el Excel descargado desde Power BI. Se conserva un registro por fecha para revisar la evolución diaria.")
        with st.form("powerbi_krs_upload_home_form",clear_on_submit=True):
            _krs_upload=st.file_uploader("Excel de Power BI",type=["xlsx","xls"],key="powerbi_krs_excel_upload_home")
            _krs_submit=st.form_submit_button("Guardar valores",type="primary",use_container_width=True)
        if _krs_submit:
            if _krs_upload is None:
                st.warning("Selecciona primero el Excel descargado desde Power BI.")
            else:
                try:
                    _snapshot=parse_powerbi_krs_excel(_krs_upload.getvalue(),_krs_upload.name)
                    if _snapshot.get("tienda") not in {"Coyoacán","Coyoacan"}:
                        st.error(f"El archivo indica tienda: {_snapshot.get('tienda','No indicada')}. Selecciona Coyoacán en Power BI para guardar estos indicadores.")
                    else:
                        _krs_history=store.setdefault("powerbi_krs_history",[])
                        _same_day=next((n for n,item in enumerate(_krs_history) if item.get("fecha")==_snapshot["fecha"] and item.get("periodo")==_snapshot["periodo"] and item.get("tienda")==_snapshot["tienda"]),None)
                        if _same_day is not None and _krs_history[_same_day].get("huella")==_snapshot["huella"]:
                            st.info("Este Excel ya está guardado para hoy; no agregué un duplicado.")
                        else:
                            if _same_day is None: _krs_history.append(_snapshot)
                            else: _krs_history[_same_day]=_snapshot
                            _krs_history.sort(key=lambda item:(str(item.get("fecha","")),str(item.get("cargado",""))))
                            store["powerbi_krs_history"]=_krs_history
                            save_store(store)
                            st.success(f"Valores guardados: {_snapshot['fecha']} · {_snapshot['periodo']} · Coyoacán.")
                        _home_latest=store.get("powerbi_krs_history",[])[-1]
                except Exception as _krs_error:
                    st.error(f"No pude leer ese Excel: {_krs_error}")

    if _home_latest:
        st.caption(f"Última carga: {_home_latest.get('fecha','')} · Periodo: {_home_latest.get('periodo','')} · Verde = en meta · Rojo = fuera de meta")
        _home_card_cols=st.columns(3,gap="small")
        for _col,_name in zip(_home_card_cols,["On Time","FNR","Mala Calidad"]):
            _item=(_home_latest.get("indicadores",{}) or {}).get(_name,{})
            _value=_item.get("valor"); _target=_item.get("meta")
            _outside=False
            if _value is not None and _target is not None:
                _outside=float(_value)<float(_target) if _name=="On Time" else float(_value)>float(_target)
            _tone="#fff1f2" if _outside else ("#f0fdf4" if _value is not None and _target is not None else "#f8fafc")
            _edge="#fca5a5" if _outside else ("#86efac" if _value is not None and _target is not None else "#cbd5e1")
            _state="FUERA DE META" if _outside else ("EN META" if _value is not None and _target is not None else "SIN META")
            _value_text=f"{float(_value):.2f}%" if _value is not None else "N/D"
            _target_text=f"Meta: {float(_target):.2f}%" if _target is not None else "Meta no incluida"
            _label_color="#b4232c" if _outside else ("#15803d" if _value is not None and _target is not None else "#6b7280")
            _col.markdown(f"<div style='background:{_tone};border:1px solid {_edge};border-radius:12px;padding:9px 12px;min-height:78px'><div style='font-size:.8rem;color:#4b5563'>{_name}</div><div style='font-size:1.35rem;font-weight:750;color:#272936'>{_value_text}</div><div style='font-size:.72rem;color:{_label_color};font-weight:700'>{_state} · {_target_text}</div></div>",unsafe_allow_html=True)
        with st.expander(f"Ver historial ({len(_krs_history)} cargas)",expanded=False):
            _history_rows=[]
            for _snap in _krs_history:
                _row={"Fecha":_snap.get("fecha",""),"Periodo":_snap.get("periodo",""),"Tienda":_snap.get("tienda","")}
                for _name in ["On Time","FNR","Mala Calidad"]:
                    _v=(_snap.get("indicadores",{}) or {}).get(_name,{}).get("valor")
                    _row[_name]=float(_v) if _v is not None else None
                _history_rows.append(_row)
            _history_df=pd.DataFrame(_history_rows)
            _history_df["Fecha"]=pd.to_datetime(_history_df["Fecha"],errors="coerce")
            _history_df=_history_df.dropna(subset=["Fecha"]).sort_values("Fecha")
            if len(_history_df)>1: st.line_chart(_history_df.set_index("Fecha")[["On Time","FNR","Mala Calidad"]],use_container_width=True)
            _show_history=_history_df.sort_values("Fecha",ascending=False).copy()
            _show_history["Fecha"]=_show_history["Fecha"].dt.strftime("%Y-%m-%d")
            st.dataframe(_show_history,use_container_width=True,hide_index=True)
    else:
        st.info("Aún no hay datos. Abre “Actualizar indicadores desde Power BI” para cargar el Excel.");

if not (ub and uf and um and up):
    st.info("Carga los 3 Excel operativos y la plantilla consolidada para comenzar. Para que el cruce sea por correo, los registros que deban asociarse deben traer CORREO / CODIGO + CORREO.")
    st.markdown("**Fuentes:** ① Pickers/Líneas · ② FNR · ③ Mala Calidad · ④ Plantilla consolidada (correo → nombre asociado, turno, supervisor y área).")
    st.stop()

try:
    personnel_config_json=json.dumps({
        "master_overrides":store.get("master_overrides",{}),
        "master_excluded":store.get("master_excluded",[]),
    },ensure_ascii=False,sort_keys=True)
    base,fnr,mc,roster=prepare_source_frames(
        ub.getvalue(),uf.getvalue(),um.getvalue(),up.getvalue(),personnel_config_json
    )
except Exception as e:
    st.error(f"Error en los archivos cargados: {e}"); st.stop()

if excluded:
    fnr=fnr[~fnr.ORDER_NUMBER.isin(excluded)].copy()
    mc=mc[~mc.ORDER_NUMBER.isin(excluded)].copy()
base=_dedupe_columns(base)

base,fnr,mc=exclude_registered_supervisors(base,fnr,mc,store)
s=summary(base,fnr,mc)
# Estado de cruce de TODOS los archivos: Pickers/Líneas + FNR + MC.
# La llave es CORREO_KEY; cualquier registro sin coincidencia con la plantilla
# se concentra en la pestaña 🧰 Herramientas para asignación manual.
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
# Registrar automáticamente pickers vistos en la operación solo si aparecen por primera vez.
_picker_registry_changed=False
for _idx,_row in base.iterrows():
    _p=str(_row.get("PICKER","")).strip()
    _src=str(_row.get("_SOURCE_PICKER","")).strip()
    if not _p: continue
    _before=set((store.get("pickers",{}) or {}).keys())
    picker_record(store,_p,aliases=[_src] if _src else [])
    if set((store.get("pickers",{}) or {}).keys())!=_before:
        _picker_registry_changed=True
if _picker_registry_changed:
    save_store(store)

if roster is not None:
    match_series=base.get("_MASTER_MATCH",pd.Series(False,index=base.index)).fillna(False).astype(bool)
    excluded_series=base.get("_EXCLUDED_PERSONNEL",pd.Series(False,index=base.index)).fillna(False).astype(bool)
    matched=int((match_series & ~excluded_series).sum())
    total=len(base); unmatched=total-matched-int(excluded_series.sum())
    matched_turn=int((match_series & ~excluded_series & base["TURNO"].astype(str).str.strip().ne("") & base["TURNO"].astype(str).str.strip().ne("No especificado") & base["TURNO"].astype(str).str.strip().ne("__EXCLUIDO__")).sum())
    with st.sidebar:
        st.success(f"Base Pickers/Líneas: {matched} identificados de {total} registros")
        st.caption(f"Con turno asignado: {matched_turn}/{total} · Excluidos: {int(excluded_series.sum())}")
        if unmatched:
            st.caption(f"Pendientes de asociar en esta base: {unmatched} registros.")
        email_match=int(base.get("_MASTER_MATCH",pd.Series(False,index=base.index)).fillna(False).astype(bool).sum())
        st.caption(f"🔑 Cruce de identidad: CORREO primero; código/nombre solo como respaldo conservador · {email_match} registros identificados")

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
        r=pd.DataFrame(columns=["PICKER","TURNO","SUPERVISOR","AREA_BASE","CORREO","CORREO_KEY","PERSONAL_TIPO","_CONTEXT_KEY","CATEGORIA"])
    else:
        r=roster.copy()
    r["PICKER"]=r.get("PICKER",pd.Series("",index=r.index)).fillna("").astype(str).str.strip()
    r["TURNO"]=r.get("TURNO_MAESTRO",pd.Series("",index=r.index)).fillna("").astype(str).str.strip()
    r["SUPERVISOR"]=r.get("SUPERVISOR",pd.Series("",index=r.index)).fillna("").astype(str).str.strip()
    r["AREA_BASE"]=r.get("AREA_MAESTRO",pd.Series("",index=r.index)).fillna("").astype(str).str.strip()
    r["PERSONAL_TIPO"]=r.get("PERSONAL_TIPO",pd.Series("",index=r.index)).fillna("").astype(str).str.strip()
    r["CORREO"]=r.get("CORREO",pd.Series("",index=r.index)).fillna("").astype(str).str.strip()
    r["CORREO_KEY"]=r["CORREO"].map(email_key)
    r=r[r["CORREO_KEY"].ne("")].copy()
    r.loc[r["TURNO"].eq(""),"TURNO"]="No especificado"
    r.loc[r["SUPERVISOR"].eq(""),"SUPERVISOR"]="No asignado"
    r.loc[r["AREA_BASE"].eq(""),"AREA_BASE"]="No especificada"
    _layout_mask=(r["PERSONAL_TIPO"].map(lambda v:"layout" in re.sub(r"[^a-z0-9]", "", norm(v)))
                  | r["AREA_BASE"].map(lambda v:"layout" in re.sub(r"[^a-z0-9]", "", norm(v))))
    r.loc[_layout_mask,"PERSONAL_TIPO"]="Layout"
    r.loc[r["PERSONAL_TIPO"].eq(""),"PERSONAL_TIPO"]="Picker"
    _leader_role_mask=r["PERSONAL_TIPO"].map(norm).isin({"supervisor","subgerente","jefe"})
    r["CATEGORIA"]=np.where(_leader_role_mask,"Supervisor","Picker")
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
    """Filtra registros por identidad estable (correo; nombre solo para no registrados)."""
    if df is None or df.empty or summary_df is None or summary_df.empty:
        return df.iloc[0:0].copy() if isinstance(df,pd.DataFrame) else pd.DataFrame()
    allowed_df=summary_df.copy()
    source=df.copy()
    allowed_df["_PERSON_KEY"]=allowed_df.apply(person_context_key,axis=1)
    source["_PERSON_KEY"]=source.apply(person_context_key,axis=1)
    allowed=set(allowed_df["_PERSON_KEY"].astype(str))
    return source[source["_PERSON_KEY"].astype(str).isin(allowed)].drop(columns=["_PERSON_KEY"],errors="ignore").copy()

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
    count_text=f"{sel_p:,} pickers · {sel_u:,} sin registrar en base e incidencias" if sel_u else f"{sel_p:,} pickers"
    st.markdown(
        f"<div class='context-banner'><div class='context-title'>🎯 {_h(label)}: {_h(scope)}</div>"
        f"<div class='context-detail'>{_h(count_text)} · {pct:.1f}% del universo de comparación · Las demás pestañas utilizan este mismo contexto.</div></div>",
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

def _severity_css(value, warning, critical):
    """Color pastel; N/D queda sin color para no sugerir un nivel desconocido."""
    try:
        number=float(value)
    except (TypeError,ValueError):
        return ""
    if pd.isna(number): return ""
    if number>=critical: return "background-color:#fde2e2;color:#991b1b;font-weight:600"
    if number>=warning: return "background-color:#fff4cc;color:#854d0e;font-weight:600"
    return "background-color:#e2f3e5;color:#166534"

def _style_severity(frame, rules):
    """Colorea incidencias y presenta conteos como enteros sin alterar sus datos."""
    styled=frame.style
    for column in ["LINEAS","PEDIDOS","FNR","MC","INCIDENCIAS","PICKERS","PRODUCTOS","Faltas","Retardos","Minutos acumulados"]:
        if column in frame.columns:
            styled=styled.format("{:,.0f}",subset=[column],na_rep="—")
    for column in ["FNR_%","MC_%","FNR %","MC %","% DEL TOTAL","% / LINEAS"]:
        if column in frame.columns:
            styled=styled.format("{:.2f}%",subset=[column],na_rep="—")
    for column,(warning,critical) in rules.items():
        if column in frame.columns:
            styled=styled.apply(
                lambda values,w=warning,c=critical:[_severity_css(value,w,c) for value in values],
                subset=[column],
            )
    return styled

def _percentile_severity_rules(frame, columns, percentile=.75):
    """Calcula corte rojo del cuartil superior de la plantilla cargada."""
    rules={}
    for column in columns:
        if column not in frame.columns: continue
        values=pd.to_numeric(frame[column],errors="coerce").replace([np.inf,-np.inf],np.nan).dropna()
        positive=values[values>0]
        if positive.empty:
            rules[column]=(1,1)
            continue
        cutoff=int(np.ceil(values.quantile(percentile)))
        if cutoff<1: cutoff=int(np.ceil(positive.quantile(percentile)))
        rules[column]=(1,max(cutoff,1))
    return rules

def _severity_tone(value, warning, critical):
    css=_severity_css(value,warning,critical)
    if "fde2e2" in css: return "red"
    if "fff4cc" in css: return "amber"
    if "e2f3e5" in css: return "green"
    return "blue"

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
sp=st.session_state.get("picker_page_picker","Todos")
stn="Todos"
ssup="Todos"
sar="Todos"
base_view=select_summary_people(base,s_view)
fnr_view=select_summary_people(fnr,s_view)
mc_view=select_summary_people(mc,s_view)

attendance_summary=store.get("attendance_summary") or {}
attendance_people=attendance_summary.get("personas",[]) if isinstance(attendance_summary,dict) else []
attendance_by_name={}
for _person in attendance_people:
    _k=token_key(_person.get("Persona",""))
    if _k: attendance_by_name.setdefault(_k,[]).append(_person)
app_people_by_name={}
if isinstance(s,pd.DataFrame) and "PICKER" in s.columns:
    for _,_row in s.iterrows():
        _name=str(_row.get("PICKER","")).strip()
        _k=token_key(_name)
        if _k:
            _category=_row.get("CATEGORIA","Picker"); _supervisor=_row.get("SUPERVISOR","No especificado")
            if pd.isna(_category) or not str(_category).strip(): _category="Picker"
            if pd.isna(_supervisor) or not str(_supervisor).strip(): _supervisor="No especificado"
            _candidate={
                "PICKER":_name,"CATEGORIA":str(_category),"SUPERVISOR":str(_supervisor),
                "CORREO":str(_row.get("CORREO","")),
                "TURNO":str(_row.get("TURNO","")),
                "AREA_BASE":str(_row.get("AREA_BASE","")),
            }
            _bucket=app_people_by_name.setdefault(_k,[])
            if not any(
                x["PICKER"]==_candidate["PICKER"]
                and x["CATEGORIA"]==_candidate["CATEGORIA"]
                and email_key(x.get("CORREO",""))==email_key(_candidate.get("CORREO",""))
                for x in _bucket
            ):
                _bucket.append(_candidate)
app_people_names=sorted({z["PICKER"] for _bucket in app_people_by_name.values() for z in _bucket},key=lambda z:z.upper())

with st.expander("🔳 Generador de códigos QR",expanded=False):
    st.caption("Genera códigos CYN-PCABA-U1–U30 o un QR personalizado sin ocupar espacio en la portada.")
    with st.container(border=True):
        _qr_select_col,_qr_preview_col=st.columns([1.5,1])
        with _qr_select_col:
            _qr_codes=[f"CYN-PCABA-U{_n}" for _n in range(1,31)]
            _qr_selection=st.selectbox("Código rápido · pickers U1–U30",["Escribir enlace o texto"]+_qr_codes,key="qr_generator_selection")
            if _qr_selection=="Escribir enlace o texto":
                _qr_content=st.text_input("Enlace o texto",placeholder="Pega un enlace o escribe texto",key="qr_generator_content")
            else:
                _qr_content=_qr_selection
                st.caption(f" Código seleccionado: **{_qr_content}**")
        with _qr_preview_col:
            if _qr_content.strip():
                try:
                    _qr_code=qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M,box_size=8,border=4)
                    _qr_code.add_data(_qr_content.strip())
                    _qr_code.make(fit=True)
                    _qr_image=_qr_code.make_image(fill_color="black",back_color="white").convert("RGB")
                    _qr_buffer=io.BytesIO()
                    _qr_image.save(_qr_buffer,format="PNG")
                    _qr_bytes=_qr_buffer.getvalue()
                    st.image(_qr_bytes,caption="Vista previa",width=145)
                    _qr_filename=re.sub(r"[^A-Za-z0-9_-]+","_",_qr_content.strip()).strip("_")[:60] or "codigo_qr"
                    st.download_button("⬇️ Descargar QR",_qr_bytes,file_name=f"{_qr_filename}.png",mime="image/png",key="download_generated_qr")
                except Exception as _qr_error:
                    st.error(f"No se pudo generar el QR: {_qr_error}")
            else:
                st.caption("Selecciona CYN-PCABA-U1–U30 para generar el QR al instante.")
    
st.divider()

_tab_labels=["🏠 Bodega y turnos / áreas","👤 Pickers y supervisores","📅 Faltas y retardos","🛡️ Seguimiento","📌 Pendientes","📦 Auditoría de pedidos","🧰 Herramientas"]
a,b,i,h,j,o,x=st.tabs(_tab_labels,on_change="rerun",key="control_fnr_mc_tabs_v7")
# Agrupa las secciones en una sola pestaña y conserva sus formularios y cálculos.
e=a  # Turnos / Áreas comparte la pestaña de Bodega.
g=b  # Supervisores comparte la pestaña de Pickers.
def _tab_active(tab):
    # Older Streamlit versions return no selected state; render normally there.
    return getattr(tab,"open",None) is not False

if _tab_active(a):
    with a:
        st.subheader(f"Resumen de bodega — {periodo}")
        st.caption("El contexto elegido aquí se comparte automáticamente con todas las pestañas.")

        _meta=store.get("upload_meta",{}) or {}
        _upload_rows=[]
        for _k,_label in [("base_picker","Pickers / Líneas"),("detalle_fnr","FNR"),("detalle_mc","Mala Calidad"),("plantilla_personal","Plantilla consolidada")]:
            _m=_meta.get(_k,{}) or {}
            _upload_rows.append({"Archivo":_label,"Última carga":_m.get("fecha","Sin registro"),"Nombre":_m.get("archivo","")})
        with st.expander("📂 Ver última carga de Excel",expanded=False):
            st.dataframe(pd.DataFrame(_upload_rows),use_container_width=True,hide_index=True)

        with st.expander("📊 Histórico · comparar con mes anterior",expanded=False):
            st.caption("Sube los mismos 3 Excel operativos del periodo anterior. La plantilla actual se usa para identificar a los mismos pickers. Al guardar el histórico, ya no tendrás que volver a subirlo para comparar.")
            _monthly_history=store.get("monthly_history",{}) or {}
            _history_saved_keys=sorted(
                _monthly_history.keys(),
                key=lambda key:str((_monthly_history.get(key,{}) or {}).get("period_key",key)),
                reverse=True,
            )
            _history_options=["➕ Cargar nuevo periodo"]+_history_saved_keys
            _history_choice=st.selectbox("Comparar contra",_history_options,key="monthly_history_choice")
            _history_snapshot=None

            if _history_choice=="➕ Cargar nuevo periodo":
                _today=datetime.now()
                _prev_month=12 if _today.month==1 else _today.month-1
                _prev_year=_today.year-1 if _today.month==1 else _today.year
                _month_names=["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
                _default_hist_label=f"{_month_names[_prev_month-1]} {_prev_year}"
                _hist_label=st.text_input("Nombre del periodo",value=_default_hist_label,key="monthly_history_label")
                _hu1,_hu2,_hu3=st.columns(3)
                with _hu1:
                    _hist_base_up=st.file_uploader("① Pickers / líneas · mes anterior",type=["xlsx","xls"],key="monthly_hist_base")
                with _hu2:
                    _hist_fnr_up=st.file_uploader("② FNR · mes anterior",type=["xlsx","xls"],key="monthly_hist_fnr")
                with _hu3:
                    _hist_mc_up=st.file_uploader("③ Mala Calidad · mes anterior",type=["xlsx","xls"],key="monthly_hist_mc")

                if _hist_base_up and _hist_fnr_up and _hist_mc_up:
                    try:
                        _hist_base,_hist_fnr,_hist_mc=prepare_history_frames(
                            _hist_base_up.getvalue(),_hist_fnr_up.getvalue(),_hist_mc_up.getvalue(),
                            roster,personnel_config_json
                        )
                        _hist_base,_hist_fnr,_hist_mc=exclude_registered_supervisors(_hist_base,_hist_fnr,_hist_mc,store)
                        _hist_summary=summary(_hist_base,_hist_fnr,_hist_mc)
                        _history_snapshot=monthly_history_snapshot(
                            _hist_label,_hist_summary,
                            {
                                "pickers":_hist_base_up.name,
                                "fnr":_hist_fnr_up.name,
                                "mc":_hist_mc_up.name,
                            }
                        )
                        st.success(f"Periodo leído: {_hist_label} · {len(_hist_summary):,} personas encontradas.")
                        if st.button("💾 Guardar / actualizar este histórico",type="primary",key="save_monthly_history"):
                            _clean_label=_hist_label.strip() or _default_hist_label
                            _history_snapshot["periodo"]=_clean_label
                            _history_snapshot["period_key"]=f"{_prev_year:04d}-{_prev_month:02d}"
                            store.setdefault("monthly_history",{})[_clean_label]=_history_snapshot
                            save_store(store)
                            st.success(f"Histórico {_clean_label} guardado.")
                    except Exception as _hist_error:
                        st.error(f"No pude procesar los 3 Excel históricos: {_hist_error}")
                else:
                    st.info("Carga los 3 archivos para generar el comparativo.")
            else:
                _history_snapshot=_monthly_history.get(_history_choice)
                if _history_snapshot:
                    st.caption(f"Histórico guardado: {_history_snapshot.get('guardado','')} · Clave: {_history_snapshot.get('period_key','Sin clave cronológica')}")
                    if st.button("🗑️ Eliminar este histórico",key="delete_monthly_history"):
                        store.setdefault("monthly_history",{}).pop(_history_choice,None)
                        save_store(store)
                        st.success("Histórico eliminado.")
                        st.rerun()

            if _history_snapshot:
                _prev_tot=_history_snapshot.get("totales",{}) or {}
                _cur_lines=float(pd.to_numeric(s.get("LINEAS",0),errors="coerce").fillna(0).sum())
                _cur_fnr=float(pd.to_numeric(s.get("FNR",0),errors="coerce").fillna(0).sum())
                _cur_mc=float(pd.to_numeric(s.get("MC",0),errors="coerce").fillna(0).sum())
                _prev_lines=float(_prev_tot.get("lineas",0) or 0)
                _prev_fnr=float(_prev_tot.get("fnr",0) or 0)
                _prev_mc=float(_prev_tot.get("mc",0) or 0)
                _cur_fnr_pct=_cur_fnr/_cur_lines*100 if _cur_lines else None
                _cur_mc_pct=_cur_mc/_cur_lines*100 if _cur_lines else None
                _prev_fnr_pct=_prev_fnr/_prev_lines*100 if _prev_lines else None
                _prev_mc_pct=_prev_mc/_prev_lines*100 if _prev_lines else None
                _cur_inc=_cur_fnr+_cur_mc
                _prev_inc=_prev_fnr+_prev_mc

                st.markdown(f"#### Actual vs {_history_snapshot.get('periodo','periodo anterior')}")
                _hc1,_hc2,_hc3,_hc4=st.columns(4)
                _line_delta=_cur_lines-_prev_lines
                _line_pct=(_line_delta/_prev_lines*100) if _prev_lines else None
                _hc1.metric(
                    "Líneas",
                    f"{_cur_lines:,.0f}",
                    f"{_line_delta:+,.0f}" + (f" ({_line_pct:+.1f}%)" if _line_pct is not None else ""),
                )
                _fnr_pp=(_cur_fnr_pct-_prev_fnr_pct) if _cur_fnr_pct is not None and _prev_fnr_pct is not None else None
                _mc_pp=(_cur_mc_pct-_prev_mc_pct) if _cur_mc_pct is not None and _prev_mc_pct is not None else None
                _hc2.metric(
                    "FNR",
                    f"{_cur_fnr:,.0f}" + (f" · {_cur_fnr_pct:.2f}%" if _cur_fnr_pct is not None else ""),
                    f"{_cur_fnr-_prev_fnr:+,.0f} inc." + (f" · {_fnr_pp:+.2f} pp" if _fnr_pp is not None else ""),
                    delta_color="inverse",
                )
                _hc3.metric(
                    "MC",
                    f"{_cur_mc:,.0f}" + (f" · {_cur_mc_pct:.2f}%" if _cur_mc_pct is not None else ""),
                    f"{_cur_mc-_prev_mc:+,.0f} inc." + (f" · {_mc_pp:+.2f} pp" if _mc_pp is not None else ""),
                    delta_color="inverse",
                )
                _hc4.metric(
                    "Incidencias totales",
                    f"{_cur_inc:,.0f}",
                    f"{_cur_inc-_prev_inc:+,.0f} vs anterior",
                    delta_color="inverse",
                )

                _history_compare=monthly_picker_comparison(s,_history_snapshot.get("pickers",[]))
                _hist_filter=st.selectbox(
                    "Ver pickers",
                    ["Todos","🔴 Más incidencias","🟢 Menos incidencias","📈 Subieron líneas","📉 Bajaron líneas","🔴 Mayor tasa de incidencias","🟢 Menor tasa de incidencias"],
                    key="monthly_history_filter",
                )
                _history_view=_history_compare.copy()
                if _hist_filter=="🔴 Más incidencias":
                    _history_view=_history_view[_history_view["Δ incidencias"]>0]
                elif _hist_filter=="🟢 Menos incidencias":
                    _history_view=_history_view[_history_view["Δ incidencias"]<0]
                elif _hist_filter=="📈 Subieron líneas":
                    _history_view=_history_view[_history_view["Δ líneas"]>0]
                elif _hist_filter=="📉 Bajaron líneas":
                    _history_view=_history_view[_history_view["Δ líneas"]<0]
                elif _hist_filter=="🔴 Mayor tasa de incidencias":
                    _history_view=_history_view[_history_view["Δ Inc./1000"]>0]
                elif _hist_filter=="🟢 Menor tasa de incidencias":
                    _history_view=_history_view[_history_view["Δ Inc./1000"]<0]

                _history_cols=[
                    "Picker",
                    "Líneas anterior","Líneas actual","Δ líneas","% Δ líneas",
                    "FNR anterior","FNR actual","Δ FNR","FNR % anterior","FNR % actual","Δ FNR pp",
                    "MC anterior","MC actual","Δ MC","MC % anterior","MC % actual","Δ MC pp",
                    "Incidencias anterior","Incidencias actual","Δ incidencias",
                    "Inc./1000 anterior","Inc./1000 actual","Δ Inc./1000",
                    "Tendencia",
                ]
                st.dataframe(_history_view[_history_cols],use_container_width=True,hide_index=True)
                st.caption("Inc./1000 líneas ayuda a separar el efecto de producir más o menos volumen: una persona puede tener más incidencias en cantidad, pero una tasa menor si también aumentaron mucho sus líneas.")
                st.info("Para comparar totales de líneas/FNR/MC, procura que ambos Excel cubran periodos equivalentes. Si el mes actual aún está parcial, usa principalmente FNR %, MC % e incidencias por 1,000 líneas.")

                _history_export=io.BytesIO()
                with pd.ExcelWriter(_history_export,engine="openpyxl") as _hist_writer:
                    _history_compare.to_excel(_hist_writer,index=False,sheet_name="Comparativo pickers")
                    pd.DataFrame([{
                        "Periodo anterior":_history_snapshot.get("periodo",""),
                        "Líneas anterior":_prev_lines,
                        "Líneas actual":_cur_lines,
                        "FNR anterior":_prev_fnr,
                        "FNR actual":_cur_fnr,
                        "MC anterior":_prev_mc,
                        "MC actual":_cur_mc,
                        "Incidencias anterior":_prev_inc,
                        "Incidencias actual":_cur_inc,
                    }]).to_excel(_hist_writer,index=False,sheet_name="Resumen")
                st.download_button(
                    "⬇️ Descargar comparativo histórico",
                    _history_export.getvalue(),
                    file_name=f"Comparativo_{re.sub(r'[^A-Za-z0-9_-]+','_',str(_history_snapshot.get('periodo','historico')))}_vs_actual.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_monthly_history_comparison",
                )

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
            ("Pickers",f"{_picker_count:,}",f"{_picker_count/_reference_picker_count*100:.1f}% del universo · {_unregistered_count:,} sin registrar en el contexto" if _reference_picker_count else f"{_unregistered_count:,} sin registrar en el contexto","blue"),
        ])
        _outside_goal=int(((s_bodega.ESTADO=="🔴 FUERA DE OBJETIVO") & s_bodega.CATEGORIA.eq("Picker")).sum())
        render_soft_kpis([
            ("FNR operativo",f"{fnr_rate:.2f}%" if fnr_rate is not None else "N/D","Excel operativo · meta < 1.50%",_severity_tone(fnr_rate,FNR_OBJ*.8,FNR_OBJ)),
            ("MC operativo",f"{mc_rate:.2f}%" if mc_rate is not None else "N/D","Excel operativo · meta < 1.00%",_severity_tone(mc_rate,MC_OBJ*.8,MC_OBJ)),
            ("Fuera objetivo",f"{_outside_goal:,}",f"de {_picker_count:,} pickers · {_unregistered_count:,} sin registrar en el contexto","red" if _outside_goal else "green"),
        ])
        if fnr_rate is not None and mc_rate is not None:
            st.caption(f"Cálculo operativo desde Excel (por pedido): {fnr_pedidos:,} con FNR / {total_pedidos:,} pedidos = {fnr_rate:.2f}% · {mc_pedidos:,} con MC / {total_pedidos:,} pedidos = {mc_rate:.2f}%. Este cálculo es distinto al acumulado importado de Power BI.")
        else:
            st.caption("No hay suficientes pedidos para calcular el KPI mensual.")

        st.subheader("Comparativo del contexto")
        render_turn_comparison(base_bodega,base_reference,fnr_bodega,fnr_reference,mc_bodega,mc_reference)

        st.subheader("Indicador operativo por líneas")
        fnr_lines_rate=F/lines*100 if lines else None
        mc_lines_rate=M/lines*100 if lines else None
        render_soft_kpis([
            ("FNR / líneas",f"{fnr_lines_rate:.2f}%" if fnr_lines_rate is not None else "N/D",f"Objetivo < {FNR_OBJ:.2f}%",_severity_tone(fnr_lines_rate,FNR_OBJ*.8,FNR_OBJ)),
            ("MC / líneas",f"{mc_lines_rate:.2f}%" if mc_lines_rate is not None else "N/D",f"Objetivo < {MC_OBJ:.2f}%",_severity_tone(mc_lines_rate,MC_OBJ*.8,MC_OBJ)),
        ])
        st.caption(f"Semáforo por picker: FNR verde < {FNR_OBJ*.8:.2f}%, amarillo {FNR_OBJ*.8:.2f}–<{FNR_OBJ:.2f}%, rojo ≥ {FNR_OBJ:.2f}% · MC verde < {MC_OBJ*.8:.2f}%, amarillo {MC_OBJ*.8:.2f}–<{MC_OBJ:.2f}%, rojo ≥ {MC_OBJ:.2f}%.")
        st.caption("En la tabla, FNR_% y MC_% significan incidencias divididas entre líneas × 100; el indicador mensual usa pedidos.")
        st.subheader("Detalle por picker")
        st.dataframe(_style_severity(s_bodega,{"FNR_%":(FNR_OBJ*.8,FNR_OBJ),"MC_%":(MC_OBJ*.8,MC_OBJ)}),use_container_width=True,hide_index=True)

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

if _tab_active(o):
    with o:
        st.subheader("📦 Auditoría rápida de pedidos")
        st.caption("Carga el archivo de pedidos, elige uno de la lista y guarda el resultado. Los archivos originales solo se consultan; las auditorías quedan guardadas por separado.")

        _audit_flash=st.session_state.pop("_order_audit_flash",None)
        if _audit_flash:
            st.success(_audit_flash)

        audit_upload=st.file_uploader(
            "1. Cargar pedidos (Excel o CSV)",
            type=["xlsx","xls","csv"],
            key="order_audit_upload",
            help="Usa el archivo de picking exportado. No se modifica ni se reemplaza.",
        )
        if audit_upload is None:
            st.info("Carga el archivo del día para ver los pedidos pendientes de auditoría.")
            st.markdown("**Columnas necesarias:** pedido, hora/slot, SKU, artículo y cantidad pedida. Picker y cantidad pickeada son opcionales.")
        else:
            try:
                _audit_raw=audit_upload.getvalue()
                if audit_upload.name.lower().endswith(".csv"):
                    try:
                        _audit_sheets={"Datos":pd.read_csv(io.BytesIO(_audit_raw),sep=None,engine="python",dtype=str,keep_default_na=False)}
                    except UnicodeDecodeError:
                        _audit_sheets={"Datos":pd.read_csv(io.BytesIO(_audit_raw),sep=None,engine="python",dtype=str,keep_default_na=False,encoding="latin-1")}
                else:
                    _audit_sheets=pd.read_excel(io.BytesIO(_audit_raw),sheet_name=None,dtype=str,keep_default_na=False)
                _audit_sheets={str(k):v for k,v in _audit_sheets.items() if v is not None and not v.empty}
                if not _audit_sheets:
                    st.error("El archivo no contiene filas con datos.")
                else:
                    _sheet_name=st.selectbox("Hoja",list(_audit_sheets),key="order_audit_sheet") if len(_audit_sheets)>1 else next(iter(_audit_sheets))
                    _audit_df=_audit_sheets[_sheet_name].copy()
                    _audit_df.columns=[str(c).strip() for c in _audit_df.columns]
                    _audit_cols=list(_audit_df.columns)
                    _normalized_cols={norm(c):c for c in _audit_cols}

                    def _audit_guess(aliases):
                        for _alias in aliases:
                            if norm(_alias) in _normalized_cols:
                                return _normalized_cols[norm(_alias)]
                        return "— Sin columna —"

                    _fields=[
                        ("pedido","Pedido",["orden","order","numero_pedido","número de pedido","order_number","order_no","order_id","numero_orden","pedido"]),
                        ("slot","Hora / slot",["slot","horario","hora entrega","hora de entrega","delivery slot","time slot","hora pedido","hora del pedido","order_time","created_at","created","fecha_hora","order_created_at","hora","time"]),
                        ("sku","SKU",["sku","seller_sku","product_sku","codigo_sku","código sku","item_sku"]),
                        ("articulo","Artículo",["producto","product","articulo","artículo","product_name","item_name","item","nombre_articulo"]),
                        ("cantidad_pedida","Cantidad",["qty pedido","qtypedido","cantidad pedida","cantidad solicitada","quantity ordered","qty_ordered","quantity","qty","cantidad","unidades","cant"]),
                        ("picker","Picker",["picker","picker asignado","picker_name","nombre picker","pickeador"]),
                        ("cantidad_pickeada","Cantidad pickeada",["qty picked","qtypicked","cantidad pickeada","cantidad preparada","quantity picked"]),
                        ("estatus","Estatus",["estatus orden","order status","estatus","status"]),
                    ]
                    with st.expander("⚙️ Corregir columnas detectadas",expanded=False):
                        st.caption("Normalmente se detectan solas. Cambia solo las que aparezcan incorrectas.")
                        _mapping={}
                        _required_keys={"pedido","slot","sku","articulo","cantidad_pedida"}
                        for _key,_label,_aliases in _fields:
                            _options=["— Sin columna —"]+_audit_cols
                            _default=_audit_guess(_aliases)
                            _mapping[_key]=st.selectbox(
                                _label,_options,index=_options.index(_default),
                                key=f"order_audit_map_{_key}",
                            )
                    if any(_mapping.get(_k)=="— Sin columna —" for _k in _required_keys):
                        st.warning("No pude detectar todas las columnas necesarias. Abre «Corregir columnas detectadas» y selecciona pedido, hora/slot, SKU, artículo y cantidad.")
                    else:
                        _audit_data=pd.DataFrame({
                            "Número de pedido":_audit_df[_mapping["pedido"]].astype(str).str.strip(),
                            "Slot":_audit_df[_mapping["slot"]].astype(str).str.strip(),
                            "SKU":_audit_df[_mapping["sku"]].astype(str).str.strip(),
                            "Artículo":_audit_df[_mapping["articulo"]].astype(str).str.strip(),
                            "Cantidad pedida":_audit_df[_mapping["cantidad_pedida"]].astype(str).str.strip(),
                            "Cantidad pickeada":_audit_df[_mapping["cantidad_pickeada"]].astype(str).str.strip() if _mapping["cantidad_pickeada"]!="— Sin columna —" else "",
                            "Picker":_audit_df[_mapping["picker"]].astype(str).str.strip() if _mapping["picker"]!="— Sin columna —" else "",
                            "Estatus":_audit_df[_mapping["estatus"]].astype(str).str.strip() if _mapping["estatus"]!="— Sin columna —" else "",
                        })
                        _audit_data=_audit_data[
                            _audit_data["Número de pedido"].ne("")
                            & _audit_data["Número de pedido"].str.casefold().ne("nan")
                            & (_audit_data["Artículo"].ne("") | _audit_data["SKU"].ne(""))
                        ]
                        if _audit_data.empty:
                            st.error("No encontré pedidos con artículo/SKU. Revisa el archivo o las columnas.")
                        else:
                            def _slot_minutes(value):
                                _match=re.search(r"(?<!\d)(\d{1,2}):(\d{2})(?!\d)",str(value or ""))
                                if not _match: return np.nan
                                _hour,_minute=int(_match.group(1)),int(_match.group(2))
                                return _hour*60+_minute if 0<=_hour<=23 and 0<=_minute<=59 else np.nan

                            _audit_data["_slot_min"]=_audit_data["Slot"].map(_slot_minutes)
                            _slot_stats=_audit_data.groupby("Número de pedido",as_index=False).agg(
                                _slot_missing=("_slot_min",lambda values:bool(values.isna().any())),
                                _slot_max=("_slot_min","max"),
                                _slot_min=("_slot_min","min"),
                            )
                            _included_orders=set(_slot_stats.loc[~_slot_stats["_slot_missing"],"Número de pedido"])
                            _bad_time_orders=int(_slot_stats["_slot_missing"].sum())
                            _eligible=_audit_data[_audit_data["Número de pedido"].isin(_included_orders)].drop(columns="_slot_min").copy()

                            if _eligible.empty:
                                st.warning(f"No hay pedidos con hora legible. Pedidos sin hora válida: {_bad_time_orders:,}.")
                            else:
                                with st.spinner("Revisando señales FNR/MC de los últimos 30 días…"):
                                    _audit_hist_fnr,_audit_hist_mc,_audit_hist_meta=audit_history_30days(
                                        uf.getvalue(),um.getvalue(),roster,base,store,
                                    )
                                _last_hist_date=datetime.now(ZoneInfo("America/Mexico_City")).date()-timedelta(days=1)
                                _first_hist_date=_last_hist_date-timedelta(days=29)
                                _historical_fnr=int(pd.to_numeric(
                                    _audit_hist_fnr.get("INCIDENCIAS",pd.Series(dtype=float)),
                                    errors="coerce",
                                ).fillna(0).sum())
                                _historical_mc=int(pd.to_numeric(
                                    _audit_hist_mc.get("INCIDENCIAS",pd.Series(dtype=float)),
                                    errors="coerce",
                                ).fillna(0).sum())
                                st.caption(
                                    f"📅 Historial de 30 días ({_first_hist_date:%d/%m}–{_last_hist_date:%d/%m})"
                                    f" · FNR: {_historical_fnr:,} · MC: {_historical_mc:,}"
                                )
                                _missing_history_dates=sum(
                                    int(m.get("fechas_ausentes",0)) for m in _audit_hist_meta.values()
                                )
                                if _missing_history_dates:
                                    st.caption(
                                        f"⚠️ {_missing_history_dates:,} renglones sin fecha legible "
                                        "no cuentan en el historial."
                                    )
                                if _audit_hist_fnr.empty and _audit_hist_mc.empty:
                                    st.warning(
                                        "Sin registros históricos fechados de FNR/MC en los 30 días previos. "
                                        "La prioridad puede ser menos representativa."
                                    )
                                _audit_detail,_risk_summary=order_risk_analysis(
                                    _eligible,s,_audit_hist_fnr,_audit_hist_mc,store.get("order_audits",[])
                                )
                                _risk_summary=_risk_summary.merge(
                                    _slot_stats[["Número de pedido","_slot_min","_slot_max"]],
                                    on="Número de pedido",how="left",validate="one_to_one",
                                )
                                _risk_summary["Cobertura equipo"]=_risk_summary.apply(
                                    lambda rr: (
                                        "Madrugada"
                                        if pd.notna(rr["_slot_max"]) and rr["_slot_max"]<600
                                        else "Día"
                                        if pd.notna(rr["_slot_min"]) and rr["_slot_min"]>=600
                                        else "Slots mixtos"
                                    ),axis=1,
                                )
                                _shift_candidates={}
                                for _frame,_turn_col in ((roster,"TURNO_MAESTRO"),(s,"TURNO")):
                                    if _frame is None or _frame.empty or _turn_col not in _frame.columns:
                                        continue
                                    for _,_rr in _frame.iterrows():
                                        _shift=str(_rr.get(_turn_col,"") or "").strip()
                                        if not _shift or norm(_shift) in {"nan","none","no_especificado","no_asignado"}:
                                            continue
                                        for _identity in [_rr.get("PICKER",""),_rr.get("CORREO",""),_rr.get("_SOURCE_PICKER","")]:
                                            _identity=str(_identity or "").strip()
                                            if not _identity or _identity.lower()=="nan":
                                                continue
                                            _keys=["n:"+person_key(_identity),"t:"+token_key(_identity)]
                                            if "@" in _identity:
                                                _keys.append("e:"+email_key(_identity))
                                            for _key in _keys:
                                                if len(_key)>2:
                                                    _shift_candidates.setdefault(_key,set()).add(_shift)

                                def _shift_for_audit_picker(value):
                                    _value=str(value or "").strip()
                                    _keys=["e:"+email_key(_value)] if "@" in _value else []
                                    _keys.extend(["n:"+person_key(_value),"t:"+token_key(_value)])
                                    for _key in _keys:
                                        _matches=_shift_candidates.get(_key,set())
                                        if len(_matches)==1:
                                            return next(iter(_matches))
                                    return "Sin identificar"

                                _shift_by_order=_audit_detail.groupby("Número de pedido")["Picker relacionado"].agg(
                                    lambda values:", ".join(sorted({_shift_for_audit_picker(v) for v in values if str(v).strip()})) or "Sin identificar"
                                )
                                _risk_summary["Turno picker"]=_risk_summary["Número de pedido"].map(_shift_by_order).fillna("Sin identificar")
                                _risk_summary["Picker nocturno"]=_risk_summary["Turno picker"].map(
                                    lambda value:any(norm(part).startswith("nocturn") or norm(part) in {"noche","22"}
                                                     for part in str(value).split(", "))
                                )
                                _risk_summary["Puntaje selección"]=_risk_summary["Puntaje alarma"]+_risk_summary["Picker nocturno"].astype(int)*2

                                _saved_audits=store.get("order_audits",[]) or []
                                _audited_orders={
                                    norm(a.get("pedido","")) for a in _saved_audits
                                    if audit_record_is_complete(a) and str(a.get("pedido","")).strip()
                                }
                                _risk_summary["Ya auditado"]=_risk_summary["Número de pedido"].map(lambda v:norm(v) in _audited_orders)

                                _coverage_label=st.selectbox(
                                    "2. Pedidos a revisar",
                                    ["Mi turno: slots antes de 10:00","Equipo de día: slots desde 10:00","Todo el día"],
                                    key="audit_coverage_simple",
                                )
                                _coverage_mode="early" if _coverage_label.startswith("Mi turno") else "day" if _coverage_label.startswith("Equipo") else "all"
                                _coverage_summary=audit_coverage_filter(_risk_summary,_coverage_mode)
                                with st.expander("Ajustar cantidad sugerida",expanded=False):
                                    _audit_percent=st.slider(
                                        "Porcentaje máximo por hora",5,30,20,1,
                                        format="%d%%",key="audit_share_percent",
                                    )
                                    _hide_done=st.checkbox("Ocultar pedidos ya auditados",value=True,key="audit_hide_done")
                                if "_audit_percent" not in locals():
                                    _audit_percent=20
                                    _hide_done=True

                                _available=_coverage_summary[
                                    ~_coverage_summary["Ya auditado"] if _hide_done
                                    else pd.Series(True,index=_coverage_summary.index)
                                ].copy()
                                _critical_hourly=audit_hourly_critical_queue(
                                    _available,_audit_percent,population=_coverage_summary,
                                )
                                _queue_mode=st.radio(
                                    "Lista",
                                    ["Sugeridos primero","Todos los pendientes"],
                                    horizontal=True,
                                    key="audit_queue_mode_simple",
                                )
                                _queue=_critical_hourly.copy() if _queue_mode=="Sugeridos primero" else _available.copy()
                                if "Puntaje selección" in _queue.columns:
                                    _queue=_queue.sort_values(["Puntaje selección","Número de pedido"],ascending=[False,True])
                                if _queue.empty and _queue_mode=="Sugeridos primero" and not _available.empty:
                                    st.info("No hay alertas prioritarias para esta hora. Cambia a «Todos los pendientes» para revisar el resto.")
                                if _bad_time_orders:
                                    st.caption(f"{_bad_time_orders:,} pedidos no tienen una hora reconocible y no entran en la lista.")
                                if not _coverage_summary.empty:
                                    _n_done=int(_coverage_summary["Ya auditado"].sum())
                                    _q1,_q2,_q3=st.columns(3)
                                    _q1.metric("Pedidos en esta cobertura",f"{len(_coverage_summary):,}")
                                    _q2.metric("Pendientes",f"{len(_available):,}")
                                    _q3.metric("Ya auditados",f"{_n_done:,}")
                                st.caption("Las sugerencias usan señales históricas FNR/MC y picker/artículo para ordenar la revisión; sirven como guía, no determinan la causa.")

                                if not _queue.empty:
                                    _search=st.text_input("Buscar número de pedido",key="audit_order_search_simple")
                                    if _search.strip():
                                        _queue=_queue[_queue["Número de pedido"].astype(str).str.contains(re.escape(_search.strip()),case=False,na=False)]
                                    _compact_cols=[c for c in ["Número de pedido","Hora auditoría","Turno picker","Prioridad","Pickers asignados"] if c in _queue.columns]
                                    st.dataframe(
                                        _queue[_compact_cols].rename(columns={"Número de pedido":"Pedido","Hora auditoría":"Hora","Turno picker":"Turno","Pickers asignados":"Picker"}),
                                        use_container_width=True,hide_index=True,
                                    )
                                    _order_options=_queue["Número de pedido"].astype(str).tolist()
                                    if _order_options:
                                        _order=st.selectbox("3. Seleccionar pedido",_order_options,key="order_audit_selected_simple")
                                        _lines=_audit_detail[_audit_detail["Número de pedido"].astype(str)==str(_order)].copy().reset_index(drop=True)
                                        _order_row=_queue[_queue["Número de pedido"].astype(str)==str(_order)].iloc[0]
                                        _slot_values=[v for v in _lines["Slot"].drop_duplicates().astype(str).tolist() if v and v.casefold()!="nan"]
                                        _order_slot="–".join(_slot_values) if _slot_values else "Sin dato"

                                        _head1,_head2,_head3=st.columns(3)
                                        _head1.metric("Pedido",str(_order))
                                        _head2.metric("Hora / slot",_order_slot)
                                        _head3.metric("Picker",str(_order_row.get("Pickers asignados","Sin asignar")))
                                        st.caption(f"Turno del picker: {_order_row.get('Turno picker','Sin identificar')}")
                                        _item_cols=[c for c in ["SKU","Artículo","Picker relacionado","Cantidad pedida","Cantidad pickeada"] if c in _lines.columns]
                                        _items=_lines[_item_cols].copy().rename(columns={"Picker relacionado":"Picker"})
                                        st.markdown("**Artículos del pedido**")
                                        st.dataframe(_items,use_container_width=True,hide_index=True)
                                        def _audit_order_incidents(source,order):
                                            if source is None or source.empty or "ORDER_NUMBER" not in source.columns:
                                                return pd.DataFrame()
                                            _mask=source["ORDER_NUMBER"].astype(str).str.strip().map(norm)==norm(order)
                                            return source.loc[_mask].copy()

                                        _audit_fnr=_audit_order_incidents(fnr,_order)
                                        _audit_mc=_audit_order_incidents(mc,_order)
                                        def _audit_incident_count(frame):
                                            if frame.empty: return 0
                                            if "INCIDENCIAS" in frame.columns:
                                                return int(pd.to_numeric(frame["INCIDENCIAS"],errors="coerce").fillna(0).sum())
                                            return int(len(frame))
                                        _audit_fnr_count=_audit_incident_count(_audit_fnr)
                                        _audit_mc_count=_audit_incident_count(_audit_mc)
                                        with st.expander("📱 Asignar validación móvil (supervisores)",expanded=False):
                                            from mobile_audit_admin import render_mobile_assignment
                                            render_mobile_assignment(
                                                _order,_order_slot,_order_row,_lines,roster,s,_audit_raw,
                                                _audit_fnr_count,_audit_mc_count,
                                            )
                                        with st.expander("Antecedentes de este pedido (FNR/MC)",expanded=False):
                                            _x1,_x2=st.columns(2)
                                            _x1.metric("FNR",_audit_fnr_count)
                                            _x2.metric("MC",_audit_mc_count)
                                            if _audit_fnr_count or _audit_mc_count:
                                                st.caption("La coincidencia ayuda a investigar; por sí sola no prueba el origen.")
                                                for _kind,_frame in (("FNR",_audit_fnr),("MC",_audit_mc)):
                                                    if not _frame.empty:
                                                        _show_cols=[c for c in ["ORDER_NUMBER","PRODUCTO","PICKER","AREA","INCIDENCIAS"] if c in _frame.columns]
                                                        st.markdown(f"**{_kind} relacionados**")
                                                        st.dataframe(_frame[_show_cols],use_container_width=True,hide_index=True)
                                            else:
                                                st.caption("Sin coincidencia actual en los archivos FNR/MC cargados.")

                                        st.markdown("### 4. Guardar auditoría")
                                        _existing_audit=next((a for a in reversed(_saved_audits) if norm(a.get("pedido",""))==norm(_order)),None)
                                        if _existing_audit:
                                            st.caption(f"Auditoría anterior: {_existing_audit.get('fecha_auditoria','')} · {_existing_audit.get('resultado','')}. La nueva revisión se agrega al historial.")
                                        _audit_key=hashlib.sha256((str(_order)+"|"+hashlib.sha256(_audit_raw).hexdigest()).encode()).hexdigest()[:14]
                                        _result=st.selectbox(
                                            "Resultado",
                                            ["Pedido correcto","Con diferencias","No se pudo validar"],
                                            key="audit_result_simple_"+_audit_key,
                                        )
                                        _line_options=[
                                            f"{idx+1}. {row.get('SKU','')} · {row.get('Artículo','')} · Cantidad {row.get('Cantidad pedida','')}"
                                            for idx,row in _lines.iterrows()
                                        ]
                                        with st.form("audit_save_form_"+_audit_key,clear_on_submit=False):
                                            _auditor=st.text_input("Quién revisó",key="audit_responsable_simple_"+_audit_key)
                                            if _result=="Con diferencias":
                                                _diff_items=st.multiselect("Artículos con diferencia",_line_options,key="audit_diff_items_"+_audit_key)
                                                _diff_notes=st.text_area("Qué debía haber y qué encontraste",key="audit_diff_notes_"+_audit_key)
                                            else:
                                                _diff_items=[]
                                                _diff_notes=""
                                            _audit_notes=st.text_area(
                                                "Motivo por el que quedó pendiente" if _result=="No se pudo validar"
                                                else "Observaciones (opcional)",
                                                key="audit_notes_simple_"+_audit_key,
                                            )
                                            if _result=="No se pudo validar":
                                                st.caption("Este intento se registrará como pendiente; el pedido seguirá disponible.")
                                                _approved=False
                                            else:
                                                _approved=st.checkbox(
                                                    "Confirmo que revisé el pedido completo",
                                                    key="audit_approved_simple_"+_audit_key,
                                                )
                                            _save_audit=st.form_submit_button("Guardar revisión",type="primary",use_container_width=True)

                                        if _save_audit:
                                            if not _auditor.strip():
                                                st.error("Escribe quién realizó la revisión.")
                                            elif _result!="No se pudo validar" and not _approved:
                                                st.error("Confirma que revisaste el pedido completo.")
                                            elif _result=="No se pudo validar" and not _audit_notes.strip():
                                                st.error("Describe por qué el pedido sigue pendiente.")
                                            elif _result=="Con diferencias" and (not _diff_items or not _diff_notes.strip()):
                                                st.error("Elige el artículo y describe la diferencia.")
                                            else:
                                                _line_records=[]
                                                for _idx,_line in _lines.iterrows():
                                                    _label=_line_options[_idx]
                                                    _line_records.append({
                                                        "SKU":str(_line.get("SKU","")),
                                                        "Artículo":str(_line.get("Artículo","")),
                                                        "Picker relacionado":str(_line.get("Picker relacionado","")),
                                                        "Cantidad pedida":str(_line.get("Cantidad pedida","")),
                                                        "Cantidad pickeada":str(_line.get("Cantidad pickeada","")),
                                                        "Diferencia encontrada":_label in _diff_items,
                                                        "Corrección / observación":_diff_notes.strip() if _label in _diff_items else "",
                                                    })
                                                _audit_record={
                                                    "id":hashlib.sha256((str(_order)+"|"+datetime.now().isoformat()).encode()).hexdigest()[:18],
                                                    "pedido":str(_order),
                                                    "slot":str(_order_slot),
                                                    "cobertura_equipo":str(_order_row.get("Cobertura equipo","")),
                                                    "turno_picker":str(_order_row.get("Turno picker","Sin identificar")),
                                                    "fecha_auditoria":datetime.now(ZoneInfo("America/Mexico_City")).isoformat(timespec="seconds"),
                                                    "picker":str(_order_row.get("Pickers asignados","")).strip(),
                                                    "auditor":_auditor.strip(),
                                                    "resultado":_result,
                                                    "pedido_validado":bool(_approved and _result!="No se pudo validar"),
                                                    "observaciones":_audit_notes.strip(),
                                                    "diferencias":len(_diff_items),
                                                    "lineas":_line_records,
                                                    "fnr_al_guardar":_audit_fnr_count,
                                                    "mc_al_guardar":_audit_mc_count,
                                                    "prioridad_preventiva":str(_order_row.get("Prioridad","")),
                                                    "nivel_foco":str(_order_row.get("Prioridad","")),
                                                    "puntaje_foco":int(_order_row.get("Puntaje foco",0) or 0),
                                                    "factores_foco":str(_order_row.get("Factores","")),
                                                    "alarma":str(_order_row.get("Alarma","")),
                                                    "fuente_archivo":audit_upload.name,
                                                }
                                                store.setdefault("order_audits",[]).append(_audit_record)
                                                try:
                                                    save_store(store)
                                                except Exception as _audit_save_error:
                                                    if store.get("order_audits") and store["order_audits"][-1].get("id")==_audit_record["id"]:
                                                        store["order_audits"].pop()
                                                    st.error(f"No se pudo guardar la auditoría: {_audit_save_error}")
                                                else:
                                                    st.session_state["_order_audit_flash"]=(
                                                        f"Pedido {_order}: intento registrado; sigue pendiente."
                                                        if _result=="No se pudo validar"
                                                        else f"Pedido {_order}: auditoría validada y guardada."
                                                    )
                                                    st.rerun()
                                else:
                                    st.info("No hay pedidos pendientes en esta lista.")
                                    if _coverage_summary.empty:
                                        st.caption("No se encontraron pedidos dentro de esta cobertura.")
            except Exception as _audit_error:
                st.error(f"No pude leer el archivo de pedidos: {_audit_error}")

        _saved_audits=store.get("order_audits",[]) or []
        with st.expander(f"📚 Historial de auditorías ({len(_saved_audits)})",expanded=False):
            if _saved_audits:
                _history=pd.DataFrame([{
                    "Pedido":rec.get("pedido",""),
                    "Fecha":rec.get("fecha_auditoria",""),
                    "Picker":rec.get("picker",""),
                    "Revisó":rec.get("auditor",""),
                    "Resultado":rec.get("resultado",""),
                    "Estado":"Validado" if audit_record_is_complete(rec) else "Pendiente",
                    "Diferencias":rec.get("diferencias",0),
                    "FNR":rec.get("fnr_al_guardar",0),
                    "MC":rec.get("mc_al_guardar",0),
                    "Observaciones":rec.get("observaciones",""),
                } for rec in _saved_audits])
                _history=_history.sort_values("Fecha",ascending=False)
                _history_search=st.text_input("Buscar en el historial",key="audit_history_search_simple")
                if _history_search.strip():
                    _history=_history[_history["Pedido"].astype(str).str.contains(re.escape(_history_search.strip()),case=False,na=False)]
                st.dataframe(_history,use_container_width=True,hide_index=True)
            else:
                st.caption("Todavía no hay auditorías guardadas.")


if _tab_active(b):
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
            st.markdown(f"<div class='justo-card'><div class='justo-kicker'>Ficha de picker</div><div class='justo-title'>{_h(sp)}</div><div class='justo-muted'>Turno: {_h(r.TURNO)} · Supervisor: {_h(r.SUPERVISOR)} · Área: {_h(r.AREA_BASE)} · Usuario: {_h(r.CORREO)}</div></div>", unsafe_allow_html=True)
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
            _picker_email=str(r.get("CORREO","")).strip()
            _picker_identity={
                "PICKER":sp,"CORREO":_picker_email,
                "TURNO":str(r.get("TURNO","")).strip(),
                "SUPERVISOR":str(r.get("SUPERVISOR","")).strip(),
                "AREA_BASE":str(r.get("AREA_BASE","")).strip(),
            }
            rec=followup_email_record(store,sp,_picker_email,aliases=_src_alias,create=False)
            st.subheader("📝 Retroalimentación y seguimiento")
            st.caption("Solo se solicita el tipo de seguimiento. La retroalimentación opcional queda dentro del mismo registro.")
            with st.form(f"accion_picker_form_{sp}", clear_on_submit=True):
                act_type=st.selectbox("Categoría",FOLLOWUP_CATEGORIES)
                act_sup=st.text_input("Supervisor",value=str(r.get("SUPERVISOR","")))
                act_motivo=st.text_area("Motivo / detalle")
                act_retro=st.text_area("Retroalimentación (opcional)",placeholder="Observación de retroalimentación, si aplica…")
                if st.form_submit_button("Guardar seguimiento", type="primary"):
                    _seguimiento={
                        "id":uuid.uuid4().hex,
                        "fecha":datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "accion":act_type,"supervisor":act_sup.strip() or "No especificado",
                        "motivo":act_motivo.strip(),"retroalimentacion":act_retro.strip()
                    }
                    try:
                        save_followup_entry(
                            store,sp,"acciones",_seguimiento,aliases=_src_alias,
                            identity_email=_picker_email,identity_fields=_picker_identity,
                        )
                    except Exception as exc:
                        st.error(str(exc))
                    else:
                        st.success("Seguimiento respaldado en la nube.")
                        st.rerun()

            st.caption("Los enlaces reutilizables se administran en la pestaña Seguimiento.")

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

if _tab_active(e):
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
        st.caption("Las celdas de tasa por líneas usan el mismo semáforo de objetivos de FNR y MC.")
        st.subheader("FNR por turno")
        st.dataframe(_style_severity(groups(fsel,bsel,"TURNO"),{"% / LINEAS":(FNR_OBJ*.8,FNR_OBJ)}),use_container_width=True,hide_index=True)
        st.subheader("MC por turno")
        st.dataframe(_style_severity(groups(msel,bsel,"TURNO"),{"% / LINEAS":(MC_OBJ*.8,MC_OBJ)}),use_container_width=True,hide_index=True)
        st.subheader("FNR por área")
        st.dataframe(_style_severity(groups(fsel,bsel,"AREA"),{"% / LINEAS":(FNR_OBJ*.8,FNR_OBJ)}),use_container_width=True,hide_index=True)
        st.subheader("MC por área")
        st.dataframe(_style_severity(groups(msel,bsel,"AREA"),{"% / LINEAS":(MC_OBJ*.8,MC_OBJ)}),use_container_width=True,hide_index=True)
        st.warning("El % / líneas por área solo aparece si existe un denominador real de líneas por área.")

if _tab_active(x):
    with x:
        st.subheader("🧰 Herramientas")
        st.caption("Herramientas de administración fuera del análisis operativo diario.")
        st.markdown("### ✉️ Cruce de correos")
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

        st.divider()
        st.markdown("### 📊 Análisis puntual de FR · otra área")
        st.caption("Carga un Excel cuando necesites revisarlo. El archivo se analiza aquí y no se incorpora a los datos operativos ni al respaldo en nube.")
        _other_area_upload=st.file_uploader("Excel de FR de otra área",type=["xlsx","xls"],key="oneoff_area_fr_upload")
        if _other_area_upload is not None:
            try:
                _other_fr=parse_oneoff_area_fr(_other_area_upload.getvalue())
                if _other_fr.empty:
                    st.info("El archivo no contiene filas con fecha, artículo y departamento.")
                else:
                    _dept_counts=(_other_fr.groupby("Departamento").size().rename("Reportes FR").reset_index()
                        .sort_values("Reportes FR",ascending=False))
                    _article_counts=(_other_fr.groupby(["Departamento","Artículo"]).size().rename("Reportes FR").reset_index()
                        .sort_values(["Reportes FR","Artículo"],ascending=[False,True]))
                    _article_totals=(_other_fr.groupby("Artículo").size().sort_values(ascending=False))
                    _daily=(_other_fr.groupby("Fecha").size().rename("Reportes FR")
                        .reindex(pd.date_range(_other_fr["Fecha"].min().normalize(),_other_fr["Fecha"].max().normalize(),freq="D"),fill_value=0)
                        .rename_axis("Fecha").reset_index())
                    _valid_orders=_other_fr["Pedido"].dropna().astype(str).str.strip()
                    _valid_orders=_valid_orders[_valid_orders.ne("")]
                    _top_day=_daily.loc[_daily["Reportes FR"].idxmax()]
                    _top_dept=_dept_counts.iloc[0]
                    _top_article=_article_totals.index[0]
                    _top_article_count=int(_article_totals.iloc[0])
                    st.caption(f"{len(_other_fr):,} registros de artículo · {_valid_orders.nunique():,} pedidos · {_other_fr['Artículo'].nunique():,} artículos · {_other_fr['Departamento'].nunique():,} departamentos. Cada registro cuenta un artículo reportado, no unidades físicas.")
                    _a1,_a2,_a3,_a4=st.columns(4)
                    _a1.metric("Registros FR",f"{len(_other_fr):,}")
                    _a2.metric("Departamento con más",str(_top_dept["Departamento"]),f"{int(_top_dept['Reportes FR']):,} registros")
                    _a3.metric("Artículo más reportado",str(_top_article),f"{_top_article_count:,} registros")
                    _a4.metric("Día con más reportes",_top_day["Fecha"].strftime("%d/%m/%Y"),f"{int(_top_day['Reportes FR']):,} registros")
                    _ch1,_ch2=st.columns(2)
                    with _ch1:
                        st.markdown("**Reportes FR por departamento**")
                        st.bar_chart(_dept_counts.set_index("Departamento"),horizontal=True)
                    with _ch2:
                        st.markdown("**Reportes FR por día**")
                        st.line_chart(_daily.set_index("Fecha"))
                    st.markdown("**Conteo por artículo**")
                    _f1,_f2=st.columns([1,2])
                    with _f1:
                        _dept_options=["Todos"]+_dept_counts["Departamento"].tolist()
                        _dept_choice=st.selectbox("Departamento",_dept_options,key="oneoff_fr_department_filter")
                    with _f2:
                        _article_query=st.text_input("Buscar artículo",key="oneoff_fr_article_search",placeholder="Escribe parte del nombre")
                    _article_view=_article_counts if _dept_choice=="Todos" else _article_counts[_article_counts["Departamento"]==_dept_choice]
                    if _article_query.strip():
                        _article_view=_article_view[_article_view["Artículo"].str.contains(_article_query.strip(),case=False,na=False)]
                    st.dataframe(_article_view,use_container_width=True,hide_index=True,height=420)
            except Exception as _other_fr_error:
                st.error(f"No pude analizar ese archivo: {_other_fr_error}")

if _tab_active(g):
    with g:
        ctx,_,_,_=global_context(context_identity)
        st.subheader("👥 Supervisores")
        st.caption("Vista operativa bajo el mismo contexto global. Los supervisores registrados por correo se excluyen automáticamente de incidencias y filtros de pickers.")
        _leaders=pd.DataFrame()
        if isinstance(roster,pd.DataFrame) and not roster.empty and "PERSONAL_TIPO" in roster.columns:
            _leader_mask=roster["PERSONAL_TIPO"].fillna("").astype(str).map(norm).isin({"supervisor","subgerente","jefe"})
            _leaders=roster.loc[_leader_mask].copy()
        st.markdown("#### Líderes de Outbound OKI · plantilla")
        if not _leaders.empty:
            _leader_columns=[c for c in ["PICKER","PERSONAL_TIPO","TURNO_MAESTRO","CORREO","AREA_MAESTRO"] if c in _leaders.columns]
            st.dataframe(
                _leaders[_leader_columns].rename(columns={
                    "PICKER":"Nombre","PERSONAL_TIPO":"Cargo","TURNO_MAESTRO":"Turno",
                    "CORREO":"Correo","AREA_MAESTRO":"Área",
                }),
                use_container_width=True,hide_index=True,
            )
        else:
            st.info("La plantilla actual no contiene personas con cargo Supervisor, Subgerente o Jefe.")
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
                .agg(PICKERS=("CORREO_KEY","nunique"),LINEAS=("LINEAS","sum"),FNR=("FNR","sum"),MC=("MC","sum"))
                .sort_values("FNR",ascending=False))
            sup_summary["FNR %"]=safe_pct(sup_summary["FNR"],sup_summary["LINEAS"])
            sup_summary["MC %"]=safe_pct(sup_summary["MC"],sup_summary["LINEAS"])
            st.dataframe(_style_severity(sup_summary,{"FNR %":(FNR_OBJ*.8,FNR_OBJ),"MC %":(MC_OBJ*.8,MC_OBJ)}),use_container_width=True,hide_index=True)
            sup_focus=st.selectbox("Supervisor",["Todos"]+sorted([str(x) for x in sup_summary["SUPERVISOR"] if str(x).strip()]))
            if sup_focus!="Todos":
                st.dataframe(sup_data[sup_data["SUPERVISOR"]==sup_focus],use_container_width=True,hide_index=True)
        else:
            st.info("No hay datos para el contexto actual.")

if _tab_active(i):
    with i:
        st.subheader("📅 Faltas y retardos")
        st.caption("Carga los dos reportes para ver incidencias por persona y turno. Si una falta no trae turno, se completa desde la plantilla consolidada.")
        _shift_cfg=store.get("attendance_shift_config",{}) or {"Mañana":[6,7,8,9,10],"Intermedio":[11],"Tarde":[13,14],"Nocturno":[22]}
        with st.expander("⚙️ Configurar horarios de turno",expanded=False):
            _sc1,_sc2,_sc3,_sc4=st.columns(4)
            with _sc1: _morning=st.text_input("Mañana",value=",".join(map(str,_shift_cfg.get("Mañana",[6,7,8,9,10]))),key="shift_cfg_morning")
            with _sc2: _middle=st.text_input("Intermedio",value=",".join(map(str,_shift_cfg.get("Intermedio",[11]))),key="shift_cfg_middle")
            with _sc3: _afternoon=st.text_input("Tarde",value=",".join(map(str,_shift_cfg.get("Tarde",[13,14]))),key="shift_cfg_afternoon")
            with _sc4: _night=st.text_input("Nocturno",value=",".join(map(str,_shift_cfg.get("Nocturno",[22]))),key="shift_cfg_night")
            if st.button("Guardar horarios",key="save_shift_config"):
                try:
                    def _hours(text):
                        return sorted({int(x.strip()) for x in str(text).split(",") if x.strip() and 0<=int(x.strip())<=23})
                    _new_cfg={"Mañana":_hours(_morning),"Intermedio":_hours(_middle),"Tarde":_hours(_afternoon),"Nocturno":_hours(_night)}
                    if not all(_new_cfg.values()): raise ValueError("Cada turno debe tener al menos una hora.")
                    store["attendance_shift_config"]=_new_cfg
                    save_store(store)
                    st.success("Horarios guardados.")
                    st.rerun()
                except Exception:
                    st.error("Usa horas de 0 a 23 separadas por comas, por ejemplo: 6,7,8,9.")
        with st.container(border=True):
            u_abs,u_late=st.columns(2)
            with u_abs:
                attendance_absence_upload=st.file_uploader("① Reporte de faltas",type=["xlsx","xls"],key="attendance_absence_upload")
            with u_late:
                attendance_tardy_upload=st.file_uploader("② Reporte de retardos",type=["xlsx","xls"],key="attendance_tardy_upload")
            st.caption("Para actualizar los resultados, selecciona ambos archivos. El resumen agregado se conserva con los datos guardados de la app y se respalda en nube si está configurado.")

        if attendance_absence_upload is not None and attendance_tardy_upload is not None:
            _abs_bytes=attendance_absence_upload.getvalue(); _late_bytes=attendance_tardy_upload.getvalue()
            _att_context=json.dumps({
                "shift":store.get("attendance_shift_config",{}),
                "roster":[
                    {"p":str(r.get("PICKER","")),"c":str(r.get("CORREO","")),"t":str(r.get("TURNO_MAESTRO","")),"s":str(r.get("SUPERVISOR","")),"a":str(r.get("AREA_MAESTRO",""))}
                    for _,r in roster.iterrows()
                ] if isinstance(roster,pd.DataFrame) else [],
            },ensure_ascii=False,sort_keys=True).encode("utf-8")
            _fingerprint=hashlib.sha256(_abs_bytes+b"\0"+_late_bytes+b"\0"+_att_context).hexdigest()
            _old_meta=store.get("attendance_meta",{}) or {}
            if _fingerprint!=_old_meta.get("fingerprint"):
                try:
                    with st.spinner("Leyendo y cruzando los reportes…"):
                        _attendance_new=build_attendance_summary(
                            _abs_bytes,_late_bytes,attendance_absence_upload.name,attendance_tardy_upload.name,
                            roster_rows=roster,shift_config=store.get("attendance_shift_config")
                        )
                    _attendance_new["fingerprint"]=_fingerprint
                    store["attendance_summary"]=_attendance_new
                    store["attendance_meta"]={"fingerprint":_fingerprint,"faltas":_safe_filename(attendance_absence_upload.name),"retardos":_safe_filename(attendance_tardy_upload.name),"fecha_carga":_attendance_new["fecha_carga"]}
                    save_store(store)
                    attendance_summary=_attendance_new
                    attendance_people=attendance_summary.get("personas",[])
                    attendance_by_name={}
                    for _person in attendance_people:
                        _k=token_key(_person.get("Persona",""))
                        if _k: attendance_by_name.setdefault(_k,[]).append(_person)
                    st.success("Reportes cargados y cruzados. Los CURP se usan solo durante el cruce y no quedan guardados en el resumen.")
                except Exception as _att_error:
                    st.error(f"No pude leer los reportes: {_att_error}")
            else:
                attendance_summary=store.get("attendance_summary") or {}
                attendance_people=attendance_summary.get("personas",[]) if isinstance(attendance_summary,dict) else []
        elif attendance_absence_upload is not None or attendance_tardy_upload is not None:
            st.info("Selecciona también el otro archivo para poder cruzar las faltas y los retardos por persona.")

        attendance_summary=store.get("attendance_summary") or attendance_summary
        attendance_people=attendance_summary.get("personas",[]) if isinstance(attendance_summary,dict) else []
        if not attendance_people:
            st.info("Carga el reporte de faltas y el de retardos para ver el resumen.")
        else:
            _period=f"{attendance_summary.get('fecha_desde','')} a {attendance_summary.get('fecha_hasta','')}".strip(" a") or "Fechas no especificadas"
            st.caption(f"Periodo de los reportes: {_period} · Última carga: {attendance_summary.get('fecha_carga','')}")
            _hours,_mins=divmod(int(attendance_summary.get("minutos_acumulados",0)),60)
            _k1,_k2,_k3,_k4=st.columns(4)
            _k1.metric("Faltas registradas",f"{int(attendance_summary.get('total_faltas',0)):,}",f"{int(attendance_summary.get('personas_con_faltas',0)):,} personas")
            _k2.metric("Retardos registrados",f"{int(attendance_summary.get('total_retardos',0)):,}",f"{int(attendance_summary.get('personas_con_retardos',0)):,} personas")
            _k3.metric("Tiempo acumulado de retardos",f"{_hours} h {_mins:02d} min")
            _k4.metric("Personas en los reportes",f"{int(attendance_summary.get('total_personas',0)):,}")

            _shift_order={"Mañana":0,"Intermedio":1,"Tarde":2,"Nocturno":3,"Sin turno":4,"Otro horario":5,"Revisar turno":6}
            _tardy_shift_df=pd.DataFrame(attendance_summary.get("retardos_por_turno",[]))
            _absence_shift_df=pd.DataFrame(attendance_summary.get("faltas_por_turno",[]))
            if not _tardy_shift_df.empty:
                _tardy_shift_df["_orden"]=_tardy_shift_df["Turno"].map(lambda z:_shift_order.get(z,9))
                _tardy_shift_df=_tardy_shift_df.sort_values("_orden").drop(columns="_orden")
            if not _absence_shift_df.empty:
                _absence_shift_df["_orden"]=_absence_shift_df["Turno"].map(lambda z:_shift_order.get(z,9))
                _absence_shift_df=_absence_shift_df.sort_values("_orden").drop(columns="_orden")
            with st.expander("Resumen por turno",expanded=True):
                _shift_col1,_shift_col2=st.columns(2)
                with _shift_col1:
                    st.markdown("**Retardos por turno**")
                    if not _tardy_shift_df.empty:
                        _tardy_shift_df["Tiempo acumulado"]=_tardy_shift_df["Minutos acumulados"].map(lambda v:f"{int(v)//60} h {int(v)%60:02d} min")
                        st.dataframe(_tardy_shift_df[["Turno","Personas","Retardos","Tiempo acumulado"]],use_container_width=True,hide_index=True)
                with _shift_col2:
                    st.markdown("**Faltas por turno habitual**")
                    if not _absence_shift_df.empty:
                        st.dataframe(_absence_shift_df,use_container_width=True,hide_index=True)
                    st.caption("El archivo de faltas no incluye turno. Se usa el turno más frecuente de retardos de cada persona; ‘Sin turno’ indica que no hubo coincidencia en el reporte de retardos.")

            _att_df=pd.DataFrame(attendance_people)
            def _attendance_app_link(name):
                _manual=store.get("attendance_links",{}).get(token_key(name),"")
                if _manual:
                    for _bucket in app_people_by_name.values():
                        for _candidate in _bucket:
                            if _candidate["PICKER"]==_manual: return _manual,_candidate["CATEGORIA"]
                _matches=app_people_by_name.get(token_key(name),[])
                if len(_matches)==1:
                    return _matches[0]["PICKER"],_matches[0]["CATEGORIA"]
                if len(_matches)>1: return "Revisar coincidencia","Ambigua"
                return "Sin coincidencia exacta","Sin registrar"
            _links=_att_df["Persona"].map(_attendance_app_link)
            _att_df["Personal en la app"]=[x[0] for x in _links]
            _att_df["Categoría en la app"]=[x[1] for x in _links]
            _att_df["Turno"]=_att_df.apply(lambda row:str(row["Turno habitual"])+("*" if row["Turno variable"] else ""),axis=1)
            with st.container(border=True):
                _f1,_f2,_f3=st.columns([2,1,1])
                _search=_f1.text_input("Buscar persona",placeholder="Escribe un nombre…",key="attendance_search")
                _shift_filter=_f2.selectbox("Turno",["Todos"]+sorted(_att_df["Turno habitual"].dropna().unique().tolist(),key=lambda z:_shift_order.get(z,9)),key="attendance_shift_filter")
                _sort=_f3.selectbox("Ordenar por",["Más faltas","Más retardos","Más minutos acumulados","Nombre"],key="attendance_sort")
                _view=_att_df.copy()
                if _search.strip(): _view=_view[_view["Persona"].map(norm).str.contains(norm(_search),regex=False)]
                if _shift_filter!="Todos": _view=_view[_view["Turno habitual"]==_shift_filter]
                _sort_col={"Más faltas":"Faltas","Más retardos":"Retardos","Más minutos acumulados":"Minutos acumulados","Nombre":"Persona"}[_sort]
                _view=_view.sort_values(_sort_col,ascending=(_sort=="Nombre"))
                _show_cols=["Persona","Turno","Faltas","Retardos","Minutos acumulados","Personal en la app","Categoría en la app"]
                _attendance_color_rules=_percentile_severity_rules(_att_df,["Faltas","Retardos","Minutos acumulados"])
                st.dataframe(_style_severity(_view[_show_cols],_attendance_color_rules),use_container_width=True,hide_index=True)
                def _attendance_cutoff_label(label,column,unit=""):
                    _rule=_attendance_color_rules.get(column)
                    if not pd.to_numeric(_att_df.get(column,pd.Series(dtype=float)),errors="coerce").gt(0).any():
                        return f"{label}: sin casos positivos (0 verde)"
                    if not _rule: return f"{label}: sin casos positivos"
                    _yellow_max=_rule[1]-1
                    if _yellow_max<_rule[0]:
                        return f"{label}: 0 verde · ≥{_rule[1]}{unit} rojo"
                    return f"{label}: 0 verde · {_rule[0]}–{_yellow_max}{unit} amarillo · ≥{_rule[1]}{unit} rojo"
                st.caption("Semáforo del periodo (corte rojo = cuartil superior de las personas cargadas): "+" · ".join([
                    _attendance_cutoff_label("Faltas","Faltas"),
                    _attendance_cutoff_label("Retardos","Retardos"),
                    _attendance_cutoff_label("Minutos","Minutos acumulados"," min"),
                ]))
                st.caption(f"{len(_view):,} personas visibles · * indica registros de la persona en más de una categoría de turno.")

            with st.container(border=True):
                st.markdown("### 📝 Vincular con Seguimiento")
                _attendance_names=sorted(_att_df["Persona"].astype(str).unique().tolist(),key=lambda z:z.upper())
                _person_selected=st.selectbox("Persona",_attendance_names,key="attendance_followup_person")
                _person_row=_att_df[_att_df["Persona"]==_person_selected].iloc[0].to_dict()
                _person_token=token_key(_person_selected)
                _candidate_matches=app_people_by_name.get(_person_token,[])
                _saved_link=(store.get("attendance_links",{}) or {}).get(_person_token,"")
                if _saved_link:
                    _linked_name=_saved_link
                    _selected_candidate=next((z for _bucket in app_people_by_name.values() for z in _bucket if z["PICKER"]==_linked_name),{})
                    st.caption(f"Vínculo guardado con {_linked_name}.")
                elif len(_candidate_matches)>1:
                    _target_names=sorted({z["PICKER"] for z in _candidate_matches})
                    _linked_name=st.selectbox("Hay más de una coincidencia. Elige a quién vincular",["No asociar"]+_target_names,key=f"attendance_link_choice_{person_key(_person_selected)}")
                    _selected_candidate=next((z for z in _candidate_matches if z["PICKER"]==_linked_name),{})
                elif len(_candidate_matches)==1:
                    _linked_name=_candidate_matches[0]["PICKER"]
                    _selected_candidate=_candidate_matches[0]
                    st.caption(f"Coincidencia con el expediente de {_linked_name}.")
                else:
                    _manual_options=["Sin asociar"]+app_people_names
                    _manual_key=f"attendance_manual_link_{person_key(_person_selected)}"
                    _manual_choice=st.selectbox("No hay coincidencia exacta. Puedes vincular manualmente con la plantilla",_manual_options,key=_manual_key)
                    if st.button("Guardar vínculo de esta persona",key=f"attendance_save_link_{person_key(_person_selected)}"):
                        if _manual_choice=="Sin asociar":
                            (store.get("attendance_links",{}) or {}).pop(_person_token,None)
                            store["attendance_links"]=store.get("attendance_links",{}) or {}
                        else:
                            store.setdefault("attendance_links",{})[_person_token]=_manual_choice
                        save_store(store); st.success("Vínculo guardado."); st.rerun()
                    _linked_name=_manual_choice if _manual_choice!="Sin asociar" else _person_selected
                    _selected_candidate=next((z for _bucket in app_people_by_name.values() for z in _bucket if z["PICKER"]==_linked_name),{})
                    if _linked_name==_person_selected: st.caption("Sin coincidencia guardada: se registrará con este nombre como personal sin registrar.")
                if _linked_name!="No asociar":
                    _same_name_candidates=[
                        z for _bucket in app_people_by_name.values() for z in _bucket
                        if z.get("PICKER")==_linked_name and email_key(z.get("CORREO",""))
                    ]
                    _same_name_emails={email_key(z.get("CORREO","")) for z in _same_name_candidates}
                    _followup_candidate=(
                        _same_name_candidates[0] if len(_same_name_emails)==1 and _same_name_candidates else {}
                    )
                    _followup_email=_followup_candidate.get("CORREO","")
                    _default_supervisor=_selected_candidate.get("SUPERVISOR","No especificado")
                    _form_suffix=person_key(_linked_name) or "persona"
                    _saved_resource_options={}
                    for _resource_index,_resource in enumerate(store.get("recursos_formatos",[]) or []):
                        _resource_url=str(_resource.get("url","")).strip()
                        if _resource_url.startswith(("https://","http://")):
                            _resource_label=f"{_resource_index+1}. {_resource.get('tipo','Documento')} · {_resource.get('titulo','Enlace')}"
                            _saved_resource_options[_resource_label]={"titulo":str(_resource.get("titulo","Documento")),"url":_resource_url}
                    with st.form(f"attendance_followup_form_{_form_suffix}",clear_on_submit=True):
                        _ac1,_ac2=st.columns([1,2])
                        _att_category=_ac1.selectbox("Categoría del registro",FOLLOWUP_CATEGORIES,key=f"attendance_followup_category_{_form_suffix}")
                        _att_supervisor=_ac2.text_input("Supervisor",value=_default_supervisor,key=f"attendance_followup_supervisor_{_form_suffix}")
                        _att_detail=st.text_area("Motivo / detalle",value=f"Faltas: {int(_person_row['Faltas'])} · Retardos: {int(_person_row['Retardos'])} · Minutos acumulados de retardo: {int(_person_row['Minutos acumulados'])} · Periodo: {_period}",key=f"attendance_followup_detail_{_form_suffix}")
                        _att_feedback=st.text_area("Retroalimentación (opcional)",key=f"attendance_followup_feedback_{_form_suffix}")
                        st.caption("Puedes elegir enlaces guardados o pegar otros enlaces de documentos.")
                        _att_saved_links=st.multiselect("Enlaces guardados (opcional)",options=list(_saved_resource_options),placeholder="Selecciona uno o varios documentos",key=f"attendance_followup_links_{_form_suffix}")
                        _att_other_links=st.text_area("Otros enlaces de documentos (opcional)",placeholder="Pega una URL por línea",height=70,key=f"attendance_followup_other_links_{_form_suffix}")
                        if st.form_submit_button("Guardar en expediente de seguimiento",type="primary"):
                            _document_links=[]
                            for _selected_label in _att_saved_links:
                                _link=dict(_saved_resource_options[_selected_label])
                                if not any(_existing.get("url")==_link.get("url") for _existing in _document_links):
                                    _document_links.append(_link)
                            _invalid_links=[]
                            for _raw_link in re.split(r"[\n,;]+",_att_other_links):
                                _url=_raw_link.strip()
                                if not _url: continue
                                if not _url.startswith(("https://","http://")):
                                    _invalid_links.append(_url)
                                elif not any(_existing.get("url")==_url for _existing in _document_links):
                                    _document_links.append({"titulo":_url,"url":_url})
                            if _invalid_links:
                                st.error("Cada enlace debe comenzar con https:// o http://. Revisa los enlaces escritos.")
                            else:
                                _new_att_followup={
                                    "id":uuid.uuid4().hex,
                                    "fecha":datetime.now().strftime("%Y-%m-%d %H:%M"),
                                    "accion":_att_category,
                                    "supervisor":_att_supervisor.strip() or "No especificado",
                                    "motivo":_att_detail.strip(),
                                    "retroalimentacion":_att_feedback.strip(),
                                    "documentos_enlaces":_document_links
                                }
                                try:
                                    save_followup_entry(
                                        store,_linked_name,"acciones",_new_att_followup,
                                        aliases=[_person_selected],
                                        identity_email=_followup_email,
                                        identity_fields=_followup_candidate,
                                    )
                                except Exception as exc:
                                    st.error(str(exc))
                                else:
                                    st.success("Seguimiento respaldado en la nube.")
                                    st.rerun()
                    _linked_email=str(_followup_email)
                    if email_key(_linked_email):
                        _stored_record=followup_email_record(
                            store,_linked_name,_linked_email,
                            aliases=[_person_selected],migrate_legacy=False,create=False,
                        )
                    else:
                        _stored_record=None
                        for _stored_name,_stored_value in (store.get("pickers",{}) or {}).items():
                            if token_key(_stored_name) in {token_key(_linked_name),token_key(_person_selected)}:
                                _stored_record=_stored_value
                                break
                    if _stored_record and _stored_record.get("acciones"):
                        st.markdown("**Historial de seguimiento de esta persona**")
                        _history_df=pd.DataFrame(_stored_record["acciones"]).rename(columns={"fecha":"Fecha","accion":"Categoría","supervisor":"Supervisor","motivo":"Motivo","retroalimentacion":"Retroalimentación"})
                        _history_columns=[z for z in ["Fecha","Categoría","Supervisor","Motivo","Retroalimentación"] if z in _history_df.columns]
                        st.dataframe(_history_df.sort_values("Fecha",ascending=False)[_history_columns],use_container_width=True,hide_index=True)
                        for _action_index,_history_action in enumerate(reversed(_stored_record["acciones"])):
                            _history_links=_history_action.get("documentos_enlaces",[]) or []
                            if _history_links:
                                st.caption(f"Documentos vinculados · {_history_action.get('fecha','')}")
                                _link_columns=st.columns(min(len(_history_links),3))
                                for _link_index,_history_link in enumerate(_history_links):
                                    _link_url=str(_history_link.get("url","")).strip()
                                    if _link_url.startswith(("https://","http://")):
                                        _link_title=str(_history_link.get("titulo") or f"Documento {_link_index+1}")
                                        _link_columns[_link_index%len(_link_columns)].link_button(f"📎 {_link_title}",_link_url,key=f"attendance_followup_doc_{_form_suffix}_{_action_index}_{_link_index}",use_container_width=True)
                    else:
                        st.caption("Todavía no hay registros en el expediente de esta persona.")

if _tab_active(h):
    with h:
        st.subheader("🛡️ Seguimiento")
        _saved_flash=st.session_state.pop("seguimiento_saved_flash","")
        if _saved_flash:
            st.success(_saved_flash)
        st.caption("Todos los registros se organizan en tres categorías: 1. Seguimiento, 2. Acta y 3. 0 Tolerancia.")
        ctx,_,_,_=global_context(context_identity)
        selected_context=apply_person_filters(
            context_identity,"Todos",
            ctx.get("turno","Todos"),ctx.get("supervisor","Todos"),ctx.get("area","Todos"),
        )
        selected_context=selected_context[selected_context["CORREO_KEY"].map(email_key).ne("")].copy()
        _scope_parts=[f"{label}: {ctx[key]}" for key,label in (("turno","Turno"),("supervisor","Supervisor"),("area","Área")) if ctx.get(key,"Todos")!="Todos"]
        _scope=" · ".join(_scope_parts) if _scope_parts else "Todos los turnos, supervisores y áreas"
        st.caption(f"La lista se toma directamente del master de plantillas e incluye personal de Layout aunque no aparezca en los reportes operativos. · {_scope}")

        all_rows=[]
        for _,rr in selected_context.iterrows():
            picker=str(rr.get("PICKER","")); _ek=email_key(rr.get("CORREO_KEY",rr.get("CORREO","")))
            _unique_name=master_name_is_unique(context_identity,picker,_ek)
            rec0=followup_email_record(
                store,picker,_ek,aliases=[str(rr.get("_SOURCE_PICKER",""))],
                migrate_legacy=_unique_name,create=False,
            )
            acciones=rec0.get("acciones",[]) or []
            documentos=rec0.get("documentos",[]) or []
            retro_legacy=[]
            for z in store.get("feedback_rows",[]) or []:
                _z_email=email_key(z.get("CORREO_KEY",z.get("CORREO","")))
                if _z_email and _z_email==_ek:
                    retro_legacy.append(z)
                elif not _z_email and _unique_name and person_key(z.get("PICKER",""))==person_key(picker):
                    retro_legacy.append(z)
            retro_new=[z for z in acciones if str(z.get("retroalimentacion","")).strip()]
            categorias=[normalize_followup_category(z.get("accion","")) for z in acciones]
            categorias += [normalize_followup_category(z.get("tipo","")) for z in documentos]
            seguimiento_count=sum(c==FOLLOWUP_CATEGORIES[0] for c in categorias)
            acta_count=sum(c==FOLLOWUP_CATEGORIES[1] for c in categorias)
            tolerancia_count=sum(c==FOLLOWUP_CATEGORIES[2] for c in categorias)
            fechas=[str(z.get("fecha","")) for z in acciones+documentos if z.get("fecha")]
            all_rows.append({"PERSONA":picker,"TIPO DE PERSONAL":rr.get("PERSONAL_TIPO","Picker"),"CORREO":str(rr.get("CORREO","")),"TURNO":rr.get("TURNO",""),"SUPERVISOR":rr.get("SUPERVISOR",""),"AREA":rr.get("AREA_BASE",""),"RETROALIMENTACIONES":len(retro_legacy)+len(retro_new),"1. SEGUIMIENTO":seguimiento_count,"2. ACTA":acta_count,"3. 0 TOLERANCIA":tolerancia_count,"DOCUMENTOS":len(documentos),"TOTAL":len(acciones)+len(documentos)+len(retro_legacy),"ÚLTIMO SEGUIMIENTO":max(fechas,default="")})
        seguimiento_df=pd.DataFrame(all_rows)
        if seguimiento_df.empty: st.info("No hay pickers disponibles para seguimiento.")
        else:
            with st.container(border=True):
                fc1,fc2,fc3,fc4=st.columns(4)
                search_seg=fc1.text_input("Buscar picker",placeholder="Nombre o correo…",key="seguimiento_global_search")
                old_filter_map={"Con actas":"Con 2. Acta","Sin actas":"Sin 2. Acta","Con cero tolerancia":"Con 3. 0 Tolerancia","Sin seguimiento":"Sin 1. Seguimiento"}
                if st.session_state.get("seguimiento_global_estado") in old_filter_map:
                    st.session_state["seguimiento_global_estado"]=old_filter_map[st.session_state["seguimiento_global_estado"]]
                old_sort_map={"ACTAS":"2. ACTA","SEGUIMIENTOS":"1. SEGUIMIENTO","CERO TOLERANCIA":"3. 0 TOLERANCIA","LLAMADAS / ADVERTENCIAS":"1. SEGUIMIENTO"}
                if st.session_state.get("seguimiento_global_sort") in old_sort_map:
                    st.session_state["seguimiento_global_sort"]=old_sort_map[st.session_state["seguimiento_global_sort"]]
                seg_filter=fc2.selectbox("Estado",["Todos","Con 1. Seguimiento","Sin 1. Seguimiento","Con 2. Acta","Sin 2. Acta","Con 3. 0 Tolerancia","Sin 3. 0 Tolerancia"],key="seguimiento_global_estado")
                seg_sort=fc3.selectbox("Ordenar por",["1. SEGUIMIENTO","2. ACTA","3. 0 TOLERANCIA","RETROALIMENTACIONES","DOCUMENTOS","TOTAL"],key="seguimiento_global_sort")
                seg_dir=fc4.selectbox("Orden",["Mayor a menor","Menor a mayor"],key="seguimiento_global_dir")
                view=seguimiento_df.copy()
                if search_seg.strip():
                    _term_seg=norm(search_seg)
                    _name_hit=view["PERSONA"].map(norm).str.contains(_term_seg,regex=False)
                    _email_hit=view["CORREO"].map(norm).str.contains(_term_seg,regex=False)
                    view=view[_name_hit|_email_hit]
                if seg_filter=="Con 1. Seguimiento": view=view[view["1. SEGUIMIENTO"]>0]
                elif seg_filter=="Sin 1. Seguimiento": view=view[view["1. SEGUIMIENTO"]==0]
                elif seg_filter=="Con 2. Acta": view=view[view["2. ACTA"]>0]
                elif seg_filter=="Sin 2. Acta": view=view[view["2. ACTA"]==0]
                elif seg_filter=="Con 3. 0 Tolerancia": view=view[view["3. 0 TOLERANCIA"]>0]
                elif seg_filter=="Sin 3. 0 Tolerancia": view=view[view["3. 0 TOLERANCIA"]==0]
                view=view.sort_values(seg_sort,ascending=(seg_dir=="Menor a mayor"))
                st.dataframe(view,use_container_width=True,hide_index=True)
            st.caption(f"{len(view):,} personas del master visibles.")

        st.divider()
        _master_options={}
        for _,_master_row in selected_context.iterrows():
            _master_email=email_key(_master_row.get("CORREO_KEY",_master_row.get("CORREO","")))
            if _master_email:
                _master_options[_master_email]=_master_row.to_dict()
        _master_labels={
            _email:f"{_row.get('PICKER','')} · {_row.get('PERSONAL_TIPO','Picker')} · {_email}"
            for _email,_row in _master_options.items()
        }
        with st.container(border=True):
            st.markdown("**🔎 Abrir expediente individual**")
            buscar=st.text_input("Buscar en la plantilla",placeholder="Nombre o correo…",key="seguimiento_picker_search")
            opciones=sorted(_master_options)
            if buscar.strip():
                term=norm(buscar); opciones=[k for k in opciones if term in norm(_master_labels[k])]
            _selected_email=st.selectbox(
                "Picker · correo de la plantilla",[""]+opciones,
                format_func=lambda k:"Selecciona un picker" if not k else _master_labels.get(k,k),
                key="seguimiento_picker_email_select_v2",
            )

        if not _selected_email:
            st.info("Selecciona un picker para consultar y documentar su expediente.")
        else:
            r=pd.Series(_master_options[_selected_email])
            seguimiento_picker=str(r.get("PICKER","")).strip()
            seguimiento_email=str(r.get("CORREO",_selected_email)).strip()
            _seguimiento_id=hashlib.sha256(_selected_email.encode("utf-8")).hexdigest()[:16]
            _seguimiento_identity={
                "PICKER":seguimiento_picker,"CORREO":seguimiento_email,
                "CORREO_KEY":_selected_email,"TURNO":str(r.get("TURNO","")),
                "SUPERVISOR":str(r.get("SUPERVISOR","")),
                "AREA_BASE":str(r.get("AREA_BASE","")),
                "PERSONAL_TIPO":str(r.get("PERSONAL_TIPO","Picker")),
            }
            _legacy_name_unique=master_name_is_unique(context_identity,seguimiento_picker,_selected_email)
            rec=followup_email_record(
                store,seguimiento_picker,seguimiento_email,
                aliases=[str(r.get("_SOURCE_PICKER",""))],
                migrate_legacy=_legacy_name_unique,create=False,
            )
            acciones=rec.get("acciones",[]) or []; documentos=rec.get("documentos",[]) or []
            st.markdown(f"<div class='justo-card'><div class='justo-kicker'>Expediente · referencia del master</div><div class='justo-title'>{_h(seguimiento_picker)}</div><div class='justo-muted'>Tipo: {_h(r.get('PERSONAL_TIPO','Picker'))} · Correo: {_h(seguimiento_email)} · Turno: {_h(r.get('TURNO',''))} · Supervisor: {_h(r.get('SUPERVISOR',''))} · Área: {_h(r.get('AREA_BASE',''))}</div></div>",unsafe_allow_html=True)
            k1,k2,k3,k4,k5=st.columns(5)
            _history_rows=[]
            for _item in acciones:
                _history_rows.append({
                    "Fecha":str(_item.get("fecha","")),
                    "Categoría":normalize_followup_category(_item.get("accion","")),
                    "Tipo de registro":"Seguimiento",
                    "Supervisor":str(_item.get("supervisor","")),
                    "Detalle":str(_item.get("motivo","")),
                    "Retroalimentación":str(_item.get("retroalimentacion","")),
                })
            for _item in documentos:
                _history_rows.append({
                    "Fecha":str(_item.get("fecha","")),
                    "Categoría":normalize_followup_category(_item.get("tipo","")),
                    "Tipo de registro":"Documento / acta",
                    "Supervisor":str(_item.get("supervisor",r.get("SUPERVISOR",""))),
                    "Detalle":" · ".join(v for v in [str(_item.get("titulo","")).strip(),str(_item.get("detalle","")).strip()] if v),
                    "Retroalimentación":"",
                })
            _history_rows.sort(key=lambda z:z["Fecha"],reverse=True)
            categorias_picker=[normalize_followup_category(z.get("accion","")) for z in acciones]
            categorias_picker += [normalize_followup_category(z.get("tipo","")) for z in documentos]
            k1.metric("Retroalimentaciones",f"{sum(1 for z in acciones if str(z.get('retroalimentacion','')).strip()):,}")
            k2.metric("1. Seguimiento",f"{sum(c==FOLLOWUP_CATEGORIES[0] for c in categorias_picker):,}")
            k3.metric("2. Acta",f"{sum(c==FOLLOWUP_CATEGORIES[1] for c in categorias_picker):,}")
            k4.metric("3. 0 Tolerancia",f"{sum(c==FOLLOWUP_CATEGORIES[2] for c in categorias_picker):,}")
            k5.metric("Documentos",f"{len(documentos):,}")
            _attendance_matches=attendance_by_name.get(token_key(seguimiento_picker),[]) if _legacy_name_unique else []
            if len(_attendance_matches)==1:
                _att_row=_attendance_matches[0]
                st.markdown("### 📅 Asistencia del periodo cargado")
                _ak1,_ak2,_ak3,_ak4=st.columns(4)
                _ak1.metric("Faltas",f"{int(_att_row.get('Faltas',0)):,}")
                _ak2.metric("Retardos",f"{int(_att_row.get('Retardos',0)):,}")
                _am=int(_att_row.get("Minutos acumulados",0)); _ah,_amin=divmod(_am,60)
                _ak3.metric("Tiempo de retardo",f"{_ah} h {_amin:02d} min")
                _ak4.metric("Turno de referencia",str(_att_row.get("Turno habitual","Sin dato"))+(" · variable" if _att_row.get("Turno variable") else ""))
                st.caption(f"Datos del periodo {attendance_summary.get('fecha_desde','')} a {attendance_summary.get('fecha_hasta','')}. El turno proviene del reporte de retardos.")
            with st.container(border=True):
                st.markdown("### 📋 Historial de seguimiento")
                if _history_rows:
                    st.dataframe(pd.DataFrame(_history_rows),use_container_width=True,hide_index=True)
                else: st.info("Este picker todavía no tiene seguimientos registrados.")
            with st.container(border=True):
                st.markdown("### 📄 Registrar seguimiento")
                st.caption("Selecciona una categoría y registra el motivo. El PDF y el enlace son opcionales.")
                recursos=store.get("recursos_formatos",[]) or []
                recurso_lookup={}
                for i,z in enumerate(recursos):
                    if str(z.get("url","")).strip():
                        recurso_lookup[f"{i+1}. {z.get('tipo','Recurso')} · {z.get('titulo','Enlace')}"]=z
                recurso_opciones=["Sin enlace guardado"]+list(recurso_lookup)
                with st.container(border=True):
                    st.markdown("**🔗 Elegir enlace guardado (opcional)**")
                    recurso_seleccionado=st.selectbox("Enlaces de seguimiento",recurso_opciones,key=f"seguimiento_enlace_guardado_{_seguimiento_id}")
                    recurso_actual=recurso_lookup.get(recurso_seleccionado)
                    if recurso_actual:
                        st.info(f"Seleccionado: {recurso_actual.get('tipo','Seguimiento')} · {recurso_actual.get('titulo','Enlace de seguimiento')}")
                        st.link_button("Abrir enlace para revisar",str(recurso_actual.get("url","")))
                enviar_copia=st.checkbox("Enviar copia por correo al guardar",value=False,key=f"seguimiento_email_{_seguimiento_id}")
                destinatarios=st.text_input("Correo(s) destinatario(s)",placeholder="persona@correo.com (separa varios con coma)",key=f"seguimiento_destinatarios_{_seguimiento_id}") if enviar_copia else ""
                with st.form(f"seguimiento_documento_form_{_seguimiento_id}",clear_on_submit=True):
                    dc1,dc2=st.columns([1,2])
                    with dc1: doc_tipo=st.selectbox("Categoría",FOLLOWUP_CATEGORIES)
                    with dc2: doc_titulo=st.text_input("Nombre / referencia",placeholder="Ej. Acta por FNR — septiembre 2026")
                    doc_detalle=st.text_area("Detalle / motivo",placeholder="Qué originó el seguimiento y cualquier dato importante…")
                    doc_pdf=st.file_uploader("📎 Adjuntar PDF",type=["pdf"],accept_multiple_files=False,key=f"seguimiento_pdf_{_seguimiento_id}")
                    if st.form_submit_button("💾 Guardar seguimiento / documento",type="primary"):
                        url=str(recurso_actual.get("url","")).strip() if recurso_actual else ""
                        titulo=doc_titulo.strip() or (doc_pdf.name if doc_pdf is not None else (str(recurso_actual.get("titulo",doc_tipo)) if recurso_actual else doc_tipo))
                        registro={
                            "id":uuid.uuid4().hex,
                            "fecha":datetime.now().strftime("%Y-%m-%d %H:%M"),
                            "tipo":doc_tipo,"titulo":titulo,
                            "detalle":doc_detalle.strip(),
                            "supervisor":str(r.get("SUPERVISOR","")),
                            "url":url,
                            "url_titulo":str(recurso_actual.get("titulo","")) if recurso_actual else "",
                            "path":"","archivo":"","destinatarios":destinatarios,
                        }
                        pdf_bytes=None
                        try:
                            if not cloud_enabled():
                                raise RuntimeError("No se guardó: el respaldo en nube está desactivado.")
                            if doc_pdf is not None:
                                pdf_bytes=doc_pdf.getvalue()
                                path,_=save_followup_pdf(pdf_bytes,seguimiento_picker,doc_pdf.name,identity_email=seguimiento_email)
                                registro["path"]=path
                                registro["archivo"]=_safe_filename(doc_pdf.name)
                            if enviar_copia:
                                registro["email_estado"]="Pendiente de envío"
                            save_followup_entry(
                                store,seguimiento_picker,"documentos",registro,
                                aliases=[str(r.get("_SOURCE_PICKER",""))],
                                identity_email=seguimiento_email,
                                identity_fields=_seguimiento_identity,
                            )
                        except Exception as exc:
                            st.error(str(exc))
                        else:
                            st.success("Documento / seguimiento respaldado en la nube.")
                            if enviar_copia:
                                email_body=f"Se registró un {doc_tipo} para {seguimiento_picker}.\\n\\nDetalle: {doc_detalle.strip() or 'Sin detalle.'}"
                                if url:
                                    email_body+=f"\\n\\nEnlace de seguimiento: {url}"
                                ok,msg=send_followup_email(
                                    f"Seguimiento {doc_tipo} · {seguimiento_picker}",
                                    email_body,destinatarios,pdf_bytes,
                                    registro.get("archivo") or "seguimiento.pdf"
                                )
                                _saved_rec=followup_email_record(
                                    store,seguimiento_picker,seguimiento_email,
                                    migrate_legacy=False,create=True,
                                )
                                _saved_doc=next(
                                    (d for d in _saved_rec.get("documentos",[])
                                     if d.get("id")==registro["id"]),None
                                )
                                if _saved_doc is not None:
                                    _saved_doc["email_estado"]="Enviado" if ok else "Pendiente: error de envío"
                                    _saved_doc["email_ultimo_envio"]=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    _saved_doc["email_mensaje"]=msg
                                    try:
                                        save_store(store)
                                    except Exception as exc:
                                        st.warning("Se respaldó el documento, pero no pudo actualizarse el resultado del correo: "+str(exc))
                                if ok: st.session_state["seguimiento_saved_flash"]=f"{doc_tipo} guardado. {msg}"
                                else: st.session_state["seguimiento_saved_flash"]=f"{doc_tipo} guardado en la nube. El correo quedó pendiente."
                            else:
                                st.session_state["seguimiento_saved_flash"]=f"{doc_tipo} guardado en la nube."
                            st.rerun()
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
                            if st.button("Quitar",key=f"seg_recurso_del_{idx}",use_container_width=True):
                                st.session_state["confirm_delete_resource_index"]=idx
                                st.rerun()
                        if st.session_state.get("confirm_delete_resource_index")==idx:
                            st.warning(f"¿Quitar el enlace ‘{recurso.get('titulo','Formato')}’ de la lista reutilizable?")
                            del_yes,del_no=st.columns(2)
                            with del_yes:
                                if st.button("Sí, quitar enlace",key=f"seg_recurso_confirm_{idx}",type="primary"):
                                    if idx < len(store.get("recursos_formatos",[])):
                                        store["recursos_formatos"].pop(idx)
                                    st.session_state.pop("confirm_delete_resource_index",None)
                                    save_store(store); st.rerun()
                            with del_no:
                                if st.button("Cancelar",key=f"seg_recurso_cancel_{idx}"):
                                    st.session_state.pop("confirm_delete_resource_index",None)
                                    st.rerun()
                with st.form("seg_recurso_form",clear_on_submit=True):
                    q1,q2=st.columns([1,2])
                    with q1: rt=st.selectbox("Categoría",FOLLOWUP_CATEGORIES)
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
                    if doc.get("email_estado"):
                        st.caption(f"Correo: {doc.get('email_estado')} · {doc.get('email_ultimo_envio','')}")
                    with st.expander("✉️ Reenviar este documento por correo",expanded=False):
                        resend_to=st.text_input(
                            "Destinatario(s)",
                            value=str(doc.get("destinatarios", "")),
                            placeholder="persona@correo.com (separa varios con coma)",
                            key=f"resend_to_{_seguimiento_id}_{doc.get('id',i)}",
                        )
                        if st.button("Reenviar correo",key=f"resend_email_{_seguimiento_id}_{doc.get('id',i)}",type="primary"):
                            email_body=f"Se comparte el documento {doc.get('tipo','Seguimiento')} de {seguimiento_picker}.\n\n{doc.get('detalle') or 'Sin detalle.'}"
                            if doc.get("url"):
                                email_body+=f"\n\nEnlace de seguimiento: {doc.get('url')}"
                            ok,msg=send_followup_email(
                                f"Seguimiento {doc.get('tipo','Documento')} · {seguimiento_picker}",
                                email_body,
                                resend_to,
                                archivo,
                                str(doc.get("archivo") or "seguimiento.pdf"),
                            )
                            doc["destinatarios"]=resend_to.strip()
                            doc["email_estado"]="Enviado" if ok else "Pendiente: error de envío"
                            doc["email_ultimo_envio"]=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            doc["email_mensaje"]=msg
                            save_store(store)
                            if ok: st.success(msg)
                            else: st.error(msg)
            else: st.info("No hay documentos vinculados a este expediente.")

if _tab_active(x):
    with x:
        st.divider()
        st.markdown("### 📤 Exportar")
        st.download_button("📥 Descargar Excel completo",export(s,fnr,mc,None if sp=="Todos" else sp,roster),f"Analisis_FNR_MC_{periodo}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.info("La app utiliza 3 Excel operativos separados (Pickers, FNR y MC) y un Master consolidado para turno, correo, supervisor y área.")
        st.success(f"Persistencia activa: {len(store.get('excluded_orders',[]))} pedidos excluidos · {len(store.get('master_overrides',{}))} asignaciones manuales · {len(store.get('master_excluded',[]))} exclusiones de personal · {len(store.get('feedback_rows',[]))} retroalimentaciones guardadas.")

        st.divider()
        st.markdown("### 💾 Respaldo")
        st.caption("Descarga una copia de asignaciones, seguimientos y procesos para conservarla como respaldo adicional.")
        backup_json=json.dumps(store,ensure_ascii=False,indent=2).encode("utf-8")
        st.download_button("Descargar respaldo de configuración",backup_json,f"Respaldo_Control_FNR_{periodo}.json","application/json",key="tools_config_backup")

if _tab_active(j):
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
                        st.markdown(f"<div class='justo-card'><div class='justo-kicker'>{_h(priority)}</div><div class='justo-title'>{_h(proc.get('titulo','Sin título'))}</div><div class='justo-muted'>Responsable: {_h(proc.get('responsable') or 'Sin asignar')} · Fecha: {_h(proc.get('fecha_objetivo') or 'Sin fecha')}</div></div>",unsafe_allow_html=True)
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
                                st.session_state["confirm_delete_process_id"]=proc.get("id")
                                st.rerun()
                        if st.session_state.get("confirm_delete_process_id")==proc.get("id"):
                            st.warning(f"¿Eliminar el proceso ‘{proc.get('titulo','Sin título')}’?")
                            yes_col,no_col=st.columns(2)
                            with yes_col:
                                if st.button("Sí, eliminar",key=f"proc_del_confirm_{proc['id']}",type="primary"):
                                    store["procesos"]=[pp for pp in store["procesos"] if pp.get("id")!=proc.get("id")]
                                    st.session_state.pop("confirm_delete_process_id",None)
                                    save_store(store); st.rerun()
                            with no_col:
                                if st.button("Cancelar",key=f"proc_del_cancel_{proc['id']}"):
                                    st.session_state.pop("confirm_delete_process_id",None)
                                    st.rerun()
