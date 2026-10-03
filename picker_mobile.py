"""Miniaplicación Streamlit móvil — publicar como app separada picker_mobile.py."""
import time
import uuid

import streamlit as st

from mobile_audit_backend import (
    MobileAuditDB, MobileAuditError, domain_allowed, verified_identity,
)


st.set_page_config(page_title="Validación Coyoacán",page_icon="📱",layout="centered")
st.title("📱 Validación Coyoacán")
st.caption("Mis pedidos · Autoverificación y auditoría cruzada")

if "auth" not in st.secrets:
    st.error("Inicio de sesión desactivado. Un administrador debe configurar [auth] antes de usar esta aplicación.")
    st.stop()

if not st.user.is_logged_in:
    st.info("Inicia sesión con tu correo individual para consultar únicamente los pedidos que te asignaron.")
    if st.button("Iniciar sesión con mi correo",type="primary",use_container_width=True):
        st.login()
    st.stop()

user = dict(st.user)
user["is_logged_in"] = st.user.is_logged_in
expiration = user.get("exp")
try:
    expired = bool(expiration and int(expiration) <= int(time.time()))
except (ValueError,TypeError):
    expired = True
if expired:
    st.warning("La sesión caducó. Inicia sesión nuevamente.")
    st.logout()
    st.stop()

email = verified_identity(user)
if not email or not domain_allowed(email,st.secrets.get("MOBILE_AUDIT_ALLOWED_DOMAINS","")):
    st.error("Este correo no está verificado o no tiene autorización para entrar.")
    st.button("Cerrar sesión",on_click=st.logout)
    st.stop()

with st.sidebar:
    st.caption(f"Usuario: {email}")
    st.button("Cerrar sesión",on_click=st.logout)

try:
    db = MobileAuditDB(
        st.secrets.get("SUPABASE_URL",""),
        st.secrets.get("SUPABASE_SECRET_KEY") or st.secrets.get("SUPABASE_SERVICE_ROLE_KEY",""),
    )
    assignments = db.mine(email)
except MobileAuditError as exc:
    st.error(str(exc))
    st.info("El administrador debe ejecutar supabase_mobile_audit.sql en el proyecto de Supabase y activar las credenciales.")
    st.stop()

pending = [a for a in assignments if a.get("status")=="pendiente"]
finished = [a for a in assignments if a.get("status")=="validado"]
q1,q2 = st.columns(2)
q1.metric("Pendientes",len(pending))
q2.metric("Validados",len(finished))

mode=st.segmented_control(
    "Tipo de revisión",
    options=["autoverificacion","auditoria_cruzada"],
    format_func=lambda v: "Mis pedidos" if v=="autoverificacion" else "Auditoría cruzada",
    default="autoverificacion",
    selection_mode="single",
    key="mobile_audit_view_mode",
)
mode=mode or "autoverificacion"
visible = [a for a in pending if a.get("mode")==mode]
st.caption(
    "Autoverificación: revisas tu preparación."
    if mode=="autoverificacion" else
    "Auditoría cruzada: revisas pedidos preparados por otra persona."
)

if not visible:
    st.success("No tienes pedidos pendientes de este tipo.")
else:
    selected_id = st.selectbox(
        "Pedido por revisar",
        [a["id"] for a in visible],
        format_func=lambda a_id: next(
            (
                f"#{a['mobile_audit_orders']['pedido']} · {a['mobile_audit_orders'].get('slot','')} · "
                f"{'Sin terminar' if a.get('last_result')=='No se pudo validar' else 'Pendiente'}"
                for a in visible if a["id"]==a_id
            ),
            str(a_id),
        ),
    )
    assignment = next(a for a in visible if a["id"]==selected_id)
    order = assignment.get("mobile_audit_orders") or {}
    st.subheader("Pedido #"+str(order.get("pedido","")))
    st.caption(
        f"Slot: {order.get('slot','Sin dato')} · Preparado por: "
        f"{order.get('picker_labels','No identificado')}"
    )

    all_items=order.get("items") or []
    items=(
        [v for v in all_items if v.get("picker_email")==email]
        if mode=="autoverificacion" else all_items
    )
    if not items:
        st.error("No hay artículos vinculados a tu correo para esta revisión. Contacta al supervisor.")
        st.stop()

    st.markdown("**Lista de artículos para verificar**")
    st.dataframe(
        [{
            "SKU":row.get("sku",""),
            "Artículo":row.get("articulo",""),
            "Pedido":row.get("cantidad_pedida",""),
            "Pickeado":row.get("cantidad_pickeada",""),
        } for row in items],
        hide_index=True,
        use_container_width=True,
    )
    st.caption(f"{len(items)} artículos por revisar.")

    result=st.radio(
        "Resultado de la revisión",
        ["Pedido correcto","Con diferencias","No se pudo validar"],
        key="mobile_result_"+selected_id,
    )
    options=list(range(len(items)))
    differences=[]
    if result=="Con diferencias":
        marked=st.multiselect(
            "¿Qué artículos presentan una diferencia?",
            options,
            format_func=lambda i: f"{i+1}. {items[i].get('sku','')} · {items[i].get('articulo','')}",
            key="mobile_diff_"+selected_id,
        )
        differences=[{
            "sku":items[idx].get("sku",""),
            "articulo":items[idx].get("articulo",""),
            "cantidad_pedida":items[idx].get("cantidad_pedida",""),
            "cantidad_pickeada":items[idx].get("cantidad_pickeada",""),
            "indice":idx,
        } for idx in marked]

    with st.form("mobile_review_"+selected_id,clear_on_submit=False):
        notes=st.text_area(
            "Describe las diferencias" if result=="Con diferencias" else
            "Motivo por el que no se pudo validar" if result=="No se pudo validar" else
            "Observaciones (opcional)",
            max_chars=1500,
        )
        if result=="No se pudo validar":
            st.caption("Este intento se guardará, pero el pedido seguirá pendiente.")
            confirmed=False
        else:
            confirmed=st.checkbox("Confirmo que revisé físicamente todos los artículos mostrados.")
        submitted=st.form_submit_button("Guardar revisión",type="primary",use_container_width=True)

    if submitted:
        if result!="No se pudo validar" and not confirmed:
            st.error("Confirma que revisaste todos los artículos antes de validar.")
        elif result=="No se pudo validar" and not notes.strip():
            st.error("Describe el motivo para poder registrar el intento.")
        elif result=="Con diferencias" and (not differences or not notes.strip()):
            st.error("Selecciona los artículos con diferencia y describe lo encontrado.")
        else:
            nonce_key="mobile_submit_nonce_"+selected_id
            nonce=st.session_state.setdefault(nonce_key,str(uuid.uuid4()))
            try:
                db.submit(
                    selected_id,email,result,notes,differences,
                    len(items) if confirmed else 0,nonce,
                )
            except MobileAuditError as exc:
                st.error(str(exc))
            else:
                st.session_state.pop(nonce_key,None)
                st.toast("Revisión registrada correctamente",icon="✅")
                st.rerun()

with st.expander(f"📚 Mis revisiones terminadas ({len(finished)})",expanded=False):
    if finished:
        st.dataframe(
            [{
                "Pedido":(a.get("mobile_audit_orders") or {}).get("pedido",""),
                "Tipo":"Propio" if a.get("mode")=="autoverificacion" else "Cruzada",
                "Resultado":a.get("last_result",""),
            } for a in finished],
            hide_index=True,use_container_width=True,
        )
    else:
        st.caption("Aún no hay revisiones terminadas.")
