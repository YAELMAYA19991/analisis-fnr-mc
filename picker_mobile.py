"""Miniaplicación móvil de auditoría para supervisores; acceso por código, sin correo."""
import time
import uuid

import streamlit as st

from mobile_audit_backend import (
    MobileAuditDB, MobileAuditError, supervisor_pin_valid,
)

st.set_page_config(page_title="Validación Coyoacán", page_icon="📱", layout="centered")
st.title("📱 Validación Coyoacán")
st.caption("Auditoría rápida desde el teléfono · acceso de supervisor")

configured_pin = str(st.secrets.get("MOBILE_AUDIT_SUPERVISOR_PIN", "")).strip()
if len(configured_pin) < 8:
    st.error("Falta configurar un código privado de supervisor (mínimo 8 caracteres) en Secrets de Streamlit.")
    st.info("Agrega MOBILE_AUDIT_SUPERVISOR_PIN en Settings → Secrets. No lo publiques en GitHub ni lo compartas en el chat.")
    st.stop()

# Mantiene abierta la sesión hasta 12 horas; el código nunca se guarda en la sesión.
authenticated = bool(st.session_state.get("mobile_supervisor_authenticated", False))
auth_time = float(st.session_state.get("mobile_supervisor_auth_time", 0) or 0)
if authenticated and time.time() - auth_time > 12 * 60 * 60:
    st.session_state.pop("mobile_supervisor_authenticated", None)
    st.session_state.pop("mobile_supervisor_name", None)
    st.session_state.pop("mobile_supervisor_auth_time", None)
    authenticated = False

if not authenticated:
    st.info("Ingresa tu nombre y el código privado de supervisor para abrir la cola de auditorías.")
    attempts = int(st.session_state.get("mobile_supervisor_pin_attempts", 0))
    if attempts >= 5:
        st.error("Se alcanzó el límite de intentos en esta sesión. Cierra y vuelve a abrir la app para intentarlo de nuevo.")
        st.stop()
    with st.form("mobile_supervisor_login"):
        supervisor_name = st.text_input("Nombre del supervisor", max_chars=100)
        entered_pin = st.text_input("Código de acceso", type="password", max_chars=100)
        login_submitted = st.form_submit_button("Entrar a auditorías", type="primary", use_container_width=True)
    if login_submitted:
        if supervisor_name.strip() and supervisor_pin_valid(entered_pin, configured_pin):
            st.session_state["mobile_supervisor_authenticated"] = True
            st.session_state["mobile_supervisor_name"] = supervisor_name.strip()
            st.session_state["mobile_supervisor_auth_time"] = time.time()
            st.session_state["mobile_supervisor_pin_attempts"] = 0
            st.rerun()
        else:
            st.session_state["mobile_supervisor_pin_attempts"] = attempts + 1
            st.error("Revisa el nombre y el código e intenta de nuevo.")
    st.stop()

supervisor_name = str(st.session_state.get("mobile_supervisor_name", "")).strip()
with st.sidebar:
    st.caption(f"Supervisor: {supervisor_name}")
    if st.button("Cerrar sesión", use_container_width=True):
        for key in (
            "mobile_supervisor_authenticated",
            "mobile_supervisor_name",
            "mobile_supervisor_auth_time",
            "mobile_supervisor_pin_attempts",
        ):
            st.session_state.pop(key, None)
        st.rerun()

try:
    db = MobileAuditDB(
        st.secrets.get("SUPABASE_URL", ""),
        st.secrets.get("SUPABASE_SECRET_KEY") or st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", ""),
    )
    assignments = db.all_assignments()
except MobileAuditError as exc:
    st.error(str(exc))
    st.info("Revisa las credenciales privadas de Supabase y que exista la migración de auditoría móvil.")
    st.stop()

pending = [a for a in assignments if a.get("status") == "pendiente"]
finished = [a for a in assignments if a.get("status") == "validado"]
q1, q2 = st.columns(2)
q1.metric("Pendientes", len(pending))
q2.metric("Validados", len(finished))

mode = st.segmented_control(
    "Tipo de revisión",
    options=["autoverificacion", "auditoria_cruzada"],
    format_func=lambda value: "Autoverificaciones" if value == "autoverificacion" else "Auditorías cruzadas",
    default="autoverificacion",
    selection_mode="single",
    key="mobile_audit_view_mode",
)
mode = mode or "autoverificacion"
visible = [a for a in pending if a.get("mode") == mode]
st.caption(
    "Revisa los artículos asociados al picker del pedido."
    if mode == "autoverificacion"
    else "Revisa el pedido completo como auditoría cruzada."
)

if not visible:
    st.success("No hay pedidos pendientes de este tipo.")
else:
    def order_label(assignment_id):
        assignment = next((row for row in visible if row.get("id") == assignment_id), {})
        order = assignment.get("mobile_audit_orders") or {}
        return (
            f"Pedido #{order.get('pedido', '')} · {order.get('slot', '')} · "
            f"{order.get('picker_labels', 'Picker no identificado')}"
        )

    selected_id = st.selectbox(
        "Pedido por revisar",
        [a["id"] for a in visible],
        format_func=order_label,
    )
    assignment = next(a for a in visible if a["id"] == selected_id)
    order = assignment.get("mobile_audit_orders") or {}
    st.subheader("Pedido #" + str(order.get("pedido", "")))
    st.caption(
        f"Slot: {order.get('slot', 'Sin dato')} · Preparado por: "
        f"{order.get('picker_labels', 'No identificado')}"
    )

    all_items = order.get("items") or []
    items = (
        [item for item in all_items if item.get("picker_email") == assignment.get("assignee_email")]
        if mode == "autoverificacion"
        else all_items
    )
    if not items:
        st.error("No hay artículos vinculados a esta asignación. Revisa el pedido antes de continuar.")
        st.stop()

    st.markdown("**Artículos para verificar**")
    st.dataframe(
        [{
            "SKU": item.get("sku", ""),
            "Artículo": item.get("articulo", ""),
            "Pedido": item.get("cantidad_pedida", ""),
            "Pickeado": item.get("cantidad_pickeada", ""),
        } for item in items],
        hide_index=True,
        use_container_width=True,
    )
    st.caption(f"{len(items)} artículos por revisar.")

    result = st.radio(
        "Resultado de la revisión",
        ["Pedido correcto", "Con diferencias", "No se pudo validar"],
        key="mobile_result_" + selected_id,
    )
    options = list(range(len(items)))
    differences = []
    if result == "Con diferencias":
        marked = st.multiselect(
            "¿Qué artículos presentan una diferencia?",
            options,
            format_func=lambda index: f"{index + 1}. {items[index].get('sku', '')} · {items[index].get('articulo', '')}",
            key="mobile_diff_" + selected_id,
        )
        differences = [{
            "sku": items[index].get("sku", ""),
            "articulo": items[index].get("articulo", ""),
            "cantidad_pedida": items[index].get("cantidad_pedida", ""),
            "cantidad_pickeada": items[index].get("cantidad_pickeada", ""),
            "indice": index,
        } for index in marked]

    with st.form("mobile_review_" + selected_id, clear_on_submit=False):
        notes = st.text_area(
            "Describe las diferencias" if result == "Con diferencias" else
            "Motivo por el que no se pudo validar" if result == "No se pudo validar" else
            "Observaciones (opcional)",
            max_chars=1500,
        )
        if result == "No se pudo validar":
            st.caption("Este intento se guardará, pero el pedido seguirá pendiente.")
            confirmed = False
        else:
            confirmed = st.checkbox("Confirmo que revisé físicamente todos los artículos mostrados.")
        submitted = st.form_submit_button("Guardar revisión", type="primary", use_container_width=True)

    if submitted:
        if result != "No se pudo validar" and not confirmed:
            st.error("Confirma que revisaste todos los artículos antes de validar.")
        elif result == "No se pudo validar" and not notes.strip():
            st.error("Describe el motivo para poder registrar el intento.")
        elif result == "Con diferencias" and (not differences or not notes.strip()):
            st.error("Selecciona los artículos con diferencia y describe lo encontrado.")
        else:
            nonce_key = "mobile_submit_nonce_" + selected_id
            nonce = st.session_state.setdefault(nonce_key, str(uuid.uuid4()))
            try:
                db.submit_supervisor(
                    selected_id, supervisor_name, result, notes, differences,
                    len(items) if confirmed else 0, nonce,
                )
            except MobileAuditError as exc:
                st.error(str(exc))
            else:
                st.session_state.pop(nonce_key, None)
                st.toast("Revisión registrada correctamente", icon="✅")
                st.rerun()

with st.expander(f"📚 Revisiones terminadas ({len(finished)})", expanded=False):
    if finished:
        st.dataframe(
            [{
                "Pedido": (row.get("mobile_audit_orders") or {}).get("pedido", ""),
                "Tipo": "Autoverificación" if row.get("mode") == "autoverificacion" else "Cruzada",
                "Resultado": row.get("last_result", ""),
            } for row in finished],
            hide_index=True,
            use_container_width=True,
        )
    else:
        st.caption("Aún no hay revisiones terminadas.")
