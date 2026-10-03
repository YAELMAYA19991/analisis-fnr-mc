"""Panel protegido de asignaciones móviles incrustado en Auditoría de pedidos."""
import hashlib

import streamlit as st

from mobile_audit_backend import (
    MobileAuditDB, MobileAuditError, admin_allowed, available_people, domain_allowed,
    email_from_source, normalized_name, order_id, order_items, owners_for_lines,
    validated_assignment, verified_identity,
)


def _staff_emails(people, allowed_domains):
    return {
        email: p for email, p in people.items()
        if domain_allowed(email, allowed_domains)
    }


def _items_with_owner_email(lines, people):
    result = order_items(lines)
    for item in result:
        name = item["picker"]
        direct = email_from_source(name)
        matches = (
            [direct] if direct and direct in people else
            [mail for mail, person in people.items() if normalized_name(name) in person["names"]]
        )
        item["picker_email"] = matches[0] if len(matches) == 1 else ""
    return result


def render_mobile_assignment(pedido, slot, row, lines, roster, summary, raw_upload):
    """Nadie sin OIDC + lista explícita de administradores puede asignar ni ver revisiones."""
    st.caption("Asigna autoverificaciones y auditorías cruzadas sin modificar el Excel original.")
    if "auth" not in st.secrets:
        st.info("Para habilitar las asignaciones, configura el inicio de sesión [auth] de esta aplicación.")
        return
    if not st.secrets.get("MOBILE_AUDIT_ADMIN_EMAILS"):
        st.info("Configura MOBILE_AUDIT_ADMIN_EMAILS con los correos autorizados para asignar pedidos.")
        return
    if not st.user.is_logged_in:
        if st.button("Iniciar sesión como supervisor",key="mobile_manager_signin"):
            st.login()
        return

    user=dict(st.user)
    user["is_logged_in"]=st.user.is_logged_in
    supervisor=verified_identity(user)
    if not admin_allowed(supervisor,st.secrets.get("MOBILE_AUDIT_ADMIN_EMAILS","")):
        st.error("Esta cuenta no está autorizada para asignar pedidos.")
        return

    people = _staff_emails(
        available_people(roster,summary),
        st.secrets.get("MOBILE_AUDIT_ALLOWED_DOMAINS",""),
    )
    owners,unknown = owners_for_lines(lines,people)
    if unknown or not owners:
        st.warning(
            "No se publicará este pedido: algunos pickers no tienen una coincidencia "
            "única con el correo de la plantilla. Revisa la identidad de: "+
            ", ".join(unknown[:5])
        )
        return

    try:
        db=MobileAuditDB(
            st.secrets.get("SUPABASE_URL",""),
            st.secrets.get("SUPABASE_SECRET_KEY") or st.secrets.get("SUPABASE_SERVICE_ROLE_KEY",""),
        )
    except MobileAuditError as exc:
        st.error(str(exc))
        return

    batch=hashlib.sha256(raw_upload).hexdigest()
    key=order_id(batch,pedido)
    st.caption("Picker(s) identificado(s): "+", ".join(
        f"{people[email]['name']} ({email})" for email in owners
    ))
    kind=st.radio(
        "Tipo de revisión para asignar",
        ["autoverificacion","auditoria_cruzada"],
        horizontal=True,
        format_func=lambda v:"Autoverificación" if v=="autoverificacion" else "Auditoría cruzada",
        key="mobile_assign_kind_"+key,
    )
    candidates = (
        owners if kind=="autoverificacion"
        else sorted(set(people)-set(owners))
    )
    if not candidates:
        st.info("No hay personal elegible para este tipo de revisión.")
    else:
        assignee=st.selectbox(
            "Responsable",
            candidates,
            format_func=lambda mail:f"{people[mail]['name']} · {mail}",
            key="mobile_assignee_"+key+"_"+kind,
        )
        if st.button("Asignar pedido al celular",type="primary",key="mobile_assign_submit_"+key+"_"+kind):
            try:
                validated_assignment(kind,assignee,owners)
                payload={
                    "id":key,
                    "batch_id":batch,
                    "pedido":str(pedido),
                    "slot":str(slot),
                    "owner_emails":owners,
                    "picker_labels":str(row.get("Pickers asignados","")),
                    "items":_items_with_owner_email(lines,people),
                    "published_by_email":supervisor,
                }
                db.publish(payload)
                assigned=db.assign({
                    "order_id":key,
                    "mode":kind,
                    "assignee_email":assignee,
                    "assigned_by_email":supervisor,
                })
                if assigned:
                    st.success(f"Pedido #{pedido} asignado a {assignee}.")
                else:
                    st.info("Esta asignación ya existía; no se creó un duplicado.")
            except (MobileAuditError,ValueError) as exc:
                st.error(str(exc))

    if st.button("Consultar validaciones móviles de este pedido",key="mobile_check_"+key):
        try:
            rows=db.assignments_for_order(key)
            if rows:
                st.dataframe([{
                    "Responsable":rec.get("assignee_email",""),
                    "Tipo":"Propio" if rec.get("mode")=="autoverificacion" else "Cruzada",
                    "Estado":"Validado" if rec.get("status")=="validado" else "Pendiente",
                    "Resultado":rec.get("last_result",""),
                } for rec in rows],hide_index=True,use_container_width=True)
                reviews=db.reviews_for_order(key)
                if reviews:
                    st.caption("Registro de intentos y hallazgos. La coincidencia con FNR/MC no determina la causa.")
                    st.dataframe([{
                        "Revisó":rec.get("reviewer_email",""),
                        "Tipo":"Propio" if rec.get("mode")=="autoverificacion" else "Cruzada",
                        "Resultado":rec.get("result",""),
                        "Detalles":rec.get("notes",""),
                        "FNR actual":str(row.get("_audit_fnr_count","")), 
                    } for rec in reviews],hide_index=True,use_container_width=True)
            else:
                st.info("Aún no hay validaciones móviles asignadas a este pedido.")
        except MobileAuditError as exc:
            st.error(str(exc))
