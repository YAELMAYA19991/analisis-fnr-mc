"""Panel público para enviar pedidos a la cola de auditoría móvil."""
import hashlib

import streamlit as st

from mobile_audit_backend import (
    MobileAuditDB, MobileAuditError, SUPABASE_PUBLIC_URL, SUPABASE_PUBLISHABLE_KEY,
    available_people, domain_allowed, email_from_source, normalized_name,
    order_id, order_items, owners_for_lines, validated_assignment,
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


def render_mobile_assignment(pedido, slot, row, lines, roster, summary, raw_upload, fnr_count=0, mc_count=0):
    """Nadie sin OIDC + lista explícita de administradores puede asignar ni ver revisiones."""
    st.caption("Asigna autoverificaciones y auditorías cruzadas sin modificar el Excel original.")
    st.caption("La publicación crea una copia en auditoría móvil; no cambia el pedido original.")
    people = _staff_emails(available_people(roster,summary), "justo.mx")
    owners,unknown = owners_for_lines(lines,people)
    if unknown or not owners:
        st.warning(
            "No se publicará este pedido: algunos pickers no tienen una coincidencia "
            "única con el correo de la plantilla. Revisa la identidad de: "+
            ", ".join(unknown[:5])
        )
        return

    try:
        db=MobileAuditDB(SUPABASE_PUBLIC_URL, SUPABASE_PUBLISHABLE_KEY)
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
        else ["auditoria-publica@coyoacan.invalid"]
    )
    if not candidates:
        st.info("No hay personal elegible para este tipo de revisión.")
    else:
        assignee=st.selectbox(
            "Responsable",
            candidates,
            format_func=lambda mail: (
                f"{people[mail]['name']} · {mail}"
                if mail in people else "Supervisión abierta"
            ),
            key="mobile_assignee_"+key+"_"+kind,
        )
        if st.button("Publicar y enviar al celular",type="primary",key="mobile_assign_submit_"+key+"_"+kind):
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
                    "published_by_email":"acceso-publico",
                }
                db.publish_and_assign_public(payload,kind,assignee)
                st.success(f"Pedido #{pedido} listo en la cola móvil. No se crean duplicados.")
            except (MobileAuditError,ValueError) as exc:
                st.error(str(exc))

    if st.button("Consultar validaciones móviles de este pedido",key="mobile_check_"+key):
        try:
            rows=db.assignments_for_order(key)
            if rows:
                st.dataframe([{
                    "Responsable":"Picker asignado" if rec.get("mode")=="autoverificacion" else "Supervisión abierta",
                    "Tipo":"Propio" if rec.get("mode")=="autoverificacion" else "Cruzada",
                    "Estado":"Validado" if rec.get("status")=="validado" else "Pendiente",
                    "Resultado":rec.get("last_result",""),
                } for rec in rows],hide_index=True,use_container_width=True)
                reviews=db.reviews_for_order(key)
                if reviews:
                    st.caption(
                        f"Cruce con archivos operativos cargados: FNR {fnr_count:,} · MC {mc_count:,}. "
                        "Una coincidencia requiere investigación y no determina la causa."
                    )
                    st.dataframe([{
                        "Revisó":rec.get("reviewer_name",""),
                        "Tipo":"Propio" if rec.get("mode")=="autoverificacion" else "Cruzada",
                        "Resultado":rec.get("result",""),
                        "Artículos con diferencia":len(rec.get("differences",[]) or []),
                        "Detalles":rec.get("notes",""),
                        "Fecha":rec.get("created_at",""),
                    } for rec in reviews],hide_index=True,use_container_width=True)
            else:
                st.info("Aún no hay validaciones móviles asignadas a este pedido.")
        except MobileAuditError as exc:
            st.error(str(exc))
