"""Servidor compartido de auditorías móviles: REST Supabase; sin claves en cliente."""
from __future__ import annotations

import hashlib
import hmac
import json
import re
import unicodedata
import urllib.error
import urllib.parse
import urllib.request


EMAIL_PATTERN = re.compile(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9.-]+[.][A-Za-z]{2,}")
MODES = {"autoverificacion", "auditoria_cruzada"}
RESULTS = {"Pedido correcto", "Con diferencias", "No se pudo validar"}


SUPABASE_PUBLIC_URL = "https://kywxkynoigufkyzzaryx.supabase.co"
SUPABASE_PUBLISHABLE_KEY = "sb_publishable_pvXrzejU-C-T4MKr_smDyQ__DyfxYm3"

class MobileAuditError(RuntimeError):
    pass


def email_from_source(value):
    match = EMAIL_PATTERN.search(str(value or ""))
    return match.group(0).lower() if match else ""


def verified_identity(user):
    """Nunca acepta un correo tecleado: sólo el claim OIDC verificado."""
    if not user or not bool(user.get("is_logged_in", False)):
        return ""
    verified = user.get("email_verified")
    if verified is not True and str(verified).lower() != "true":
        return ""
    raw = str(user.get("email", "")).strip().lower()
    return raw if EMAIL_PATTERN.fullmatch(raw) else ""


def email_set(raw):
    if isinstance(raw, (tuple, list, set)):
        candidates = raw
    else:
        candidates = str(raw or "").replace(";", ",").replace(chr(10), ",").split(",")
    return {str(v).strip().lower() for v in candidates if EMAIL_PATTERN.fullmatch(str(v).strip().lower())}


def admin_allowed(email, configured):
    return bool(email) and email in email_set(configured)


def supervisor_pin_valid(entered, configured):
    """Compara el código privado sin exponerlo ni usar correo para iniciar sesión."""
    candidate = str(entered or "")
    expected = str(configured or "")
    return len(expected) >= 8 and hmac.compare_digest(candidate, expected)


def domain_allowed(email, configured):
    domains = {v.strip().lstrip("@").lower() for v in str(configured or "").replace(";", ",").replace(chr(10), ",").split(",") if v.strip()}
    return bool(email) and bool(domains) and email.rsplit("@", 1)[-1] in domains


def normalized_name(value):
    raw = unicodedata.normalize("NFKD", str(value or "").lower())
    raw = "".join(c for c in raw if not unicodedata.combining(c))
    return " ".join(re.findall(r"[a-z0-9]+", raw))


def available_people(roster, summary=None):
    """Sólo personas con identidad CORREO establecida en las fuentes existentes."""
    people = {}
    for frame in (roster, summary):
        if frame is None or frame.empty:
            continue
        for row in frame.to_dict("records"):
            email = email_from_source(row.get("CORREO", ""))
            if not email:
                continue
            name = str(row.get("PICKER", "") or "").strip()
            if not name or name.lower() in {"nan", "none"}:
                continue
            record = people.setdefault(email, {"email": email, "name": name, "names": set()})
            record["name"] = name
            for field in ("PICKER", "_SOURCE_PICKER", "_MASTER_PICKER"):
                candidate = normalized_name(row.get(field, ""))
                if candidate and candidate not in {"sin asignar", "nan"}:
                    record["names"].add(candidate)
    return people


def owners_for_lines(lines, people):
    """Falla cerrado si un picker del pedido no se vincula de manera única al correo."""
    names = {str(v).strip() for v in lines["Picker relacionado"].tolist()}
    owners = set()
    unresolved = []
    for name in sorted(names):
        direct_email = email_from_source(name)
        matches = (
            [direct_email] if direct_email and direct_email in people
            else [email for email, person in people.items() if normalized_name(name) in person["names"]]
        )
        if len(matches) != 1:
            unresolved.append(name)
        else:
            owners.add(matches[0])
    return sorted(owners), unresolved


def order_id(batch_sha, order_number):
    digest = hashlib.sha256((str(batch_sha) + "|" + str(order_number)).encode("utf-8")).hexdigest()
    return digest


def order_items(lines):
    return [{
        "sku": str(row.get("SKU", "") or ""),
        "articulo": str(row.get("Artículo", "") or ""),
        "picker": str(row.get("Picker relacionado", "") or ""),
        "cantidad_pedida": str(row.get("Cantidad pedida", "") or ""),
        "cantidad_pickeada": str(row.get("Cantidad pickeada", "") or ""),
    } for row in lines.to_dict("records")]


def validated_assignment(mode, assignee, owners):
    if mode not in MODES:
        raise ValueError("Tipo de revisión desconocido.")
    if not assignee or not EMAIL_PATTERN.fullmatch(assignee):
        raise ValueError("Correo del responsable no válido.")
    if not owners:
        raise ValueError("Antes de asignar hay que identificar al picker del pedido.")
    if mode == "autoverificacion" and assignee not in owners:
        raise ValueError("La autoverificación debe asignarse al picker que preparó el pedido.")
    if mode == "auditoria_cruzada" and assignee in owners:
        raise ValueError("No se puede auditar como tercero un pedido propio.")
    return True


def _request(url, key, method, resource, *, params=None, payload=None, prefer=""):
    if not url or not key:
        raise MobileAuditError("Configura SUPABASE_URL y SUPABASE_SECRET_KEY en el servidor.")
    if not str(url).startswith("https://"):
        raise MobileAuditError("La dirección de Supabase debe usar HTTPS.")
    query = urllib.parse.urlencode(params or {})
    target = url.rstrip("/") + "/rest/v1/" + resource + (("?" + query) if query else "")
    raw = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(target, method=method, data=raw)
    req.add_header("apikey", key)
    if not key.startswith(("sb_secret_", "sb_publishable_")):
        req.add_header("Authorization", "Bearer " + key)
    req.add_header("Content-Type", "application/json")
    req.add_header("Accept", "application/json")
    if prefer:
        req.add_header("Prefer", prefer)
    try:
        with urllib.request.urlopen(req, timeout=18) as response:
            body = response.read().decode("utf-8")
            return json.loads(body) if body else []
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", "replace")
        try:
            error = json.loads(body)
            message = str(error.get("message", ""))[:220]
        except ValueError:
            message = ""
        raise MobileAuditError(
            f"Supabase devolvió HTTP {exc.code}. {message or 'Comprueba las tablas y sus permisos.'}"
        ) from exc
    except (urllib.error.URLError, TimeoutError) as exc:
        raise MobileAuditError("No se pudo conectar con Supabase. Intenta nuevamente.") from exc


class MobileAuditDB:
    def __init__(self, url, secret_key):
        self.url = str(url or "").rstrip("/")
        self.key = str(secret_key or "")
        if not self.url or not self.key:
            raise MobileAuditError("No hay credenciales del servidor para auditoría móvil.")

    def request(self, method, resource, **kwargs):
        return _request(self.url, self.key, method, resource, **kwargs)

    def publish(self, order):
        """Repetir la publicación del mismo lote no sobrescribe revisiones."""
        return self.request(
            "POST", "mobile_audit_orders",
            params={"on_conflict": "id"}, payload=[order],
            prefer="resolution=ignore-duplicates,return=representation",
        )

    def publish_and_assign_public(self, order, mode, assignee_email):
        """Publica una copia para auditoría y crea la asignación sin tocar el pedido original."""
        return self.request(
            "POST", "rpc/publish_mobile_audit_order_public",
            payload={
                "p_order": order,
                "p_mode": mode,
                "p_assignee_email": assignee_email,
            },
        )

    def assign(self, assignment):
        return self.request(
            "POST", "mobile_audit_assignments",
            params={"on_conflict": "order_id,mode,assignee_email"},
            payload=[assignment],
            prefer="resolution=ignore-duplicates,return=representation",
        )

    def mine(self, email):
        if not EMAIL_PATTERN.fullmatch(email or ""):
            raise MobileAuditError("No se pudo identificar el correo de la sesión.")
        return self.request(
            "GET", "mobile_audit_assignments",
            params={
                "select": "id,mode,status,last_result,order_id,mobile_audit_orders(pedido,slot,items,picker_labels)",
                "assignee_email": "eq." + email,
                "order": "created_at.desc",
                "limit": "100",
            },
        )

    def all_assignments(self):
        """Obtiene la cola pública mediante una función que devuelve datos de pedido saneados."""
        return self.request("POST", "rpc/list_mobile_audit_assignments_public", payload={})

    def assignments_for_order(self, order_key):
        """Filtra la cola pública saneada por el ID del pedido."""
        return [row for row in self.all_assignments() if row.get("order_id") == order_key]

    def reviews_for_order(self, order_key):
        """Consulta historial sin correo de auditor, por RPC público limitado."""
        return self.request(
            "POST", "rpc/list_mobile_audit_reviews_public",
            payload={"p_order_id": order_key},
        )

    def submit(self, assignment_id, email, result, notes, differences, checked_items, nonce):
        if result not in RESULTS or not isinstance(differences, list):
            raise MobileAuditError("Resultado de validación no admitido.")
        return self.request(
            "POST", "rpc/submit_mobile_audit_review",
            payload={
                "p_assignment_id": assignment_id,
                "p_reviewer_email": email,
                "p_result": result,
                "p_notes": str(notes or "").strip(),
                "p_differences": differences,
                "p_checked_items": int(checked_items),
                "p_idempotency_key": nonce,
            },
        )


    def submit_supervisor(self, assignment_id, reviewer_name, result, notes, differences, checked_items, nonce):
        """Registra revisión móvil con nombre del supervisor y sin identidad de correo."""
        return self.request(
            "POST", "rpc/submit_mobile_audit_review_supervisor",
            payload={
                "p_assignment_id": assignment_id,
                "p_reviewer_name": str(reviewer_name or "").strip(),
                "p_result": result,
                "p_notes": str(notes or "").strip(),
                "p_differences": differences,
                "p_checked_items": int(checked_items),
                "p_idempotency_key": nonce,
            },
        )
