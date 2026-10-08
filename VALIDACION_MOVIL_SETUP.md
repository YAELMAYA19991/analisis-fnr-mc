# Validación Coyoacán — activación de la miniaplicación móvil

## Archivos

- `app.py`: mantiene la auditoría compacta de FNR–MC y permite asignar pedidos a cada correo corporativo desde la sección plegada **Asignar validación móvil (supervisores)**.
- `picker_mobile.py`: miniaplicación pública para celulares; no pide correo ni código y registra el nombre de quien audita.
- `mobile_audit_backend.py`: identidad, reglas de negocio y consultas seguras de servidor.
- `mobile_audit_admin.py`: panel de administración dentro de la auditoría FNR–MC.
- `supabase_mobile_audit.sql`: crea únicamente las nuevas tablas para validación móvil; no modifica las tablas existentes.
- `supabase_supervisor_pin_migration.sql`: agrega el nombre del supervisor a las revisiones móviles sin borrar datos existentes.
- `tests/test_mobile_audit.py`: regresiones de permisos y separación de funciones.

## Requisitos ANTES de dar acceso al personal

1. Entra a **Supabase → proyecto conectado a FNR–MC → SQL Editor**. Revisa y ejecuta `supabase_mobile_audit.sql` una única vez. Confirma las tablas `mobile_audit_orders`, `mobile_audit_assignments` y `mobile_audit_reviews`.
2. La miniaplicación móvil ya no necesita Google OAuth ni URL de retorno. El panel de asignaciones dentro de FNR–MC conserva su configuración actual; este cambio no modifica esa app principal.
3. En **Streamlit Community Cloud**, crea otra aplicación desde el repositorio `YAELMAYA19991/analisis-fnr-mc`, rama `main`, archivo principal **`picker_mobile.py`**. Conserva la aplicación principal existente `app.py` sin reemplazarla.
4. En los Secrets de la miniaplicación conserva las credenciales privadas existentes de Supabase. No subas claves al repositorio ni las compartas en un chat.
5. En la aplicación FNR–MC: carga plantilla de personal y Excel de pedidos, selecciona un pedido, abre **Asignar validación móvil (supervisores)**, inicia sesión y asigna la autoverificación a su picker; la auditoría cruzada a un correo corporativo distinto.
6. En la miniaplicación: abre el enlace sin iniciar sesión ni introducir códigos, escribe el nombre de quien audita, revisa los artículos y registra el resultado. Comprueba el historial en FNR–MC.

## Ejemplo de Secrets para la aplicación móvil

Reemplaza los valores de ejemplo en la configuración privada de Streamlit:

```toml
SUPABASE_URL = "https://TU-PROYECTO.supabase.co"
SUPABASE_SECRET_KEY = "TU-CLAVE-PRIVADA-DE-SERVIDOR"
```

**Importante:** `SUPABASE_SECRET_KEY` es privada, de uso exclusivo en el servidor. Nunca debe estar en código del repositorio, un enlace ni un código QR.

## Secrets adicionales de la aplicación FNR–MC original

Conserva los Secrets existentes y agrega/ajusta:

```toml
MOBILE_AUDIT_ALLOWED_DOMAINS = "empresa.com"
MOBILE_AUDIT_ADMIN_EMAILS = "supervisor@empresa.com,admin@empresa.com"

[auth]
redirect_uri = "https://fnr-mc-coyoacan.streamlit.app/oauth2callback"
cookie_secret = "OTRA_CADENA_ALEATORIA_LARGA"
client_id = "CLIENT_ID_DEL_PROVEEDOR"
client_secret = "CLIENT_SECRET_DEL_PROVEEDOR"
server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"
```

Las dos apps pueden usar el mismo proyecto Supabase, pero cada una necesita sus propios Secrets y una `redirect_uri` exacta. Registra ambas URL de retorno ante el proveedor OAuth. Si ya existe una sección `[auth]`, **edítala** en vez de duplicarla.

## Reglas de operación

- La miniapp es pública: cualquier persona con el enlace puede ver los pedidos asignados y registrar auditorías o diferencias. El nombre del auditor es autodeclarado y no comprueba su identidad.
- Las asignaciones existentes siguen asociadas internamente al picker del roster para limitar los artículos de autoverificación; el picker no inicia sesión ni escribe su correo.
- Nadie puede auditar como tercero un pedido propio; la base de datos también valida esta regla.
- La miniapp muestra a quien tenga el enlace la cola de pedidos asignados; en autoverificación presenta únicamente los artículos asociados al picker asignado.
- **Pedido correcto** y **Con diferencias** cierran la asignación después de confirmar todos los artículos. **No se pudo validar** deja el pedido pendiente y registra el intento.
- El Excel original no se modifica. Los datos quedan en las tres tablas nuevas de Supabase. La pestaña FNR–MC permite consultar validaciones y compararlas con las coincidencias de FNR/MC en los archivos actuales.
- El 20 % de selección por hora en la auditoría principal se mantiene; la asignación móvil es manual para que no altere los criterios de selección.
- Un pedido con un picker no identificado por correo se bloquea antes de publicarse.
- Protege la distribución del enlace: cualquier persona que lo reciba puede consultar los pedidos y registrar revisiones. La app no pide correo ni PIN.

## Pruebas recomendadas antes de producción

1. Abre la miniapp en un teléfono: debe mostrar la cola sin pedir correo ni código.
2. Cuenta picker B: ver una auditoría cruzada del pedido de A, no poder autoauditar ese pedido.
3. Registrar "No se pudo validar" y verificar que el pedido permanece pendiente.
4. Registrar "Con diferencias" e inspeccionar el registro en FNR–MC.
5. Recargar y comprobar persistencia después de un cierre de sesión.
6. Recargar el Excel y probar una asignación duplicada: no debe crear dos revisiones.
7. Confirma que el registro de auditorías y diferencias aparezca en FNR–MC y que los pedidos originales no se hayan modificado.

El enlace público permite consultar pedidos y enviar revisiones. La clave de Supabase permanece únicamente en Secrets del servidor. La app FNR–MC original y los pedidos originales no se modifican con esta configuración.
