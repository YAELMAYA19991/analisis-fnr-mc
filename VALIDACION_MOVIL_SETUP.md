# Validación Coyoacán — activación de la miniaplicación móvil

## Archivos

- `app.py`: mantiene la auditoría compacta de FNR–MC y permite asignar pedidos a cada correo corporativo desde la sección plegada **Asignar validación móvil (supervisores)**.
- `picker_mobile.py`: miniaplicación pública para celulares; no pide correo ni código y registra el nombre de quien audita.
- `mobile_audit_backend.py`: acceso público limitado por funciones de Supabase, sin clave privada de servidor en la miniapp.
- `mobile_audit_admin.py`: panel de administración dentro de la auditoría FNR–MC.
- `supabase_mobile_audit.sql`: crea únicamente las nuevas tablas para validación móvil; no modifica las tablas existentes.
- `supabase_supervisor_pin_migration.sql`: agrega el nombre del supervisor a las revisiones móviles sin borrar datos existentes.
- `supabase_public_mobile_audit_rpc.sql`: limita el acceso público a lectura saneada y registro de revisiones.
- `supabase_public_mobile_publish_rpc.sql`: permite enviar el pedido seleccionado a la cola de auditoría, sin editar el original.
- `tests/test_mobile_audit.py`: regresiones de permisos y separación de funciones.

## Requisitos ANTES de dar acceso al personal

1. Entra a **Supabase → proyecto conectado a FNR–MC → SQL Editor**. Revisa y ejecuta `supabase_mobile_audit.sql` una única vez. Confirma las tablas `mobile_audit_orders`, `mobile_audit_assignments` y `mobile_audit_reviews`.
2. La miniaplicación móvil ya no necesita Google OAuth ni URL de retorno. El panel de asignaciones dentro de FNR–MC conserva su configuración actual; este cambio no modifica esa app principal.
3. En **Streamlit Community Cloud**, crea otra aplicación desde el repositorio `YAELMAYA19991/analisis-fnr-mc`, rama `main`, archivo principal **`picker_mobile.py`**. Conserva la aplicación principal existente `app.py` sin reemplazarla.
4. La miniaplicación ya no requiere Secrets de Supabase. Usa una clave publishable pública y funciones restringidas; la clave privada de servidor no se copia a esta app.
5. En FNR–MC: carga plantilla y pedidos, selecciona cada pedido que se quiera auditar, abre **Asignar validación móvil (supervisores)**, elige el tipo y pulsa **Publicar y enviar al celular**. No solicita correo ni inicio de sesión.
6. En la miniaplicación: abre el enlace sin iniciar sesión ni introducir códigos, escribe el nombre de quien audita, revisa los artículos y registra el resultado. Comprueba el historial en FNR–MC.

## Configuración privada de FNR–MC

La app móvil no requiere claves privadas en Streamlit Community Cloud. La clave publishable está diseñada para uso público y sus permisos están limitados en la base de datos; no es una clave de servicio.

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

La configuración OAuth anterior aplica únicamente a FNR–MC. La miniapp móvil no utiliza Google OAuth ni requiere inicio de sesión.

## Reglas de operación

- Las dos pantallas móviles son públicas: cualquier persona con los enlaces puede publicar pedidos seleccionados desde FNR–MC, consultar la cola y registrar auditorías o diferencias. El nombre del auditor es autodeclarado y no comprueba su identidad.
- Las asignaciones existentes siguen asociadas internamente al picker del roster para limitar los artículos de autoverificación; el picker no inicia sesión ni escribe su correo.
- Nadie puede auditar como tercero un pedido propio; la base de datos también valida esta regla.
- La miniapp muestra a quien tenga el enlace la cola de pedidos asignados; en autoverificación presenta únicamente los artículos asociados al picker asignado y omite correos internos.
- **Pedido correcto** y **Con diferencias** cierran la asignación después de confirmar todos los artículos. **No se pudo validar** deja el pedido pendiente y registra el intento.
- El Excel original no se modifica. Los datos quedan en las tres tablas nuevas de Supabase. La pestaña FNR–MC permite consultar validaciones y compararlas con las coincidencias de FNR/MC en los archivos actuales.
- El 20 % de selección por hora en la auditoría principal se mantiene; la asignación móvil es manual para que no altere los criterios de selección.
- Un pedido con un picker no identificado por correo se bloquea antes de publicarse.
- Protege la distribución del enlace: cualquier persona que lo reciba puede consultar los pedidos y registrar revisiones. La app no pide correo ni PIN.

## Pruebas recomendadas antes de producción

1. Abre la miniapp en un teléfono: debe mostrar la cola sin pedir correo ni código.
2. En una autoverificación, confirma que se muestran sólo los artículos del picker asignado y que no aparecen correos internos.
3. Registrar "No se pudo validar" y verificar que el pedido permanece pendiente.
4. Registrar "Con diferencias" e inspeccionar el registro en FNR–MC.
5. Recargar la app y comprobar que los registros guardados siguen apareciendo.
6. Recargar el Excel y probar una asignación duplicada: no debe crear dos revisiones.
7. Confirma que el registro de auditorías y diferencias aparezca en FNR–MC y que los pedidos originales no se hayan modificado.

La app móvil y la sección de publicación funcionan sin código, correo o cuenta. La clave publishable de Supabase es pública; las tablas siguen bloqueadas para acceso directo y los cambios se envían por funciones limitadas. La app FNR–MC original y los pedidos originales no se modifican con esta configuración.
