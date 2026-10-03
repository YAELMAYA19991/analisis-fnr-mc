# Validación Coyoacán — activación de la miniaplicación móvil

## Archivos

- `app.py`: mantiene la auditoría compacta de FNR–MC y permite asignar pedidos a cada correo corporativo desde la sección plegada **Asignar validación móvil (supervisores)**.
- `picker_mobile.py`: nueva miniaplicación independiente para celulares.
- `mobile_audit_backend.py`: identidad, reglas de negocio y consultas seguras de servidor.
- `mobile_audit_admin.py`: panel de administración dentro de la auditoría FNR–MC.
- `supabase_mobile_audit.sql`: crea únicamente las nuevas tablas para validación móvil; no modifica las tablas existentes.
- `tests/test_mobile_audit.py`: regresiones de permisos y separación de funciones.

## Requisitos ANTES de dar acceso al personal

1. Entra a **Supabase → proyecto conectado a FNR–MC → SQL Editor**. Revisa y ejecuta `supabase_mobile_audit.sql` una única vez. Confirma las tablas `mobile_audit_orders`, `mobile_audit_assignments` y `mobile_audit_reviews`.
2. En **Google Cloud → Google Auth Platform**, configura un cliente OAuth Web para Streamlit (o usa Microsoft Entra si las cuentas corporativas son Microsoft). Agrega como URL de retorno **cada** dirección de Streamlit terminada en `/oauth2callback`.
3. En **Streamlit Community Cloud**, crea otra aplicación desde el repositorio `YAELMAYA19991/analisis-fnr-mc`, rama `main`, archivo principal **`picker_mobile.py`**. Conserva la aplicación principal existente `app.py` sin reemplazarla.
4. Configura los Secrets de **ambas aplicaciones** desde el menú Settings → Secrets. No subas claves al repositorio ni las compartas en un chat.
5. En la aplicación FNR–MC: carga plantilla de personal y Excel de pedidos, selecciona un pedido, abre **Asignar validación móvil (supervisores)**, inicia sesión y asigna la autoverificación a su picker; la auditoría cruzada a un correo corporativo distinto.
6. En la miniaplicación: ingresa con el correo asignado, revisa los artículos, registra el resultado y comprueba el historial en FNR–MC.

## Ejemplo de Secrets para la aplicación móvil

Reemplaza los valores de ejemplo en la configuración privada de Streamlit:

```toml
SUPABASE_URL = "https://TU-PROYECTO.supabase.co"
SUPABASE_SECRET_KEY = "TU-CLAVE-PRIVADA-DE-SERVIDOR"
MOBILE_AUDIT_ALLOWED_DOMAINS = "empresa.com"

[auth]
redirect_uri = "https://TU-MINIAPP.streamlit.app/oauth2callback"
cookie_secret = "GENERAR_UNA_CADENA_ALEATORIA_LARGA"
client_id = "CLIENT_ID_DEL_PROVEEDOR"
client_secret = "CLIENT_SECRET_DEL_PROVEEDOR"
server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"
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

- Sólo un usuario con identidad OIDC verificada y correo corporativo autorizado puede entrar a la miniapp; sólo los correos de `MOBILE_AUDIT_ADMIN_EMAILS` pueden asignar pedidos desde FNR–MC.
- Nadie puede auditar como tercero un pedido propio; la base de datos también valida esta regla.
- El picker ve solamente los pedidos asignados a su correo, y en autoverificación únicamente los artículos que él preparó.
- **Pedido correcto** y **Con diferencias** cierran la asignación después de confirmar todos los artículos. **No se pudo validar** deja el pedido pendiente y registra el intento.
- El Excel original no se modifica. Los datos quedan en las tres tablas nuevas de Supabase. La pestaña FNR–MC permite consultar validaciones y compararlas con las coincidencias de FNR/MC en los archivos actuales.
- El 20 % de selección por hora en la auditoría principal se mantiene; la asignación móvil es manual para que no altere los criterios de selección.
- Un pedido con un picker no identificado por correo se bloquea antes de publicarse.
- El enlace de la miniapp puede compartirse entre empleados, pero sólo muestra los pedidos del correo autenticado.

## Pruebas recomendadas antes de producción

1. Cuenta picker A: ver su autoverificación y sus propias líneas; no poder consultar la asignación de B.
2. Cuenta picker B: ver una auditoría cruzada del pedido de A, no poder autoauditar ese pedido.
3. Registrar "No se pudo validar" y verificar que el pedido permanece pendiente.
4. Registrar "Con diferencias" e inspeccionar el registro en FNR–MC.
5. Recargar y comprobar persistencia después de un cierre de sesión.
6. Recargar el Excel y probar una asignación duplicada: no debe crear dos revisiones.
7. Verificar que sin configuración OIDC o sin dominio corporativo permitido la aplicación deniega acceso.

Esta integración **no queda publicada** simplemente por tener el código en GitHub: deben completarse los pasos de Supabase, OAuth, Secrets y creación de la segunda app en Streamlit.
