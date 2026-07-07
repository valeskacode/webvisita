# -*- coding: utf-8 -*-
"""
utils/onedrive.py
Integración con Microsoft Graph API para subir reportes (Word/PDF),
el historial Excel y las fotos de visita directamente a OneDrive,
sin necesitar el cliente de escritorio instalado en el servidor.

Usa el flujo "Client Credentials" (app-only, sin login interactivo de
usuario). Esto requiere que en Azure AD → App registrations →
appvisitareporte → "Permisos de API" el permiso Files.ReadWrite.All
esté agregado como permiso de tipo **Aplicación** (no Delegado), y que
un administrador global haya presionado "Otorgar consentimiento
administrativo". Si el permiso quedó como Delegado, esta integración
fallará con error 401/403 sin importar que las credenciales sean
correctas, porque el flujo Client Credentials no tiene un usuario
autenticado detrás.

Requiere las credenciales de la app registrada en Azure AD:
  - CLIENT_ID     (ID de la aplicación)
  - CLIENT_SECRET (secreto de cliente — el VALOR, no el ID del secreto)
  - TENANT_ID     (ID del directorio / inquilino)

Estas credenciales se configuran en Streamlit Cloud → Settings →
Secrets (nunca en el código ni en el repositorio), con esta forma:

    [graph]
    client_id = "f4f4f84d-93c0-4450-97f6-8fdf25f14de1"
    client_secret = "el-valor-real-del-secreto"
    tenant_id = "f3831aea-ec1b-461b-b42f-ca26f9f78551"
    onedrive_user = "auditoria@cajaarequipa.pe"
    onedrive_carpeta = "Auditoria/VisitaClientes"

También se aceptan variables de entorno (GRAPH_CLIENT_ID, etc.) como
alternativa, útil para pruebas locales.
"""
import io
import os

import requests

try:
    import streamlit as st
except Exception:  # pragma: no cover
    st = None


def _secret(key: str, default: str = "") -> str:
    """Lee una credencial desde st.secrets['graph'][key], luego desde
    st.secrets[key] (compatibilidad), y por último desde una variable
    de entorno GRAPH_<KEY>."""
    if st is not None:
        try:
            if "graph" in st.secrets and key in st.secrets["graph"]:
                return str(st.secrets["graph"][key])
        except Exception:
            pass
        try:
            if key in st.secrets:
                return str(st.secrets[key])
        except Exception:
            pass
    return os.environ.get(f"GRAPH_{key.upper()}", default)


# --------------------------------------------------------------------------
# CONFIGURACIÓN — lee desde Streamlit Secrets (o variables de entorno)
# --------------------------------------------------------------------------
CLIENT_ID     = _secret("client_id")
CLIENT_SECRET = _secret("client_secret")
TENANT_ID     = _secret("tenant_id")

# Correo / UPN del usuario de OneDrive donde se guardarán los archivos.
# Ej: "auditoria@cajaarequipa.pe"
ONEDRIVE_USER = _secret("onedrive_user")

# Ruta de la carpeta DENTRO de ese OneDrive donde caerán los reportes.
# Ej: "Auditoria/VisitaClientes"
# Deja vacío ("") para que quede en la raíz del OneDrive.
ONEDRIVE_CARPETA = _secret("onedrive_carpeta", "Auditoria/VisitaClientes")

# Link de "Compartir" de la carpeta de OneDrive/SharePoint donde está el
# Excel con la cartera de clientes (el mismo que se sube manualmente en
# "Búsqueda"). Configúralo en Secrets como base_share_url si cambia.
ONEDRIVE_BASE_SHARE_URL = _secret(
    "base_share_url",
    "https://cajaarequipape-my.sharepoint.com/:f:/g/personal/vherrera_cajaarequipa_pe/"
    "IgDFL8eh0Jv9TY56xnomN8DgASKbGXNZRG1OF15eSzXgVj0?e=WSdtmm",
)

# URL base de Graph API
GRAPH_URL = "https://graph.microsoft.com/v1.0"

# --------------------------------------------------------------------------
# TOKEN — Client Credentials Flow (app-only, sin login interactivo).
# Necesita permiso "Files.ReadWrite.All" en la app de Azure.
# --------------------------------------------------------------------------
_token_cache: dict = {}


def _obtener_token() -> str:
    """Obtiene (o reutiliza en caché) un token de acceso de Graph API
    usando el flujo Client Credentials (sin login del usuario)."""
    import time
    ahora = time.time()
    if _token_cache.get("expires_at", 0) > ahora + 60:
        return _token_cache["access_token"]

    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=15)
    resp.raise_for_status()
    data = resp.json()
    _token_cache["access_token"] = data["access_token"]
    _token_cache["expires_at"] = ahora + data.get("expires_in", 3600)
    return _token_cache["access_token"]


def _headers() -> dict:
    return {"Authorization": f"Bearer {_obtener_token()}"}


def credenciales_configuradas() -> bool:
    """True si las tres credenciales mínimas están presentes."""
    return bool(CLIENT_ID and CLIENT_SECRET and TENANT_ID and ONEDRIVE_USER)


# --------------------------------------------------------------------------
# OPERACIONES CON ONEDRIVE
# --------------------------------------------------------------------------

def _ruta_onedrive(nombre_archivo: str, subcarpeta: str = "") -> str:
    """Construye la ruta completa dentro del OneDrive del usuario."""
    base = ONEDRIVE_CARPETA.strip("/")
    if subcarpeta:
        base = f"{base}/{subcarpeta.strip('/')}"
    if base:
        return f"{base}/{nombre_archivo}"
    return nombre_archivo


def _upload_url(ruta_en_onedrive: str) -> str:
    """URL de Graph API para subir un archivo por ruta."""
    ruta_enc = requests.utils.quote(ruta_en_onedrive, safe="/")
    return f"{GRAPH_URL}/users/{ONEDRIVE_USER}/drive/root:/{ruta_enc}:/content"


def subir_archivo(nombre_archivo: str, contenido_bytes: bytes,
                  subcarpeta: str = "") -> tuple[bool, str]:
    """Sube un archivo a OneDrive.

    Parámetros:
        nombre_archivo: nombre del archivo (con extensión).
        contenido_bytes: contenido del archivo en bytes.
        subcarpeta: subcarpeta adicional dentro de ONEDRIVE_CARPETA.

    Retorna:
        (True, url_web_del_archivo) si tuvo éxito.
        (False, mensaje_de_error)   si falló.
    """
    if not credenciales_configuradas():
        return False, "Credenciales de Graph API no configuradas."
    try:
        ruta = _ruta_onedrive(nombre_archivo, subcarpeta)
        url  = _upload_url(ruta)
        # Graph API acepta archivos ≤ 4 MB con un PUT simple.
        # Para archivos más grandes habría que usar upload session.
        resp = requests.put(url, headers=_headers(), data=contenido_bytes,
                            timeout=60)
        resp.raise_for_status()
        web_url = resp.json().get("webUrl", "")
        return True, web_url
    except Exception as e:
        return False, str(e)


def subir_reporte(nombre_archivo: str, contenido_bytes: bytes) -> tuple[bool, str]:
    """Atajo para subir un reporte Word/PDF a la subcarpeta 'Reportes'."""
    return subir_archivo(nombre_archivo, contenido_bytes, subcarpeta="Reportes")


def subir_historial(contenido_bytes: bytes) -> tuple[bool, str]:
    """Sube el Excel de historial de visitas a OneDrive."""
    return subir_archivo("historial_visitas.xlsx", contenido_bytes, subcarpeta="Historial")


def subir_foto(nombre_archivo: str, foto_bytes: bytes,
               agencia: str = "") -> tuple[bool, str]:
    """Sube una foto de verificación a la subcarpeta Fotos/<agencia>."""
    sub = f"Fotos/{agencia}" if agencia else "Fotos"
    return subir_archivo(nombre_archivo, foto_bytes, subcarpeta=sub)


def listar_carpeta(subcarpeta: str = "") -> list[dict]:
    """Lista los archivos en la carpeta configurada (o subcarpeta).

    Retorna lista de dicts con: name, size, lastModifiedDateTime, webUrl.
    """
    if not credenciales_configuradas():
        return []
    try:
        base = ONEDRIVE_CARPETA.strip("/")
        if subcarpeta:
            base = f"{base}/{subcarpeta.strip('/')}"
        ruta_enc = requests.utils.quote(base, safe="/")
        url = f"{GRAPH_URL}/users/{ONEDRIVE_USER}/drive/root:/{ruta_enc}:/children"
        resp = requests.get(url, headers=_headers(), timeout=15)
        resp.raise_for_status()
        items = resp.json().get("value", [])
        return [
            {
                "name": i.get("name"),
                "size": i.get("size"),
                "fecha": i.get("lastModifiedDateTime", "")[:10],
                "webUrl": i.get("webUrl", ""),
            }
            for i in items if not i.get("folder")  # solo archivos, no carpetas
        ]
    except Exception:
        return []


def test_conexion() -> tuple[bool, str]:
    """Prueba rápida de conexión — verifica el token y el acceso al drive."""
    if not credenciales_configuradas():
        return False, "Falta configurar CLIENT_ID, CLIENT_SECRET, TENANT_ID o ONEDRIVE_USER."
    try:
        _obtener_token()
        url  = f"{GRAPH_URL}/users/{ONEDRIVE_USER}/drive"
        resp = requests.get(url, headers=_headers(), timeout=10)
        resp.raise_for_status()
        nombre = resp.json().get("owner", {}).get("user", {}).get("displayName", ONEDRIVE_USER)
        return True, f"Conectado correctamente al OneDrive de: {nombre}"
    except requests.HTTPError as e:
        codigo = e.response.status_code if e.response else "?"
        if codigo == 401:
            return False, ("Error 401 — Credenciales incorrectas, secreto vencido, "
                            "o falta otorgar consentimiento administrativo en Azure.")
        if codigo == 403:
            return False, ("Error 403 — La app no tiene permiso 'Files.ReadWrite.All' de tipo "
                            "**Aplicación** en Azure (si está como 'Delegado' no funciona con "
                            "este flujo), o falta el consentimiento administrativo.")
        if codigo == 404:
            return False, f"Error 404 — El usuario '{ONEDRIVE_USER}' no se encontró en este tenant."
        return False, f"Error HTTP {codigo}: {e}"
    except Exception as e:
        return False, f"Error de conexión: {e}"


# --------------------------------------------------------------------------
# CARPETA — link directo para abrir en el navegador / app de OneDrive
# --------------------------------------------------------------------------

def carpeta_weburl(subcarpeta: str = "") -> str:
    """Devuelve el webUrl (link para abrir en el navegador) de la carpeta
    configurada en ONEDRIVE_CARPETA, o de una subcarpeta dentro de ella."""
    if not credenciales_configuradas():
        return ""
    try:
        base = ONEDRIVE_CARPETA.strip("/")
        if subcarpeta:
            base = f"{base}/{subcarpeta.strip('/')}" if base else subcarpeta.strip("/")
        ruta_enc = requests.utils.quote(base, safe="/")
        url = f"{GRAPH_URL}/users/{ONEDRIVE_USER}/drive/root:/{ruta_enc}"
        resp = requests.get(url, headers=_headers(), timeout=10)
        resp.raise_for_status()
        return resp.json().get("webUrl", "")
    except Exception:
        return ""


# --------------------------------------------------------------------------
# BASE DE CLIENTES — descargar el Excel directamente desde un link
# de OneDrive/SharePoint compartido (sin subirlo manualmente desde el móvil)
# --------------------------------------------------------------------------

def _encode_share_url(url: str) -> str:
    """Codifica una URL para usarla con el endpoint /shares/ de Graph API."""
    import base64
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return "u!" + b64


def listar_archivos_share(url_compartida: str) -> list[dict]:
    """Lista los archivos dentro de una carpeta compartida de OneDrive/SharePoint
    a partir de su link de 'Compartir' (los que empiezan con .../:f:/...)."""
    if not credenciales_configuradas():
        return []
    try:
        share_id = _encode_share_url(url_compartida)
        resp = requests.get(
            f"{GRAPH_URL}/shares/{share_id}/driveItem/children",
            headers=_headers(), timeout=20,
        )
        resp.raise_for_status()
        return resp.json().get("value", [])
    except Exception:
        return []


def descargar_item(drive_id: str, item_id: str) -> bytes:
    """Descarga el contenido binario de un archivo por su driveId/itemId."""
    resp = requests.get(
        f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/content",
        headers=_headers(), timeout=60,
    )
    resp.raise_for_status()
    return resp.content


def descargar_excel_base(url_compartida: str):
    """Busca el primer archivo Excel dentro de la carpeta compartida y lo
    descarga. Retorna (contenido_bytes, nombre_archivo, fecha_modificado) o
    (None, mensaje_error, None) si algo falla."""
    if not credenciales_configuradas():
        return None, "Credenciales de Graph API no configuradas.", None
    try:
        items = listar_archivos_share(url_compartida)
        if not items:
            return None, "No se encontraron archivos en la carpeta compartida.", None
        for it in items:
            nombre = it.get("name", "")
            if it.get("folder") or not nombre.lower().endswith((".xlsx", ".xls")):
                continue
            drive_id = (it.get("parentReference") or {}).get("driveId")
            item_id = it.get("id")
            if not (drive_id and item_id):
                continue
            contenido = descargar_item(drive_id, item_id)
            fecha = it.get("lastModifiedDateTime", "")[:16].replace("T", " ")
            return contenido, nombre, fecha
        return None, "No se encontró ningún archivo .xlsx en esa carpeta.", None
    except Exception as e:
        return None, f"Error al descargar: {e}", None
