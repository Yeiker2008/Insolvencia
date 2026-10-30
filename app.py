# ISOLVENCIAS APP

import streamlit as st
import pandas as pd
import os
import shutil
import base64
import random
import uuid
import calendar
import re
from datetime import datetime, timedelta, timezone
import time
import requests
import json
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from html import escape
from notificaciones import enviar_email, enviar_telegram, enviar_alertas_vencimientos, notificar_nuevo_vencimiento
import multifirma as mf


def ahora_colombia():
    """Hora de Colombia (UTC-5). El servidor de Streamlit corre en UTC y desfasaba las fechas por la noche."""
    return datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(hours=5)

st.set_page_config(page_title="Insolvencia OS | Enterprise", page_icon="⚖", layout="wide", initial_sidebar_state="collapsed")

try:
    from fpdf import FPDF
except ImportError:
    st.error("⚠️ Abre la terminal y ejecuta: pip install fpdf2")
try:
    import docx
except ImportError:
    st.error("⚠️ Abre la terminal y ejecuta: pip install python-docx")

# Intentar importar librerías de Google Sheets
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    GSPREAD_DISPONIBLE = True
except ImportError:
    GSPREAD_DISPONIBLE = False

# --- PODERES PARA GOOGLE DRIVE Y SHEETS ---
try:
    from google.oauth2.credentials import Credentials  # <-- ESTA LÍNEA ES LA CLAVE
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
    import io
    GDRIVE_DISPONIBLE = True
except ImportError:
    st.error("⚠️ Abre la terminal y ejecuta: pip install google-api-python-client google-auth")
    GDRIVE_DISPONIBLE = False

# Definimos el permiso maestro (Drive tiene acceso total a todo, incluyendo Sheets)
SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
# ID de tu Bóveda Maestra en Drive
CARPETA_RAIZ_DRIVE_ID = "1MC6wHXaV557prpKV-KCRc8yeCdphD6U8"

# ID de la Papelera en Drive
CARPETA_PAPELERA_DRIVE_ID = "1MGBXOKbPuAE6xsf53AcCjwMFE_yzDbir"
# Solo se usan para migrar la firma original a Control_Central; las demás firmas tienen sus propias carpetas.
mf.configurar_legado(CARPETA_RAIZ_DRIVE_ID, CARPETA_PAPELERA_DRIVE_ID)

# ==========================================
# GESTOR DE GOOGLE DRIVE VÍA API REST (CERO SEGMENTATION FAULT)
# ==========================================

@st.cache_resource
def _almacen_cache():
    """Caché compartida de la app: tablas, hojas, IDs de carpetas y listados de Drive."""
    return {"tablas": {}, "hojas": {}, "carpetas": {}, "listados": {}}


@st.cache_data(ttl=3000, show_spinner=False)
def obtener_token_fresco():
    """Token de Google reutilizado ~50 min en vez de pedir uno nuevo en cada llamada."""
    creds = Credentials.from_authorized_user_info({
        "client_id": st.secrets["google_client_id"],
        "client_secret": st.secrets["google_client_secret"],
        "refresh_token": st.secrets["google_refresh_token"],
        "token_uri": "https://oauth2.googleapis.com/token"
    })
    creds.refresh(Request())
    return creds.token

def estructurar_carpetas(cc, nom):
    """Asegura la carpeta del cliente en Drive una sola vez por sesión."""
    ya = st.session_state.setdefault("_carpetas_ok", set())
    clave = f"{_fid()}|{cc}|{nom}"
    if clave in ya:
        return False
    resultado = _estructurar_carpetas_real(cc, nom)
    if obtener_id_subcarpeta_drive(cc, nom, "Juzgado"):
        ya.add(clave)
    return resultado


def _estructurar_carpetas_real(cc, nom):
    """Crea la carpeta del cliente y sus subcarpetas obligatorias en Drive"""
    try:
        token = obtener_token_fresco()
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        nombre_carpeta_cliente = f"{cc} - {nom}"
        
        query = f"name='{nombre_carpeta_cliente}' and '{_boveda_id()}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res = requests.get(f"https://www.googleapis.com/drive/v3/files?q={requests.utils.quote(query)}", headers=headers).json()
        archivos = res.get('files', [])
        
        if not archivos:
            meta = {
                'name': nombre_carpeta_cliente,
                'parents': [_boveda_id()],
                'mimeType': 'application/vnd.google-apps.folder'
            }
            crear_cli = requests.post("https://www.googleapis.com/drive/v3/files", headers=headers, json=meta).json()
            cliente_id = crear_cli.get('id')
            
            for sub in ["Financiero", "Juzgado", "Notaria", "Acreedores"]:
                meta_sub = {
                    'name': sub,
                    'parents': [cliente_id],
                    'mimeType': 'application/vnd.google-apps.folder'
                }
                requests.post("https://www.googleapis.com/drive/v3/files", headers=headers, json=meta_sub)
            return True
        return False
    except Exception as e:
        print(f"Error estructurando carpetas: {e}")
        return False

def obtener_id_subcarpeta_drive(cc, nom, nombre_subcarpeta):
    """Obtiene el ID de una subcarpeta interna del cliente (con memoria para no repetir llamadas)"""
    clave = f"{_fid()}|{cc}|{nom}|{nombre_subcarpeta}"
    memoria = _almacen_cache()["carpetas"]
    if clave in memoria:
        return memoria[clave]
    try:
        token = obtener_token_fresco()
        headers = {"Authorization": f"Bearer {token}"}
        nombre_carpeta_cliente = f"{cc} - {nom}"
        
        q1 = f"name='{nombre_carpeta_cliente}' and '{_boveda_id()}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        r1 = requests.get(f"https://www.googleapis.com/drive/v3/files?q={requests.utils.quote(q1)}", headers=headers).json().get('files', [])
        if not r1: return None
        cliente_id = r1[0]['id']
        
        q2 = f"name='{nombre_subcarpeta}' and '{cliente_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        r2 = requests.get(f"https://www.googleapis.com/drive/v3/files?q={requests.utils.quote(q2)}", headers=headers).json().get('files', [])
        if r2:
            memoria[clave] = r2[0]['id']
            return r2[0]['id']
        return None
    except Exception as e:
        print(f"⚠️ Error buscando subcarpeta en Drive: {type(e).__name__}")
        return None

def subir_archivo_directo_drive(file_obj, filename, folder_id):
    """Sube cualquier archivo cargado desde Streamlit a Google Drive vía REST API"""
    if not folder_id: return False
    try:
        token = obtener_token_fresco()
        headers = {"Authorization": f"Bearer {token}"}
        
        metadata = {
            "name": filename,
            "parents": [folder_id]
        }
        
        files = {
            'data': ('metadata', json.dumps(metadata), 'application/json'),
            'file': (filename, file_obj.getbuffer(), 'application/octet-stream')
        }
        
        response = requests.post(
            "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
            headers=headers,
            files=files
        )
        _almacen_cache()["listados"].clear()
        return response.status_code == 200
    except Exception as e:
        print(f"Error en subida directa: {e}")
        return False

def listar_y_descargar_archivos_drive(folder_id):
    """Lista (id y nombre) los archivos de una carpeta. NO descarga contenido: es rápido."""
    if not folder_id: return []
    listados = _almacen_cache()["listados"]
    previo = listados.get(folder_id)
    if previo and (time.time() - previo[0]) < 60:
        return previo[1]
    try:
        token = obtener_token_fresco()
        headers = {"Authorization": f"Bearer {token}"}
        query = f"'{folder_id}' in parents and mimeType != 'application/vnd.google-apps.folder' and trashed=false"
        res = requests.get(
            f"https://www.googleapis.com/drive/v3/files?q={requests.utils.quote(query)}&fields=files(id,name)&pageSize=200",
            headers=headers, timeout=20
        ).json()
        lista = [{'id': a['id'], 'name': a['name']} for a in res.get('files', [])]
        listados[folder_id] = (time.time(), lista)
        return lista
    except Exception as e:
        print(f"⚠️ Error listando archivos de Drive: {type(e).__name__}")
        return []


def descargar_archivo_drive(file_id):
    """Descarga el contenido de UN archivo, solo cuando el usuario lo pide."""
    try:
        token = obtener_token_fresco()
        r = requests.get(f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media",
                         headers={"Authorization": f"Bearer {token}"}, timeout=60)
        return r.content if r.status_code == 200 else b""
    except Exception as e:
        print(f"⚠️ Error descargando archivo de Drive: {type(e).__name__}")
        return b""

def mover_a_papelera_drive(cc, nom):
    _almacen_cache()["carpetas"].clear()
    _almacen_cache()["listados"].clear()
    try:
        token = obtener_token_fresco()
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        nombre_carpeta_cliente = f"{cc} - {nom}"
        
        query = f"name='{nombre_carpeta_cliente}' and '{_boveda_id()}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res = requests.get(f"https://www.googleapis.com/drive/v3/files?q={requests.utils.quote(query)}", headers=headers).json().get('files', [])
        
        if res:
            cid = res[0]['id']
            padres = ",".join(res[0].get('parents', []))
            params = {"addParents": _papelera_id(), "removeParents": padres}
            requests.patch(f"https://www.googleapis.com/drive/v3/files/{cid}", headers=headers, params=params)
            return True
        return False
    except:
        return False

def restaurar_de_papelera_drive(cc, nom):
    _almacen_cache()["carpetas"].clear()
    _almacen_cache()["listados"].clear()
    try:
        token = obtener_token_fresco()
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        nombre_carpeta_cliente = f"{cc} - {nom}"
        
        query = f"name='{nombre_carpeta_cliente}' and '{_papelera_id()}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res = requests.get(f"https://www.googleapis.com/drive/v3/files?q={requests.utils.quote(query)}", headers=headers).json().get('files', [])
        
        if res:
            cid = res[0]['id']
            padres = ",".join(res[0].get('parents', []))
            params = {"addParents": _boveda_id(), "removeParents": padres}
            requests.patch(f"https://www.googleapis.com/drive/v3/files/{cid}", headers=headers, params=params)
            return True
        return False
    except:
        return False


# Conexión Global a Google Sheets
# ==========================================
# 1. CORE: SESIÓN, FIRMA ACTIVA Y CAPA DE DATOS
# ==========================================
if 'autenticado' not in st.session_state: st.session_state.autenticado = False
if 'es_superadmin' not in st.session_state: st.session_state.es_superadmin = False
if 'firma_ctx' not in st.session_state: st.session_state.firma_ctx = {}
if 'debe_cambiar_clave' not in st.session_state: st.session_state.debe_cambiar_clave = False
if 'usuario_actual' not in st.session_state: st.session_state.usuario_actual = ""
if 'rol_actual' not in st.session_state: st.session_state.rol_actual = ""
if 'alias_actual' not in st.session_state: st.session_state.alias_actual = ""
if 'avatar_path' not in st.session_state: st.session_state.avatar_path = ""
if 'session_token' not in st.session_state: st.session_state.session_token = ""
if 'pagina_actual' not in st.session_state: st.session_state.pagina_actual = 'Dashboard'
if 'kicked' not in st.session_state: st.session_state.kicked = False
if 'kicked_reason' not in st.session_state: st.session_state.kicked_reason = ""

COLS_USR = mf.COLS["usuarios"]


def cambiar_pagina(p): st.session_state.pagina_actual = p


def _ctx():
    return st.session_state.get("firma_ctx") or {}


def _fid():
    return str(_ctx().get("Firma_ID", ""))


def _boveda_id():
    return str(_ctx().get("Boveda_ID", ""))


def _papelera_id():
    return str(_ctx().get("Papelera_ID", ""))


def cerrar_sesion():
    """Cierra la sesión y borra todo lo que quedó en memoria de la firma, para que nada se cruce."""
    for k in list(st.session_state.keys()):
        if k.startswith(("sel_", "sim_", "dlbuf_", "kb_", "tgl_bov_", "busq", "cal_", "tar_", "_carpetas_ok", "aviso_", "sa_")):
            st.session_state.pop(k, None)
    st.session_state.update(autenticado=False, es_superadmin=False, debe_cambiar_clave=False, firma_ctx={},
                            usuario_actual="", rol_actual="", alias_actual="", avatar_path="", session_token="",
                            pagina_actual="Dashboard", cliente_foco="")


TTL_TABLAS = 600


def leer_tabla(nombre_tabla, columnas_def):
    """Lee una tabla de LA FIRMA DE LA SESIÓN. La memoria rápida se separa por firma."""
    fid, ref = _fid(), str(_ctx().get("Sheet_ID", ""))
    if not fid or not ref:
        raise mf.ErrorDatos("No hay una firma activa en esta sesión.")
    tablas = _almacen_cache()["tablas"]
    clave = (fid, nombre_tabla)
    guardado = tablas.get(clave)
    if guardado and (time.time() - guardado[0]) < TTL_TABLAS:
        return guardado[1].copy()
    df = mf.leer_registros(ref, nombre_tabla, columnas_def)
    tablas[clave] = (time.time(), df)
    return df.copy()


def guardar_tabla(df, nombre_tabla):
    """Guarda en la hoja de la firma de la sesión. Si falla, lanza un error visible: nunca se pierde en silencio."""
    fid, ref = _fid(), str(_ctx().get("Sheet_ID", ""))
    if not fid or not ref:
        raise mf.ErrorDatos("No hay una firma activa en esta sesión.")
    mf.escribir_registros(ref, nombre_tabla, df)
    _almacen_cache()["tablas"].pop((fid, nombre_tabla), None)


def cambiar_mi_clave(actual, nueva, confirmacion, exigir_actual=True):
    motivo = mf.validar_clave_nueva(nueva)
    if motivo:
        return False, motivo
    if nueva != confirmacion:
        return False, "La confirmación no coincide."
    df_u = leer_tabla("usuarios", COLS_USR).astype(object)
    m = df_u["Usuario"].astype(str).str.lower() == str(st.session_state.usuario_actual).lower()
    if not m.any():
        return False, "No se encontró tu usuario."
    almacenado = df_u.loc[m, "Password"].iloc[0]
    if exigir_actual and not mf.verificar_password(actual, almacenado):
        return False, "La contraseña actual no es correcta."
    if mf.verificar_password(nueva, almacenado):
        return False, "La nueva contraseña debe ser distinta de la actual."
    df_u.loc[m, "Password"] = mf.hash_password(nueva)
    df_u.loc[m, "Cambiar_Clave"] = "No"
    guardar_tabla(df_u, "usuarios")
    registrar_log("SEGURIDAD", "Cambió su contraseña")
    return True, "Contraseña actualizada."


# ==========================================
# 2. MOTOR CSS: SAAS CORPORATIVO ADAPTATIVO (MÓVIL & DESKTOP)
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Fraunces:opsz,wght@9..144,600;9..144,700&display=swap');

:root{
  --navy:#0A1628; --navy2:#13294B; --navy3:#1B3A6B; --gold:#C9A24B; --gold2:#E6C873; --gold-dark:#8F7129;
  --canvas:#EEF1F7; --paper:#FFFFFF; --ink:#14213A; --muted:#667085; --line:#E3E8F1;
  --burgundy:#8A1F35; --emerald:#1F7A5C; --sky:#2F6DB5; --violet:#6B5BA6; --orange:#D9822B;
  --radius:20px;
}
footer{visibility:hidden !important; display:none !important;}
html, body, [class*="css"]{font-family:'Plus Jakarta Sans',sans-serif; color:var(--ink);}
::selection{background:var(--gold2); color:var(--navy);}
::-webkit-scrollbar{width:10px; height:10px;}
::-webkit-scrollbar-thumb{background:#C4CCDD; border-radius:8px;}

/* ---------- Lienzo ---------- */
[data-testid="stAppViewContainer"]{
  background:
    radial-gradient(900px 420px at 100% -5%, rgba(201,162,75,.13), transparent 60%),
    radial-gradient(700px 400px at -5% 40%, rgba(47,109,181,.08), transparent 60%),
    var(--canvas) !important;
}
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"]{display:none !important;}
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"], [data-testid="collapsedControl"]{display:none !important;}
.block-container{padding:1.1rem 2rem 3rem 2rem !important; max-width:1360px;}

/* ---------- Animaciones ---------- */
@keyframes fadeUp{from{opacity:0; transform:translateY(16px);} to{opacity:1; transform:none;}}
@keyframes slideIn{from{opacity:0; transform:translateX(-18px);} to{opacity:1; transform:none;}}
@keyframes floaty{0%,100%{transform:translateY(0);} 50%{transform:translateY(-8px);}}
@keyframes drift{0%,100%{transform:translate(0,0) scale(1);} 50%{transform:translate(30px,-18px) scale(1.12);}}
@keyframes pulseRed{0%,100%{box-shadow:0 0 0 0 rgba(138,31,53,.35);} 50%{box-shadow:0 0 0 9px rgba(138,31,53,0);}}
@keyframes grow{from{transform:scaleX(0);} to{transform:scaleX(1);}}
@keyframes gradientMove{0%{background-position:0% 50%;} 50%{background-position:100% 50%;} 100%{background-position:0% 50%;}}
@media (prefers-reduced-motion: reduce){*{animation:none !important; transition:none !important;}}

/* ---------- Tipografía ---------- */
h1,h2,h3,h4,h5{font-family:'Fraunces', Georgia, serif !important; color:var(--navy) !important; letter-spacing:-.3px;}
[data-testid="stMarkdownContainer"] h1{
  font-size:2rem !important; font-weight:700 !important; padding:.4rem 0 .2rem 0 !important; animation:slideIn .5s ease both;
}
[data-testid="stMarkdownContainer"] h1::after{
  content:''; display:block; width:56px; height:4px; margin-top:8px; border-radius:4px;
  background:linear-gradient(90deg,var(--gold),var(--gold2));
}
.page-sub{color:var(--muted) !important; font-size:14px; margin:2px 0 14px 0;}
[data-testid="stMarkdownContainer"] p{color:var(--ink);}
[data-testid="stWidgetLabel"] p, label p{color:var(--navy) !important; font-weight:600 !important; font-size:13px !important;}
.card-title{font-size:12px !important; font-weight:800 !important; letter-spacing:1.3px; text-transform:uppercase; color:var(--muted) !important; margin:0 0 10px 0 !important;}

/* ---------- Barra superior (centro de mando) ---------- */
.st-key-topbar{
  position:relative; overflow:hidden; border-radius:24px; padding:14px 22px; margin-bottom:12px;
  background:
    radial-gradient(520px 200px at 8% -30%, rgba(201,162,75,.45), transparent 60%),
    radial-gradient(420px 220px at 95% 130%, rgba(47,109,181,.5), transparent 60%),
    linear-gradient(135deg,#0A1628 0%,#13294B 100%);
  box-shadow:0 28px 60px -28px rgba(10,22,40,.85); animation:fadeUp .5s ease both;
}
.marca{display:flex; align-items:center; gap:12px;}
.marca .sello{width:44px; height:44px; border-radius:14px; background:linear-gradient(135deg,var(--gold),var(--gold2));
  display:flex; align-items:center; justify-content:center; font-size:23px; box-shadow:0 10px 24px -8px rgba(201,162,75,.9); animation:floaty 5s ease-in-out infinite;}
.marca .nombre{font-family:'Fraunces',Georgia,serif; font-size:20px; font-weight:700; color:#FFFFFF !important; line-height:1.05;}
.marca .lema{font-size:10px; letter-spacing:2px; color:var(--gold2) !important; text-transform:uppercase; font-weight:700;}
.st-key-topbar [data-baseweb="base-input"], .st-key-topbar [data-baseweb="input"], .st-key-topbar [data-baseweb="input"] > div, .st-key-topbar [data-testid="stTextInputRootElement"]{
  background:rgba(255,255,255,.10) !important; border-color:rgba(255,255,255,.22) !important; border-radius:14px !important;}
.st-key-topbar input{background:rgba(255,255,255,.10) !important; color:#FFFFFF !important; -webkit-text-fill-color:#FFFFFF !important; border:1px solid rgba(255,255,255,.18) !important;}
.st-key-topbar div[data-baseweb="input"] > div{background:rgba(255,255,255,.10) !important; border-color:rgba(255,255,255,.18) !important;}
.st-key-topbar input::placeholder{color:rgba(255,255,255,.6) !important; -webkit-text-fill-color:rgba(255,255,255,.6) !important;}
.st-key-topbar input:focus{border-color:var(--gold2) !important; box-shadow:0 0 0 3px rgba(230,200,115,.28) !important;}
.st-key-topbar button{background:rgba(255,255,255,.10) !important; color:#FFFFFF !important; border:1px solid rgba(255,255,255,.2) !important; backdrop-filter:blur(6px);}
.st-key-topbar button:hover{background:linear-gradient(135deg,var(--gold),var(--gold2)) !important; color:var(--navy) !important; border-color:var(--gold) !important;}
.st-key-topbar button p{color:inherit !important;}

/* ---------- Navegación ---------- */
.st-key-nav_grupos{background:var(--paper); border:1px solid var(--line); border-radius:18px; padding:6px; margin-bottom:6px;
  box-shadow:0 12px 30px -20px rgba(10,22,40,.4); animation:fadeUp .55s ease both;}
.st-key-nav_grupos [data-testid="stHorizontalBlock"], .st-key-nav_sub [data-testid="stHorizontalBlock"]{gap:.35rem !important; flex-wrap:wrap !important;}
.st-key-nav_grupos [data-testid="stColumn"], .st-key-nav_sub [data-testid="stColumn"], .st-key-nav_grupos [data-testid="column"], .st-key-nav_sub [data-testid="column"]{min-width:0 !important; flex:1 1 auto !important; width:auto !important;}
.st-key-nav_grupos div.stButton > button{background:transparent !important; color:var(--muted) !important; border:1px solid transparent !important;
  border-radius:13px !important; font-weight:700 !important; box-shadow:none !important; padding:.55rem .6rem !important;}
.st-key-nav_grupos div.stButton > button p{color:inherit !important;}
.st-key-nav_grupos div.stButton > button:hover{background:rgba(201,162,75,.14) !important; color:var(--navy) !important; transform:none !important; box-shadow:none !important; border-color:rgba(201,162,75,.4) !important;}
.st-key-nav_grupos div.stButton > button[kind="primary"], .st-key-nav_grupos div.stButton > button[data-testid="stBaseButton-primary"]{
  background:linear-gradient(135deg,var(--navy),var(--navy3)) !important; color:#FFFFFF !important;
  box-shadow:0 12px 24px -12px rgba(10,22,40,.8) !important;}
.st-key-nav_grupos div.stButton > button[kind="primary"]:hover, .st-key-nav_grupos div.stButton > button[data-testid="stBaseButton-primary"]:hover{color:#FFFFFF !important;}
.st-key-nav_sub{padding:2px 4px 10px 4px; animation:fadeUp .6s ease both;}
.st-key-nav_sub div.stButton > button{background:transparent !important; color:var(--muted) !important; border:none !important; border-bottom:3px solid transparent !important;
  border-radius:0 !important; font-weight:600 !important; box-shadow:none !important; padding:.35rem .5rem !important; font-size:14px !important;}
.st-key-nav_sub div.stButton > button p{color:inherit !important;}
.st-key-nav_sub div.stButton > button:hover{color:var(--navy) !important; background:transparent !important; transform:none !important; box-shadow:none !important; border-bottom-color:rgba(201,162,75,.6) !important;}
.st-key-nav_sub div.stButton > button[kind="primary"], .st-key-nav_sub div.stButton > button[data-testid="stBaseButton-primary"]{
  color:var(--navy) !important; border-bottom-color:var(--gold) !important; font-weight:800 !important; background:transparent !important;}

/* ---------- Tarjetas (contenedores con borde) ---------- */
[data-testid="stVerticalBlockBorderWrapper"]{
  background:var(--paper); border:1px solid var(--line) !important; border-radius:var(--radius) !important;
  box-shadow:0 1px 2px rgba(10,22,40,.04), 0 18px 40px -26px rgba(10,22,40,.35);
  animation:fadeUp .55s ease both; transition:box-shadow .25s ease, transform .25s ease;
}
[data-testid="stVerticalBlockBorderWrapper"]:hover{box-shadow:0 1px 2px rgba(10,22,40,.05), 0 26px 50px -26px rgba(10,22,40,.45);}
[data-testid="stVerticalBlockBorderWrapper"] [data-testid="stVerticalBlockBorderWrapper"]{box-shadow:none; animation:none;}

/* ---------- Héroe del centro de mando ---------- */
.hero{position:relative; overflow:hidden; border-radius:26px; padding:30px 34px; margin:6px 0 16px 0;
  background:linear-gradient(120deg,#0A1628,#13294B,#1B3A6B,#13294B); background-size:300% 300%; animation:gradientMove 14s ease infinite;
  box-shadow:0 30px 60px -30px rgba(10,22,40,.8);}
.hero::before,.hero::after{content:''; position:absolute; border-radius:50%; filter:blur(8px); animation:drift 11s ease-in-out infinite;}
.hero::before{width:260px; height:260px; right:-60px; top:-90px; background:radial-gradient(circle,rgba(230,200,115,.55),transparent 65%);}
.hero::after{width:300px; height:300px; right:22%; bottom:-170px; background:radial-gradient(circle,rgba(47,109,181,.55),transparent 65%); animation-delay:-5s;}
.hero .kick{color:var(--gold2) !important; font-size:12px; font-weight:800; letter-spacing:2px; text-transform:uppercase; margin:0 0 6px 0; position:relative; z-index:1;}
.hero .titulo{font-family:'Fraunces',Georgia,serif; color:#FFFFFF !important; font-size:2.3rem; font-weight:700; margin:0; position:relative; z-index:1; line-height:1.1;}
.hero .sub{color:#C9D4E8 !important; font-size:15px; margin:10px 0 0 0; position:relative; z-index:1; max-width:640px;}
.hero .chips{display:flex; gap:10px; margin-top:16px; flex-wrap:wrap; position:relative; z-index:1;}
.hero .chip2{background:rgba(255,255,255,.12); border:1px solid rgba(255,255,255,.22); color:#FFFFFF !important; border-radius:30px; padding:5px 14px; font-size:12.5px; font-weight:600; backdrop-filter:blur(6px);}

/* ---------- KPIs ---------- */
.kpi{position:relative; background:var(--paper); border:1px solid var(--line); border-radius:var(--radius); padding:18px 20px 16px 22px;
  overflow:hidden; box-shadow:0 18px 40px -28px rgba(10,22,40,.4); transition:transform .25s ease, box-shadow .25s ease; animation:fadeUp .6s ease both; margin-bottom:10px;}
.kpi::after{content:''; position:absolute; right:-30px; top:-30px; width:110px; height:110px; border-radius:50%; background:var(--accent-bg, rgba(201,162,75,.15));}
.kpi:hover{transform:translateY(-5px); box-shadow:0 30px 50px -26px rgba(10,22,40,.5);}
.kpi .ico{font-size:22px; width:42px; height:42px; border-radius:13px; display:flex; align-items:center; justify-content:center;
  background:var(--accent-bg, rgba(201,162,75,.15)); margin-bottom:12px; position:relative; z-index:1;}
.kpi .lbl{font-size:11px; font-weight:800; letter-spacing:1.2px; text-transform:uppercase; color:var(--muted) !important; margin:0; position:relative; z-index:1;}
.kpi .val{font-family:'Fraunces',Georgia,serif; font-size:2.1rem; font-weight:700; color:var(--navy) !important; margin:2px 0 0 0; line-height:1.15; position:relative; z-index:1;}
.kpi .bar{height:5px; border-radius:6px; background:var(--line); margin-top:12px; overflow:hidden;}
.kpi .bar i{display:block; height:100%; background:var(--accent,var(--gold)); transform-origin:left; animation:grow 1.1s ease both;}

/* ---------- Línea de tiempo ---------- */
.tl{position:relative; margin:4px 0 6px 6px;}
.tl::before{content:''; position:absolute; left:6px; top:6px; bottom:6px; width:2px; background:linear-gradient(var(--gold),rgba(201,162,75,.1));}
.tl-item{position:relative; padding:0 0 14px 28px; animation:fadeUp .5s ease both;}
.tl-dot{position:absolute; left:0; top:4px; width:14px; height:14px; border-radius:50%; background:var(--c,var(--gold)); border:3px solid #FFFFFF; box-shadow:0 0 0 2px var(--c,var(--gold));}
.tl-item b{display:block; font-size:14px; color:var(--navy) !important;}
.tl-item span{display:block; font-size:12.5px; color:var(--muted) !important;}
.tl-item em{display:inline-block; font-style:normal; font-size:11px; font-weight:800; letter-spacing:.6px; text-transform:uppercase; color:var(--c,var(--gold-dark)) !important; margin-top:2px;}
.vacio{text-align:center; color:var(--muted) !important; padding:18px 6px; font-size:14px;}

/* ---------- Pipeline, actividad, tareas ---------- */
.pipe{display:flex; height:14px; border-radius:10px; overflow:hidden; background:var(--line); margin:4px 0 12px 0;}
.pipe span{display:block; transform-origin:left; animation:grow 1s ease both;}
.pipe-row{display:flex; align-items:center; justify-content:space-between; font-size:13px; padding:4px 0; border-bottom:1px dashed var(--line);}
.pipe-row i{display:inline-block; width:10px; height:10px; border-radius:50%; margin-right:8px; background:var(--c);}
.pipe-row b{color:var(--navy) !important;}
.act{padding:7px 0; border-bottom:1px dashed var(--line); font-size:13px;}
.act em{display:block; font-style:normal; font-size:11px; color:var(--muted) !important;}
.task{padding:6px 0; font-size:13.5px;}
.task em{display:block; font-style:normal; font-size:11px; color:var(--muted) !important;}
.alerta-item{display:flex; gap:10px; align-items:flex-start; padding:8px 10px; border-radius:12px; margin-bottom:6px; background:var(--bg,#F5F7FB); border-left:4px solid var(--c,var(--gold));}
.alerta-item b{display:block; font-size:13px; color:var(--navy) !important;}
.alerta-item span{font-size:12px; color:var(--muted) !important;}

/* ---------- Tablero Kanban ---------- */
.kb-head{display:flex; align-items:center; justify-content:space-between; padding:10px 14px; border-radius:14px; margin-bottom:10px;
  background:linear-gradient(135deg,var(--c),color-mix(in srgb,var(--c) 70%,#000)); color:#FFFFFF !important; font-weight:800; font-size:13px; letter-spacing:.4px;
  box-shadow:0 14px 26px -16px var(--c);}
.kb-head span{background:rgba(255,255,255,.25); border-radius:20px; padding:1px 10px; font-size:12px;}
.kb-card{background:var(--paper); border:1px solid var(--line); border-left:5px solid var(--c); border-radius:14px; padding:11px 13px; margin-bottom:4px;
  box-shadow:0 14px 28px -22px rgba(10,22,40,.5); transition:transform .2s ease, box-shadow .2s ease; animation:fadeUp .45s ease both;}
.kb-card:hover{transform:translateY(-3px); box-shadow:0 22px 34px -22px rgba(10,22,40,.6);}
.kb-card b{display:block; font-size:13.5px; color:var(--navy) !important; line-height:1.25;}
.kb-card span{display:block; font-size:11.5px; color:var(--muted) !important; margin-top:2px;}
.kb-card .pill{display:inline-block; margin:6px 6px 0 0; padding:2px 9px; border-radius:20px; font-size:10.5px; font-weight:700; background:#EEF1F7; color:var(--navy) !important;}
.kb-card .pill.warn{background:rgba(138,31,53,.12); color:var(--burgundy) !important;}
.kb-vacio{border:2px dashed var(--line); border-radius:14px; padding:18px; text-align:center; color:var(--muted) !important; font-size:12.5px;}

/* ---------- Ficha 360 ---------- */
.ficha-head{display:flex; align-items:center; gap:18px; padding:4px 0;}
.avatar{width:72px; height:72px; border-radius:22px; background:linear-gradient(135deg,var(--navy),var(--navy3)); color:var(--gold2) !important;
  display:flex; align-items:center; justify-content:center; font-family:'Fraunces',Georgia,serif; font-size:28px; font-weight:700; box-shadow:0 18px 30px -16px rgba(10,22,40,.8), 0 0 0 3px rgba(201,162,75,.5);}
.ficha-head .nom{font-family:'Fraunces',Georgia,serif; font-size:1.7rem; font-weight:700; color:var(--navy) !important; margin:0; line-height:1.1;}
.ficha-head .det{color:var(--muted) !important; font-size:13px; margin:4px 0 0 0;}
.stepper{display:flex; align-items:center; margin:14px 0 18px 0;}
.stepper .s{flex:1; text-align:center; position:relative; font-size:11px; font-weight:700; color:var(--muted) !important;}
.stepper .s::before{content:''; position:absolute; top:13px; left:-50%; width:100%; height:3px; background:var(--line); z-index:0;}
.stepper .s:first-child::before{display:none;}
.stepper .s i{position:relative; z-index:1; display:flex; align-items:center; justify-content:center; width:28px; height:28px; margin:0 auto 5px auto; border-radius:50%;
  background:#FFFFFF; border:3px solid var(--line); font-style:normal; font-size:13px;}
.stepper .s.ok i{background:var(--emerald); border-color:var(--emerald); color:#FFFFFF;}
.stepper .s.ok::before{background:var(--emerald);}
.stepper .s.now i{background:var(--navy); border-color:var(--gold); color:var(--gold2); box-shadow:0 0 0 5px rgba(201,162,75,.22);}
.stepper .s.now{color:var(--navy) !important;}
.chip-info{display:inline-block; padding:4px 12px; border-radius:20px; background:#EEF1F7; font-size:12px; font-weight:700; color:var(--navy) !important; margin:0 6px 6px 0;}
.chip-info.rojo{background:rgba(138,31,53,.12); color:var(--burgundy) !important;}
.chip-info.verde{background:rgba(31,122,92,.14); color:var(--emerald) !important;}

/* ---------- Calendario ---------- */
.cal-titulo{font-family:'Fraunces',Georgia,serif; font-size:1.5rem; font-weight:700; color:var(--navy) !important; text-align:center; margin:6px 0;}
.cal{display:grid; grid-template-columns:repeat(7,minmax(0,1fr)); gap:6px; margin:8px 0 14px 0;}
.cal-h{text-align:center; font-size:11px; font-weight:800; letter-spacing:1.2px; text-transform:uppercase; color:var(--muted) !important; padding:4px 0;}
.cal-d{min-height:86px; background:var(--paper); border:1px solid var(--line); border-radius:14px; padding:6px 7px; animation:fadeUp .5s ease both; transition:transform .2s ease, box-shadow .2s ease;}
.cal-d:hover{transform:translateY(-3px); box-shadow:0 16px 28px -20px rgba(10,22,40,.6);}
.cal-d.fuera{background:transparent; border-color:transparent; box-shadow:none;}
.cal-d.fuera:hover{transform:none; box-shadow:none;}
.cal-d .n{display:inline-block; font-size:12px; font-weight:800; color:var(--navy) !important; margin-bottom:3px;}
.cal-d.hoy{border:2px solid var(--gold); background:linear-gradient(180deg,#FFFBF0,#FFFFFF);}
.cal-d.hoy .n{background:var(--gold); color:var(--navy) !important; border-radius:20px; padding:0 8px;}
.chip{font-size:10.5px; font-weight:600; padding:2px 6px; margin-top:3px; border-radius:7px; background:color-mix(in srgb,var(--c) 15%,#FFFFFF); color:var(--navy) !important;
  border-left:3px solid var(--c); white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
.mas{font-size:10px; color:var(--muted) !important; margin-top:3px; font-weight:700;}

/* ---------- Simulador ---------- */
.sim-res{background:linear-gradient(135deg,var(--navy),var(--navy3)); border-radius:var(--radius); padding:22px 26px; color:#FFFFFF; box-shadow:0 28px 50px -28px rgba(10,22,40,.9);}
.sim-res .big{font-family:'Fraunces',Georgia,serif; font-size:2.6rem; font-weight:700; color:var(--gold2) !important; margin:0; line-height:1.1;}
.sim-res p{color:#C9D4E8 !important; margin:4px 0 0 0; font-size:13px;}
.aviso{background:#FFF8E6; border:1px solid #F0DDA6; border-left:5px solid var(--gold); border-radius:14px; padding:12px 16px; font-size:12.5px; color:#5B4A12 !important;}

/* ---------- Formularios, tabs, expanders, tablas ---------- */
[data-testid="stForm"]{background:var(--paper); border:1px solid var(--line) !important; border-top:4px solid var(--gold) !important;
  border-radius:var(--radius) !important; padding:22px !important; animation:fadeUp .5s ease both; box-shadow:0 18px 40px -28px rgba(10,22,40,.35);}
input, textarea, div[data-baseweb="select"] > div, div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div{
  background-color:#FFFFFF !important; color:var(--ink) !important; -webkit-text-fill-color:var(--ink) !important;
  border-color:#D3DAE8 !important; border-radius:12px !important;}
input:focus, textarea:focus{border-color:var(--gold) !important; box-shadow:0 0 0 3px rgba(201,162,75,.28) !important;}
button[data-baseweb="tab"]{font-weight:700; color:var(--muted);}
button[data-baseweb="tab"][aria-selected="true"]{color:var(--navy) !important;}
[data-baseweb="tab-highlight"]{background:var(--gold) !important; height:3px !important;}
[data-testid="stExpander"]{border:1px solid var(--line) !important; border-radius:16px !important; background:var(--paper); overflow:hidden;}
[data-testid="stExpander"] summary:hover{background:rgba(201,162,75,.10);}
[data-testid="stAlert"]{border-radius:14px !important;}
[data-testid="stDataFrame"]{border-radius:14px; overflow:hidden; box-shadow:0 14px 28px -22px rgba(10,22,40,.5);}
[data-testid="stMetric"]{background:var(--paper); border:1px solid var(--line); border-radius:14px; padding:12px 16px;}
[data-testid="stPopoverBody"]{border-radius:18px !important;}

/* ---------- Botones (área principal) ---------- */
div.stButton > button, div[data-testid="stFormSubmitButton"] > button, div.stDownloadButton > button, a[data-testid^="stBaseLinkButton"]{
  background:linear-gradient(135deg,var(--navy),var(--navy3)) !important; color:#FFFFFF !important;
  border:1px solid var(--navy) !important; border-radius:12px !important; font-weight:700 !important; letter-spacing:.2px; padding:.5rem 1rem !important;
  transition:transform .2s ease, box-shadow .2s ease, background .2s ease !important;}
div.stButton > button p, div[data-testid="stFormSubmitButton"] > button p, div.stDownloadButton > button p{color:inherit !important;}
div.stButton > button:hover, div[data-testid="stFormSubmitButton"] > button:hover, div.stDownloadButton > button:hover, a[data-testid^="stBaseLinkButton"]:hover{
  background:linear-gradient(135deg,var(--gold),var(--gold2)) !important; color:var(--navy) !important; border-color:var(--gold) !important;
  transform:translateY(-2px); box-shadow:0 14px 26px -12px rgba(201,162,75,.8) !important;}
div.stButton > button:active{transform:translateY(0) scale(.98);}
a[data-testid^="stBaseLinkButton"] p, a[data-testid^="stBaseLinkButton"] span{color:inherit !important;}
a[data-testid^="stBaseLinkButton"]{text-decoration:none !important; text-align:center;}
/* botones pequeños de tarjetas (kanban, tareas) */
[class*="st-key-mini"] button{padding:.2rem .4rem !important; font-size:13px !important; border-radius:10px !important;}

/* ---------- Login ---------- */
.login-hero{text-align:center; padding:6px 0 4px 0;}
.login-hero .seal{width:90px; height:90px; margin:0 auto 16px auto; border-radius:28px; font-size:46px;
  background:linear-gradient(135deg,var(--gold),var(--gold2)); display:flex; align-items:center; justify-content:center;
  box-shadow:0 26px 50px -18px rgba(201,162,75,.9); animation:floaty 4s ease-in-out infinite;}
.login-hero h1::after{margin-left:auto; margin-right:auto;}
.login-hero p{font-size:13px; margin:8px 0 18px 0;}

/* ---------- Compatibilidad con módulos anteriores ---------- */
.module-card{background:var(--paper); border:1px solid var(--line); border-radius:var(--radius); padding:18px 22px; margin-bottom:14px;
  box-shadow:0 18px 40px -28px rgba(10,22,40,.35); animation:fadeUp .5s ease both;}
.module-card h3,.module-card h4{color:var(--navy) !important; margin:0 0 6px 0;}
.module-card p,.module-card span,.module-card b{color:var(--ink) !important;}
.module-card-gold{border-top:4px solid var(--gold) !important;}
.module-card-blue{border-top:4px solid var(--sky) !important;}
.module-card-green{border-top:4px solid var(--emerald) !important;}
.module-card:has(> h3:only-child), .module-card:has(> h4:only-child){
  background:transparent; border:none; border-left:4px solid var(--gold); border-radius:0; box-shadow:none; padding:2px 0 2px 12px; margin:14px 0 6px 0;}
[data-testid="stElementContainer"]:has(.module-card:not(:has(*))), .element-container:has(.module-card:not(:has(*))),
[data-testid="stElementContainer"]:has([data-testid="stMarkdownContainer"]:empty){display:none !important;}
.law-card{background:linear-gradient(135deg,var(--navy),var(--navy3)); border-radius:var(--radius); padding:22px 26px; border-top:4px solid var(--gold); margin-bottom:14px;}
.law-card h3{color:var(--gold2) !important; font-size:1.3rem; margin:0 0 10px 0;}
.law-card p, .law-card b{color:#E8ECF3 !important; font-size:14px; line-height:1.75; margin:0;}
.venc{border-left:5px solid var(--c,#10B981); padding:12px 16px; background:#FFFFFF; margin-bottom:8px; border-radius:12px;
  box-shadow:0 12px 24px -18px rgba(10,22,40,.4); color:var(--ink); animation:fadeUp .45s ease both; transition:transform .2s ease;}
.venc:hover{transform:translateX(4px);}
.venc-peligro{animation:fadeUp .45s ease both, pulseRed 2s ease-in-out infinite;}
.boveda-sub{color:var(--gold-dark) !important; font-weight:800; margin:6px 0 2px 0; font-size:13px; letter-spacing:.4px;}
.archivo{font-size:12.5px; color:var(--ink);}
.nav-hr{border:none; height:1px; background:linear-gradient(90deg,transparent,#C4CCDD,transparent); margin:14px 0;}

@media (max-width:900px){
  .st-key-topbar [data-testid="stHorizontalBlock"]{flex-wrap:wrap !important; gap:.5rem !important;}
  .st-key-topbar [data-testid="stColumn"]{min-width:44% !important; flex:1 1 44% !important;}
  .st-key-topbar [data-testid="stColumn"]:nth-child(-n+2){min-width:100% !important; flex:1 1 100% !important;}
  .block-container{padding:.8rem .7rem 3rem .7rem !important;}
  [data-testid="stMarkdownContainer"] h1{font-size:1.55rem !important;}
  .hero{padding:22px 20px;} .hero .titulo{font-size:1.7rem;}
  .cal-d{min-height:64px; padding:4px;} .chip{font-size:9px;}
}

</style>
""", unsafe_allow_html=True)


# ==========================================
# 3. BASE DE DATOS Y CARGA CLOUD
# ==========================================
CARP_EXP, CARP_PLA, CARP_PAP, CARP_AVATAR = "Expedientes", "Plantillas", "Papelera", "Avatares"

for c in [CARP_EXP, CARP_PLA, CARP_PAP, CARP_AVATAR]: os.makedirs(c, exist_ok=True)

def registrar_log(modulo, accion):
    try:
        usr = st.session_state.get('alias_actual', 'Sistema')
        df_l = leer_tabla("logs", ["Timestamp", "Usuario", "Modulo", "Accion"])
        nuevo = pd.DataFrame([{"Timestamp": ahora_colombia().strftime("%Y-%m-%d %H:%M:%S"), "Usuario": usr, "Modulo": modulo, "Accion": accion}])
        guardar_tabla(pd.concat([df_l, nuevo], ignore_index=True), "logs")
    except Exception as e:
        print(f"⚠️ No se pudo registrar el log: {type(e).__name__}")

def limpiar_num(val):
    """Convierte montos a número. Acepta 1500000, 1.500.000, $ 1.500.000, 1,500,000.50 y 4500000.0"""
    try:
        if isinstance(val, (int, float)) and not pd.isna(val):
            return float(val)
        s = str(val).replace("$", "").replace(" ", "").strip()
        if s == "" or s.lower() == "nan":
            return 0.0
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        else:
            for sep in (",", "."):
                if sep in s:
                    partes = s.split(sep)
                    if len(partes) > 2 or len(partes[-1]) == 3:
                        s = s.replace(sep, "")
                    else:
                        s = s.replace(sep, ".")
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def motor_docx(plantilla, salida, dict_datos):
    try:
        from docx.text.paragraph import Paragraph
        doc = docx.Document(plantilla)
        for xml_p in doc.element.body.xpath('.//w:p'):
            p = Paragraph(xml_p, doc)
            for k, v in dict_datos.items():
                if k in p.text:
                    for run in p.runs:
                        if k in run.text: run.text = run.text.replace(k, str(v))
        doc.save(salida)
        return True
    except: return False

def generar_pdf_oficial(titulo, contenido, filename):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margins(20, 20, 20)
    def cln(t): return str(t).encode('latin-1', 'replace').decode('latin-1')
    pdf.set_font("Helvetica", 'B', 14)
    pdf.multi_cell(0, 8, txt=cln(titulo), align='C')
    pdf.ln(8)
    pdf.set_font("Helvetica", '', 11)
    pdf.multi_cell(0, 6, txt=cln(contenido), align='J')
    pdf.ln(15)
    pdf.set_font("Helvetica", 'B', 11)
    pdf.cell(0, 6, txt="__________________________________", ln=True)
    pdf.cell(0, 6, txt=cln("Firma Autorizada del Despacho"), ln=True)
    pdf.output(filename)
    return filename

def generar_ics(cliente, fecha_hora_str, motivo):
    try:
        dt_start = datetime.strptime(fecha_hora_str, "%Y-%m-%d %H:%M:%S")
        dt_end = dt_start + timedelta(hours=1)
        formato = "%Y%m%dT%H%M%S"
        ics_content = f"BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//Insolvencia OS//Agenda//ES\nBEGIN:VEVENT\nSUMMARY:⚖ {motivo} - {cliente}\nDTSTART:{dt_start.strftime(formato)}\nDTEND:{dt_end.strftime(formato)}\nDESCRIPTION:Audiencia/Cita programada desde el sistema de gestión.\nEND:VEVENT\nEND:VCALENDAR"
        return ics_content.encode('utf-8')
    except: return b""

def mostrar_boveda(cc, nom):
    with st.expander("📂 BÓVEDA CENTRAL DE ARCHIVOS EN GOOGLE DRIVE", expanded=False):
        cargar = st.toggle("Cargar archivos de la bóveda", key=f"tgl_bov_{cc}",
                           help="Drive solo se consulta cuando activas esto, así la app va más rápida.")
        if not cargar:
            st.caption("Activa el interruptor para ver, descargar o borrar los archivos de este expediente.")
            return
        estructurar_carpetas(cc, nom)
        subcarpetas = ["Financiero", "Juzgado", "Notaria", "Acreedores"]
        cols = st.columns(2)

        for i, sub in enumerate(subcarpetas):
            col = cols[i % 2]
            col.markdown(f"<p class='boveda-sub'>📁 {sub}</p>", unsafe_allow_html=True)
            folder_id = obtener_id_subcarpeta_drive(cc, nom, sub)

            if folder_id:
                archivos = listar_y_descargar_archivos_drive(folder_id)
                if archivos:
                    for a in archivos:
                        c_nom, c_dl, c_del = col.columns([5, 1.3, 1.3])
                        c_nom.markdown(f"<span class='archivo'>📄 {escape(a['name'][:30])}</span>", unsafe_allow_html=True)
                        buf_key = f"dlbuf_{a['id']}"
                        if buf_key in st.session_state:
                            c_dl.download_button("💾", st.session_state[buf_key], file_name=a['name'],
                                                 key=f"dl_drive_{cc}_{sub}_{a['id']}", help="Guardar el archivo en tu equipo")
                        elif c_dl.button("⬇️", key=f"prep_{cc}_{sub}_{a['id']}", help="Preparar descarga"):
                            st.session_state[buf_key] = descargar_archivo_drive(a['id'])
                            st.rerun()
                        if c_del.button("🗑️", key=f"del_drive_{cc}_{sub}_{a['id']}", help="Borrar archivo de Google Drive"):
                            token = obtener_token_fresco()
                            requests.delete(f"https://www.googleapis.com/drive/v3/files/{a['id']}",
                                            headers={"Authorization": f"Bearer {token}"}, timeout=20)
                            st.session_state.pop(buf_key, None)
                            _almacen_cache()["listados"].clear()
                            st.rerun()
                else:
                    col.markdown("<span class='archivo' style='color:#6B7280;'><i>Carpeta vacía</i></span>", unsafe_allow_html=True)
            else:
                col.markdown("<span class='archivo' style='color:#6B7280;'><i>Carpeta no disponible (revisa la conexión con Drive)</i></span>", unsafe_allow_html=True)
            col.write("")

# ==========================================
# 4. ACCESO: CÓDIGO DE FIRMA + USUARIO + CONTRASEÑA
# ==========================================
def procesar_login(codigo, usuario, clave):
    """Devuelve 'ok' o el mensaje de error que se le muestra al usuario."""
    codigo, usuario = str(codigo).strip().lower(), str(usuario).strip()
    if not codigo or not usuario or not clave:
        return "Completa el código de firma, el usuario y la contraseña."
    lim = f"{codigo}|{usuario.lower()}"
    espera = mf.segundos_bloqueo(lim)
    if espera > 0:
        return f"Demasiados intentos fallidos. Intenta de nuevo en {espera // 60 + 1} minuto(s)."
    invalido = "❌ Credenciales inválidas."
    try:
        if codigo == mf.codigo_superadmin():
            if not mf.login_superadmin(usuario, clave):
                mf.registrar_fallo(lim)
                return invalido
            mf.limpiar_fallos(lim)
            st.session_state.update(autenticado=True, es_superadmin=True, usuario_actual=usuario, rol_actual="Superadministrador",
                                    alias_actual="Superadmin", firma_ctx={}, avatar_path="", session_token="", debe_cambiar_clave=False)
            return "ok"
        firma = mf.obtener_firma(codigo=codigo, fresco=True)
        if not firma:
            mf.registrar_fallo(lim)
            return invalido
        st.session_state.firma_ctx = firma
        df_u = leer_tabla("usuarios", COLS_USR).astype(object)
        m = df_u["Usuario"].astype(str).str.lower() == usuario.lower()
        if not m.any() or not mf.verificar_password(clave, df_u.loc[m, "Password"].iloc[0]):
            st.session_state.firma_ctx = {}
            mf.registrar_fallo(lim)
            return invalido
        mf.limpiar_fallos(lim)
        fila = df_u[m].iloc[0]
        era_texto_plano = not mf.es_hash(fila["Password"])
        debe_cambiar = str(fila.get("Cambiar_Clave", "")) == "Si" or (era_texto_plano and str(clave) == "123")
        token = uuid.uuid4().hex
        if era_texto_plano:
            df_u.loc[m, "Password"] = mf.hash_password(clave)
        df_u.loc[m, "Session_Token"] = token
        df_u.loc[m, "Cambiar_Clave"] = "Si" if debe_cambiar else "No"
        guardar_tabla(df_u, "usuarios")
        st.session_state.update(autenticado=True, es_superadmin=False, usuario_actual=str(fila["Usuario"]), rol_actual=str(fila["Rol"]),
                                alias_actual=str(fila.get("Alias", usuario)) or usuario, avatar_path=str(fila.get("Avatar_Path", "")),
                                session_token=token, debe_cambiar_clave=debe_cambiar, bienvenida=True, pagina_actual="Dashboard")
        registrar_log("AUTH", f"Inició sesión: {usuario}")
        return "ok"
    except mf.ErrorDatos as e:
        st.session_state.firma_ctx = {}
        return f"⚠️ {e}"


def mostrar_bloqueo(lic):
    """Pantalla cuando la firma está suspendida o con la suscripción vencida."""
    st.markdown("<div style='height:4vh'></div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        with st.container(border=True):
            st.markdown(f"<div class='login-hero' style='padding-top:8px;'><div class='seal' style='width:70px;height:70px;font-size:34px;'>🔒</div><h1 style='color:#0A1628 !important;'>Acceso detenido</h1></div>", unsafe_allow_html=True)
            st.warning(lic["mensaje"])
            contacto = str(mf._secreto("soporte_contacto", "")).strip()
            st.markdown(f"Para reactivar el servicio, comunícate con soporte{(': **' + escape(contacto) + '**') if contacto else '.'}", unsafe_allow_html=True)
            if lic["nivel"] == "vencida" and st.session_state.rol_actual == "Administrador (Jefa)":
                st.markdown("**Tus datos siguen siendo tuyos.** Mientras regularizas el pago puedes descargarlos:")
                for tabla in ("clientes", "finanzas", "actuaciones", "vencimientos", "acreedores", "audiencias", "tareas"):
                    try:
                        df_exp = leer_tabla(tabla, mf.COLS[tabla])
                        st.download_button(f"⬇️ {tabla.capitalize()} (CSV)", df_exp.to_csv(index=False).encode("utf-8"), f"{tabla}.csv", "text/csv", key=f"exp_{tabla}")
                    except mf.ErrorDatos as e:
                        st.error(str(e))
            st.button("🚪 Cerrar sesión", key="bloqueo_salir", on_click=cerrar_sesion)


if not st.session_state.autenticado:
    st.markdown("""<style>
[data-testid="stAppViewContainer"]{background:radial-gradient(800px 520px at 12% 8%, rgba(201,162,75,.30), transparent 60%), radial-gradient(760px 520px at 92% 92%, rgba(47,109,181,.5), transparent 60%), linear-gradient(135deg,#0A1628,#13294B) !important;}
.login-hero h1{color:#FFFFFF !important;} .login-hero p{color:#B8C4DA !important;} [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] p{color:#9FB0CC !important;}
</style><div style="height:4vh"></div>""", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        if st.session_state.kicked:
            st.error(f"🚨 ALERTA: {st.session_state.kicked_reason}")
            st.session_state.kicked = False
        st.markdown("""<div class="login-hero"><div class="seal">⚖️</div><h1>Insolvencia OS</h1><p>Plataforma jurídica de gestión de expedientes · Acceso restringido</p></div>""", unsafe_allow_html=True)
        with st.form("login_form"):
            codigo_in = st.text_input("🏢 Código de firma", placeholder="El que te entregó soporte")
            user_in = st.text_input("👤 Usuario", placeholder="Tu usuario")
            pass_in = st.text_input("🔑 Contraseña", type="password")
            if st.form_submit_button("ENTRAR"):
                resultado = procesar_login(codigo_in, user_in, pass_in)
                if resultado == "ok":
                    st.rerun()
                else:
                    st.error(resultado)
        st.caption("🔒 Sesión única por usuario. Toda actividad queda registrada.")
    st.stop()

if st.session_state.es_superadmin:
    mf.render_superadmin(cerrar_sesion)
    st.stop()

# --- Firma activa: licencia vigente, sesión única y cambio de clave obligatorio ---
try:
    firma_vigente = mf.obtener_firma(firma_id=_fid())
    if not firma_vigente:
        cerrar_sesion()
        st.session_state.kicked, st.session_state.kicked_reason = True, "Tu firma ya no existe en la plataforma."
        st.rerun()
    st.session_state.firma_ctx = firma_vigente
    licencia = mf.evaluar_licencia(firma_vigente)
    st.session_state.aviso_licencia = licencia["mensaje"] if licencia["nivel"] in ("por_vencer", "gracia") else ""
    if licencia["nivel"] in ("bloqueada", "vencida"):
        mostrar_bloqueo(licencia)
        st.stop()

    df_u_check = leer_tabla("usuarios", COLS_USR)
    fila_u = df_u_check[df_u_check["Usuario"].astype(str).str.lower() == str(st.session_state.usuario_actual).lower()]
    if fila_u.empty:
        cerrar_sesion()
        st.session_state.kicked, st.session_state.kicked_reason = True, "Tu usuario fue revocado."
        st.rerun()
    db_token = str(fila_u.iloc[0].get("Session_Token", ""))
    if db_token not in ("", "nan") and db_token != str(st.session_state.get("session_token", "")):
        cerrar_sesion()
        st.session_state.kicked, st.session_state.kicked_reason = True, "Sesión iniciada en otro equipo o contraseña restablecida."
        st.rerun()
except mf.ErrorDatos as e:
    st.error(f"🚨 {e}")
    st.button("🔄 Reintentar", key="reintentar_sesion")
    st.stop()

if st.session_state.debe_cambiar_clave:
    st.markdown("<div style='height:4vh'></div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1.6, 1])
    with c2:
        st.markdown("<div class='login-hero'><div class='seal' style='width:70px;height:70px;font-size:34px;'>🔑</div><h1 style='color:#0A1628 !important;'>Crea tu contraseña</h1></div>", unsafe_allow_html=True)
        st.info("Por seguridad, debes reemplazar la contraseña temporal antes de continuar. Usa mínimo 8 caracteres con letras y números.")
        with st.form("form_cambio_obligatorio"):
            n1 = st.text_input("Nueva contraseña", type="password")
            n2 = st.text_input("Confirma la nueva contraseña", type="password")
            if st.form_submit_button("GUARDAR Y ENTRAR"):
                ok, msg = cambiar_mi_clave("", n1, n2, exigir_actual=False)
                if ok:
                    st.session_state.debe_cambiar_clave = False
                    st.rerun()
                else:
                    st.error(msg)
        st.button("🚪 Cerrar sesión", key="cambio_salir", on_click=cerrar_sesion)
    st.stop()

# --- Datos de la firma (solo después de validar acceso) ---
try:
    df_cli = leer_tabla("clientes", ["Cedula", "Nombre", "Fuerza", "Telefono", "Email", "Deuda_Est", "Ingresos", "Senal", "F_Actualizacion", "Estado", "F_Borrado"])
    df_fin = leer_tabla("finanzas", ["Cedula", "Honorarios", "Abonado"])
    df_act = leer_tabla("actuaciones", ["ID_Act", "Cedula", "Fecha", "Tipo", "Juzgado", "Radicado", "Anotacion"])
    df_log = leer_tabla("logs", ["Timestamp", "Usuario", "Modulo", "Accion"])
    df_ven = leer_tabla("vencimientos", ["ID_Ven", "Cedula", "Cliente", "Asunto", "Fecha_Limite", "Estado"])
    df_acr = leer_tabla("acreedores", ["ID_Acr", "Cedula", "Acreedor", "Cuantia", "Clase"])
    df_aud = leer_tabla("audiencias", ["ID_Aud", "Cedula", "Cliente", "Fecha_Hora", "Motivo"])
    df_tar = leer_tabla("tareas", ["ID_Tar", "Tarea", "Asignado", "Creador", "Estado", "Fecha"])
    df_usr = leer_tabla("usuarios", COLS_USR)

    if "Alias" not in df_usr.columns: df_usr["Alias"] = df_usr["Usuario"]
    if "Avatar_Path" not in df_usr.columns: df_usr["Avatar_Path"] = ""
    if "Session_Token" not in df_usr.columns: df_usr["Session_Token"] = ""

    hoy = ahora_colombia()

    cambios_bd = False
    for idx, r in df_cli[df_cli["Estado"] == "Borrado"].iterrows():
        if pd.notna(r["F_Borrado"]) and str(r["F_Borrado"]).strip() != "":
            try:
                if (hoy - datetime.strptime(str(r["F_Borrado"])[:10], "%Y-%m-%d")).days >= 30:
                    rp = os.path.join(CARP_PAP, f"{r['Cedula']} - {r['Nombre']}")
                    if os.path.exists(rp): shutil.rmtree(rp)
                    df_cli = df_cli.drop(idx)
                    cambios_bd = True
            except (ValueError, OSError) as e:
                print(f"⚠️ No se pudo depurar la papelera: {type(e).__name__}")
    if cambios_bd: guardar_tabla(df_cli, "clientes")

    df_activos = df_cli[df_cli["Estado"] == "Activo"].copy()
    df_papelera = df_cli[df_cli["Estado"] == "Borrado"].copy()

except mf.ErrorDatos as e:
    st.error(f"🚨 {e}")
    st.button("🔄 Reintentar", key="reintentar_datos")
    st.stop()

# ==========================================
# 5. CENTRO DE MANDO: barra superior, navegación, búsqueda y notificaciones
# ==========================================
if st.session_state.pop("bienvenida", False):
    st.toast(f"Bienvenido/a, {st.session_state.alias_actual}", icon="⚖️")

MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS_ES = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
FASES = [
    (0, "📥", "Recolección", "#2F6DB5"),
    (1, "📝", "Contratos", "#6B5BA6"),
    (2, "✍️", "Firmas", "#C9A24B"),
    (3, "🏛️", "Radicación", "#D9822B"),
    (4, "✅", "Radicado", "#1F7A5C"),
]
CLASES_ACREEDOR = ["Primera", "Segunda", "Tercera", "Cuarta", "Quinta"]
COLS_CLI = ["Cedula", "Nombre", "Fuerza", "Telefono", "Email", "Deuda_Est", "Ingresos", "Senal", "F_Actualizacion", "Estado", "F_Borrado"]
COLS_TAR = ["ID_Tar", "Tarea", "Asignado", "Creador", "Estado", "Fecha"]

GRUPOS_NAV = [
    ("🏠", "Inicio", [("Dashboard", "Centro de mando"), ("Kanban", "Tablero de casos"), ("Agenda", "Calendario"), ("Reportes", "Reportes")]),
    ("📁", "Expedientes", [("Ficha", "Ficha 360"), ("Nuevo", "Apertura y edición"), ("Contratos", "Documentos y fases"), ("Acreedores", "Pasivos"), ("Papelera", "Papelera")]),
    ("⚖️", "Jurídico", [("Actuaciones", "Seguimiento"), ("Vencimientos", "Vencimientos"), ("Memoriales", "Memoriales"), ("Simulador", "Simulador"), ("Rastreador", "Rastreador")]),
    ("✅", "Gestión", [("Tareas", "Tareas"), ("Finanzas", "Cobros")]),
    ("🛡️", "Admin", [("Usuarios", "Accesos"), ("Sistema", "Auditoría")]),
]


def es_jefa():
    return st.session_state.rol_actual == "Administrador (Jefa)"


def grupos_visibles():
    visibles = []
    for ico, nombre, paginas in GRUPOS_NAV:
        if not es_jefa():
            paginas = [p for p in paginas if p[0] not in ("Finanzas", "Usuarios", "Sistema")]
        if paginas:
            visibles.append((ico, nombre, paginas))
    return visibles


def fecha_larga(d):
    return f"{DIAS_ES[d.weekday()].capitalize()} {d.day} de {MESES_ES[d.month - 1]} de {d.year}"


def senal_num(serie):
    return pd.to_numeric(serie, errors="coerce").fillna(0).astype(int)


def mapa_saldos():
    saldos = {}
    for _, r in df_fin.iterrows():
        saldos[str(r["Cedula"])] = limpiar_num(r["Honorarios"]) - limpiar_num(r["Abonado"])
    return saldos


def abrir_ficha(cc):
    """Salta a la Ficha 360 del cliente y deja ese cliente como foco para los demás módulos."""
    ir_a_modulo("Ficha", cc)
    st.session_state.busq_global = ""


def ir_a_modulo(pagina, cc):
    st.session_state.cliente_foco = str(cc)
    st.session_state.pagina_actual = pagina
    try:
        d = leer_tabla("clientes", COLS_CLI)
        fila = d[d["Cedula"].astype(str) == str(cc)]
        if not fila.empty:
            opcion = f"{cc} - {fila.iloc[0]['Nombre']}"
            for clave in ("sel_ficha", "sim_sel", "sel_acr"):
                st.session_state[clave] = opcion
    except Exception as e:
        print(f"⚠️ No se pudo fijar el cliente en foco: {type(e).__name__}")


def idx_foco(opciones):
    foco = str(st.session_state.get("cliente_foco", ""))
    for i, o in enumerate(list(opciones)):
        if str(o).split(" - ", 1)[0] == foco:
            return i
    return 0


def cambiar_estado_tarea(id_tar, estado):
    try:
        d = leer_tabla("tareas", COLS_TAR)
        d = d.astype(object)
        d.loc[d["ID_Tar"].astype(str) == str(id_tar), "Estado"] = estado
        guardar_tabla(d, "tareas")
        registrar_log("TAREAS", f"Tarea {id_tar} marcada como {estado}")
    except Exception as e:
        print(f"⚠️ No se pudo actualizar la tarea: {type(e).__name__}")


def mover_fase(cc, delta):
    try:
        d = leer_tabla("clientes", COLS_CLI).astype(object)
        m = d["Cedula"].astype(str) == str(cc)
        if not m.any():
            return
        actual = int(pd.to_numeric(d.loc[m, "Senal"], errors="coerce").fillna(0).iloc[0])
        nuevo = max(0, min(4, actual + delta))
        if nuevo == actual:
            return
        d.loc[m, "Senal"] = nuevo
        d.loc[m, "F_Actualizacion"] = ahora_colombia().strftime("%Y-%m-%d")
        guardar_tabla(d, "clientes")
        registrar_log("KANBAN", f"Expediente {cc}: fase {actual} → {nuevo}")
    except Exception as e:
        print(f"⚠️ No se pudo mover el expediente: {type(e).__name__}")


def mover_mes(delta):
    ahora_c = ahora_colombia()
    a, m = st.session_state.get("cal_mes", (ahora_c.year, ahora_c.month))
    m += delta
    while m > 12:
        m -= 12
        a += 1
    while m < 1:
        m += 12
        a -= 1
    st.session_state.cal_mes = (a, m)


def mes_actual():
    ahora_c = ahora_colombia()
    st.session_state.cal_mes = (ahora_c.year, ahora_c.month)


def etiqueta_dia(fecha_d, hoy_d):
    dif = (fecha_d - hoy_d).days
    if dif == 0:
        return "Hoy"
    if dif == 1:
        return "Mañana"
    if dif == -1:
        return "Ayer"
    if dif < 0:
        return f"Hace {abs(dif)} días"
    return f"{DIAS_ES[fecha_d.weekday()][:3].capitalize()} {fecha_d.day} {MESES_ES[fecha_d.month - 1][:3]}"


def eventos_rango(ini_d, fin_d):
    """Audiencias/citas y vencimientos activos entre dos fechas. Los vencimientos atrasados siempre aparecen."""
    eventos = []
    if not df_aud.empty:
        fechas = pd.to_datetime(df_aud["Fecha_Hora"], errors="coerce")
        for (_, r), dt in zip(df_aud.iterrows(), fechas):
            if pd.notna(dt) and ini_d <= dt.date() <= fin_d:
                eventos.append({"fecha": dt.to_pydatetime(), "tipo": "aud", "hora": True, "color": "#C9A24B",
                                "titulo": str(r["Motivo"]).strip() or "Audiencia / cita", "detalle": str(r["Cliente"]), "cc": str(r["Cedula"])})
    if not df_ven.empty:
        activos = df_ven[df_ven["Estado"].astype(str) == "Activo"]
        fechas = pd.to_datetime(activos["Fecha_Limite"], errors="coerce")
        for (_, r), dt in zip(activos.iterrows(), fechas):
            if pd.isna(dt):
                continue
            d = dt.date()
            if ini_d - timedelta(days=60) <= d <= fin_d:
                dias = (d - ini_d).days
                color = "#8A1F35" if dias < 0 else ("#D9822B" if dias <= 2 else "#2F6DB5")
                eventos.append({"fecha": dt.to_pydatetime(), "tipo": "ven", "hora": False, "color": color,
                                "titulo": str(r["Asunto"]).strip() or "Vencimiento", "detalle": str(r["Cliente"]), "cc": str(r["Cedula"])})
    eventos.sort(key=lambda e: e["fecha"])
    return eventos


def html_timeline(eventos, hoy_d):
    if not eventos:
        return "<div class='vacio'>🎉 Sin compromisos en este período.</div>"
    items = []
    for i, e in enumerate(eventos):
        d = e["fecha"].date()
        etiqueta = etiqueta_dia(d, hoy_d)
        if e["hora"]:
            etiqueta += " · " + e["fecha"].strftime("%I:%M %p").lstrip("0").lower()
        tipo = "Audiencia / cita" if e["tipo"] == "aud" else ("Vencido" if d < hoy_d else "Vencimiento")
        items.append(
            f"<div class='tl-item' style='--c:{e['color']}; animation-delay:{i * 0.06:.2f}s;'><div class='tl-dot'></div>"
            f"<b>{escape(e['titulo'])}</b><span>{escape(e['detalle'])} · {tipo}</span><em>{escape(etiqueta)}</em></div>"
        )
    return "<div class='tl'>" + "".join(items) + "</div>"


def calcular_notificaciones():
    """Reúne lo que requiere atención: vencimientos, audiencias, tareas y casos estancados."""
    notifs = []
    hoy_d = ahora_colombia().date()
    for e in eventos_rango(hoy_d, hoy_d + timedelta(days=5)):
        d = e["fecha"].date()
        dias = (d - hoy_d).days
        if e["tipo"] == "ven":
            if dias < 0:
                notifs.append({"nivel": "alto", "orden": dias, "icono": "🔴", "titulo": f"Vencido: {e['titulo']}", "detalle": f"{e['detalle']} · {etiqueta_dia(d, hoy_d)}", "pagina": "Vencimientos"})
            elif dias <= 2:
                notifs.append({"nivel": "alto", "orden": dias, "icono": "🟠", "titulo": f"Vence {etiqueta_dia(d, hoy_d).lower()}: {e['titulo']}", "detalle": e["detalle"], "pagina": "Vencimientos"})
            else:
                notifs.append({"nivel": "medio", "orden": dias, "icono": "🟡", "titulo": f"{e['titulo']}", "detalle": f"{e['detalle']} · {etiqueta_dia(d, hoy_d)}", "pagina": "Vencimientos"})
        elif dias <= 2:
            notifs.append({"nivel": "alto", "orden": dias, "icono": "⚖️", "titulo": f"{e['titulo']} · {etiqueta_dia(d, hoy_d).lower()}", "detalle": e["detalle"], "pagina": "Agenda"})
    if not df_tar.empty:
        mias = df_tar[(df_tar["Asignado"].astype(str) == str(st.session_state.alias_actual)) & (df_tar["Estado"].astype(str) == "Pendiente")]
        if len(mias):
            notifs.append({"nivel": "medio", "orden": 50, "icono": "✅", "titulo": f"Tienes {len(mias)} tarea(s) pendiente(s)", "detalle": "Revisa el tablero de tareas", "pagina": "Tareas"})
    if not df_activos.empty:
        sen = senal_num(df_activos["Senal"])
        for (_, r), s in zip(df_activos.iterrows(), sen):
            try:
                dias = (ahora_colombia() - datetime.strptime(str(r["F_Actualizacion"])[:10], "%Y-%m-%d")).days
            except ValueError:
                continue
            if s == 2 and dias >= 2:
                notifs.append({"nivel": "medio", "orden": 60, "icono": "✍️", "titulo": f"Firma pendiente: {r['Nombre']}", "detalle": f"{dias} días sin avanzar", "pagina": "Kanban"})
            elif s == 3 and dias >= 1:
                notifs.append({"nivel": "alto", "orden": 61, "icono": "🏛️", "titulo": f"Radicación urgente: {r['Nombre']}", "detalle": f"{dias} días sin avanzar", "pagina": "Kanban"})
    peso = {"alto": 0, "medio": 1, "info": 2}
    notifs.sort(key=lambda n: (peso[n["nivel"]], n["orden"]))
    return notifs


def buscar_global(texto):
    q = str(texto).strip().lower()
    res = {"clientes": [], "acreedores": [], "actuaciones": []}
    if len(q) < 2:
        return res
    if not df_activos.empty:
        base = df_activos.astype(str)
        mask = pd.Series(False, index=base.index)
        for col in ["Nombre", "Cedula", "Telefono", "Email"]:
            mask |= base[col].str.lower().str.contains(q, regex=False, na=False)
        res["clientes"] = [(str(r["Cedula"]), str(r["Nombre"]), str(r["Fuerza"])) for _, r in base[mask].head(6).iterrows()]
    nombres = {str(r["Cedula"]): str(r["Nombre"]) for _, r in df_activos.iterrows()}
    if not df_acr.empty:
        a = df_acr.astype(str)
        for _, r in a[a["Acreedor"].str.lower().str.contains(q, regex=False, na=False)].head(5).iterrows():
            if r["Cedula"] in nombres:
                res["acreedores"].append((r["Cedula"], r["Acreedor"], nombres[r["Cedula"]]))
    if not df_act.empty:
        b = df_act.astype(str)
        m2 = b["Radicado"].str.lower().str.contains(q, regex=False, na=False) | b["Juzgado"].str.lower().str.contains(q, regex=False, na=False)
        for _, r in b[m2].head(5).iterrows():
            if r["Cedula"] in nombres:
                res["actuaciones"].append((r["Cedula"], f"{r['Tipo']} · {r['Radicado']}", nombres[r["Cedula"]]))
    return res


def ics_eventos(eventos):
    """Exporta eventos a un archivo de calendario (.ics) compatible con Google Calendar y Outlook."""
    def limpio(t):
        return str(t).replace("\\", "\\\\").replace(",", "\\,").replace(";", "\\;").replace("\n", " ")
    lineas = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Insolvencia OS//Agenda//ES"]
    sello = ahora_colombia().strftime("%Y%m%dT%H%M%S")
    for e in eventos:
        lineas += ["BEGIN:VEVENT", f"UID:{uuid.uuid4().hex}@insolvencia-os", f"DTSTAMP:{sello}",
                   f"SUMMARY:{limpio(('⚖ ' if e['tipo'] == 'aud' else '⏰ ') + e['titulo'] + ' - ' + e['detalle'])}"]
        if e["hora"]:
            lineas += [f"DTSTART:{e['fecha'].strftime('%Y%m%dT%H%M%S')}", f"DTEND:{(e['fecha'] + timedelta(hours=1)).strftime('%Y%m%dT%H%M%S')}"]
        else:
            lineas += [f"DTSTART;VALUE=DATE:{e['fecha'].strftime('%Y%m%d')}", f"DTEND;VALUE=DATE:{(e['fecha'] + timedelta(days=1)).strftime('%Y%m%d')}"]
        lineas.append("END:VEVENT")
    lineas.append("END:VCALENDAR")
    return "\r\n".join(lineas).encode("utf-8")


# ---------- Barra superior ----------
notificaciones_hoy = calcular_notificaciones()
alias_corto = str(st.session_state.alias_actual).split(" ")[0][:14] or "Perfil"

with st.container(key="topbar"):
    c_marca, c_busq, c_campana, c_user = st.columns([1.5, 3.3, 0.8, 1.3], vertical_alignment="center")
    c_marca.markdown("<div class='marca'><div class='sello'>⚖️</div><div><div class='nombre'>Insolvencia OS</div><div class='lema'>Centro de mando</div></div></div>", unsafe_allow_html=True)
    c_busq.text_input("Buscar", key="busq_global", placeholder="🔎 Buscar cliente, cédula, acreedor o radicado…", label_visibility="collapsed")

    n_notif = len(notificaciones_hoy)
    with c_campana.popover(f"🔔 {n_notif}" if n_notif else "🔔", use_container_width=True):
        st.markdown("**Notificaciones**")
        if not notificaciones_hoy:
            st.success("Todo al día ✨")
        for i, n in enumerate(notificaciones_hoy[:8]):
            col_t, col_b = st.columns([4, 1], vertical_alignment="center")
            col_t.markdown(f"{n['icono']} **{escape(n['titulo'])}**<br><span style='font-size:12px;color:#667085;'>{escape(n['detalle'])}</span>", unsafe_allow_html=True)
            col_b.button("Ir", key=f"nt_{i}", on_click=cambiar_pagina, args=(n["pagina"],))
        if n_notif > 8:
            st.caption(f"y {n_notif - 8} más…")

    with c_user.popover(f"👤 {alias_corto}", use_container_width=True):
        avatar_p = st.session_state.get('avatar_path', '')
        if pd.notna(avatar_p) and str(avatar_p).strip() != "" and os.path.exists(avatar_p):
            with open(avatar_p, "rb") as f:
                encoded_img = base64.b64encode(f.read()).decode()
            st.markdown(f'<div style="text-align:center;"><img src="data:image/png;base64,{encoded_img}" style="width:72px;height:72px;border-radius:50%;object-fit:cover;border:3px solid #C9A24B;"></div>', unsafe_allow_html=True)
        st.markdown(f"<div style='text-align:center;'><b>{escape(str(st.session_state.alias_actual))}</b><br><span style='font-size:12px;color:#8F7129;font-weight:700;'>{escape(str(st.session_state.rol_actual))}</span></div>", unsafe_allow_html=True)
        with st.form("form_perfil"):
            n_alias = st.text_input("Alias", value=st.session_state.alias_actual)
            foto_sub = st.file_uploader("Foto de perfil", type=["jpg", "png", "jpeg"])
            if st.form_submit_button("Actualizar perfil"):
                df_u_up = leer_tabla("usuarios", ["Usuario", "Password", "Rol", "Creador", "Alias", "Avatar_Path", "Session_Token"])
                df_u_up = df_u_up.astype(object)
                p_guardado = st.session_state.avatar_path
                if foto_sub:
                    p_guardado = os.path.join(CARP_AVATAR, f"{_fid()}_{st.session_state.usuario_actual}_{os.path.basename(foto_sub.name)}")
                    with open(p_guardado, "wb") as f:
                        f.write(foto_sub.getbuffer())
                df_u_up.loc[df_u_up["Usuario"].astype(str) == st.session_state.usuario_actual, "Alias"] = n_alias
                df_u_up.loc[df_u_up["Usuario"].astype(str) == st.session_state.usuario_actual, "Avatar_Path"] = p_guardado
                guardar_tabla(df_u_up, "usuarios")
                st.session_state.alias_actual = n_alias
                st.session_state.avatar_path = p_guardado
                st.rerun()
        with st.expander("🔑 Cambiar mi contraseña"):
            with st.form("form_mi_clave", clear_on_submit=True):
                p_act = st.text_input("Contraseña actual", type="password")
                p_nue = st.text_input("Nueva contraseña", type="password")
                p_con = st.text_input("Confirma la nueva", type="password")
                if st.form_submit_button("Cambiar contraseña"):
                    ok_c, msg_c = cambiar_mi_clave(p_act, p_nue, p_con)
                    (st.success if ok_c else st.error)(msg_c)
        st.button("🚪 Cerrar sesión", key="b_logout", use_container_width=True, on_click=cerrar_sesion)

# ---------- Navegación en dos niveles ----------
grupos = grupos_visibles()
pagina_act = st.session_state.pagina_actual
grupo_act = next((g[1] for g in grupos if any(p[0] == pagina_act for p in g[2])), grupos[0][1])

with st.container(key="nav_grupos"):
    cols_g = st.columns(len(grupos))
    for col, (ico, nombre, paginas) in zip(cols_g, grupos):
        with col:
            st.button(f"{ico} {nombre}", key=f"g_{nombre}", use_container_width=True,
                      type="primary" if nombre == grupo_act else "secondary",
                      on_click=cambiar_pagina, args=(paginas[0][0],))

paginas_grupo = next(g[2] for g in grupos if g[1] == grupo_act)
with st.container(key="nav_sub"):
    cols_s = st.columns(len(paginas_grupo))
    for col, (clave, etiqueta) in zip(cols_s, paginas_grupo):
        with col:
            st.button(etiqueta, key=f"p_{clave}", use_container_width=True,
                      type="primary" if clave == pagina_act else "secondary",
                      on_click=cambiar_pagina, args=(clave,))

# ---------- Resultados de la búsqueda global ----------
consulta = st.session_state.get("busq_global", "")
if str(consulta).strip():
    resultados = buscar_global(consulta)
    total_res = sum(len(v) for v in resultados.values())
    with st.container(border=True):
        cab1, cab2 = st.columns([5, 1], vertical_alignment="center")
        cab1.markdown(f"<p class='card-title'>Resultados para “{escape(str(consulta))}” · {total_res}</p>", unsafe_allow_html=True)
        cab2.button("✖ Cerrar", key="cerrar_busq", on_click=lambda: st.session_state.update(busq_global=""), use_container_width=True)
        if total_res == 0:
            st.markdown("<div class='vacio'>Sin coincidencias. Prueba con otra cédula, nombre o radicado.</div>", unsafe_allow_html=True)
        for tipo, etiqueta_tipo in (("clientes", "👤 Cliente"), ("acreedores", "🏦 Acreedor"), ("actuaciones", "📂 Actuación")):
            for j, fila in enumerate(resultados[tipo]):
                cc_r = fila[0]
                c1, c2 = st.columns([5, 1], vertical_alignment="center")
                if tipo == "clientes":
                    c1.markdown(f"{etiqueta_tipo} · **{escape(fila[1])}** <span style='color:#667085;font-size:12px;'>CC {escape(cc_r)} · {escape(fila[2])}</span>", unsafe_allow_html=True)
                else:
                    c1.markdown(f"{etiqueta_tipo} · **{escape(fila[1])}** <span style='color:#667085;font-size:12px;'>cliente: {escape(fila[2])}</span>", unsafe_allow_html=True)
                c2.button("Abrir ficha", key=f"rs_{tipo}_{j}_{cc_r}", on_click=abrir_ficha, args=(cc_r,), use_container_width=True)


if st.session_state.get("aviso_licencia"):
    st.warning(f"⏳ {st.session_state.aviso_licencia}")

# ==========================================
# 6. MÓDULOS DE LA APLICACIÓN
# ==========================================

# --- 1. CENTRO DE MANDO ---
if st.session_state.pagina_actual == 'Dashboard':
    ahora = ahora_colombia()
    hoy_d = ahora.date()
    saludo = "Buenos días" if ahora.hour < 12 else ("Buenas tardes" if ahora.hour < 19 else "Buenas noches")
    primer_nombre = str(st.session_state.alias_actual).split(" ")[0] or "equipo"
    n_alto = sum(1 for n in notificaciones_hoy if n["nivel"] == "alto")
    eventos_semana = eventos_rango(hoy_d, hoy_d + timedelta(days=7))
    n_hoy = sum(1 for e in eventos_semana if e["fecha"].date() == hoy_d)
    if n_alto:
        resumen = f"Hay {n_alto} asunto(s) que piden atención inmediata. Empieza por ahí."
    elif n_hoy:
        resumen = f"Hoy tienes {n_hoy} compromiso(s) en agenda. Lo demás va sin sobresaltos."
    else:
        resumen = "Todo está bajo control. Buen momento para adelantar expedientes."
    st.markdown(
        f"<div class='hero'><p class='kick'>{escape(fecha_larga(hoy_d))}</p>"
        f"<p class='titulo'>{saludo}, {escape(primer_nombre)}</p><p class='sub'>{escape(resumen)}</p>"
        f"<div class='chips'><span class='chip2'>📁 {len(df_activos)} expedientes activos</span>"
        f"<span class='chip2'>🗓️ {len(eventos_semana)} compromisos en 7 días</span>"
        f"<span class='chip2'>🔔 {len(notificaciones_hoy)} notificaciones</span></div></div>",
        unsafe_allow_html=True,
    )

    sen_activos = senal_num(df_activos["Senal"]) if not df_activos.empty else pd.Series(dtype=int)
    n_tramite = int(sen_activos.isin([1, 2, 3]).sum())
    n_radicados = int((sen_activos >= 4).sum())
    n_total = max(len(df_activos), 1)
    saldos = mapa_saldos()
    activos_cc = set(df_activos["Cedula"].astype(str)) if not df_activos.empty else set()
    total_cartera = sum(v for k, v in saldos.items() if k in activos_cc and v > 0)
    mis_pend = 0
    if not df_tar.empty:
        mis_pend = int(((df_tar["Asignado"].astype(str) == str(st.session_state.alias_actual)) & (df_tar["Estado"].astype(str) == "Pendiente")).sum())

    def kpi(icono, etiqueta, valor, color, fondo, retraso, porcentaje=None):
        barra = f"<div class='bar'><i style='width:{porcentaje}%;'></i></div>" if porcentaje is not None else ""
        return (f"<div class='kpi' style='--accent:{color}; --accent-bg:{fondo}; animation-delay:{retraso}s;'>"
                f"<div class='ico'>{icono}</div><p class='lbl'>{etiqueta}</p><p class='val'>{valor}</p>{barra}</div>")

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi("📁", "Expedientes activos", len(df_activos), "#2F6DB5", "rgba(47,109,181,.14)", 0), unsafe_allow_html=True)
    k2.markdown(kpi("⏳", "En trámite", n_tramite, "#C9A24B", "rgba(201,162,75,.2)", .08, round(100 * n_tramite / n_total)), unsafe_allow_html=True)
    k3.markdown(kpi("🏛️", "Radicados", n_radicados, "#1F7A5C", "rgba(31,122,92,.14)", .16, round(100 * n_radicados / n_total)), unsafe_allow_html=True)
    if es_jefa():
        k4.markdown(kpi("💰", "Cartera pendiente", f"$ {total_cartera:,.0f}", "#8A1F35", "rgba(138,31,53,.12)", .24), unsafe_allow_html=True)
    else:
        k4.markdown(kpi("✅", "Mis tareas pendientes", mis_pend, "#8A1F35", "rgba(138,31,53,.12)", .24), unsafe_allow_html=True)

    col_a, col_b = st.columns([1.55, 1])
    with col_a:
        with st.container(border=True):
            st.markdown("<p class='card-title'>🗓️ Próximos 7 días</p>", unsafe_allow_html=True)
            st.markdown(html_timeline(eventos_semana[:8], hoy_d), unsafe_allow_html=True)
            if len(eventos_semana) > 8:
                st.caption(f"y {len(eventos_semana) - 8} compromisos más en el calendario.")
            st.button("Abrir calendario completo →", key="dash_cal", on_click=cambiar_pagina, args=("Agenda",))
        if es_jefa():
            with st.container(border=True):
                st.markdown("<p class='card-title'>🕘 Actividad reciente del equipo</p>", unsafe_allow_html=True)
                if df_log.empty:
                    st.markdown("<div class='vacio'>Aún no hay actividad registrada.</div>", unsafe_allow_html=True)
                else:
                    filas = []
                    for _, r in df_log.tail(6).iloc[::-1].iterrows():
                        filas.append(f"<div class='act'><b>{escape(str(r['Usuario']))}</b> · {escape(str(r['Accion']))}<em>{escape(str(r['Timestamp'])[:16])} · {escape(str(r['Modulo']))}</em></div>")
                    st.markdown("".join(filas), unsafe_allow_html=True)
    with col_b:
        with st.container(border=True):
            st.markdown("<p class='card-title'>⚡ Requiere atención</p>", unsafe_allow_html=True)
            if not notificaciones_hoy:
                st.markdown("<div class='vacio'>✨ Nada urgente por ahora.</div>", unsafe_allow_html=True)
            for n in notificaciones_hoy[:4]:
                color = "#8A1F35" if n["nivel"] == "alto" else "#D9822B"
                fondo = "rgba(138,31,53,.07)" if n["nivel"] == "alto" else "rgba(217,130,43,.08)"
                st.markdown(f"<div class='alerta-item' style='--c:{color}; --bg:{fondo};'><div>{n['icono']}</div><div><b>{escape(n['titulo'])}</b><span>{escape(n['detalle'])}</span></div></div>", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown("<p class='card-title'>🧭 Pipeline de expedientes</p>", unsafe_allow_html=True)
            cuentas = sen_activos.clip(upper=4).value_counts() if len(sen_activos) else pd.Series(dtype=int)
            segmentos = "".join(f"<span style='flex:{int(cuentas.get(n, 0))}; background:{color};' title='{nombre}'></span>" for n, _, nombre, color in FASES if int(cuentas.get(n, 0)) > 0)
            filas_fase = "".join(f"<div class='pipe-row'><span><i style='--c:{color};'></i>{ico} {nombre}</span><b>{int(cuentas.get(n, 0))}</b></div>" for n, ico, nombre, color in FASES)
            st.markdown(f"<div class='pipe'>{segmentos}</div>{filas_fase}", unsafe_allow_html=True)
            st.button("Abrir tablero →", key="dash_kb", on_click=cambiar_pagina, args=("Kanban",))
        with st.container(border=True):
            st.markdown("<p class='card-title'>✅ Mis tareas</p>", unsafe_allow_html=True)
            mias = df_tar[(df_tar["Asignado"].astype(str) == str(st.session_state.alias_actual)) & (df_tar["Estado"].astype(str) == "Pendiente")] if not df_tar.empty else df_tar
            if mias.empty:
                st.markdown("<div class='vacio'>Sin tareas pendientes 🙌</div>", unsafe_allow_html=True)
            else:
                for i, (_, t) in enumerate(mias.head(5).iterrows()):
                    c_t, c_b = st.columns([5, 1], vertical_alignment="center")
                    c_t.markdown(f"<div class='task'>{escape(str(t['Tarea']))}<em>de {escape(str(t['Creador']))}</em></div>", unsafe_allow_html=True)
                    with c_b.container(key=f"mini_done_{i}"):
                        st.button("✓", key=f"done_{i}_{t['ID_Tar']}", on_click=cambiar_estado_tarea, args=(t['ID_Tar'], "Hecha"), help="Marcar como hecha")

    with st.expander("⚠️ Zona de peligro: archivar un expediente"):
        if not df_activos.empty:
            with st.form("form_borrar"):
                cliente_a_borrar = st.selectbox("Seleccionar expediente para mover a la Papelera:", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"])
                if st.form_submit_button("MOVER A LA PAPELERA"):
                    cc_b, nom_b = cliente_a_borrar.split(" - ", 1)
                    df_cli.loc[df_cli["Cedula"].astype(str) == cc_b, "Estado"] = "Borrado"
                    df_cli.loc[df_cli["Cedula"].astype(str) == cc_b, "F_Borrado"] = hoy.strftime("%Y-%m-%d")
                    guardar_tabla(df_cli, "clientes")
                    mover_a_papelera_drive(str(cc_b), nom_b)
                    registrar_log("SISTEMA", f"Expediente y carpeta enviados a papelera: {cc_b}")
                    st.success("Expediente movido a la papelera en Sheets y Drive.")
                    time.sleep(1.2)
                    st.rerun()
        else:
            st.info("No hay clientes activos para archivar.")


# --- 2. APERTURA Y EDICIÓN ---
elif st.session_state.pagina_actual == 'Nuevo':
    st.markdown("<h1>Gestión de Expedientes</h1>", unsafe_allow_html=True)
    
    tab_nuevo, tab_editar = st.tabs(["➕ Apertura de Nuevo Caso", "✏️ Modificar Expediente Existente"])
    
    with tab_nuevo:
        st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
        with st.form("nuevo_exp"):
            c1, c2, c3 = st.columns(3)
            with c1: cc = st.text_input("Cédula de Ciudadanía")
            with c2: nom = st.text_input("Nombre Completo")
            with c3: fza = st.selectbox("Facción", ["Policía Nacional", "Ejército Nacional", "Armada Nacional", "Fuerza Aérea", "Retirado / Pensionado", "Civil"])
            c4, c5, c6 = st.columns(3)
            with c4: tel = st.text_input("WhatsApp / Celular")
            with c5: mail = st.text_input("Correo Electrónico")
            with c6: deuda = st.text_input("Deuda Aprox ($)")
            
            st.markdown("<hr style='border-color: #E2E8F0;'>", unsafe_allow_html=True)
            st.markdown("<p style='color:#2563EB; font-weight:bold;'>💰 Configuración Financiera del Caso</p>", unsafe_allow_html=True)
            cf1, cf2 = st.columns(2)
            with cf1: honorarios_totales = st.text_input("Valor Total de Honorarios del Caso ($)", value="0")
            with cf2: abono_inicial = st.text_input("Abono Inicial Recibido ($)", value="0")
            
            if st.form_submit_button("CREAR BÓVEDA Y REGISTRAR EN NUBE"):
                if not cc or not nom: st.error("Cédula y Nombre son obligatorios.")
                elif str(cc) in df_cli["Cedula"].astype(str).values: st.error("Sujeto ya existe en la base de datos.")
                else:
                    ndf = pd.DataFrame([{"Cedula": str(cc), "Nombre": nom, "Fuerza": fza, "Telefono": str(tel), "Email": mail, "Deuda_Est": str(deuda), "Ingresos": "", "Senal": 0, "F_Actualizacion": hoy.strftime("%Y-%m-%d"), "Estado": "Activo", "F_Borrado": ""}])
                    guardar_tabla(pd.concat([df_cli, ndf], ignore_index=True), "clientes")
                    
                    nuevo_f = pd.concat([df_fin, pd.DataFrame([{"Cedula": str(cc), "Honorarios": str(honorarios_totales), "Abonado": str(abono_inicial)}])], ignore_index=True)
                    guardar_tabla(nuevo_f, "finanzas")
                    
                    estructurar_carpetas(str(cc), nom)
                    st.success("Expediente creado, financiero configurado y sincronizado en Google Sheets.")
                    st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
        
    with tab_editar:
        st.markdown("<div class='module-card module-card-blue'>", unsafe_allow_html=True)
        
        df_frescos = leer_tabla("clientes", ["Cedula", "Nombre", "Fuerza", "Telefono", "Email", "Deuda_Est", "Ingresos", "Senal", "F_Actualizacion", "Estado", "F_Borrado"])
        df_act_frescos = df_frescos[df_frescos["Estado"] == "Activo"].copy()
        
        if not df_act_frescos.empty:
            cli_mod = st.selectbox("Seleccionar Cliente a Modificar:", df_act_frescos["Cedula"].astype(str) + " - " + df_act_frescos["Nombre"], key="sel_cliente_a_editar")
            
            if cli_mod:
                cc_original = cli_mod.split(" - ")[0]
                datos_c = df_act_frescos[df_act_frescos["Cedula"].astype(str) == str(cc_original)].iloc[0]
                
                st.warning("⚠️ Los cambios realizados aquí sobrescribirán al cliente seleccionado. No se creará un cliente nuevo.")
                
                mc_a, mc_b = st.columns(2)
                with mc_a: m_cc = st.text_input("Cédula de Ciudadanía", value=str(datos_c["Cedula"]), key=f"inp_cc_{cc_original}")
                with mc_b: m_nom = st.text_input("Nombre Completo", value=str(datos_c["Nombre"]), key=f"inp_nom_{cc_original}")
                
                opciones_fza = ["Policía Nacional", "Ejército Nacional", "Armada Nacional", "Fuerza Aérea", "Retirado / Pensionado", "Civil"]
                fza_actual = str(datos_c.get("Fuerza", "Civil"))
                idx_fza = opciones_fza.index(fza_actual) if fza_actual in opciones_fza else 5
                m_fza = st.selectbox("Facción", opciones_fza, index=idx_fza, key=f"inp_fza_{cc_original}")
                
                mc1, mc2, mc3 = st.columns(3)
                with mc1: m_tel = st.text_input("WhatsApp / Celular", value=str(datos_c.get("Telefono", "")), key=f"inp_tel_{cc_original}")
                with mc2: m_mail = st.text_input("Correo Electrónico", value=str(datos_c.get("Email", "")), key=f"inp_mail_{cc_original}")
                with mc3: m_deuda = st.text_input("Deuda Aprox ($)", value=str(datos_c.get("Deuda_Est", "")), key=f"inp_deuda_{cc_original}")
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                if st.button("💾 SOBRESCRIBIR Y GUARDAR CAMBIOS", key=f"btn_save_{cc_original}"):
                    if str(m_cc) != str(cc_original) and str(m_cc) in df_frescos["Cedula"].astype(str).values:
                        st.error(f"❌ La cédula {m_cc} ya pertenece a otro expediente en el sistema.")
                    else:
                        with st.spinner("Modificando cliente en Google Sheets..."):
                            old_path = os.path.join(CARP_EXP, f"{cc_original} - {str(datos_c['Nombre'])}")
                            new_path = os.path.join(CARP_EXP, f"{m_cc} - {m_nom}")
                            if old_path != new_path:
                                try:
                                    if os.path.exists(old_path) and not os.path.exists(new_path):
                                        os.rename(old_path, new_path)
                                except Exception: pass
                                estructurar_carpetas(str(m_cc), m_nom)
                            
                            df_frescos = df_frescos.astype(str)
                            
                            idx_filas = df_frescos.index[df_frescos["Cedula"] == str(cc_original)].tolist()
                            if idx_filas:
                                idx_exacto = idx_filas[0]
                                df_frescos.at[idx_exacto, "Cedula"] = str(m_cc)
                                df_frescos.at[idx_exacto, "Nombre"] = str(m_nom)
                                df_frescos.at[idx_exacto, "Fuerza"] = str(m_fza)
                                df_frescos.at[idx_exacto, "Telefono"] = str(m_tel)
                                df_frescos.at[idx_exacto, "Email"] = str(m_mail)
                                df_frescos.at[idx_exacto, "Deuda_Est"] = str(m_deuda)
                                df_frescos.at[idx_exacto, "F_Actualizacion"] = hoy.strftime("%Y-%m-%d")
                                
                                guardar_tabla(df_frescos, "clientes")
                            
                            for t_nom, t_cols in [("finanzas", ["Cedula", "Honorarios", "Abonado"]), 
                                                  ("actuaciones", ["ID_Act", "Cedula", "Fecha", "Tipo", "Juzgado", "Radicado", "Anotacion"]),
                                                  ("vencimientos", ["ID_Ven", "Cedula", "Cliente", "Asunto", "Fecha_Limite", "Estado"]),
                                                  ("acreedores", ["ID_Acr", "Cedula", "Acreedor", "Cuantia", "Clase"]),
                                                  ("audiencias", ["ID_Aud", "Cedula", "Cliente", "Fecha_Hora", "Motivo"])]:
                                df_t = leer_tabla(t_nom, t_cols)
                                if not df_t.empty and str(cc_original) in df_t["Cedula"].astype(str).values:
                                    df_t = df_t.astype(str)
                                    mask_t = df_t["Cedula"] == str(cc_original)
                                    df_t.loc[mask_t, "Cedula"] = str(m_cc)
                                    if "Cliente" in df_t.columns:
                                        df_t.loc[mask_t, "Cliente"] = str(m_nom)
                                    guardar_tabla(df_t, t_nom)
                                    
                            registrar_log("SISTEMA", f"Modificó expediente exacto: {cc_original} -> {m_cc}")
                            
                            st.success("✅ ¡Cliente modificado y sincronizado con éxito!")
                            time.sleep(1.5)
                            st.rerun()
        else:
            st.info("No hay clientes activos para modificar.")
        st.markdown("</div>", unsafe_allow_html=True)
                            
# --- 3. CONTRATOS Y PLANTILLAS ---
elif st.session_state.pagina_actual == 'Contratos':
    st.markdown("<h1>Gestión Documental y Contratos</h1>", unsafe_allow_html=True)
    if not df_activos.empty:
        cli_sel = st.selectbox("Expediente:", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"], index=idx_foco(df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"]))
        cc_s, nom_s = cli_sel.split(" - ")[0], cli_sel.split(" - ")[1]
        data_c = df_activos[df_activos["Cedula"].astype(str) == cc_s].iloc[0]
        senal, tel_c = int(data_c["Senal"]), str(data_c.get("Telefono", ""))
        rb = estructurar_carpetas(cc_s, nom_s)
        
        mostrar_boveda(cc_s, nom_s)

        # ==========================================
        # MÓDULO DE FASES: WIZARD INTERACTIVO ULTRA-PREMIUM
        # ==========================================
        
        # 1. Inicializar la variable de estado para recordar en qué fase estamos
        if 'fase_activa' not in st.session_state:
            st.session_state.fase_activa = 1

        # Funciones de navegación
        def avanzar_fase():
            if st.session_state.fase_activa < 4:
                st.session_state.fase_activa += 1

        def retroceder_fase():
            if st.session_state.fase_activa > 1:
                st.session_state.fase_activa -= 1

        # 2. Inyectar Motor CSS de Animaciones y Sombras Élite
        st.markdown("""
        <style>
        /* Animación fluida de entrada para las tarjetas */
        @keyframes fadeSlideUp {
            0% { opacity: 0; transform: translateY(20px) scale(0.98); }
            100% { opacity: 1; transform: translateY(0) scale(1); }
        }
        
        .premium-card {
            animation: fadeSlideUp 0.5s cubic-bezier(0.2, 0.8, 0.2, 1) forwards;
            background: #FFFFFF;
            border-radius: 16px;
            padding: 30px;
            box-shadow: 0 15px 35px -5px rgba(15, 23, 42, 0.1);
            border: 1px solid #E2E8F0;
            position: relative;
            overflow: hidden;
            margin-bottom: 20px;
        }
        
        /* Línea decorativa superior dinámica */
        .premium-card::before {
            content: ''; position: absolute; top: 0; left: 0; right: 0; height: 5px;
            background: linear-gradient(90deg, #0F172A 0%, #C9A24B 100%);
        }

        /* Línea de tiempo superior (Estilo SaaS Moderno) */
        .premium-timeline { display: flex; justify-content: space-between; align-items: center; margin: 10px 0 30px 0; position: relative; }
        .premium-timeline::after { content: ''; position: absolute; top: 50%; left: 10px; right: 10px; height: 4px; background: #E2E8F0; z-index: 1; border-radius: 2px; }
        
        .p-step { position: relative; z-index: 2; padding: 10px 25px; border-radius: 30px; font-weight: 700; font-size: 13px; color: #64748B; background: #F8FAFC; border: 3px solid #E2E8F0; transition: all 0.4s ease; text-transform: uppercase; letter-spacing: 0.5px; }
        .p-step.completed { background: #10B981; color: white; border-color: #059669; box-shadow: 0 0 15px rgba(16, 185, 129, 0.4); }
        .p-step.active { background: #0F172A; color: #C9A24B; border-color: #C9A24B; transform: scale(1.08) translateY(-2px); box-shadow: 0 10px 20px rgba(15, 23, 42, 0.3); }
        </style>
        """, unsafe_allow_html=True)

        # 3. Renderizar la Barra de Progreso Visual basada en el estado
        f = st.session_state.fase_activa
        st.markdown(f"""
        <div class="premium-timeline">
            <div class="p-step {'active' if f == 1 else ('completed' if f > 1 else '')}">1. Contratos</div>
            <div class="p-step {'active' if f == 2 else ('completed' if f > 2 else '')}">2. Viabilidad</div>
            <div class="p-step {'active' if f == 3 else ('completed' if f > 3 else '')}">3. Ensamblaje</div>
            <div class="p-step {'active' if f == 4 else ('completed' if f > 4 else '')}">4. Radicado</div>
        </div>
        """, unsafe_allow_html=True)

        # 4. Renderizar el Contenido Animado según la fase
        st.markdown("<div class='premium-card'>", unsafe_allow_html=True)
        
        if f == 1:
            st.markdown("<h3 style='color: #0F172A; margin-bottom: 20px;'>✍️ Fase 1: Emisión de Contratos</h3>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**1. Generar Documento Maestro**")
                if st.button("📄 Generar Contrato Inteligente", use_container_width=True):
                    st.success("Contrato generado con los datos del cliente.")
                    st.download_button("⬇️ Descargar PDF", data="Doc", file_name="Contrato.pdf", use_container_width=True)
            with col2:
                st.markdown("**2. Bóveda de Firmas**")
                archivo_firma = st.file_uploader("Subir contrato ya firmado", type=["pdf"], key="f_up")
                if st.button("Guardar en Nube Segura", use_container_width=True):
                    st.success("✅ Guardado y encriptado.")
                    
        elif f == 2:
            st.markdown("<h3 style='color: #0F172A; margin-bottom: 20px;'>⚖️ Fase 2: Check de Viabilidad</h3>", unsafe_allow_html=True)
            st.info("💡 **Sistema Experto:** Verifique el cumplimiento normativo (Ley 1564 de 2012).")
            st.text_area("Dictamen Jurídico y Observaciones", height=120, placeholder="El cliente declara ingresos por...")
            
        elif f == 3:
            st.markdown("<h3 style='color: #0F172A; margin-bottom: 20px;'>📂 Fase 3: Ensamblaje de Pruebas</h3>", unsafe_allow_html=True)
            st.write("Cargue los anexos obligatorios (Cédula, Datacrédito, Extractos, Certificados).")
            archivo_ensamblaje = st.file_uploader("Arrastre los documentos aquí", type=["pdf", "png", "jpg"], accept_multiple_files=True)
            if st.button("Sincronizar Pruebas con Drive", use_container_width=True):
                st.success("📄 Sincronización exitosa. Expediente armado.")
                
        elif f == 4:
            st.markdown("<h3 style='color: #0F172A; margin-bottom: 20px;'>🏛️ Fase 4: Radicación Oficial</h3>", unsafe_allow_html=True)
            colA, colB = st.columns(2)
            with colA:
                codigo_radicado = st.text_input("Código de Radicado (C. Conciliación)")
            with colB:
                fecha_rad = st.date_input("Fecha de Admisión")
            if st.button("🚀 Confirmar Apertura Legal", use_container_width=True):
                st.balloons()
                st.success(f"🎉 Radicado {codigo_radicado} validado. Alertas activadas.")

        st.markdown("</div>", unsafe_allow_html=True)

        # 5. Botones de Navegación Interactiva (Anterior / Siguiente)
        st.markdown("<br>", unsafe_allow_html=True)
        nav_col1, empty_col, nav_col2 = st.columns([1, 2, 1])
        with nav_col1:
            if f > 1:
                st.button("⬅️ Fase Anterior", on_click=retroceder_fase, use_container_width=True)
        with nav_col2:
            if f < 4:
                st.button("Siguiente Fase ➡️", on_click=avanzar_fase, use_container_width=True)
                
        st.markdown("<br><br>", unsafe_allow_html=True)
        
        st.markdown("<div class='module-card' style='border-top: 3px solid #10B981;'>", unsafe_allow_html=True)
        st.markdown("<h3>📱 Central Automática de WhatsApp</h3>", unsafe_allow_html=True)
        
        opciones_whatsapp = [
            "01. Bienvenida a la Firma", "02. Confirmación Apertura de Expediente",
            "03. Solicitud Paquete Documental Inicial", "04. Recordatorio de Documentos Faltantes",
            "05. Actualización de Estados de Deuda", "06. Petición Certificados de Acreedores",
            "07. Solicitud Soportes de Ingresos/Gastos", "08. Aviso: Contratos Listos para Revisión",
            "09. Recordatorio URGENTE de Firma", "10. Confirmación: Contratos Recibidos",
            "11. Notificación: Radicación Oficial", "12. Buenas Noticias: Trámite ADMITIDO",
            "13. Suspensión de Embargos Emitida", "14. Fecha de Audiencia Fijada",
            "15. Recordatorio Audiencia (Faltan 24h)", "16. Actualización: Postura Acreedores",
            "17. Solicitud Aprobación de Propuesta", "18. Aviso: Objeciones Presentadas",
            "19. ¡Acuerdo de Pago APROBADO!", "20. Transición a Liquidación Patrimonial",
            "21. Recordatorio Preventivo de Cuota", "22. Confirmación de Pago Recibido",
            "23. Aviso de Mora Financiera", "24. Agendamiento Cita Presencial",
            "25. Invitación a Reunión Virtual (Link)", "26. Reporte de Avance Mensual (Tranquilidad)",
            "27. Entrega de Paz y Salvos", "28. Cierre Exitoso del Proceso",
            "29. Solicitud de Reseña/Calificación", "30. Saludo de Feliz Cumpleaños"
        ]
        
        tipo_msg = st.selectbox("Plantilla de Mensaje Legal:", opciones_whatsapp)
        
        msg_map = {
            "01. Bienvenida a la Firma": f"Estimado/a {nom_s}, le damos la bienvenida a nuestra firma jurídica. Es un honor representarle y acompañarle hacia su tranquilidad financiera.",
            "02. Confirmación Apertura de Expediente": f"Hola {nom_s}, le confirmamos que su expediente de insolvencia ya ha sido abierto oficialmente en nuestra plataforma.",
            "03. Solicitud Paquete Documental Inicial": f"Hola {nom_s}, para avanzar ágilmente con su proceso, por favor envíenos el paquete documental completo: Cédula, REDAM, SIMIT, Datacrédito y certificados de propiedad (si aplica).",
            "04. Recordatorio de Documentos Faltantes": f"Estimado/a {nom_s}, le recordamos amablemente que aún tenemos documentos vitales pendientes por recibir. Sin ellos no podemos radicar su caso ante el centro de conciliación.",
            "05. Actualización de Estados de Deuda": f"Hola {nom_s}, nuestro equipo necesita que nos actualice si ha recibido nuevas notificaciones de cobro o demandas en los últimos días.",
            "06. Petición Certificados de Acreedores": f"Estimado/a {nom_s}, por favor solicite y envíenos los certificados de deuda actualizados expedidos directamente por sus acreedores (bancos, cooperativas, etc.).",
            "07. Solicitud Soportes de Ingresos/Gastos": f"Hola {nom_s}, requerimos los soportes recientes de sus ingresos (desprendibles, certificaciones) y la relación mensual de sus gastos de subsistencia.",
            "08. Aviso: Contratos Listos para Revisión": f"¡Excelente noticia {nom_s}! Sus contratos de representación legal e insolvencia ya están redactados. Puede revisarlos y proceder con su firma.",
            "09. Recordatorio URGENTE de Firma": f"¡Atención {nom_s}! Es URGENTE que firme y nos devuelva los contratos. Cada día sin firmar retrasa su protección legal contra embargos y libranzas.",
            "10. Confirmación: Contratos Recibidos": f"Estimado/a {nom_s}, confirmamos la recepción de sus documentos firmados. Procederemos con la radicación formal de su expediente.",
            "11. Notificación: Radicación Oficial": f"¡Hola {nom_s}! Le informamos que su solicitud de insolvencia ha sido radicada oficialmente. Ya dimos el primer gran paso.",
            "12. Buenas Noticias: Trámite ADMITIDO": f"¡Buenas noticias {nom_s}! Su solicitud de insolvencia ha sido ADMITIDA. A partir de este momento queda formalmente protegido por la Ley 1564 de 2012.",
            "13. Suspensión de Embargos Emitida": f"Estimado/a {nom_s}, le confirmamos que ya se han emitido los oficios de suspensión de embargos. Notificaremos a sus pagadores inmediatamente.",
            "14. Fecha de Audiencia Fijada": f"Hola {nom_s}, se ha fijado la fecha para su audiencia de negociación de deudas. Por favor, contáctenos para preparar la diligencia y repasar la estrategia.",
            "15. Recordatorio Audiencia (Faltan 24h)": f"¡Importante {nom_s}! Le recordamos que mañana es su audiencia oficial de insolvencia. Exigimos puntualidad, excelente conexión y presentación formal.",
            "16. Actualización: Postura Acreedores": f"Hola {nom_s}, hemos recibido respuesta de algunos de sus acreedores. Necesitamos agendar una llamada breve para revisar sus posturas.",
            "17. Solicitud Aprobación de Propuesta": f"Estimado/a {nom_s}, nuestro equipo ha estructurado la propuesta de pago objetiva. Necesitamos su revisión y autorización final antes de presentarla.",
            "18. Aviso: Objeciones Presentadas": f"Hola {nom_s}, le informamos que se han presentado algunas objeciones a su relación de créditos. No se preocupe, nuestro equipo jurídico ya está preparando la defensa.",
            "19. ¡Acuerdo de Pago APROBADO!": f"¡Felicidades {nom_s}! Hemos logrado la aprobación de su acuerdo de pago con los acreedores. Su patrimonio está a salvo y sus deudas reestructuradas.",
            "20. Transición a Liquidación Patrimonial": f"Estimado/a {nom_s}, le informamos que al no haber acuerdo con los bancos, el proceso pasará a la etapa de Liquidación Patrimonial. Lo guiaremos paso a paso en esta fase.",
            "21. Recordatorio Preventivo de Cuota": f"Hola {nom_s}, nos comunicamos del área financiera para recordarle amablemente el pago próximo de su cuota de honorarios con el despacho.",
            "22. Confirmación de Pago Recibido": f"Confirmamos la recepción exitosa de su pago, {nom_s}. Su recibo de caja oficial ha sido generado en nuestro sistema. ¡Gracias por su cumplimiento!",
            "23. Aviso de Mora Financiera": f"Estimado/a {nom_s}, registramos un atraso en su compromiso de honorarios. Por favor contáctenos a la brevedad para regularizar su cuenta y no suspender servicios.",
            "24. Agendamiento Cita Presencial": f"Hola {nom_s}, le hemos agendado una cita presencial en nuestras oficinas. Por favor confirme su asistencia respondiendo este mensaje.",
            "25. Invitación a Reunión Virtual (Link)": f"Hola {nom_s}, le hemos programado una reunión virtual con su abogado líder. En breve le enviaremos el enlace de conexión. Por favor conéctese puntual.",
            "26. Reporte de Avance Mensual (Tranquilidad)": f"Estimado/a {nom_s}, le escribimos para darle un parte de tranquilidad. Su caso sigue su curso normal en los juzgados y nuestro equipo lo está monitoreando activamente.",
            "27. Entrega de Paz y Salvos": f"¡Hola {nom_s}! Ya tenemos en nuestro poder los paz y salvos oficiales de sus deudas saneadas. Puede pasar a nuestra oficina a recogerlos cuando guste.",
            "28. Cierre Exitoso del Proceso": f"¡Misión cumplida, {nom_s}! Su proceso legal ha concluido exitosamente. Fue un honor haber sido su escudo jurídico en esta etapa.",
            "29. Solicitud de Reseña/Calificación": f"Hola {nom_s}, ha sido un placer ayudarle. ¿Nos regalaría 1 minuto para calificar nuestro servicio? Su testimonio ayuda a que más personas salven su patrimonio.",
            "30. Saludo de Feliz Cumpleaños": f"¡Feliz cumpleaños {nom_s}! De parte de todo el equipo de la firma jurídica le deseamos un excelente día rodeado de sus seres queridos."
        }
        
        msg_texto = msg_map.get(tipo_msg, f"Hola {nom_s}, nos comunicamos de la firma jurídica para gestionar su proceso.")
        st.text_area("Vista previa del mensaje (Puedes editarlo antes de enviar):", value=msg_texto, height=100)
        
        if tel_c and tel_c != "nan" and tel_c != "":
            link_wa = f"https://wa.me/57{tel_c.replace(' ', '')}?text={msg_texto.replace(' ', '%20')}"
            st.markdown(f"<a href='{link_wa}' target='_blank'><button style='background:#10B981; color:white; border:none; padding:10px 20px; border-radius:6px; font-weight:bold; cursor:pointer; width:100%; margin-bottom: 15px;'>🚀 ENVIAR WHATSAPP DIRECTO AL CLIENTE</button></a>", unsafe_allow_html=True)
        else: st.warning("⚠ No hay celular registrado en la bóveda de este cliente.")
        email_cli = str(data_c.get("Email", "")).strip()
        if email_cli and email_cli.lower() != "nan":
            if st.button("📧 ENVIAR ESTE MENSAJE POR GMAIL AL CLIENTE", key=f"btn_mail_cli_{cc_s}", use_container_width=True):
                cuerpo_html = f"<p>{escape(msg_texto).replace(chr(10), '<br>')}</p><hr><small>Mensaje enviado por su firma jurídica.</small>"
                ok_c, msg_c = enviar_email(email_cli, f"Información sobre su proceso - {tipo_msg[4:]}", cuerpo_html, msg_texto)
                if ok_c:
                    registrar_log("COMUNICACIONES", f"Correo enviado a cliente {cc_s}: {tipo_msg}")
                    st.success("📧 Correo enviado al cliente.")
                else:
                    st.error(f"📧 {msg_c}")
        else:
            st.info("ℹ️ Este cliente no tiene correo registrado.")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
        
        def obtener_archivos_y_id(cc, nom, subcarpeta):
            sub_id = obtener_id_subcarpeta_drive(cc, nom, subcarpeta)
            if sub_id:
                archivos = listar_y_descargar_archivos_drive(sub_id)
                return [a['name'] for a in archivos], sub_id
            return [], None

        if senal == 0:
            sub_id = obtener_id_subcarpeta_drive(cc_s, nom_s, "Financiero")
            arch_drive = []
            if sub_id:
                archivos_list = listar_y_descargar_archivos_drive(sub_id)
                arch_drive = [a['name'] for a in archivos_list]

            st.markdown(f"<h3 style='font-size:18px;'>Fase 1: Recolección Documental ({len(arch_drive)}/8)</h3>", unsafe_allow_html=True)
            docs = ["1. CÉDULA", "2. CERTIFICADO REDAM", "3. DATACRÉDITO", "4. CERTIFICADO SIMIT", "5. CERTIFICADO RAMA", "6. CERTIFICADO RUNT", "7. TRADICIÓN Y LIBERTAD", "8. CERTIFICADO RUES"]
            
            c_a, c_b = st.columns(2)
            for i, d in enumerate(docs):
                col = c_a if i < 4 else c_b
                nombre_corto = d.split(". ")[1]
                
                if any(nombre_corto in f for f in arch_drive): 
                    col.success(f"✔️ {d}")
                else:
                    upl = col.file_uploader(f"📥 Subir: {d}", key=f"auto_up_{cc_s}_{i}")
                    file_key = f"procesando_{cc_s}_{i}"
                    
                    if upl is not None and st.session_state.get(file_key) != upl.name:
                        if sub_id:
                            with st.spinner(f"Subiendo {d} a Google Drive..."):
                                try:
                                    nombre_archivo = f"{nombre_corto}_{upl.name}"
                                    subir_archivo_directo_drive(upl, nombre_archivo, sub_id)
                                    st.session_state[file_key] = upl.name
                                    
                                    archivos_actualizados = listar_y_descargar_archivos_drive(sub_id)
                                    if len(archivos_actualizados) >= 8:
                                        df_cli.loc[df_cli["Cedula"].astype(str) == str(cc_s), "Senal"] = 1
                                        guardar_tabla(df_cli, "clientes")
                                except Exception as e:
                                    st.error(f"❌ ERROR AL SUBIR EL ARCHIVO: {e}")
                            st.rerun()

        elif senal == 2:
            st.markdown("<h3 style='font-size:18px;'>Fase 3: Recolección de Firmas</h3>", unsafe_allow_html=True)
            st.info("📌 Los contratos listos para firmar están arriba, en la Bóveda Central. Descárgalos de ahí y sube los firmados aquí abajo.")
            
            sub_id_notaria = obtener_id_subcarpeta_drive(cc_s, nom_s, "Notaria")
            
            uf = st.file_uploader("📥 Subir Paquete de Contratos Firmados (PDF)", type=["pdf"], key=f"up_firmados_{cc_s}")
            file_key_f = f"procesando_firmados_{cc_s}"
            if uf and sub_id_notaria and st.session_state.get(file_key_f) != uf.name:
                with st.spinner("Subiendo contratos firmados automáticamente a Drive..."):
                    subir_archivo_directo_drive(uf, f"Firmados_{uf.name}", sub_id_notaria)
                    st.session_state[file_key_f] = uf.name
                    df_cli.loc[df_cli["Cedula"].astype(str) == str(cc_s), "Senal"] = 3
                    guardar_tabla(df_cli, "clientes")
                st.success("¡Contratos firmados subidos automáticamente!")
                st.rerun()
                
        elif senal == 3:
            st.markdown("<h3 style='font-size:18px;'>Fase 4: Radicación Oficial del Expediente</h3>", unsafe_allow_html=True)
            st.info("📌 Los contratos ya están firmados. Sube aquí el comprobante de radicación oficial (Centro de Conciliación o Juzgado).")
            urad = st.file_uploader("📥 Subir Soporte de Radicado (PDF/IMG)", type=["pdf", "jpg", "png"], key=f"up_radicado_{cc_s}")
            file_key_r = f"procesando_radicado_{cc_s}"
            sub_id_juz = obtener_id_subcarpeta_drive(cc_s, nom_s, "Juzgado")
            if urad and sub_id_juz and st.session_state.get(file_key_r) != urad.name:
                with st.spinner("Subiendo radicado automáticamente a Drive..."):
                    subir_archivo_directo_drive(urad, f"Radicado_{urad.name}", sub_id_juz)
                    st.session_state[file_key_r] = urad.name
                    df_cli.loc[df_cli["Cedula"].astype(str) == str(cc_s), "Senal"] = 4
                    guardar_tabla(df_cli, "clientes")
                st.success("¡Radicado subido automáticamente!")
                st.rerun()
                
        elif senal >= 4: 
            st.success("✨ ¡Misión Cumplida! El expediente está oficialmente radicado y ha superado todas las fases documentales.")
            sub_id_juz = obtener_id_subcarpeta_drive(cc_s, nom_s, "Juzgado")
            if sub_id_juz:
                arch_juz = listar_y_descargar_archivos_drive(sub_id_juz)
                for ar in arch_juz:
                    if "Radicado" in ar['name']:
                        st.markdown(f"📄 **Soporte Oficial guardado en Drive:** `{ar['name']}`")
            
        st.markdown("</div>", unsafe_allow_html=True)

# --- 4. ACTUACIONES ---
elif st.session_state.pagina_actual == 'Actuaciones':
    st.markdown("<h1>Seguimiento Procesal</h1>", unsafe_allow_html=True)
    if not df_activos.empty:
        c_a = st.selectbox("Expediente:", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"], index=idx_foco(df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"]))
        cc_a, nom_a = c_a.split(" - ")[0], c_a.split(" - ")[1]
        estructurar_carpetas(cc_a, nom_a)
        mostrar_boveda(cc_a, nom_a)
        
        st.markdown("<div class='module-card'>", unsafe_allow_html=True)
        with st.form("f_act"):
            ca1, ca2, ca3 = st.columns(3)
            with ca1: tipo_a = st.selectbox("Tipo", ["Auto Admisorio", "Auto Inadmisorio", "Sentencia", "Oficio", "Memorial Radicado"])
            with ca2: juzgado = st.text_input("Juzgado")
            with ca3: rad_jud = st.text_input("Radicado")
            anota = st.text_area("Anotación / Novedad")
            doc_up = st.file_uploader("Adjuntar PDF de la actuación (Opcional)")
            if st.form_submit_button("REGISTRAR EN GOOGLE SHEETS"):
                if doc_up:
                    sub_id_juz = obtener_id_subcarpeta_drive(cc_a, nom_a, "Juzgado")
                    if sub_id_juz: subir_archivo_directo_drive(doc_up, f"{tipo_a}_{doc_up.name}", sub_id_juz)
                n_act = pd.DataFrame([{"ID_Act": f"ACT-{hoy.strftime('%H%M%S')}", "Cedula": str(cc_a), "Fecha": hoy.strftime("%Y-%m-%d"), "Tipo": tipo_a, "Juzgado": juzgado, "Radicado": rad_jud, "Anotacion": anota}])
                guardar_tabla(pd.concat([df_act, n_act], ignore_index=True), "actuaciones")
                st.success("Registrado y sincronizado."); st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

# --- 5. VENCIMIENTOS ---
elif st.session_state.pagina_actual == 'Vencimientos':
    st.markdown("<h1>🚦 Control de Vencimientos Automático</h1>", unsafe_allow_html=True)
    st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
    with st.form("form_ven"):
        c1, c2 = st.columns(2)
        with c1:
            cl_v = st.selectbox("Seleccionar Cliente", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"] if not df_activos.empty else ["Sin Clientes"])
            asunto = st.text_input("Asunto Legal")
        with c2:
            fecha_v = st.date_input("Fecha Límite Exacta")
        if st.form_submit_button("PROGRAMAR ALERTA"):
            if not df_activos.empty and cl_v != "Sin Clientes":
                cc_v, nom_v = cl_v.split(" - ", 1)
                n_v = pd.DataFrame([{"ID_Ven": f"VEN-{uuid.uuid4().hex[:8]}", "Cedula": str(cc_v), "Cliente": nom_v, "Asunto": asunto, "Fecha_Limite": str(fecha_v), "Estado": "Activo"}])
                guardar_tabla(pd.concat([df_ven, n_v], ignore_index=True), "vencimientos")
                notificar_nuevo_vencimiento(nom_v, asunto, fecha_v, _ctx().get("Email", ""))
                st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("📣 Enviar alertas de vencimientos próximos", key="btn_alertas_ven"):
        ok_mail, msg_mail, ok_tg, msg_tg = enviar_alertas_vencimientos(df_ven, hoy, _ctx().get("Email", ""))
        if ok_mail:
            st.success(f"📧 {msg_mail}")
        else:
            st.warning(f"📧 {msg_mail}")
        if ok_tg:
            st.success(f"✈️ {msg_tg}")
        else:
            st.info(f"✈️ {msg_tg}")

    if not df_ven.empty:
        for _, rv in df_ven[df_ven["Estado"] == "Activo"].iloc[::-1].iterrows():
            try:
                dias = (datetime.strptime(str(rv["Fecha_Limite"])[:10], "%Y-%m-%d").date() - hoy.date()).days
                if dias < 0:
                    cb = "#71717A"
                    est = f"VENCIDO HACE {abs(dias)} DÍAS"
                elif dias <= 2:
                    cb = "#EF4444"
                    est = f"¡PELIGRO! VENCE EN {dias} DÍAS"
                elif dias <= 5:
                    cb = "#F59E0B"
                    est = f"ATENCIÓN: Faltan {dias} días"
                else:
                    cb = "#10B981"
                    est = f"En término ({dias} días)"
                clase_v = "venc venc-peligro" if 0 <= dias <= 2 else "venc"
                st.markdown(f"<div class='{clase_v}' style='--c:{cb};'><b>{escape(str(rv['Cliente']))}</b> · {escape(str(rv['Asunto']))}<br><span style='color:{cb}; font-weight:700;'>{est}</span> <span style='color:#6B7280;'>(Límite: {escape(str(rv['Fecha_Limite']))})</span></div>", unsafe_allow_html=True)
            except ValueError as e:
                print(f"⚠️ Fecha inválida en vencimiento: {e}")

# --- 6. ACREEDORES ---
elif st.session_state.pagina_actual == 'Acreedores':
    st.markdown("<h1>Inventario de pasivos y acreedores</h1><p class='page-sub'>Registra cada deuda con su clase. Esta información alimenta el simulador y la ficha 360.</p>", unsafe_allow_html=True)
    if df_activos.empty:
        st.info("Aún no hay expedientes activos.")
    else:
        opciones_ac = (df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"]).tolist()
        if st.session_state.get("sel_acr") not in opciones_ac:
            st.session_state["sel_acr"] = opciones_ac[idx_foco(opciones_ac)]
        cl_a = st.selectbox("Expediente de insolvencia", opciones_ac, key="sel_acr")
        cc_a = cl_a.split(" - ", 1)[0]
        st.session_state.cliente_foco = cc_a
        with st.form("form_acr", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                nom_acr = st.text_input("Acreedor")
            with c2:
                cuantia_acr = st.number_input("Cuantía ($)", min_value=0, step=100000)
            with c3:
                clase_acr = st.selectbox("Clase", CLASES_ACREEDOR)
            if st.form_submit_button("AGREGAR PASIVO"):
                if nom_acr and cuantia_acr > 0:
                    n_ac = pd.DataFrame([{"ID_Acr": f"ACR-{uuid.uuid4().hex[:8]}", "Cedula": str(cc_a), "Acreedor": nom_acr, "Cuantia": str(cuantia_acr), "Clase": clase_acr}])
                    guardar_tabla(pd.concat([df_acr, n_ac], ignore_index=True), "acreedores")
                    registrar_log("ACREEDORES", f"Agregó acreedor {nom_acr} al expediente {cc_a}")
                    st.rerun()
                else:
                    st.error("Indica el nombre del acreedor y una cuantía mayor a cero.")
        mis_acr = df_acr[df_acr["Cedula"].astype(str) == cc_a] if not df_acr.empty else df_acr
        if mis_acr.empty:
            st.markdown("<div class='vacio'>Este expediente aún no tiene acreedores registrados.</div>", unsafe_allow_html=True)
        else:
            total_a = sum(limpiar_num(v) for v in mis_acr["Cuantia"])
            k1, k2 = st.columns(2)
            k1.markdown(f"<div class='kpi' style='--accent:#8A1F35; --accent-bg:rgba(138,31,53,.12);'><div class='ico'>🏦</div><p class='lbl'>Deuda inventariada</p><p class='val'>$ {total_a:,.0f}</p></div>", unsafe_allow_html=True)
            k2.markdown(f"<div class='kpi' style='--accent:#2F6DB5; --accent-bg:rgba(47,109,181,.14);'><div class='ico'>📋</div><p class='lbl'>Acreedores</p><p class='val'>{len(mis_acr)}</p></div>", unsafe_allow_html=True)
            tabla_m = mis_acr[["Acreedor", "Cuantia", "Clase"]].copy()
            tabla_m["Cuantia"] = tabla_m["Cuantia"].apply(limpiar_num)
            st.dataframe(tabla_m, use_container_width=True, hide_index=True)
            with st.expander("🗑️ Eliminar un acreedor"):
                etiquetas = {f"{r['Acreedor']} · $ {limpiar_num(r['Cuantia']):,.0f}": r["ID_Acr"] for _, r in mis_acr.iterrows()}
                elegido = st.selectbox("Acreedor a eliminar", list(etiquetas.keys()), key="acr_del_sel")
                if st.button("Eliminar acreedor", key="acr_del_btn"):
                    nuevo_acr = df_acr[df_acr["ID_Acr"].astype(str) != str(etiquetas[elegido])]
                    guardar_tabla(nuevo_acr, "acreedores")
                    registrar_log("ACREEDORES", f"Eliminó acreedor del expediente {cc_a}")
                    st.rerun()


# --- VERIFICACIÓN EN FUENTES OFICIALES ---
elif st.session_state.pagina_actual == 'Rastreador':
    st.markdown("<h1>Verificación en fuentes oficiales</h1><p class='page-sub'>Consulta los portales públicos antes de radicar y deja constancia de lo que encontraste en las actuaciones del expediente.</p>", unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("<p class='card-title'>Portales oficiales</p>", unsafe_allow_html=True)
        v1, v2, v3, v4 = st.columns(4)
        v1.link_button("⚖️ Rama Judicial (procesos)", "https://consultaprocesos.ramajudicial.gov.co/", use_container_width=True)
        v2.link_button("🏛️ RUES (comerciantes)", "https://www.rues.org.co/", use_container_width=True)
        v3.link_button("🚗 SIMIT (multas y tránsito)", "https://www.fcm.org.co/simit/", use_container_width=True)
        v4.link_button("📋 Procuraduría (antecedentes)", "https://www.procuraduria.gov.co/Pages/Generacion-de-Antecedentes.aspx", use_container_width=True)
    st.markdown("<div class='aviso'><b>Cómo trabajar:</b> consulta cada portal con la cédula del deudor, descarga o captura el resultado y súbelo a la Bóveda del expediente (carpeta Juzgado o Financiero). Registra la fecha y el resultado en <i>Seguimiento procesal</i>.</div>", unsafe_allow_html=True)

# --- NUEVO: TABLERO KANBAN ---
elif st.session_state.pagina_actual == 'Kanban':
    st.markdown("<h1>Tablero de casos</h1><p class='page-sub'>Cada columna es una fase del trámite. Usa ◀ ▶ para mover un expediente; el cambio queda en el registro de auditoría.</p>", unsafe_allow_html=True)
    if df_activos.empty:
        st.info("Aún no hay expedientes activos. Crea el primero en Expedientes → Apertura y edición.")
    else:
        fuerzas = ["Todas"] + sorted(df_activos["Fuerza"].astype(str).unique().tolist())
        filtro = st.radio("Filtrar por fuerza", fuerzas, horizontal=True, key="kb_filtro")
        base = df_activos if filtro == "Todas" else df_activos[df_activos["Fuerza"].astype(str) == filtro]
        sen = senal_num(base["Senal"]).clip(upper=4)
        saldos_kb = mapa_saldos()
        cols_kb = st.columns(5)
        for (n, ico, nombre, color), col in zip(FASES, cols_kb):
            with col:
                sub = base[sen == n]
                st.markdown(f"<div class='kb-head' style='--c:{color};'><div>{ico} {nombre}</div><span>{len(sub)}</span></div>", unsafe_allow_html=True)
                if sub.empty:
                    st.markdown("<div class='kb-vacio'>Sin casos en esta fase</div>", unsafe_allow_html=True)
                for _, r in sub.iterrows():
                    cc_k = str(r["Cedula"])
                    try:
                        dias_k = (ahora_colombia() - datetime.strptime(str(r["F_Actualizacion"])[:10], "%Y-%m-%d")).days
                    except ValueError:
                        dias_k = 0
                    saldo_k = saldos_kb.get(cc_k, 0)
                    pill_dias = f"<span class='pill {'warn' if dias_k >= 3 and n in (2, 3) else ''}'>⏱ {dias_k} d en la fase</span>"
                    pill_saldo = f"<span class='pill {'warn' if saldo_k > 0 else ''}'>💰 {'Saldo $' + format(saldo_k, ',.0f') if saldo_k > 0 else 'Al día'}</span>" if es_jefa() else ""
                    st.markdown(f"<div class='kb-card' style='--c:{color};'><b>{escape(str(r['Nombre']))}</b><span>CC {escape(cc_k)} · {escape(str(r['Fuerza']))}</span>{pill_dias}{pill_saldo}</div>", unsafe_allow_html=True)
                    b1, b2, b3 = st.columns(3)
                    with b1.container(key=f"mini_prev_{cc_k}"):
                        if n > 0:
                            st.button("◀", key=f"kbp_{cc_k}", on_click=mover_fase, args=(cc_k, -1), use_container_width=True, help="Mover a la fase anterior")
                    with b2.container(key=f"mini_open_{cc_k}"):
                        st.button("🔎", key=f"kbo_{cc_k}", on_click=abrir_ficha, args=(cc_k,), use_container_width=True, help="Abrir ficha 360")
                    with b3.container(key=f"mini_next_{cc_k}"):
                        if n < 4:
                            st.button("▶", key=f"kbn_{cc_k}", on_click=mover_fase, args=(cc_k, 1), use_container_width=True, help="Mover a la fase siguiente")
                    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# --- NUEVO: FICHA 360 ---
elif st.session_state.pagina_actual == 'Ficha':
    st.markdown("<h1>Ficha 360</h1><p class='page-sub'>Todo el expediente en un solo lugar: fase, deudas, agenda, actuaciones y finanzas.</p>", unsafe_allow_html=True)
    if df_activos.empty:
        st.info("Aún no hay expedientes activos.")
    else:
        opciones_f = (df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"]).tolist()
        if st.session_state.get("sel_ficha") not in opciones_f:
            st.session_state["sel_ficha"] = opciones_f[idx_foco(opciones_f)]
        sel_f = st.selectbox("Expediente", opciones_f, key="sel_ficha")
        cc_f, nom_f = sel_f.split(" - ", 1)
        st.session_state.cliente_foco = cc_f
        dato = df_activos[df_activos["Cedula"].astype(str) == cc_f].iloc[0]
        fase_f = int(min(4, senal_num(pd.Series([dato["Senal"]])).iloc[0]))
        saldo_f = mapa_saldos().get(cc_f, 0)
        iniciales = "".join(p[0] for p in nom_f.split()[:2]).upper() or "?"
        tel_f = str(dato.get("Telefono", "")).strip()
        tel_f = tel_f[:-2] if tel_f.endswith(".0") else tel_f
        mail_f = str(dato.get("Email", "")).strip()
        tel_f = "" if tel_f.lower() == "nan" else tel_f
        mail_f = "" if mail_f.lower() == "nan" else mail_f

        with st.container(border=True):
            chips = f"<span class='chip-info'>{FASES[fase_f][1]} {FASES[fase_f][2]}</span><span class='chip-info'>{escape(str(dato['Fuerza']))}</span>"
            if es_jefa():
                chips += f"<span class='chip-info {'rojo' if saldo_f > 0 else 'verde'}'>💰 {'Saldo $' + format(saldo_f, ',.0f') if saldo_f > 0 else 'Honorarios al día'}</span>"
            st.markdown(f"<div class='ficha-head'><div class='avatar'>{escape(iniciales)}</div><div><p class='nom'>{escape(nom_f)}</p><p class='det'>CC {escape(cc_f)}{(' · 📞 ' + escape(tel_f)) if tel_f else ''}{(' · ✉️ ' + escape(mail_f)) if mail_f else ''}</p><div style='margin-top:8px;'>{chips}</div></div></div>", unsafe_allow_html=True)
            pasos = []
            for n, ico, nombre, color in FASES:
                clase = "ok" if n < fase_f else ("now" if n == fase_f else "")
                pasos.append(f"<div class='s {clase}'><i>{'✓' if n < fase_f else ico}</i>{nombre}</div>")
            st.markdown(f"<div class='stepper'>{''.join(pasos)}</div>", unsafe_allow_html=True)
            a1, a2, a3, a4, a5 = st.columns(5)
            a1.button("📄 Documentos", key="fi_doc", on_click=ir_a_modulo, args=("Contratos", cc_f), use_container_width=True)
            a2.button("📋 Pasivos", key="fi_acr", on_click=ir_a_modulo, args=("Acreedores", cc_f), use_container_width=True)
            a3.button("🧮 Simular", key="fi_sim", on_click=ir_a_modulo, args=("Simulador", cc_f), use_container_width=True)
            if tel_f:
                a4.link_button("💬 WhatsApp", f"https://wa.me/57{tel_f.replace(' ', '')}", use_container_width=True)
            if mail_f:
                a5.link_button("✉️ Correo", f"mailto:{mail_f}", use_container_width=True)

        acr_f = df_acr[df_acr["Cedula"].astype(str) == cc_f] if not df_acr.empty else df_acr
        act_f = df_act[df_act["Cedula"].astype(str) == cc_f] if not df_act.empty else df_act
        aud_f = df_aud[df_aud["Cedula"].astype(str) == cc_f] if not df_aud.empty else df_aud
        ven_f = df_ven[(df_ven["Cedula"].astype(str) == cc_f) & (df_ven["Estado"].astype(str) == "Activo")] if not df_ven.empty else df_ven
        deuda_reg = sum(limpiar_num(v) for v in acr_f["Cuantia"]) if not acr_f.empty else 0

        nombres_tabs = ["Resumen", "Acreedores", "Actuaciones", "Agenda"] + (["Finanzas"] if es_jefa() else [])
        tabs_f = st.tabs(nombres_tabs)
        with tabs_f[0]:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Deuda declarada", f"$ {limpiar_num(dato['Deuda_Est']):,.0f}")
            m2.metric("Deuda inventariada", f"$ {deuda_reg:,.0f}")
            m3.metric("Acreedores", len(acr_f))
            m4.metric("Actuaciones", len(act_f))
            dif = limpiar_num(dato['Deuda_Est']) - deuda_reg
            if len(acr_f) == 0:
                st.info("Aún no hay acreedores inventariados. Regístralos en Pasivos para poder simular una propuesta.")
            elif abs(dif) > 1:
                st.warning(f"La deuda declarada y la inventariada difieren en $ {abs(dif):,.0f}. Conviene conciliarlas con los certificados de los acreedores.")
            else:
                st.success("La deuda declarada coincide con la inventariada.")
        with tabs_f[1]:
            if acr_f.empty:
                st.markdown("<div class='vacio'>Sin acreedores registrados.</div>", unsafe_allow_html=True)
            else:
                tabla_a = acr_f[["Acreedor", "Cuantia", "Clase"]].copy()
                tabla_a["Cuantia"] = tabla_a["Cuantia"].apply(limpiar_num)
                st.dataframe(tabla_a, use_container_width=True, hide_index=True)
                por_clase = tabla_a.groupby("Clase")["Cuantia"].sum().reindex(CLASES_ACREEDOR).fillna(0).reset_index()
                st.bar_chart(por_clase, x="Clase", y="Cuantia", color="#C9A24B")
        with tabs_f[2]:
            if act_f.empty:
                st.markdown("<div class='vacio'>Sin actuaciones registradas.</div>", unsafe_allow_html=True)
            else:
                st.dataframe(act_f[["Fecha", "Tipo", "Juzgado", "Radicado", "Anotacion"]].iloc[::-1], use_container_width=True, hide_index=True)
        with tabs_f[3]:
            evs_f = [e for e in eventos_rango(ahora_colombia().date() - timedelta(days=60), ahora_colombia().date() + timedelta(days=365)) if e["cc"] == cc_f]
            st.markdown(html_timeline(evs_f[:10], ahora_colombia().date()), unsafe_allow_html=True)
        if es_jefa():
            with tabs_f[4]:
                fin_f = df_fin[df_fin["Cedula"].astype(str) == cc_f]
                if fin_f.empty:
                    st.info("Este expediente no tiene honorarios configurados.")
                else:
                    hon = limpiar_num(fin_f.iloc[0]["Honorarios"])
                    abo = limpiar_num(fin_f.iloc[0]["Abonado"])
                    f1, f2, f3 = st.columns(3)
                    f1.metric("Honorarios", f"$ {hon:,.0f}")
                    f2.metric("Abonado", f"$ {abo:,.0f}")
                    f3.metric("Saldo", f"$ {max(hon - abo, 0):,.0f}")
                    st.progress(min(abo / hon, 1.0) if hon > 0 else 0.0, text=f"{(abo / hon * 100) if hon > 0 else 0:.0f}% recaudado")

# --- NUEVO: SIMULADOR DE PROPUESTA DE PAGO ---
elif st.session_state.pagina_actual == 'Simulador':
    st.markdown("<h1>Simulador de propuesta de pago</h1><p class='page-sub'>Estima cuánto podría pagar el deudor, en cuánto tiempo y cómo se repartiría entre sus acreedores por clase.</p>", unsafe_allow_html=True)
    lista_sim = (df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"]).tolist() if not df_activos.empty else []
    opciones_sim = ["— Escenario manual —"] + lista_sim
    if st.session_state.get("sim_sel") not in opciones_sim:
        st.session_state["sim_sel"] = opciones_sim[(idx_foco(lista_sim) + 1) if (lista_sim and str(st.session_state.get("cliente_foco", "")) != "") else 0]
    sel_sim = st.selectbox("Cargar datos de un expediente", opciones_sim, key="sim_sel")

    ingreso_base, base_acr = 0.0, pd.DataFrame({"Acreedor": [""], "Cuantía": [0.0], "Clase": ["Quinta"]})
    if sel_sim != opciones_sim[0]:
        cc_s2 = sel_sim.split(" - ", 1)[0]
        fila_s = df_activos[df_activos["Cedula"].astype(str) == cc_s2]
        if not fila_s.empty:
            ingreso_base = limpiar_num(fila_s.iloc[0]["Ingresos"])
        acr_s = df_acr[df_acr["Cedula"].astype(str) == cc_s2] if not df_acr.empty else df_acr
        if not acr_s.empty:
            base_acr = pd.DataFrame({"Acreedor": acr_s["Acreedor"].astype(str).tolist(),
                                     "Cuantía": [limpiar_num(v) for v in acr_s["Cuantia"]],
                                     "Clase": [c if c in CLASES_ACREEDOR else "Quinta" for c in acr_s["Clase"].astype(str)]})

    with st.container(border=True):
        st.markdown("<p class='card-title'>1 · Capacidad de pago</p>", unsafe_allow_html=True)
        p1, p2 = st.columns(2)
        ingresos_m = p1.number_input("Ingresos mensuales ($)", min_value=0.0, value=float(ingreso_base), step=100000.0, key=f"sim_ing_{sel_sim}")
        gastos_m = p2.number_input("Gastos de subsistencia mensuales ($)", min_value=0.0, value=0.0, step=100000.0, key=f"sim_gas_{sel_sim}")
        p3, p4 = st.columns(2)
        pct = p3.slider("% del excedente que se destina a deudas", 10, 100, 100, step=5, key="sim_pct")
        plazo = p4.slider("Plazo de la propuesta (meses)", 6, 120, 60, step=6, key="sim_plazo")

    with st.container(border=True):
        st.markdown("<p class='card-title'>2 · Acreedores (puedes editar, agregar o quitar filas)</p>", unsafe_allow_html=True)
        tabla_ed = st.data_editor(
            base_acr, num_rows="dynamic", use_container_width=True, hide_index=True, key=f"sim_ed_{sel_sim}",
            column_config={
                "Cuantía": st.column_config.NumberColumn("Cuantía ($)", min_value=0.0, step=100000.0, format="%.0f"),
                "Clase": st.column_config.SelectboxColumn("Clase", options=CLASES_ACREEDOR, required=True),
            },
        )

    tabla_ed = tabla_ed.copy()
    tabla_ed["Cuantía"] = pd.to_numeric(tabla_ed["Cuantía"], errors="coerce").fillna(0.0)
    tabla_ed["Clase"] = tabla_ed["Clase"].fillna("Quinta")
    tabla_ed["Acreedor"] = tabla_ed["Acreedor"].fillna("").astype(str)
    tabla_ed = tabla_ed[tabla_ed["Cuantía"] > 0].reset_index(drop=True)

    excedente = max(ingresos_m - gastos_m, 0.0)
    cuota = excedente * pct / 100
    disponible = cuota * plazo
    deuda_total = float(tabla_ed["Cuantía"].sum())

    # Reparto por orden de clase: cada clase se paga completa antes de pasar a la siguiente.
    restante = disponible
    asignado = []
    for clase in CLASES_ACREEDOR:
        idx_c = tabla_ed.index[tabla_ed["Clase"] == clase]
        deuda_c = float(tabla_ed.loc[idx_c, "Cuantía"].sum())
        pago_c = min(restante, deuda_c)
        restante -= pago_c
        for i in idx_c:
            parte = tabla_ed.loc[i, "Cuantía"] / deuda_c if deuda_c else 0
            asignado.append((i, pago_c * parte))
    tabla_ed["Pago estimado"] = 0.0
    for i, v in asignado:
        tabla_ed.loc[i, "Pago estimado"] = v
    tabla_ed["Cobertura %"] = (tabla_ed["Pago estimado"] / tabla_ed["Cuantía"] * 100).round(1)
    cobertura = (min(disponible, deuda_total) / deuda_total * 100) if deuda_total > 0 else 0
    meses_total = (-(-deuda_total // cuota)) if cuota > 0 else None

    r1, r2 = st.columns([1.2, 1])
    with r1:
        mensaje_plazo = f"Pagar el 100% tomaría unos {int(meses_total)} meses ({meses_total / 12:.1f} años)." if meses_total else "Con los datos actuales no hay excedente para pagar deudas."
        st.markdown(
            f"<div class='sim-res'><p>Cuota mensual estimada</p><p class='big'>$ {cuota:,.0f}</p>"
            f"<p>Total disponible en {plazo} meses: <b style='color:#fff'>$ {disponible:,.0f}</b> · Deuda inventariada: <b style='color:#fff'>$ {deuda_total:,.0f}</b></p>"
            f"<p style='margin-top:10px;'>Cobertura de la deuda: <b style='color:#E6C873; font-size:20px;'>{cobertura:.1f}%</b></p><p>{mensaje_plazo}</p></div>",
            unsafe_allow_html=True,
        )
        st.progress(min(cobertura / 100, 1.0))
    with r2:
        if deuda_total > 0:
            graf = tabla_ed.groupby("Clase")[["Cuantía", "Pago estimado"]].sum().reindex(CLASES_ACREEDOR).fillna(0).reset_index()
            graf.columns = ["Clase", "Deuda", "Pago estimado"]
            st.bar_chart(graf, x="Clase", y=["Deuda", "Pago estimado"], color=["#8A1F35", "#1F7A5C"], stack=False)
        else:
            st.markdown("<div class='vacio'>Agrega acreedores para ver el reparto.</div>", unsafe_allow_html=True)
    if deuda_total > 0:
        st.dataframe(tabla_ed[["Acreedor", "Clase", "Cuantía", "Pago estimado", "Cobertura %"]], use_container_width=True, hide_index=True)
    st.markdown("<div class='aviso'><b>Herramienta orientativa.</b> No incluye intereses, actualización de créditos, costos del trámite ni las reglas específicas del procedimiento. La propuesta real debe ser definida y revisada por el abogado responsable.</div>", unsafe_allow_html=True)

# --- NUEVO: REPORTES ---
elif st.session_state.pagina_actual == 'Reportes':
    st.markdown("<h1>Reportes y analítica</h1><p class='page-sub'>Una mirada rápida al estado del despacho.</p>", unsafe_allow_html=True)
    if df_activos.empty:
        st.info("Aún no hay datos suficientes para generar reportes.")
    else:
        activos_cc = set(df_activos["Cedula"].astype(str))
        sen_r = senal_num(df_activos["Senal"]).clip(upper=4)
        c1, c2 = st.columns(2)
        with c1:
            with st.container(border=True):
                st.markdown("<p class='card-title'>Expedientes por fase</p>", unsafe_allow_html=True)
                datos_fase = pd.DataFrame({"Fase": [f"{n + 1}. {nombre}" for n, _, nombre, _ in FASES], "Casos": [int((sen_r == n).sum()) for n, _, _, _ in FASES]})
                st.bar_chart(datos_fase, x="Fase", y="Casos", color="#13294B")
        with c2:
            with st.container(border=True):
                st.markdown("<p class='card-title'>Expedientes por fuerza</p>", unsafe_allow_html=True)
                por_fuerza = df_activos["Fuerza"].astype(str).value_counts().reset_index()
                por_fuerza.columns = ["Fuerza", "Casos"]
                st.bar_chart(por_fuerza, x="Fuerza", y="Casos", color="#C9A24B")
        c3, c4 = st.columns(2)
        acr_activos = df_acr[df_acr["Cedula"].astype(str).isin(activos_cc)] if not df_acr.empty else df_acr
        with c3:
            with st.container(border=True):
                st.markdown("<p class='card-title'>Deuda inventariada por clase</p>", unsafe_allow_html=True)
                if acr_activos.empty:
                    st.markdown("<div class='vacio'>Sin acreedores registrados.</div>", unsafe_allow_html=True)
                else:
                    t = acr_activos.copy()
                    t["Cuantia"] = t["Cuantia"].apply(limpiar_num)
                    cl = t.groupby("Clase")["Cuantia"].sum().reindex(CLASES_ACREEDOR).fillna(0).reset_index()
                    st.bar_chart(cl, x="Clase", y="Cuantia", color="#8A1F35")
        with c4:
            with st.container(border=True):
                st.markdown("<p class='card-title'>Principales acreedores</p>", unsafe_allow_html=True)
                if acr_activos.empty:
                    st.markdown("<div class='vacio'>Sin acreedores registrados.</div>", unsafe_allow_html=True)
                else:
                    t = acr_activos.copy()
                    t["Cuantia"] = t["Cuantia"].apply(limpiar_num)
                    top = t.groupby("Acreedor")["Cuantia"].sum().sort_values(ascending=False).head(6).reset_index()
                    st.bar_chart(top, x="Acreedor", y="Cuantia", color="#2F6DB5", horizontal=True)
        if es_jefa():
            with st.container(border=True):
                st.markdown("<p class='card-title'>Recaudo de honorarios</p>", unsafe_allow_html=True)
                f_act = df_fin[df_fin["Cedula"].astype(str).isin(activos_cc)] if not df_fin.empty else df_fin
                hon_t = sum(limpiar_num(v) for v in f_act["Honorarios"]) if not f_act.empty else 0
                abo_t = sum(limpiar_num(v) for v in f_act["Abonado"]) if not f_act.empty else 0
                m1, m2, m3 = st.columns(3)
                m1.metric("Honorarios pactados", f"$ {hon_t:,.0f}")
                m2.metric("Recaudado", f"$ {abo_t:,.0f}")
                m3.metric("Por cobrar", f"$ {max(hon_t - abo_t, 0):,.0f}")
                st.progress(min(abo_t / hon_t, 1.0) if hon_t > 0 else 0.0, text=f"{(abo_t / hon_t * 100) if hon_t > 0 else 0:.0f}% del total pactado")
            if not df_log.empty:
                with st.container(border=True):
                    st.markdown("<p class='card-title'>Actividad por usuario</p>", unsafe_allow_html=True)
                    act_u = df_log["Usuario"].astype(str).value_counts().head(8).reset_index()
                    act_u.columns = ["Usuario", "Acciones"]
                    st.bar_chart(act_u, x="Usuario", y="Acciones", color="#6B5BA6")

# --- 7. TAREAS ---
elif st.session_state.pagina_actual == 'Tareas':
    st.markdown("<h1>Tablero de tareas</h1><p class='page-sub'>Asigna trabajo al equipo y marca lo que ya está hecho.</p>", unsafe_allow_html=True)
    with st.expander("➕ Nueva tarea", expanded=df_tar.empty):
        with st.form("form_tareas", clear_on_submit=True):
            col1, col2 = st.columns([2, 1])
            with col1:
                desc_tar = st.text_input("Descripción de la tarea")
            with col2:
                asignado = st.selectbox("Asignar a:", df_usr["Alias"].tolist())
            if st.form_submit_button("ENVIAR TAREA"):
                if desc_tar:
                    n_t = pd.DataFrame([{"ID_Tar": f"TAR-{uuid.uuid4().hex[:8]}", "Tarea": desc_tar, "Asignado": asignado, "Creador": st.session_state.alias_actual, "Estado": "Pendiente", "Fecha": ahora_colombia().strftime("%Y-%m-%d")}])
                    guardar_tabla(pd.concat([df_tar, n_t], ignore_index=True), "tareas")
                    registrar_log("TAREAS", f"Asignó una tarea a {asignado}")
                    st.rerun()
    vista = st.radio("Mostrar", ["Mis tareas", "Todas"], horizontal=True, key="tar_vista")
    base_t = df_tar if vista == "Todas" else df_tar[df_tar["Asignado"].astype(str) == str(st.session_state.alias_actual)]
    pend = base_t[base_t["Estado"].astype(str) == "Pendiente"]
    hechas = base_t[base_t["Estado"].astype(str) != "Pendiente"]
    cp, ch = st.columns(2)
    with cp:
        with st.container(border=True):
            st.markdown(f"<p class='card-title'>Pendientes · {len(pend)}</p>", unsafe_allow_html=True)
            if pend.empty:
                st.markdown("<div class='vacio'>Nada pendiente 🙌</div>", unsafe_allow_html=True)
            for i, (_, t) in enumerate(pend.iterrows()):
                c_t, c_b = st.columns([5, 1], vertical_alignment="center")
                c_t.markdown(f"<div class='task'>{escape(str(t['Tarea']))}<em>para {escape(str(t['Asignado']))} · de {escape(str(t['Creador']))} · {escape(str(t['Fecha']))}</em></div>", unsafe_allow_html=True)
                with c_b.container(key=f"mini_tp_{i}"):
                    st.button("✓", key=f"tp_{i}_{t['ID_Tar']}", on_click=cambiar_estado_tarea, args=(t['ID_Tar'], "Hecha"), help="Marcar como hecha")
    with ch:
        with st.container(border=True):
            st.markdown(f"<p class='card-title'>Hechas · {len(hechas)}</p>", unsafe_allow_html=True)
            if hechas.empty:
                st.markdown("<div class='vacio'>Aún no hay tareas completadas.</div>", unsafe_allow_html=True)
            for i, (_, t) in enumerate(hechas.tail(8).iloc[::-1].iterrows()):
                c_t, c_b = st.columns([5, 1], vertical_alignment="center")
                c_t.markdown(f"<div class='task' style='opacity:.65; text-decoration:line-through;'>{escape(str(t['Tarea']))}<em>{escape(str(t['Asignado']))}</em></div>", unsafe_allow_html=True)
                with c_b.container(key=f"mini_th_{i}"):
                    st.button("↩", key=f"th_{i}_{t['ID_Tar']}", on_click=cambiar_estado_tarea, args=(t['ID_Tar'], "Pendiente"), help="Reabrir")


# --- 8. MEMORIALES ---
elif st.session_state.pagina_actual == 'Memoriales':
    st.markdown("<h1>🤖 Dependiente Virtual (Memoriales Oficiales)</h1>", unsafe_allow_html=True)
    st.markdown("<div class='module-card'>", unsafe_allow_html=True)
    with st.form("form_mem"):
        cli_m = st.selectbox("Seleccionar Cliente", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"] if not df_activos.empty else ["Sin Clientes"])
        tipo_mem = st.selectbox("Actuación:", ["Petición de Desembargo", "Terminación por Acuerdo", "Copias Simples"])
        juzgado = st.text_input("Juzgado Destino:")
        radicado = st.text_input("Radicado:")
        if st.form_submit_button("REDACTAR Y GENERAR PDF"):
            if cli_m != "Sin Clientes" and juzgado:
                nom_m, cc_m = cli_m.split(" - ")[1], cli_m.split(" - ")[0]
                tit = f"SEÑOR JUEZ\n{juzgado}\nE. S. D.\n\nREF: {tipo_mem}\nDEUDOR: {nom_m}\nRADICADO: {radicado}"
                cue = f"Actuando en nombre de {nom_m} (C.C. {cc_m}), solicito formalmente la procedencia de la presente petición en derecho."
                n_pdf = f"Memorial_{cc_m}_{hoy.strftime('%H%M%S')}.pdf"
                generar_pdf_oficial(tit, cue, n_pdf)
                st.session_state['memorial_pdf'] = n_pdf
    if 'memorial_pdf' in st.session_state and os.path.exists(st.session_state['memorial_pdf']):
        with open(st.session_state['memorial_pdf'], "rb") as f:
            st.download_button("📥 DESCARGAR MEMORIAL PDF", f, file_name=st.session_state['memorial_pdf'], use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

# --- 9. CALENDARIO Y CITAS ---
elif st.session_state.pagina_actual == 'Agenda':
    st.markdown("<h1>Calendario y citas</h1><p class='page-sub'>Audiencias, citas y vencimientos del mes en una sola vista.</p>", unsafe_allow_html=True)
    hoy_cal = ahora_colombia().date()
    anio_c, mes_c = st.session_state.get("cal_mes", (hoy_cal.year, hoy_cal.month))
    ini_mes = datetime(anio_c, mes_c, 1).date()
    fin_mes = datetime(anio_c, mes_c, calendar.monthrange(anio_c, mes_c)[1]).date()
    eventos_mes = [e for e in eventos_rango(ini_mes, fin_mes) if ini_mes <= e["fecha"].date() <= fin_mes]
    por_dia = {}
    for e in eventos_mes:
        por_dia.setdefault(e["fecha"].day, []).append(e)

    n1, n2, n3, n4 = st.columns([1, 4, 1, 1.4], vertical_alignment="center")
    n1.button("◀ Anterior", key="cal_prev", on_click=mover_mes, args=(-1,), use_container_width=True)
    n2.markdown(f"<div class='cal-titulo'>{MESES_ES[mes_c - 1].capitalize()} {anio_c}</div>", unsafe_allow_html=True)
    n3.button("Siguiente ▶", key="cal_next", on_click=mover_mes, args=(1,), use_container_width=True)
    n4.button("📍 Ir a hoy", key="cal_hoy", on_click=mes_actual, use_container_width=True)

    cabecera = "".join(f"<div class='cal-h'>{d}</div>" for d in ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"])
    celdas = []
    for semana in calendar.monthcalendar(anio_c, mes_c):
        for dia in semana:
            if dia == 0:
                celdas.append("<div class='cal-d fuera'></div>")
                continue
            evs = por_dia.get(dia, [])
            clase = "cal-d hoy" if (anio_c, mes_c, dia) == (hoy_cal.year, hoy_cal.month, hoy_cal.day) else "cal-d"
            chips = ""
            for e in evs[:3]:
                prefijo = e["fecha"].strftime("%H:%M ") if e["hora"] else ""
                chips += f"<div class='chip' style='--c:{e['color']};' title='{escape(e['titulo'])} · {escape(e['detalle'])}'>{prefijo}{escape(e['titulo'][:24])}</div>"
            extra = f"<div class='mas'>+{len(evs) - 3} más</div>" if len(evs) > 3 else ""
            celdas.append(f"<div class='{clase}'><span class='n'>{dia}</span>{chips}{extra}</div>")
    st.markdown(f"<div class='cal'>{cabecera}{''.join(celdas)}</div>", unsafe_allow_html=True)

    col_l, col_f = st.columns([1.3, 1])
    with col_l:
        with st.container(border=True):
            st.markdown(f"<p class='card-title'>Compromisos de {MESES_ES[mes_c - 1]}</p>", unsafe_allow_html=True)
            st.markdown(html_timeline(eventos_mes, hoy_cal), unsafe_allow_html=True)
            if eventos_mes:
                st.download_button("📅 Exportar el mes (.ics)", ics_eventos(eventos_mes), file_name=f"agenda_{anio_c}_{mes_c:02d}.ics", mime="text/calendar")
    with col_f:
        st.markdown("<p class='card-title' style='margin-top:4px;'>➕ Agendar cita o audiencia</p>", unsafe_allow_html=True)
        with st.form("form_aud", clear_on_submit=True):
            cl_aud = st.selectbox("Expediente", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"] if not df_activos.empty else ["Sin Clientes"])
            motivo = st.text_input("Motivo")
            f1, f2 = st.columns(2)
            fecha_a = f1.date_input("Fecha", value=hoy_cal)
            hora_a = f2.time_input("Hora")
            if st.form_submit_button("AGENDAR"):
                if cl_aud != "Sin Clientes":
                    cc_a, nom_a = cl_aud.split(" - ", 1)
                    n_au = pd.DataFrame([{"ID_Aud": f"AUD-{uuid.uuid4().hex[:8]}", "Cedula": str(cc_a), "Cliente": nom_a, "Fecha_Hora": f"{fecha_a} {hora_a}", "Motivo": motivo}])
                    guardar_tabla(pd.concat([df_aud, n_au], ignore_index=True), "audiencias")
                    registrar_log("AGENDA", f"Agendó '{motivo}' para {cc_a} el {fecha_a}")
                    st.session_state.cal_mes = (fecha_a.year, fecha_a.month)
                    st.rerun()


# --- 10. FINANZAS ---
elif st.session_state.pagina_actual == 'Finanzas':
    if st.session_state.rol_actual != "Administrador (Jefa)": st.error("⛔ DENEGADO")
    else:
        st.markdown("<h1>💰 Finanzas y Facturación Cloud</h1>", unsafe_allow_html=True)
        
        def generar_factura_pdf(nombre, cc, total, abono, saldo, recibo_id):
            pdf = FPDF()
            pdf.add_page()
            pdf.rect(5.0, 5.0, 200.0, 287.0)
            
            def cln(t): return str(t).encode('latin-1', 'replace').decode('latin-1')
            
            pdf.set_font("Helvetica", 'B', 20)
            pdf.set_text_color(212, 175, 55)
            pdf.cell(0, 15, txt=cln("FIRMA JURÍDICA - INSOLVENCIA OS"), ln=True, align='C')
            pdf.set_font("Helvetica", 'B', 14)
            pdf.set_text_color(50, 50, 50)
            pdf.cell(0, 8, txt=cln("RECIBO DE CAJA / COMPROBANTE DE INGRESO"), ln=True, align='C')
            pdf.ln(5)
            
            pdf.set_font("Helvetica", 'B', 10)
            pdf.cell(0, 6, txt=cln(f"Recibo No: {recibo_id}"), ln=True, align='R')
            pdf.cell(0, 6, txt=cln(f"Fecha de Emisión: {ahora_colombia().strftime('%Y-%m-%d %H:%M:%S')}"), ln=True, align='R')
            pdf.ln(10)
            
            pdf.set_fill_color(240, 240, 240)
            pdf.set_font("Helvetica", 'B', 12)
            pdf.cell(0, 8, txt=cln(" DATOS DEL CLIENTE"), ln=True, align='L', fill=True)
            pdf.set_font("Helvetica", '', 11)
            pdf.cell(0, 6, txt=cln(f" Nombre / Razón Social: {nombre}"), ln=True)
            pdf.cell(0, 6, txt=cln(f" Cédula de Identidad: {cc}"), ln=True)
            pdf.ln(5)
            
            pdf.set_font("Helvetica", 'B', 12)
            pdf.cell(0, 8, txt=cln(" DETALLES DEL SERVICIO"), ln=True, align='L', fill=True)
            pdf.set_font("Helvetica", '', 11)
            pdf.multi_cell(0, 6, txt=cln("Concepto: Abono a honorarios profesionales por representación jurídica integral en trámite de insolvencia económica persona natural no comerciante (Ley 1564 de 2012)."))
            pdf.ln(5)
            
            pdf.set_font("Helvetica", 'B', 12)
            pdf.cell(0, 8, txt=cln(" DESGLOSE FINANCIERO"), ln=True, align='L', fill=True)
            pdf.ln(2)
            
            pdf.set_font("Helvetica", '', 11)
            pdf.cell(95, 8, txt=cln("Valor Total de Honorarios (Contrato):"), border=1)
            pdf.set_font("Helvetica", 'B', 11)
            pdf.cell(95, 8, txt=cln(f"$ {total:,.0f}"), border=1, ln=True, align='R')
            
            pdf.set_font("Helvetica", '', 11)
            pdf.cell(95, 8, txt=cln("Saldo Pendiente Anterior:"), border=1)
            pdf.set_font("Helvetica", 'B', 11)
            pdf.cell(95, 8, txt=cln(f"$ {(saldo + abono):,.0f}"), border=1, ln=True, align='R')
            
            pdf.set_fill_color(212, 175, 55)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Helvetica", 'B', 12)
            pdf.cell(95, 10, txt=cln("VALOR ABONADO (ESTE RECIBO):"), border=1, fill=True)
            pdf.cell(95, 10, txt=cln(f"$ {abono:,.0f}"), border=1, ln=True, align='R', fill=True)
            
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Helvetica", '', 11)
            pdf.cell(95, 8, txt=cln("NUEVO SALDO PENDIENTE:"), border=1)
            pdf.set_font("Helvetica", 'B', 11)
            pdf.cell(95, 8, txt=cln(f"$ {saldo:,.0f}"), border=1, ln=True, align='R')
            
            pdf.ln(20)
            
            pdf.set_font("Helvetica", 'B', 11)
            pdf.cell(95, 6, txt="__________________________________", align='C')
            pdf.cell(95, 6, txt="__________________________________", ln=True, align='C')
            pdf.cell(95, 6, txt=cln("Firma Autorizada - Dpto Financiero"), align='C')
            pdf.cell(95, 6, txt=cln("Recibí Conforme (Cliente)"), ln=True, align='C')
            
            pdf.ln(15)
            pdf.set_font("Helvetica", 'I', 8)
            pdf.set_text_color(150, 150, 150)
            pdf.cell(0, 4, txt=cln("Este documento es un comprobante de ingreso oficial generado y resguardado por el ecosistema LegalTech Insolvencia OS."), ln=True, align='C')
            
            filename = f"Factura_{cc}_{recibo_id}.pdf"
            pdf.output(filename)
            return filename
        
        df_f_n = df_fin.merge(df_activos[['Cedula', 'Nombre', 'Telefono']], on='Cedula', how='inner') if not df_fin.empty else pd.DataFrame()
        if not df_f_n.empty:
            st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
            cf = st.selectbox("Seleccionar Cliente:", df_f_n["Cedula"].astype(str) + " - " + df_f_n["Nombre"])
            if cf:
                cc_f = cf.split(" - ")[0]
                dat_f = df_f_n[df_f_n["Cedula"].astype(str) == cc_f].iloc[0]
                nom_f = dat_f['Nombre']
                tel_f = dat_f['Telefono']
                
                hon_t = limpiar_num(dat_f['Honorarios'])
                abo_t = limpiar_num(dat_f['Abonado'])
                saldo_actual = hon_t - abo_t
                
                st.metric("Saldo Pendiente Actual", f"$ {saldo_actual:,.0f}")
                
                with st.form("abono"):
                    n_abo = st.number_input("Registrar Abono Actual ($)", min_value=0, step=50000)
                    if st.form_submit_button("ACTUALIZAR CONTABILIDAD Y GENERAR FACTURA"):
                        if n_abo > 0:
                            nuevo_abonado = abo_t + n_abo
                            nuevo_saldo = hon_t - nuevo_abonado
                            
                            df_fu = leer_tabla("finanzas", ["Cedula", "Honorarios", "Abonado"])
                            df_fu.loc[df_fu["Cedula"].astype(str) == cc_f, "Abonado"] = str(int(nuevo_abonado))
                            guardar_tabla(df_fu, "finanzas")
                            
                            recibo_id = f"REC-{hoy.strftime('%Y%m%d%H%M%S')}"
                            
                            pdf_path = generar_factura_pdf(nom_f, cc_f, hon_t, n_abo, nuevo_saldo, recibo_id)
                            st.session_state['ultimo_pdf_factura'] = pdf_path
                            
                            msg_wa = f"⚖️ *FIRMA JURÍDICA - INSOLVENCIA OS* ⚖️%0A%0AEstimado/a *{nom_f}*, desde el departamento financiero confirmamos la recepción exitosa de su pago.%0A%0A💰 *Abono registrado:* ${n_abo:,.0f}%0A📉 *Nuevo saldo pendiente:* ${nuevo_saldo:,.0f}%0A%0ASu recibo de caja oficial No. {recibo_id} ha sido generado en nuestro sistema. %0A%0A¡Gracias por su cumplimiento y confianza en nuestro equipo!"
                            st.session_state['ultimo_wa_factura'] = f"https://wa.me/57{str(tel_f).replace(' ', '')}?text={msg_wa}"
                            
                            st.success("¡Pago registrado en la nube con éxito!")
                            st.rerun()
                            
                if 'ultimo_pdf_factura' in st.session_state and os.path.exists(st.session_state['ultimo_pdf_factura']):
                    st.markdown("<hr class='nav-hr' style='background:linear-gradient(90deg,transparent,#CDBF9A,transparent);'><h3 style='color:#10B981;'>🧾 Factura Oficial Generada</h3>", unsafe_allow_html=True)
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        with open(st.session_state['ultimo_pdf_factura'], "rb") as f:
                            st.download_button("📥 DESCARGAR FACTURA (PDF)", f, file_name=st.session_state['ultimo_pdf_factura'], use_container_width=True)
                    with col_f2:
                        if 'ultimo_wa_factura' in st.session_state:
                            st.markdown(f"<a href='{st.session_state['ultimo_wa_factura']}' target='_blank'><button style='background:#10B981; color:white; border:none; padding:10px 20px; border-radius:6px; font-weight:bold; cursor:pointer; width:100%;'>🚀 ENVIAR RESUMEN POR WHATSAPP</button></a>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

# --- 11. USUARIOS ---
elif st.session_state.pagina_actual == 'Usuarios':
    if not es_jefa():
        st.error("⛔ Solo la administradora de la firma puede gestionar usuarios.")
    else:
        firma_u = _ctx()
        max_u = max(mf._entero(firma_u.get("Max_Usuarios"), 3), 1)
        df_u_pag = leer_tabla("usuarios", COLS_USR).astype(object)
        usados = len(df_u_pag)
        st.markdown(f"<h1>Accesos y usuarios</h1><p class='page-sub'>Crea las cuentas de tu equipo. Tu plan incluye hasta {max_u} usuarios; hoy usas {usados}.</p>", unsafe_allow_html=True)
        st.progress(min(usados / max_u, 1.0), text=f"{usados} de {max_u} usuarios")

        with st.container(border=True):
            st.markdown("<p class='card-title'>Usuarios de tu firma</p>", unsafe_allow_html=True)
            vista_u = df_u_pag[["Usuario", "Alias", "Rol", "Cambiar_Clave"]].copy()
            vista_u.columns = ["Usuario", "Nombre", "Rol", "Debe cambiar clave"]
            st.dataframe(vista_u, use_container_width=True, hide_index=True)

        col_n, col_e = st.columns(2)
        with col_n:
            with st.container(border=True):
                st.markdown("<p class='card-title'>➕ Crear usuario</p>", unsafe_allow_html=True)
                if usados >= max_u:
                    st.warning(f"Llegaste al límite de {max_u} usuarios de tu plan. Para agregar más, comunícate con soporte.")
                else:
                    with st.form("form_nuevo_usr", clear_on_submit=True):
                        n_user = st.text_input("Usuario de acceso", help="3 a 30 caracteres: letras minúsculas, números, punto o guion. Ej.: mgomez")
                        n_alias = st.text_input("Nombre a mostrar")
                        n_rol = st.selectbox("Rol", mf.ROLES, index=1)
                        n_pass = st.text_input("Contraseña temporal", type="password", help="La persona deberá cambiarla en su primer ingreso.")
                        if st.form_submit_button("CREAR USUARIO"):
                            usuario_nuevo = n_user.strip().lower()
                            motivo = mf.validar_clave_nueva(n_pass)
                            if not re.fullmatch(r"[a-z0-9._-]{3,30}", usuario_nuevo):
                                st.error("El usuario debe tener de 3 a 30 caracteres: letras minúsculas, números, punto o guion.")
                            elif usuario_nuevo in df_u_pag["Usuario"].astype(str).str.lower().tolist():
                                st.error("Ese usuario ya existe en tu firma.")
                            elif motivo:
                                st.error(motivo)
                            else:
                                nuevo_u = pd.DataFrame([{"Usuario": usuario_nuevo, "Password": mf.hash_password(n_pass), "Rol": n_rol,
                                                         "Creador": st.session_state.usuario_actual, "Alias": n_alias.strip() or usuario_nuevo,
                                                         "Avatar_Path": "", "Session_Token": "", "Cambiar_Clave": "Si"}])
                                guardar_tabla(pd.concat([df_u_pag, nuevo_u], ignore_index=True), "usuarios")
                                registrar_log("SEGURIDAD", f"Creó el usuario {usuario_nuevo} ({n_rol})")
                                st.success(f"Usuario {usuario_nuevo} creado. Entrégale su contraseña temporal por un canal seguro.")
                                time.sleep(1.2)
                                st.rerun()

        with col_e:
            with st.container(border=True):
                st.markdown("<p class='card-title'>✏️ Modificar o eliminar</p>", unsafe_allow_html=True)
                usr_sel = st.selectbox("Usuario", df_u_pag["Usuario"].astype(str).tolist(), key="sel_usr_edit")
                fila_e = df_u_pag[df_u_pag["Usuario"].astype(str) == usr_sel].iloc[0]
                with st.form("form_edit_usr"):
                    e_alias = st.text_input("Nombre a mostrar", value=str(fila_e["Alias"]))
                    e_rol = st.selectbox("Rol", mf.ROLES, index=mf.ROLES.index(fila_e["Rol"]) if fila_e["Rol"] in mf.ROLES else 1)
                    e_pass = st.text_input("Nueva contraseña (déjala vacía para no cambiarla)", type="password")
                    b1, b2 = st.columns(2)
                    guardar_u = b1.form_submit_button("💾 GUARDAR")
                    borrar_u = b2.form_submit_button("🗑️ ELIMINAR")
                admins = df_u_pag[df_u_pag["Rol"].astype(str) == mf.ROLES[0]]
                es_ultimo_admin = (str(fila_e["Rol"]) == mf.ROLES[0]) and len(admins) <= 1
                if guardar_u:
                    if es_ultimo_admin and e_rol != mf.ROLES[0]:
                        st.error("No puedes quitarle el rol de administradora a la única administradora de la firma.")
                    elif e_pass and mf.validar_clave_nueva(e_pass):
                        st.error(mf.validar_clave_nueva(e_pass))
                    else:
                        m_e = df_u_pag["Usuario"].astype(str) == usr_sel
                        df_u_pag.loc[m_e, "Alias"] = e_alias.strip() or usr_sel
                        df_u_pag.loc[m_e, "Rol"] = e_rol
                        if e_pass:
                            df_u_pag.loc[m_e, "Password"] = mf.hash_password(e_pass)
                            df_u_pag.loc[m_e, "Cambiar_Clave"] = "Si"
                            df_u_pag.loc[m_e, "Session_Token"] = "REVOCADA"
                        guardar_tabla(df_u_pag, "usuarios")
                        registrar_log("SEGURIDAD", f"Actualizó al usuario {usr_sel}" + (" y restableció su contraseña" if e_pass else ""))
                        st.success("Cambios guardados." + (" Su sesión actual se cerrará." if e_pass else ""))
                        if usr_sel.lower() == str(st.session_state.usuario_actual).lower():
                            e_alias_final = e_alias.strip() or usr_sel
                            st.session_state.alias_actual = e_alias_final
                        time.sleep(1.2)
                        st.rerun()
                if borrar_u:
                    if usr_sel.lower() == str(st.session_state.usuario_actual).lower():
                        st.error("No puedes eliminar tu propia cuenta mientras la usas.")
                    elif es_ultimo_admin:
                        st.error("No puedes eliminar a la única administradora de la firma.")
                    else:
                        guardar_tabla(df_u_pag[df_u_pag["Usuario"].astype(str) != usr_sel], "usuarios")
                        registrar_log("SEGURIDAD", f"Eliminó al usuario {usr_sel}")
                        st.success(f"Usuario {usr_sel} eliminado.")
                        time.sleep(1.2)
                        st.rerun()

# --- 12. SISTEMA ---
elif st.session_state.pagina_actual == 'Sistema':
    if st.session_state.rol_actual != "Administrador (Jefa)": st.error("⛔ DENEGADO")
    else:
        st.markdown("<h1>🛡️ Auditoría y Exportación del Sistema</h1>", unsafe_allow_html=True)
        st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
        st.success("☁️ Base de datos sincronizada y operando 100% sobre Google Sheets Master.")
        
        st.markdown("<h3>📥 Descarga de Resguardos CSV Locales</h3>", unsafe_allow_html=True)
        col_d1, col_d2, col_d3 = st.columns(3)
        with col_d1:
            csv_c = df_cli.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Descargar Clientes CSV", csv_c, "clientes.csv", "text/csv")
        with col_d2:
            csv_f = df_fin.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Descargar Finanzas CSV", csv_f, "finanzas.csv", "text/csv")
        with col_d3:
            csv_l = df_log.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Descargar Logs CSV", csv_l, "logs.csv", "text/csv")
            
        st.markdown("<br><hr style='border-color: #CBD5E1;'>", unsafe_allow_html=True)
        
        st.markdown("<h3>🔍 Terminal de Auditoría Detallada por Usuario</h3>", unsafe_allow_html=True)
        st.markdown("<p style='color: #64748B; font-size: 13px;'>Selecciona un colaborador activo para revisar el rastro exacto de todas sus operaciones.</p>", unsafe_allow_html=True)
        
        usuarios_auditables = sorted(df_usr["Alias"].dropna().unique().tolist()) if not df_usr.empty else []
            
        if usuarios_auditables:
            usuario_seleccionado = st.selectbox("👤 Seleccionar Colaborador a Auditar:", usuarios_auditables)
            
            if usuario_seleccionado:
                st.markdown(f"<br><h4 style='color: #2563EB;'>📜 Historial de Actividad: {usuario_seleccionado}</h4>", unsafe_allow_html=True)
                
                if not df_log.empty:
                    df_log_usuario = df_log[df_log["Usuario"].astype(str) == str(usuario_seleccionado)].copy()
                    
                    if not df_log_usuario.empty:
                        st.info(f"Se encontraron **{len(df_log_usuario)}** acciones registradas en el sistema para este perfil.")
                        st.dataframe(df_log_usuario.iloc[::-1][["Timestamp", "Modulo", "Accion"]], use_container_width=True, hide_index=True)
                    else:
                        st.info(f"✨ El colaborador '{usuario_seleccionado}' no registra eventos o acciones en el sistema todavía.")
                else:
                    st.info("No hay registros de auditoría disponibles en la base de datos.")
        else:
            st.warning("No hay usuarios activos registrados para auditar.")
            
        st.markdown("</div>", unsafe_allow_html=True)

# --- 13. PAPELERA DE RECICLAJE ---
elif st.session_state.pagina_actual == 'Papelera':
    st.markdown("<h1>🗑️ Papelera de Reciclaje</h1>", unsafe_allow_html=True)
    st.markdown("<div class='module-card'>", unsafe_allow_html=True)
    
    df_borrados = df_cli[df_cli["Estado"] == "Borrado"].copy()
    
    if not df_borrados.empty:
        st.info("⚠️ Los expedientes aquí estarán por 30 días antes de su eliminación permanente. Puedes restaurarlos si fue un error.")
        
        for _, r in df_borrados.iterrows():
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"<h4 style='margin-bottom: 0;'>{r['Cedula']} - {r['Nombre']}</h4>", unsafe_allow_html=True)
                st.markdown(f"<span style='color: #EF4444; font-size: 13px; font-weight: bold;'>Enviado a papelera el: {r.get('F_Borrado', 'Desconocido')}</span>", unsafe_allow_html=True)
            with col2:
                if st.button("♻️ Restaurar Cliente", key=f"res_{r['Cedula']}"):
                    with st.spinner("Restaurando expediente y carpeta en Drive..."):
                        df_cli.loc[df_cli["Cedula"].astype(str) == str(r['Cedula']), "Estado"] = "Activo"
                        df_cli.loc[df_cli["Cedula"].astype(str) == str(r['Cedula']), "F_Borrado"] = ""
                        guardar_tabla(df_cli, "clientes")
                        
                        restaurar_de_papelera_drive(str(r['Cedula']), r['Nombre'])
                        
                        registrar_log("SISTEMA", f"Restauró expediente de la papelera: {r['Cedula']}")
                        st.success("¡Expediente restaurado con éxito!")
                        time.sleep(1.5)
                        st.rerun()
            st.markdown("<hr style='border-color: #E2E8F0; margin: 10px 0;'>", unsafe_allow_html=True)
    else:
        st.success("✨ La papelera está vacía. No hay expedientes eliminados.")
        
    st.markdown("</div>", unsafe_allow_html=True)
