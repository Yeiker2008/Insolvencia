# ISOLVENCIAS APP

import streamlit as st
import pandas as pd
import os
import shutil
import base64
import random
import uuid
from datetime import datetime, timedelta
import time
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

# --- PODERES PARA GOOGLE DRIVE (LA BÓVEDA MAESTRA) ---
try:
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
    import io
    GDRIVE_DISPONIBLE = True
except ImportError:
    st.error("⚠️ Abre la terminal y ejecuta: pip install google-api-python-client")
    GDRIVE_DISPONIBLE = False

# Definimos los permisos necesarios (Sheets y Drive)
SCOPES = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]

# ID de tu Bóveda Maestra en Drive (¡REEMPLAZA ESTO!)
CARPETA_RAIZ_DRIVE_ID = "1MC6wHXaV557prpKV-KCRc8yeCdphD6U8"

# ID de la Papelera en Drive (¡NUEVO!)
CARPETA_PAPELERA_DRIVE_ID = "1MGBXOKbPuAE6xsf53AcCjwMFE_yzDbir"

# Conexión Global a Google Drive
@st.cache_resource
def conectar_gdrive():
    if not GDRIVE_DISPONIBLE: return None
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
            drive_service = build('drive', 'v3', credentials=creds)
            return drive_service
    except Exception as e:
        st.error(f"🚨 ERROR CONECTANDO A GOOGLE DRIVE: {e}")
    return None

gc_drive = conectar_gdrive()

def mover_a_papelera_drive(cc, nom):
    if not gc_drive: return False
    
    nombre_carpeta_cliente = f"{cc} - {nom}"
    
    try:
        # 1. Buscar la carpeta del cliente dentro de la Bóveda Maestra
        query = f"name='{nombre_carpeta_cliente}' and '{CARPETA_RAIZ_DRIVE_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        resultados = gc_drive.files().list(q=query, fields="files(id, parents)").execute()
        archivos = resultados.get('files', [])
        
        if archivos:
            carpeta_id = archivos[0]['id']
            # Obtener los padres actuales (la Bóveda) para quitarlos
            padres_viejos = ",".join(archivos[0].get('parents', []))
            
            # 2. Mover la carpeta (Se le agrega el padre Papelera y se le quita el padre Bóveda)
            gc_drive.files().update(
                fileId=carpeta_id,
                addParents=CARPETA_PAPELERA_DRIVE_ID,
                removeParents=padres_viejos,
                fields='id, parents'
            ).execute()
            return True
            
        return False
        
    except Exception as e:
        print(f"Error moviendo la carpeta a la Papelera de Drive: {e}")
        return False

def restaurar_de_papelera_drive(cc, nom):
    if not gc_drive: return False
    nombre_carpeta_cliente = f"{cc} - {nom}"
    try:
        # Buscar la carpeta en la Papelera
        query = f"name='{nombre_carpeta_cliente}' and '{CARPETA_PAPELERA_DRIVE_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        resultados = gc_drive.files().list(q=query, fields="files(id, parents)").execute()
        archivos = resultados.get('files', [])
        
        if archivos:
            carpeta_id = archivos[0]['id']
            padres_viejos = ",".join(archivos[0].get('parents', []))
            
            # Mover de vuelta a la Bóveda Maestra
            gc_drive.files().update(
                fileId=carpeta_id,
                addParents=CARPETA_RAIZ_DRIVE_ID,
                removeParents=padres_viejos,
                fields='id, parents'
            ).execute()
            return True
        return False
    except Exception as e:
        print(f"Error restaurando carpeta en Drive: {e}")
        return False

def obtener_id_subcarpeta_drive(cc, nom, nombre_subcarpeta):
    """Busca y retorna el ID de la subcarpeta interna (ej: '01_Docs_Viabilidad') en Google Drive"""
    if not gc_drive: return None
    nombre_carpeta_cliente = f"{cc} - {nom}"
    try:
        # 1. Buscar carpeta principal del cliente en la Bóveda
        query = f"name='{nombre_carpeta_cliente}' and '{CARPETA_RAIZ_DRIVE_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res = gc_drive.files().list(q=query, fields="files(id)").execute()
        carpetas = res.get('files', [])
        if not carpetas: return None
        cliente_id = carpetas[0]['id']
        
        # 2. Buscar la subcarpeta interna
        query_sub = f"name='{nombre_subcarpeta}' and '{cliente_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res_sub = gc_drive.files().list(q=query_sub, fields="files(id)").execute()
        subcarpetas = res_sub.get('files', [])
        if subcarpetas: return subcarpetas[0]['id']
        return None
    except Exception as e:
        print(f"Error obteniendo subcarpeta en Drive: {e}")
        return None

def subir_archivo_a_drive(file_obj, filename, folder_id):
    """Sube un archivo cargado desde Streamlit directamente a una carpeta de Google Drive"""
    if not gc_drive or not folder_id: return False
    try:
        # Guardar archivo temporal en disco local para transferir a la API
        temp_path = f"temp_{filename}"
        with open(temp_path, "wb") as f:
            f.write(file_obj.getbuffer())
            
        file_metadata = {'name': filename, 'parents': [folder_id]}
        media = MediaFileUpload(temp_path, resumable=True)
        gc_drive.files().create(body=file_metadata, media_body=media, fields='id').execute()
        
        if os.path.exists(temp_path): os.remove(temp_path)
        return True
    except Exception as e:
        print(f"Error subiendo archivo a Drive: {e}")
        return False

def listar_y_descargar_archivos_drive(folder_id):
    """Lista todos los archivos de una subcarpeta en Drive y retorna una lista con su nombre y contenido binario"""
    if not gc_drive or not folder_id: return []
    try:
        query = f"'{folder_id}' in parents and mimeType != 'application/vnd.google-apps.folder' and trashed=false"
        res = gc_drive.files().list(q=query, fields="files(id, name)").execute()
        archivos = res.get('files', [])
        lista_descargas = []
        for a in archivos:
            request = gc_drive.files().get_media(fileId=a['id'])
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
            fh.seek(0)
            lista_descargas.append({'id': a['id'], 'name': a['name'], 'content': fh.getvalue()})
        return lista_descargas
    except Exception as e:
        print(f"Error listando archivos de Drive: {e}")
        return []

# ==========================================
# 1. CORE & CONFIGURACIÓN DE SESIÓN Y GSHEETS
# ==========================================
st.set_page_config(page_title="Insolvencia OS | Enterprise", page_icon="⚖", layout="wide", initial_sidebar_state="expanded")

if 'autenticado' not in st.session_state: st.session_state.autenticado = False
if 'usuario_actual' not in st.session_state: st.session_state.usuario_actual = ""
if 'rol_actual' not in st.session_state: st.session_state.rol_actual = ""
if 'alias_actual' not in st.session_state: st.session_state.alias_actual = ""
if 'avatar_path' not in st.session_state: st.session_state.avatar_path = ""
if 'session_token' not in st.session_state: st.session_state.session_token = ""
if 'pagina_actual' not in st.session_state: st.session_state.pagina_actual = 'Dashboard'
if 'kicked' not in st.session_state: st.session_state.kicked = False
if 'kicked_reason' not in st.session_state: st.session_state.kicked_reason = ""

def cambiar_pagina(p): st.session_state.pagina_actual = p

# Conexión y Auto-creación Global en Google Sheets con aviso visual
@st.cache_resource
def conectar_gsheets():
    if not GSPREAD_DISPONIBLE: return None
    try:
        if "gcp_service_account" in st.secrets:
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            client = gspread.authorize(creds)
            
            nombre_hoja = "DB_Insolvencia_Master"
            
            try:
                sheet = client.open(nombre_hoja)
            except gspread.SpreadsheetNotFound:
                sheet = client.create(nombre_hoja)
                pestañas = ["clientes", "finanzas", "actuaciones", "vencimientos", "acreedores", "audiencias", "tareas", "usuarios", "logs"]
                first_sheet = sheet.get_sheet_by_id(0)
                first_sheet.update_title(pestañas[0])
                for p in pestañas[1:]:
                    sheet.add_worksheet(title=p, rows="100", cols="20")
                try:
                    sheet.share('chincuenta5025@gmail.com', perm_type='user', role='writer')
                except:
                    pass
            
            return sheet
    except Exception as e:
        st.error(f"🚨 ERROR CRÍTICO CREANDO/CONECTANDO GSHEETS: {e}")
    return None

gc_sheet = conectar_gsheets()

# Inyectamos Caché Ultrarrápida (Guarda en RAM por 2 minutos o hasta que haya cambios)
@st.cache_data(ttl=120, show_spinner=False)
def leer_tabla(nombre_tabla, columnas_def):
    if gc_sheet:
        try:
            worksheet = gc_sheet.worksheet(nombre_tabla)
            data = worksheet.get_all_records()
            df_gs = pd.DataFrame(data)
            if df_gs.empty:
                return pd.DataFrame(columns=columnas_def)
            return df_gs
        except Exception:
            pass
    
    arch = f"db_{nombre_tabla}.csv"
    if not os.path.exists(arch):
        pd.DataFrame(columns=columnas_def).to_csv(arch, index=False)
    return pd.read_csv(arch)

def guardar_tabla(df, nombre_tabla):
    arch = f"db_{nombre_tabla}.csv"
    df.to_csv(arch, index=False)
    if gc_sheet:
        try:
            worksheet = gc_sheet.worksheet(nombre_tabla)
            worksheet.clear()
            worksheet.update([df.columns.values.tolist()] + df.fillna("").values.tolist())
        except Exception as e:
            print(f"Error guardando en GSheets {nombre_tabla}: {e}")
            
    # MAGIA: Vaciamos la RAM para que la app descargue los datos frescos de inmediato
    leer_tabla.clear()

# ==========================================
# 2. MOTOR CSS: SAAS CORPORATIVO LIMPIO (ANTI-FATIGA VISUAL + ANIMACIONES 3D)
# ==========================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    #MainMenu, footer, header {visibility: hidden !important; display: none !important;}
    [data-testid="stHeader"] {display: none !important; visibility: hidden !important;}
    
    /* Fuentes y fondos generales (Área de trabajo clara) */
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #F8FAFC !important; color: #1E293B !important; }
    h1, h2, h3 { color: #0F172A !important; font-weight: 700; letter-spacing: -0.5px; }
    [data-testid="stAppViewContainer"] { background-color: #F4F7F8 !important; background-image: none !important; }
    
    /* Sidebar: Oscuro y elegante */
    [data-testid="stSidebar"] { background-color: #0F172A !important; border-right: 1px solid #1E293B !important; }
    /* Textos del sidebar (Títulos y párrafos) en blanco, sin dañar los inputs ni botones */
    [data-testid="stSidebar"] p, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] span, [data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] { color: #F8FAFC !important; }
    
    /* 🔥 BOTONES DEL SIDEBAR: Visibles, elegantes y con animación de agrandado */
    [data-testid="stSidebar"] div.stButton > button:first-child {
        background-color: #1E293B !important; /* Fondo oscuro sutil */
        color: #F8FAFC !important; /* Letras blancas siempre visibles */
        border: 1px solid #334155 !important;
        transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1) !important; /* Transición súper suave */
    }
    [data-testid="stSidebar"] div.stButton > button:first-child:hover {
        background-color: #2563EB !important; /* Azul eléctrico al pasar el mouse */
        border-color: #2563EB !important;
        color: #FFFFFF !important; 
        transform: scale(1.05) !important; /* EFECTO 3D: EL BOTÓN SE AGRANDA */
        box-shadow: 0 5px 15px rgba(37,99,235,0.4) !important; /* Brillo azul */
        z-index: 10;
    }

    /* 🔥 BOTONES DEL ÁREA PRINCIPAL: Blancos con azul y animación de agrandado */
    div.stButton > button:first-child { 
        background: #FFFFFF !important; 
        color: #2563EB !important; 
        border: 2px solid #2563EB !important; 
        border-radius: 6px !important; 
        padding: 10px 20px !important; 
        font-weight: 600 !important; 
        font-size: 13px !important; 
        transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1) !important; 
        width: 100%;
    }
    div.stButton > button:first-child:hover { 
        background: #2563EB !important; 
        color: #FFFFFF !important; 
        box-shadow: 0 8px 20px rgba(37,99,235,0.3) !important; 
        transform: scale(1.03) translateY(-2px) !important; /* EFECTO 3D: SE AGRANDA Y SE LEVANTA */
    }
    
    /* Tarjetas de Módulos */
    .module-card { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; padding: 30px; margin-bottom: 25px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -1px rgba(0,0,0,0.03); }
    .module-card p { color: #475569 !important; } 
    
    .module-card-gold { border-top: 3px solid #2563EB !important; } 
    .module-card-blue { border-top: 3px solid #0EA5E9 !important; }
    .module-card-green { border-top: 3px solid #10B981 !important; }
    
    /* Inputs y Formularios: Blancos con borde gris oscuro para que se vean en cualquier lado */
    input, textarea, select, div[data-baseweb="select"] > div, div[data-baseweb="input"] > div { background-color: #FFFFFF !important; color: #0F172A !important; -webkit-text-fill-color: #0F172A !important; border: 1px solid #94A3B8 !important; border-radius: 6px !important; }
    input:focus, textarea:focus { border-color: #2563EB !important; box-shadow: 0 0 0 2px rgba(37,99,235,0.2) !important; }
    
    /* Línea de Tiempo de Fases */
    .timeline { display: flex; justify-content: space-between; align-items: center; margin: 30px 0; position: relative; }
    .timeline::before { content: ''; position: absolute; top: 50%; left: 0; right: 0; height: 2px; background: #E2E8F0; z-index: 1; }
    .step { position: relative; z-index: 2; background: #F8FAFC; padding: 8px 16px; border-radius: 20px; border: 2px solid #E2E8F0; color: #64748B; font-weight: 600; font-size: 12px; display: flex; align-items: center; text-transform: uppercase; }
    .step.active { border-color: #2563EB; color: #2563EB; box-shadow: 0 0 0 4px rgba(37,99,235,0.1); background: #FFFFFF; }
    .step.completed { border-color: #10B981; color: #FFFFFF; background: #10B981; }
    
    .alerta-roja { animation: blinker 1.5s linear infinite; color: #EF4444 !important; font-weight: bold;}
    @keyframes blinker { 50% { opacity: 0; } }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. BASE DE DATOS Y CARGA CLOUD
# ==========================================
CARP_EXP, CARP_PLA, CARP_PAP, CARP_AVATAR = "Expedientes", "Plantillas", "Papelera", "Avatares"
ABOGADOS = ["David Alexander Arias Posse", "Maria Jose Ospino Perez", "Yeiker Cardona"]

for c in [CARP_EXP, CARP_PLA, CARP_PAP, CARP_AVATAR]: os.makedirs(c, exist_ok=True)
for a in ABOGADOS: os.makedirs(os.path.join(CARP_PLA, a), exist_ok=True)

df_cli = leer_tabla("clientes", ["Cedula", "Nombre", "Fuerza", "Telefono", "Email", "Deuda_Est", "Ingresos", "Senal", "F_Actualizacion", "Estado", "F_Borrado"])
df_fin = leer_tabla("finanzas", ["Cedula", "Honorarios", "Abonado"])
df_act = leer_tabla("actuaciones", ["ID_Act", "Cedula", "Fecha", "Tipo", "Juzgado", "Radicado", "Anotacion"])
df_log = leer_tabla("logs", ["Timestamp", "Usuario", "Modulo", "Accion"])
df_ven = leer_tabla("vencimientos", ["ID_Ven", "Cedula", "Cliente", "Asunto", "Fecha_Limite", "Estado"])
df_acr = leer_tabla("acreedores", ["ID_Acr", "Cedula", "Acreedor", "Cuantia", "Clase"])
df_aud = leer_tabla("audiencias", ["ID_Aud", "Cedula", "Cliente", "Fecha_Hora", "Motivo"])
df_tar = leer_tabla("tareas", ["ID_Tar", "Tarea", "Asignado", "Creador", "Estado", "Fecha"])
df_usr = leer_tabla("usuarios", ["Usuario", "Password", "Rol", "Creador", "Alias", "Avatar_Path", "Session_Token"])

if df_usr.empty:
    df_usr = pd.DataFrame([{"Usuario": "admin", "Password": "123", "Rol": "Administrador (Jefa)", "Creador": "Sistema", "Alias": "Administración", "Avatar_Path": "", "Session_Token": ""}])
    guardar_tabla(df_usr, "usuarios")

if "Alias" not in df_usr.columns: df_usr["Alias"] = df_usr["Usuario"]
if "Avatar_Path" not in df_usr.columns: df_usr["Avatar_Path"] = ""
if "Session_Token" not in df_usr.columns: df_usr["Session_Token"] = ""

hoy = datetime.now()

cambios_bd = False
for idx, r in df_cli[df_cli["Estado"] == "Borrado"].iterrows():
    if pd.notna(r["F_Borrado"]) and str(r["F_Borrado"]).strip() != "":
        try:
            if (hoy - datetime.strptime(str(r["F_Borrado"])[:10], "%Y-%m-%d")).days >= 30:
                rp = os.path.join(CARP_PAP, f"{r['Cedula']} - {r['Nombre']}")
                if os.path.exists(rp): shutil.rmtree(rp)
                df_cli = df_cli.drop(idx)
                cambios_bd = True
        except: pass
if cambios_bd: guardar_tabla(df_cli, "clientes")

df_activos = df_cli[df_cli["Estado"] == "Activo"].copy()
df_papelera = df_cli[df_cli["Estado"] == "Borrado"].copy()

def registrar_log(modulo, accion):
    try:
        global df_log
        usr = st.session_state.get('alias_actual', 'Sistema')
        nuevo = pd.DataFrame([{"Timestamp": hoy.strftime("%Y-%m-%d %H:%M:%S"), "Usuario": usr, "Modulo": modulo, "Accion": accion}])
        df_log = pd.concat([df_log, nuevo], ignore_index=True)
        guardar_tabla(df_log, "logs")
    except: pass

def estructurar_carpetas(cc, nom):
    # Si por alguna razón no hay conexión a Drive, cancela para no causar errores
    if not gc_drive: return False
    
    nombre_carpeta_cliente = f"{cc} - {nom}"
    
    try:
        # 1. Verificar si el cliente ya tiene carpeta en Drive para no duplicarla
        query = f"name='{nombre_carpeta_cliente}' and '{CARPETA_RAIZ_DRIVE_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        resultados = gc_drive.files().list(q=query, fields="files(id, name)").execute()
        archivos = resultados.get('files', [])
        
        if not archivos:
            # 2. Si no existe, crear la carpeta principal del cliente en la Bóveda Maestra
            metadata_cliente = {
                'name': nombre_carpeta_cliente,
                'parents': [CARPETA_RAIZ_DRIVE_ID],
                'mimeType': 'application/vnd.google-apps.folder'
            }
            carpeta_cliente = gc_drive.files().create(body=metadata_cliente, fields='id').execute()
            cliente_id = carpeta_cliente.get('id')
            
            # 3. Crear las subcarpetas internas obligatorias
            subcarpetas = ["Financiero", "Juzgado", "Notaria", "Acreedores"]
            for sub in subcarpetas:
                metadata_sub = {
                    'name': sub,
                    'parents': [cliente_id],
                    'mimeType': 'application/vnd.google-apps.folder'
                }
                gc_drive.files().create(body=metadata_sub, fields='id').execute()
            return True
            
        return False
        
    except Exception as e:
        print(f"Error creando carpetas en Drive: {e}")
        return False

def limpiar_num(val):
    try: return float(str(val).replace("$","").replace(".","").replace(",","").strip())
    except: return 0.0

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
        estructurar_carpetas(cc, nom)
        subcarpetas = ["Financiero", "Juzgado", "Notaria", "Acreedores"]
        cols = st.columns(2)
        
        for i, sub in enumerate(subcarpetas):
            col = cols[i % 2]
            col.markdown(f"<p style='color:#2563EB; margin-bottom:2px; font-weight:bold;'>{sub}</p>", unsafe_allow_html=True)
            folder_id = obtener_id_subcarpeta_drive(cc, nom, sub)
            
            if folder_id:
                archivos = listar_y_descargar_archivos_drive(folder_id)
                if archivos:
                    for a in archivos:
                        c_down, c_del = col.columns([4, 1])
                        c_down.download_button(
                            f"📄 {a['name'][:20]}...", 
                            a['content'], 
                            file_name=a['name'], 
                            key=f"dl_drive_{cc}_{sub}_{a['id']}", 
                            use_container_width=True
                        )
                        if c_del.button("🗑️", key=f"del_drive_{cc}_{sub}_{a['id']}", help="Borrar archivo de Google Drive"):
                            gc_drive.files().delete(fileId=a['id']).execute()
                            st.rerun()
                else:
                    col.markdown("<span style='color:#64748B; font-size:12px;'><i>Carpeta vacía</i></span>", unsafe_allow_html=True)
            col.write("")

# ==========================================
# 4. PANTALLA DE LOGIN Y SEGURIDAD ANTICLONACIÓN
# ==========================================
if st.session_state.autenticado:
    try:
        df_u_check = leer_tabla("usuarios", ["Usuario", "Password", "Rol", "Creador", "Alias", "Avatar_Path", "Session_Token"])
        user_match = df_u_check[df_u_check["Usuario"].astype(str) == str(st.session_state.usuario_actual)]
        if user_match.empty:
            st.session_state.autenticado = False; st.session_state.kicked = True; st.session_state.kicked_reason = "Usuario revocado."
            st.rerun()
        else:
            db_token = str(user_match.iloc[0].get("Session_Token", ""))
            local_token = str(st.session_state.get("session_token", ""))
            if db_token != local_token and db_token != "nan" and db_token != "":
                st.session_state.autenticado = False; st.session_state.kicked = True; st.session_state.kicked_reason = "Sesión iniciada en otro equipo."
                st.rerun()
    except: pass

if not st.session_state.autenticado:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        if st.session_state.kicked:
            st.error(f"🚨 ALERTA: {st.session_state.kicked_reason}"); st.session_state.kicked = False
        st.markdown("""<div class="module-card module-card-gold" style="text-align: center; padding: 40px;"><span style="font-size: 50px;">⚖️</span><h2>ACCESO RESTRINGIDO</h2><p style="color: #A1A1AA; font-size: 13px; margin-bottom: 30px;">Plataforma LegalTech Cifrada (Google Cloud DB)</p>""", unsafe_allow_html=True)
        with st.form("login_form"):
            user_input = st.text_input("👤 Usuario")
            pass_input = st.text_input("🔑 Contraseña", type="password")
            if st.form_submit_button("AUTENTICAR CREDENCIALES"):
                df_u = leer_tabla("usuarios", ["Usuario", "Password", "Rol", "Creador", "Alias", "Avatar_Path", "Session_Token"])
                user_match = df_u[(df_u["Usuario"].astype(str) == user_input) & (df_u["Password"].astype(str) == pass_input)]
                if not user_match.empty:
                    st.session_state.autenticado = True
                    st.session_state.usuario_actual = user_input
                    st.session_state.rol_actual = user_match.iloc[0]["Rol"]
                    st.session_state.alias_actual = user_match.iloc[0].get("Alias", user_input)
                    st.session_state.avatar_path = str(user_match.iloc[0].get("Avatar_Path", ""))
                    nuevo_token = uuid.uuid4().hex
                    st.session_state.session_token = nuevo_token
                    df_u.loc[df_u["Usuario"].astype(str) == user_input, "Session_Token"] = nuevo_token
                    guardar_tabla(df_u, "usuarios")
                    registrar_log("AUTH", f"Inició sesión: {user_input}")
                    st.rerun()
                else: st.error("❌ Credenciales inválidas.")
        st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

# ==========================================
# 5. SIDEBAR ORDENADO
# ==========================================
with st.sidebar:
    avatar_p = st.session_state.get('avatar_path', '')
    if pd.notna(avatar_p) and str(avatar_p).strip() != "" and os.path.exists(avatar_p):
        with open(avatar_p, "rb") as f: encoded_img = base64.b64encode(f.read()).decode()
        st.markdown(f"""<div style="text-align:center; margin-bottom: 10px;"><img src="data:image/png;base64,{encoded_img}" style="width: 70px; height: 70px; border-radius: 50%; object-fit: cover; border: 2px solid #D4AF37; box-shadow: 0 0 15px rgba(212,175,55,0.3);"></div>""", unsafe_allow_html=True)
    else:
        st.markdown("""<div style="width: 60px; height: 60px; background: #D4AF37; border-radius: 50%; margin: 0 auto 10px auto; display:flex; align-items:center; justify-content:center; box-shadow: 0 0 15px rgba(212,175,55,0.3);"><span style="font-size: 30px;">⚖️</span></div>""", unsafe_allow_html=True)

    st.markdown(f"""<div style="text-align:center; padding-bottom: 10px; margin-bottom: 15px;"><h3 style="color: #F8FAFC !important; font-size: 18px; margin:0;">{st.session_state.alias_actual}</h3><span style="color: #3B82F6; font-size: 11px; font-weight: bold;">{st.session_state.rol_actual}</span></div>""", unsafe_allow_html=True)
    
    with st.expander("⚙️ Mi Perfil"):
        with st.form("form_perfil"):
            n_alias = st.text_input("Alias", value=st.session_state.alias_actual)
            foto_sub = st.file_uploader("Foto de Perfil", type=["jpg", "png", "jpeg"])
            if st.form_submit_button("Actualizar Perfil"):
                df_u_up = leer_tabla("usuarios", ["Usuario", "Password", "Rol", "Creador", "Alias", "Avatar_Path", "Session_Token"])
                p_guardado = st.session_state.avatar_path
                if foto_sub:
                    p_guardado = os.path.join(CARP_AVATAR, f"{st.session_state.usuario_actual}_{foto_sub.name}")
                    with open(p_guardado, "wb") as f: f.write(foto_sub.getbuffer())
                df_u_up.loc[df_u_up["Usuario"].astype(str) == st.session_state.usuario_actual, "Alias"] = n_alias
                df_u_up.loc[df_u_up["Usuario"].astype(str) == st.session_state.usuario_actual, "Avatar_Path"] = p_guardado
                guardar_tabla(df_u_up, "usuarios")
                st.session_state.alias_actual = n_alias; st.session_state.avatar_path = p_guardado
                st.success("Perfil Actualizado"); st.rerun()

    st.markdown("<hr style='border-color: #27272A; margin-top: 0;'>", unsafe_allow_html=True)
    
    st.markdown("<p style='color:#71717A; font-size:11px; font-weight:bold; letter-spacing:1px;'>🏠 INICIO & GESTIÓN</p>", unsafe_allow_html=True)
    if st.button("📊 Portal Ejecutivo", key="b_dash"): cambiar_pagina("Dashboard")
    if st.button("📝 Apertura de Casos", key="b_nuev"): cambiar_pagina("Nuevo")
    if st.button("📅 Agenda y Citas", key="b_agen"): cambiar_pagina("Agenda")
    if st.button("✅ Gestor de Tareas", key="b_tar"): cambiar_pagina("Tareas")
    if st.button("🗑️ Papelera de Reciclaje", key="b_pap"): cambiar_pagina("Papelera") # <-- ESTA ES LA LÍNEA NUEVA
    
    st.markdown("<br><p style='color:#71717A; font-size:11px; font-weight:bold; letter-spacing:1px;'>⚖️ OPERACIÓN JURÍDICA</p>", unsafe_allow_html=True)
    if st.button("⚙️ Contratos y Docs", key="b_cont"): cambiar_pagina("Contratos")
    if st.button("📂 Seguimiento Procesal", key="b_act"): cambiar_pagina("Actuaciones")
    if st.button("🚦 Control Vencimientos", key="b_ven"): cambiar_pagina("Vencimientos")
    if st.button("🤖 Dependiente Virtual", key="b_mem"): cambiar_pagina("Memoriales")
    
    st.markdown("<br><p style='color:#71717A; font-size:11px; font-weight:bold; letter-spacing:1px;'>💰 FINANZAS & ADMIN</p>", unsafe_allow_html=True)
    if st.button("📋 Pasivos y Acreedores", key="b_acr"): cambiar_pagina("Acreedores")
    
    if st.session_state.rol_actual == "Administrador (Jefa)":
        if st.button("💰 Cobros y Facturación", key="b_fin"): cambiar_pagina("Finanzas")
        if st.button("👥 Accesos y Seguridad", key="b_usr"): cambiar_pagina("Usuarios")
        if st.button("🛡️ Auditoría y Exportación", key="b_sis"): cambiar_pagina("Sistema")
        
    st.markdown("<hr style='border-color: #27272A;'>", unsafe_allow_html=True)
    if st.button("🚪 Cerrar Sesión"): st.session_state.autenticado = False; st.rerun()

# ==========================================
# 6. MÓDULOS DE LA APLICACIÓN (TODOS INTACTOS)
# ==========================================

# --- 1. PORTAL EJECUTIVO ---
if st.session_state.pagina_actual == 'Dashboard':
    st.markdown("<h1>Portal Ejecutivo Legal</h1><p style='margin-bottom: 20px;'>Centro de operaciones sincronizado a Google Sheets.</p>", unsafe_allow_html=True)
    total_cartera = sum([(limpiar_num(r["Honorarios"]) - limpiar_num(r["Abonado"])) for _, r in df_fin[df_fin["Cedula"].astype(str).isin(df_activos["Cedula"].astype(str))].iterrows() if limpiar_num(r["Honorarios"]) > limpiar_num(r["Abonado"])])
    
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"<div class='module-card module-card-blue'><h4>Expedientes Activos</h4><h2>{len(df_activos)}</h2></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='module-card module-card-gold'><h4>En Trámite</h4><h2>{len(df_activos[df_activos['Senal'].isin([1,2,3])])}</h2></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='module-card module-card-green'><h4>Radicados</h4><h2>{len(df_activos[df_activos['Senal'] >= 4])}</h2></div>", unsafe_allow_html=True)
    if st.session_state.rol_actual == "Administrador (Jefa)": c4.markdown(f"<div class='module-card module-card-gold'><h4>Cartera Pendiente</h4><h2 style='color:#FACC15;'>$ {total_cartera:,.0f}</h2></div>", unsafe_allow_html=True)

    st.markdown("<div class='module-card'><h3>🔍 Buscador Rápido de Expedientes</h3>", unsafe_allow_html=True)
    busq = st.text_input("Ingresa nombre o cédula para encontrar al cliente de inmediato:")
    if busq:
        res = df_activos[df_activos['Nombre'].str.contains(busq, case=False, na=False) | df_activos['Cedula'].astype(str).str.contains(busq, na=False)]
        if not res.empty: st.dataframe(res[["Cedula", "Nombre", "Fuerza", "Telefono", "Email"]], use_container_width=True, hide_index=True)
        else: st.info("No hay resultados coincidentes.")
    st.markdown("</div>", unsafe_allow_html=True)
    
    col_info1, col_info2 = st.columns([1.5, 1])
    with col_info1:
        st.markdown("""
            <div class='module-card module-card-gold'>
                <h3 style='color: #FACC15; font-size: 20px; margin-bottom: 15px;'>📜 Ley 1564 de 2012</h3>
                <p style='color: #D4D4D8; font-size: 14px; line-height: 1.6;'>
                <b>Ventajas principales de la Insolvencia:</b><br>
                • Suspensión inmediata de embargos y procesos.<br>
                • Protección del patrimonio familiar.<br>
                • Reactivación financiera y paz mental.
                </p>
            </div>
        """, unsafe_allow_html=True)
    with col_info2:
        st.markdown("<div class='module-card'><h3 style='font-size: 18px; color:#F8FAFC !important;'>🔔 Alertas de Proceso</h3>", unsafe_allow_html=True)
        hay_alertas = False
        if not df_activos.empty:
            for _, r in df_activos.iterrows():
                try:
                    dias = (hoy - datetime.strptime(str(r["F_Actualizacion"])[:10], "%Y-%m-%d")).days
                    if r["Senal"] == 2 and dias >= 2:
                        hay_alertas = True; st.warning(f"⏳ **Firma pendiente:** {r['Nombre']} ({dias} días).")
                    elif r["Senal"] == 3 and dias >= 1:
                        hay_alertas = True; st.error(f"🚨 **Radicación urgente:** {r['Nombre']} ({dias} días).")
                except: pass
        if not hay_alertas: st.success("✨ Expedientes fluyendo con normalidad.")
        st.markdown("</div>", unsafe_allow_html=True)
        
    # --- ZONA DE PELIGRO: MOVER A PAPELERA ---
    st.markdown("<hr style='border-color: #27272A;'><h3 style='color:#EF4444;'>⚠️ Zona de Peligro: Gestión de Archivo</h3>", unsafe_allow_html=True)
    st.markdown("<div class='module-card'>", unsafe_allow_html=True)
    if not df_activos.empty:
        with st.form("form_borrar"):
            cliente_a_borrar = st.selectbox("Seleccionar expediente para mover a la Papelera:", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"])
            if st.form_submit_button("MOVER A LA PAPELERA"):
                cc_b = cliente_a_borrar.split(" - ")[0]
                nom_b = cliente_a_borrar.split(" - ")[1] # Extraemos el nombre para Google Drive
                
                # 1. Cambiar estado a "Borrado" en Google Sheets
                df_cli.loc[df_cli["Cedula"].astype(str) == cc_b, "Estado"] = "Borrado"
                df_cli.loc[df_cli["Cedula"].astype(str) == cc_b, "F_Borrado"] = hoy.strftime("%Y-%m-%d")
                guardar_tabla(df_cli, "clientes")
                
                # 2. 🔥 LA MAGIA DE GOOGLE DRIVE 🔥: Mover carpeta a la Papelera de Drive
                mover_a_papelera_drive(str(cc_b), nom_b)
                
                registrar_log("SISTEMA", f"Expediente y carpeta enviados a papelera: {cc_b}")
                st.success("¡Expediente movido a la papelera correctamente en Sheets y Drive!")
                time.sleep(1.5) # Pausa corta para que se vea el mensaje
                st.rerun()
    else:
        st.info("No hay clientes activos para archivar.")
    st.markdown("</div>", unsafe_allow_html=True)
    
# --- 2. APERTURA Y EDICIÓN ---
elif st.session_state.pagina_actual == 'Nuevo':
    st.markdown("<h1>Gestión de Expedientes</h1>", unsafe_allow_html=True)
    
    # Creamos las dos pestañas de navegación
    tab_nuevo, tab_editar = st.tabs(["➕ Apertura de Nuevo Caso", "✏️ Modificar Expediente Existente"])
    
    # ---------------------------------------------
    # PESTAÑA 1: CREAR NUEVO
    # ---------------------------------------------
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
            with c6: deuda = st.text_input("Deuda Aprox (\$)")
            
            st.markdown("<hr style='border-color: #E2E8F0;'>", unsafe_allow_html=True)
            st.markdown("<p style='color:#2563EB; font-weight:bold;'>💰 Configuración Financiera del Caso</p>", unsafe_allow_html=True)
            cf1, cf2 = st.columns(2)
            with cf1: honorarios_totales = st.text_input("Valor Total de Honorarios del Caso (\$)", value="0")
            with cf2: abono_inicial = st.text_input("Abono Inicial Recibido (\$)", value="0")
            
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
        
   # ---------------------------------------------
    # PESTAÑA 2: MODIFICAR CLIENTE (BLINDADO CONTRA TYPEERROR)
    # ---------------------------------------------
    with tab_editar:
        st.markdown("<div class='module-card module-card-blue'>", unsafe_allow_html=True)
        
        # Leemos los datos en vivo para asegurar refresco
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
                            # 1. Renombrar carpeta física
                            old_path = os.path.join(CARP_EXP, f"{cc_original} - {str(datos_c['Nombre'])}")
                            new_path = os.path.join(CARP_EXP, f"{m_cc} - {m_nom}")
                            if old_path != new_path:
                                try:
                                    if os.path.exists(old_path) and not os.path.exists(new_path):
                                        os.rename(old_path, new_path)
                                except Exception: pass
                                estructurar_carpetas(str(m_cc), m_nom)
                            
                            # 2. CONVERSIÓN OBLIGATORIA A TEXTO DE TODA LA TABLA (SOLUCIÓN A TYPEERROR)
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
                            
                            # 3. Efecto Cascada en las demás tablas
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
                            
# --- 3. CONTRATOS E INICIO Y LAS 30 PLANTILLAS DE WHATSAPP ---
elif st.session_state.pagina_actual == 'Contratos':
    st.markdown("<h1>Gestión Documental y Contratos</h1>", unsafe_allow_html=True)
    if not df_activos.empty:
        cli_sel = st.selectbox("Expediente:", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"])
        cc_s, nom_s = cli_sel.split(" - ")[0], cli_sel.split(" - ")[1]
        data_c = df_activos[df_activos["Cedula"].astype(str) == cc_s].iloc[0]
        senal, tel_c = int(data_c["Senal"]), str(data_c.get("Telefono", ""))
        rb = estructurar_carpetas(cc_s, nom_s)
        
        mostrar_boveda(cc_s, nom_s)

        c1, c2, c3, c4 = ["completed" if senal >= i else ("active" if senal == i-1 else "") for i in range(1, 5)]
        st.markdown(f'<div class="module-card" style="padding: 10px 30px;"><div class="timeline"><div class="step {c1}">1. Viabilidad</div><div class="step {c2}">2. Ensamblaje</div><div class="step {c3}">3. Firmas</div><div class="step {c4}">4. Radicado</div></div></div>', unsafe_allow_html=True)
        
        # --- CENTRAL DE WHATSAPP CON LAS 30 PLANTILLAS ORIGINALES ---
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
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
        st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
        
        # --- MOTOR DE GOOGLE DRIVE PARA CONTRATOS ---
        def obtener_archivos_y_id(cc, nom, subcarpeta):
            if not gc_drive: return [], None
            try:
                q1 = f"name='{cc} - {nom}' and '{CARPETA_RAIZ_DRIVE_ID}' in parents and trashed=false"
                r1 = gc_drive.files().list(q=q1, fields="files(id)").execute()
                if not r1.get('files'): return [], None
                cli_id = r1['files'][0]['id']
                
                q2 = f"name='{subcarpeta}' and '{cli_id}' in parents and trashed=false"
                r2 = gc_drive.files().list(q=q2, fields="files(id)").execute()
                if not r2.get('files'): return [], None
                sub_id = r2['files'][0]['id']
                
                q3 = f"'{sub_id}' in parents and mimeType != 'application/vnd.google-apps.folder' and trashed=false"
                r3 = gc_drive.files().list(q=q3, fields="files(name)").execute()
                return [a['name'] for a in r3.get('files', [])], sub_id
            except: return [], None

        def subir_a_drive_mem(file_obj, filename, folder_id):
            if not gc_drive or not folder_id: return
            try:
                from googleapiclient.http import MediaIoBaseUpload
                bytes_data = file_obj.getvalue()
                fh = io.BytesIO(bytes_data)
                
                meta = {'name': filename, 'parents': [folder_id]}
                # Quitamos el resumable=True para que suba directo y rápido
                media = MediaIoBaseUpload(fh, mimetype='application/pdf')
                
                gc_drive.files().create(body=meta, media_body=media, fields='id').execute()
            except Exception as e:
                print(f"Error en subir_a_drive_mem: {e}")

        def subir_doc_local_a_drive(ruta_local, filename, folder_id):
            if not gc_drive or not folder_id: return
            try:
                from googleapiclient.http import MediaFileUpload
                meta = {'name': filename, 'parents': [folder_id]}
                media = MediaFileUpload(ruta_local, resumable=True)
                gc_drive.files().create(body=meta, media_body=media, fields='id').execute()
            except Exception as e:
                print(f"Error en subir_doc_local_a_drive: {e}")

        if senal == 0:
            arch_drive, sub_id = obtener_archivos_y_id(cc_s, nom_s, "Financiero")
            st.markdown(f"<h3 style='font-size:18px;'>Fase 1: Recolección Documental ({len(arch_drive)}/8)</h3>", unsafe_allow_html=True)
            docs = ["1. CÉDULA", "2. CERTIFICADO REDAM", "3. DATACRÉDITO", "4. CERTIFICADO SIMIT", "5. CERTIFICADO RAMA", "6. CERTIFICADO RUNT", "7. TRADICIÓN Y LIBERTAD", "8. CERTIFICADO RUES"]
            
            with st.form("form_subida_masiva"):
                c_a, c_b = st.columns(2)
                uploaders = {}
                
                for i, d in enumerate(docs):
                    col = c_a if i < 4 else c_b
                    nombre_corto = d.split(". ")[1]
                    
                    if any(nombre_corto in f for f in arch_drive): 
                        col.success(f"✔️ {d} (Ya en Drive)")
                    else:
                        # Guardamos la referencia de cada uploader dentro del formulario
                        uploaders[nombre_corto] = col.file_uploader(f"📥 Subir: {d}", key=f"up_masivo_{i}")
                
                st.markdown("<br>", unsafe_allow_html=True)
                # Un solo botón maestro para procesar todo lo que se adjuntó
                btn_enviar_todo = st.form_submit_button("🚀 SUBIR TODOS LOS DOCUMENTOS SELECCIONADOS A DRIVE")
                
                if btn_enviar_todo:
                    if sub_id:
                        subidos_count = 0
                        with st.spinner("Subiendo documentos a Google Drive en lote..."):
                            for nombre_corto, upl in uploaders.items():
                                if upl is not None:
                                    nombre_archivo = f"{nombre_corto}_{upl.name}"
                                    subir_a_drive_mem(upl, nombre_archivo, sub_id)
                                    subidos_count += 1
                        
                        if subidos_count > 0:
                            st.success(f"¡Se subieron {subidos_count} documento(s) con éxito a Google Drive!")
                            
                            # Verificamos si ya se completaron los 8 documentos en total
                            arch_actualizados, _ = obtener_archivos_y_id(cc_s, nom_s, "Financiero")
                            if len(arch_actualizados) >= 8:
                                df_cli.loc[df_cli["Cedula"].astype(str) == str(cc_s), "Senal"] = 1
                                guardar_tabla(df_cli, "clientes")
                                st.info("🎉 ¡Fase 1 completada! El expediente avanzará a la Fase 2.")
                            
                            time.sleep(1.5)
                            st.rerun()
                        else:
                            st.warning("⚠️ No seleccionaste ningún archivo nuevo para subir.")
                        
        elif senal == 1:
            st.markdown("<h3 style='font-size:18px;'>Fase 2: Motor de Contratos</h3>", unsafe_allow_html=True)
            with st.form("motor"):
                ca, cb = st.columns(2)
                with ca: abo = st.selectbox("Abogado Titular de la Firma", ABOGADOS)
                with cb: h_n = st.text_input("Honorarios Totales (\$)"); h_l = st.text_input("Honorarios (En Letras)")
                cc_col, cd, ce = st.columns(3)
                with cc_col: cuo_n = st.text_input("Valor Cuota (\$)"); cuo_l = st.text_input("Valor Cuota (Letras)")
                with cd: ciu_e = st.text_input("Ciudad Expedición C.C."); ciu_r = st.text_input("Ciudad Residencia Actual")
                with ce: dir_r = st.text_input("Dirección de Residencia")
                
                if st.form_submit_button("GENERAR DOCUMENTOS Y SUBIR A NUBE"):
                    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
                    dic_r = { "«NOMBRES»": nom_s, "«CEDULA»": cc_s, "«EXP_CC»": ciu_e, "«DIRECCION»": dir_r, "«CIUDAD_DIRECCION»": ciu_r, "«CORREO»": data_c.get("Email",""), "«TELEFONO»": tel_c, "«TOTAL_PAGO»": h_n, "«TOTAL_PAGO_LETRA»": h_l, "«VALOR_CUOTA»": cuo_n, "«VALOR_CUOTA_LETRA»": cuo_l, "«DIA»": str(hoy.day), "«MES»": meses[hoy.month - 1] }
                    rutas = { "CONTRATO": os.path.join(CARP_PLA, abo, "CONTRATO.docx"), "PODER_JUZGADO": os.path.join(CARP_PLA, abo, "PODER_JUZGADO.docx") }
                    
                    _, sub_id_notaria = obtener_archivos_y_id(cc_s, nom_s, "Notaria")
                    
                    err = False
                    with st.spinner("Generando y subiendo contratos a Google Drive..."):
                        for n, r_p in rutas.items():
                            if os.path.exists(r_p): 
                                temp_salida = f"temp_{n}_{nom_s}.docx"
                                motor_docx(r_p, temp_salida, dic_r)
                                if sub_id_notaria:
                                    subir_doc_local_a_drive(temp_salida, f"{n}_{nom_s}.docx", sub_id_notaria)
                                if os.path.exists(temp_salida): os.remove(temp_salida)
                            else: 
                                err = True; st.error(f"Falta plantilla {n} en carpeta {abo}.")
                        
                        if not err:
                            if str(cc_s) not in df_fin["Cedula"].astype(str).values:
                                nuevo_f = pd.concat([df_fin, pd.DataFrame([{"Cedula": str(cc_s), "Honorarios": str(h_n), "Abonado": "0"}])], ignore_index=True)
                                guardar_tabla(nuevo_f, "finanzas")
                            df_cli.loc[df_cli["Cedula"].astype(str) == str(cc_s), "Senal"] = 2
                            guardar_tabla(df_cli, "clientes"); st.rerun()

        elif senal == 2:
            st.markdown("<h3 style='font-size:18px;'>Fase 3: Recolección de Firmas</h3>", unsafe_allow_html=True)
            st.info("📌 Los contratos listos para firmar están arriba, en la Bóveda Central. Descárgalos de ahí y sube los firmados aquí abajo.")
            
            _, sub_id_notaria = obtener_archivos_y_id(cc_s, nom_s, "Notaria")
            
            uf = st.file_uploader("📥 Subir Paquete de Contratos Firmados (PDF)", type=["pdf"])
            if uf and sub_id_notaria:
                with st.spinner("Subiendo a Drive..."):
                    subir_a_drive_mem(uf, f"Firmados_{uf.name}", sub_id_notaria)
                    df_cli.loc[df_cli["Cedula"].astype(str) == str(cc_s), "Senal"] = 3
                    guardar_tabla(df_cli, "clientes"); st.rerun()
                    
        elif senal == 3:
            st.markdown("<h3 style='font-size:18px;'>Fase 4: Radicación Oficial del Expediente</h3>", unsafe_allow_html=True)
            st.info("📌 Los contratos ya están firmados. Sube aquí el comprobante de radicación oficial (Centro de Conciliación o Juzgado).")
            urad = st.file_uploader("📥 Subir Soporte de Radicado (PDF/IMG)", type=["pdf", "jpg", "png"])
            _, sub_id_juz = obtener_archivos_y_id(cc_s, nom_s, "Juzgado")
            if urad and sub_id_juz:
                with st.spinner("Subiendo Radicado a Drive..."):
                    subir_a_drive_mem(urad, f"Radicado_{urad.name}", sub_id_juz)
                    df_cli.loc[df_cli["Cedula"].astype(str) == str(cc_s), "Senal"] = 4
                    guardar_tabla(df_cli, "clientes"); st.rerun()
                    
        elif senal >= 4: 
            st.success("✨ ¡Misión Cumplida! El expediente está oficialmente radicado y ha superado todas las fases documentales.")
            arch_juz, _ = obtener_archivos_y_id(cc_s, nom_s, "Juzgado")
            if arch_juz:
                for ar in arch_juz:
                    if "Radicado" in ar:
                        st.markdown(f"📄 **Soporte Oficial guardado en Drive:** `{ar}`")
            
        st.markdown("</div>", unsafe_allow_html=True)

# --- 4. ACTUACIONES ---
elif st.session_state.pagina_actual == 'Actuaciones':
    st.markdown("<h1>Seguimiento Procesal</h1>", unsafe_allow_html=True)
    if not df_activos.empty:
        c_a = st.selectbox("Expediente:", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"])
        cc_a, nom_a = c_a.split(" - ")[0], c_a.split(" - ")[1]
        rb = estructurar_carpetas(cc_a, nom_a)
        r4 = os.path.join(rb, "04_Log_Judicial")
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
                    with open(os.path.join(r4, f"{tipo_a}_{doc_up.name}"), "wb") as f: f.write(doc_up.getbuffer())
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
        with c2: fecha_v = st.date_input("Fecha Límite Exacta")
        if st.form_submit_button("PROGRAMAR ALERTA"):
            if not df_activos.empty and cl_v != "Sin Clientes":
                cc_v, nom_v = cl_v.split(" - ")[0], cl_v.split(" - ")[1]
                n_v = pd.DataFrame([{"ID_Ven": f"VEN-{hoy.strftime('%H%M%S')}", "Cedula": str(cc_v), "Cliente": nom_v, "Asunto": asunto, "Fecha_Limite": str(fecha_v), "Estado": "Activo"}])
                guardar_tabla(pd.concat([df_ven, n_v], ignore_index=True), "vencimientos"); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    if not df_ven.empty:
        for _, rv in df_ven[df_ven["Estado"] == "Activo"].iloc[::-1].iterrows():
            try:
                dias = (datetime.strptime(str(rv["Fecha_Limite"])[:10], "%Y-%m-%d").date() - hoy.date()).days
                if dias < 0: cb = "#71717A"; est = f"VENCIDO HACE {abs(dias)} DÍAS"
                elif dias <= 2: cb = "#EF4444"; est = f"¡PELIGRO! VENCE EN {dias} DÍAS"
                elif dias <= 5: cb = "#F59E0B"; est = f"ATENCIÓN: Faltan {dias} días"
                else: cb = "#10B981"; est = f"En término ({dias} días)"
                st.markdown(f"<div style='border-left: 5px solid {cb}; padding: 12px; background:#121214; margin-bottom:8px; border-radius:6px;'><b>{rv['Cliente']}</b> - {rv['Asunto']}<br>{est} (Límite: {rv['Fecha_Limite']})</div>", unsafe_allow_html=True)
            except: pass

# --- 6. ACREEDORES ---
elif st.session_state.pagina_actual == 'Acreedores':
    st.markdown("<h1>📋 Inventario de Pasivos y Acreedores</h1>", unsafe_allow_html=True)
    if not df_activos.empty:
        st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
        with st.form("form_acr"):
            cl_a = st.selectbox("Expediente de Insolvencia", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"])
            c1, c2, c3 = st.columns(3)
            with c1: nom_acr = st.text_input("Acreedor")
            with c2: cuantia_acr = st.number_input("Cuantía ($)", min_value=0, step=100000)
            with c3: clase_acr = st.selectbox("Clase", ["Primera", "Segunda", "Tercera", "Cuarta", "Quinta"])
            if st.form_submit_button("AGREGAR PASIVO"):
                cc_a = cl_a.split(" - ")[0]
                n_ac = pd.DataFrame([{"ID_Acr": f"ACR-{hoy.strftime('%H%M%S')}", "Cedula": str(cc_a), "Acreedor": nom_acr, "Cuantia": str(cuantia_acr), "Clase": clase_acr}])
                guardar_tabla(pd.concat([df_acr, n_ac], ignore_index=True), "acreedores"); st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

# --- 7. TAREAS ---
elif st.session_state.pagina_actual == 'Tareas':
    st.markdown("<h1>✅ Tablero de Misiones Privado</h1>", unsafe_allow_html=True)
    st.markdown("<div class='module-card module-card-blue'>", unsafe_allow_html=True)
    with st.form("form_tareas"):
        col1, col2 = st.columns([2, 1])
        with col1: desc_tar = st.text_input("Descripción de la Tarea")
        with col2: asignado = st.selectbox("Asignar a:", df_usr["Alias"].tolist())
        if st.form_submit_button("ENVIAR MISIÓN"):
            if desc_tar:
                n_t = pd.DataFrame([{"ID_Tar": f"TAR-{hoy.strftime('%H%M%S')}", "Tarea": desc_tar, "Asignado": asignado, "Creador": st.session_state.alias_actual, "Estado": "Pendiente", "Fecha": hoy.strftime("%Y-%m-%d")}])
                guardar_tabla(pd.concat([df_tar, n_t], ignore_index=True), "tareas"); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

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

# --- 9. AGENDA ---
elif st.session_state.pagina_actual == 'Agenda':
    st.markdown("<h1>📅 Agenda de Citas</h1>", unsafe_allow_html=True)
    st.markdown("<div class='module-card module-card-blue'>", unsafe_allow_html=True)
    with st.form("form_aud"):
        c1, c2 = st.columns(2)
        with c1:
            cl_aud = st.selectbox("Expediente", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"] if not df_activos.empty else ["Sin Clientes"])
            motivo = st.text_input("Motivo")
        with c2:
            fecha_a = st.date_input("Fecha")
            hora_a = st.time_input("Hora")
        if st.form_submit_button("AGENDAR"):
            if cl_aud != "Sin Clientes":
                cc_a, nom_a = cl_aud.split(" - ")[0], cl_aud.split(" - ")[1]
                n_au = pd.DataFrame([{"ID_Aud": f"AUD-{hoy.strftime('%H%M%S')}", "Cedula": str(cc_a), "Cliente": nom_a, "Fecha_Hora": f"{fecha_a} {hora_a}", "Motivo": motivo}])
                guardar_tabla(pd.concat([df_aud, n_au], ignore_index=True), "audiencias"); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)
    
# --- 10. FINANZAS ---
elif st.session_state.pagina_actual == 'Finanzas':
    if st.session_state.rol_actual != "Administrador (Jefa)": st.error("⛔ DENEGADO")
    else:
        st.markdown("<h1>💰 Finanzas y Facturación Cloud</h1>", unsafe_allow_html=True)
        
        # --- MOTOR GENERADOR DE FACTURAS PRO ---
        def generar_factura_pdf(nombre, cc, total, abono, saldo, recibo_id):
            pdf = FPDF()
            pdf.add_page()
            # Borde de página estético
            pdf.rect(5.0, 5.0, 200.0, 287.0)
            
            def cln(t): return str(t).encode('latin-1', 'replace').decode('latin-1')
            
            # Encabezado Oficial
            pdf.set_font("Helvetica", 'B', 20)
            pdf.set_text_color(212, 175, 55) # Color Dorado
            pdf.cell(0, 15, txt=cln("FIRMA JURÍDICA - INSOLVENCIA OS"), ln=True, align='C')
            pdf.set_font("Helvetica", 'B', 14)
            pdf.set_text_color(50, 50, 50)
            pdf.cell(0, 8, txt=cln("RECIBO DE CAJA / COMPROBANTE DE INGRESO"), ln=True, align='C')
            pdf.ln(5)
            
            # Datos del Documento
            pdf.set_font("Helvetica", 'B', 10)
            pdf.cell(0, 6, txt=cln(f"Recibo No: {recibo_id}"), ln=True, align='R')
            pdf.cell(0, 6, txt=cln(f"Fecha de Emisión: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"), ln=True, align='R')
            pdf.ln(10)
            
            # Cuadro 1: Datos del Cliente
            pdf.set_fill_color(240, 240, 240)
            pdf.set_font("Helvetica", 'B', 12)
            pdf.cell(0, 8, txt=cln(" DATOS DEL CLIENTE"), ln=True, align='L', fill=True)
            pdf.set_font("Helvetica", '', 11)
            pdf.cell(0, 6, txt=cln(f" Nombre / Razón Social: {nombre}"), ln=True)
            pdf.cell(0, 6, txt=cln(f" Cédula de Identidad: {cc}"), ln=True)
            pdf.ln(5)
            
            # Cuadro 2: Concepto del Servicio
            pdf.set_font("Helvetica", 'B', 12)
            pdf.cell(0, 8, txt=cln(" DETALLES DEL SERVICIO"), ln=True, align='L', fill=True)
            pdf.set_font("Helvetica", '', 11)
            pdf.multi_cell(0, 6, txt=cln("Concepto: Abono a honorarios profesionales por representación jurídica integral en trámite de insolvencia económica persona natural no comerciante (Ley 1564 de 2012)."))
            pdf.ln(5)
            
            # Cuadro 3: Desglose Contable
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
            
            pdf.set_fill_color(212, 175, 55) # Dorado
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
            
            # Firmas
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
        
        # --- LÓGICA DE FINANZAS ---
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
                            
                            # 1. Actualizar BD en Google Sheets
                            df_fu = leer_tabla("finanzas", ["Cedula", "Honorarios", "Abonado"])
                            df_fu.loc[df_fu["Cedula"].astype(str) == cc_f, "Abonado"] = str(int(nuevo_abonado))
                            guardar_tabla(df_fu, "finanzas")
                            
                            # 2. Generar Recibo ID
                            recibo_id = f"REC-{hoy.strftime('%Y%m%d%H%M%S')}"
                            
                            # 3. Generar PDF Pro
                            pdf_path = generar_factura_pdf(nom_f, cc_f, hon_t, n_abo, nuevo_saldo, recibo_id)
                            st.session_state['ultimo_pdf_factura'] = pdf_path
                            
                            # 4. Generar Link de WhatsApp
                            msg_wa = f"⚖️ *FIRMA JURÍDICA - INSOLVENCIA OS* ⚖️%0A%0AEstimado/a *{nom_f}*, desde el departamento financiero confirmamos la recepción exitosa de su pago.%0A%0A💰 *Abono registrado:* ${n_abo:,.0f}%0A📉 *Nuevo saldo pendiente:* ${nuevo_saldo:,.0f}%0A%0ASu recibo de caja oficial No. {recibo_id} ha sido generado en nuestro sistema. %0A%0A¡Gracias por su cumplimiento y confianza en nuestro equipo!"
                            st.session_state['ultimo_wa_factura'] = f"https://wa.me/57{str(tel_f).replace(' ', '')}?text={msg_wa}"
                            
                            st.success("¡Pago registrado en la nube con éxito!")
                            st.rerun()
                            
                # --- ZONA DE DESCARGA Y ENVÍO ---
                if 'ultimo_pdf_factura' in st.session_state and os.path.exists(st.session_state['ultimo_pdf_factura']):
                    st.markdown("<hr style='border-color: #27272A;'><h3 style='color:#10B981;'>🧾 Factura Oficial Generada</h3>", unsafe_allow_html=True)
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
    if st.session_state.rol_actual != "Administrador (Jefa)": st.error("⛔ DENEGADO")
    else:
        st.markdown("<h1>👥 Gestión de Accesos y Seguridad</h1>", unsafe_allow_html=True)
        
        # 1. Mostrar lista de usuarios
        st.markdown("<div class='module-card'>", unsafe_allow_html=True)
        st.markdown("<h3 style='font-size: 18px; margin-bottom: 10px;'>📋 Usuarios Activos en el Sistema</h3>", unsafe_allow_html=True)
        df_mostrar = df_usr[["Usuario", "Alias", "Rol", "Creador"]].copy()
        st.dataframe(df_mostrar, use_container_width=True, hide_index=True)
        st.markdown("</div>", unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        
        # 2. Panel Izquierdo: Crear Usuario
        with col1:
            st.markdown("<div class='module-card module-card-blue'>", unsafe_allow_html=True)
            st.markdown("<h3 style='font-size: 18px;'>➕ Crear Nuevo Usuario</h3>", unsafe_allow_html=True)
            with st.form("form_nuevo_usr"):
                n_user = st.text_input("Usuario de Login (Ej: mario_abogado)")
                n_pass = st.text_input("Contraseña", type="password")
                n_alias = st.text_input("Nombre a mostrar (Alias / Nombre Real)")
                n_rol = st.selectbox("Rol en la plataforma", ["Abogado / Operativo", "Administrador (Jefa)"])
                if st.form_submit_button("CREAR CREDENCIAL"):
                    if n_user and n_pass:
                        if str(n_user) in df_usr["Usuario"].astype(str).values:
                            st.error("❌ Ese usuario ya existe.")
                        else:
                            alias_final = n_alias if n_alias else n_user
                            nu = pd.DataFrame([{"Usuario": str(n_user), "Password": str(n_pass), "Rol": str(n_rol), "Creador": str(st.session_state.usuario_actual), "Alias": str(alias_final), "Avatar_Path": "", "Session_Token": ""}])
                            guardar_tabla(pd.concat([df_usr, nu], ignore_index=True), "usuarios")
                            st.success(f"Usuario {n_user} creado con éxito.")
                            time.sleep(1.5)
                            st.rerun()
                    else:
                        st.error("⚠ Usuario y Contraseña son obligatorios.")
            st.markdown("</div>", unsafe_allow_html=True)

        # 3. Panel Derecho: Editar, Eliminar y EXPULSAR Usuario
        with col2:
            st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
            st.markdown("<h3 style='font-size: 18px;'>✏️ Modificar o Eliminar Cuenta</h3>", unsafe_allow_html=True)
            
            usr_sel = st.selectbox("Selecciona un usuario a gestionar:", df_usr["Usuario"].tolist())
            
            if usr_sel:
                usr_data = df_usr[df_usr["Usuario"] == usr_sel].iloc[0]
                with st.form("form_edit_usr"):
                    e_alias = st.text_input("Cambiar Nombre / Alias", value=str(usr_data.get("Alias", usr_sel)))
                    e_pass = st.text_input("Nueva Contraseña (Dejar en blanco para no cambiar)", type="password")
                    idx_rol = 0 if usr_data["Rol"] == "Abogado / Operativo" else 1
                    e_rol = st.selectbox("Cambiar Rol", ["Abogado / Operativo", "Administrador (Jefa)"], index=idx_rol)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    col_btn1, col_btn2 = st.columns(2)
                    
                    with col_btn1:
                        btn_guardar = st.form_submit_button("💾 GUARDAR CAMBIOS")
                    with col_btn2:
                        st.markdown("""<style>div[data-testid="stFormSubmitButton"] button:contains('ELIMINAR') { background: #EF4444 !important; border-color: #EF4444 !important; color: white !important; }</style>""", unsafe_allow_html=True)
                        btn_eliminar = st.form_submit_button("🗑️ ELIMINAR USUARIO")
                        
                    if btn_guardar:
                        df_usr = df_usr.astype(str) # Blindaje de Pandas
                        
                        df_usr.loc[df_usr["Usuario"] == str(usr_sel), "Alias"] = str(e_alias)
                        df_usr.loc[df_usr["Usuario"] == str(usr_sel), "Rol"] = str(e_rol)
                        
                        clave_cambiada = False
                        if e_pass.strip():  # Si escribiste una clave nueva...
                            df_usr.loc[df_usr["Usuario"] == str(usr_sel), "Password"] = str(e_pass)
                            # 🔥 LA MAGIA: Invalidamos el Token de Sesión en la Base de Datos
                            df_usr.loc[df_usr["Usuario"] == str(usr_sel), "Session_Token"] = "REVOCADO_POR_SEGURIDAD"
                            clave_cambiada = True
                            
                        guardar_tabla(df_usr, "usuarios")
                        registrar_log("SEGURIDAD", f"Actualizó credenciales de: {usr_sel}")
                        
                        if clave_cambiada:
                            st.success(f"✅ Contraseña cambiada. La sesión de '{usr_sel}' ha sido cerrada forzosamente.")
                            # Si te cambiaste la clave a ti mismo, el sistema te saca a ti también de inmediato
                            if usr_sel == st.session_state.usuario_actual:
                                st.session_state.autenticado = False
                                st.session_state.kicked = True
                                st.session_state.kicked_reason = "Tu contraseña fue actualizada. Inicia sesión nuevamente por seguridad."
                        else:
                            st.success("✅ Cambios guardados correctamente.")
                            
                        time.sleep(2.5)
                        st.rerun()
                        
                    if btn_eliminar:
                        if usr_sel == st.session_state.usuario_actual:
                            st.error("❌ ¡Peligro! No puedes eliminar tu propia cuenta mientras la usas.")
                        elif usr_sel == "admin" and len(df_usr[df_usr["Rol"] == "Administrador (Jefa)"]) == 1:
                            st.error("❌ No puedes eliminar al único Administrador del sistema.")
                        else:
                            df_usr = df_usr[df_usr["Usuario"] != usr_sel]
                            guardar_tabla(df_usr, "usuarios")
                            registrar_log("SEGURIDAD", f"Eliminó al usuario: {usr_sel}")
                            st.success(f"🗑️ Usuario {usr_sel} eliminado permanentemente.")
                            time.sleep(1.5)
                            st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
            
# --- 12. SISTEMA ---
elif st.session_state.pagina_actual == 'Sistema':
    if st.session_state.rol_actual != "Administrador (Jefa)": st.error("⛔ DENEGADO")
    else:
        st.markdown("<h1>🛡️ Auditoría y Exportación del Sistema</h1>", unsafe_allow_html=True)
        st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
        st.success("☁️ Base de datos sincronizada y operando 100% sobre Google Sheets Master.")
        
        # --- BOTONES DE DESCARGA DE RESGUARDOS ---
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
        
        # --- TERMINAL DE AUDITORÍA QUIRÚRGICA POR PERFIL ---
        st.markdown("<h3>🔍 Terminal de Auditoría Detallada por Usuario</h3>", unsafe_allow_html=True)
        st.markdown("<p style='color: #64748B; font-size: 13px;'>Selecciona un colaborador activo para revisar el rastro exacto de todas sus operaciones.</p>", unsafe_allow_html=True)
        
        # 🔥 SOLUCIÓN: Extraemos ÚNICAMENTE los usuarios que existen ACTUALMENTE en la base de datos
        usuarios_auditables = sorted(df_usr["Alias"].dropna().unique().tolist()) if not df_usr.empty else []
            
        if usuarios_auditables:
            usuario_seleccionado = st.selectbox("👤 Seleccionar Colaborador a Auditar:", usuarios_auditables)
            
            if usuario_seleccionado:
                st.markdown(f"<br><h4 style='color: #2563EB;'>📜 Historial de Actividad: {usuario_seleccionado}</h4>", unsafe_allow_html=True)
                
                # Filtrar logs de este usuario específico
                if not df_log.empty:
                    df_log_usuario = df_log[df_log["Usuario"].astype(str) == str(usuario_seleccionado)].copy()
                    
                    if not df_log_usuario.empty:
                        # Mostrar el número total de acciones registradas
                        st.info(f"Se encontraron **{len(df_log_usuario)}** acciones registradas en el sistema para este perfil.")
                        # Mostrar la tabla ordenada de más reciente a más antigua
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
    
    # Buscamos los clientes que tienen el estado en "Borrado"
    df_borrados = df_cli[df_cli["Estado"] == "Borrado"].copy()
    
    if not df_borrados.empty:
        st.info("⚠️ Los expedientes aquí estarán por 30 días antes de su eliminación permanente. Puedes restaurarlos si fue un error.")
        
        for _, r in df_borrados.iterrows():
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"<h4 style='margin-bottom: 0;'>{r['Cedula']} - {r['Nombre']}</h4>", unsafe_allow_html=True)
                st.markdown(f"<span style='color: #EF4444; font-size: 13px; font-weight: bold;'>Enviado a papelera el: {r.get('F_Borrado', 'Desconocido')}</span>", unsafe_allow_html=True)
            with col2:
                # Botón de rescate
                if st.button("♻️ Restaurar Cliente", key=f"res_{r['Cedula']}"):
                    with st.spinner("Restaurando expediente y carpeta en Drive..."):
                        # 1. Volver a activarlo en Google Sheets
                        df_cli.loc[df_cli["Cedula"].astype(str) == str(r['Cedula']), "Estado"] = "Activo"
                        df_cli.loc[df_cli["Cedula"].astype(str) == str(r['Cedula']), "F_Borrado"] = ""
                        guardar_tabla(df_cli, "clientes")
                        
                        # 2. Devolver la carpeta de Google Drive a la Bóveda
                        restaurar_de_papelera_drive(str(r['Cedula']), r['Nombre'])
                        
                        registrar_log("SISTEMA", f"Restauró expediente de la papelera: {r['Cedula']}")
                        st.success("¡Expediente restaurado con éxito!")
                        time.sleep(1.5)
                        st.rerun()
            st.markdown("<hr style='border-color: #E2E8F0; margin: 10px 0;'>", unsafe_allow_html=True)
    else:
        st.success("✨ La papelera está vacía. No hay expedientes eliminados.")
        
    st.markdown("</div>", unsafe_allow_html=True)
