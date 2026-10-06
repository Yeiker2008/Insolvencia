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

# ==========================================
# 1. CORE & CONFIGURACIÓN DE SESIÓN
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

# ==========================================
# 2. MOTOR CSS: OBSIDIAN & GOLD 
# ==========================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;700&family=Playfair+Display:wght@600;800;900&display=swap');
    #MainMenu, footer {visibility: hidden;}
    [data-testid="collapsedControl"] { visibility: visible !important; display: block !important; background-color: #121214 !important; border: 1px solid #D4AF37 !important; border-radius: 6px !important; color: #D4AF37 !important; z-index: 999999; }
    [data-testid="collapsedControl"] svg { fill: #D4AF37 !important; }
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #050505 !important; color: #E4E4E7 !important; }
    h1, h2, h3 { font-family: 'Playfair Display', serif !important; color: #FACC15 !important; letter-spacing: -0.5px; }
    [data-testid="stAppViewContainer"] { background-color: #050505 !important; background-image: linear-gradient(rgba(255,255,255,0.02) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.02) 1px, transparent 1px); background-size: 30px 30px; }
    [data-testid="stSidebar"] { background-color: #09090B !important; border-right: 1px solid #27272A !important; }
    input, textarea, select, div[data-baseweb="select"] > div, div[data-baseweb="input"] > div { background-color: #121214 !important; color: #F8FAFC !important; -webkit-text-fill-color: #F8FAFC !important; border: 1px solid #27272A !important; border-radius: 6px !important; }
    input:focus, textarea:focus { border-color: #D4AF37 !important; box-shadow: 0 0 10px rgba(212,175,55,0.2) !important; }
    .module-card { background: linear-gradient(145deg, #121214 0%, #09090B 100%); border: 1px solid #27272A; border-radius: 12px; padding: 30px; margin-bottom: 25px; box-shadow: 0 10px 30px rgba(0,0,0,0.8); }
    .module-card-gold { border-top: 3px solid #D4AF37; }
    .module-card-blue { border-top: 3px solid #3B82F6; }
    .module-card-green { border-top: 3px solid #10B981; }
    div.stButton > button:first-child { background: #121214 !important; color: #D4AF37 !important; border: 1px solid #D4AF37 !important; border-radius: 4px !important; padding: 10px 20px !important; font-weight: 700 !important; text-transform: uppercase; font-size: 13px !important; transition: all 0.3s ease !important; width: 100%;}
    div.stButton > button:first-child:hover { background: #D4AF37 !important; color: #050505 !important; box-shadow: 0 0 25px rgba(212,175,55,0.4) !important; transform: translateY(-2px); }
    .timeline { display: flex; justify-content: space-between; align-items: center; margin: 30px 0; position: relative; }
    .timeline::before { content: ''; position: absolute; top: 50%; left: 0; right: 0; height: 2px; background: #27272A; z-index: 1; }
    .step { position: relative; z-index: 2; background: #050505; padding: 8px 16px; border-radius: 20px; border: 2px solid #27272A; color: #71717A; font-weight: 600; font-size: 12px; display: flex; align-items: center; text-transform: uppercase; }
    .step.active { border-color: #D4AF37; color: #D4AF37; box-shadow: 0 0 15px rgba(212,175,55,0.3); background: #121214; }
    .step.completed { border-color: #10B981; color: #10B981; background: #064E3B; }
    .alerta-roja { animation: blinker 1.5s linear infinite; color: #EF4444 !important; font-weight: bold;}
    @keyframes blinker { 50% { opacity: 0; } }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. BASE DE DATOS Y FUNCIONES NÚCLEO
# ==========================================
ARCH_CLI, ARCH_FIN, ARCH_ACT, ARCH_LOG = "db_clientes.csv", "db_finanzas.csv", "db_actuaciones.csv", "db_logs.csv"
ARCH_USR, ARCH_VEN, ARCH_ACR, ARCH_AUD, ARCH_TAR = "db_usuarios.csv", "db_vencimientos.csv", "db_acreedores.csv", "db_audiencias.csv", "db_tareas.csv"
CARP_EXP, CARP_PLA, CARP_PAP, CARP_AVATAR = "Expedientes", "Plantillas", "Papelera", "Avatares"
ABOGADOS = ["David Alexander Arias Posse", "Maria Jose Ospino Perez", "Yeiker Cardona"]

for c in [CARP_EXP, CARP_PLA, CARP_PAP, CARP_AVATAR]: os.makedirs(c, exist_ok=True)
for a in ABOGADOS: os.makedirs(os.path.join(CARP_PLA, a), exist_ok=True)

if not os.path.exists(ARCH_CLI): pd.DataFrame(columns=["Cedula", "Nombre", "Fuerza", "Telefono", "Email", "Deuda_Est", "Ingresos", "Senal", "F_Actualizacion", "Estado", "F_Borrado"]).to_csv(ARCH_CLI, index=False)
if not os.path.exists(ARCH_FIN): pd.DataFrame(columns=["Cedula", "Honorarios", "Abonado"]).to_csv(ARCH_FIN, index=False)
if not os.path.exists(ARCH_ACT): pd.DataFrame(columns=["ID_Act", "Cedula", "Fecha", "Tipo", "Juzgado", "Radicado", "Anotacion"]).to_csv(ARCH_ACT, index=False)
if not os.path.exists(ARCH_LOG): pd.DataFrame(columns=["Timestamp", "Usuario", "Modulo", "Accion"]).to_csv(ARCH_LOG, index=False)
if not os.path.exists(ARCH_VEN): pd.DataFrame(columns=["ID_Ven", "Cedula", "Cliente", "Asunto", "Fecha_Limite", "Estado"]).to_csv(ARCH_VEN, index=False)
if not os.path.exists(ARCH_ACR): pd.DataFrame(columns=["ID_Acr", "Cedula", "Acreedor", "Cuantia", "Clase"]).to_csv(ARCH_ACR, index=False)
if not os.path.exists(ARCH_AUD): pd.DataFrame(columns=["ID_Aud", "Cedula", "Cliente", "Fecha_Hora", "Motivo"]).to_csv(ARCH_AUD, index=False)
if not os.path.exists(ARCH_TAR): pd.DataFrame(columns=["ID_Tar", "Tarea", "Asignado", "Creador", "Estado", "Fecha"]).to_csv(ARCH_TAR, index=False)
if not os.path.exists(ARCH_USR): pd.DataFrame([{"Usuario": "admin", "Password": "123", "Rol": "Administrador (Jefa)", "Creador": "Sistema", "Alias": "Administración", "Avatar_Path": "", "Session_Token": ""}]).to_csv(ARCH_USR, index=False)

df_usr = pd.read_csv(ARCH_USR)
if "Alias" not in df_usr.columns: df_usr["Alias"] = df_usr["Usuario"]
if "Avatar_Path" not in df_usr.columns: df_usr["Avatar_Path"] = ""
if "Session_Token" not in df_usr.columns: df_usr["Session_Token"] = ""
df_usr.to_csv(ARCH_USR, index=False)

df_cli = pd.read_csv(ARCH_CLI)
df_fin = pd.read_csv(ARCH_FIN)
df_act = pd.read_csv(ARCH_ACT)
df_log = pd.read_csv(ARCH_LOG)
df_ven = pd.read_csv(ARCH_VEN)
df_acr = pd.read_csv(ARCH_ACR)
df_aud = pd.read_csv(ARCH_AUD)
df_tar = pd.read_csv(ARCH_TAR)
hoy = datetime.now()

cambios_bd = False
for idx, r in df_cli[df_cli["Estado"] == "Borrado"].iterrows():
    if pd.notna(r["F_Borrado"]) and str(r["F_Borrado"]).strip() != "":
        if (hoy - datetime.strptime(r["F_Borrado"], "%Y-%m-%d")).days >= 30:
            rp = os.path.join(CARP_PAP, f"{r['Cedula']} - {r['Nombre']}")
            if os.path.exists(rp): shutil.rmtree(rp)
            df_cli = df_cli.drop(idx)
            cambios_bd = True
if cambios_bd: df_cli.to_csv(ARCH_CLI, index=False)

df_activos = df_cli[df_cli["Estado"] == "Activo"].copy()
df_papelera = df_cli[df_cli["Estado"] == "Borrado"].copy()

def registrar_log(modulo, accion):
    try:
        df_l = pd.read_csv(ARCH_LOG)
        usr = st.session_state.get('alias_actual', 'Sistema')
        nuevo = pd.DataFrame([{"Timestamp": hoy.strftime("%Y-%m-%d %H:%M:%S"), "Usuario": usr, "Modulo": modulo, "Accion": accion}])
        pd.concat([df_l, nuevo], ignore_index=True).to_csv(ARCH_LOG, index=False)
    except: pass

def estructurar_carpetas(cedula, nombre):
    base = os.path.join(CARP_EXP, f"{cedula} - {nombre}")
    for c in ["01_Docs_Viabilidad", "02_Contratos_Firmas", "03_Soporte_Radicacion", "04_Log_Judicial", "05_Pruebas_Anexos"]:
        os.makedirs(os.path.join(base, c), exist_ok=True)
    return base

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
        ics_content = f"BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//Insolvencia OS//Agenda//ES\nBEGIN:VEVENT\nSUMMARY:⚖️️ {motivo} - {cliente}\nDTSTART:{dt_start.strftime(formato)}\nDTEND:{dt_end.strftime(formato)}\nDESCRIPTION:Audiencia/Cita programada desde el sistema de gestión.\nEND:VEVENT\nEND:VCALENDAR"
        return ics_content.encode('utf-8')
    except: return b""

def mostrar_boveda(cc, nom):
    with st.expander("📂 BÓVEDA CENTRAL DE ARCHIVOS DEL EXPEDIENTE", expanded=False):
        rb = estructurar_carpetas(cc, nom)
        carpetas = ["01_Docs_Viabilidad", "02_Contratos_Firmas", "03_Soporte_Radicacion", "04_Log_Judicial", "05_Pruebas_Anexos"]
        cols = st.columns(2)
        for i, c in enumerate(carpetas):
            col = cols[i % 2]
            archivos = os.listdir(os.path.join(rb, c))
            col.markdown(f"<p style='color:#D4AF37; margin-bottom:2px; font-weight:bold;'>{c.replace('_', ' ').title()}</p>", unsafe_allow_html=True)
            if archivos:
                for a in archivos:
                    ruta = os.path.join(rb, c, a)
                    if os.path.isfile(ruta):
                        with open(ruta, "rb") as f:
                            col.download_button(f"📄 {a[:25]}...", f, file_name=a, key=f"dl_{cc}_{c}_{a}")
            else:
                col.markdown("<span style='color:#71717A; font-size:12px;'><i>Carpeta vacía</i></span>", unsafe_allow_html=True)
            col.write("")

# ==========================================
# 4. PANTALLA DE LOGIN Y SEGURIDAD ANTICLONACIÓN
# ==========================================
if st.session_state.autenticado:
    try:
        df_u_check = pd.read_csv(ARCH_USR)
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
        st.markdown("""<div class="module-card module-card-gold" style="text-align: center; padding: 40px;"><span style="font-size: 50px;">⚖️</span><h2>ACCESO RESTRINGIDO</h2><p style="color: #A1A1AA; font-size: 13px; margin-bottom: 30px;">Plataforma LegalTech Cifrada</p>""", unsafe_allow_html=True)
        with st.form("login_form"):
            user_input = st.text_input("👤 Usuario")
            pass_input = st.text_input("🔑 Contraseña", type="password")
            if st.form_submit_button("AUTENTICAR CREDENCIALES"):
                df_u = pd.read_csv(ARCH_USR)
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
                    df_u.to_csv(ARCH_USR, index=False)
                    registrar_log("AUTH", f"Inició sesión: {user_input}")
                    st.rerun()
                else: st.error("❌ Credenciales inválidas.")
        st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

# ==========================================
# 5. SIDEBAR ORDENADO POR SECCIONES
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
                df_u_up = pd.read_csv(ARCH_USR)
                p_guardado = st.session_state.avatar_path
                if foto_sub:
                    p_guardado = os.path.join(CARP_AVATAR, f"{st.session_state.usuario_actual}_{foto_sub.name}")
                    with open(p_guardado, "wb") as f: f.write(foto_sub.getbuffer())
                df_u_up.loc[df_u_up["Usuario"].astype(str) == st.session_state.usuario_actual, "Alias"] = n_alias
                df_u_up.loc[df_u_up["Usuario"].astype(str) == st.session_state.usuario_actual, "Avatar_Path"] = p_guardado
                df_u_up.to_csv(ARCH_USR, index=False)
                st.session_state.alias_actual = n_alias; st.session_state.avatar_path = p_guardado
                st.success("Perfil Actualizado"); st.rerun()

    st.markdown("<hr style='border-color: #27272A; margin-top: 0;'>", unsafe_allow_html=True)
    
    st.markdown("<p style='color:#71717A; font-size:11px; font-weight:bold; letter-spacing:1px;'>🏠 INICIO & GESTIÓN</p>", unsafe_allow_html=True)
    if st.button("📊 Portal Ejecutivo", key="b_dash"): cambiar_pagina("Dashboard")
    if st.button("📝 Apertura de Casos", key="b_nuev"): cambiar_pagina("Nuevo")
    if st.button("📅 Agenda y Citas", key="b_agen"): cambiar_pagina("Agenda")
    if st.button("✅ Gestor de Tareas", key="b_tar"): cambiar_pagina("Tareas")
    
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
# 6. MÓDULOS DE LA APLICACIÓN
# ==========================================

# --- 1. PORTAL EJECUTIVO ---
if st.session_state.pagina_actual == 'Dashboard':
    st.markdown("<h1>Portal Ejecutivo Legal</h1><p style='margin-bottom: 20px;'>Centro de operaciones e inteligencia del despacho.</p>", unsafe_allow_html=True)
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
                dias = (hoy - datetime.strptime(r["F_Actualizacion"], "%Y-%m-%d")).days
                if r["Senal"] == 2 and dias >= 2:
                    hay_alertas = True; st.warning(f"⏳ **Firma pendiente:** {r['Nombre']} ({dias} días).")
                elif r["Senal"] == 3 and dias >= 1:
                    hay_alertas = True; st.error(f"🚨 **Radicación urgente:** {r['Nombre']} ({dias} días).")
        if not hay_alertas: st.success("✨ Expedientes fluyendo con normalidad.")
        st.markdown("</div>", unsafe_allow_html=True)

# --- 2. APERTURA ---
elif st.session_state.pagina_actual == 'Nuevo':
    st.markdown("<h1>Apertura de Expediente</h1>", unsafe_allow_html=True)
    st.markdown("<div class='module-card'>", unsafe_allow_html=True)
    with st.form("nuevo_exp"):
        c1, c2, c3 = st.columns(3)
        with c1: cc = st.text_input("Cédula de Ciudadanía")
        with c2: nom = st.text_input("Nombre Completo")
        with c3: fza = st.selectbox("Facción", ["Policía Nacional", "Ejército Nacional", "Armada Nacional", "Fuerza Aérea", "Retirado / Pensionado", "Civil"])
        c4, c5, c6 = st.columns(3)
        with c4: tel = st.text_input("WhatsApp / Celular")
        with c5: mail = st.text_input("Correo Electrónico")
        with c6: deuda = st.text_input("Deuda Aprox ($)")
        if st.form_submit_button("CREAR BÓVEDA EN SERVIDOR"):
            if not cc or not nom: st.error("Cédula y Nombre son obligatorios.")
            elif str(cc) in df_cli["Cedula"].astype(str).values: st.error("Sujeto ya existe en la base de datos.")
            else:
                ndf = pd.DataFrame([{"Cedula": cc, "Nombre": nom, "Fuerza": fza, "Telefono": tel, "Email": mail, "Deuda_Est": deuda, "Ingresos": "", "Senal": 0, "F_Actualizacion": hoy.strftime("%Y-%m-%d"), "Estado": "Activo", "F_Borrado": ""}])
                pd.concat([df_cli, ndf], ignore_index=True).to_csv(ARCH_CLI, index=False)
                estructurar_carpetas(cc, nom); st.success("Expediente Centralizado Exitosamente."); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# --- 3. CONTRATOS E INICIO ---
elif st.session_state.pagina_actual == 'Contratos':
    st.markdown("<h1>Gestión Documental y Contratos</h1>", unsafe_allow_html=True)
    if not df_activos.empty:
        cli_sel = st.selectbox("Expediente:", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"])
        cc_s, nom_s = cli_sel.split(" - ")[0], cli_sel.split(" - ")[1]
        data_c = df_activos[df_activos["Cedula"].astype(str) == cc_s].iloc[0]
        senal, tel_c = data_c["Senal"], str(data_c.get("Telefono", ""))
        rb = estructurar_carpetas(cc_s, nom_s)
        r1, r2, r3 = [os.path.join(rb, c) for c in ["01_Docs_Viabilidad", "02_Contratos_Firmas", "03_Soporte_Radicacion"]]
        
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
        if senal == 0:
            arch = os.listdir(r1)
            st.markdown(f"<h3 style='font-size:18px;'>Fase 1: Recolección Documental ({len(arch)}/8)</h3>", unsafe_allow_html=True)
            docs = ["1. CÉDULA", "2. CERTIFICADO REDAM", "3. DATACRÉDITO", "4. CERTIFICADO SIMIT", "5. CERTIFICADO RAMA", "6. CERTIFICADO RUNT", "7. TRADICIÓN Y LIBERTAD", "8. CERTIFICADO RUES"]
            c_a, c_b = st.columns(2)
            for i, d in enumerate(docs):
                col = c_a if i < 4 else c_b
                if any(d.split(". ")[1] in f for f in arch): col.success(f"✔️ {d}")
                else:
                    upl = col.file_uploader(f"📥 Subir: {d}", key=f"up_{i}")
                    if upl:
                        with open(os.path.join(r1, f"{d.split('. ')[1]}_{upl.name}"), "wb") as f: f.write(upl.getbuffer())
                        if len(os.listdir(r1)) == 8:
                            df_cli.loc[df_cli["Cedula"].astype(str) == str(cc_s), "Senal"] = 1
                            df_cli.to_csv(ARCH_CLI, index=False)
                        st.rerun()
        
        elif senal == 1:
            st.markdown("<h3 style='font-size:18px;'>Fase 2: Motor de Contratos</h3>", unsafe_allow_html=True)
            with st.form("motor"):
                ca, cb = st.columns(2)
                with ca: abo = st.selectbox("Abogado Titular de la Firma", ABOGADOS)
                with cb: h_n = st.text_input("Honorarios Totales ($)"); h_l = st.text_input("Honorarios (En Letras)")
                cc, cd, ce = st.columns(3)
                with cc: cuo_n = st.text_input("Valor Cuota ($)"); cuo_l = st.text_input("Valor Cuota (Letras)")
                with cd: ciu_e = st.text_input("Ciudad Expedición C.C."); ciu_r = st.text_input("Ciudad Residencia Actual")
                with ce: dir_r = st.text_input("Dirección de Residencia")
                
                if st.form_submit_button("GENERAR DOCUMENTOS Y AVANZAR FASE"):
                    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
                    dic_r = { "«NOMBRES»": nom_s, "«CEDULA»": cc_s, "«EXP_CC»": ciu_e, "«DIRECCION»": dir_r, "«CIUDAD_DIRECCION»": ciu_r, "«CORREO»": data_c.get("Email",""), "«TELEFONO»": tel_c, "«TOTAL_PAGO»": h_n, "«TOTAL_PAGO_LETRA»": h_l, "«VALOR_CUOTA»": cuo_n, "«VALOR_CUOTA_LETRA»": cuo_l, "«DIA»": str(hoy.day), "«MES»": meses[hoy.month - 1] }
                    rutas = { "CONTRATO": os.path.join(CARP_PLA, abo, "CONTRATO.docx"), "PODER_JUZGADO": os.path.join(CARP_PLA, abo, "PODER_JUZGADO.docx") }
                    
                    err = False
                    for n, r_p in rutas.items():
                        if os.path.exists(r_p): motor_docx(r_p, os.path.join(r2, f"{n}_{nom_s}.docx"), dic_r)
                        else: err = True; st.error(f"Falta plantilla {n} en carpeta {abo}.")
                    if not err:
                        dff = pd.read_csv(ARCH_FIN)
                        if str(cc_s) not in dff["Cedula"].astype(str).values: pd.concat([dff, pd.DataFrame([{"Cedula": cc_s, "Honorarios": h_n, "Abonado": "0"}])], ignore_index=True).to_csv(ARCH_FIN, index=False)
                        df_cli.loc[df_cli["Cedula"].astype(str) == str(cc_s), "Senal"] = 2
                        df_cli.to_csv(ARCH_CLI, index=False); st.rerun()
        
        elif senal == 2:
            st.markdown("<h3 style='font-size:18px;'>Fase 3: Recolección de Firmas</h3>", unsafe_allow_html=True)
            for dg in os.listdir(r2):
                ruta = os.path.join(r2, dg)
                if os.path.isfile(ruta):
                    with open(ruta, "rb") as f: st.download_button(f"📥 Imprimir Documento: {dg}", f, file_name=dg)
            uf = st.file_uploader("📥 Subir Paquete de Contratos Firmados (PDF)", type=["pdf"])
            if uf:
                with open(os.path.join(r2, f"Firmados_{uf.name}"), "wb") as f: f.write(uf.getbuffer())
                df_cli.loc[df_cli["Cedula"].astype(str) == str(cc_s), "Senal"] = 3
                df_cli.to_csv(ARCH_CLI, index=False); st.rerun()
        
        elif senal >= 3: st.success("✨ El expediente ha superado las fases documentales. Revise el módulo de Seguimiento Procesal.")
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
            if st.form_submit_button("REGISTRAR ACTUACIÓN EN SISTEMA"):
                if doc_up:
                    with open(os.path.join(r4, f"{tipo_a}_{doc_up.name}"), "wb") as f: f.write(doc_up.getbuffer())
                n_act = pd.DataFrame([{"ID_Act": f"ACT-{hoy.strftime('%H%M%S')}", "Cedula": cc_a, "Fecha": hoy.strftime("%Y-%m-%d"), "Tipo": tipo_a, "Juzgado": juzgado, "Radicado": rad_jud, "Anotacion": anota}])
                pd.concat([df_act, n_act], ignore_index=True).to_csv(ARCH_ACT, index=False); st.success("Registrado con éxito."); st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("### 📂 Historial de Actuaciones:")
        if not df_act.empty:
            acts_cliente = df_act[df_act["Cedula"].astype(str) == cc_a]
            if not acts_cliente.empty:
                st.dataframe(acts_cliente[["Fecha", "Tipo", "Juzgado", "Radicado", "Anotacion"]], use_container_width=True, hide_index=True)
            else: st.info("No hay actuaciones jurídicas registradas para este expediente.")

# --- 5. VENCIMIENTOS ---
elif st.session_state.pagina_actual == 'Vencimientos':
    st.markdown("<h1>🚦 Control de Vencimientos Automático</h1><p style='margin-bottom: 20px;'>Evita caducidades legales con el semáforo inteligente.</p>", unsafe_allow_html=True)
    st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
    with st.form("form_ven"):
        c1, c2 = st.columns(2)
        with c1:
            cl_v = st.selectbox("Seleccionar Cliente", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"] if not df_activos.empty else ["Sin Clientes"])
            asunto = st.text_input("Asunto Legal (Ej: Descorrer traslado 3 días)")
        with c2:
            fecha_v = st.date_input("Fecha Límite Exacta")
        if st.form_submit_button("PROGRAMAR ALERTA DE VENCIMIENTO"):
            if not df_activos.empty and cl_v != "Sin Clientes":
                cc_v, nom_v = cl_v.split(" - ")[0], cl_v.split(" - ")[1]
                n_v = pd.DataFrame([{"ID_Ven": f"VEN-{hoy.strftime('%H%M%S')}", "Cedula": cc_v, "Cliente": nom_v, "Asunto": asunto, "Fecha_Limite": str(fecha_v), "Estado": "Activo"}])
                pd.concat([df_ven, n_v], ignore_index=True).to_csv(ARCH_VEN, index=False); st.success("Plazo registrado en el motor."); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    if not df_ven.empty:
        for _, rv in df_ven[df_ven["Estado"] == "Activo"].iloc[::-1].iterrows():
            dias = (datetime.strptime(rv["Fecha_Limite"], "%Y-%m-%d").date() - hoy.date()).days
            if dias < 0:
                cb = "#71717A"; estado_txt = f"<span style='color:#71717A; font-weight:bold;'>VENCIDO HACE {abs(dias)} DÍAS</span>"
            elif dias <= 2:
                cb = "#EF4444"; estado_txt = f"<span class='alerta-roja'>¡PELIGRO INMINENTE! VENCE EN {dias} DÍAS</span>"
            elif dias <= 5:
                cb = "#F59E0B"; estado_txt = f"<span style='color:#F59E0B; font-weight:bold;'>ATENCIÓN: Faltan {dias} días</span>"
            else:
                cb = "#10B981"; estado_txt = f"<span style='color:#10B981; font-weight:bold;'>En término adecuado ({dias} días)</span>"
                
            st.markdown(f"<div style='background: #121214; border: 1px solid #27272A; border-left: 5px solid {cb}; padding: 15px; border-radius: 8px; margin-bottom: 10px;'><b style='color:white;'>{rv['Cliente']}</b> - {rv['Asunto']}<br>{estado_txt} (Fecha Límite: {rv['Fecha_Limite']})</div>", unsafe_allow_html=True)
            if st.button(f"Marcar como Cumplido", key=f"c_{rv['ID_Ven']}"):
                df_v_up = pd.read_csv(ARCH_VEN)
                df_v_up.loc[df_v_up["ID_Ven"] == rv["ID_Ven"], "Estado"] = "Completado"
                df_v_up.to_csv(ARCH_VEN, index=False); st.rerun()

# --- 6. ACREEDORES ---
elif st.session_state.pagina_actual == 'Acreedores':
    st.markdown("<h1>📋 Inventario de Pasivos y Acreedores</h1>", unsafe_allow_html=True)
    if not df_activos.empty:
        st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
        with st.form("form_acr"):
            cl_a = st.selectbox("Expediente de Insolvencia", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"])
            c1, c2, c3 = st.columns(3)
            with c1: nom_acr = st.text_input("Nombre Entidad o Persona (Acreedor)")
            with c2: cuantia_acr = st.number_input("Cuantía Adeudada ($)", min_value=0, step=100000)
            with c3: clase_acr = st.selectbox("Clase de Acreencia", ["Primera (Laboral / Fiscal)", "Segunda (Prendaria)", "Tercera (Hipotecaria)", "Cuarta (Proveedores)", "Quinta (Quirografaria / Bancos)"])
            if st.form_submit_button("AGREGAR PASIVO AL EXPEDIENTE"):
                cc_a = cl_a.split(" - ")[0]
                n_ac = pd.DataFrame([{"ID_Acr": f"ACR-{hoy.strftime('%H%M%S')}", "Cedula": cc_a, "Acreedor": nom_acr, "Cuantia": str(cuantia_acr), "Clase": clase_acr}])
                pd.concat([df_acr, n_ac], ignore_index=True).to_csv(ARCH_ACR, index=False); st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("### 📊 Relación de Créditos Global")
        if not df_acr.empty:
            df_acr_full = df_acr.merge(df_activos[['Cedula', 'Nombre']], on='Cedula', how='inner')
            st.dataframe(df_acr_full[["Nombre", "Acreedor", "Cuantia", "Clase"]], use_container_width=True, hide_index=True)

# --- 7. TAREAS ---
elif st.session_state.pagina_actual == 'Tareas':
    st.markdown("<h1>✅ Tablero de Misiones Privado</h1>", unsafe_allow_html=True)
    st.markdown("<div class='module-card module-card-blue'>", unsafe_allow_html=True)
    with st.form("form_tareas"):
        col1, col2 = st.columns([2, 1])
        with col1: desc_tar = st.text_input("Descripción de la Tarea a delegar")
        with col2: asignado = st.selectbox("Asignar al miembro:", df_usr["Alias"].tolist())
        if st.form_submit_button("ENVIAR MISIÓN"):
            if desc_tar:
                n_t = pd.DataFrame([{"ID_Tar": f"TAR-{hoy.strftime('%H%M%S')}", "Tarea": desc_tar, "Asignado": asignado, "Creador": st.session_state.alias_actual, "Estado": "Pendiente", "Fecha": hoy.strftime("%Y-%m-%d")}])
                pd.concat([df_tar, n_t], ignore_index=True).to_csv(ARCH_TAR, index=False); st.success("Misión asignada."); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    tareas_v = df_tar[(df_tar["Asignado"] == st.session_state.alias_actual) | (df_tar["Creador"] == st.session_state.alias_actual)]
    c_p, c_c = st.columns(2)
    with c_p:
        st.markdown("<h3 style='color: #F59E0B;'>⚠️ Misiones Pendientes</h3>", unsafe_allow_html=True)
        pend = tareas_v[tareas_v["Estado"] == "Pendiente"]
        if not pend.empty:
            for _, tar in pend.iloc[::-1].iterrows():
                st.markdown(f"<div style='border: 1px solid #F59E0B; padding: 15px; border-radius: 8px; margin-bottom: 10px; background: rgba(245, 158, 11, 0.05);'><b style='color:white;'>{tar['Tarea']}</b><br><span style='color:#A1A1AA; font-size:12px;'>Asignado a: {tar['Asignado']} | Por: {tar['Creador']}</span></div>", unsafe_allow_html=True)
                if st.button(f"✔️ Terminar", key=f"t_{tar['ID_Tar']}"):
                    df_tu = pd.read_csv(ARCH_TAR)
                    df_tu.loc[df_tu["ID_Tar"] == tar["ID_Tar"], "Estado"] = "Completada"
                    df_tu.to_csv(ARCH_TAR, index=False); st.rerun()
        else: st.info("No tienes misiones pendientes.")
    with c_c:
        st.markdown("<h3 style='color: #10B981;'>✅ Últimas Completadas</h3>", unsafe_allow_html=True)
        for _, tar in tareas_v[tareas_v["Estado"] == "Completada"].tail(5).iloc[::-1].iterrows():
            st.markdown(f"<div style='border: 1px solid #064E3B; background: rgba(16,185,129,0.05); padding: 15px; border-radius: 8px; margin-bottom: 10px;'><s style='color:#10B981; font-weight:bold;'>{tar['Tarea']}</s><br><span style='color:#71717A; font-size:12px;'>Asignada a {tar['Asignado']}. Finalizada.</span></div>", unsafe_allow_html=True)

# --- 8. MEMORIALES ---
elif st.session_state.pagina_actual == 'Memoriales':
    st.markdown("<h1>🤖 Dependiente Virtual (Memoriales Oficiales)</h1><p style='margin-bottom: 20px;'>Redacta documentos basados en la Ley 1564 y expórtalos en formato PDF oficial para el juzgado.</p>", unsafe_allow_html=True)
    st.markdown("<div class='module-card'>", unsafe_allow_html=True)
    
    with st.form("form_mem"):
        cli_m = st.selectbox("Seleccionar Cliente", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"] if not df_activos.empty else ["Sin Clientes"])
        tipo_mem = st.selectbox("Tipo de Actuación Legal / Petición:", [
            "Petición de Desembargo por Aceptación al Trámite", 
            "Solicitud de Terminación por Acuerdo de Pago", 
            "Solicitud de Copias Simples para Archivo", 
            "Sustitución de Poder a Nuevo Abogado",
            "Desistimiento de la Acción por Acuerdo"
        ])
        juzgado = st.text_input("Despacho Judicial o Centro de Conciliación Destino:")
        radicado = st.text_input("Número de Radicado (21 dígitos si aplica):")
        submit_mem = st.form_submit_button("📝 REDACTAR, DIAGRAMAR Y GENERAR PDF")
        
    if submit_mem:
        if cli_m != "Sin Clientes" and juzgado:
            nom_m, cc_m = cli_m.split(" - ")[1], cli_m.split(" - ")[0]
            titulo = f"SEÑOR JUEZ / CONCILIADOR\n{juzgado}\nE. S. D.\n\nREF: {tipo_mem}\nDEUDOR: {nom_m}\nRADICADO: {radicado}"
            
            if "Desembargo" in tipo_mem:
                cuerpo = f"Yo, en calidad de apoderado de {nom_m}, identificado(a) con cédula {cc_m}, respetuosamente me dirijo a su Despacho para solicitar que, en pleno cumplimiento de los efectos de la aceptación al trámite de insolvencia de persona natural no comerciante (Ley 1564 de 2012), se sirva DECRETAR EL INMEDIATO LEVANTAMIENTO DE LOS EMBARGOS que pesan sobre los bienes o salarios de mi poderdante y librar los oficios correspondientes a los pagadores o registradores pertinentes."
            elif "Terminación" in tipo_mem:
                cuerpo = f"En nombre y representación de {nom_m}, portador de la cédula de ciudadanía {cc_m}, acudo a su estrado judicial para solicitar de manera formal la TERMINACIÓN Y ARCHIVO DEFINITIVO del presente proceso. Lo anterior, toda vez que en el Centro de Conciliación se logró consolidar un Acuerdo de Pago de Insolvencia con las mayorías legales exigidas, el cual aporta paz y salvo respecto a las medidas ejecutivas vigentes."
            elif "Copias" in tipo_mem:
                cuerpo = f"Quien suscribe, actuando como mandatario judicial del señor(a) {nom_m} (C.C. {cc_m}), de manera comedida solicito a este Despacho autorizar y expedir copias simples, a mi costa, de la totalidad del expediente y/o las últimas actuaciones relevantes, con el único fin de ejercer el debido control, defensa y archivo interno de nuestra firma jurídica."
            elif "Sustitución" in tipo_mem:
                cuerpo = f"Yo, apoderado actual y debidamente reconocido dentro del proceso adelantado en contra de {nom_m}, identificado(a) con cédula {cc_m}, mediante el presente escrito manifiesto a su Despacho que SUSTITUYO el poder a mí conferido. Ruego reconocer personería jurídica al nuevo profesional del derecho para continuar con las etapas procesales correspondientes."
            else:
                cuerpo = f"Actuando como apoderado de {nom_m} (C.C. {cc_m}), por medio del presente memorial presento formal DESISTIMIENTO de las pretensiones de la demanda incoada. Solicitamos la terminación anormal del proceso sin condena en costas, en razón a que las partes han conciliado extraprocesalmente sus diferencias."

            cuerpo_completo = f"{cuerpo}\n\nAgradeciendo la atención y celeridad procesal otorgada a la presente petición.\n\nAtentamente,"
            nombre_archivo = f"Memorial_{cc_m}_{hoy.strftime('%H%M%S')}.pdf"
            generar_pdf_oficial(titulo, cuerpo_completo, nombre_archivo)
            st.session_state['memorial_pdf'] = nombre_archivo
        else:
            st.error("⚠ Es obligatorio seleccionar un cliente y escribir el nombre del Juzgado destino.")
            
    if 'memorial_pdf' in st.session_state and os.path.exists(st.session_state['memorial_pdf']):
        st.success("✅ Memorial redactado, diagramado y estructurado en PDF exitosamente.")
        with open(st.session_state['memorial_pdf'], "rb") as f:
            st.download_button("📥 DESCARGAR MEMORIAL OFICIAL (PDF) PARA FIRMA", f, file_name=st.session_state['memorial_pdf'], use_container_width=True)
            
    st.markdown("</div>", unsafe_allow_html=True)

# --- 9. AGENDA ---
elif st.session_state.pagina_actual == 'Agenda':
    st.markdown("<h1>📅 Agenda de Citas y Audiencias</h1>", unsafe_allow_html=True)
    st.markdown("<div class='module-card module-card-blue'>", unsafe_allow_html=True)
    with st.form("form_aud"):
        c1, c2 = st.columns(2)
        with c1:
            cl_aud = st.selectbox("Expediente / Cliente", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"] if not df_activos.empty else ["Sin Clientes"])
            motivo_aud = st.text_input("Motivo (Ej: Audiencia de Negociación)")
        with c2: 
            fecha_aud = st.date_input("Fecha Programada")
            hora_aud = st.time_input("Hora de la Cita")
        if st.form_submit_button("AGENDAR EVENTO"):
            if cl_aud != "Sin Clientes":
                cc_a, nom_a = cl_aud.split(" - ")[0], cl_aud.split(" - ")[1]
                n_au = pd.DataFrame([{"ID_Aud": f"AUD-{hoy.strftime('%H%M%S')}", "Cedula": cc_a, "Cliente": nom_a, "Fecha_Hora": f"{fecha_aud} {hora_aud}", "Motivo": motivo_aud}])
                pd.concat([df_aud, n_au], ignore_index=True).to_csv(ARCH_AUD, index=False); st.success("Agendado correctamente."); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)
    
    if not df_aud.empty:
        st.markdown("### 📌 Próximos Eventos Sincronizables")
        for _, r in df_aud.iloc[::-1].iterrows():
            col_d, col_b = st.columns([4,1])
            col_d.markdown(f"**{r['Cliente']}**: {r['Motivo']} - 🗓️ {r['Fecha_Hora']}")
            b_ics = generar_ics(r['Cliente'], r['Fecha_Hora'], r['Motivo'])
            if b_ics: col_b.download_button("📲 .ics", b_ics, file_name=f"Cita_{r['ID_Aud']}.ics", key=f"ics_{r['ID_Aud']}")
            st.divider()

# --- 10. FINANZAS ---
elif st.session_state.pagina_actual == 'Finanzas':
    if st.session_state.rol_actual != "Administrador (Jefa)": st.error("⛔ ACCESO DENEGADO")
    else:
        st.markdown("<h1>💰 Finanzas y Facturación Central</h1>", unsafe_allow_html=True)
        df_f_n = df_fin.merge(df_activos[['Cedula', 'Nombre', 'Telefono']], on='Cedula', how='inner') if not df_fin.empty else pd.DataFrame()
        if not df_f_n.empty:
            st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
            cf = st.selectbox("Seleccionar Libreta Financiera del Cliente:", df_f_n["Cedula"].astype(str) + " - " + df_f_n["Nombre"])
            if cf:
                cc_f = cf.split(" - ")[0]
                dat_f = df_f_n[df_f_n["Cedula"].astype(str) == cc_f].iloc[0]
                hon_t = limpiar_num(dat_f['Honorarios']); abo_t = limpiar_num(dat_f['Abonado']); saldo = hon_t - abo_t
                nom_f = dat_f['Nombre']; tel_f = str(dat_f['Telefono']).replace(" ", "")
                
                c1, c2, c3 = st.columns(3)
                c1.metric("Honorarios Pactados", f"$ {hon_t:,.0f}")
                c2.metric("Total Pagado (Abonado)", f"$ {abo_t:,.0f}")
                c3.metric("Saldo Adeudado a la Fecha", f"$ {saldo:,.0f}")
                
                st.divider()
                st.markdown("### 🧾 Registrar Pago y Emitir Soporte Oficial")
                with st.form("abono"):
                    n_abo = st.number_input("Monto Recibido en Caja Hoy ($)", min_value=0, step=50000)
                    if st.form_submit_button("REGISTRAR PAGO Y GENERAR FACTURA PDF"):
                        if n_abo > 0:
                            df_fu = pd.read_csv(ARCH_FIN)
                            n_abonado = abo_t + n_abo
                            n_saldo = hon_t - n_abonado
                            df_fu.loc[df_fu["Cedula"].astype(str) == cc_f, "Abonado"] = str(int(n_abonado))
                            df_fu.to_csv(ARCH_FIN, index=False)
                            st.session_state["pago_exitoso"] = True
                            st.session_state["datos_recibo"] = {"nom": nom_f, "cc": cc_f, "tel": tel_f, "abono": n_abo, "saldo": n_saldo}
                            st.rerun()
                        else: st.error("Ingresa un abono mayor a $0.")

                if st.session_state.get("pago_exitoso", False):
                    d_r = st.session_state["datos_recibo"]
                    st.success("✅ Pago registrado en la base de datos contable.")
                    n_pdf = f"Recibo_{d_r['cc']}_{hoy.strftime('%H%M%S')}.pdf"
                    if d_r['saldo'] <= 0:
                        tit_p = "CERTIFICADO OFICIAL DE PAZ Y SALVO"
                        cue_p = f"La firma jurídica certifica mediante el presente documento contable que el señor(a) {d_r['nom']}, identificado(a) con la cédula de ciudadanía No. {d_r['cc']}, ha cancelado el 100% de los honorarios pactados.\n\nPor consiguiente, el titular SE ENCUENTRA A PAZ Y SALVO por todo concepto financiero con esta firma de abogados."
                    else:
                        tit_p = f"RECIBO DE CAJA - RC-{hoy.strftime('%Y%m%d%H%M')}"
                        cue_p = f"Recibimos de: {d_r['nom']}\nIdentificación: C.C. {d_r['cc']}\nValor Recibido: $ {d_r['abono']:,.0f} COP\n\nResumen de la Cuenta Actualizada en el Sistema:\n- El pago ha sido ingresado a la cartera.\n- NUEVO SALDO RESTANTE: $ {d_r['saldo']:,.0f} COP."
                    
                    generar_pdf_oficial(tit_p, cue_p, n_pdf)
                    col_p1, col_p2 = st.columns(2)
                    with col_p1:
                        with open(n_pdf, "rb") as f:
                            st.download_button("📥 DESCARGAR RECIBO/PAZ Y SALVO (PDF)", f, file_name=n_pdf, use_container_width=True)
                    with col_p2:
                        if d_r['tel'] and d_r['tel'] != "nan":
                            if d_r['saldo'] <= 0: msg_w = f"¡Excelente noticia {d_r['nom']}! Hemos registrado tu último pago y tu cuenta quedó en $0. Eres oficialmente libre de deudas con nosotros."
                            else: msg_w = f"Hola {d_r['nom']}, confirmamos la recepción de tu pago por ${d_r['abono']:,.0f}. Tu nuevo saldo pendiente es de ${d_r['saldo']:,.0f}."
                            link_wa = f"https://wa.me/57{d_r['tel']}?text={msg_w.replace(' ', '%20')}"
                            st.markdown(f"<a href='{link_wa}' target='_blank'><button style='background:#10B981; color:white; border:none; padding:10px 20px; border-radius:6px; font-weight:bold; cursor:pointer; width:100%;'>📱 AVISARLE AL CLIENTE POR WHATSAPP</button></a>", unsafe_allow_html=True)
                        else: st.warning("El cliente no tiene celular registrado.")
                    if st.button("Finalizar Proceso de Pago"): st.session_state.pop("pago_exitoso"); st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

# --- 11. USUARIOS ---
elif st.session_state.pagina_actual == 'Usuarios':
    if st.session_state.rol_actual != "Administrador (Jefa)": st.error("⛔ ACCESO DENEGADO")
    else:
        st.markdown("<h1>Administración de Accesos y Seguridad</h1><p style='margin-bottom: 20px;'>Control total de credenciales, cambio de claves y eliminación de personal.</p>", unsafe_allow_html=True)
        st.markdown("<div class='module-card module-card-gold'>", unsafe_allow_html=True)
        st.markdown("<h3>➕ Crear Nueva Credencial</h3>", unsafe_allow_html=True)
        with st.form("form_nuevo_usr"):
            c1, c2, c3 = st.columns(3)
            with c1: n_user = st.text_input("Usuario (Login de Red)")
            with c2: n_pass = st.text_input("Contraseña de Acceso", type="password")
            with c3: n_rol = st.selectbox("Rol Asignado", ["Abogado / Operativo", "Auxiliar Jurídico", "Administrador (Jefa)"])
            if st.form_submit_button("REGISTRAR NUEVO USUARIO ENCRIPTADO"):
                if not n_user or not n_pass: st.error("Campos obligatorios incompletos.")
                elif n_user in df_usr["Usuario"].astype(str).values: st.error("El usuario ya existe en la red.")
                else:
                    nu = pd.DataFrame([{"Usuario": n_user, "Password": n_pass, "Rol": n_rol, "Creador": st.session_state.usuario_actual, "Alias": n_user, "Avatar_Path": "", "Session_Token": ""}])
                    pd.concat([df_usr, nu], ignore_index=True).to_csv(ARCH_USR, index=False)
                    st.success(f"Credenciales creadas exitosamente para {n_user}."); st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
        df_u_actual = pd.read_csv(ARCH_USR)
        st.dataframe(df_u_actual[["Usuario", "Alias", "Rol", "Creador"]], use_container_width=True, hide_index=True)

# --- 12. SISTEMA ---
elif st.session_state.pagina_actual == 'Sistema':
    if st.session_state.rol_actual != "Administrador (Jefa)": st.error("⛔ DENEGADO")
    else:
        st.markdown("<h1>Exportación General y Auditoría</h1>", unsafe_allow_html=True)
        st.markdown("<div class='module-card'>", unsafe_allow_html=True)
        st.markdown("<h3>📊 Data Dump (Copia de Seguridad Excel)</h3>", unsafe_allow_html=True)
        df_dump = df_cli.merge(df_fin, on="Cedula", how="left")
        st.download_button("📥 DESCARGAR BASE DE DATOS GLOBAL DE LA FIRMA (CSV)", df_dump.to_csv(index=False).encode('utf-8'), f"Respaldo_Firma_{hoy.strftime('%Y%m%d')}.csv", "text/csv")
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<div class='module-card' style='border-top: 3px solid #EF4444;'>", unsafe_allow_html=True)
        st.markdown("<h3 style='color: #EF4444 !important;'>🗑️️ Papelera de Reciclaje</h3><p>Eliminar o restaurar expedientes.</p>", unsafe_allow_html=True)
        if not df_activos.empty:
            cb = st.selectbox("Mover expediente a la papelera:", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"])
            if st.button("ENVIAR A PAPELERA"):
                c_b = cb.split(" - ")[0]
                df_bd = pd.read_csv(ARCH_CLI)
                df_bd.loc[df_bd["Cedula"].astype(str) == c_b, "Estado"] = "Borrado"
                df_bd.loc[df_bd["Cedula"].astype(str) == c_b, "F_Borrado"] = hoy.strftime("%Y-%m-%d")
                df_bd.to_csv(ARCH_CLI, index=False); st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
