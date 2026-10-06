import streamlit as st
import pandas as pd
import os
import shutil
import base64
import uuid
from datetime import datetime, timedelta
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
st.set_page_config(page_title="Insolvencia OS | V12 Platinum", page_icon="⚖", layout="wide", initial_sidebar_state="expanded")

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
# 2. MOTOR CSS: MIDNIGHT SLATE & CHAMPAGNE
# ==========================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Playfair+Display:wght@600;800&display=swap');
    #MainMenu, footer {visibility: hidden;}
    
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #0B1120 !important; color: #F8FAFC !important; }
    h1, h2, h3 { font-family: 'Playfair Display', serif !important; color: #C8A951 !important; letter-spacing: -0.5px; }
    
    [data-testid="stAppViewContainer"] { background-color: #0F172A !important; }
    [data-testid="stSidebar"] { background-color: #0B1120 !important; border-right: 1px solid #1E293B !important; }
    
    /* Inputs y Textareas */
    input, textarea, select, div[data-baseweb="select"] > div, div[data-baseweb="input"] > div { 
        background-color: #1E293B !important; color: #F8FAFC !important; -webkit-text-fill-color: #F8FAFC !important; 
        border: 1px solid #334155 !important; border-radius: 10px !important; transition: all 0.3s;
    }
    input:focus, textarea:focus { border-color: #C8A951 !important; box-shadow: 0 0 12px rgba(200,169,81,0.2) !important; }
    
    /* Tarjetas Soft UI */
    .module-card { background: linear-gradient(145deg, #1E293B 0%, #0F172A 100%); border: 1px solid #334155; border-radius: 20px; padding: 25px; margin-bottom: 25px; box-shadow: 0 10px 40px rgba(0,0,0,0.3); }
    .module-card-gold { border-top: 4px solid #C8A951; }
    .module-card-blue { border-top: 4px solid #3B82F6; }
    .module-card-green { border-top: 4px solid #10B981; }
    
    /* Botones Premium */
    div.stButton > button:first-child { background: #1E293B !important; color: #C8A951 !important; border: 1px solid #C8A951 !important; border-radius: 8px !important; padding: 12px 24px !important; font-weight: 600 !important; font-size: 13px !important; transition: all 0.3s ease !important; width: 100%; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    div.stButton > button:first-child:hover { background: #C8A951 !important; color: #0F172A !important; box-shadow: 0 8px 20px rgba(200,169,81,0.3) !important; transform: translateY(-2px); }
    
    /* Pestañas (Tabs) Estilizadas */
    .stTabs [data-baseweb="tab-list"] { gap: 8px; background-color: #1E293B; padding: 10px; border-radius: 12px; }
    .stTabs [data-baseweb="tab"] { background-color: transparent !important; border-radius: 8px !important; color: #94A3B8 !important; border: none !important; }
    .stTabs [aria-selected="true"] { background-color: #334155 !important; color: #C8A951 !important; font-weight: bold; }
    
    .timeline { display: flex; justify-content: space-between; align-items: center; margin: 20px 0; position: relative; }
    .timeline::before { content: ''; position: absolute; top: 50%; left: 0; right: 0; height: 3px; background: #334155; z-index: 1; border-radius: 2px;}
    .step { position: relative; z-index: 2; background: #0F172A; padding: 10px 20px; border-radius: 25px; border: 2px solid #334155; color: #94A3B8; font-weight: 600; font-size: 12px; display: flex; align-items: center; text-transform: uppercase; letter-spacing: 0.5px;}
    .step.active { border-color: #C8A951; color: #C8A951; box-shadow: 0 0 20px rgba(200,169,81,0.2); background: #1E293B; }
    .step.completed { border-color: #10B981; color: #10B981; background: #064E3B; }
    
    .alerta-roja { color: #EF4444 !important; font-weight: bold; background: rgba(239, 68, 68, 0.1); padding: 4px 8px; border-radius: 6px; }
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
    df_l = pd.read_csv(ARCH_LOG)
    usr = st.session_state.get('alias_actual', 'Sistema')
    nuevo = pd.DataFrame([{"Timestamp": hoy.strftime("%Y-%m-%d %H:%M:%S"), "Usuario": usr, "Modulo": modulo, "Accion": accion}])
    pd.concat([df_l, nuevo], ignore_index=True).to_csv(ARCH_LOG, index=False)

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
    except Exception as e: 
        return False

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
        ics_content = f"BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//Insolvencia OS//Agenda//ES\nBEGIN:VEVENT\nSUMMARY:⚖️ {motivo} - {cliente}\nDTSTART:{dt_start.strftime(formato)}\nDTEND:{dt_end.strftime(formato)}\nDESCRIPTION:Audiencia/Cita programada desde el sistema de gestión.\nEND:VEVENT\nEND:VCALENDAR"
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
            col.markdown(f"<p style='color:#C8A951; margin-bottom:2px; font-weight:bold;'>{c.replace('_', ' ').title()}</p>", unsafe_allow_html=True)
            if archivos:
                for a in archivos:
                    ruta = os.path.join(rb, c, a)
                    if os.path.isfile(ruta):
                        with open(ruta, "rb") as f:
                            col.download_button(f"📄 {a[:25]}...", f, file_name=a, key=f"dl_{cc}_{c}_{a}")
            else:
                col.markdown("<span style='color:#94A3B8; font-size:12px;'><i>Carpeta vacía</i></span>", unsafe_allow_html=True)
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
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        if st.session_state.kicked:
            st.error(f"🚨 ALERTA: {st.session_state.kicked_reason}"); st.session_state.kicked = False
        st.markdown("""<div class="module-card module-card-gold" style="text-align: center; padding: 40px;"><span style="font-size: 55px; display:block; margin-bottom:10px;">⚖️</span><h2 style="font-size:32px;">ACCESO LEGALTECH</h2><p style="color: #94A3B8; font-size: 14px; margin-bottom: 30px;">Plataforma Cifrada V12</p>""", unsafe_allow_html=True)
        with st.form("login_form"):
            user_input = st.text_input("👤 Identificador de Red")
            pass_input = st.text_input("🔑 Clave de Acceso", type="password")
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
# 5. SIDEBAR ESTILIZADA Y CONDENSADA
# ==========================================
with st.sidebar:
    avatar_p = st.session_state.get('avatar_path', '')
    if pd.notna(avatar_p) and str(avatar_p).strip() != "" and os.path.exists(avatar_p):
        with open(avatar_p, "rb") as f: encoded_img = base64.b64encode(f.read()).decode()
        st.markdown(f"""<div style="text-align:center; margin-bottom: 10px;"><img src="data:image/png;base64,{encoded_img}" style="width: 80px; height: 80px; border-radius: 50%; object-fit: cover; border: 2px solid #C8A951; box-shadow: 0 4px 15px rgba(200,169,81,0.2);"></div>""", unsafe_allow_html=True)
    else:
        st.markdown("""<div style="width: 80px; height: 80px; background: #1E293B; border-radius: 50%; margin: 0 auto 10px auto; display:flex; align-items:center; justify-content:center; border: 2px solid #C8A951;"><span style="font-size: 35px;">⚖️</span></div>""", unsafe_allow_html=True)

    st.markdown(f"""<div style="text-align:center; padding-bottom: 10px; margin-bottom: 15px;"><h3 style="color: #F8FAFC !important; font-size: 18px; margin:0;">{st.session_state.alias_actual}</h3><span style="color: #C8A951; font-size: 12px; font-weight: 600;">{st.session_state.rol_actual}</span></div>""", unsafe_allow_html=True)
    
    with st.expander("⚙️ Configurar Perfil"):
        with st.form("form_perfil"):
            n_alias = st.text_input("Alias", value=st.session_state.alias_actual)
            foto_sub = st.file_uploader("Foto", type=["jpg", "png", "jpeg"])
            if st.form_submit_button("Guardar"):
                df_u_up = pd.read_csv(ARCH_USR)
                p_guardado = st.session_state.avatar_path
                if foto_sub:
                    p_guardado = os.path.join(CARP_AVATAR, f"{st.session_state.usuario_actual}_{foto_sub.name}")
                    with open(p_guardado, "wb") as f: f.write(foto_sub.getbuffer())
                df_u_up.loc[df_u_up["Usuario"].astype(str) == st.session_state.usuario_actual, "Alias"] = n_alias
                df_u_up.loc[df_u_up["Usuario"].astype(str) == st.session_state.usuario_actual, "Avatar_Path"] = p_guardado
                df_u_up.to_csv(ARCH_USR, index=False)
                st.session_state.alias_actual = n_alias; st.session_state.avatar_path = p_guardado
                st.success("Actualizado"); st.rerun()

    st.markdown("<hr style='border-color: #1E293B; margin-top: 5px;'>", unsafe_allow_html=True)
    
    st.markdown("<p style='color:#64748B; font-size:10px; font-weight:bold; letter-spacing:1.5px; margin-bottom:8px;'>HOME & WORKSPACE</p>", unsafe_allow_html=True)
    if st.button("📊 Portal Inteligente", key="b_dash"): cambiar_pagina("Dashboard")
    if st.button("📝 Apertura de Casos", key="b_nuev"): cambiar_pagina("Nuevo")
    if st.button("📅 Agenda Sincronizada", key="b_agen"): cambiar_pagina("Agenda")
    if st.button("✅ Gestor de Tareas", key="b_tar"): cambiar_pagina("Tareas")
    
    st.markdown("<br><p style='color:#64748B; font-size:10px; font-weight:bold; letter-spacing:1.5px; margin-bottom:8px;'>LEGAL OPERATIONS</p>", unsafe_allow_html=True)
    if st.button("🔄 Expediente 360°", key="b_360"): cambiar_pagina("Expediente360")
    if st.button("🚦 Control Vencimientos", key="b_ven"): cambiar_pagina("Vencimientos")
    if st.button("🤖 Auto-Memoriales", key="b_mem"): cambiar_pagina("Memoriales")
    
    st.markdown("<br><p style='color:#64748B; font-size:10px; font-weight:bold; letter-spacing:1.5px; margin-bottom:8px;'>FINANCE & ADMIN</p>", unsafe_allow_html=True)
    if st.button("📋 Pasivos Registrados", key="b_acr"): cambiar_pagina("Acreedores")
    if st.session_state.rol_actual == "Administrador (Jefa)":
        if st.button("💰 Flujo de Caja", key="b_fin"): cambiar_pagina("Finanzas")
        if st.button("👥 Credenciales", key="b_usr"): cambiar_pagina("Usuarios")
        if st.button("🛡️ Sistema y Respaldo", key="b_sis"): cambiar_pagina("Sistema")
        
    st.markdown("<hr style='border-color: #1E293B;'>", unsafe_allow_html=True)
    if st.button("🚪 Cerrar Sesión Segura"): st.session_state.autenticado = False; st.rerun()

# ==========================================
# 6. MÓDULOS DE LA APLICACIÓN
# ==========================================

# --- 1. PORTAL EJECUTIVO ---
if st.session_state.pagina_actual == 'Dashboard':
    st.markdown("<h1>Centro de Operaciones V12</h1><p style='margin-bottom: 25px; color:#94A3B8;'>Monitoreo global de expedientes y proyecciones financieras.</p>", unsafe_allow_html=True)
    total_cartera = sum([(limpiar_num(r["Honorarios"]) - limpiar_num(r["Abonado"])) for _, r in df_fin[df_fin["Cedula"].astype(str).isin(df_activos["Cedula"].astype(str))].iterrows() if limpiar_num(r["Honorarios"]) > limpiar_num(r["Abonado"])])
    
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"<div class='module-card module-card-blue'><h4>Casos Activos</h4><h2 style='font-size:38px; margin:0;'>{len(df_activos)}</h2></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='module-card module-card-gold'><h4>En Ensamblaje</h4><h2 style='font-size:38px; margin:0;'>{len(df_activos[df_activos['Senal'].isin([1,2,3])])}</h2></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='module-card module-card-green'><h4>Radicados Oficiales</h4><h2 style='font-size:38px; margin:0;'>{len(df_activos[df_activos['Senal'] >= 4])}</h2></div>", unsafe_allow_html=True)
    if st.session_state.rol_actual == "Administrador (Jefa)": c4.markdown(f"<div class='module-card module-card-gold'><h4>Cartera Flotante</h4><h2 style='font-size:28px; margin:0; color:#C8A951;'>$ {total_cartera:,.0f}</h2></div>", unsafe_allow_html=True)

    col_izq, col_der = st.columns([1.6, 1])
    with col_izq:
        st.markdown("<div class='module-card'><h3>🔍 Búsqueda Instantánea</h3>", unsafe_allow_html=True)
        busq = st.text_input("Localizar cliente por nombre o cédula:")
        if busq:
            res = df_activos[df_activos['Nombre'].str.contains(busq, case=False, na=False) | df_activos['Cedula'].astype(str).str.contains(busq, na=False)]
            if not res.empty: st.dataframe(res[["Cedula", "Nombre", "Fuerza", "Telefono"]], use_container_width=True, hide_index=True)
            else: st.info("No hay coincidencias en la base de datos activa.")
        st.markdown("</div>", unsafe_allow_html=True)
        
    with col_der:
        st.markdown("<div class='module-card module-card-gold'><h3>🧮 Test de Viabilidad (1564)</h3>", unsafe_allow_html=True)
        ingreso = st.number_input("Ingresos Mensuales", min_value=0, step=100000, key="cal_ing")
        gastos = st.number_input("Gastos Subsistencia", min_value=0, step=100000, key="cal_gas")
        deuda = st.number_input("Deuda Total Aprox.", min_value=0, step=500000, key="cal_deu")
        if st.button("Calcular Proyección"):
            disp = ingreso - gastos
            if disp > 0 and deuda > 0:
                meses = deuda / disp
                st.success(f"Disponibilidad Legal: **${disp:,.0f}**")
                if meses <= 60: st.info(f"Capital cubierto en aprox. **{meses:.0f} meses**. Idóneo para Acuerdo de Pago.")
                else: st.warning(f"Tomaría {meses:.0f} meses. Perfil sugerido para Liquidación Patrimonial.")
            else: st.error("No hay margen disponible o falta ingresar la deuda.")
        st.markdown("</div>", unsafe_allow_html=True)

# --- 2. APERTURA ---
elif st.session_state.pagina_actual == 'Nuevo':
    st.markdown("<h1>Apertura de Expediente</h1>", unsafe_allow_html=True)
    st.markdown("<div class='module-card'>", unsafe_allow_html=True)
    with st.form("nuevo_exp"):
        c1, c2, c3 = st.columns(3)
        with c1: cc = st.text_input("Cédula de Ciudadanía")
        with c2: nom = st.text_input("Nombre Completo")
        with c3: fza = st.selectbox("Perfil / Ocupación", ["Civil", "Policía Nacional", "Ejército Nacional", "Armada Nacional", "Fuerza Aérea", "Retirado / Pensionado"])
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

# --- 3. EXPEDIENTE 360° (Consolidación de Contratos, WhatsApp y Reportes) ---
elif st.session_state.pagina_actual == 'Expediente360':
    st.markdown("<h1>Expediente Digital 360°</h1><p style='color:#94A3B8; margin-bottom:20px;'>Todo el universo del cliente en una sola pantalla unificada.</p>", unsafe_allow_html=True)
    if df_activos.empty: st.warning("No hay clientes activos.")
    else:
        st.markdown("<div class='module-card'>", unsafe_allow_html=True)
        cli_sel = st.selectbox("🔍 Cargar Expediente:", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"])
        cc_s, nom_s = cli_sel.split(" - ")[0], cli_sel.split(" - ")[1]
        data_c = df_activos[df_activos["Cedula"].astype(str) == cc_s].iloc[0]
        senal = data_c["Senal"]
        tel_c = str(data_c.get("Telefono", ""))
        rb = estructurar_carpetas(cc_s, nom_s)
        
        # Timeline Visual Estilizado
        c1, c2, c3, c4 = ["completed" if senal >= i else ("active" if senal == i-1 else "") for i in range(1, 5)]
        st.markdown(f'<div class="timeline" style="margin-top:10px; margin-bottom:30px;"><div class="step {c1}">1. Docs</div><div class="step {c2}">2. Motor</div><div class="step {c3}">3. Firmas</div><div class="step {c4}">4. Radicado</div></div>', unsafe_allow_html=True)
        
        # Bóveda Central (Expander)
        mostrar_boveda(cc_s, nom_s)
        
        # PESTAÑAS (TABS) V12
        tab_resumen, tab_docs, tab_acts, tab_wa = st.tabs(["📊 Resumen & Reporte", "📁 Documentos & Motor", "⚖️ Actuaciones", "📱 Central WhatsApp"])
        
        with tab_resumen:
            col_res1, col_res2 = st.columns(2)
            with col_res1:
                st.markdown("### Perfil Legal")
                st.write(f"**Identificación:** {cc_s} | **Perfil:** {data_c['Fuerza']}")
                st.write(f"**Contacto:** {tel_c} | **Email:** {data_c.get('Email', 'N/A')}")
                st.write(f"**Pasivo Declarado:** ${data_c.get('Deuda_Est', '0')}")
                
                fin_cli = df_fin[df_fin["Cedula"].astype(str) == cc_s]
                if not fin_cli.empty:
                    d_fin = fin_cli.iloc[0]
                    h = limpiar_num(d_fin['Honorarios']); a = limpiar_num(d_fin['Abonado'])
                    st.markdown(f"<div style='background:#1E293B; padding:15px; border-radius:10px; border: 1px solid #334155; margin-top:15px;'><b>Estado Financiero:</b><br>Honorarios: ${h:,.0f}<br>Abonado: ${a:,.0f}<br><span style='color:#C8A951;'>Saldo Actual: ${h-a:,.0f}</span></div>", unsafe_allow_html=True)
            
            with col_res2:
                st.markdown("### Reporte de Avance Oficial (PDF)")
                st.info("Exporta un documento formal con el resumen del caso para enviarlo al cliente y brindarle tranquilidad.")
                if st.button("📄 Generar Reporte de Estado para Cliente"):
                    rep_name = f"Reporte_Avance_{cc_s}_{hoy.strftime('%H%M%S')}.pdf"
                    fase_txt = ["Fase 1: Recolección Documental", "Fase 2: Redacción de Contratos", "Fase 3: Pendiente Firmas del Cliente", "Fase 4: Radicación y Trámite Legal"][min(senal, 3)]
                    acts = df_act[df_act["Cedula"].astype(str) == cc_s]
                    acts_txt = "Aún no hay actuaciones procesales registradas en los juzgados." if acts.empty else f"Última novedad procesal: {acts.iloc[-1]['Tipo']} en {acts.iloc[-1]['Juzgado']} (Fecha: {acts.iloc[-1]['Fecha']})."
                    
                    texto_rep = f"REPORTE OFICIAL DE ESTADO PROCESAL\n\nFecha de emisión: {hoy.strftime('%Y-%m-%d')}\nTitular del Proceso: {nom_s}\nDocumento de Identidad: {cc_s}\n\nFase Actual del Proceso:\n{fase_txt}\n\nREPORTE JURÍDICO:\n{acts_txt}\n\nMensaje del Equipo Jurídico:\nSu trámite de insolvencia avanza conforme a los tiempos y estrategias estipuladas. Nuestro equipo sigue monitoreando cada etapa para garantizar la protección de su patrimonio. Agradecemos su confianza en nuestra firma."
                    
                    generar_pdf_oficial(f"ESTADO DE PROCESO LEGAL", texto_rep, rep_name)
                    st.session_state['rep_pdf'] = rep_name
                    st.rerun()
                
                if 'rep_pdf' in st.session_state and os.path.exists(st.session_state['rep_pdf']):
                    with open(st.session_state['rep_pdf'], "rb") as f:
                        st.download_button("📥 DESCARGAR REPORTE PDF GENERADO", f, file_name=st.session_state['rep_pdf'], use_container_width=True)

        with tab_docs:
            r1, r2 = os.path.join(rb, "01_Docs_Viabilidad"), os.path.join(rb, "02_Contratos_Firmas")
            if senal == 0:
                arch = os.listdir(r1)
                st.markdown(f"### Fase 1: Recolección Documental ({len(arch)}/8)")
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
                st.markdown("### Fase 2: Motor de Ensamblaje")
                with st.form("motor"):
                    ca, cb = st.columns(2)
                    with ca: abo = st.selectbox("Abogado Titular de la Firma", ABOGADOS)
                    with cb: h_n = st.text_input("Honorarios Totales ($)"); h_l = st.text_input("Honorarios (En Letras)")
                    cc_col, cd, ce = st.columns(3)
                    with cc_col: cuo_n = st.text_input("Valor Cuota ($)"); cuo_l = st.text_input("Valor Cuota (Letras)")
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
                st.markdown("### Fase 3: Recolección de Firmas")
                for dg in os.listdir(r2):
                    ruta = os.path.join(r2, dg)
                    if os.path.isfile(ruta):
                        with open(ruta, "rb") as f: st.download_button(f"📥 Imprimir Documento: {dg}", f, file_name=dg)
                uf = st.file_uploader("📥 Subir Paquete de Contratos Firmados (PDF)", type=["pdf"])
                if uf:
                    with open(os.path.join(r2, f"Firmados_{uf.name}"), "wb") as f: f.write(uf.getbuffer())
                    df_cli.loc[df_cli["Cedula"].astype(str) == str(cc_s), "Senal"] = 3
                    df_cli.to_csv(ARCH_CLI, index=False); st.rerun()
            
            elif senal >= 3: st.success("✨ El expediente documental está completado y listo para operaciones procesales.")

        with tab_acts:
            st.markdown("### ⚖️ Registro y Log Judicial")
            r4 = os.path.join(rb, "04_Log_Judicial")
            with st.form("f_act_360"):
                ca1, ca2, ca3 = st.columns(3)
                with ca1: tipo_a = st.selectbox("Tipo de Novedad", ["Auto Admisorio", "Auto Inadmisorio", "Sentencia", "Oficio", "Memorial Radicado"])
                with ca2: juzgado = st.text_input("Despacho Judicial")
                with ca3: rad_jud = st.text_input("Radicado")
                anota = st.text_area("Anotación / Novedad")
                doc_up = st.file_uploader("Adjuntar PDF de la actuación (Opcional)")
                if st.form_submit_button("REGISTRAR ACTUACIÓN EN SISTEMA"):
                    if doc_up:
                        with open(os.path.join(r4, f"{tipo_a}_{doc_up.name}"), "wb") as f: f.write(doc_up.getbuffer())
                    n_act = pd.DataFrame([{"ID_Act": f"ACT-{hoy.strftime('%H%M%S')}", "Cedula": cc_s, "Fecha": hoy.strftime("%Y-%m-%d"), "Tipo": tipo_a, "Juzgado": juzgado, "Radicado": rad_jud, "Anotacion": anota}])
                    pd.concat([df_act, n_act], ignore_index=True).to_csv(ARCH_ACT, index=False); st.success("Registrado con éxito."); st.rerun()
            
            acts_c = df_act[df_act["Cedula"].astype(str) == cc_s]
            if not acts_c.empty: st.dataframe(acts_c[["Fecha", "Tipo", "Juzgado", "Radicado", "Anotacion"]], use_container_width=True, hide_index=True)

        with tab_wa:
            st.markdown("### 📱 Central Automática de WhatsApp (Tus 30 Plantillas)")
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
            st.text_area("Vista previa del mensaje (Puedes editarlo antes de enviar):", value=msg_texto, height=120)
            
            if tel_c and tel_c != "nan" and tel_c != "":
                link_wa = f"https://wa.me/57{tel_c.replace(' ', '')}?text={msg_texto.replace(' ', '%20')}"
                st.markdown(f"<a href='{link_wa}' target='_blank'><button style='background:#10B981; color:white; border:none; padding:12px 20px; border-radius:6px; font-weight:bold; cursor:pointer; width:100%; margin-bottom: 15px;'>🚀 ENVIAR WHATSAPP DIRECTO</button></a>", unsafe_allow_html=True)
            else: st.warning("⚠ No hay celular registrado para este cliente.")
        
        st.markdown("</div>", unsafe_allow_html=True)

# --- 5. VENCIMIENTOS (SEMAFORO) ---
elif st.session_state.pagina_actual == 'Vencimientos':
    st.markdown("<h1>🚦 Control de Vencimientos Automático</h1><p style='margin-bottom: 20px; color:#94A3B8;'>Evita caducidades legales con el semáforo inteligente.</p>", unsafe_allow_html=True)
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
                pd.concat([df_ven, n_v], ignore_index=True).to_csv(ARCH_VEN, index=False); st.success("Plazo registrado."); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    if not df_ven.empty:
        for _, rv in df_ven[df_ven["Estado"] == "Activo"].iloc[::-1].iterrows():
            dias = (datetime.strptime(rv["Fecha_Limite"], "%Y-%m-%d").date() - hoy.date()).days
            if dias < 0:
                cb = "#334155"; estado_txt = f"<span style='color:#94A3B8; font-weight:bold;'>VENCIDO HACE {abs(dias)} DÍAS</span>"
            elif dias <= 2:
                cb = "#EF4444"; estado_txt = f"<span class='alerta-roja'>¡PELIGRO INMINENTE! VENCE EN {dias} DÍAS</span>"
            elif dias <= 5:
                cb = "#F59E0B"; estado_txt = f"<span style='color:#F59E0B; font-weight:bold;'>ATENCIÓN: Faltan {dias} días</span>"
            else:
                cb = "#10B981"; estado_txt = f"<span style='color:#10B981; font-weight:bold;'>En término adecuado ({dias} días)</span>"
                
            st.markdown(f"<div style='background: #1E293B; border: 1px solid #334155; border-left: 5px solid {cb}; padding: 15px; border-radius: 8px; margin-bottom: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'><b style='color:white;'>{rv['Cliente']}</b> - {rv['Asunto']}<br>{estado_txt} (Fecha Límite: {rv['Fecha_Limite']})</div>", unsafe_allow_html=True)
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
                st.markdown(f"<div style='border: 1px solid #F59E0B; padding: 15px; border-radius: 8px; margin-bottom: 10px; background: rgba(245, 158, 11, 0.05);'><b style='color:white;'>{tar['Tarea']}</b><br><span style='color:#94A3B8; font-size:12px;'>Asignado a: {tar['Asignado']} | Por: {tar['Creador']}</span></div>", unsafe_allow_html=True)
                if st.button(f"✔️ Terminar", key=f"t_{tar['ID_Tar']}"):
                    df_tu = pd.read_csv(ARCH_TAR)
                    df_tu.loc[df_tu["ID_Tar"] == tar["ID_Tar"], "Estado"] = "Completada"
                    df_tu.to_csv(ARCH_TAR, index=False); st.rerun()
        else: st.info("No tienes misiones pendientes.")
    with c_c:
        st.markdown("<h3 style='color: #10B981;'>✅ Últimas Completadas</h3>", unsafe_allow_html=True)
        for _, tar in tareas_v[tareas_v["Estado"] == "Completada"].tail(5).iloc[::-1].iterrows():
            st.markdown(f"<div style='border: 1px solid #064E3B; background: rgba(16,185,129,0.05); padding: 15px; border-radius: 8px; margin-bottom: 10px;'><s style='color:#10B981; font-weight:bold;'>{tar['Tarea']}</s><br><span style='color:#94A3B8; font-size:12px;'>Asignada a {tar['Asignado']}. Finalizada.</span></div>", unsafe_allow_html=True)

# --- 8. DEPENDIENTE VIRTUAL (MEMORIALES PRO Y PDF) ---
elif st.session_state.pagina_actual == 'Memoriales':
    st.markdown("<h1>🤖 Dependiente Virtual (Memoriales Oficiales)</h1><p style='margin-bottom: 20px; color:#94A3B8;'>Redacta documentos basados en la Ley 1564 y expórtalos en formato PDF oficial para el juzgado.</p>", unsafe_allow_html=True)
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
        
        # SOLO el botón de submit va dentro del form (EL FIX SE MANTIENE INTACTO)
        submit_mem = st.form_submit_button("📝 REDACTAR, DIAGRAMAR Y GENERAR PDF")
        
    # FUERA DEL FORM: Aquí procesamos la información y mostramos la descarga para evitar el error
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
            
            # Generar PDF Oficial
            nombre_archivo = f"Memorial_{cc_m}_{hoy.strftime('%H%M%S')}.pdf"
            generar_pdf_oficial(titulo, cuerpo_completo, nombre_archivo)
            
            # Guardamos la ruta en la sesión
            st.session_state['memorial_pdf'] = nombre_archivo
        else:
            st.error("⚠ Es obligatorio seleccionar un cliente y escribir el nombre del Juzgado destino.")
            
    # Mostrar el botón de descarga
    if 'memorial_pdf' in st.session_state and os.path.exists(st.session_state['memorial_pdf']):
        st.success("✅ Memorial redactado, diagramado y estructurado en PDF exitosamente.")
        with open(st.session_state['memorial_pdf'], "rb") as f:
            st.download_button("📥 DESCARGAR MEMORIAL OFICIAL (PDF) PARA FIRMA", f, file_name=st.session_state['memorial_pdf'], use_container_width=True)
            
    st.markdown("</div>", unsafe_allow_html=True)

# --- 9. AGENDA (Y GENERADOR ICS V12) ---
elif st.session_state.pagina_actual == 'Agenda':
    st.markdown("<h1>📅 Agenda de Citas y Audiencias</h1>", unsafe_allow_html=True)
    st.markdown("<div class='module-card'>", unsafe_allow_html=True)
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
        st.markdown("### 📌 Próximos Eventos y Sincronización")
        for _, r in df_aud.iloc[::-1].iterrows():
            col_d, col_b = st.columns([4,1])
            col_d.markdown(f"<div style='background:#1E293B; padding:12px; border-radius:8px; border:1px solid #334155;'><b>{r['Cliente']}</b>: {r['Motivo']} - 🗓️ {r['Fecha_Hora']}</div>", unsafe_allow_html=True)
            b_ics = generar_ics(r['Cliente'], r['Fecha_Hora'], r['Motivo'])
            if b_ics: 
                col_b.download_button("📲 Añadir (.ics)", b_ics, file_name=f"Cita_{r['ID_Aud']}.ics", key=f"ics_{r['ID_Aud']}")

# --- 10. FINANZAS (FACTURACIÓN, GRÁFICO, RECIBOS PDF Y WHATSAPP) ---
elif st.session_state.pagina_actual == 'Finanzas':
    if st.session_state.rol_actual != "Administrador (Jefa)": st.error("⛔ ACCESO DENEGADO")
    else:
        st.markdown("<h1>💰 Finanzas y Facturación Central</h1>", unsafe_allow_html=True)
        
        # Gráfico V12 de Rendimiento Financiero
        st.markdown("### 📊 Rendimiento Global de Cartera")
        dat_plot = pd.DataFrame({
            "Métricas": ["Total Facturado", "Total Recaudado"],
            "Montos": [
                sum([limpiar_num(x) for x in df_fin["Honorarios"]]), 
                sum([limpiar_num(x) for x in df_fin["Abonado"]])
            ]
        }).set_index("Métricas")
        st.bar_chart(dat_plot, color="#C8A951", height=250)

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

                # Procesamiento post-pago y generación PDF
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
                            if d_r['saldo'] <= 0: msg_w = f"¡Excelente noticia {d_r['nom']}! Hemos registrado tu último pago y tu cuenta quedó en $0. Eres oficialmente libre de deudas con nosotros. Adjúntale el PDF de Paz y Salvo que acabas de descargar."
                            else: msg_w = f"Hola {d_r['nom']}, confirmamos la recepción de tu pago por ${d_r['abono']:,.0f}. Tu nuevo saldo pendiente es de ${d_r['saldo']:,.0f}. Quedamos a tu entera disposición."
                            
                            link_wa = f"https://wa.me/57{d_r['tel']}?text={msg_w.replace(' ', '%20')}"
                            st.markdown(f"<a href='{link_wa}' target='_blank'><button style='background:#10B981; color:white; border:none; padding:10px 20px; border-radius:6px; font-weight:bold; cursor:pointer; width:100%;'>📱 AVISARLE AL CLIENTE POR WHATSAPP</button></a>", unsafe_allow_html=True)
                        else: st.warning("El cliente no tiene celular registrado en su expediente.")
                    
                    if st.button("Finalizar Proceso de Pago"): st.session_state.pop("pago_exitoso"); st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

# --- 11. USUARIOS (SISTEMA COMPLETO RESTAURADO) ---
elif st.session_state.pagina_actual == 'Usuarios':
    if st.session_state.rol_actual != "Administrador (Jefa)": st.error("⛔ ACCESO DENEGADO")
    else:
        st.markdown("<h1>Administración de Accesos y Seguridad</h1><p style='margin-bottom: 20px; color:#94A3B8;'>Control total de credenciales, cambio de claves y eliminación de personal.</p>", unsafe_allow_html=True)
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
        col_u1, col_u2 = st.columns(2)
        
        with col_u1:
            st.markdown("<div class='module-card'>", unsafe_allow_html=True)
            st.markdown("<h3>🔑 Forzar Cambio de Contraseña</h3>", unsafe_allow_html=True)
            with st.form("form_cambiar_pass"):
                u_sel_pass = st.selectbox("Seleccionar Usuario", df_u_actual["Usuario"].tolist())
                nueva_pass = st.text_input("Escribir Nueva Contraseña", type="password")
                if st.form_submit_button("ACTUALIZAR CONTRASEÑA"):
                    if nueva_pass:
                        df_u_actual.loc[df_u_actual["Usuario"].astype(str) == str(u_sel_pass), "Password"] = nueva_pass
                        df_u_actual.to_csv(ARCH_USR, index=False)
                        st.success(f"Contraseña de {u_sel_pass} actualizada."); st.rerun()
                    else: st.error("La contraseña no puede estar vacía.")
            st.markdown("</div>", unsafe_allow_html=True)
            
        with col_u2:
            st.markdown("<div class='module-card' style='border-top: 4px solid #EF4444;'>", unsafe_allow_html=True)
            st.markdown("<h3 style='color: #EF4444 !important;'>🗑️ Revocar Acceso y Expulsar</h3>", unsafe_allow_html=True)
            with st.form("form_eliminar_usr"):
                u_sel_del = st.selectbox("Seleccionar Usuario a Borrar", df_u_actual["Usuario"].tolist(), key="del_u")
                if st.form_submit_button("ELIMINAR DEFINITIVAMENTE"):
                    if str(u_sel_del) == "admin": st.error("⚠️ El administrador principal ('admin') no puede ser eliminado del sistema.")
                    elif str(u_sel_del) == str(st.session_state.usuario_actual): st.error("⚠ No puedes borrar la sesión en la que estás actualmente logueado.")
                    else:
                        df_u_del = df_u_actual[df_u_actual["Usuario"].astype(str) != str(u_sel_del)]
                        df_u_del.to_csv(ARCH_USR, index=False)
                        st.success(f"El acceso para '{u_sel_del}' ha sido revocado. Se le expulsará si está en línea."); st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("### 👥 Personal Autorizado en Plataforma")
        st.dataframe(df_u_actual[["Usuario", "Alias", "Rol", "Creador"]], use_container_width=True, hide_index=True)

# --- 12. SISTEMA (DUMP Y PAPELERA RESTAURADA) ---
elif st.session_state.pagina_actual == 'Sistema':
    if st.session_state.rol_actual != "Administrador (Jefa)": st.error("⛔ DENEGADO")
    else:
        st.markdown("<h1>Exportación General y Auditoría</h1>", unsafe_allow_html=True)
        st.markdown("<div class='module-card'>", unsafe_allow_html=True)
        st.markdown("<h3>📊 Data Dump (Copia de Seguridad Excel)</h3>", unsafe_allow_html=True)
        df_dump = df_cli.merge(df_fin, on="Cedula", how="left")
        st.download_button("📥 DESCARGAR BASE DE DATOS GLOBAL DE LA FIRMA (CSV)", df_dump.to_csv(index=False).encode('utf-8'), f"Respaldo_Firma_{hoy.strftime('%Y%m%d')}.csv", "text/csv")
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("<div class='module-card' style='border-top: 4px solid #EF4444;'>", unsafe_allow_html=True)
        st.markdown("<h3 style='color: #EF4444 !important;'>🗑️ Papelera de Reciclaje</h3><p style='color:#94A3B8;'>Eliminar o restaurar expedientes. El servidor borra la carpeta física automáticamente pasados 30 días del envío a papelera.</p>", unsafe_allow_html=True)
        
        if not df_activos.empty:
            cb = st.selectbox("Mover expediente a la papelera (Ocultar del sistema):", df_activos["Cedula"].astype(str) + " - " + df_activos["Nombre"])
            if st.button("ENVIAR A PAPELERA"):
                c_b = cb.split(" - ")[0]
                df_bd = pd.read_csv(ARCH_CLI)
                df_bd.loc[df_bd["Cedula"].astype(str) == c_b, "Estado"] = "Borrado"
                df_bd.loc[df_bd["Cedula"].astype(str) == c_b, "F_Borrado"] = hoy.strftime("%Y-%m-%d")
                df_bd.to_csv(ARCH_CLI, index=False); st.rerun()
                
        st.divider()
        if not df_papelera.empty:
            cr = st.selectbox("Restaurar desde la papelera:", df_papelera["Cedula"].astype(str) + " - " + df_papelera["Nombre"])
            if st.button("RESTAURAR EXPEDIENTE AL CLÚSTER ACTIVO"):
                c_r = cr.split(" - ")[0]
                df_bd = pd.read_csv(ARCH_CLI)
                df_bd.loc[df_bd["Cedula"].astype(str) == c_r, "Estado"] = "Activo"
                df_bd.loc[df_bd["Cedula"].astype(str) == c_r, "F_Borrado"] = ""
                df_bd.to_csv(ARCH_CLI, index=False); st.rerun()
        else:
            st.info("La papelera de reciclaje está vacía.")
        st.markdown("</div>", unsafe_allow_html=True)