"""Núcleo multi-firma de Insolvencia OS.

- Contraseñas con hash (PBKDF2) y bloqueo por intentos fallidos.
- Control_Central: hoja maestra con firmas, licencias y pagos (solo la ve el superadministrador).
- Alta automática de firmas: carpeta en Drive (Bóveda y Papelera), hoja de datos propia y usuario dueño.
- Panel del superadministrador.

Cada firma tiene su propia hoja de cálculo: los datos de una firma nunca comparten tabla con los de otra.
"""
import calendar
import hashlib
import hmac
import os
import re
import secrets
import time
import uuid
from datetime import datetime, timedelta, timezone
from html import escape

import pandas as pd
import requests
import streamlit as st

# Modo de pruebas locales (solo para desarrollo): guarda en carpetas en vez de Google.
LOCAL = os.environ.get("INSOLVENCIA_LOCAL") == "1"
DIR_LOCAL = os.environ.get("INSOLVENCIA_LOCAL_DIR", "db_local")

NOMBRE_CONTROL = "Control_Central"
NOMBRE_CARPETA_FIRMAS = "Insolvencia_OS_Firmas"
HOJA_PRINCIPAL_ANTIGUA = "DB_Insolvencia_Master"

ESTADOS = ["Prueba", "Activa", "Suspendida", "Baja"]
PLANES = ["Prueba", "Básico", "Profesional", "Despacho", "Propia"]
ROLES = ["Administrador (Jefa)", "Abogado / Operativo"]

COLS = {
    "clientes": ["Cedula", "Nombre", "Fuerza", "Telefono", "Email", "Deuda_Est", "Ingresos", "Senal", "F_Actualizacion", "Estado", "F_Borrado"],
    "finanzas": ["Cedula", "Honorarios", "Abonado"],
    "actuaciones": ["ID_Act", "Cedula", "Fecha", "Tipo", "Juzgado", "Radicado", "Anotacion"],
    "vencimientos": ["ID_Ven", "Cedula", "Cliente", "Asunto", "Fecha_Limite", "Estado"],
    "acreedores": ["ID_Acr", "Cedula", "Acreedor", "Cuantia", "Clase"],
    "audiencias": ["ID_Aud", "Cedula", "Cliente", "Fecha_Hora", "Motivo"],
    "tareas": ["ID_Tar", "Tarea", "Asignado", "Creador", "Estado", "Fecha"],
    "usuarios": ["Usuario", "Password", "Rol", "Creador", "Alias", "Avatar_Path", "Session_Token", "Cambiar_Clave"],
    "logs": ["Timestamp", "Usuario", "Modulo", "Accion"],
}
COLS_FIRMAS = ["Firma_ID", "Codigo", "Nombre", "Contacto", "Email", "Plan", "Estado", "Valor_Mensual", "Fecha_Inicio",
               "Fecha_Vence", "Dias_Gracia", "Max_Usuarios", "Sheet_ID", "Carpeta_ID", "Boveda_ID", "Papelera_ID", "Notas"]
COLS_PAGOS = ["ID_Pago", "Firma_ID", "Fecha", "Monto", "Medio", "Referencia", "Meses", "Vence_Hasta", "Registrado_Por"]
COLS_CONTROL = {"firmas": COLS_FIRMAS, "pagos": COLS_PAGOS}

# Recursos de la firma original (la que ya usaba la app antes de volverla multi-firma).
LEGADO = {"raiz": "", "papelera": ""}


def configurar_legado(raiz, papelera):
    LEGADO["raiz"] = raiz
    LEGADO["papelera"] = papelera


class ErrorDatos(Exception):
    """No se pudo leer o escribir en la base de datos (Google Sheets / Drive)."""


def ahora_co():
    return datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(hours=5)


# ==========================================================
# Contraseñas y bloqueo por intentos
# ==========================================================
ITERACIONES = 200_000


def hash_password(clave):
    sal = secrets.token_bytes(16)
    h = hashlib.pbkdf2_hmac("sha256", str(clave).encode("utf-8"), sal, ITERACIONES)
    return f"pbkdf2${ITERACIONES}${sal.hex()}${h.hex()}"


def es_hash(valor):
    return str(valor).startswith("pbkdf2$")


def verificar_password(clave, almacenado):
    """Acepta hash PBKDF2 y, para migrar cuentas antiguas, contraseñas en texto plano."""
    almacenado = str(almacenado)
    clave = str(clave)
    if es_hash(almacenado):
        try:
            _, it, sal, h = almacenado.split("$")
            nuevo = hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), bytes.fromhex(sal), int(it))
            return hmac.compare_digest(nuevo.hex(), h)
        except (ValueError, TypeError):
            return False
    if almacenado in ("", "nan"):
        return False
    return hmac.compare_digest(clave.encode("utf-8"), almacenado.encode("utf-8"))


def validar_clave_nueva(clave):
    """Devuelve None si es válida, o el motivo por el que no lo es."""
    clave = str(clave)
    if len(clave) < 8:
        return "La contraseña debe tener al menos 8 caracteres."
    if not re.search(r"[A-Za-z]", clave) or not re.search(r"\d", clave):
        return "La contraseña debe combinar letras y números."
    if clave.lower() in ("12345678", "password1", "contraseña1"):
        return "Esa contraseña es demasiado común."
    return None


def clave_temporal():
    return secrets.token_urlsafe(9)


@st.cache_resource
def _caches():
    return {"gc": None, "ss": {}, "ws": {}, "control": None, "firmas": None, "carpeta_firmas": None, "intentos": {}}


MAX_INTENTOS = 5
BLOQUEO_SEG = 15 * 60


def segundos_bloqueo(clave):
    dato = _caches()["intentos"].get(clave)
    if not dato:
        return 0
    hasta = dato.get("hasta", 0)
    return max(0, int(hasta - time.time()))


def registrar_fallo(clave):
    intentos = _caches()["intentos"]
    dato = intentos.setdefault(clave, {"n": 0, "hasta": 0})
    if dato["hasta"] and dato["hasta"] < time.time():
        dato.update(n=0, hasta=0)
    dato["n"] += 1
    if dato["n"] >= MAX_INTENTOS:
        dato["hasta"] = time.time() + BLOQUEO_SEG
        dato["n"] = 0


def limpiar_fallos(clave):
    _caches()["intentos"].pop(clave, None)


# ==========================================================
# Superadministrador (credenciales en Secrets, nunca en una hoja)
# ==========================================================
def _secreto(nombre, defecto=""):
    try:
        return st.secrets.get(nombre, defecto)
    except Exception:  # no hay archivo de secretos
        return defecto


def codigo_superadmin():
    return str(_secreto("superadmin_codigo", "superadmin")).strip().lower()


def login_superadmin(usuario, clave):
    u = str(_secreto("superadmin_user", "")).strip().lower()
    if not u or str(usuario).strip().lower() != u:
        return False
    h = str(_secreto("superadmin_hash", ""))
    if h:
        return verificar_password(clave, h)
    p = str(_secreto("superadmin_password", ""))
    return bool(p) and hmac.compare_digest(str(clave).encode("utf-8"), p.encode("utf-8"))


# ==========================================================
# Google (Sheets + Drive)
# ==========================================================
def _creds():
    from google.oauth2.credentials import Credentials
    return Credentials.from_authorized_user_info({
        "client_id": st.secrets["google_client_id"],
        "client_secret": st.secrets["google_client_secret"],
        "refresh_token": st.secrets["google_refresh_token"],
        "token_uri": "https://oauth2.googleapis.com/token",
    })


@st.cache_data(ttl=3000, show_spinner=False)
def _token():
    from google.auth.transport.requests import Request
    c = _creds()
    c.refresh(Request())
    return c.token


def _gc():
    cache = _caches()
    if cache["gc"] is None:
        import gspread
        cache["gc"] = gspread.authorize(_creds())
    return cache["gc"]


def _hdr():
    return {"Authorization": f"Bearer {_token()}", "Content-Type": "application/json"}


def _drive_buscar(nombre, padre, mime=None):
    nombre = nombre.replace("'", "\\'")
    q = f"name='{nombre}' and '{padre}' in parents and trashed=false"
    if mime:
        q += f" and mimeType='{mime}'"
    r = requests.get("https://www.googleapis.com/drive/v3/files", params={"q": q, "fields": "files(id,name)", "pageSize": 5},
                     headers=_hdr(), timeout=20)
    r.raise_for_status()
    archivos = r.json().get("files", [])
    return archivos[0]["id"] if archivos else None


def _drive_crear_carpeta(nombre, padre=None):
    meta = {"name": nombre, "mimeType": "application/vnd.google-apps.folder"}
    if padre:
        meta["parents"] = [padre]
    r = requests.post("https://www.googleapis.com/drive/v3/files", json=meta, headers=_hdr(), timeout=20)
    r.raise_for_status()
    return r.json()["id"]


def _carpeta_firmas():
    cache = _caches()
    if cache["carpeta_firmas"]:
        return cache["carpeta_firmas"]
    fija = str(_secreto("drive_firmas_id", "")).strip()
    if fija:
        cache["carpeta_firmas"] = fija
        return fija
    existente = _drive_buscar(NOMBRE_CARPETA_FIRMAS, "root", "application/vnd.google-apps.folder")
    cache["carpeta_firmas"] = existente or _drive_crear_carpeta(NOMBRE_CARPETA_FIRMAS)
    return cache["carpeta_firmas"]


def _numericise(v):
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return ""
        try:
            return int(s)
        except ValueError:
            try:
                return float(s)
            except ValueError:
                return v
    return v


def _ss(ref):
    cache = _caches()
    if ref not in cache["ss"]:
        cache["ss"][ref] = _gc().open_by_key(ref)
    return cache["ss"][ref]


def _ws(ref, tabla, cols):
    cache = _caches()
    clave = (ref, tabla)
    if clave not in cache["ws"]:
        import gspread
        ss = _ss(ref)
        try:
            cache["ws"][clave] = ss.worksheet(tabla)
        except gspread.WorksheetNotFound:
            ws = ss.add_worksheet(title=tabla, rows=200, cols=max(len(cols), 10))
            ws.update(values=[cols], range_name="A1")
            cache["ws"][clave] = ws
    return cache["ws"][clave]


def _olvidar(ref):
    cache = _caches()
    cache["ss"].pop(ref, None)
    for k in [k for k in cache["ws"] if k[0] == ref]:
        cache["ws"].pop(k, None)


def _col_letra(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _ruta_local(ref, tabla):
    carpeta = os.path.join(DIR_LOCAL, str(ref))
    os.makedirs(carpeta, exist_ok=True)
    return os.path.join(carpeta, f"{tabla}.csv")


def leer_registros(ref, tabla, cols):
    """Lee una tabla de la hoja `ref`. Lanza ErrorDatos si no puede (nunca devuelve datos falsos)."""
    if LOCAL:
        ruta = _ruta_local(ref, tabla)
        if not os.path.exists(ruta):
            pd.DataFrame(columns=cols).to_csv(ruta, index=False)
        df = pd.read_csv(ruta, dtype=str, keep_default_na=False)
        df = df.apply(lambda col: col.map(_numericise))
    else:
        try:
            data = _ws(ref, tabla, cols).get_all_records()
            df = pd.DataFrame(data)
        except Exception as e:
            _olvidar(ref)
            raise ErrorDatos(f"No se pudo leer '{tabla}' en Google Sheets ({type(e).__name__}). Intenta de nuevo.") from e
    if df.empty:
        df = pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df


def escribir_registros(ref, tabla, df):
    """Guarda la tabla completa. Escribe primero y limpia sobrantes después, para no perder datos si algo falla."""
    if LOCAL:
        df.to_csv(_ruta_local(ref, tabla), index=False)
        return
    try:
        cols = list(df.columns)
        ws = _ws(ref, tabla, cols)
        valores = [cols] + df.fillna("").values.tolist()
        n_filas, n_cols = len(valores), len(cols)
        if ws.row_count < n_filas + 20 or ws.col_count < n_cols:
            ws.resize(rows=max(ws.row_count, n_filas + 50), cols=max(ws.col_count, n_cols))
        ws.update(values=valores, range_name="A1")
        if ws.row_count > n_filas:
            ws.batch_clear([f"A{n_filas + 1}:{_col_letra(max(n_cols, 26))}"])
    except Exception as e:
        _olvidar(ref)
        raise ErrorDatos(f"No se pudo guardar '{tabla}' en Google Sheets ({type(e).__name__}). Tus cambios NO se guardaron: intenta de nuevo.") from e


# ==========================================================
# Control_Central
# ==========================================================
def _crear_hoja_con_tablas(titulo, carpeta_id, tablas):
    """Crea una hoja de cálculo dentro de una carpeta, con una pestaña por tabla y sus encabezados."""
    ss = _gc().create(titulo, folder_id=carpeta_id)
    nombres = list(tablas.keys())
    primera = ss.sheet1
    primera.update_title(nombres[0])
    primera.update(values=[tablas[nombres[0]]], range_name="A1")
    for n in nombres[1:]:
        ws = ss.add_worksheet(title=n, rows=200, cols=max(len(tablas[n]), 10))
        ws.update(values=[tablas[n]], range_name="A1")
    return ss


def _sembrar_principal(control_ss):
    """Registra la firma que ya usaba la app (si existe), para no perder sus datos."""
    try:
        ss_viejo = _gc().open(HOJA_PRINCIPAL_ANTIGUA)
    except Exception:
        return
    fila = {c: "" for c in COLS_FIRMAS}
    fila.update({
        "Firma_ID": "F" + uuid.uuid4().hex[:8], "Codigo": "principal", "Nombre": "Firma principal", "Plan": "Propia",
        "Estado": "Activa", "Valor_Mensual": 0, "Fecha_Inicio": ahora_co().strftime("%Y-%m-%d"), "Fecha_Vence": "2099-12-31",
        "Dias_Gracia": 0, "Max_Usuarios": 50, "Sheet_ID": ss_viejo.id, "Carpeta_ID": LEGADO["raiz"],
        "Boveda_ID": LEGADO["raiz"], "Papelera_ID": LEGADO["papelera"], "Notas": "Firma original migrada automáticamente",
    })
    ws = control_ss.worksheet("firmas")
    ws.update(values=[COLS_FIRMAS, [fila[c] for c in COLS_FIRMAS]], range_name="A1")


def control_ref():
    if LOCAL:
        return "_control"
    cache = _caches()
    if cache["control"]:
        return cache["control"]
    fijo = str(_secreto("control_sheet_id", "")).strip()
    if fijo:
        cache["control"] = fijo
        return fijo
    try:
        raiz = _carpeta_firmas()
        existente = _drive_buscar(NOMBRE_CONTROL, raiz, "application/vnd.google-apps.spreadsheet")
        if existente:
            cache["control"] = existente
        else:
            ss = _crear_hoja_con_tablas(NOMBRE_CONTROL, raiz, COLS_CONTROL)
            _sembrar_principal(ss)
            cache["control"] = ss.id
    except Exception as e:
        raise ErrorDatos(f"No se pudo abrir o crear Control_Central ({type(e).__name__}).") from e
    return cache["control"]


def control_leer(tabla):
    return leer_registros(control_ref(), tabla, COLS_CONTROL[tabla])


def control_guardar(tabla, df):
    escribir_registros(control_ref(), tabla, df)
    if tabla == "firmas":
        _caches()["firmas"] = None


def firmas_df(fresco=False):
    cache = _caches()
    item = cache["firmas"]
    if item and not fresco and (time.time() - item[0]) < 45:
        return item[1].copy()
    df = control_leer("firmas").astype(object)
    cache["firmas"] = (time.time(), df)
    return df.copy()


def _fila_a_dict(fila):
    d = {c: ("" if pd.isna(fila[c]) else str(fila[c]).strip()) for c in COLS_FIRMAS}
    d["Codigo"] = d["Codigo"].lower()
    return d


def obtener_firma(firma_id=None, codigo=None, fresco=False):
    df = firmas_df(fresco=fresco)
    if df.empty:
        return None
    if firma_id:
        sel = df[df["Firma_ID"].astype(str) == str(firma_id)]
    else:
        sel = df[df["Codigo"].astype(str).str.lower() == str(codigo).strip().lower()]
    return _fila_a_dict(sel.iloc[0]) if not sel.empty else None


# ==========================================================
# Licencias
# ==========================================================
def _fecha(valor):
    try:
        return datetime.strptime(str(valor)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def _entero(valor, defecto=0):
    try:
        return int(float(str(valor).replace(",", ".")))
    except ValueError:
        return defecto


def _num(valor):
    try:
        s = str(valor).replace("$", "").replace(" ", "").strip()
        if s in ("", "nan"):
            return 0.0
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
        else:
            for sep in (",", "."):
                if sep in s:
                    partes = s.split(sep)
                    s = s.replace(sep, "") if (len(partes) > 2 or len(partes[-1]) == 3) else s.replace(sep, ".")
        return float(s)
    except ValueError:
        return 0.0


def evaluar_licencia(firma, hoy=None):
    """nivel: ok | por_vencer | gracia | vencida | bloqueada."""
    hoy = (hoy or ahora_co()).date()
    estado = str(firma.get("Estado", "")).strip()
    if estado in ("Suspendida", "Baja"):
        return {"nivel": "bloqueada", "dias": 0,
                "mensaje": "El acceso de tu firma está suspendido. Comunícate con soporte para reactivarlo."}
    vence = _fecha(firma.get("Fecha_Vence"))
    gracia = max(_entero(firma.get("Dias_Gracia"), 0), 0)
    if vence is None:
        return {"nivel": "vencida", "dias": 0, "mensaje": "Tu suscripción no tiene una fecha de vigencia válida."}
    if hoy <= vence:
        dias = (vence - hoy).days
        if dias <= 5:
            return {"nivel": "por_vencer", "dias": dias,
                    "mensaje": f"Tu suscripción vence en {dias} día(s) ({vence.strftime('%d/%m/%Y')}). Renueva para no perder el acceso."}
        return {"nivel": "ok", "dias": dias, "mensaje": ""}
    if hoy <= vence + timedelta(days=gracia):
        restan = (vence + timedelta(days=gracia) - hoy).days
        return {"nivel": "gracia", "dias": restan,
                "mensaje": f"Tu suscripción venció el {vence.strftime('%d/%m/%Y')}. Tienes {restan} día(s) de gracia antes de que se bloquee el acceso."}
    return {"nivel": "vencida", "dias": 0,
            "mensaje": f"Tu suscripción venció el {vence.strftime('%d/%m/%Y')} y terminó el período de gracia."}


def sumar_meses(fecha, meses):
    m = fecha.month - 1 + int(meses)
    anio = fecha.year + m // 12
    mes = m % 12 + 1
    dia = min(fecha.day, calendar.monthrange(anio, mes)[1])
    return fecha.replace(year=anio, month=mes, day=dia)


# ==========================================================
# Alta de firmas
# ==========================================================
def crear_firma(codigo, nombre, contacto, email, plan, valor, max_usuarios, vence, gracia, usuario_dueno, alias_dueno):
    codigo = str(codigo).strip().lower()
    usuario_dueno = str(usuario_dueno).strip().lower()
    if not re.fullmatch(r"[a-z0-9]{3,20}", codigo):
        return {"ok": False, "error": "El código debe tener de 3 a 20 letras minúsculas o números, sin espacios."}
    if codigo == codigo_superadmin():
        return {"ok": False, "error": "Ese código está reservado."}
    if not re.fullmatch(r"[a-z0-9._-]{3,30}", usuario_dueno):
        return {"ok": False, "error": "El usuario del dueño debe tener de 3 a 30 caracteres (letras, números, punto, guion)."}
    if not str(nombre).strip():
        return {"ok": False, "error": "Escribe el nombre de la firma."}
    try:
        existentes = firmas_df(fresco=True)
        if not existentes.empty and codigo in existentes["Codigo"].astype(str).str.lower().tolist():
            return {"ok": False, "error": f"Ya existe una firma con el código '{codigo}'."}

        firma_id = "F" + uuid.uuid4().hex[:8]
        clave = clave_temporal()
        if LOCAL:
            ref, carpeta, boveda, papelera = firma_id, "local", "local", "local"
            for tabla, cols in COLS.items():
                escribir_registros(ref, tabla, pd.DataFrame(columns=cols))
        else:
            raiz = _carpeta_firmas()
            carpeta = _drive_crear_carpeta(f"Firma_{codigo}", raiz)
            boveda = _drive_crear_carpeta("Bóveda", carpeta)
            papelera = _drive_crear_carpeta("Papelera", carpeta)
            ss = _crear_hoja_con_tablas(f"BD_{codigo}", carpeta, COLS)
            ref = ss.id

        dueno = pd.DataFrame([{
            "Usuario": usuario_dueno, "Password": hash_password(clave), "Rol": ROLES[0], "Creador": "Superadmin",
            "Alias": alias_dueno or usuario_dueno, "Avatar_Path": "", "Session_Token": "", "Cambiar_Clave": "Si",
        }])
        escribir_registros(ref, "usuarios", dueno)

        fila = {c: "" for c in COLS_FIRMAS}
        fila.update({
            "Firma_ID": firma_id, "Codigo": codigo, "Nombre": str(nombre).strip(), "Contacto": contacto, "Email": email,
            "Plan": plan, "Estado": "Prueba" if plan == "Prueba" else "Activa", "Valor_Mensual": valor,
            "Fecha_Inicio": ahora_co().strftime("%Y-%m-%d"), "Fecha_Vence": vence.strftime("%Y-%m-%d"),
            "Dias_Gracia": gracia, "Max_Usuarios": max_usuarios, "Sheet_ID": ref, "Carpeta_ID": carpeta,
            "Boveda_ID": boveda, "Papelera_ID": papelera, "Notas": "",
        })
        control_guardar("firmas", pd.concat([existentes.astype(object), pd.DataFrame([fila])], ignore_index=True))
        return {"ok": True, "firma_id": firma_id, "codigo": codigo, "usuario": usuario_dueno, "clave": clave}
    except Exception as e:
        return {"ok": False, "error": f"No se pudo completar el alta ({type(e).__name__}: {e}). Revisa Drive: puede haber quedado una carpeta a medias."}


def restablecer_clave(firma, usuario):
    """Genera una contraseña temporal para un usuario de una firma (soporte)."""
    df = leer_registros(firma["Sheet_ID"], "usuarios", COLS["usuarios"]).astype(object)
    m = df["Usuario"].astype(str).str.lower() == str(usuario).lower()
    if not m.any():
        return None
    clave = clave_temporal()
    df.loc[m, "Password"] = hash_password(clave)
    df.loc[m, "Cambiar_Clave"] = "Si"
    df.loc[m, "Session_Token"] = ""
    escribir_registros(firma["Sheet_ID"], "usuarios", df)
    return clave


def registrar_pago(firma_id, monto, medio, referencia, meses, reactivar=True, por="Superadmin"):
    firmas = firmas_df(fresco=True).astype(object)
    m = firmas["Firma_ID"].astype(str) == str(firma_id)
    if not m.any():
        raise ErrorDatos("La firma no existe.")
    hoy = ahora_co().date()
    vence_actual = _fecha(firmas.loc[m, "Fecha_Vence"].iloc[0]) or hoy
    base = vence_actual if vence_actual >= hoy else hoy
    nueva = sumar_meses(base, meses)
    firmas.loc[m, "Fecha_Vence"] = nueva.strftime("%Y-%m-%d")
    if reactivar and str(firmas.loc[m, "Estado"].iloc[0]) in ("Suspendida", "Prueba"):
        firmas.loc[m, "Estado"] = "Activa"
    pagos = control_leer("pagos").astype(object)
    fila = {"ID_Pago": "PAG-" + uuid.uuid4().hex[:8], "Firma_ID": firma_id, "Fecha": hoy.strftime("%Y-%m-%d"), "Monto": monto,
            "Medio": medio, "Referencia": referencia, "Meses": meses, "Vence_Hasta": nueva.strftime("%Y-%m-%d"), "Registrado_Por": por}
    control_guardar("pagos", pd.concat([pagos, pd.DataFrame([fila])], ignore_index=True))
    control_guardar("firmas", firmas)
    return nueva


def enviar_recordatorios(dias_aviso=5):
    """Correo a cada firma que vence pronto o ya venció. Resumen para el superadministrador."""
    from notificaciones import enviar_email
    firmas = firmas_df(fresco=True)
    hoy = ahora_co().date()
    resultados, resumen = [], []
    for _, r in firmas.iterrows():
        f = _fila_a_dict(r)
        if f["Estado"] not in ("Activa", "Prueba"):
            continue
        vence = _fecha(f["Fecha_Vence"])
        if vence is None or f["Codigo"] == "principal":
            continue
        dias = (vence - hoy).days
        if dias > dias_aviso:
            continue
        estado = f"vence en {dias} día(s)" if dias >= 0 else f"venció hace {abs(dias)} día(s)"
        resumen.append(f"• {f['Nombre']} ({f['Codigo']}): {estado}")
        if f["Email"]:
            cuerpo = (f"<p>Hola, equipo de <b>{escape(f['Nombre'])}</b>.</p>"
                      f"<p>Te recordamos que tu suscripción a Insolvencia OS {estado} ({vence.strftime('%d/%m/%Y')}).</p>"
                      f"<p>Para mantener tu acceso, realiza el pago y avísanos para registrarlo.</p>")
            ok, msg = enviar_email(f["Email"], "Recordatorio de tu suscripción a Insolvencia OS", cuerpo,
                                   f"Tu suscripción {estado} ({vence.strftime('%d/%m/%Y')}).")
            resultados.append((f["Nombre"], ok, msg))
        else:
            resultados.append((f["Nombre"], False, "La firma no tiene correo registrado."))
    if resumen:
        destino = _secreto("alertas_email_destino", _secreto("gmail_usuario", ""))
        if destino:
            enviar_email(destino, "Cobros próximos - Insolvencia OS", "<br>".join(resumen), "\n".join(resumen))
    return resultados


# ==========================================================
# Panel del superadministrador
# ==========================================================
SA_PAGINAS = [("resumen", "📊 Resumen"), ("firmas", "🏢 Firmas"), ("nueva", "➕ Nueva firma"), ("pagos", "💳 Pagos"), ("cuenta", "🔐 Mi cuenta")]


def _kpi(icono, etiqueta, valor, color, fondo):
    return (f"<div class='kpi' style='--accent:{color}; --accent-bg:{fondo};'><div class='ico'>{icono}</div>"
            f"<p class='lbl'>{etiqueta}</p><p class='val'>{valor}</p></div>")


def _ir(pagina):
    st.session_state.sa_pagina = pagina


def render_superadmin(cerrar_cb):
    if "sa_pagina" not in st.session_state:
        st.session_state.sa_pagina = "resumen"
    pag = st.session_state.sa_pagina

    with st.container(key="topbar"):
        c1, c2, c3 = st.columns([2, 3, 1.3], vertical_alignment="center")
        c1.markdown("<div class='marca'><div class='sello'>⚖️</div><div><div class='nombre'>Insolvencia OS</div><div class='lema'>Superadministrador</div></div></div>", unsafe_allow_html=True)
        c2.markdown("<span class='chip2' style='background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.22);color:#fff;border-radius:30px;padding:5px 14px;font-size:12.5px;font-weight:600;'>🔒 Solo tú ves este panel</span>", unsafe_allow_html=True)
        c3.button("🚪 Cerrar sesión", key="sa_logout", use_container_width=True, on_click=cerrar_cb)

    with st.container(key="nav_grupos"):
        cols = st.columns(len(SA_PAGINAS))
        for col, (clave, etiqueta) in zip(cols, SA_PAGINAS):
            with col:
                st.button(etiqueta, key=f"sa_{clave}", use_container_width=True,
                          type="primary" if clave == pag else "secondary", on_click=_ir, args=(clave,))

    try:
        firmas = firmas_df()
    except ErrorDatos as e:
        st.error(f"🚨 {e}")
        st.info("Revisa que los Secrets de Google estén bien (google_client_id, google_client_secret y google_refresh_token).")
        return

    hoy = ahora_co().date()
    if pag == "resumen":
        _sa_resumen(firmas, hoy)
    elif pag == "firmas":
        _sa_firmas(firmas, hoy)
    elif pag == "nueva":
        _sa_nueva(hoy)
    elif pag == "pagos":
        _sa_pagos(firmas, hoy)
    else:
        _sa_cuenta(firmas)


def _sa_resumen(firmas, hoy):
    st.markdown("<h1>Control de firmas</h1><p class='page-sub'>Estado de cada despacho que usa la plataforma.</p>", unsafe_allow_html=True)
    activas = gracia = vencidas = susp = 0
    mrr = 0.0
    filas = []
    for _, r in firmas.iterrows():
        f = _fila_a_dict(r)
        lic = evaluar_licencia(f)
        nivel = lic["nivel"]
        if nivel in ("ok", "por_vencer"):
            activas += 1
            mrr += _num(f["Valor_Mensual"])
        elif nivel == "gracia":
            gracia += 1
        elif nivel == "vencida":
            vencidas += 1
        else:
            susp += 1
        etiqueta = {"ok": "🟢 Al día", "por_vencer": "🟠 Por vencer", "gracia": "🟡 En gracia", "vencida": "🔴 Vencida", "bloqueada": "⛔ Suspendida"}[nivel]
        vence = _fecha(f["Fecha_Vence"])
        filas.append({"Firma": f["Nombre"], "Código": f["Codigo"], "Plan": f["Plan"], "Estado": etiqueta,
                      "Vence": vence.strftime("%d/%m/%Y") if vence else "—",
                      "Días": (vence - hoy).days if vence else None, "Valor mensual": _num(f["Valor_Mensual"])})
    k = st.columns(4)
    k[0].markdown(_kpi("✅", "Firmas al día", activas, "#1F7A5C", "rgba(31,122,92,.14)"), unsafe_allow_html=True)
    k[1].markdown(_kpi("🟡", "En período de gracia", gracia, "#C9A24B", "rgba(201,162,75,.2)"), unsafe_allow_html=True)
    k[2].markdown(_kpi("⛔", "Vencidas o suspendidas", vencidas + susp, "#8A1F35", "rgba(138,31,53,.12)"), unsafe_allow_html=True)
    k[3].markdown(_kpi("💰", "Ingreso mensual (al día)", f"$ {mrr:,.0f}", "#2F6DB5", "rgba(47,109,181,.14)"), unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("<p class='card-title'>Todas las firmas</p>", unsafe_allow_html=True)
        if not filas:
            st.markdown("<div class='vacio'>Aún no hay firmas. Crea la primera en “Nueva firma”.</div>", unsafe_allow_html=True)
        else:
            st.dataframe(pd.DataFrame(filas).sort_values("Días", na_position="last"), use_container_width=True, hide_index=True)
    with st.container(border=True):
        st.markdown("<p class='card-title'>📣 Recordatorios de cobro</p>", unsafe_allow_html=True)
        st.caption("Envía un correo a cada firma que vence en 5 días o menos (o ya venció) y te manda el resumen a ti.")
        if st.button("Enviar recordatorios ahora", key="sa_recordar"):
            res = enviar_recordatorios()
            if not res:
                st.success("No hay firmas por vencer en los próximos 5 días.")
            for nombre, ok, msg in res:
                (st.success if ok else st.warning)(f"{nombre}: {msg}")


def _sa_firmas(firmas, hoy):
    st.markdown("<h1>Firmas</h1><p class='page-sub'>Edita planes y vencimientos, suspende o reactiva el acceso.</p>", unsafe_allow_html=True)
    if firmas.empty:
        st.info("Aún no hay firmas registradas.")
        return
    opciones = {f"{r['Nombre']} · {r['Codigo']}": str(r["Firma_ID"]) for _, r in firmas.iterrows()}
    nombre_sel = st.selectbox("Firma", list(opciones.keys()), key="sa_sel_firma")
    firma = obtener_firma(firma_id=opciones[nombre_sel])
    lic = evaluar_licencia(firma)
    etiqueta = {"ok": "🟢 Al día", "por_vencer": "🟠 Por vencer", "gracia": "🟡 En gracia", "vencida": "🔴 Vencida", "bloqueada": "⛔ Suspendida"}[lic["nivel"]]
    st.markdown(f"**Estado de acceso:** {etiqueta}" + (f" — {escape(lic['mensaje'])}" if lic["mensaje"] else ""), unsafe_allow_html=True)

    a, b = st.columns(2)
    if str(firma["Estado"]) in ("Suspendida", "Baja"):
        if a.button("✅ Reactivar acceso", key="sa_reactivar", use_container_width=True):
            _cambiar_estado(firma["Firma_ID"], "Activa")
            st.rerun()
    else:
        if a.button("⛔ Suspender acceso ahora", key="sa_suspender", use_container_width=True):
            _cambiar_estado(firma["Firma_ID"], "Suspendida")
            st.rerun()
    b.caption("La suspensión bloquea a la dueña y a todos los usuarios de la firma en menos de un minuto. Los datos se conservan.")

    with st.form("sa_form_firma"):
        c1, c2 = st.columns(2)
        nombre = c1.text_input("Nombre de la firma", value=firma["Nombre"])
        contacto = c2.text_input("Persona de contacto", value=firma["Contacto"])
        c3, c4 = st.columns(2)
        email = c3.text_input("Correo de contacto (recibe avisos de cobro)", value=firma["Email"])
        plan = c4.selectbox("Plan", PLANES, index=PLANES.index(firma["Plan"]) if firma["Plan"] in PLANES else 0)
        c5, c6, c7 = st.columns(3)
        valor = c5.number_input("Valor mensual ($)", min_value=0.0, step=10000.0, value=float(_num(firma["Valor_Mensual"])))
        max_u = c6.number_input("Usuarios máximos", min_value=1, step=1, value=max(_entero(firma["Max_Usuarios"], 3), 1))
        gracia = c7.number_input("Días de gracia", min_value=0, step=1, value=max(_entero(firma["Dias_Gracia"], 0), 0))
        c8, c9 = st.columns(2)
        vence = c8.date_input("Vigente hasta", value=_fecha(firma["Fecha_Vence"]) or hoy)
        estado = c9.selectbox("Estado", ESTADOS, index=ESTADOS.index(firma["Estado"]) if firma["Estado"] in ESTADOS else 1)
        notas = st.text_area("Notas internas", value=firma["Notas"])
        if st.form_submit_button("GUARDAR CAMBIOS"):
            df = firmas_df(fresco=True).astype(object)
            m = df["Firma_ID"].astype(str) == firma["Firma_ID"]
            for campo, val in {"Nombre": nombre, "Contacto": contacto, "Email": email, "Plan": plan, "Valor_Mensual": valor,
                               "Max_Usuarios": int(max_u), "Dias_Gracia": int(gracia), "Fecha_Vence": vence.strftime("%Y-%m-%d"),
                               "Estado": estado, "Notas": notas}.items():
                df.loc[m, campo] = val
            control_guardar("firmas", df)
            st.success("Cambios guardados.")
            st.rerun()

    with st.expander("🔑 Soporte: restablecer la contraseña de un usuario de esta firma"):
        try:
            usuarios = leer_registros(firma["Sheet_ID"], "usuarios", COLS["usuarios"])
            lista = usuarios["Usuario"].astype(str).tolist()
        except ErrorDatos as e:
            st.error(str(e))
            lista = []
        if lista:
            elegido = st.selectbox("Usuario", lista, key="sa_reset_user")
            if st.button("Generar contraseña temporal", key="sa_reset_btn"):
                clave = restablecer_clave(firma, elegido)
                st.success("Listo. Entrégala por un canal seguro; deberá cambiarla al entrar.")
                st.code(f"Código: {firma['Codigo']}\nUsuario: {elegido}\nContraseña temporal: {clave}")
        st.caption("Aquí solo ves los nombres de usuario. No tienes acceso al contenido de los expedientes desde este panel.")


def _cambiar_estado(firma_id, estado):
    df = firmas_df(fresco=True).astype(object)
    df.loc[df["Firma_ID"].astype(str) == str(firma_id), "Estado"] = estado
    control_guardar("firmas", df)


def _sa_nueva(hoy):
    st.markdown("<h1>Nueva firma</h1><p class='page-sub'>Crea su hoja de datos, su Bóveda y Papelera en Drive y el usuario de la dueña, todo en un paso.</p>", unsafe_allow_html=True)
    with st.form("sa_form_nueva", clear_on_submit=False):
        c1, c2 = st.columns(2)
        codigo = c1.text_input("Código de la firma", help="Es lo que escribirán para entrar. Ej.: bufetegomez")
        nombre = c2.text_input("Nombre de la firma")
        c3, c4 = st.columns(2)
        contacto = c3.text_input("Persona de contacto")
        email = c4.text_input("Correo (recibe avisos de cobro)")
        c5, c6 = st.columns(2)
        usuario = c5.text_input("Usuario de la dueña", help="Ej.: mgomez")
        alias = c6.text_input("Nombre de la dueña")
        c7, c8, c9 = st.columns(3)
        plan = c7.selectbox("Plan", PLANES, index=0)
        valor = c8.number_input("Valor mensual ($)", min_value=0.0, step=10000.0)
        max_u = c9.number_input("Usuarios máximos", min_value=1, step=1, value=3)
        c10, c11 = st.columns(2)
        dias = c10.number_input("Días de prueba o primer período", min_value=1, step=1, value=30)
        gracia = c11.number_input("Días de gracia", min_value=0, step=1, value=5)
        enviado = st.form_submit_button("CREAR FIRMA")
    if enviado:
        with st.spinner("Creando carpetas, hoja de datos y usuario…"):
            res = crear_firma(codigo, nombre, contacto, email, plan, valor, int(max_u), hoy + timedelta(days=int(dias)), int(gracia), usuario, alias)
        if res["ok"]:
            st.success("✅ Firma creada. Entrega estos datos a la dueña por un canal seguro (se muestran solo una vez).")
            st.code(f"Código de firma: {res['codigo']}\nUsuario: {res['usuario']}\nContraseña temporal: {res['clave']}")
            st.caption("Al entrar por primera vez, la app le pedirá cambiar la contraseña.")
        else:
            st.error(res["error"])


def _sa_pagos(firmas, hoy):
    st.markdown("<h1>Pagos</h1><p class='page-sub'>Cada pago registrado extiende automáticamente la vigencia de la firma.</p>", unsafe_allow_html=True)
    if firmas.empty:
        st.info("Aún no hay firmas registradas.")
        return
    opciones = {f"{r['Nombre']} · {r['Codigo']}": str(r["Firma_ID"]) for _, r in firmas.iterrows()}
    with st.form("sa_form_pago", clear_on_submit=True):
        nombre_sel = st.selectbox("Firma", list(opciones.keys()))
        c1, c2, c3 = st.columns(3)
        monto = c1.number_input("Monto recibido ($)", min_value=0.0, step=10000.0)
        meses = c2.number_input("Meses que cubre", min_value=1, max_value=24, step=1, value=1)
        medio = c3.selectbox("Medio", ["Transferencia", "Nequi", "Daviplata", "Wompi", "Efectivo", "Otro"])
        referencia = st.text_input("Referencia o comprobante")
        reactivar = st.checkbox("Reactivar la firma si estaba suspendida o en prueba", value=True)
        if st.form_submit_button("REGISTRAR PAGO"):
            if monto <= 0:
                st.error("Indica el monto recibido.")
            else:
                nueva = registrar_pago(opciones[nombre_sel], monto, medio, referencia, int(meses), reactivar)
                st.success(f"Pago registrado. La firma queda vigente hasta el {nueva.strftime('%d/%m/%Y')}.")
    with st.container(border=True):
        st.markdown("<p class='card-title'>Historial de pagos</p>", unsafe_allow_html=True)
        pagos = control_leer("pagos")
        if pagos.empty:
            st.markdown("<div class='vacio'>Aún no hay pagos registrados.</div>", unsafe_allow_html=True)
        else:
            nombres = {str(r["Firma_ID"]): r["Nombre"] for _, r in firmas.iterrows()}
            vista = pagos.copy()
            vista["Firma"] = vista["Firma_ID"].astype(str).map(nombres).fillna("—")
            st.dataframe(vista[["Fecha", "Firma", "Monto", "Medio", "Referencia", "Meses", "Vence_Hasta"]].iloc[::-1], use_container_width=True, hide_index=True)


def _sa_cuenta(firmas):
    st.markdown("<h1>Mi cuenta</h1><p class='page-sub'>Seguridad de tu acceso y estado del sistema.</p>", unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("<p class='card-title'>🔐 Generar el hash de tu contraseña</p>", unsafe_allow_html=True)
        st.caption("Escribe la contraseña que quieres usar. Copia el resultado en Secrets como superadmin_hash y borra superadmin_password.")
        nueva = st.text_input("Nueva contraseña", type="password", key="sa_nueva_clave")
        if nueva:
            motivo = validar_clave_nueva(nueva)
            if motivo:
                st.warning(motivo)
            else:
                st.code(f'superadmin_hash = "{hash_password(nueva)}"')
    with st.container(border=True):
        st.markdown("<p class='card-title'>🩺 Estado del sistema</p>", unsafe_allow_html=True)
        if st.button("Probar conexión con Google", key="sa_probar"):
            try:
                n = len(firmas_df(fresco=True))
                if not LOCAL:
                    _token()
                st.success(f"Conexión correcta. Control_Central responde y tiene {n} firma(s).")
            except Exception as e:
                st.error(f"Falló la conexión: {type(e).__name__}: {e}")
    with st.container(border=True):
        st.markdown("<p class='card-title'>💾 Respaldo de Control_Central</p>", unsafe_allow_html=True)
        st.caption("Descarga tus firmas y pagos. Guárdalos en un lugar distinto a tu Drive.")
        st.download_button("⬇️ Firmas (CSV)", firmas.to_csv(index=False).encode("utf-8"), "firmas.csv", "text/csv", key="sa_dl_firmas")
        try:
            st.download_button("⬇️ Pagos (CSV)", control_leer("pagos").to_csv(index=False).encode("utf-8"), "pagos.csv", "text/csv", key="sa_dl_pagos")
        except ErrorDatos as e:
            st.warning(str(e))
