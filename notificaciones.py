import smtplib
import ssl
from datetime import datetime
from email.message import EmailMessage
from html import escape

import requests
import streamlit as st


def enviar_email(destinatario, asunto, cuerpo_html, cuerpo_texto=""):
    """Envía un correo por Gmail SMTP. Devuelve (ok, mensaje)."""
    try:
        remitente = st.secrets["gmail_usuario"]
        clave = st.secrets["gmail_app_password"]
        nombre = st.secrets.get("gmail_nombre_remitente", "Insolvencia OS")
        if not destinatario:
            return False, "No hay correo de destino."

        msg = EmailMessage()
        msg["Subject"] = asunto
        msg["From"] = f"{nombre} <{remitente}>"
        msg["To"] = destinatario
        msg.set_content(cuerpo_texto or "Este mensaje requiere un lector de correo con soporte HTML.")
        msg.add_alternative(cuerpo_html, subtype="html")

        contexto = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto, timeout=20) as servidor:
            servidor.login(remitente, clave)
            servidor.send_message(msg)
        return True, "Correo enviado."
    except KeyError as e:
        return False, f"Falta el secreto {e} en Streamlit Secrets."
    except smtplib.SMTPAuthenticationError:
        return False, "Gmail rechazó la contraseña de aplicación."
    except Exception as e:
        print(f"⚠️ Error enviando correo: {type(e).__name__}")
        return False, f"No se pudo enviar el correo ({type(e).__name__})."


def enviar_telegram(texto, chat_id=None):
    """Envía un mensaje por Telegram. Devuelve (ok, mensaje)."""
    try:
        token = st.secrets["telegram_bot_token"]
        destino = chat_id or st.secrets["telegram_chat_id"]
        r = requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            json={"chat_id": destino, "text": texto},
            timeout=15,
        )
        if r.status_code == 200:
            return True, "Telegram enviado."
        return False, f"Telegram respondió con error {r.status_code}."
    except KeyError as e:
        return False, f"Falta el secreto {e} en Streamlit Secrets."
    except requests.RequestException as e:
        print(f"⚠️ Error de red con Telegram: {type(e).__name__}")
        return False, "No se pudo conectar con Telegram."


def enviar_alertas_vencimientos(df_ven, hoy, destino=""):
    """Arma y envía las alertas de vencimientos. Devuelve (ok_mail, msg_mail, ok_tg, msg_tg)."""
    urgentes = []
    if not df_ven.empty:
        for _, rv in df_ven[df_ven["Estado"] == "Activo"].iterrows():
            try:
                fecha = datetime.strptime(str(rv["Fecha_Limite"])[:10], "%Y-%m-%d").date()
                dias = (fecha - hoy.date()).days
                if dias <= 5:
                    urgentes.append((dias, str(rv["Cliente"]), str(rv["Asunto"]), str(rv["Fecha_Limite"])[:10]))
            except ValueError as e:
                print(f"⚠️ Fecha inválida en vencimiento: {e}")
    urgentes.sort(key=lambda x: x[0])

    if urgentes:
        lineas = []
        for dias, cli, asu, fec in urgentes:
            if dias < 0:
                estado = f"VENCIDO hace {abs(dias)} día(s)"
            else:
                estado = f"vence en {dias} día(s)"
            lineas.append(f"• {cli} - {asu} ({estado}, límite {fec})")
        cuerpo = "\n".join(lineas)
    else:
        cuerpo = "Sin vencimientos en los próximos 5 días. (Mensaje de prueba del sistema)"

    texto = f"⚖️ Alertas Insolvencia OS\n\n{cuerpo}"
    html = "<h3>⚖️ Alertas Insolvencia OS</h3><p>" + escape(cuerpo).replace("\n", "<br>") + "</p>"
    if not destino:
        return False, "Tu firma no tiene un correo de avisos configurado. Pídele a soporte que lo registre.", False, "Telegram no aplica."
    ok_mail, msg_mail = enviar_email(destino, "⚖️ Alertas de vencimientos", html, texto)
    return ok_mail, msg_mail, False, "Telegram no se usa para avisos de las firmas."


def notificar_nuevo_vencimiento(cliente, asunto, fecha_limite, destino=""):
    """Avisa por correo a la firma cuando se programa un vencimiento. Nunca rompe el guardado."""
    if not destino:
        return
    try:
        texto = (
            "⚖️ Nuevo vencimiento programado\n\n"
            f"Cliente: {cliente}\n"
            f"Asunto: {asunto}\n"
            f"Fecha límite: {fecha_limite}"
        )
        html = (
            "<h3>⚖️ Nuevo vencimiento programado</h3><p>"
            f"<b>Cliente:</b> {escape(str(cliente))}<br>"
            f"<b>Asunto:</b> {escape(str(asunto))}<br>"
            f"<b>Fecha límite:</b> {escape(str(fecha_limite))}</p>"
        )
        enviar_email(destino, "⚖️ Nuevo vencimiento programado", html, texto)
    except Exception as e:
        print(f"⚠️ No se pudo notificar el nuevo vencimiento: {type(e).__name__}")
