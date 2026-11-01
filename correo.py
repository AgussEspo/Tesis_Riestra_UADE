# -*- coding: utf-8 -*-
"""
Envío del informe PDF por correo electrónico (SMTP) - Club Atlético Deportivo Riestra
=====================================================================================

Usa solo la librería estándar de Python (smtplib). No requiere servicios externos: se necesita una
cuenta de correo desde la cual salen los mensajes (por ejemplo, una cuenta de Gmail con "contraseña de
aplicación"). Las credenciales se guardan en los secretos de la app, nunca en el código.

Secretos (ver .streamlit/secrets.toml.example):
    SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD   obligatorios
    SMTP_REMITENTE          opcional: nombre y dirección que ve el destinatario (por defecto, SMTP_USER)
    SMTP_SEGURIDAD          opcional: "starttls" (por defecto, puerto 587), "ssl" (puerto 465) o "ninguna" (solo pruebas locales)
    EMAIL_DOMINIOS_PERMITIDOS  opcional: lista separada por comas; si se define, solo se envía a esos dominios
    EMAIL_MAX_POR_SESION    opcional: tope de envíos por visitante (por defecto 5)
"""
import csv
import os
import re
import smtplib
import ssl
import socket
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, parseaddr

import pandas as pd

MAX_DESTINATARIOS = 5
MAX_ADJUNTO_MB = 20
_RE_EMAIL = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}$")

CAMPOS_ENVIOS = ["fecha_hora", "destinatarios", "informe", "estado", "detalle"]


class ErrorCorreo(Exception):
    """Error con mensaje listo para mostrar a la persona usuaria."""


def leer_configuracion(secreto):
    """`secreto(nombre, defecto)` lee de st.secrets o del entorno. Devuelve el dict de configuración o None si falta algo."""
    cfg = {"host": str(secreto("SMTP_HOST", "")).strip(), "puerto": str(secreto("SMTP_PORT", "587")).strip(),
           "usuario": str(secreto("SMTP_USER", "")).strip(), "clave": str(secreto("SMTP_PASSWORD", "")),
           "remitente": str(secreto("SMTP_REMITENTE", "")).strip(),
           "seguridad": str(secreto("SMTP_SEGURIDAD", "")).strip().lower(),
           "dominios": [d.strip().lower().lstrip("@") for d in str(secreto("EMAIL_DOMINIOS_PERMITIDOS", "")).split(",") if d.strip()],
           "max_sesion": int(secreto("EMAIL_MAX_POR_SESION", 5))}
    if not (cfg["host"] and cfg["usuario"] and cfg["clave"] and cfg["puerto"].isdigit()):
        return None
    cfg["puerto"] = int(cfg["puerto"])
    if not cfg["seguridad"]:
        cfg["seguridad"] = "ssl" if cfg["puerto"] == 465 else "starttls"
    cfg["remitente"] = cfg["remitente"] or cfg["usuario"]
    return cfg


def separar_destinatarios(texto):
    """Devuelve (lista_valida_sin_duplicados, lista_invalida). Acepta comas, punto y coma, espacios o saltos de línea."""
    crudos = [x.strip() for x in re.split(r"[,;\s]+", str(texto or "")) if x.strip()]
    validos, invalidos = [], []
    for x in crudos:
        direccion = parseaddr(x)[1]
        if direccion and _RE_EMAIL.match(direccion) and len(direccion) <= 254:
            if direccion.lower() not in [v.lower() for v in validos]:
                validos.append(direccion)
        else:
            invalidos.append(x)
    return validos, invalidos


def validar_envio(cfg, destinatarios, pdf_bytes):
    if not destinatarios:
        raise ErrorCorreo("Ingresá al menos una dirección de correo.")
    if len(destinatarios) > MAX_DESTINATARIOS:
        raise ErrorCorreo(f"Se pueden enviar hasta {MAX_DESTINATARIOS} destinatarios por vez.")
    if cfg["dominios"]:
        fuera = [d for d in destinatarios if d.rsplit("@", 1)[1].lower() not in cfg["dominios"]]
        if fuera:
            raise ErrorCorreo("Solo se permite enviar a estos dominios: " + ", ".join(cfg["dominios"]) + ". No permitidos: " + ", ".join(fuera) + ".")
    if len(pdf_bytes) > MAX_ADJUNTO_MB * 1024 * 1024:
        raise ErrorCorreo(f"El PDF supera los {MAX_ADJUNTO_MB} MB.")


def armar_mensaje(cfg, destinatarios, asunto, cuerpo, pdf_bytes, nombre_pdf):
    msg = EmailMessage()
    msg["Subject"] = asunto
    msg["From"] = cfg["remitente"] if "@" in cfg["remitente"] and "<" in cfg["remitente"] else formataddr(("Club Atlético Deportivo Riestra", cfg["remitente"]))
    msg["To"] = ", ".join(destinatarios)
    msg.set_content(cuerpo)
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=nombre_pdf)
    return msg


def enviar_pdf(cfg, destinatarios, asunto, cuerpo, pdf_bytes, nombre_pdf):
    """Envía el PDF adjunto. Lanza ErrorCorreo con un mensaje claro si algo falla."""
    validar_envio(cfg, destinatarios, pdf_bytes)
    msg = armar_mensaje(cfg, destinatarios, asunto, cuerpo, pdf_bytes, nombre_pdf)
    try:
        if cfg["seguridad"] == "ssl":
            servidor = smtplib.SMTP_SSL(cfg["host"], cfg["puerto"], timeout=30, context=ssl.create_default_context())
        else:
            servidor = smtplib.SMTP(cfg["host"], cfg["puerto"], timeout=30)
        with servidor:
            servidor.ehlo()
            if cfg["seguridad"] == "starttls":
                servidor.starttls(context=ssl.create_default_context())
                servidor.ehlo()
            if cfg["seguridad"] != "ninguna":
                servidor.login(cfg["usuario"], cfg["clave"])
            rechazados = servidor.send_message(msg, to_addrs=destinatarios)
    except smtplib.SMTPAuthenticationError:
        raise ErrorCorreo("El servidor rechazó el usuario o la contraseña. Con Gmail hace falta una 'contraseña de aplicación', no la clave habitual.")
    except smtplib.SMTPRecipientsRefused:
        raise ErrorCorreo("El servidor rechazó todas las direcciones de destino. Revisá que estén bien escritas.")
    except smtplib.SMTPSenderRefused:
        raise ErrorCorreo("El servidor rechazó la dirección remitente. Revisá SMTP_USER y SMTP_REMITENTE.")
    except (smtplib.SMTPConnectError, smtplib.SMTPServerDisconnected, socket.timeout, socket.gaierror, ConnectionError, TimeoutError):
        raise ErrorCorreo("No se pudo conectar con el servidor de correo. Revisá SMTP_HOST, SMTP_PORT y la conexión a internet.")
    except ssl.SSLError:
        raise ErrorCorreo("Falló la conexión segura con el servidor. Probá con SMTP_SEGURIDAD = \"ssl\" y puerto 465, o \"starttls\" y puerto 587.")
    except smtplib.SMTPException as e:
        raise ErrorCorreo(f"El servidor de correo devolvió un error: {type(e).__name__}.")
    if rechazados:
        raise ErrorCorreo("Se envió, pero el servidor rechazó estas direcciones: " + ", ".join(rechazados) + ".")
    return True


def cuerpo_por_defecto(k, introduccion):
    t = k["totales"]
    return (f"Hola,\n\nSe adjunta el Reporte Financiero Ejecutivo de Club Atlético Deportivo Riestra.\n\n"
            f"Informe: {k['periodo']['tipo'].capitalize()} - {k['periodo']['etiqueta']}\n\n"
            f"{introduccion}\n\n"
            "Este mensaje fue generado automáticamente por el Agente de Reporte Financiero.\n")


def registrar_envio(ruta, destinatarios, informe, estado, detalle=""):
    fila = {"fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "destinatarios": ", ".join(destinatarios),
            "informe": informe, "estado": estado, "detalle": detalle}
    nuevo = not os.path.exists(ruta)
    try:
        with open(ruta, "a", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=CAMPOS_ENVIOS)
            if nuevo:
                w.writeheader()
            w.writerow(fila)
    except OSError:
        pass  # el registro nunca debe romper el envío


def leer_envios(ruta):
    if not os.path.exists(ruta):
        return pd.DataFrame(columns=CAMPOS_ENVIOS)
    return pd.read_csv(ruta)
