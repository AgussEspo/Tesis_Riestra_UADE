#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Agente de Reporte Financiero Ejecutivo - Club Atlético Deportivo Riestra  (v2)
==============================================================================

Núcleo del agente (sin interfaz). Lo usan:
  - app.py            -> aplicación web (Streamlit)
  - este mismo archivo -> línea de comandos

Etapas:
  1. FUENTE    : lee el Excel desde archivo, Google Sheets, OneDrive o una URL.
  2. CÁLCULO   : código determinista; calcula todos los indicadores del período
                 (mensual, trimestral o anual) y los compara con el período anterior.
  3. ANÁLISIS  : redacción automática con reglas fijas, usando SOLO los indicadores calculados.
  4. REPORTE   : arma el PDF ejecutivo con el template del club.
  5. REGISTRO  : guarda cada informe generado (período y cifras).

Línea de comandos:
    python agente_riestra.py datos.xlsx --tipo mensual --periodo 2026-10
    python agente_riestra.py datos.xlsx --tipo trimestral --periodo 2026-T3
    python agente_riestra.py datos.xlsx --tipo anual --periodo 2026
"""
import argparse
import base64
import csv
import io
import json
import os
import re
import sys
import tempfile
import unicodedata
from datetime import datetime
from xml.sax.saxutils import escape

import pandas as pd

import flujo_caja as fc
import hashlib
import numbers
import planificacion as pl
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (BaseDocTemplate, Frame, HRFlowable, Image,
                                KeepTogether, NextPageTemplate, PageBreak,
                                PageTemplate, Paragraph, Spacer, Table,
                                TableStyle)

try:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    pdfmetrics.registerFont(TTFont("DejaVu", os.path.join(matplotlib.get_data_path(), "fonts", "ttf", "DejaVuSans.ttf")))
    FUENTE_SIMBOLOS = "DejaVu"
except Exception:  # sin la fuente, los círculos del semáforo se dibujan con el símbolo estándar
    FUENTE_SIMBOLOS = "Helvetica"

# ----------------------------------------------------------------------------
# Constantes
# ----------------------------------------------------------------------------
CLUB = "Club Atlético Deportivo Riestra"
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
MESES_ABR = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep",
             "oct", "nov", "dic"]
TIPOS = {"mensual": 1, "trimestral": 3, "anual": 12}
MESES_TENDENCIA = 6

NEGRO = colors.HexColor("#111111")
GRIS = colors.HexColor("#6B6B6B")
GRIS_MEDIO = colors.HexColor("#B5B5B5")
GRIS_CLARO = colors.HexColor("#EFEFEF")
VERDE = colors.HexColor("#1B7F4B")
ROJO = colors.HexColor("#B3261E")
HEX_NEGRO, HEX_GRIS, HEX_VERDE = "#111111", "#9A9A9A", "#1B7F4B"
# Paleta de los gráficos (colores por rubro, consistentes entre informes)
HEX_ING, HEX_EGR, HEX_RES = "#2F6DB5", "#E8833A", "#1B7F4B"
HEX_COBRO, HEX_CUOTA = "#2E9E6B", "#7B5EA7"
PALETA = ["#2F6DB5", "#2E9E6B", "#E8833A", "#7B5EA7", "#4FB3D9", "#D1495B", "#E9B44C", "#5C7080"]
COLOR_CAT = {
    "Sponsor": "#2F6DB5", "LPF": "#4FB3D9", "Entradas": "#2E9E6B", "Merchandising": "#E9B44C",
    "Eventos": "#7B5EA7", "Cuota social": "#5C7080",
    "Sueldos Jugadores": "#D1495B", "Cuerpo Tecnico": "#E8833A", "Servicios": "#E9B44C",
    "Mantenimiento Cancha": "#2E9E6B", "Seguridad": "#7B5EA7", "Inferiores": "#4FB3D9", "Vestuarios": "#5C7080",
}


def color_cat(nombre, i):
    return COLOR_CAT.get(nombre, PALETA[i % len(PALETA)])


COLUMNAS = {
    "Socios": ["ID_Socio", "Estado", "Fecha_Alta"],
    "Pago_Cuota_Social": ["Fecha", "ID_Socio", "Monto"],
    "Ingresos": ["Fecha", "Monto", "Categoria"],
    "Egresos": ["Fecha", "Monto", "Categoria"],
}


# ----------------------------------------------------------------------------
# Formato de números (estilo argentino)
# ----------------------------------------------------------------------------
def _miles(txt):
    return txt.replace(",", "X").replace(".", ",").replace("X", ".")


def ars(n):
    return ("-" if n < 0 else "") + "$\u00a0" + _miles(f"{abs(n):,.0f}")


def ars_m(n):
    return ("-" if n < 0 and round(abs(n) / 1e6, 1) > 0 else "") + "$\u00a0" + _miles(f"{abs(n) / 1e6:,.1f}") + "\u00a0M"


def pct(x, d=1):
    return f"{x * 100:.{d}f}".replace(".", ",") + "%"


def signo_pct(x, d=1):
    return ("+" if x >= 0 else "") + pct(x, d)


def _cap(txt):
    return txt[:1].upper() + txt[1:]


# ----------------------------------------------------------------------------
# 1. FUENTE DE DATOS
# ----------------------------------------------------------------------------
def _norm(txt):
    """Minúsculas, sin tildes y con guiones bajos: 'Categoría Gasto' -> 'categoria_gasto'."""
    t = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "_", t.lower()).strip("_")


ALIAS_COL = {
    "id_socio": "ID_Socio", "idsocio": "ID_Socio", "id": "ID_Socio", "nombre": "Nombre", "apellido": "Apellido",
    "fecha_alta": "Fecha_Alta", "alta": "Fecha_Alta", "estado": "Estado", "fecha": "Fecha", "socio": "Socio",
    "monto": "Monto", "importe": "Monto", "categoria": "Categoria", "rubro": "Categoria",
    "descripcion": "Descripcion", "detalle": "Descripcion", "concepto": "Descripcion", "tipo": "Tipo",
    "actividad": "Actividad", "cuenta": "Cuenta", "saldo": "Saldo", "observacion": "Observacion",
    "fecha_vencimiento": "Fecha_Vencimiento", "vencimiento": "Fecha_Vencimiento", "contraparte": "Contraparte",
    "mes": "Fecha", "periodo": "Fecha",
}
ALIAS_HOJA = {
    "socios": "Socios", "pago_cuota_social": "Pago_Cuota_Social", "pagos_cuota_social": "Pago_Cuota_Social",
    "pagos": "Pago_Cuota_Social", "pagos_de_cuota": "Pago_Cuota_Social", "cuota_social": "Pago_Cuota_Social",
    "ingresos": "Ingresos", "egresos": "Egresos", "gastos": "Egresos",
    "saldos": "Saldos", "saldo": "Saldos", "saldos_iniciales": "Saldos",
    "compromisos": "Compromisos", "vencimientos": "Compromisos",
    "presupuesto_aprobado": "Presupuesto_Aprobado", "presupuesto": "Presupuesto_Aprobado", "ppto": "Presupuesto_Aprobado",
    "presupuesto_anual": "Presupuesto_Aprobado",
}
# Hojas opcionales (flujo de caja) y sus columnas obligatorias cuando están presentes
COLUMNAS_OPCIONALES = {"Saldos": ["Fecha", "Cuenta", "Saldo"], "Compromisos": ["Fecha_Vencimiento", "Tipo", "Monto"],
                       "Presupuesto_Aprobado": ["Fecha", "Tipo", "Categoria", "Monto"]}
# Rubros de egreso que forman la masa salarial (se busca la palabra dentro del nombre de la categoría)
PALABRAS_SALARIO = ("sueldo", "salario", "remuneracion", "cuerpo_tecnico")
SALARIO_VERDE, SALARIO_ROJO = 0.60, 0.80
# Palabras que marcan un movimiento como extraordinario (además de la regla de no recurrencia)
PALABRAS_EXTRA = ("extraordinari", "streamer", "amistoso", "excepcional", "venta de jugador", "pase de jugador",
                  "indemniz", "multa", "donacion", "subsidio", "premio")


def _numero(valor):
    """Convierte '1.234.567', '17,000', '$ 1.234,56' o 1234 en número."""
    if isinstance(valor, (int, float)):
        return valor
    s = str(valor).strip().replace("$", "").replace(" ", "").replace("\u00a0", "")
    if not s:
        return float("nan")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", "") if re.fullmatch(r"-?\d{1,3}(,\d{3})+", s) else s.replace(",", ".")
    elif "." in s and re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def _serie_numerica(s):
    return s if pd.api.types.is_numeric_dtype(s) else s.map(_numero)


def _serie_fecha(s):
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    txt = s.astype(str).str.strip()
    iso = txt.str.match(r"^\d{4}-\d{1,2}-\d{1,2}")
    out = pd.to_datetime(txt.where(iso, None), errors="coerce")
    if (~iso).any():
        out = out.fillna(pd.to_datetime(txt.where(~iso, None), errors="coerce", dayfirst=True))
    return out


def marcar_extraordinarios(df):
    """Agrega 'extraordinario' (bool) y 'motivo_extra' a Ingresos/Egresos.
    Reglas (en este orden):
      1. Si el archivo trae una columna 'Tipo', manda lo que diga (Extraordinario / Ordinario).
      2. Palabras clave en la descripción o el rubro (streamer, amistoso, premio, multa...).
      3. No recurrencia: la descripción aparece en menos de la mitad de los meses con datos
         (solo si hay al menos 3 meses de historia)."""
    df = df.copy()
    desc = df["Descripcion"].fillna("").astype(str) if "Descripcion" in df.columns else pd.Series("", index=df.index)
    texto = (df["Categoria"].astype(str) + " " + desc).map(_norm)
    mes = df["Fecha"].dt.to_period("M")
    n_hist = mes.nunique()
    clave = df["Categoria"].astype(str) + "|" + desc.map(_norm)
    meses_clave = pd.Series(mes.values, index=df.index).groupby(clave).nunique()
    veces = clave.map(meses_clave)

    extra = pd.Series(False, index=df.index)
    motivo = pd.Series("", index=df.index)
    for kw in PALABRAS_EXTRA:
        hit = texto.str.contains(_norm(kw), regex=False)
        motivo = motivo.where(~hit | (motivo != ""), f"Palabra clave: {kw.strip()}")
        extra |= hit
    if n_hist >= 3:
        rare = (veces / n_hist) < 0.5
        motivo = motivo.where(~rare | (motivo != ""), "No recurrente: " + veces.astype(str) + f" de {n_hist} meses")
        extra |= rare
    if "Tipo" in df.columns:
        t = df["Tipo"].astype(str).map(_norm)
        es_x, es_o = t.str.contains("extra"), t.str.contains("ordinari") & ~t.str.contains("extra")
        extra = (extra & ~es_o) | es_x
        motivo = motivo.where(~es_x, "Marcado como extraordinario en el archivo").where(~es_o, "")
    df["extraordinario"] = extra
    df["motivo_extra"] = motivo.where(extra, "")
    return df


def preparar_datos(hojas):
    """Valida hojas/columnas, normaliza formatos, separa lo no operativo y marca movimientos extraordinarios.
    `hojas` = dict nombre -> DataFrame. Hojas opcionales: Saldos y Compromisos (flujo de caja)."""
    hojas = {ALIAS_HOJA.get(_norm(k), str(k).strip()): v for k, v in hojas.items()}
    faltan = [h for h in COLUMNAS if h not in hojas]
    if faltan:
        raise ValueError(f"Faltan estas pestañas/archivos: {', '.join(faltan)}. Se esperan: {', '.join(COLUMNAS)}.")
    for hoja, cols in list(COLUMNAS.items()) + [(h, c) for h, c in COLUMNAS_OPCIONALES.items() if h in hojas]:
        df = hojas[hoja]
        df.columns = [ALIAS_COL.get(_norm(c), str(c).strip()) for c in df.columns]
        falt = [c for c in cols if c not in df.columns]
        if falt:
            raise ValueError(f"En '{hoja}' faltan las columnas: {', '.join(falt)}.")
    socios = hojas["Socios"].copy()
    pagos = hojas["Pago_Cuota_Social"].copy()
    ing = hojas["Ingresos"].copy()
    egr = hojas["Egresos"].copy()
    socios["Fecha_Alta"] = _serie_fecha(socios["Fecha_Alta"])
    for d in (socios, pagos):
        try:
            d["ID_Socio"] = pd.to_numeric(d["ID_Socio"])
        except (ValueError, TypeError):
            d["ID_Socio"] = d["ID_Socio"].astype(str).str.strip()
    for d in (pagos, ing, egr):
        d["Fecha"] = _serie_fecha(d["Fecha"])
        d["Monto"] = _serie_numerica(d["Monto"])
        d.dropna(subset=["Fecha", "Monto"], inplace=True)
    avisos = []
    # Actividad y cuenta de cada movimiento (si faltan las columnas: todo es operativo y está en el Banco)
    for nombre, d in (("Ingresos", ing), ("Egresos", egr)):
        if "Actividad" not in d.columns:
            d["Actividad"] = fc.ACT_OPERATIVA
        norm = d["Actividad"].map(fc.normalizar_actividad)
        raros = int(norm.isna().sum())
        if raros:
            avisos.append(f"En '{nombre}' hay {raros} movimientos con una Actividad que no se reconoce "
                          f"(se esperan: {', '.join(fc.ACTIVIDADES)}); se los trató como operativos.")
        d["Actividad"] = norm.fillna(fc.ACT_OPERATIVA)
        d["Cuenta"] = fc.normalizar_cuentas(d["Cuenta"]) if "Cuenta" in d.columns else fc.CUENTA_DEFECTO
    saldos = compromisos = None
    if "Saldos" in hojas:
        saldos = hojas["Saldos"].copy()
        saldos["Fecha"] = _serie_fecha(saldos["Fecha"])
        saldos["Saldo"] = _serie_numerica(saldos["Saldo"])
        saldos = saldos.dropna(subset=["Fecha", "Saldo"]).reset_index(drop=True)
        saldos["Cuenta"] = fc.normalizar_cuentas(saldos["Cuenta"])
        if saldos.empty:
            saldos = None
            avisos.append("La hoja Saldos no tiene filas válidas: no se calcula el flujo de caja.")
    if "Compromisos" in hojas:
        compromisos = hojas["Compromisos"].copy()
        compromisos["Fecha_Vencimiento"] = _serie_fecha(compromisos["Fecha_Vencimiento"])
        compromisos["Monto"] = _serie_numerica(compromisos["Monto"])
        compromisos["Tipo"] = compromisos["Tipo"].map(fc.normalizar_tipo_compromiso)
        n0 = len(compromisos)
        compromisos = compromisos.dropna(subset=["Fecha_Vencimiento", "Monto", "Tipo"]).reset_index(drop=True)
        if len(compromisos) < n0:
            avisos.append(f"Se descartaron {n0 - len(compromisos)} filas de Compromisos con fecha, monto o Tipo (Cobro/Pago) inválidos.")
        act = compromisos["Actividad"].map(fc.normalizar_actividad) if "Actividad" in compromisos.columns else pd.Series(fc.ACT_OPERATIVA, index=compromisos.index)
        compromisos["Actividad"] = act.fillna(fc.ACT_OPERATIVA)
        compromisos["Estado"] = (compromisos["Estado"].map(fc.normalizar_estado_compromiso)
                                 if "Estado" in compromisos.columns else "Confirmado")
        if "Cuenta" in compromisos.columns:
            compromisos["Cuenta"] = fc.normalizar_cuentas(compromisos["Cuenta"])
    ppto = None
    if "Presupuesto_Aprobado" in hojas:
        ppto = hojas["Presupuesto_Aprobado"].copy()
        n0 = len(ppto)
        ppto["Fecha"] = _serie_fecha(ppto["Fecha"])
        ppto["Monto"] = _serie_numerica(ppto["Monto"])
        tn = ppto["Tipo"].map(lambda v: "Ingreso" if _norm(v).startswith("ing") else ("Egreso" if _norm(v).startswith(("egr", "gast")) else None))
        ppto["Tipo"] = tn
        ppto = ppto.dropna(subset=["Fecha", "Monto", "Tipo", "Categoria"]).reset_index(drop=True)
        ppto["Categoria"] = ppto["Categoria"].astype(str).str.strip()
        ppto["mes"] = ppto["Fecha"].dt.to_period("M")
        if len(ppto) < n0:
            avisos.append(f"Se descartaron {n0 - len(ppto)} filas de Presupuesto_Aprobado con fecha, tipo (Ingreso/Egreso), categoría o monto inválidos.")
        if ppto.empty:
            ppto = None
            avisos.append("La hoja Presupuesto_Aprobado no tiene filas válidas: no se compara contra el presupuesto aprobado.")
    if saldos is None and (ing["Actividad"] != fc.ACT_OPERATIVA).any():
        avisos.append("Hay movimientos de inversión, financiación o transferencias, pero falta la hoja Saldos: "
                      "no se puede armar el flujo de caja (el resultado del período sí se calcula).")
    # Solo la actividad operativa forma parte del resultado; el resto va al flujo de caja
    caja = {"ing": ing.copy(), "egr": egr.copy(), "saldos": saldos, "compromisos": compromisos, "avisos": avisos}
    ing = ing[ing["Actividad"] == fc.ACT_OPERATIVA].copy()
    egr = egr[egr["Actividad"] == fc.ACT_OPERATIVA].copy()
    if min(len(pagos), len(ing), len(egr)) == 0:
        raise ValueError("Alguna de las tablas de movimientos no tiene datos válidos (revisá fechas y montos).")
    return {"socios": socios, "pagos": pagos, "ing": marcar_extraordinarios(ing), "egr": marcar_extraordinarios(egr), "caja": caja,
            "presupuesto_aprobado": ppto}


def huella_datos(datos):
    """Huella (SHA-256 abreviada) del contenido normalizado de todas las tablas usadas: permite saber con qué versión de los
    datos se generó cada informe. No depende del orden de filas ni de columnas ni del formato del archivo de origen."""
    caja = datos.get("caja") or {}
    tablas = [("socios", datos["socios"]), ("pagos", datos["pagos"]), ("ingresos", caja.get("ing", datos["ing"])),
              ("egresos", caja.get("egr", datos["egr"])), ("saldos", caja.get("saldos")), ("compromisos", caja.get("compromisos")),
              ("presupuesto", datos.get("presupuesto_aprobado"))]
    h, filas = hashlib.sha256(), {}
    for nombre, df in tablas:
        if df is None:
            continue
        d = df[sorted(c for c in df.columns if c not in ("extraordinario", "motivo_extra", "mes"))].copy()
        for c in d.columns:
            d[c] = d[c].map(lambda v: "" if pd.isna(v) else (f"{float(v):.2f}" if isinstance(v, numbers.Real) and not isinstance(v, bool) else str(v)))
        d = d.sort_values(list(d.columns)).reset_index(drop=True)
        h.update(nombre.encode() + d.to_csv(index=False).encode("utf-8"))
        filas[nombre] = int(len(d))
    return {"huella": h.hexdigest()[:12].upper(), "filas": filas}


def cargar_excel(fuente):
    """`fuente`: ruta, archivo subido o BytesIO."""
    return preparar_datos(pd.read_excel(fuente, sheet_name=None))


def _leer_csv(contenido):
    texto = None
    for enc in ("utf-8-sig", "latin-1"):
        try:
            texto = contenido.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    muestra = texto[:4096]
    try:
        sep = csv.Sniffer().sniff(muestra, delimiters=",;\t|").delimiter
    except csv.Error:
        sep = ";" if muestra.count(";") > muestra.count(",") else ","
    return pd.read_csv(io.StringIO(texto), sep=sep, dtype=str, keep_default_na=False)


def _tabla_de_csv(nombre, df):
    n = _norm(os.path.splitext(os.path.basename(nombre))[0])
    for clave, tabla in (("presupuesto", "Presupuesto_Aprobado"), ("ppto", "Presupuesto_Aprobado"), ("saldo", "Saldos"), ("compromiso", "Compromisos"), ("vencimiento", "Compromisos"),
                         ("ingreso", "Ingresos"), ("egreso", "Egresos"), ("gasto", "Egresos"),
                         ("pago", "Pago_Cuota_Social"), ("cuota", "Pago_Cuota_Social"), ("socio", "Socios")):
        if clave in n:
            return tabla
    cols = {ALIAS_COL.get(_norm(c), c) for c in df.columns}
    if {"Estado", "Fecha_Alta"} <= cols:
        return "Socios"
    if {"ID_Socio", "Monto", "Fecha"} <= cols:
        return "Pago_Cuota_Social"
    raise ValueError(f"No pude identificar qué tabla es '{nombre}'. Nombrá el archivo con la palabra "
                     f"'socios', 'pagos', 'ingresos' o 'egresos' (y, de forma opcional, 'saldos', 'compromisos' o 'presupuesto').")


def cargar_archivos(items):
    """`items`: lista de (nombre_archivo, bytes). Acepta un .xlsx, o varios .csv (uno por tabla)."""
    hojas = {}
    for nombre, contenido in items:
        ext = os.path.splitext(nombre)[1].lower()
        if ext in (".xlsx", ".xlsm", ".xls"):
            if len(items) > 1:
                raise ValueError("Subí un solo Excel, o bien los archivos CSV (no ambos a la vez).")
            return cargar_excel(io.BytesIO(contenido))
        if ext != ".csv":
            raise ValueError(f"Formato no soportado: '{nombre}'. Usá .xlsx o .csv.")
        df = _leer_csv(contenido)
        tabla = _tabla_de_csv(nombre, df)
        if tabla in hojas:
            raise ValueError(f"Hay dos archivos para la tabla '{tabla}'.")
        hojas[tabla] = df
    return preparar_datos(hojas)


def cargar_rutas(rutas):
    """Rutas a archivos o a una carpeta con los CSV."""
    items = []
    for r in rutas:
        if os.path.isdir(r):
            items += [(f, open(os.path.join(r, f), "rb").read()) for f in sorted(os.listdir(r))
                      if f.lower().endswith((".csv", ".xlsx"))]
        else:
            items.append((r, open(r, "rb").read()))
    return cargar_archivos(items)


def _bytes_a_datos(contenido):
    if not contenido[:2] == b"PK":  # los .xlsx son archivos zip
        raise ValueError("El enlace no devolvió un archivo Excel. Revisá que esté compartido "
                         "como 'Cualquier persona con el enlace puede ver'.")
    return cargar_excel(io.BytesIO(contenido))


def _get(url):
    import requests
    try:
        return requests.get(url, timeout=30, allow_redirects=True,
                            headers={"User-Agent": "Mozilla/5.0"})
    except Exception as e:
        raise ValueError(f"No se pudo conectar ({type(e).__name__}). Revisá la conexión a internet.")


def cargar_google_sheets(url):
    m = re.search(r"/spreadsheets/d/([\w-]+)", url)
    if not m:
        raise ValueError("El enlace no parece ser de Google Sheets.")
    r = _get(f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=xlsx")
    if r.status_code != 200 or "text/html" in r.headers.get("content-type", ""):
        raise ValueError("No se pudo leer la hoja de Google. Compartila como 'Cualquier persona "
                         "con el enlace: Lector'. Si es un .xlsx subido a Drive, abrilo y usá "
                         "Archivo > Guardar como hoja de cálculo de Google.")
    return _bytes_a_datos(r.content)


def cargar_onedrive(url):
    """Enlaces públicos de OneDrive personal (1drv.ms / onedrive.live.com)."""
    codificado = "u!" + base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
    r = _get(f"https://api.onedrive.com/v1.0/shares/{codificado}/root/content")
    if r.status_code != 200:
        raise ValueError("No se pudo descargar desde OneDrive. Verificá que el enlace sea público "
                         "('Cualquier persona con el enlace puede ver') o subí el archivo manualmente.")
    return _bytes_a_datos(r.content)


def cargar_desde_url(url):
    url = url.strip()
    if "docs.google.com/spreadsheets" in url:
        return cargar_google_sheets(url)
    if any(h in url for h in ("1drv.ms", "onedrive.live.com", "onedrive.com")):
        return cargar_onedrive(url)
    r = _get(url)
    if r.status_code != 200:
        raise ValueError(f"No se pudo descargar el archivo (código {r.status_code}).")
    return _bytes_a_datos(r.content)


# ----------------------------------------------------------------------------
# 2. PERÍODOS Y CÁLCULO
# ----------------------------------------------------------------------------
def rango_datos(datos):
    fechas = [datos[k]["Fecha"] for k in ("pagos", "ing", "egr")]
    return min(f.min() for f in fechas).normalize(), max(f.max() for f in fechas).normalize()


def periodos_disponibles(datos):
    """Devuelve {'mensual': [(etiqueta, clave)...], ...} con lo más reciente primero."""
    fmin, fmax = rango_datos(datos)
    meses = pd.period_range(fmin.to_period("M"), fmax.to_period("M"), freq="M")
    out = {"mensual": [], "trimestral": [], "anual": []}
    for p in meses:
        out["mensual"].append((f"{_cap(MESES[p.month - 1])} {p.year}", f"{p.year}-{p.month:02d}"))
        q = (p.month - 1) // 3 + 1
        clave_q = f"{p.year}-T{q}"
        if clave_q not in [c for _, c in out["trimestral"]]:
            out["trimestral"].append((etiqueta_periodo("trimestral", clave_q), clave_q))
        if str(p.year) not in [c for _, c in out["anual"]]:
            out["anual"].append((etiqueta_periodo("anual", str(p.year)), str(p.year)))
    return {t: v[::-1] for t, v in out.items()}


def rango_periodo(tipo, clave):
    """Devuelve (desde, hasta, n_meses_esperados) nominales del período."""
    if tipo == "mensual":
        y, m = (int(x) for x in clave.split("-"))
        desde = pd.Timestamp(y, m, 1)
    elif tipo == "trimestral":
        y, q = clave.upper().split("-T")
        desde = pd.Timestamp(int(y), 3 * (int(q) - 1) + 1, 1)
    elif tipo == "anual":
        desde = pd.Timestamp(int(clave), 1, 1)
    else:
        raise ValueError(f"Tipo de período inválido: {tipo}")
    n = TIPOS[tipo]
    return desde, desde + pd.offsets.MonthEnd(n), n


def etiqueta_periodo(tipo, clave):
    desde, hasta, _ = rango_periodo(tipo, clave)
    if tipo == "mensual":
        return f"{_cap(MESES[desde.month - 1])} {desde.year}"
    if tipo == "trimestral":
        return f"T{(desde.month - 1) // 3 + 1} {desde.year} ({MESES_ABR[desde.month - 1]}–{MESES_ABR[hasta.month - 1]})"
    return f"Año {desde.year}"


def _filtrar(df, desde, hasta):
    return df[(df["Fecha"] >= desde) & (df["Fecha"] < hasta + pd.Timedelta(days=1))].copy()


def _totales(datos, desde, hasta):
    """Totales de un rango; None si no hay datos (sirve para el período anterior)."""
    i, e, p = (_filtrar(datos[k], desde, hasta) for k in ("ing", "egr", "pagos"))
    if len(i) + len(e) + len(p) == 0:
        return None
    ing_t = int(i["Monto"].sum() + p["Monto"].sum())
    egr_t = int(e["Monto"].sum())
    return {"ingresos": ing_t, "egresos": egr_t, "resultado": ing_t - egr_t,
            "margen": (ing_t - egr_t) / ing_t if ing_t else 0}


def _metricas_mes(datos, per):
    socios = datos["socios"]
    ing, egr, pagos = datos["ing"], datos["egr"], datos["pagos"]
    desde, hasta = per.start_time, per.end_time.normalize()  # _filtrar ya incluye todo el último día
    i_m = int(_filtrar(ing, desde, hasta)["Monto"].sum())
    e_m = int(_filtrar(egr, desde, hasta)["Monto"].sum())
    pm = _filtrar(pagos, desde, hasta)
    c_m = int(pm["Monto"].sum())
    base = socios[socios["Fecha_Alta"] <= hasta]
    pagantes = int(pm[pm["ID_Socio"].isin(base["ID_Socio"])]["ID_Socio"].nunique())
    tot = i_m + c_m
    return {
        "mes": str(per), "etiqueta": f"{MESES_ABR[per.month - 1]}-{str(per.year)[2:]}",
        "nombre": f"{MESES[per.month - 1]} {per.year}",
        "ingresos": tot, "cuotas": c_m, "egresos": e_m, "resultado": tot - e_m,
        "margen": (tot - e_m) / tot if tot else 0,
        "pagantes": pagantes, "base_socios": int(len(base)),
        "tasa_cobro": pagantes / len(base) if len(base) else 0,
        "cuota_prom": float(pm["Monto"].mean()) if len(pm) else 0.0,
    }


def estado_cuenta_socios(datos, hasta=None):
    """Estado de cuenta de TODOS los socios al cierre de `hasta` (por defecto, el último dato).

    Criterio (estimado, porque no hay una cuota fija en los datos):
      - Cada socio debe una cuota por mes, desde el mes de su alta (o desde el primer mes con datos).
      - Mes impago = mes sin ningún pago registrado.
      - Deuda estimada = suma, por cada mes impago, de la cuota promedio que pagó el club ese mes.
    """
    socios, pagos = datos["socios"], datos["pagos"]
    fmin, fmax = rango_datos(datos)
    hasta = min(pd.Timestamp(hasta), fmax) if hasta is not None else fmax
    pagos = pagos[pagos["Fecha"] < hasta.normalize() + pd.Timedelta(days=1)].copy()
    pagos["mes"] = pagos["Fecha"].dt.to_period("M")
    meses = pd.period_range(fmin.to_period("M"), hasta.to_period("M"), freq="M")
    cuota_mes = pagos.groupby("mes")["Monto"].mean()
    cuota_def = float(pagos["Monto"].mean()) if len(pagos) else 0.0
    meses_pagados = pagos.groupby("ID_Socio")["mes"].agg(lambda x: set(x))
    ultimo = pagos.groupby("ID_Socio")["Fecha"].max()
    total_pagado = pagos.groupby("ID_Socio")["Monto"].sum()

    filas = []
    for _, r in socios.iterrows():
        sid = r["ID_Socio"]
        alta = r["Fecha_Alta"]
        inicio = max(meses[0], alta.to_period("M")) if pd.notna(alta) else meses[0]
        esperados = [m for m in meses if m >= inicio]
        pag = meses_pagados.get(sid, set())
        impagos = [m for m in esperados if m not in pag]
        deuda = sum(float(cuota_mes.get(m, cuota_def)) for m in impagos)
        filas.append({
            "ID_Socio": sid, "Socio": f"{r.get('Nombre', '')} {r.get('Apellido', '')}".strip(),
            "Estado": r["Estado"], "Fecha_Alta": alta,
            "Meses_esperados": len(esperados), "Meses_pagados": len(esperados) - len(impagos),
            "Meses_impagos": len(impagos), "Deuda_estimada": int(round(deuda)),
            "Detalle_impagos": ", ".join(f"{MESES_ABR[m.month - 1]}-{str(m.year)[2:]}" for m in impagos),
            "Ultimo_pago": ultimo.get(sid, pd.NaT), "Total_pagado": int(total_pagado.get(sid, 0)),
        })
    df = pd.DataFrame(filas)
    return df.sort_values(["Deuda_estimada", "Meses_impagos", "ID_Socio"], ascending=[False, False, True]).reset_index(drop=True)


def _resumen_deudores(datos, hasta, top=5, anonimo=False):
    df = estado_cuenta_socios(datos, hasta)
    fmin, _ = rango_datos(datos)
    con = df[df["Deuda_estimada"] > 0]
    total = int(df["Deuda_estimada"].sum())
    primeros = con.head(top)
    return {
        "al_cierre_de": pd.Timestamp(hasta).strftime("%Y-%m-%d"),
        "calculada_desde": f"{MESES[fmin.month - 1]} {fmin.year}",
        "total_socios": int(len(df)), "socios_con_deuda": int(len(con)),
        "deuda_total": total, "anonimo": bool(anonimo),
        "deuda_morosos": int(df.loc[df["Estado"] == "Moroso", "Deuda_estimada"].sum()),
        "deuda_activos": int(df.loc[df["Estado"] == "Activo", "Deuda_estimada"].sum()),
        "participacion_top": float(primeros["Deuda_estimada"].sum() / total) if total else 0.0,
        "top": [{
            "puesto": i + 1, "id": int(r["ID_Socio"]) if str(r["ID_Socio"]).isdigit() else str(r["ID_Socio"]),
            "socio": "" if anonimo else r["Socio"], "estado": r["Estado"], "meses_impagos": int(r["Meses_impagos"]),
            "deuda": int(r["Deuda_estimada"]),
            "ultimo_pago": r["Ultimo_pago"].strftime("%d/%m/%Y") if pd.notna(r["Ultimo_pago"]) else "sin pagos",
        } for i, (_, r) in enumerate(primeros.iterrows())],
    }


def _lista_extra(df, tipo):
    out = []
    for _, r in df[df["extraordinario"]].sort_values("Fecha").iterrows():
        out.append({"fecha": r["Fecha"].strftime("%d/%m/%Y"), "tipo": tipo, "rubro": str(r["Categoria"]),
                    "descripcion": str(r["Descripcion"]) if "Descripcion" in df.columns else "",
                    "monto": int(r["Monto"]), "motivo": r["motivo_extra"]})
    return out


def _resumen_extraordinarios(ing, egr, resultado):
    li, le = _lista_extra(ing, "Ingreso"), _lista_extra(egr, "Egreso")
    ti, te = sum(x["monto"] for x in li), sum(x["monto"] for x in le)
    neto = ti - te
    res_ord = resultado - neto
    return {"ingresos": li, "egresos": le, "total_ingresos": ti, "total_egresos": te, "neto": neto,
            "cantidad": len(li) + len(le), "resultado_ordinario": res_ord,
            "peso_en_resultado": neto / resultado if resultado else 0.0}


# ----------------------------------------------------------------------------
# PLANIFICACIÓN: proyección, presupuesto y semáforo
# ----------------------------------------------------------------------------
METODOS = {"tendencia": "tendencia lineal", "promedio3": "promedio de los últimos 3 meses"}
ESTADOS = ("verde", "amarillo", "rojo")


def series_ordinarias(datos, hasta):
    """Series mensuales ORDINARIAS (sin extraordinarios) hasta `hasta`.
    Devuelve (ingresos_por_rubro, egresos_por_rubro, nota). El mes en curso se excluye si está muy incompleto."""
    fmin, _ = rango_datos(datos)
    hasta = pd.Timestamp(hasta)
    ultimo, nota = hasta.to_period("M"), ""
    if hasta.day < 20 and hasta < hasta + pd.offsets.MonthEnd(0):
        ultimo = ultimo - 1
        nota = f"Se excluyó {pl.nombre_mes(ultimo + 1)} por estar incompleto."
    meses = pd.period_range(fmin.to_period("M"), ultimo, freq="M")
    corte = ultimo.end_time

    def tabla(df):
        d = df[(df["Fecha"] <= corte) & ~df["extraordinario"]].copy()
        d["mes"] = d["Fecha"].dt.to_period("M")
        return d.groupby(["mes", "Categoria"])["Monto"].sum().unstack(fill_value=0.0).reindex(meses, fill_value=0.0)

    ing, egr = tabla(datos["ing"]), tabla(datos["egr"])
    pg = datos["pagos"][datos["pagos"]["Fecha"] <= corte].copy()
    pg["mes"] = pg["Fecha"].dt.to_period("M")
    ing["Cuota social"] = pg.groupby("mes")["Monto"].sum().reindex(meses, fill_value=0.0)
    return ing, egr, nota


def calcular_planificacion(datos, hasta, desde=None, extra=None, metodo="tendencia", meses_presupuesto=12):
    """Proyección (3/6/12 + un horizonte opcional) y presupuesto de los próximos meses.
    Si se indica `desde` (inicio del período del informe) y hay historial anterior, calcula también la
    ejecución: presupuesto armado con datos previos vs. lo real del período."""
    if metodo not in METODOS:
        metodo = "tendencia"
    ing_r, egr_r, nota = series_ordinarias(datos, hasta)
    horizontes = list(pl.HORIZONTES)
    extra = int(extra) if extra else 0
    if extra and extra not in horizontes:
        horizontes.append(extra)
    ing_s, egr_s = ing_r.sum(axis=1), egr_r.sum(axis=1)
    proy = pl.proyectar(ing_s, egr_s, horizontes, metodo)
    out = {"proyeccion": None, "presupuesto": None, "ejecucion": None, "nota_mes_incompleto": nota}
    if proy:
        proy["extra"] = extra if extra and extra not in pl.HORIZONTES else None
        proy["metodo_txt"] = METODOS[metodo]
        proy["historial"] = [{"etiqueta": pl.etiqueta_mes(p), "nombre": pl.nombre_mes(p), "ingresos": float(ing_s[p]),
                              "egresos": float(egr_s[p]), "resultado": float(ing_s[p] - egr_s[p])} for p in ing_r.index]
        out["proyeccion"] = proy
    h = int(meses_presupuesto)
    pres = pl.presupuesto(ing_r, egr_r, h)
    if pres:
        ult = ing_r.index[-1]
        pres["periodos"] = [pl.nombre_mes(ult + (j + 1)) for j in range(h)]
        pres["etiquetas"] = [pl.etiqueta_mes(ult + (j + 1)) for j in range(h)]
        out["presupuesto"] = pres
        if h >= 6:
            tri = []
            for a in range(0, h - h % 3, 3):
                ii, ee = sum(pres["serie_ingresos"][a:a + 3]), sum(pres["serie_egresos"][a:a + 3])
                tri.append({"nombre": f"{pres['etiquetas'][a]} a {pres['etiquetas'][a + 2]}", "ingresos": ii, "egresos": ee,
                            "reserva": ee * pl.RESERVA, "resultado": ii - ee - ee * pl.RESERVA})
            pres["trimestres"] = tri
    if desde is not None:
        p0 = pd.Timestamp(desde).to_period("M")
        antes_i, antes_e = ing_r[ing_r.index < p0], egr_r[egr_r.index < p0]
        real_i, real_e = ing_r[ing_r.index >= p0], egr_r[egr_r.index >= p0]
        if len(antes_i) >= pl.MIN_MESES and len(real_i) >= 1:
            b = pl.presupuesto(antes_i, antes_e, len(real_i))
            if b:
                out["ejecucion"] = pl.ejecucion(b, real_i.sum().to_dict(), real_e.sum().to_dict())
                out["ejecucion"]["origen"] = "previo"
        aprobada = ejecucion_aprobada(datos, desde, hasta)
        if aprobada:  # si hay presupuesto aprobado, manda sobre el armado con datos previos
            out["ejecucion"] = aprobada
    return out


def ejecucion_aprobada(datos, desde, hasta):
    """Compara el presupuesto aprobado (hoja Presupuesto_Aprobado) con lo real ORDINARIO del período, mes por mes con presupuesto.
    Devuelve None si no hay hoja o ningún mes del período tiene presupuesto."""
    pa = datos.get("presupuesto_aprobado")
    if pa is None:
        return None
    hasta = pd.Timestamp(hasta)
    ultimo = hasta.to_period("M")
    if hasta.day < 20 and hasta < hasta + pd.offsets.MonthEnd(0):
        ultimo = ultimo - 1
    desde = max(pd.Timestamp(desde), rango_datos(datos)[0])  # el período nominal puede empezar antes del primer dato
    meses = list(pd.period_range(desde.to_period("M"), ultimo, freq="M"))
    cubiertos = [m for m in meses if (pa["mes"] == m).any()]
    if not cubiertos:
        return None
    b = pa[pa["mes"].isin(cubiertos)]
    nombre = {}
    for cat in b["Categoria"]:
        nombre.setdefault(_norm(cat), cat)

    def real(df):
        d = df[df["Fecha"].dt.to_period("M").isin(cubiertos) & ~df["extraordinario"]]
        out = {}
        for cat, v in d.groupby("Categoria")["Monto"].sum().items():
            out[nombre.get(_norm(cat), str(cat))] = out.get(nombre.get(_norm(cat), str(cat)), 0.0) + float(v)
        return out
    real_i, real_e = real(datos["ing"]), real(datos["egr"])
    pg = datos["pagos"]
    cuota = float(pg[pg["Fecha"].dt.to_period("M").isin(cubiertos)]["Monto"].sum())
    if cuota:
        real_i[nombre.get(_norm("Cuota social"), "Cuota social")] = cuota

    def filas(tipo):
        g = b[b["Tipo"] == tipo].groupby("Categoria")["Monto"].sum()
        acc = {}
        for cat, v in g.items():
            acc[nombre[_norm(cat)]] = acc.get(nombre[_norm(cat)], 0.0) + float(v)
        return [{"rubro": c, "total": v} for c, v in sorted(acc.items(), key=lambda kv: -kv[1])]
    fi, fe = filas("Ingreso"), filas("Egreso")
    ti, te = sum(f["total"] for f in fi), sum(f["total"] for f in fe)
    ej = pl.ejecucion({"meses": len(cubiertos), "filas_ingresos": fi, "filas_egresos": fe, "total_ingresos": ti,
                       "total_egresos": te, "resultado_operativo": ti - te}, real_i, real_e)
    ej.update({"origen": "aprobado", "meses_cubiertos": len(cubiertos), "meses_periodo": len(meses),
               "rango": (pl.etiqueta_mes(cubiertos[0]) + (f" a {pl.etiqueta_mes(cubiertos[-1])}" if len(cubiertos) > 1 else ""))})
    return ej


def _estado(v, ok, mal, mayor_mejor=True):
    if mayor_mejor:
        return "verde" if v >= ok else ("rojo" if v < mal else "amarillo")
    return "verde" if v <= ok else ("rojo" if v > mal else "amarillo")


def calcular_semaforo(k):
    """Semáforo de hasta 16 indicadores con umbrales explícitos (verde = bien, amarillo = regular, rojo = mal).
    Los indicadores que no se pueden calcular por falta de datos simplemente no se incluyen."""
    t, s, c, x, d = k["totales"], k["socios"], k["comparacion"], k["extraordinarios"], k["deudores"]
    pr, ej = k.get("proyeccion"), k.get("ejecucion")
    it = []

    def add(grupo, nombre, valor, estado, criterio):
        it.append({"grupo": grupo, "indicador": nombre, "valor": valor, "estado": estado, "criterio": criterio})

    cob = t["ingresos"] / t["egresos"] if t["egresos"] else 0
    ic, ec = list(k["ingresos_cat"].items()), list(k["egresos_cat"].items())
    add("Resultados", "Margen del período", pct(t["margen"]), _estado(t["margen"], 0.20, 0.05), "Verde ≥ 20% · Rojo < 5%")
    add("Resultados", "Cobertura de egresos", f"{cob:.1f}".replace(".", ",") + " veces", _estado(cob, 1.2, 1.0),
        "Verde ≥ 1,2 veces · Rojo < 1,0")
    if c:
        add("Resultados", "Ingresos vs período anterior", signo_pct(c["var_ingresos"]), _estado(c["var_ingresos"], 0.0, -0.05),
            "Verde ≥ 0% · Rojo < -5%")
        dif = c["var_egresos"] - c["var_ingresos"]
        add("Resultados", "Egresos vs ingresos (ritmo)", f"{signo_pct(c['var_egresos'])} vs {signo_pct(c['var_ingresos'])}",
            _estado(dif, 0.0, 0.05, False), "Verde: egresos crecen igual o menos · Rojo: más de 5 pts por encima")
    add("Estructura", "Concentración de ingresos", f"{ic[0][0]}: {pct(ic[0][1] / t['ingresos'])}",
        _estado(ic[0][1] / t["ingresos"], 0.30, 0.50, False), "Verde < 30% · Rojo > 50% en un solo rubro")
    add("Estructura", "Concentración de egresos", f"{ec[0][0]}: {pct(ec[0][1] / t['egresos'])}",
        _estado(ec[0][1] / t["egresos"], 0.40, 0.60, False), "Verde < 40% · Rojo > 60% en un solo rubro")
    ms = k.get("masa_salarial")
    if ms:
        add("Estructura", "Masa salarial sobre ingresos", pct(ms["sobre_ingresos"], 0), _estado(ms["sobre_ingresos"], SALARIO_VERDE, SALARIO_ROJO, False),
            "Sueldos y cuerpo técnico sobre ingresos ordinarios: Verde ≤ 60% · Rojo > 80%")
    pe = abs(x["peso_en_resultado"]) if x["cantidad"] else 0.0
    add("Estructura", "Dependencia de extraordinarios", pct(pe), _estado(pe, 0.10, 0.25, False),
        "Verde ≤ 10% del resultado · Rojo > 25%")
    add("Socios", "Morosidad del padrón", pct(s["tasa_morosidad"], 0), _estado(s["tasa_morosidad"], 0.10, 0.20, False),
        "Verde ≤ 10% · Rojo > 20%")
    add("Socios", "Tasa de cobro de cuotas", pct(s["tasa_cobro_promedio"], 0), _estado(s["tasa_cobro_promedio"], 0.90, 0.75),
        "Verde ≥ 90% · Rojo < 75%")
    add("Socios", "Cuota social sobre egresos", pct(s["cuotas_sobre_egresos"]), _estado(s["cuotas_sobre_egresos"], 0.25, 0.10),
        "Verde ≥ 25% · Rojo < 10%")
    rd = d["deuda_total"] / d["cobrado_historico"] if d.get("cobrado_historico") else 0.0
    add("Socios", "Deuda estimada sobre lo cobrado", pct(rd), _estado(rd, 0.10, 0.25, False), "Verde ≤ 10% · Rojo > 25%")
    if pr:
        h3 = pr["horizontes"][0]
        est = "rojo" if h3["resultado"] < 0 else ("amarillo" if h3["res_pesimista"] < 0 or h3["margen"] < 0.10 else "verde")
        add("Futuro", "Resultado proyectado (3 meses)", f"{ars_m(h3['resultado'])}  (margen {pct(h3['margen'], 0)})", est,
            "Verde: positivo aun en el escenario pesimista · Rojo: resultado negativo")
    if ej:
        add("Futuro", "Ejecución presupuestaria", signo_pct(ej["desvio_resultado_pct"]), _estado(ej["desvio_resultado_pct"], -0.05, -0.15),
            ("Resultado real vs presupuesto aprobado: Verde ≥ -5% · Rojo < -15%" if ej.get("origen") == "aprobado"
         else "Resultado real vs presupuesto previo: Verde ≥ -5% · Rojo < -15%"))
    cj = k.get("caja")
    if cj:
        cob_c = cj["periodo"].get("cobertura")
        if cob_c is not None:
            add("Caja", "Cobertura de caja", f"{cob_c:.1f}".replace(".", ",") + " meses", _estado(cob_c, fc.COBERTURA_VERDE, fc.COBERTURA_ROJO),
                "Saldo final sobre el egreso mensual promedio (con cuotas de deuda): Verde ≥ 3 meses · Rojo < 1 mes")
        pj = cj.get("proyeccion")
        if pj:
            h6 = next((x for x in pj["horizontes"] if x["meses"] == 6), pj["horizontes"][-1])
            est = "rojo" if h6["saldo_minimo"] < 0 else ("amarillo" if h6["saldo_minimo_pesimista"] < 0 else "verde")
            add("Caja", f"Saldo de caja proyectado ({h6['meses']} meses)", f"{ars_m(h6['saldo'])}  (mínimo {ars_m(h6['saldo_minimo'])})", est,
                "Verde: nunca negativo, ni en el escenario pesimista · Rojo: saldo negativo en el escenario base")

    ev = it
    n = {e: sum(1 for i in it if i["estado"] == e) for e in ESTADOS}
    score = sum({"verde": 2, "amarillo": 1, "rojo": 0}[i["estado"]] for i in ev) / len(ev) if ev else 0
    grave = it[0]["estado"] == "rojo" or it[1]["estado"] == "rojo"
    general = "rojo" if grave or score < 0.8 else ("verde" if score >= 1.6 and n["rojo"] == 0 else "amarillo")
    return {"items": it, "conteo": n, "puntaje": score, "general": general, "evaluados": len(ev)}


def _no_operativos(datos, desde, hasta):
    """Movimientos de inversión, financiación y transferencias del rango: fuera del resultado, dentro del flujo de caja."""
    out = {}
    caja = datos.get("caja")
    if not caja:
        return out
    for clave, df in (("ingresos", caja["ing"]), ("egresos", caja["egr"])):
        d = _filtrar(df, desde, hasta)
        for act, g in d[d["Actividad"] != fc.ACT_OPERATIVA].groupby("Actividad"):
            o = out.setdefault(act, {"ingresos": 0, "egresos": 0, "n": 0})
            o[clave] += int(g["Monto"].sum())
            o["n"] += int(len(g))
    return out


def calcular_caja(k, datos, hasta):
    """Flujo de caja histórico del período y proyectado. None si el archivo no trae la hoja Saldos."""
    caja = datos.get("caja")
    if not caja or caja.get("saldos") is None:
        return None
    h = fc.historico(caja["ing"], caja["egr"], datos["pagos"], caja["saldos"], hasta)
    per = k["periodo"]
    res = fc.resumen_periodo(h, per["desde"], per["hasta"]) if h else None
    if not res:
        return None
    out = {"periodo": res, "avisos": list(caja["avisos"]) + h["avisos"], "cuentas": h["cuentas"], "proyeccion": None,
           "hay_compromisos": caja.get("compromisos") is not None}
    pr = k.get("proyeccion")
    if pr:
        ult = pd.Period(pr["serie"][0]["mes"]) - 1
        previas = [f for f in h["filas"] if pd.Period(f["mes"]) <= ult]
        saldo0 = previas[-1]["saldo_final"] if previas else h["filas"][0]["saldo_inicial"]
        out["proyeccion"] = fc.proyectar(saldo0, ult, pr["serie"], [x["meses"] for x in pr["horizontes"]], caja.get("compromisos"))
    return out


def _completar_planificacion(k, datos, hasta, desde_periodo, extra, metodo, meses_presupuesto):
    pl_ = calcular_planificacion(datos, hasta, desde_periodo, extra, metodo, meses_presupuesto)
    k["proyeccion"], k["presupuesto"], k["ejecucion"] = pl_["proyeccion"], pl_["presupuesto"], pl_["ejecucion"]
    k["nota_mes_incompleto"] = pl_["nota_mes_incompleto"]
    pg = datos["pagos"]
    k["deudores"]["cobrado_historico"] = int(pg[pg["Fecha"] <= pd.Timestamp(hasta) + pd.Timedelta(days=1)]["Monto"].sum())
    k["caja"] = calcular_caja(k, datos, hasta)
    k["semaforo"] = calcular_semaforo(k)
    return k


def calcular_kpis(datos, tipo, clave, proyeccion_extra=None, metodo="tendencia", meses_presupuesto=12, anonimo=False):
    socios = datos["socios"]
    fmin, fmax = rango_datos(datos)
    desde_nom, hasta_nom, n_esp = rango_periodo(tipo, clave)
    desde, hasta = max(desde_nom, fmin), min(hasta_nom, fmax)
    if desde > hasta:
        raise ValueError("No hay datos en el período elegido.")

    ing, egr, pagos = (_filtrar(datos[k], desde, hasta) for k in ("ing", "egr", "pagos"))
    meses = pd.period_range(desde.to_period("M"), hasta.to_period("M"), freq="M")
    parcial = len(meses) < n_esp

    ing_cat = {k: int(v) for k, v in ing.groupby("Categoria")["Monto"].sum().items()}
    cuotas_total = int(pagos["Monto"].sum())
    ing_cat["Cuota social"] = cuotas_total
    egr_cat = {k: int(v) for k, v in egr.groupby("Categoria")["Monto"].sum().items()}
    ing_cat = dict(sorted(ing_cat.items(), key=lambda kv: -kv[1]))
    egr_cat = dict(sorted(egr_cat.items(), key=lambda kv: -kv[1]))
    ing_total, egr_total = sum(ing_cat.values()), sum(egr_cat.values())
    resultado = ing_total - egr_total

    # tendencia: meses del período + contexto previo (para gráficos)
    n_tend = max(MESES_TENDENCIA, len(meses))
    ini_tend = max(meses[-1] - (n_tend - 1), fmin.to_period("M"))
    tendencia = [_metricas_mes(datos, p) for p in pd.period_range(ini_tend, meses[-1], freq="M")]
    en_periodo = {str(p) for p in meses}
    for t in tendencia:
        t["en_periodo"] = t["mes"] in en_periodo
    mensual = [t for t in tendencia if t["en_periodo"]]

    # comparación con el período anterior equivalente
    d_prev = desde_nom - pd.DateOffset(months=n_esp)
    h_prev = d_prev + pd.offsets.MonthEnd(n_esp)
    prev = _totales(datos, d_prev, h_prev)
    comparacion = None
    if prev:
        clave_prev = {"mensual": f"{d_prev.year}-{d_prev.month:02d}",
                      "trimestral": f"{d_prev.year}-T{(d_prev.month - 1) // 3 + 1}",
                      "anual": str(d_prev.year)}[tipo]
        comparacion = dict(prev, etiqueta=etiqueta_periodo(tipo, clave_prev))
        comparacion["var_ingresos"] = ing_total / prev["ingresos"] - 1 if prev["ingresos"] else 0
        comparacion["var_egresos"] = egr_total / prev["egresos"] - 1 if prev["egresos"] else 0

    # etiquetas / frases
    etiqueta = etiqueta_periodo(tipo, clave)
    rango_txt = f"{MESES_ABR[meses[0].month - 1]}–{MESES_ABR[meses[-1].month - 1]}"
    if tipo == "anual" and parcial:
        etiqueta += f" (parcial: {rango_txt})"
    frase = {"mensual": f"durante {MESES[desde.month - 1]} de {desde.year}",
             "trimestral": f"durante el {etiqueta}",
             "anual": f"durante el año {desde.year}" + (f" (período parcial: {rango_txt})" if parcial else "")}[tipo]

    activos = int((socios["Estado"] == "Activo").sum())
    morosos = int((socios["Estado"] == "Moroso").sum())
    ids_morosos = socios.loc[socios["Estado"] == "Moroso", "ID_Socio"]
    chk = pagos.merge(socios[["ID_Socio", "Fecha_Alta"]], on="ID_Socio", how="left")

    k = {
        "periodo": {"tipo": tipo, "clave": clave, "etiqueta": etiqueta, "frase": frase,
                    "desde": desde.strftime("%Y-%m-%d"), "hasta": hasta.strftime("%Y-%m-%d"),
                    "n_meses": len(meses), "parcial": parcial},
        "totales": {"ingresos": int(ing_total), "egresos": int(egr_total),
                    "resultado": int(resultado), "margen": resultado / ing_total if ing_total else 0},
        "comparacion": comparacion,
        "ingresos_cat": ing_cat, "egresos_cat": egr_cat,
        "mensual": mensual, "tendencia": tendencia,
        "socios": {
            "total": int(len(socios)), "activos": activos, "morosos": morosos,
            "tasa_morosidad": morosos / len(socios) if len(socios) else 0,
            "recaudacion_cuotas": cuotas_total,
            "cuota_promedio": float(pagos["Monto"].mean()) if len(pagos) else 0.0,
            "cuotas_sobre_ingresos": cuotas_total / ing_total if ing_total else 0,
            "cuotas_sobre_egresos": cuotas_total / egr_total if egr_total else 0,
            "tasa_cobro_promedio": float(sum(x["tasa_cobro"] for x in mensual) / len(mensual)),
        },
        "extraordinarios": _resumen_extraordinarios(ing, egr, int(resultado)),
        "deudores": _resumen_deudores(datos, hasta, anonimo=anonimo),
        "anonimo": bool(anonimo),
        "trazabilidad": dict(huella_datos(datos), validado_por=""),
        "calidad_datos": {
            "pagos_previos_al_alta": int((chk["Fecha"] < chk["Fecha_Alta"]).sum()),
            "morosos_con_pagos": int(pagos[pagos["ID_Socio"].isin(ids_morosos)]["ID_Socio"].nunique()),
            "cuota_no_esta_en_ingresos": "Cuota social" not in set(ing["Categoria"].astype(str)),
            "no_operativos": _no_operativos(datos, desde, hasta),
            "avisos_carga": list((datos.get("caja") or {}).get("avisos", [])),
        },
    }
    sal = {c: v for c, v in egr_cat.items() if any(p_ in _norm(c) for p_ in PALABRAS_SALARIO)}
    ing_ord = ing_total - k["extraordinarios"]["total_ingresos"]
    k["masa_salarial"] = ({"monto": sum(sal.values()), "rubros": sal, "ingresos_ordinarios": ing_ord,
                           "sobre_ingresos": sum(sal.values()) / ing_ord} if sal and ing_ord > 0 else None)
    if len(meses) == 1:  # informe mensual: detalle de movimientos
        def movs(df):
            cols = [c for c in ("Categoria", "Descripcion") if c in df.columns]
            return [{"fecha": r["Fecha"].strftime("%d/%m"), "rubro": str(r["Categoria"]),
                     "descripcion": str(r["Descripcion"]) if "Descripcion" in cols else "",
                     "monto": int(r["Monto"]), "extra": bool(r.get("extraordinario", False))}
                    for _, r in df.sort_values("Fecha").iterrows()]
        mi = movs(ing)
        mi.append({"fecha": "mes", "rubro": "Cuota social", "descripcion": f"{len(pagos)} pagos de socios", "monto": cuotas_total})
        k["movimientos"] = {"ingresos": mi, "egresos": movs(egr)}
    return _completar_planificacion(k, datos, hasta, desde_nom, proyeccion_extra, metodo, meses_presupuesto)


# ----------------------------------------------------------------------------
# 3. ANÁLISIS (redacción automática)
# ----------------------------------------------------------------------------
def _lista_y(items):
    return items[0] if len(items) == 1 else ", ".join(items[:-1]) + " y " + items[-1]


def etq_socio(r, con_n=False):
    """Nombre del socio para mostrar. En modo anónimo no hay nombre: solo el número de socio."""
    if not r.get("socio"):
        return f"Socio N° {r['id']}"
    return f"{r['socio']} (N° {r['id']})" if con_n else r["socio"]


def analisis_local(k):
    """Redactor de respaldo (sin IA): reglas fijas sobre los mismos indicadores."""
    t, s, c = k["totales"], k["socios"], k["comparacion"]
    m, tend, per = k["mensual"], k["tendencia"], k["periodo"]
    ic, ec = list(k["ingresos_cat"].items()), list(k["egresos_cat"].items())
    sup = t["resultado"] >= 0
    ref = (c["etiqueta"].lower() if per["tipo"] == "mensual" else c["etiqueta"]) if c else ""
    sube = lambda v: "crecieron" if v >= 0 else "cayeron"

    intro = (f"El presente informe resume la situación económico-financiera de {CLUB} {per['frase']}. "
             f"En ese lapso el club obtuvo ingresos por {ars_m(t['ingresos'])} y afrontó egresos por {ars_m(t['egresos'])}, "
             f"lo que dejó un resultado {'superavitario' if sup else 'deficitario'} de {ars_m(abs(t['resultado']))}, "
             f"equivalente a un margen de {pct(t['margen'])} sobre los ingresos.")

    ing = (f"Los ingresos totales fueron de {ars_m(t['ingresos'])}. El rubro principal fue {ic[0][0]}, con {ars_m(ic[0][1])} "
           f"({pct(ic[0][1] / t['ingresos'])} del total), seguido por {ic[1][0]} ({pct(ic[1][1] / t['ingresos'])})")
    ing += (f" y {ic[2][0]} ({pct(ic[2][1] / t['ingresos'])}). " if len(ic) > 2 else ". ")
    if ic[0][1] / t["ingresos"] >= 0.40:
        ing += f"Esto implica una alta concentración de los ingresos en {ic[0][0]}. "
    if c:
        ing += f"Respecto de {ref}, los ingresos {sube(c['var_ingresos'])} {pct(abs(c['var_ingresos']))}."

    egr = (f"Los egresos totales fueron de {ars_m(t['egresos'])}. El rubro de mayor peso fue {ec[0][0]}, con {ars_m(ec[0][1])} "
           f"({pct(ec[0][1] / t['egresos'])} del gasto), seguido por {ec[1][0]} ({pct(ec[1][1] / t['egresos'])}). ")
    if ec[0][1] / t["egresos"] >= 0.50:
        egr += "La estructura de costos depende en buena medida de ese rubro. "
    ms = k.get("masa_salarial")
    if ms:
        egr += (f"La masa salarial ({_lista_y(list(ms['rubros']))}) suma {ars_m(ms['monto'])} y equivale al {pct(ms['sobre_ingresos'], 0)} "
                f"de los ingresos ordinarios. ")
    if c:
        egr += f"Respecto de {ref}, los egresos {sube(c['var_egresos'])} {pct(abs(c['var_egresos']))}."

    cob = str(round(t["ingresos"] / t["egresos"], 1)).replace(".", ",") if t["egresos"] else "-"
    res = (f"El resultado del período fue {'positivo' if sup else 'negativo'} en {ars_m(abs(t['resultado']))}, con un margen de "
           f"{pct(t['margen'])}. Los ingresos cubrieron {cob} veces los egresos. ")
    if len(m) > 1:
        mejor, peor = max(m, key=lambda x: x["resultado"]), min(m, key=lambda x: x["resultado"])
        res += (f"El mejor mes fue {mejor['nombre']} ({ars_m(mejor['resultado'])}) y el de menor resultado fue "
                f"{peor['nombre']} ({ars_m(peor['resultado'])}). ")
    if c:
        res += f"En comparación con {ref}, el resultado pasó de {ars_m(c['resultado'])} a {ars_m(t['resultado'])}."

    soc = (f"El padrón cuenta con {s['total']} socios: {s['activos']} activos y {s['morosos']} morosos "
           f"({pct(s['tasa_morosidad'], 0)} del padrón). La recaudación por cuota social fue de {ars_m(s['recaudacion_cuotas'])}, "
           f"el {pct(s['cuotas_sobre_ingresos'])} de los ingresos. ")
    ini, fin = tend[0], tend[-1]
    if len(tend) > 1 and ini["cuota_prom"]:
        v = fin["cuota_prom"] / ini["cuota_prom"] - 1
        soc += (f"La cuota promedio pasó de {ars(ini['cuota_prom'])} en {ini['nombre']} a {ars(fin['cuota_prom'])} en "
                f"{fin['nombre']} ({signo_pct(v)}), y en promedio abonó el {pct(s['tasa_cobro_promedio'], 0)} de los socios por mes.")
    else:
        soc += f"La cuota promedio fue de {ars(s['cuota_promedio'])} y abonó el {pct(s['tasa_cobro_promedio'], 0)} de los socios."

    d, x = k["deudores"], k["extraordinarios"]
    if d["socios_con_deuda"] == 0:
        deu = f"Al cierre del período ningún socio registra cuotas impagas."
    else:
        top = d["top"]
        nombres = _lista_y([f"{etq_socio(r)} ({ars(r['deuda'])})" for r in top])
        deu = (f"Al cierre del período, {d['socios_con_deuda']} de {d['total_socios']} socios registran cuotas impagas, con una deuda "
               f"estimada total de {ars(d['deuda_total'])}. Los {len(top)} mayores deudores son {nombres} y concentran el "
               f"{pct(d['participacion_top'])} de esa deuda. ")
        n_act = sum(1 for r in top if r["estado"] != "Moroso")
        deu += (f"{n_act} de ellos figuran como activos en el padrón pese a tener cuotas impagas, por lo que conviene revisar su estado. "
                if n_act else "Todos figuran como morosos en el padrón. ")
        deu += f"La deuda es una estimación (meses sin pago por la cuota promedio del club, desde {d['calculada_desde']})."

    if x["cantidad"] == 0:
        extra = "En el período no se registraron ingresos ni egresos extraordinarios, por lo que el resultado refleja la operatoria habitual del club."
    else:
        partes = []
        for lista, nombre, total in ((x["ingresos"], "ingreso", x["total_ingresos"]), (x["egresos"], "egreso", x["total_egresos"])):
            if lista:
                det = _lista_y([f"{r['descripcion']} ({r['fecha']}, {ars_m(r['monto'])})" for r in lista])
                partes.append(f"{len(lista)} {nombre}{'s' if len(lista) > 1 else ''} extraordinario{'s' if len(lista) > 1 else ''} "
                              f"por {ars_m(total)}: {det}")
        extra = ("Se identificaron " + " y ".join(partes) + ". Sin considerarlos, el resultado del período habría sido de "
                 f"{ars_m(x['resultado_ordinario'])} en lugar de {ars_m(t['resultado'])}"
                 + (f" (el {pct(abs(x['peso_en_resultado']))} del resultado proviene de estos movimientos)" if t["resultado"] else "")
                 + ". Al no ser recurrentes, no conviene usarlos como base para proyectar los próximos períodos.")

    sm, pr, pp, ej = k["semaforo"], k.get("proyeccion"), k.get("presupuesto"), k.get("ejecucion")
    nom = {"verde": [], "amarillo": [], "rojo": []}
    for i in sm["items"]:
        if i["estado"] in nom:
            nom[i["estado"]].append(i["indicador"].lower())
    gen = {"verde": "buena", "amarillo": "regular, con focos de atención", "rojo": "crítica"}[sm["general"]]
    sem = (f"De {sm['evaluados']} indicadores evaluados, {sm['conteo']['verde']} están en verde, {sm['conteo']['amarillo']} en amarillo y "
           f"{sm['conteo']['rojo']} en rojo, por lo que la situación general es {gen}. ")
    if nom["verde"]:
        sem += f"Vienen bien: {_lista_y(nom['verde'][:5])}. "
    if nom["amarillo"]:
        sem += f"Requieren seguimiento (amarillo): {_lista_y(nom['amarillo'][:5])}. "
    sem += (f"Están en rojo: {_lista_y(nom['rojo'])}." if nom["rojo"] else "No hay indicadores en rojo.")

    if not pr:
        proy_t = "No hay historial suficiente (se necesitan al menos 3 meses con datos) para proyectar los próximos meses."
    else:
        hz = pr["horizontes"]
        partes = [f"a {h['meses']} meses (hasta {h['hasta']}) un resultado de {ars_m(h['resultado'])} sobre ingresos de {ars_m(h['ingresos'])} y "
                  f"egresos de {ars_m(h['egresos'])}" for h in hz[:3]]
        proy_t = (f"Proyectando solo la operatoria ordinaria (sin extraordinarios) con el método de {pr['metodo_txt']}, se estima {partes[0]}; {partes[1]}; y {partes[2]}. "
                  f"En el escenario pesimista el resultado a 3 meses sería de {ars_m(hz[0]['res_pesimista'])} y en el optimista de {ars_m(hz[0]['res_optimista'])}. ")
        if pr.get("extra"):
            hx = next(h for h in hz if h["meses"] == pr["extra"])
            proy_t += f"Para el horizonte elegido de {hx['meses']} meses se estima un resultado de {ars_m(hx['resultado'])} (confiabilidad {hx['confiabilidad'].lower()}). "
        proy_t += (f"La confiabilidad es {hz[0]['confiabilidad'].lower()} a 3 meses y {hz[2]['confiabilidad'].lower()} a 12, ya que se dispone de solo "
                   f"{pr['meses_historial']} meses de historial; son valores nominales y sin estacionalidad.")

    if not pp:
        pres_t = "No hay historial suficiente para elaborar un presupuesto."
    else:
        c_ = pp["criterio"]
        pres_t = (f"Para los próximos {pp['meses']} meses ({pp['periodos'][0]} a {pp['periodos'][-1]}) se presupuestan ingresos por {ars_m(pp['total_ingresos'])} "
                  f"y egresos por {ars_m(pp['total_egresos'])}, más una reserva de contingencia de {ars_m(pp['reserva'])} ({pct(c_['reserva'], 0)} de los egresos), "
                  f"lo que deja un resultado de {ars_m(pp['resultado_final'])}. Se aplicó un criterio prudente: la tendencia de cada rubro con ingresos "
                  f"{pct(c_['prudencia_ingresos'], 0)} por debajo y egresos {pct(c_['prudencia_egresos'], 0)} por encima. "
                  f"El mayor ingreso presupuestado es {pp['filas_ingresos'][0]['rubro']} ({ars_m(pp['filas_ingresos'][0]['total'])}) y el mayor egreso es "
                  f"{pp['filas_egresos'][0]['rubro']} ({ars_m(pp['filas_egresos'][0]['total'])}). ")
    if ej and ej.get("origen") == "aprobado":
        cob_ = (f" ({ej['meses_cubiertos']} de {ej['meses_periodo']} meses del período tienen presupuesto)" if ej["meses_cubiertos"] < ej["meses_periodo"] else "")
        pres_t += ((" " if pp else "") + f"Frente al presupuesto aprobado{cob_}, el resultado real ordinario del período ({ars_m(ej['real_resultado'])}) "
                   f"{'superó' if ej['desvio_resultado'] >= 0 else 'quedó por debajo de'} lo presupuestado ({ars_m(ej['presupuesto_resultado'])}) "
                   f"en {pct(abs(ej['desvio_resultado_pct']))}. Los ingresos reales fueron {ars_m(ej['real_ingresos'])} contra {ars_m(ej['presupuesto_ingresos'])} "
                   f"presupuestados y los egresos {ars_m(ej['real_egresos'])} contra {ars_m(ej['presupuesto_egresos'])}.")
    elif pp and ej:
        pres_t += (f"Aplicando el mismo criterio con los datos previos, el resultado real del período ({ars_m(ej['real_resultado'])}) "
                   f"{'superó' if ej['desvio_resultado'] >= 0 else 'quedó por debajo de'} lo presupuestado ({ars_m(ej['presupuesto_resultado'])}) "
                   f"en {pct(abs(ej['desvio_resultado_pct']))}.")
    elif pp:
        pres_t += "No hay historial previo suficiente para medir la ejecución de este presupuesto en el período."

    # ---- flujo de caja (solo si el archivo trae la hoja Saldos)
    cj = k.get("caja")
    caja_t = caja_proy_t = ""
    pj = cj.get("proyeccion") if cj else None
    if cj:
        rp = cj["periodo"]
        caja_t = (f"Medido por caja, el saldo del club {per['frase']} pasó de {ars_m(rp['saldo_inicial'])} a {ars_m(rp['saldo_final'])}, "
                  f"una variación neta de {ars_m(rp['neto'])}. ")
        comp = [f"la actividad operativa {'aportó' if rp['operativo'] >= 0 else 'consumió'} {ars_m(abs(rp['operativo']))}"]
        if abs(rp["inversion"]) >= 1:
            comp.append(f"las inversiones {'aportaron' if rp['inversion'] >= 0 else 'demandaron'} {ars_m(abs(rp['inversion']))}")
        if abs(rp["financiacion"]) >= 1:
            comp.append(f"la financiación {'aportó' if rp['financiacion'] >= 0 else 'demandó'} {ars_m(abs(rp['financiacion']))} netos de cuotas")
        caja_t += "En ese lapso " + _lista_y(comp) + ". "
        if abs(rp["inversion"]) >= 1 or abs(rp["financiacion"]) >= 1:
            caja_t += (f"Por eso el flujo de caja ({ars_m(rp['neto'])}) difiere del resultado del período ({ars_m(t['resultado'])}): "
                       "las inversiones y la financiación mueven la caja, pero no forman parte del resultado. ")
        if rp["meses"] > 1:
            caja_t += f"El saldo más bajo del período fue de {ars_m(rp['minimo']['saldo'])} al cierre de {rp['minimo']['mes']}. "
        if len(rp["cuentas_final"]) > 1:
            caja_t += "Al cierre, el dinero se reparte en " + _lista_y([f"{c_} ({ars_m(v_)})" for c_, v_ in rp["cuentas_final"].items()]) + ". "
        if rp.get("cobertura") is not None:
            caja_t += (f"Con ese saldo, la caja cubre {rp['cobertura']:.1f}".replace(".", ",") +
                       f" meses de egresos operativos y cuotas de deuda (promedio mensual de los últimos {fc.MESES_REFERENCIA} meses: {ars_m(rp['egreso_referencia'])}). ")
        if rp.get("anterior"):
            va = rp["anterior"]
            caja_t += (f"Respecto del período anterior, el saldo final {'subió' if rp['saldo_final'] >= va['saldo_final'] else 'bajó'} "
                       f"de {ars_m(va['saldo_final'])} a {ars_m(rp['saldo_final'])}.")
    if pj:
        hz_ = pj["horizontes"]
        partes = [f"{ars_m(h_['saldo'])} a {h_['meses']} meses (hasta {h_['hasta']})" for h_ in hz_[:3]]
        caja_proy_t = (f"Partiendo de un saldo de {ars_m(pj['saldo_inicial'])} al cierre de {pj['ultimo_mes_historial']} y sumando a la proyección "
                       f"ordinaria los compromisos conocidos, el escenario base estima un saldo de {_lista_y(partes)}. ")
        if pj["compromisos"]:
            grandes = sorted(pj["compromisos"], key=lambda r_: -r_["monto"])[:3]
            det = _lista_y([f"{r_['concepto'].lower()} ({r_['fecha']}, {ars_m(r_['monto'])})" for r_ in grandes])
            caja_proy_t += (f"Se consideraron {len(pj['compromisos'])} compromisos ({pj['n_confirmados']} confirmados y {pj['n_probables']} probables); "
                            f"los de mayor monto son {det}. ")
        else:
            caja_proy_t += "No hay compromisos cargados para este horizonte, por lo que la caja se proyecta solo con la operatoria ordinaria. "
        hx = next((h_ for h_ in hz_ if h_["meses"] == 6), hz_[-1])
        caja_proy_t += f"El saldo más bajo estimado dentro de {hx['meses']} meses es de {ars_m(hx['saldo_minimo'])} ({hx['mes_minimo']}). "
        if hx["saldo_minimo_pesimista"] >= 0:
            caja_proy_t += "En el escenario pesimista el saldo se mantiene positivo en todo ese lapso. "
        else:
            caja_proy_t += f"En el escenario pesimista el saldo quedaría en negativo desde {hx['primer_mes_negativo_pesimista']}. "
        if pr and pr.get("extra"):
            hxe = next((h_ for h_ in hz_ if h_["meses"] == pr["extra"]), None)
            if hxe:
                caja_proy_t += f"Para el horizonte elegido de {hxe['meses']} meses se estima un saldo de {ars_m(hxe['saldo'])}. "
        caja_proy_t += "Los vencimientos de deuda, inversiones y otros pagos se toman de la hoja Compromisos; los ingresos y egresos habituales se proyectan por tendencia."

    alertas, temas = [], []
    if cj and cj["periodo"].get("cobertura") is not None and cj["periodo"]["cobertura"] < 2:
        alertas.append(f"La caja cubre solo {cj['periodo']['cobertura']:.1f}".replace(".", ",") + " meses de egresos operativos y cuotas de deuda.")
        temas.append("el nivel de caja")
    if pj:
        hx_ = next((h_ for h_ in pj["horizontes"] if h_["meses"] == 6), pj["horizontes"][-1])
        if hx_["saldo_minimo_pesimista"] < 0:
            alertas.append(f"En el escenario pesimista la caja quedaría en negativo desde {hx_['primer_mes_negativo_pesimista']}.")
            temas.append("la proyección de caja")
    if pr and pr["horizontes"][0]["res_pesimista"] < 0:
        alertas.append(f"En el escenario pesimista, el resultado proyectado a 3 meses sería negativo ({ars_m(pr['horizontes'][0]['res_pesimista'])}).")
        temas.append("el escenario pesimista de la proyección")
    if x["cantidad"] and abs(x["peso_en_resultado"]) >= 0.10:
        alertas.append(f"El {pct(abs(x['peso_en_resultado']))} del resultado proviene de movimientos extraordinarios no recurrentes.")
        temas.append("la dependencia de ingresos extraordinarios")
    if d["socios_con_deuda"] and any(r["estado"] != "Moroso" for r in d["top"]):
        alertas.append("Hay socios entre los mayores deudores que figuran como activos: el estado del padrón no refleja la deuda real.")
    if ic[0][1] / t["ingresos"] >= 0.40:
        alertas.append(f"Alta concentración de ingresos: {ic[0][0]} explica el {pct(ic[0][1] / t['ingresos'])} del total y una baja "
                       f"o demora en ese rubro impactaría de forma directa en el resultado.")
        temas.append(f"la dependencia de {ic[0][0]}")
    if ec[0][1] / t["egresos"] >= 0.50:
        alertas.append(f"{ec[0][0]} concentra el {pct(ec[0][1] / t['egresos'])} del gasto, lo que vuelve poco flexible la estructura de costos.")
        temas.append(f"el peso de {ec[0][0]} en el gasto")
    if ms and ms["sobre_ingresos"] > SALARIO_VERDE:
        alertas.append(f"La masa salarial absorbe el {pct(ms['sobre_ingresos'], 0)} de los ingresos ordinarios (referencia: hasta {pct(SALARIO_VERDE, 0)}).")
        temas.append("el peso de la masa salarial")
    if c and c["var_egresos"] > c["var_ingresos"] and c["var_egresos"] > 0:
        alertas.append(f"Los egresos crecen más rápido que los ingresos respecto de {ref} ({signo_pct(c['var_egresos'])} vs {signo_pct(c['var_ingresos'])}).")
        temas.append("el ritmo de crecimiento de los egresos")
    if s["tasa_morosidad"] >= 0.10:
        alertas.append(f"La morosidad alcanza el {pct(s['tasa_morosidad'], 0)} del padrón ({s['morosos']} de {s['total']} socios).")
        temas.append("la morosidad de socios")
    if s["cuotas_sobre_ingresos"] < 0.05:
        alertas.append(f"La cuota social cubre solo el {pct(s['cuotas_sobre_egresos'])} de los egresos, por lo que el club depende de ingresos no societarios.")
        temas.append("el bajo aporte de la cuota social")
    neg = [x["nombre"] for x in m if x["resultado"] < 0]
    if neg:
        alertas.append("Meses con resultado negativo: " + ", ".join(neg) + ".")
        temas.append("los meses con resultado negativo")
    if not alertas:
        alertas.append("No se detectan alertas relevantes con los criterios definidos.")

    cierre = (f"En síntesis, el club cerró {per['frase']} con un resultado {'superavitario' if sup else 'deficitario'}"
              f"{', sostenido principalmente por ' + ic[0][0] if sup else ''}. ")
    cierre += (f"Los principales aspectos a monitorear son {_lista_y(temas[:4])}, de modo de sostener este desempeño en los próximos períodos."
               if temas else "No se identifican riesgos relevantes, por lo que se recomienda mantener el seguimiento periódico.")

    reco = []
    if ic[0][1] / t["ingresos"] >= 0.40:
        reco.append(f"Diversificar ingresos: ampliar entradas, merchandising y alquiler de espacios para reducir la dependencia de {ic[0][0]}.")
    if s["tasa_morosidad"] >= 0.10:
        reco.append("Implementar un plan de regularización para socios morosos (recordatorios, planes de pago y débito automático).")
    if s["cuotas_sobre_ingresos"] < 0.05:
        reco.append("Revisar la política de cuota social y de captación de socios para que aporte una proporción mayor al financiamiento del club.")
    if d["socios_con_deuda"]:
        reco.append(f"Gestionar de forma directa el cobro de los {len(d['top'])} mayores deudores (contacto personal y plan de pagos), "
                    f"que concentran el {pct(d['participacion_top'])} de la deuda estimada.")
    if x["cantidad"] and abs(x["peso_en_resultado"]) >= 0.10:
        reco.append("Presupuestar sin contar los ingresos extraordinarios y destinarlos a reservas o inversiones puntuales.")
    if cj and cj["periodo"].get("cobertura") is not None and cj["periodo"]["cobertura"] < fc.COBERTURA_VERDE:
        reco.append(f"Reconstituir un fondo de caja equivalente a, al menos, {fc.COBERTURA_VERDE:.0f} meses de egresos (hoy cubre {cj['periodo']['cobertura']:.1f}".replace(".", ",") + ").")
    if pj:
        comp_pagos = [r_ for r_ in pj["compromisos"] if r_["tipo"] == "Pago" and r_["actividad"] in (fc.ACT_INVERSION, fc.ACT_FINANCIACION)]
        if comp_pagos:
            reco.append(f"Programar los pagos comprometidos de inversión y deuda ({ars_m(sum(r_['monto'] for r_ in comp_pagos))} en los próximos "
                        f"{len(pj['serie'])} meses) contra el saldo de caja proyectado, para no tensionar la liquidez.")
    reco.append("Destinar parte del superávit a un fondo de reserva y a inversión en infraestructura e inferiores."
                if sup else "Definir un plan de reducción de gastos y refuerzo de ingresos para revertir el déficit.")
    if ej and ej.get("origen") == "aprobado" and abs(ej["desvio_resultado_pct"]) > 0.15:
        reco.append("Analizar con la Comisión Directiva los desvíos más grandes contra el presupuesto aprobado y, si corresponde, reformularlo.")
    if pp:
        reco.append(f"Tomar como referencia el presupuesto propuesto a {pp['meses']} meses y controlar mensualmente el desvío por rubro.")
    reco.append("Repetir este reporte periódicamente para monitorear la evolución de los indicadores.")
    return {"introduccion": intro, "ingresos": ing, "egresos": egr, "resultado": res, "socios": soc,
            "deudores": deu, "extraordinarios": extra, "semaforo": sem, "proyeccion": proy_t, "presupuesto": pres_t,
            "caja": caja_t, "caja_proyeccion": caja_proy_t,
            "alertas": alertas, "cierre": cierre, "recomendaciones": reco[:6]}


def generar_analisis(k):
    """Redacta el análisis del informe con reglas fijas sobre los indicadores calculados."""
    return analisis_local(k)


# ----------------------------------------------------------------------------
# 4. REPORTE PDF
# ----------------------------------------------------------------------------
def _estilo_ejes(ax):
    for lado in ("top", "right", "left"):
        ax.spines[lado].set_visible(False)
    ax.spines["bottom"].set_color("#BBBBBB")
    ax.tick_params(colors="#444444", labelsize=7.5, length=0)
    ax.yaxis.grid(True, color="#E5E5E5", linewidth=0.7)
    ax.set_axisbelow(True)


def grafico_evolucion(k, ruta):
    t = k["tendencia"]
    x = list(range(len(t)))
    fig, ax = plt.subplots(figsize=(7.1, 2.9), dpi=200)
    w = 0.38
    al = [1.0 if v["en_periodo"] else 0.35 for v in t]
    primero = next(i for i in x if t[i]["en_periodo"])
    for i in x:
        ax.bar(i - w / 2, t[i]["ingresos"] / 1e6, w, color=HEX_ING, alpha=al[i], label="Ingresos" if i == primero else None)
        ax.bar(i + w / 2, t[i]["egresos"] / 1e6, w, color=HEX_EGR, alpha=al[i], label="Egresos" if i == primero else None)
    ax.plot(x, [v["resultado"] / 1e6 for v in t], color=HEX_RES, marker="o", markersize=4.5, linewidth=2, label="Resultado")
    ax.set_xticks(x)
    ax.set_xticklabels([v["etiqueta"] for v in t], fontsize=7 if len(t) > 9 else 7.5)
    ax.set_ylabel("Millones de $", fontsize=8, color="#444444")
    _estilo_ejes(ax)
    ax.legend(loc="upper left", ncol=3, fontsize=7.5, frameon=False)
    fig.tight_layout()
    fig.savefig(ruta)
    plt.close(fig)
    return 7.1 / 2.9


def grafico_composicion(k, ruta):
    fig, axs = plt.subplots(1, 2, figsize=(7.1, 3.0), dpi=200)
    for ax, datos, titulo in ((axs[0], k["ingresos_cat"], "Ingresos por rubro"),
                              (axs[1], k["egresos_cat"], "Egresos por rubro")):
        total = sum(datos.values())
        nombres = list(datos.keys())
        orden = nombres[::-1]
        valores = [datos[n] / 1e6 for n in orden]
        cols = [color_cat(n, nombres.index(n)) for n in orden]
        ax.barh(orden, valores, color=cols, height=0.62)
        mx = max(valores)
        for i, n in enumerate(orden):
            ax.text(valores[i] + mx * 0.02, i, pct(datos[n] / total, 0), va="center", fontsize=7.5, color="#222222")
        ax.set_xlim(0, mx * 1.2)
        ax.set_title(titulo, fontsize=9, fontweight="bold", loc="left", color="#111111")
        for lado in ("top", "right", "bottom"):
            ax.spines[lado].set_visible(False)
        ax.spines["left"].set_color("#BBBBBB")
        ax.tick_params(labelsize=7.5, length=0, colors="#444444")
        ax.set_xticks([])
    fig.tight_layout()
    fig.savefig(ruta)
    plt.close(fig)
    return 7.1 / 3.0


def grafico_socios(k, ruta):
    t = k["tendencia"]
    x = list(range(len(t)))
    etiq = [v["etiqueta"].split("-")[0] for v in t]
    fig, axs = plt.subplots(1, 2, figsize=(7.1, 2.5), dpi=200)
    ax = axs[0]
    vals = [v["tasa_cobro"] * 100 for v in t]
    for i in x:
        ax.bar(i, vals[i], color=HEX_COBRO, width=0.6, alpha=1.0 if t[i]["en_periodo"] else 0.35)
        ax.text(i, vals[i] + 1, f"{vals[i]:.0f}%", ha="center", fontsize=7, color="#222222")
    ax.set_ylim(0, 105)
    ax.set_xticks(x)
    ax.set_xticklabels(etiq, fontsize=7)
    ax.set_title("Socios que pagaron la cuota (%)", fontsize=9, fontweight="bold", loc="left")
    _estilo_ejes(ax)
    ax = axs[1]
    ax.plot(x, [v["cuota_prom"] / 1000 for v in t], color=HEX_CUOTA, marker="o", markersize=4.5, linewidth=2)
    ax.set_xticks(x)
    ax.set_xticklabels(etiq, fontsize=7)
    ax.set_title("Cuota promedio abonada (miles de $)", fontsize=9, fontweight="bold", loc="left")
    _estilo_ejes(ax)
    fig.tight_layout()
    fig.savefig(ruta)
    plt.close(fig)
    return 7.1 / 2.5


def grafico_proyeccion(proy, ruta, max_meses=None):
    """Historial ordinario (línea llena) + proyección (línea punteada) con banda del 80 %."""
    hist = proy["historial"][-9:]
    serie = proy["serie"][:max_meses] if max_meses else proy["serie"]
    nh, n = len(hist), len(hist) + len(serie)
    xh, xp = list(range(nh)), list(range(nh, n))
    fig, ax = plt.subplots(figsize=(7.1, 3.0), dpi=200)
    for clave, color, nombre in (("ingresos", HEX_ING, "Ingresos"), ("egresos", HEX_EGR, "Egresos")):
        ax.plot(xh, [v[clave] / 1e6 for v in hist], color=color, linewidth=2, marker="o", markersize=3.5, label=nombre)
        puente = [hist[-1][clave] / 1e6] + [v[clave] / 1e6 for v in serie]
        ax.plot([nh - 1] + xp, puente, color=color, linewidth=1.8, linestyle="--")
        ax.fill_between(xp, [v[clave + "_piso"] / 1e6 for v in serie], [v[clave + "_techo"] / 1e6 for v in serie],
                        color=color, alpha=0.13, linewidth=0)
    ax.plot([nh - 1] + xp, [hist[-1]["resultado"] / 1e6] + [v["resultado"] / 1e6 for v in serie], color=HEX_RES,
            linewidth=1.8, linestyle="--", marker="o", markersize=2.5, label="Resultado proyectado")
    ax.axvline(nh - 0.5, color="#888888", linewidth=0.8, linestyle=":")
    ymax = ax.get_ylim()[1]
    ax.text(nh - 0.35, ymax * 0.97, "Proyección", fontsize=7.5, color="#555555", va="top")
    ax.text(nh - 0.65, ymax * 0.97, "Historial", fontsize=7.5, color="#555555", va="top", ha="right")
    paso = max(1, -(-n // 14))
    etiq = [v["etiqueta"] for v in hist] + [v["etiqueta"] for v in serie]
    ax.set_xticks(range(0, n, paso))
    ax.set_xticklabels([etiq[i] for i in range(0, n, paso)], fontsize=7, rotation=45 if n > 12 else 0, ha="right" if n > 12 else "center")
    ax.set_ylabel("Millones de $ por mes", fontsize=8, color="#444444")
    _estilo_ejes(ax)
    ax.legend(loc="upper left", bbox_to_anchor=(0, 0.9), fontsize=7.5, frameon=False)
    fig.tight_layout()
    fig.savefig(ruta)
    plt.close(fig)
    return 7.1 / 3.0


def grafico_caja(k, ruta):
    """Flujo de caja mensual por actividad (barras apiladas) y saldo de cierre (línea)."""
    f = k["caja"]["periodo"]["filas"]
    x = list(range(len(f)))
    fig, ax = plt.subplots(figsize=(7.1, 3.0), dpi=200)
    pos, neg = [0.0] * len(f), [0.0] * len(f)
    for clave, nombre, color in (("operativo", "Operativa", HEX_ING), ("inversion", "Inversión", HEX_EGR), ("financiacion", "Financiación", HEX_CUOTA)):
        for i in x:
            v = f[i][clave] / 1e6
            if abs(v) < 1e-9:
                continue
            base = pos[i] if v >= 0 else neg[i]
            ax.bar(i, v, 0.58, bottom=base, color=color, alpha=1.0 if f[i]["en_periodo"] else 0.35)
            if v >= 0:
                pos[i] += v
            else:
                neg[i] += v
    saldo = [v["saldo_final"] / 1e6 for v in f]
    ax.plot(x, saldo, color=HEX_NEGRO, marker="o", markersize=4.5, linewidth=2, label="Saldo de cierre", zorder=5)
    if len(f) <= 12:
        for i in x:
            ax.annotate(f"{saldo[i]:,.0f}".replace(",", "."), (i, saldo[i]), textcoords="offset points", xytext=(0, 6), ha="center", fontsize=6.5, color="#111111")
    ax.axhline(0, color="#888888", linewidth=0.8)
    ax.set_xticks(x)
    ax.set_xticklabels([v["etiqueta"] for v in f], fontsize=7 if len(f) > 9 else 7.5)
    ax.set_ylabel("Millones de $", fontsize=8, color="#444444")
    lo, hi = ax.get_ylim()
    ax.set_ylim(lo, hi + (hi - lo) * 0.16)
    _estilo_ejes(ax)
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch
    manijas = [Line2D([0], [0], color=HEX_NEGRO, marker="o", markersize=4.5, linewidth=2, label="Saldo de cierre"),
               Patch(color=HEX_ING, label="Operativa"), Patch(color=HEX_EGR, label="Inversión"), Patch(color=HEX_CUOTA, label="Financiación")]
    ax.legend(handles=manijas, loc="upper left", ncol=4, fontsize=7.5, frameon=False)
    fig.tight_layout()
    fig.savefig(ruta)
    plt.close(fig)
    return 7.1 / 3.0


def grafico_caja_proyectada(k, ruta):
    """Saldo de caja: historial (línea llena) + proyección base (punteada) con banda pesimista-optimista."""
    hist = k["caja"]["periodo"]["filas"][-6:]
    serie = k["caja"]["proyeccion"]["serie"]
    nh, n = len(hist), len(hist) + len(serie)
    xh, xp = list(range(nh)), list(range(nh, n))
    fig, ax = plt.subplots(figsize=(7.1, 3.0), dpi=200)
    ax.plot(xh, [v["saldo_final"] / 1e6 for v in hist], color=HEX_NEGRO, linewidth=2, marker="o", markersize=3.5, label="Saldo histórico")
    base = [v["saldo"] / 1e6 for v in serie]
    ax.plot([nh - 1] + xp, [hist[-1]["saldo_final"] / 1e6] + base, color=HEX_ING, linewidth=1.9, linestyle="--", marker="o", markersize=2.5,
            label="Saldo proyectado (base)")
    ax.fill_between(xp, [v["saldo_pesimista"] / 1e6 for v in serie], [v["saldo_optimista"] / 1e6 for v in serie],
                    color=HEX_ING, alpha=0.14, linewidth=0, label="Banda pesimista – optimista")
    ax.axhline(0, color="#888888", linewidth=0.8)
    ax.axvline(nh - 0.5, color="#888888", linewidth=0.8, linestyle=":")
    imin = min(range(len(serie)), key=lambda i: serie[i]["saldo"])
    ax.annotate(f"mín. {base[imin]:,.0f}".replace(",", "."), (xp[imin], base[imin]), textcoords="offset points", xytext=(0, -13), ha="center", fontsize=7, color="#444444")
    ax.annotate(f"{base[-1]:,.0f}".replace(",", "."), (xp[-1], base[-1]), textcoords="offset points", xytext=(0, 7), ha="center", fontsize=7, color=HEX_ING)
    lo, hi = ax.get_ylim()
    ax.set_ylim(min(lo, 0), hi + (hi - lo) * 0.15)
    ymax = ax.get_ylim()[1]
    ax.text(nh - 0.35, ymax * 0.99, "Proyección", fontsize=7.5, color="#555555", va="top")
    ax.text(nh - 0.65, ymax * 0.99, "Historial", fontsize=7.5, color="#555555", va="top", ha="right")
    paso = max(1, -(-n // 14))
    etiq = [v["etiqueta"] for v in hist] + [v["etiqueta"] for v in serie]
    ax.set_xticks(range(0, n, paso))
    ax.set_xticklabels([etiq[i] for i in range(0, n, paso)], fontsize=7, rotation=45 if n > 12 else 0, ha="right" if n > 12 else "center")
    ax.set_ylabel("Millones de $", fontsize=8, color="#444444")
    _estilo_ejes(ax)
    ax.legend(loc="upper left", bbox_to_anchor=(0, 0.93), fontsize=7.5, frameon=False)
    fig.tight_layout()
    fig.savefig(ruta)
    plt.close(fig)
    return 7.1 / 3.0


def generar_graficos(k, carpeta):
    """Genera los gráficos como PNG y devuelve {nombre: (ruta, relación_aspecto)}."""
    os.makedirs(carpeta, exist_ok=True)
    out = {}
    for nombre, fn in (("evolucion", grafico_evolucion), ("composicion", grafico_composicion), ("socios", grafico_socios)):
        ruta = os.path.join(carpeta, f"{nombre}.png")
        out[nombre] = (ruta, fn(k, ruta))
    if k.get("proyeccion"):
        ruta = os.path.join(carpeta, "proyeccion.png")
        out["proyeccion"] = (ruta, grafico_proyeccion(k["proyeccion"], ruta))
    if k.get("caja"):
        ruta = os.path.join(carpeta, "caja.png")
        out["caja"] = (ruta, grafico_caja(k, ruta))
        if k["caja"].get("proyeccion"):
            ruta = os.path.join(carpeta, "caja_proyeccion.png")
            out["caja_proyeccion"] = (ruta, grafico_caja_proyectada(k, ruta))
    return out


def buscar_logo(carpeta):
    """Devuelve la ruta del escudo (escudo.png/jpg/jpeg) dentro de `carpeta`, o None."""
    for nombre in ("escudo.png", "escudo.jpg", "escudo.jpeg", "Escudo.png", "Escudo.jpg"):
        ruta = os.path.join(carpeta, nombre)
        if os.path.exists(ruta):
            return ruta
    return None


def construir_pdf(k, a, salida, logo=None, graficos=None, detalle="completo"):
    resumido = detalle == "resumido"
    W, H = A4
    mg = 1.8 * cm
    ancho = W - 2 * mg
    graficos = graficos or generar_graficos(k, tempfile.mkdtemp())

    base = ParagraphStyle("base", fontName="Helvetica", fontSize=9.5, leading=13.5, textColor=NEGRO)
    cuerpo = ParagraphStyle("cuerpo", parent=base, alignment=TA_JUSTIFY, spaceAfter=2)
    h = ParagraphStyle("h", parent=base, fontName="Helvetica-Bold", fontSize=12.5, leading=15, spaceBefore=12, spaceAfter=3)
    sub_st = ParagraphStyle("sub", parent=base, fontName="Helvetica-Bold", fontSize=10.5, leading=13, spaceBefore=9, spaceAfter=2)
    peq = ParagraphStyle("peq", parent=base, fontSize=7.5, leading=10, textColor=GRIS)
    vineta = ParagraphStyle("vi", parent=base, leftIndent=12, bulletIndent=0, spaceAfter=3.5)
    nota = ParagraphStyle("nota", parent=base, fontSize=8, leading=11, textColor=GRIS, leftIndent=10, bulletIndent=0, spaceAfter=2)
    celda = ParagraphStyle("celda", parent=base, fontSize=8, leading=10)

    def seccion(titulo):
        return [Paragraph(titulo.upper(), h), HRFlowable(width="100%", thickness=1.4, color=NEGRO, spaceAfter=6)]

    def subtitulo(titulo):
        return Paragraph(titulo, sub_st)

    def bloque(titulo, texto):
        return KeepTogether([subtitulo(titulo), Paragraph(escape(texto), cuerpo)])

    def bullets(items, estilo=vineta):
        return [Paragraph(escape(x), estilo, bulletText="\u2022") for x in items]

    tipo_txt = {"mensual": "Informe mensual", "trimestral": "Informe trimestral", "anual": "Informe anual"}[k["periodo"]["tipo"]]
    if resumido:
        tipo_txt += " (resumen)"

    def pegar_escudo(c, x, y, lado):
        """Dibuja el escudo sobre un fondo blanco para que se vea bien en el encabezado negro."""
        try:
            pad = 0.12 * cm
            c.setFillColor(colors.white)
            c.roundRect(x - pad, y - pad, lado + 2 * pad, lado + 2 * pad, 0.25 * cm, stroke=0, fill=1)
            c.drawImage(logo, x, y, lado, lado, mask="auto", preserveAspectRatio=True, anchor="c")
            return True
        except Exception:
            return False  # un escudo inválido nunca debe romper el informe

    def dibujar(c, doc, grande):
        alto = 4.2 * cm if grande else 1.7 * cm
        c.saveState()
        c.setFillColor(NEGRO)
        c.rect(0, H - alto, W, alto, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFillAlpha(0.10)
        x0 = W - 6.5 * cm
        while x0 < W:
            c.rect(x0, H - alto, 0.45 * cm, alto, stroke=0, fill=1)
            x0 += 0.9 * cm
        c.setFillAlpha(1)
        x_txt = mg
        if logo and os.path.exists(logo):
            lado = 2.6 * cm if grande else 1.0 * cm
            if pegar_escudo(c, mg + 0.12 * cm, H - alto + (alto - lado) / 2, lado):
                x_txt = mg + lado + (0.7 if grande else 0.55) * cm
        if grande:
            c.setFont("Helvetica", 8.5)
            c.setFillColor(GRIS_MEDIO)
            c.drawString(x_txt, H - 1.15 * cm, CLUB.upper())
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 22)
            c.drawString(x_txt, H - 2.25 * cm, "Reporte Financiero Ejecutivo")
            c.setFont("Helvetica", 10.5)
            c.setFillColor(GRIS_MEDIO)
            c.drawString(x_txt, H - 3.1 * cm, f"{tipo_txt}  |  {k['periodo']['etiqueta']}")
        else:
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 10)
            c.drawString(x_txt, H - 1.05 * cm, "Reporte Financiero Ejecutivo")
            c.setFont("Helvetica", 8.5)
            c.setFillColor(GRIS_MEDIO)
            c.drawRightString(W - mg, H - 1.05 * cm, f"{CLUB}  |  {k['periodo']['etiqueta']}")
        c.setStrokeColor(GRIS_MEDIO)
        c.setLineWidth(0.5)
        c.line(mg, 1.35 * cm, W - mg, 1.35 * cm)
        c.setFont("Helvetica", 7.5)
        c.setFillColor(GRIS)
        c.drawString(mg, 0.9 * cm, f"{CLUB}  |  Documento generado automáticamente a partir de los registros del club")
        c.drawRightString(W - mg, 0.9 * cm, f"Página {doc.page}")
        c.restoreState()

    doc = BaseDocTemplate(salida, pagesize=A4, title=f"Reporte Financiero Ejecutivo - {CLUB} - {k['periodo']['etiqueta']}",
                          author=CLUB, leftMargin=mg, rightMargin=mg)
    kw = dict(leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([
        PageTemplate("primera", [Frame(mg, 1.7 * cm, ancho, H - 1.7 * cm - 4.7 * cm, id="f1", **kw)], onPage=lambda c, d: dibujar(c, d, True)),
        PageTemplate("resto", [Frame(mg, 1.7 * cm, ancho, H - 1.7 * cm - 2.2 * cm, id="f2", **kw)], onPage=lambda c, d: dibujar(c, d, False)),
    ])

    t, s, comp, m = k["totales"], k["socios"], k["comparacion"], k["mensual"]

    def tarjeta(titulo, valor, sub, color=NEGRO):
        return [Paragraph(titulo.upper(), ParagraphStyle("t", parent=peq, fontName="Helvetica-Bold")),
                Paragraph(valor, ParagraphStyle("v", parent=base, fontName="Helvetica-Bold", fontSize=14.5, leading=19, textColor=color)),
                Paragraph(sub, ParagraphStyle("s", parent=peq, fontSize=7))]

    sin_prev = "Sin período anterior<br/>con datos"
    col_res = VERDE if t["resultado"] >= 0 else ROJO
    kpis = Table([[
        tarjeta("Ingresos totales", ars_m(t["ingresos"]),
                f"{signo_pct(comp['var_ingresos'])} vs {escape(comp['etiqueta'])}<br/>Anterior: {ars_m(comp['ingresos'])}" if comp else sin_prev),
        tarjeta("Egresos totales", ars_m(t["egresos"]),
                f"{signo_pct(comp['var_egresos'])} vs {escape(comp['etiqueta'])}<br/>Anterior: {ars_m(comp['egresos'])}" if comp else sin_prev),
        tarjeta("Resultado", ars_m(t["resultado"]),
                f"{'Superávit' if t['resultado'] >= 0 else 'Déficit'} del período<br/>" +
                (f"Anterior: {ars_m(comp['resultado'])}" if comp else "&nbsp;"), col_res),
        tarjeta("Margen", pct(t["margen"]), "Resultado sobre ingresos<br/>" +
                (f"Anterior: {pct(comp['margen'])}" if comp else "&nbsp;"), col_res),
    ]], colWidths=[ancho / 4] * 4)
    kpis.setStyle(TableStyle([
        ("LINEABOVE", (0, 0), (-1, 0), 2.5, NEGRO), ("BACKGROUND", (0, 0), (-1, -1), GRIS_CLARO),
        ("LINEAFTER", (0, 0), (-2, -1), 3, colors.white), ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))

    def tabla_mensual():
        filas = [["Mes", "Ingresos", "Egresos", "Resultado", "Margen"]]
        for v in m:
            filas.append([_cap(v["nombre"]), ars_m(v["ingresos"]), ars_m(v["egresos"]), ars_m(v["resultado"]), pct(v["margen"])])
        filas.append(["Total", ars_m(t["ingresos"]), ars_m(t["egresos"]), ars_m(t["resultado"]), pct(t["margen"])])
        tb = Table(filas, colWidths=[ancho * 0.28, ancho * 0.21, ancho * 0.21, ancho * 0.18, ancho * 0.12], repeatRows=1)
        est = [("FONT", (0, 0), (-1, -1), "Helvetica", 8.5), ("BACKGROUND", (0, 0), (-1, 0), NEGRO),
               ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8.5),
               ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 8.5), ("LINEABOVE", (0, -1), (-1, -1), 1, NEGRO),
               ("ALIGN", (1, 0), (-1, -1), "RIGHT"), ("TOPPADDING", (0, 0), (-1, -1), 3.5),
               ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5), ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, GRIS_CLARO])]
        for i, v in enumerate(m + [{"resultado": t["resultado"]}], start=1):
            est.append(("TEXTCOLOR", (3, i), (3, i), VERDE if v["resultado"] >= 0 else ROJO))
        tb.setStyle(TableStyle(est))
        return KeepTogether([subtitulo("Detalle mensual"), tb])

    def tabla_movimientos(titulo, lista, total):
        filas = [["Fecha", "Rubro", "Descripción", "Monto"]]
        for r in lista:
            filas.append([r["fecha"], Paragraph(escape(r["rubro"]), celda), Paragraph(escape(r["descripcion"] + (" [extraordinario]" if r.get("extra") else "")), celda), ars(r["monto"])])
        filas.append(["", "", "Total", ars(total)])
        tb = Table(filas, colWidths=[ancho * 0.10, ancho * 0.24, ancho * 0.46, ancho * 0.20], repeatRows=1)
        tb.setStyle(TableStyle([
            ("FONT", (0, 0), (-1, -1), "Helvetica", 8), ("BACKGROUND", (0, 0), (-1, 0), NEGRO),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8),
            ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 8), ("LINEABOVE", (0, -1), (-1, -1), 1, NEGRO),
            ("ALIGN", (3, 0), (3, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, GRIS_CLARO])]))
        return KeepTogether([subtitulo(titulo), tb])

    def mini(titulo, valor, sub=""):
        return [Paragraph(titulo.upper(), ParagraphStyle("mt", parent=peq, fontName="Helvetica-Bold")),
                Paragraph(valor, ParagraphStyle("mv", parent=base, fontName="Helvetica-Bold", fontSize=13, leading=17)),
                Paragraph(sub or "&nbsp;", ParagraphStyle("ms", parent=peq, fontSize=7))]

    socios_t = Table([[
        mini("Padrón de socios", str(s["total"]), f"{s['activos']} activos"),
        mini("Morosos", str(s["morosos"]), f"{pct(s['tasa_morosidad'], 0)} del padrón"),
        mini("Recaudación cuotas", ars_m(s["recaudacion_cuotas"]), f"{pct(s['cuotas_sobre_ingresos'])} de los ingresos"),
        mini("Cuota promedio", ars(s["cuota_promedio"]), "por pago registrado"),
    ]], colWidths=[ancho / 4] * 4)
    socios_t.setStyle(TableStyle([
        ("LINEABOVE", (0, 0), (-1, 0), 2.5, NEGRO), ("BACKGROUND", (0, 0), (-1, -1), GRIS_CLARO),
        ("LINEAFTER", (0, 0), (-2, -1), 3, colors.white), ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7)]))

    def img(nombre):
        ruta, ratio = graficos[nombre]
        return Image(ruta, width=ancho, height=ancho / ratio)

    d, x = k["deudores"], k["extraordinarios"]

    def tabla_extraordinarios():
        filas = [["Fecha", "Tipo", "Rubro", "Descripción", "Monto"]]
        for r in x["ingresos"] + x["egresos"]:
            filas.append([r["fecha"], r["tipo"], Paragraph(escape(r["rubro"]), celda), Paragraph(escape(r["descripcion"]), celda), ars(r["monto"])])
        tb = Table(filas, colWidths=[ancho * 0.13, ancho * 0.10, ancho * 0.17, ancho * 0.40, ancho * 0.20], repeatRows=1)
        est = [("FONT", (0, 0), (-1, -1), "Helvetica", 8), ("BACKGROUND", (0, 0), (-1, 0), NEGRO),
               ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8),
               ("ALIGN", (4, 0), (4, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
               ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
               ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, GRIS_CLARO])]
        for i, r in enumerate(x["ingresos"] + x["egresos"], start=1):
            est.append(("TEXTCOLOR", (1, i), (1, i), VERDE if r["tipo"] == "Ingreso" else ROJO))
        tb.setStyle(TableStyle(est))
        tira = Table([[
            mini("Resultado del período", ars_m(t["resultado"])),
            mini("Efecto de extraordinarios", ars_m(x["neto"]), "ingresos menos egresos extraordinarios"),
            mini("Resultado sin extraordinarios", ars_m(x["resultado_ordinario"]), "operatoria habitual"),
        ]], colWidths=[ancho / 3] * 3)
        tira.setStyle(TableStyle([
            ("LINEABOVE", (0, 0), (-1, 0), 2.5, NEGRO), ("BACKGROUND", (0, 0), (-1, -1), GRIS_CLARO),
            ("LINEAFTER", (0, 0), (-2, -1), 3, colors.white), ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7)]))
        return [tb, Spacer(1, 6), tira]

    def tabla_deudores():
        filas = [["#", "Socio", "Estado", "Meses impagos", "Deuda estimada", "Último pago"]]
        for r in d["top"]:
            filas.append([str(r["puesto"]), Paragraph(escape(etq_socio(r, True)), celda), r["estado"],
                          str(r["meses_impagos"]), ars(r["deuda"]), r["ultimo_pago"]])
        tb = Table(filas, colWidths=[ancho * 0.06, ancho * 0.31, ancho * 0.13, ancho * 0.16, ancho * 0.18, ancho * 0.16], repeatRows=1)
        est = [("FONT", (0, 0), (-1, -1), "Helvetica", 8.5), ("BACKGROUND", (0, 0), (-1, 0), NEGRO),
               ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8.5),
               ("ALIGN", (3, 0), (4, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
               ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
               ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, GRIS_CLARO])]
        for i, r in enumerate(d["top"], start=1):
            est.append(("TEXTCOLOR", (2, i), (2, i), ROJO if r["estado"] == "Moroso" else GRIS))
        tb.setStyle(TableStyle(est))
        pie = Paragraph(f"Deuda estimada total del club: <b>{ars(d['deuda_total'])}</b> en {d['socios_con_deuda']} de {d['total_socios']} socios "
                        f"(morosos: {ars(d['deuda_morosos'])}; activos: {ars(d['deuda_activos'])}). Calculada desde {d['calculada_desde']}.", peq)
        return [tb, Spacer(1, 3), pie]

    # ---------- semáforo, proyección y presupuesto
    sm, pr, pp, ej = k["semaforo"], k.get("proyeccion"), k.get("presupuesto"), k.get("ejecucion")
    HEX_EST = {"verde": "#1B7F4B", "amarillo": "#F2B705", "rojo": "#B3261E"}

    def punto(estado, tam=13):
        return f'<font name="{FUENTE_SIMBOLOS}" color="{HEX_EST[estado]}" size="{tam}">\u25cf</font>'

    def tabla_semaforo():
        n_ = sm["conteo"]
        centro = ParagraphStyle("centro", parent=base, alignment=1, leading=18)
        etq = ParagraphStyle("etq", parent=peq, fontName="Helvetica-Bold", alignment=1)
        resumen = Table([[[Paragraph("SITUACIÓN GENERAL", etq), Paragraph(punto(sm["general"], 24), ParagraphStyle("c1", parent=centro, leading=28))],
                          [Paragraph("INDICADORES", etq),
                           Paragraph(f"{punto('verde', 15)} <b>{n_['verde']}</b> &nbsp;&nbsp;&nbsp; {punto('amarillo', 15)} <b>{n_['amarillo']}</b> "
                                     f"&nbsp;&nbsp;&nbsp; {punto('rojo', 15)} <b>{n_['rojo']}</b>", ParagraphStyle("c2", parent=centro, leading=28, fontSize=12))]]],
                         colWidths=[ancho * 0.4, ancho * 0.6])
        resumen.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), GRIS_CLARO), ("LINEABOVE", (0, 0), (-1, 0), 2.5, NEGRO),
                                     ("LINEAFTER", (0, 0), (0, 0), 3, colors.white), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                     ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
        filas = [["Estado", "Indicador", "Valor", "Criterio"]]
        for i in sm["items"]:
            filas.append([Paragraph(punto(i["estado"], 14), ParagraphStyle("pt", parent=centro, leading=14)),
                          Paragraph(escape(i["indicador"]), celda), Paragraph(escape(i["valor"]), celda),
                          Paragraph(escape(i["criterio"]), ParagraphStyle("cr", parent=celda, fontSize=7, textColor=GRIS))])
        tb = Table(filas, colWidths=[ancho * 0.09, ancho * 0.28, ancho * 0.26, ancho * 0.37], repeatRows=1)
        tb.setStyle(TableStyle([("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8), ("BACKGROUND", (0, 0), (-1, 0), NEGRO),
                                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("ALIGN", (0, 0), (0, 0), "CENTER"),
                                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 3),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("LINEBELOW", (0, 1), (-1, -1), 0.4, GRIS_CLARO)]))
        return [resumen, Spacer(1, 6), tb]

    def rango_m(lo, hi):
        return f"{ars_m(lo)} – {ars_m(hi)}"

    def tabla_proyeccion():
        filas = [["Horizonte", "Hasta", "Ingresos", "Egresos", "Resultado", "Pesimista – optimista", "Confiabilidad"]]
        for h_ in pr["horizontes"]:
            filas.append([f"{h_['meses']} meses", h_["hasta"].capitalize(), ars_m(h_["ingresos"]), ars_m(h_["egresos"]), ars_m(h_["resultado"]),
                          rango_m(h_["res_pesimista"], h_["res_optimista"]), h_["confiabilidad"]])
        tb = Table(filas, colWidths=[ancho * w for w in (0.10, 0.16, 0.13, 0.13, 0.13, 0.22, 0.13)], repeatRows=1)
        est = [("FONT", (0, 0), (-1, -1), "Helvetica", 8), ("BACKGROUND", (0, 0), (-1, 0), NEGRO), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
               ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8), ("ALIGN", (2, 0), (5, -1), "RIGHT"), ("TOPPADDING", (0, 0), (-1, -1), 3.5),
               ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, GRIS_CLARO])]
        for n_, h_ in enumerate(pr["horizontes"], start=1):
            est.append(("TEXTCOLOR", (4, n_), (4, n_), VERDE if h_["resultado"] >= 0 else ROJO))
            est.append(("TEXTCOLOR", (6, n_), (6, n_), {"Aceptable": VERDE, "Baja": colors.HexColor("#9A6B00")}.get(h_["confiabilidad"], ROJO)))
            if pr.get("extra") == h_["meses"]:
                est.append(("FONT", (0, n_), (-1, n_), "Helvetica-Bold", 8))
        tb.setStyle(TableStyle(est))
        return tb

    def tabla_presupuesto():
        h_ = pp["meses"]
        filas = [["Rubro", "Actual (prom. mensual)", "Presupuesto mensual", f"Total {h_} meses", "Var."]]
        est = [("FONT", (0, 0), (-1, -1), "Helvetica", 8), ("BACKGROUND", (0, 0), (-1, 0), NEGRO), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
               ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8), ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
               ("TOPPADDING", (0, 0), (-1, -1), 2.8), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.8)]

        def bloque_filas(titulo, lista, total):
            filas.append([titulo, "", "", "", ""])
            est.extend([("BACKGROUND", (0, len(filas) - 1), (-1, len(filas) - 1), GRIS_CLARO), ("FONT", (0, len(filas) - 1), (-1, len(filas) - 1), "Helvetica-Bold", 8)])
            for f in lista:
                filas.append(["   " + f["rubro"], ars_m(f["actual_mensual"]), ars_m(f["mensual"]), ars_m(f["total"]), signo_pct(f["variacion"], 0)])
            act, men = sum(f["actual_mensual"] for f in lista), sum(f["mensual"] for f in lista)
            filas.append([f"Total {titulo.lower()}", ars_m(act), ars_m(men), ars_m(total), signo_pct(men / act - 1, 0) if act else "-"])
            est.extend([("FONT", (0, len(filas) - 1), (-1, len(filas) - 1), "Helvetica-Bold", 8), ("LINEABOVE", (0, len(filas) - 1), (-1, len(filas) - 1), 0.6, NEGRO)])

        bloque_filas("Ingresos", pp["filas_ingresos"], pp["total_ingresos"])
        bloque_filas("Egresos", pp["filas_egresos"], pp["total_egresos"])
        filas.append([f"Reserva de contingencia ({pct(pp['criterio']['reserva'], 0)} de egresos)", "", "", ars_m(pp["reserva"]), ""])
        filas.append(["Resultado presupuestado", "", "", ars_m(pp["resultado_final"]), pct(pp["margen"], 0)])
        n_ = len(filas) - 1
        est.extend([("FONT", (0, n_), (-1, n_), "Helvetica-Bold", 8.5), ("LINEABOVE", (0, n_), (-1, n_), 1.2, NEGRO),
                    ("TEXTCOLOR", (0, n_), (-1, n_), VERDE if pp["resultado_final"] >= 0 else ROJO)])
        tb = Table(filas, colWidths=[ancho * w for w in (0.34, 0.20, 0.18, 0.18, 0.10)], repeatRows=1)
        tb.setStyle(TableStyle(est))
        return tb

    def tabla_trimestres():
        filas = [["Trimestre", "Ingresos", "Egresos", "Reserva", "Resultado"]]
        for t_ in pp["trimestres"]:
            filas.append([t_["nombre"], ars_m(t_["ingresos"]), ars_m(t_["egresos"]), ars_m(t_["reserva"]), ars_m(t_["resultado"])])
        tb = Table(filas, colWidths=[ancho * w for w in (0.28, 0.18, 0.18, 0.18, 0.18)], repeatRows=1)
        tb.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "Helvetica", 8), ("BACKGROUND", (0, 0), (-1, 0), NEGRO), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                                ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8), ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                                ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, GRIS_CLARO])]))
        return tb

    def tabla_ejecucion():
        filas = [["Rubro", "Presupuesto", "Real", "Desvío", "Desvío %"]]
        est = [("FONT", (0, 0), (-1, -1), "Helvetica", 7.5), ("BACKGROUND", (0, 0), (-1, 0), NEGRO), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
               ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 7.5), ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
               ("TOPPADDING", (0, 0), (-1, -1), 2.2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2)]

        def bloque_ej(titulo, lista, pres_t, real_t, favorable_si_mayor):
            filas.append([titulo, "", "", "", ""])
            est.extend([("BACKGROUND", (0, len(filas) - 1), (-1, len(filas) - 1), GRIS_CLARO), ("FONT", (0, len(filas) - 1), (-1, len(filas) - 1), "Helvetica-Bold", 7.5)])
            for f in lista:
                if max(f["presupuesto"], f["real"]) <= 0:
                    continue
                filas.append(["   " + f["rubro"], ars_m(f["presupuesto"]), ars_m(f["real"]), ars_m(f["desvio"]), signo_pct(f["desvio_pct"], 0) if f["presupuesto"] else "n/c"])
                bien = (f["desvio"] >= 0) == favorable_si_mayor
                est.append(("TEXTCOLOR", (3, len(filas) - 1), (4, len(filas) - 1), VERDE if bien else ROJO))
            filas.append([f"Total {titulo.lower()}", ars_m(pres_t), ars_m(real_t), ars_m(real_t - pres_t), signo_pct(real_t / pres_t - 1, 0) if pres_t else "n/c"])
            est.extend([("FONT", (0, len(filas) - 1), (-1, len(filas) - 1), "Helvetica-Bold", 7.5), ("LINEABOVE", (0, len(filas) - 1), (-1, len(filas) - 1), 0.6, NEGRO)])

        bloque_ej("Ingresos", ej["ingresos"], ej["presupuesto_ingresos"], ej["real_ingresos"], True)
        bloque_ej("Egresos", ej["egresos"], ej["presupuesto_egresos"], ej["real_egresos"], False)
        filas.append(["Resultado", ars_m(ej["presupuesto_resultado"]), ars_m(ej["real_resultado"]), ars_m(ej["desvio_resultado"]), signo_pct(ej["desvio_resultado_pct"], 0)])
        n_ = len(filas) - 1
        est.extend([("FONT", (0, n_), (-1, n_), "Helvetica-Bold", 8), ("LINEABOVE", (0, n_), (-1, n_), 1.2, NEGRO),
                    ("TEXTCOLOR", (3, n_), (4, n_), VERDE if ej["desvio_resultado"] >= 0 else ROJO)])
        tb = Table(filas, colWidths=[ancho * w for w in (0.34, 0.17, 0.17, 0.17, 0.15)], repeatRows=1)
        tb.setStyle(TableStyle(est))
        return tb

    # ---------- flujo de caja
    cj = k.get("caja")
    pj = cj.get("proyeccion") if cj else None

    def col_signo(v):
        return GRIS if abs(v) < 1 else (VERDE if v > 0 else ROJO)

    def estilo_base(tam=8, rel=True):
        est = [("FONT", (0, 0), (-1, -1), "Helvetica", tam), ("BACKGROUND", (0, 0), (-1, 0), NEGRO), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
               ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", tam), ("ALIGN", (1, 0), (-1, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
               ("TOPPADDING", (0, 0), (-1, -1), 3.2), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.2)]
        if rel:
            est.append(("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, GRIS_CLARO]))
        return est

    def tira_caja():
        rp = cj["periodo"]
        cob = rp.get("cobertura")
        t_ = Table([[mini("Saldo inicial", ars_m(rp["saldo_inicial"]), "al inicio del período"),
                     mini("Variación neta", ars_m(rp["neto"]), "operativo + inversión + financiación"),
                     mini("Saldo final", ars_m(rp["saldo_final"]), f"mínimo {ars_m(rp['minimo']['saldo'])}" if rp["meses"] > 1 else "al cierre del mes"),
                     mini("Cobertura de caja", (f"{cob:.1f}".replace(".", ",") + " meses") if cob is not None else "n/c",
                          "de egresos y cuotas de deuda")]], colWidths=[ancho / 4] * 4)
        t_.setStyle(TableStyle([("LINEABOVE", (0, 0), (-1, 0), 2.5, NEGRO), ("BACKGROUND", (0, 0), (-1, -1), GRIS_CLARO),
                                ("LINEAFTER", (0, 0), (-2, -1), 3, colors.white), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                ("LEFTPADDING", (0, 0), (-1, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7)]))
        return t_

    def tabla_caja():
        rp = cj["periodo"]
        filas = [["Mes", "Saldo inicial", "Operativo", "Inversión", "Financiación", "Variación", "Saldo final"]]
        for f in rp["filas"]:
            filas.append([_cap(f["nombre"]), ars_m(f["saldo_inicial"]), ars_m(f["operativo"]), ars_m(f["inversion"]),
                          ars_m(f["financiacion"]), ars_m(f["neto"]), ars_m(f["saldo_final"])])
        total = rp["meses"] > 1
        if total:
            filas.append(["Total del período", ars_m(rp["saldo_inicial"]), ars_m(rp["operativo"]), ars_m(rp["inversion"]),
                          ars_m(rp["financiacion"]), ars_m(rp["neto"]), ars_m(rp["saldo_final"])])
        tb = Table(filas, colWidths=[ancho * w for w in (0.19, 0.14, 0.13, 0.13, 0.14, 0.13, 0.14)], repeatRows=1)
        est = estilo_base(8)
        for i, f in enumerate(rp["filas"], start=1):
            if not f["en_periodo"]:
                est.append(("TEXTCOLOR", (0, i), (-1, i), GRIS_MEDIO))
                continue
            for col, clave in ((2, "operativo"), (3, "inversion"), (4, "financiacion"), (5, "neto")):
                est.append(("TEXTCOLOR", (col, i), (col, i), col_signo(f[clave])))
            est.append(("FONT", (6, i), (6, i), "Helvetica-Bold", 8))
        if total:
            n_ = len(filas) - 1
            est.extend([("FONT", (0, n_), (-1, n_), "Helvetica-Bold", 8), ("LINEABOVE", (0, n_), (-1, n_), 1, NEGRO)])
            for col, clave in ((2, "operativo"), (3, "inversion"), (4, "financiacion"), (5, "neto")):
                est.append(("TEXTCOLOR", (col, n_), (col, n_), col_signo(rp[clave])))
        tb.setStyle(TableStyle(est))
        return tb

    def tabla_caja_horizontes():
        filas = [["Horizonte", "Hasta", "Flujo neto", "Saldo (base)", "Pesimista – optimista", "Saldo mínimo"]]
        for h_ in pj["horizontes"]:
            filas.append([f"{h_['meses']} meses", h_["hasta"].capitalize(), ars_m(h_["flujo_neto"]), ars_m(h_["saldo"]),
                          rango_m(h_["saldo_pesimista"], h_["saldo_optimista"]), ars_m(h_["saldo_minimo"])])
        tb = Table(filas, colWidths=[ancho * w for w in (0.11, 0.17, 0.15, 0.16, 0.27, 0.14)], repeatRows=1)
        est = estilo_base(8)
        for i, h_ in enumerate(pj["horizontes"], start=1):
            est.append(("TEXTCOLOR", (2, i), (2, i), col_signo(h_["flujo_neto"])))
            est.append(("TEXTCOLOR", (5, i), (5, i), VERDE if h_["saldo_minimo"] >= 0 else ROJO))
            if k["proyeccion"] and k["proyeccion"].get("extra") == h_["meses"]:
                est.append(("FONT", (0, i), (-1, i), "Helvetica-Bold", 8))
        tb.setStyle(TableStyle(est))
        return tb

    def tabla_caja_meses():
        sj = pj["serie"][:6]
        filas = [["Mes", "Operativo", "Inversión", "Financiación", "Variación", "Saldo (base)"]]
        for f in sj:
            filas.append([_cap(f["nombre"]), ars_m(f["operativo"]), ars_m(f["inversion"]), ars_m(f["financiacion"]), ars_m(f["neto"]), ars_m(f["saldo"])])
        tb = Table(filas, colWidths=[ancho * w for w in (0.22, 0.15, 0.15, 0.16, 0.15, 0.17)], repeatRows=1)
        est = estilo_base(8)
        for i, f in enumerate(sj, start=1):
            for col, clave in ((1, "operativo"), (2, "inversion"), (3, "financiacion"), (4, "neto")):
                est.append(("TEXTCOLOR", (col, i), (col, i), col_signo(f[clave])))
            est.append(("FONT", (5, i), (5, i), "Helvetica-Bold", 8))
        tb.setStyle(TableStyle(est))
        return tb

    def tabla_compromisos():
        lista = pj["compromisos"][:14]
        filas = [["Fecha", "Tipo", "Concepto", "Actividad", "Estado", "Monto"]]
        for r in lista:
            filas.append([r["fecha"], r["tipo"], Paragraph(escape(r["concepto"]), celda), r["actividad"], r["estado"], ars_m(r["monto"])])
        tb = Table(filas, colWidths=[ancho * w for w in (0.12, 0.08, 0.37, 0.15, 0.12, 0.16)], repeatRows=1)
        est = estilo_base(8)
        est.append(("ALIGN", (1, 0), (4, -1), "LEFT"))
        for i, r in enumerate(lista, start=1):
            est.append(("TEXTCOLOR", (1, i), (1, i), VERDE if r["tipo"] == "Cobro" else ROJO))
        tb.setStyle(TableStyle(est))
        extra_n = len(pj["compromisos"]) - len(lista)
        return [tb] + ([Spacer(1, 2), Paragraph(f"Y {extra_n} compromisos más dentro del horizonte.", peq)] if extra_n > 0 else [])

    q = k["calidad_datos"]
    notas = []
    if k["periodo"]["parcial"]:
        notas.append("Período parcial: se informan solo los meses con datos disponibles.")
    if q["cuota_no_esta_en_ingresos"]:
        notas.append("Los pagos de cuota social se suman como un rubro de ingreso adicional, ya que no figuran en la hoja de Ingresos.")
    if q["morosos_con_pagos"]:
        notas.append(f"{q['morosos_con_pagos']} socios con estado Moroso registran pagos de cuota en el período; conviene actualizar el estado en el padrón.")
    if q["pagos_previos_al_alta"]:
        notas.append(f"{q['pagos_previos_al_alta']} pagos tienen fecha anterior al alta del socio; se excluyen del cálculo de tasa de cobro.")
    notas.append("Deuda estimada = meses sin ningún pago registrado (desde el alta del socio) por la cuota promedio del club en cada mes. "
                 "No es una deuda contable: no hay una cuota fija en los datos.")
    notas.append("Se consideran extraordinarios los movimientos marcados como tales en el archivo (columna Tipo), los que contienen "
                 "palabras clave (streamer, amistoso, premio, multa...) o los que no se repiten en la mayoría de los meses.")
    if pr:
        notas.append(f"Proyección: método de {pr['metodo_txt']} sobre {pr['meses_historial']} meses de historial ordinario (sin extraordinarios), "
                     "con intervalo del 80%. Son valores nominales (incluyen la inflación implícita en la tendencia), sin estacionalidad y con "
                     "variación limitada a +40%/-30% del promedio reciente. No es una garantía de resultados.")
    if pp:
        notas.append(f"Presupuesto: tendencia por rubro (tope +15%/-20% sobre el promedio reciente), con ingresos {pct(pp['criterio']['prudencia_ingresos'], 0)} "
                     f"menores, egresos {pct(pp['criterio']['prudencia_egresos'], 0)} mayores y reserva de {pct(pp['criterio']['reserva'], 0)} de los egresos. "
                     "Es una propuesta elaborada a partir de los datos, no un presupuesto aprobado.")
    if q.get("no_operativos"):
        det = "; ".join(f"{a_.lower()}: {v_['n']} movimientos" + (f" (ingresos {ars_m(v_['ingresos'])}, egresos {ars_m(v_['egresos'])})" if a_ != fc.ACT_TRANSFERENCIA else "")
                        for a_, v_ in q["no_operativos"].items())
        notas.append("El resultado del período incluye solo la actividad operativa. Quedaron fuera del resultado, y se informan solo en el flujo de caja: "
                     f"{det}.")
    if cj:
        notas.append("Flujo de caja (método directo): la fecha de cada movimiento se toma como la fecha efectiva en que se movió el dinero. El saldo parte de la "
                     "hoja Saldos y suma, mes a mes, la actividad operativa, las inversiones y la financiación; las transferencias entre cuentas del club no "
                     "son ingresos ni gastos. Las cuotas sociales se consideran cobradas en la cuenta Banco. La cobertura de caja es el saldo final sobre el "
                     f"egreso mensual promedio de los últimos {fc.MESES_REFERENCIA} meses (operativo más cuotas de deuda).")
    if pj:
        nota_c = ("Flujo de caja proyectado: parte del saldo al cierre del último mes histórico y suma a la proyección ordinaria de ingresos y egresos "
                  "(sin extraordinarios) los compromisos de la hoja Compromisos. Escenario base: todos los compromisos; pesimista: piso de ingresos, techo de "
                  "egresos y sin los cobros probables; optimista: lo contrario.")
        if pj["ignorados_vencidos"]:
            nota_c += f" {pj['ignorados_vencidos']} compromisos con fecha anterior al inicio de la proyección no se tuvieron en cuenta."
        if not pj["compromisos"]:
            nota_c += " No hay compromisos cargados dentro del horizonte."
        notas.append(nota_c)
    if k.get("nota_mes_incompleto"):
        notas.append(k["nota_mes_incompleto"])
    notas.append("Semáforo: los umbrales de cada indicador figuran en la tabla y son criterios de gestión orientativos, ajustables al criterio de la comisión directiva.")
    notas.append("La cantidad de socios, activos y morosos corresponde al padrón actual. La tasa de cobro mensual es la cantidad "
                 "de socios que pagaron sobre los socios dados de alta hasta ese mes.")
    if k["periodo"]["tipo"] == "mensual":
        notas.append("En los gráficos, los meses más claros son contexto previo al período del informe.")
    if k.get("masa_salarial"):
        notas.append("Masa salarial = rubros de egreso cuyo nombre incluye sueldo, salario, remuneración o cuerpo técnico, sobre los ingresos ordinarios "
                     f"(sin extraordinarios). Referencia orientativa: verde hasta {pct(SALARIO_VERDE, 0)} y rojo desde {pct(SALARIO_ROJO, 0)}.")
    if ej and ej.get("origen") == "aprobado":
        notas.append(f"Presupuesto aprobado: se tomó de la hoja Presupuesto_Aprobado ({ej['rango']}) y se comparó con lo real ordinario del período "
                     "(sin extraordinarios), solo en los meses que tienen presupuesto cargado.")
    if k.get("anonimo"):
        notas.append("Datos personales protegidos: el informe no incluye nombres de socios, solo su número (Ley 25.326).")
    tz = k["trazabilidad"]
    traz_notas = [f"Trazabilidad: huella de los datos {tz['huella']} (" + ", ".join(f"{n_}: {c_} filas" for n_, c_ in tz["filas"].items()) + "). "
                  "Dos informes con la misma huella se generaron con los mismos datos."]
    traz_notas.append(f"Informe revisado y validado por: {tz['validado_por']}." if tz.get("validado_por")
                      else "Sin validación profesional registrada: el responsable debe revisar este informe antes de distribuirlo.")
    notas += traz_notas
    notas.append("Análisis redactado con reglas automáticas a partir de los indicadores calculados.")

    # ---- armado: 1. Introducción | 2. Desarrollo | 3. Cierre (las secciones del desarrollo se numeran solas)
    contador = iter(range(1, 100))

    def num(titulo):
        return f"2.{next(contador)} {titulo}"

    tit_ej = ("Ejecución: presupuesto aprobado vs. real ordinario del período" if ej and ej.get("origen") == "aprobado"
              else "Ejecución: presupuesto armado con datos previos vs. real del período")
    story = [NextPageTemplate("resto"), Spacer(1, 0.2 * cm), kpis]
    story += seccion("1. Introducción") + [Paragraph(escape(a["introduccion"]), cuerpo)]

    if resumido:
        # Versión resumida: cifras clave, evolución, caja, semáforo, puntos de atención, proyección y cierre.
        story += seccion("2. Lo más importante")
        story += [KeepTogether([subtitulo("Resultado y evolución"), Paragraph(escape(a["resultado"]), cuerpo), Spacer(1, 4), img("evolucion")]), Spacer(1, 6)]
        if cj:
            story += [KeepTogether([subtitulo("Flujo de caja"), Paragraph(escape(a["caja"]), cuerpo), Spacer(1, 4), tira_caja()]), Spacer(1, 6)]
        story.append(KeepTogether([subtitulo("Puntos de atención")] + bullets(a["alertas"][:5])))
        story += [Spacer(1, 6), KeepTogether([subtitulo("Semáforo de indicadores"), Paragraph(escape(a["semaforo"]), cuerpo), Spacer(1, 5)] + tabla_semaforo()), Spacer(1, 6)]
        if pr:
            story += [KeepTogether([subtitulo("Proyección a 3, 6 y 12 meses"), tabla_proyeccion()]), Spacer(1, 6)]
        if pj:
            story += [KeepTogether([subtitulo("Flujo de caja proyectado"), Paragraph(escape(a["caja_proyeccion"]), cuerpo), Spacer(1, 4),
                                    Paragraph("<b>Saldo de caja proyectado por horizonte</b>", peq), Spacer(1, 2), tabla_caja_horizontes()]), Spacer(1, 6)]
        notas_res = [n_ for n_ in (["Período parcial: se informan solo los meses con datos disponibles."] if k["periodo"]["parcial"] else [])]
        if k.get("nota_mes_incompleto"):
            notas_res.append(k["nota_mes_incompleto"])
        notas_res += ["Versión resumida del informe: el detalle de ingresos y egresos por rubro, movimientos extraordinarios, deudores, presupuesto y "
                      "tablas mensuales está en el informe completo.",
                      "Semáforo: los umbrales de cada indicador figuran en la tabla y son criterios de gestión orientativos.",
                      "Proyecciones y flujo de caja proyectado son estimaciones a partir de los datos, no garantías de resultados."]
        notas_res += traz_notas + (["Datos personales protegidos: el informe no incluye nombres de socios, solo su número (Ley 25.326)."] if k.get("anonimo") else [])
        notas_res += ["Análisis redactado con reglas automáticas a partir de los indicadores calculados."]
        story.append(KeepTogether(seccion("3. Cierre") + [Paragraph(escape(a["cierre"]), cuerpo)]
                                  + [subtitulo("Recomendaciones")] + bullets(a["recomendaciones"])
                                  + [Spacer(1, 8), Paragraph("<b>Notas</b>", peq)] + bullets(notas_res, nota)))
        doc.build(story)
        return salida

    story += seccion("2. Desarrollo")
    story += [bloque(num("Ingresos"), a["ingresos"]), bloque(num("Egresos"), a["egresos"]), Spacer(1, 4), img("composicion")]
    story += [KeepTogether([subtitulo(num("Resultado y evolución")), Paragraph(escape(a["resultado"]), cuerpo), Spacer(1, 4), img("evolucion")]), Spacer(1, 6)]
    if "movimientos" in k:
        story += [tabla_movimientos("Ingresos del mes", k["movimientos"]["ingresos"], t["ingresos"]), Spacer(1, 6),
                  tabla_movimientos("Egresos del mes", k["movimientos"]["egresos"], t["egresos"])]
    else:
        story += [tabla_mensual()]
    if cj:
        story += [Spacer(1, 6), KeepTogether([subtitulo(num("Flujo de caja")), Paragraph(escape(a["caja"]), cuerpo), Spacer(1, 4), tira_caja()]),
                  Spacer(1, 6), img("caja"), Spacer(1, 6),
                  KeepTogether([Paragraph("<b>Flujo de caja por mes</b>", peq), Spacer(1, 2), tabla_caja()]), Spacer(1, 6)]
    story += [KeepTogether([subtitulo(num("Movimientos extraordinarios")), Paragraph(escape(a["extraordinarios"]), cuerpo)]
                           + ([Spacer(1, 4)] + tabla_extraordinarios() if x["cantidad"] else [])), Spacer(1, 4)]
    story += [bloque(num("Socios y cuota social"), a["socios"]), Spacer(1, 4), socios_t, Spacer(1, 6), img("socios")]
    story += [KeepTogether([subtitulo(num(f"Los {len(d['top'])} mayores deudores")), Paragraph(escape(a["deudores"]), cuerpo), Spacer(1, 4)] + tabla_deudores())]
    if pp:
        story += [KeepTogether([subtitulo(num(f"Presupuesto propuesto a {pp['meses']} meses")), Paragraph(escape(a["presupuesto"]), cuerpo), Spacer(1, 4)]), tabla_presupuesto()]
        if pp.get("trimestres"):
            story += [Spacer(1, 6), KeepTogether([Paragraph("<b>Presupuesto por trimestre</b>", peq), Spacer(1, 2), tabla_trimestres()])]
        if ej:
            story += [Spacer(1, 6), KeepTogether([Paragraph(f"<b>{tit_ej}</b>", peq), Spacer(1, 2), tabla_ejecucion()])]
        story.append(Spacer(1, 6))
    else:
        story += [bloque(num("Presupuesto"), a["presupuesto"])]
        if ej:
            story += [Spacer(1, 4), KeepTogether([Paragraph(f"<b>{tit_ej}</b>", peq), Spacer(1, 2), tabla_ejecucion()]), Spacer(1, 6)]
    story.append(KeepTogether([subtitulo(num("Puntos de atención"))] + bullets(a["alertas"])))
    story += [Spacer(1, 6), KeepTogether([subtitulo(num("Semáforo de indicadores")), Paragraph(escape(a["semaforo"]), cuerpo), Spacer(1, 5)] + tabla_semaforo()), Spacer(1, 6)]
    if pr:
        sub_p = num("Proyección a 3, 6 y 12 meses") + (f" (y a {pr['extra']} meses)" if pr.get("extra") else "")
        story += [KeepTogether([subtitulo(sub_p), Paragraph(escape(a["proyeccion"]), cuerpo), Spacer(1, 4), img("proyeccion")]),
                  Spacer(1, 4), KeepTogether([tabla_proyeccion()]), Spacer(1, 6)]
    else:
        story += [bloque(num("Proyección"), a["proyeccion"])]
    if pj:
        story += [Spacer(1, 4), KeepTogether([subtitulo("Flujo de caja proyectado"), Paragraph(escape(a["caja_proyeccion"]), cuerpo)]),
                  Spacer(1, 4), img("caja_proyeccion"), Spacer(1, 4), KeepTogether([Paragraph("<b>Saldo de caja proyectado por horizonte</b>", peq), Spacer(1, 2), tabla_caja_horizontes()]),
                  Spacer(1, 6), KeepTogether([Paragraph("<b>Detalle de los primeros meses</b>", peq), Spacer(1, 2), tabla_caja_meses()])]
        if pj["compromisos"]:
            story += [Spacer(1, 6), KeepTogether([Paragraph("<b>Compromisos considerados en la proyección</b>", peq), Spacer(1, 2)] + tabla_compromisos())]
        story.append(Spacer(1, 6))

    # El cierre completo (cierre + recomendaciones + notas) va en un solo bloque: si no entra en lo que queda de la página,
    # pasa entero a la siguiente, así nunca queda cortado.
    story.append(KeepTogether(seccion("3. Cierre") + [Paragraph(escape(a["cierre"]), cuerpo)]
                              + [subtitulo("Recomendaciones")] + bullets(a["recomendaciones"])
                              + [Spacer(1, 8), Paragraph("<b>Notas metodológicas</b>", peq)] + bullets(notas, nota)))
    doc.build(story)
    return salida


# ----------------------------------------------------------------------------
# 5. REGISTRO DE CORRIDAS (seguimiento)
# ----------------------------------------------------------------------------
CAMPOS_REGISTRO = ["fecha_hora", "fuente", "tipo", "periodo", "ingresos", "egresos", "resultado", "detalle", "nombres_ocultos",
                   "huella_datos", "validado_por", "resumen"]


def registrar_corrida(ruta, fuente, k, a, detalle="completo"):
    fila = {
        "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "fuente": fuente,
        "tipo": k["periodo"]["tipo"], "periodo": k["periodo"]["etiqueta"],
        "ingresos": k["totales"]["ingresos"], "egresos": k["totales"]["egresos"], "resultado": k["totales"]["resultado"],
        "detalle": detalle, "nombres_ocultos": "sí" if k.get("anonimo") else "no",
        "huella_datos": k["trazabilidad"]["huella"], "validado_por": k["trazabilidad"].get("validado_por", ""),
        "resumen": a["introduccion"],
    }
    try:
        if os.path.exists(ruta):  # si el archivo viene de una versión anterior con otras columnas, se conserva aparte
            with open(ruta, encoding="utf-8") as f:
                if f.readline().strip().split(",") != CAMPOS_REGISTRO:
                    os.replace(ruta, ruta + ".anterior")
        nuevo = not os.path.exists(ruta)
        with open(ruta, "a", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=CAMPOS_REGISTRO)
            if nuevo:
                w.writeheader()
            w.writerow(fila)
    except OSError:
        pass  # el registro nunca debe romper el informe
    return fila


def leer_registro(ruta):
    if not os.path.exists(ruta):
        return pd.DataFrame(columns=CAMPOS_REGISTRO)
    return pd.read_csv(ruta)


# ----------------------------------------------------------------------------
# Línea de comandos
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Agente de reporte financiero ejecutivo - Riestra")
    ap.add_argument("fuente", nargs="+", help="ruta al .xlsx, carpeta o archivos .csv, o URL (Google Sheets / OneDrive)")
    ap.add_argument("--tipo", choices=list(TIPOS), default="mensual")
    ap.add_argument("--periodo", help="mensual: 2026-10 | trimestral: 2026-T3 | anual: 2026 (por defecto: el último disponible)")
    ap.add_argument("--salida", default="Reporte_Ejecutivo_Riestra.pdf")
    ap.add_argument("--logo")
    ap.add_argument("--proyeccion", type=int, default=0, metavar="MESES", help="horizonte adicional de proyección, de 1 a 36 meses (además de 3/6/12)")
    ap.add_argument("--metodo", choices=list(METODOS), default="tendencia", help="método de proyección")
    ap.add_argument("--presupuesto-meses", type=int, choices=(3, 6, 12), default=12)
    ap.add_argument("--detalle", choices=("completo", "resumido"), default="completo", help="nivel de detalle del PDF")
    ap.add_argument("--sin-nombres", action="store_true", help="no mostrar nombres de socios (solo su número)")
    ap.add_argument("--validado-por", default="", help="nombre de quien revisó y validó el informe (queda en el PDF y en el registro)")
    ap.add_argument("--registro", default="registro_corridas.csv")
    args = ap.parse_args()

    datos = cargar_desde_url(args.fuente[0]) if args.fuente[0].startswith("http") else cargar_rutas(args.fuente)
    clave = args.periodo or periodos_disponibles(datos)[args.tipo][0][1]
    k = calcular_kpis(datos, args.tipo, clave, args.proyeccion or None, args.metodo, args.presupuesto_meses, anonimo=args.sin_nombres)
    k["trazabilidad"]["validado_por"] = args.validado_por.strip()
    a = generar_analisis(k)
    construir_pdf(k, a, args.salida, args.logo or buscar_logo(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")), detalle=args.detalle)
    registrar_corrida(args.registro, " + ".join(args.fuente), k, a, args.detalle)
    print(f"Reporte generado: {args.salida}  ({k['periodo']['etiqueta']})")


if __name__ == "__main__":
    main()
