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
  3. ANÁLISIS  : Claude redacta el análisis usando SOLO los indicadores calculados.
                 Si no hay API key o falla la llamada, usa un redactor local.
  4. REPORTE   : arma el PDF ejecutivo con el template del club.
  5. REGISTRO  : guarda cada corrida (período, cifras, origen del texto, tokens).

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
from datetime import datetime
from xml.sax.saxutils import escape

import pandas as pd
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

# ----------------------------------------------------------------------------
# Constantes
# ----------------------------------------------------------------------------
CLUB = "Club Atlético Deportivo Riestra"
MODELO_POR_DEFECTO = "claude-sonnet-5-5"
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
    return "$\u00a0" + _miles(f"{n:,.0f}")


def ars_m(n):
    return "$\u00a0" + _miles(f"{n / 1e6:,.1f}") + "\u00a0M"


def pct(x, d=1):
    return f"{x * 100:.{d}f}".replace(".", ",") + "%"


def signo_pct(x, d=1):
    return ("+" if x >= 0 else "") + pct(x, d)


def _cap(txt):
    return txt[:1].upper() + txt[1:]


# ----------------------------------------------------------------------------
# 1. FUENTE DE DATOS
# ----------------------------------------------------------------------------
def preparar_datos(hojas):
    """Valida hojas/columnas y convierte fechas. `hojas` = dict nombre -> DataFrame."""
    hojas = {str(k).strip(): v for k, v in hojas.items()}
    faltan = [h for h in COLUMNAS if h not in hojas]
    if faltan:
        raise ValueError(f"Faltan estas pestañas en el archivo: {', '.join(faltan)}. "
                         f"Se esperan: {', '.join(COLUMNAS)}.")
    for hoja, cols in COLUMNAS.items():
        df = hojas[hoja]
        df.columns = [str(c).strip() for c in df.columns]
        falt = [c for c in cols if c not in df.columns]
        if falt:
            raise ValueError(f"En la pestaña '{hoja}' faltan las columnas: {', '.join(falt)}.")
    socios = hojas["Socios"].copy()
    pagos = hojas["Pago_Cuota_Social"].copy()
    ing = hojas["Ingresos"].copy()
    egr = hojas["Egresos"].copy()
    socios["Fecha_Alta"] = pd.to_datetime(socios["Fecha_Alta"], errors="coerce")
    for d in (pagos, ing, egr):
        d["Fecha"] = pd.to_datetime(d["Fecha"], errors="coerce")
        d["Monto"] = pd.to_numeric(d["Monto"], errors="coerce")
        d.dropna(subset=["Fecha", "Monto"], inplace=True)
    if min(len(pagos), len(ing), len(egr)) == 0:
        raise ValueError("Alguna de las pestañas de movimientos no tiene datos válidos.")
    return {"socios": socios, "pagos": pagos, "ing": ing, "egr": egr}


def cargar_excel(fuente):
    """`fuente`: ruta, archivo subido o BytesIO."""
    return preparar_datos(pd.read_excel(fuente, sheet_name=None))


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
    desde, hasta = per.start_time, per.end_time
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


def calcular_kpis(datos, tipo, clave):
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
        "calidad_datos": {
            "pagos_previos_al_alta": int((chk["Fecha"] < chk["Fecha_Alta"]).sum()),
            "morosos_con_pagos": int(pagos[pagos["ID_Socio"].isin(ids_morosos)]["ID_Socio"].nunique()),
            "cuota_no_esta_en_ingresos": "Cuota social" not in set(ing["Categoria"].astype(str)),
        },
    }
    if len(meses) == 1:  # informe mensual: detalle de movimientos
        def movs(df):
            cols = [c for c in ("Categoria", "Descripcion") if c in df.columns]
            return [{"fecha": r["Fecha"].strftime("%d/%m"), "rubro": str(r["Categoria"]),
                     "descripcion": str(r["Descripcion"]) if "Descripcion" in cols else "",
                     "monto": int(r["Monto"])} for _, r in df.sort_values("Fecha").iterrows()]
        mi = movs(ing)
        mi.append({"fecha": "mes", "rubro": "Cuota social", "descripcion": f"{len(pagos)} pagos de socios", "monto": cuotas_total})
        k["movimientos"] = {"ingresos": mi, "egresos": movs(egr)}
    return k


# ----------------------------------------------------------------------------
# 3. ANÁLISIS (IA con respaldo local)
# ----------------------------------------------------------------------------
PROMPT_SISTEMA = (
    "Sos analista financiero del Club Atlético Deportivo Riestra. Redactás para la "
    "comisión directiva (lectores no técnicos). Usá EXCLUSIVAMENTE las cifras del JSON "
    "recibido; no inventes datos ni hagas cálculos que no estén en el JSON. El campo "
    "periodo.tipo indica si el informe es mensual, trimestral o anual; si hay "
    "'comparacion', contrastá con ese período anterior. Montos en pesos argentinos. "
    "Tono ejecutivo, claro y sobrio, español neutro. Respondé SOLO con un JSON válido, "
    "sin texto extra ni bloques de código, con esta forma: "
    '{"resumen": "3 a 4 oraciones", "puntos_clave": ["4 ítems"], '
    '"alertas": ["2 a 4 ítems"], "recomendaciones": ["3 a 4 ítems accionables"]}'
)


def analisis_con_claude(k):
    import anthropic  # pip install anthropic
    modelo = os.environ.get("CLAUDE_MODEL", MODELO_POR_DEFECTO)
    carga = {c: v for c, v in k.items() if c != "movimientos"}
    r = anthropic.Anthropic().messages.create(
        model=modelo, max_tokens=1500, system=PROMPT_SISTEMA,
        messages=[{"role": "user", "content": json.dumps(carga, ensure_ascii=False)}],
    )
    txt = "".join(b.text for b in r.content if b.type == "text")
    txt = re.sub(r"^```(?:json)?|```$", "", txt.strip(), flags=re.M).strip()
    a = json.loads(txt)
    for campo in ("resumen", "puntos_clave", "alertas", "recomendaciones"):
        if campo not in a:
            raise ValueError(f"Falta el campo {campo}")
    return a, {"modelo": modelo, "tokens_in": r.usage.input_tokens, "tokens_out": r.usage.output_tokens}


def analisis_local(k):
    """Redactor de respaldo (sin IA): reglas fijas sobre los mismos indicadores."""
    t, s, c = k["totales"], k["socios"], k["comparacion"]
    m, tend = k["mensual"], k["tendencia"]
    ic, ec = list(k["ingresos_cat"].items()), list(k["egresos_cat"].items())
    sup = t["resultado"] >= 0

    res = (f"{_cap(k['periodo']['frase'])}, el club registró ingresos por {ars_m(t['ingresos'])} y egresos por "
           f"{ars_m(t['egresos'])}, con un resultado {'superavitario' if sup else 'deficitario'} de "
           f"{ars_m(abs(t['resultado']))} (margen de {pct(t['margen'])} sobre los ingresos). ")
    if c:
        ref = c["etiqueta"].lower() if k["periodo"]["tipo"] == "mensual" else c["etiqueta"]
        res += (f"Respecto de {ref}, los ingresos {'crecieron' if c['var_ingresos'] >= 0 else 'cayeron'} "
                f"{pct(abs(c['var_ingresos']))} y los egresos {'crecieron' if c['var_egresos'] >= 0 else 'cayeron'} "
                f"{pct(abs(c['var_egresos']))}; el resultado pasó de {ars_m(c['resultado'])} a {ars_m(t['resultado'])}. ")
    res += (f"La cuota social aportó {ars_m(s['recaudacion_cuotas'])} ({pct(s['cuotas_sobre_ingresos'])} de los "
            f"ingresos) y hay {s['morosos']} socios morosos sobre un padrón de {s['total']}.")

    puntos = [
        f"{ic[0][0]} es la principal fuente de ingresos: {ars_m(ic[0][1])} ({pct(ic[0][1] / t['ingresos'])} del total), "
        f"seguida por {ic[1][0]} con {pct(ic[1][1] / t['ingresos'])}.",
        f"El mayor egreso es {ec[0][0]}: {ars_m(ec[0][1])} ({pct(ec[0][1] / t['egresos'])} del gasto total).",
    ]
    if len(m) > 1:
        mejor, peor = max(m, key=lambda x: x["resultado"]), min(m, key=lambda x: x["resultado"])
        puntos.append(f"El mejor resultado mensual fue en {mejor['nombre']} ({ars_m(mejor['resultado'])}) y el más "
                      f"bajo en {peor['nombre']} ({ars_m(peor['resultado'])}).")
    else:
        cob = t["ingresos"] / t["egresos"] if t["egresos"] else 0
        puntos.append(f"Los ingresos del mes cubrieron {str(round(cob, 1)).replace('.', ',')} veces los egresos.")
    ini, fin = tend[0], tend[-1]
    if len(tend) > 1 and ini["cuota_prom"]:
        v = fin["cuota_prom"] / ini["cuota_prom"] - 1
        puntos.append(f"La cuota promedio pasó de {ars(ini['cuota_prom'])} en {ini['nombre']} a {ars(fin['cuota_prom'])} "
                      f"en {fin['nombre']} ({signo_pct(v)}); en el período abonó en promedio el "
                      f"{pct(s['tasa_cobro_promedio'], 0)} de los socios por mes.")
    else:
        puntos.append(f"La cuota promedio fue de {ars(s['cuota_promedio'])}; abonó el "
                      f"{pct(s['tasa_cobro_promedio'], 0)} de los socios.")

    alertas = []
    if ic[0][1] / t["ingresos"] >= 0.40:
        alertas.append(f"Alta concentración de ingresos: {ic[0][0]} explica el {pct(ic[0][1] / t['ingresos'])} del total. "
                       f"Una baja o demora en ese rubro impactaría de forma directa en el resultado.")
    if ec[0][1] / t["egresos"] >= 0.50:
        alertas.append(f"{ec[0][0]} concentra el {pct(ec[0][1] / t['egresos'])} del gasto: la estructura de costos es poco flexible.")
    if c and c["var_egresos"] > c["var_ingresos"] and c["var_egresos"] > 0:
        alertas.append(f"Los egresos crecen más rápido que los ingresos respecto de {c['etiqueta']} "
                       f"({signo_pct(c['var_egresos'])} vs {signo_pct(c['var_ingresos'])}).")
    if s["tasa_morosidad"] >= 0.10:
        alertas.append(f"Morosidad de {pct(s['tasa_morosidad'], 0)} ({s['morosos']} de {s['total']} socios).")
    if s["cuotas_sobre_ingresos"] < 0.05:
        alertas.append(f"La cuota social cubre solo el {pct(s['cuotas_sobre_egresos'])} de los egresos: el club depende "
                       f"de ingresos no societarios.")
    neg = [x["nombre"] for x in m if x["resultado"] < 0]
    if neg:
        alertas.append("Meses con resultado negativo: " + ", ".join(neg) + ".")
    if not alertas:
        alertas.append("No se detectan alertas relevantes con los criterios definidos.")

    reco = []
    if ic[0][1] / t["ingresos"] >= 0.40:
        reco.append(f"Diversificar ingresos: ampliar entradas, merchandising y alquiler de espacios para reducir la "
                    f"dependencia de {ic[0][0]}.")
    if s["tasa_morosidad"] >= 0.10:
        reco.append("Implementar un plan de regularización para socios morosos (recordatorios, planes de pago y débito automático).")
    if s["cuotas_sobre_ingresos"] < 0.05:
        reco.append("Revisar la política de cuota social y de captación de socios para que aporte una proporción mayor "
                    "al financiamiento del club.")
    reco.append("Destinar parte del superávit a un fondo de reserva y a inversión en infraestructura e inferiores."
                if sup else "Definir un plan de reducción de gastos y refuerzo de ingresos para revertir el déficit.")
    reco.append("Repetir este reporte periódicamente para monitorear la evolución de los indicadores.")
    return {"resumen": res, "puntos_clave": puntos, "alertas": alertas, "recomendaciones": reco}


def generar_analisis(k, usar_ia=True):
    """Devuelve (analisis, origen, info) donde origen es 'ia' o 'local'."""
    if usar_ia and os.environ.get("ANTHROPIC_API_KEY"):
        try:
            a, info = analisis_con_claude(k)
            return a, "ia", info
        except Exception as e:
            return analisis_local(k), "local", {"aviso": f"Falló la IA ({type(e).__name__}: {e}). Se usó el redactor local."}
    return analisis_local(k), "local", {}


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
        ax.bar(i - w / 2, t[i]["ingresos"] / 1e6, w, color=HEX_NEGRO, alpha=al[i], label="Ingresos" if i == primero else None)
        ax.bar(i + w / 2, t[i]["egresos"] / 1e6, w, color=HEX_GRIS, alpha=al[i], label="Egresos" if i == primero else None)
    ax.plot(x, [v["resultado"] / 1e6 for v in t], color=HEX_VERDE, marker="o", markersize=4, linewidth=1.8, label="Resultado")
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
    for ax, datos, titulo, color in ((axs[0], k["ingresos_cat"], "Ingresos por rubro", HEX_NEGRO),
                                      (axs[1], k["egresos_cat"], "Egresos por rubro", HEX_GRIS)):
        total = sum(datos.values())
        nombres = list(datos.keys())[::-1]
        valores = [datos[n] / 1e6 for n in nombres]
        ax.barh(nombres, valores, color=color, height=0.62)
        mx = max(valores)
        for i, n in enumerate(nombres):
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
        ax.bar(i, vals[i], color=HEX_NEGRO, width=0.6, alpha=1.0 if t[i]["en_periodo"] else 0.35)
        ax.text(i, vals[i] + 1, f"{vals[i]:.0f}%", ha="center", fontsize=7, color="#222222")
    ax.set_ylim(0, 105)
    ax.set_xticks(x)
    ax.set_xticklabels(etiq, fontsize=7)
    ax.set_title("Socios que pagaron la cuota (%)", fontsize=9, fontweight="bold", loc="left")
    _estilo_ejes(ax)
    ax = axs[1]
    ax.plot(x, [v["cuota_prom"] / 1000 for v in t], color=HEX_NEGRO, marker="o", markersize=4, linewidth=1.8)
    ax.set_xticks(x)
    ax.set_xticklabels(etiq, fontsize=7)
    ax.set_title("Cuota promedio abonada (miles de $)", fontsize=9, fontweight="bold", loc="left")
    _estilo_ejes(ax)
    fig.tight_layout()
    fig.savefig(ruta)
    plt.close(fig)
    return 7.1 / 2.5


def generar_graficos(k, carpeta):
    """Genera los 3 gráficos como PNG y devuelve {nombre: (ruta, relación_aspecto)}."""
    os.makedirs(carpeta, exist_ok=True)
    out = {}
    for nombre, fn in (("evolucion", grafico_evolucion), ("composicion", grafico_composicion), ("socios", grafico_socios)):
        ruta = os.path.join(carpeta, f"{nombre}.png")
        out[nombre] = (ruta, fn(k, ruta))
    return out


def construir_pdf(k, a, salida, logo=None, origen_analisis="local", graficos=None):
    W, H = A4
    mg = 1.8 * cm
    ancho = W - 2 * mg
    graficos = graficos or generar_graficos(k, tempfile.mkdtemp())

    base = ParagraphStyle("base", fontName="Helvetica", fontSize=9.5, leading=13.5, textColor=NEGRO)
    cuerpo = ParagraphStyle("cuerpo", parent=base, alignment=TA_JUSTIFY)
    h = ParagraphStyle("h", parent=base, fontName="Helvetica-Bold", fontSize=12.5, leading=15, spaceBefore=12, spaceAfter=3)
    peq = ParagraphStyle("peq", parent=base, fontSize=7.5, leading=10, textColor=GRIS)
    vineta = ParagraphStyle("vi", parent=base, leftIndent=12, bulletIndent=0, spaceAfter=3.5)
    nota = ParagraphStyle("nota", parent=base, fontSize=8, leading=11, textColor=GRIS, leftIndent=10, bulletIndent=0, spaceAfter=2)
    celda = ParagraphStyle("celda", parent=base, fontSize=8, leading=10)

    def seccion(titulo):
        return [Paragraph(titulo.upper(), h), HRFlowable(width="100%", thickness=1.4, color=NEGRO, spaceAfter=6)]

    def bullets(items, estilo=vineta):
        return [Paragraph(escape(x), estilo, bulletText="\u2022") for x in items]

    tipo_txt = {"mensual": "Informe mensual", "trimestral": "Informe trimestral", "anual": "Informe anual"}[k["periodo"]["tipo"]]

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
            lado = alto - 1.0 * cm
            c.drawImage(logo, mg, H - alto + 0.5 * cm, lado, lado, mask="auto", preserveAspectRatio=True)
            x_txt = mg + lado + 0.5 * cm
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

    # detalle: tabla mensual (trimestral/anual) o movimientos (mensual)
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
        return seccion("Detalle mensual") + [tb]

    def tabla_movimientos(titulo, lista, total):
        filas = [["Fecha", "Rubro", "Descripción", "Monto"]]
        for r in lista:
            filas.append([r["fecha"], Paragraph(escape(r["rubro"]), celda), Paragraph(escape(r["descripcion"]), celda), ars(r["monto"])])
        filas.append(["", "", "Total", ars(total)])
        tb = Table(filas, colWidths=[ancho * 0.10, ancho * 0.24, ancho * 0.46, ancho * 0.20], repeatRows=1)
        tb.setStyle(TableStyle([
            ("FONT", (0, 0), (-1, -1), "Helvetica", 8), ("BACKGROUND", (0, 0), (-1, 0), NEGRO),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8),
            ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 8), ("LINEABOVE", (0, -1), (-1, -1), 1, NEGRO),
            ("ALIGN", (3, 0), (3, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, GRIS_CLARO])]))
        return seccion(titulo) + [tb]

    # socios
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
    notas.append("La cantidad de socios, activos y morosos corresponde al padrón actual. La tasa de cobro mensual es la cantidad "
                 "de socios que pagaron sobre los socios dados de alta hasta ese mes.")
    if k["periodo"]["tipo"] == "mensual":
        notas.append("En los gráficos, los meses más claros son contexto previo al período del informe.")
    notas.append("Análisis redactado por IA (Claude) a partir de los indicadores calculados."
                 if origen_analisis == "ia" else "Análisis redactado con reglas automáticas a partir de los indicadores calculados.")

    story = [NextPageTemplate("resto"), Spacer(1, 0.2 * cm), kpis]
    story += seccion("Resumen ejecutivo") + [Paragraph(escape(a["resumen"]), cuerpo)]
    story += seccion("Puntos clave") + bullets(a["puntos_clave"])
    story.append(KeepTogether(seccion("Evolución mensual") + [img("evolucion")]))
    story.append(PageBreak())
    story.append(KeepTogether(seccion("Composición de ingresos y egresos") + [img("composicion")]))
    if "movimientos" in k:
        story.append(KeepTogether(tabla_movimientos("Detalle de ingresos del mes", k["movimientos"]["ingresos"], t["ingresos"])))
        story.append(Spacer(1, 6))
        story.append(KeepTogether(tabla_movimientos("Detalle de egresos del mes", k["movimientos"]["egresos"], t["egresos"])))
    else:
        story.append(KeepTogether(tabla_mensual()))
    story.append(PageBreak())
    story += seccion("Socios y cuota social") + [Spacer(1, 2), socios_t, Spacer(1, 6), img("socios")]
    story.append(KeepTogether(seccion("Alertas") + bullets(a["alertas"])))
    story.append(KeepTogether(seccion("Recomendaciones") + bullets(a["recomendaciones"])))
    story.append(Spacer(1, 8))
    story += [Paragraph("<b>Notas metodológicas</b>", peq)] + bullets(notas, nota)
    doc.build(story)
    return salida


# ----------------------------------------------------------------------------
# 5. REGISTRO DE CORRIDAS (seguimiento)
# ----------------------------------------------------------------------------
CAMPOS_REGISTRO = ["fecha_hora", "fuente", "tipo", "periodo", "ingresos", "egresos", "resultado",
                   "origen_analisis", "modelo", "tokens_in", "tokens_out", "resumen"]


def registrar_corrida(ruta, fuente, k, a, origen, info):
    fila = {
        "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "fuente": fuente,
        "tipo": k["periodo"]["tipo"], "periodo": k["periodo"]["etiqueta"],
        "ingresos": k["totales"]["ingresos"], "egresos": k["totales"]["egresos"], "resultado": k["totales"]["resultado"],
        "origen_analisis": origen, "modelo": (info or {}).get("modelo", ""),
        "tokens_in": (info or {}).get("tokens_in", ""), "tokens_out": (info or {}).get("tokens_out", ""),
        "resumen": a["resumen"],
    }
    nuevo = not os.path.exists(ruta)
    try:
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
    ap.add_argument("fuente", help="ruta al .xlsx o URL (Google Sheets / OneDrive)")
    ap.add_argument("--tipo", choices=list(TIPOS), default="mensual")
    ap.add_argument("--periodo", help="mensual: 2026-10 | trimestral: 2026-T3 | anual: 2026 (por defecto: el último disponible)")
    ap.add_argument("--salida", default="Reporte_Ejecutivo_Riestra.pdf")
    ap.add_argument("--logo")
    ap.add_argument("--sin-ia", action="store_true")
    ap.add_argument("--registro", default="registro_corridas.csv")
    args = ap.parse_args()

    datos = cargar_desde_url(args.fuente) if args.fuente.startswith("http") else cargar_excel(args.fuente)
    clave = args.periodo or periodos_disponibles(datos)[args.tipo][0][1]
    k = calcular_kpis(datos, args.tipo, clave)
    a, origen, info = generar_analisis(k, usar_ia=not args.sin_ia)
    if info.get("aviso"):
        print("[aviso]", info["aviso"])
    construir_pdf(k, a, args.salida, args.logo, origen)
    registrar_corrida(args.registro, args.fuente, k, a, origen, info)
    print(f"Reporte generado: {args.salida}  ({k['periodo']['etiqueta']}, análisis: {origen})")


if __name__ == "__main__":
    main()
