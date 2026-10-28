# -*- coding: utf-8 -*-
"""
Aplicación web del Agente Financiero - Club Atlético Deportivo Riestra.

Ejecutar en local:   streamlit run app.py
Publicar:            ver README.md (Streamlit Community Cloud)
"""
import io
import os
import tempfile

import pandas as pd
import streamlit as st

import agente_riestra as ag

AQUI = os.path.dirname(os.path.abspath(__file__))
RUTA_EJEMPLO = os.path.join(AQUI, "datos_ejemplo", "Riestra_Tesis_Finanzas_V2.xlsx")
RUTA_REGISTRO = os.path.join(AQUI, "registro_corridas.csv")
CARPETA_ASSETS = os.path.join(AQUI, "assets")

st.set_page_config(page_title="Agente Financiero | Riestra", page_icon="📊", layout="wide")


# ----------------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------------
def secreto(nombre, defecto=""):
    """Lee un secreto de Streamlit o una variable de entorno."""
    try:
        if nombre in st.secrets:
            return st.secrets[nombre]
    except Exception:
        pass
    return os.environ.get(nombre, defecto)


def mostrar_imagen(ruta):
    try:
        st.image(ruta, width="stretch")
    except Exception:  # versiones anteriores de Streamlit
        st.image(ruta, use_container_width=True)


@st.cache_data(ttl=300, show_spinner="Leyendo datos desde el enlace...")
def leer_url(url):
    return ag.cargar_desde_url(url)


@st.cache_data(show_spinner=False)
def leer_archivo(contenido):
    return ag.cargar_excel(io.BytesIO(contenido))


@st.cache_data(show_spinner=False)
def leer_ejemplo():
    return ag.cargar_excel(RUTA_EJEMPLO)


# ----------------------------------------------------------------------------
# Estilo y acceso
# ----------------------------------------------------------------------------
st.markdown(
    """
    <style>
      .encabezado {background:#111; color:#fff; padding:1.1rem 1.4rem; border-radius:6px; margin-bottom:1rem;
                   background-image:repeating-linear-gradient(90deg, rgba(255,255,255,0.07) 0 14px, transparent 14px 28px);}
      .encabezado small {color:#b5b5b5; letter-spacing:.08em;}
      .encabezado h2 {margin:.15rem 0 0 0; color:#fff;}
    </style>
    <div class="encabezado"><small>CLUB ATLÉTICO DEPORTIVO RIESTRA</small>
    <h2>Agente de Reporte Financiero</h2></div>
    """,
    unsafe_allow_html=True,
)

clave_app = secreto("APP_PASSWORD")
if clave_app and not st.session_state.get("autenticado"):
    ingreso = st.text_input("Contraseña de acceso", type="password")
    if ingreso and ingreso == clave_app:
        st.session_state["autenticado"] = True
        st.rerun()
    elif ingreso:
        st.error("Contraseña incorrecta.")
    st.stop()

api_key = secreto("ANTHROPIC_API_KEY")
if api_key:
    os.environ["ANTHROPIC_API_KEY"] = str(api_key)
hay_ia = bool(os.environ.get("ANTHROPIC_API_KEY"))
max_ia = int(secreto("MAX_IA_POR_SESION", 10))
st.session_state.setdefault("ia_usos", 0)

# ----------------------------------------------------------------------------
# Barra lateral: fuente de datos y período
# ----------------------------------------------------------------------------
with st.sidebar:
    st.subheader("1. Datos")
    fuente = st.radio("Fuente", ["Datos de ejemplo", "Enlace (Google Sheets / OneDrive)", "Subir archivo Excel"])
    url = archivo = None
    if fuente.startswith("Enlace"):
        url = st.text_input("Enlace del archivo", value=str(secreto("URL_DATOS_DEFAULT")),
                            help="Debe estar compartido como 'Cualquier persona con el enlace puede ver'.")
    elif fuente.startswith("Subir"):
        archivo = st.file_uploader("Archivo .xlsx", type=["xlsx"])

    if st.button("Cargar datos"):
        try:
            if fuente == "Datos de ejemplo":
                datos, etiqueta = leer_ejemplo(), "datos de ejemplo"
            elif url is not None:
                if not url.strip():
                    raise ValueError("Pegá el enlace del archivo.")
                if st.session_state.get("ultima_url") == url:
                    leer_url.clear()
                datos, etiqueta = leer_url(url.strip()), url.strip()
                st.session_state["ultima_url"] = url
            else:
                if archivo is None:
                    raise ValueError("Subí un archivo .xlsx.")
                datos, etiqueta = leer_archivo(archivo.getvalue()), f"archivo: {archivo.name}"
            st.session_state["datos"], st.session_state["fuente_txt"] = datos, etiqueta
            st.session_state.pop("informe", None)
            st.success("Datos cargados.")
        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Error inesperado al leer los datos: {e}")

    if "datos" not in st.session_state:  # la app abre lista para usar
        try:
            st.session_state["datos"], st.session_state["fuente_txt"] = leer_ejemplo(), "datos de ejemplo"
        except Exception as e:
            st.error(f"No se pudieron cargar los datos de ejemplo: {e}")

    datos = st.session_state.get("datos")
    if datos is not None:
        st.caption(f"Fuente activa: {st.session_state['fuente_txt']}")
        fmin, fmax = ag.rango_datos(datos)
        st.caption(f"Datos disponibles: {fmin:%d/%m/%Y} a {fmax:%d/%m/%Y}")

        st.subheader("2. Informe")
        tipo = st.radio("Tipo", list(ag.TIPOS), horizontal=True, format_func=str.capitalize)
        opciones = ag.periodos_disponibles(datos)[tipo]
        elegido = st.selectbox("Período", opciones, format_func=lambda o: o[0])
        usar_ia = st.checkbox("Redactar análisis con IA (Claude)", value=hay_ia, disabled=not hay_ia)
        if not hay_ia:
            st.caption("Sin API key configurada: se usa el redactor automático.")
        st.subheader("3. Escudo (opcional)")
        escudo = st.file_uploader("Escudo del club (PNG o JPG)", type=["png", "jpg", "jpeg"],
                                  help="Si no lo subís, se usa assets/escudo.png cuando existe.")
        generar = st.button("Generar informe", type="primary")
    else:
        escudo = None
        generar = False

# ----------------------------------------------------------------------------
# Generación
# ----------------------------------------------------------------------------
if generar and datos is not None:
    try:
        with st.spinner("Calculando indicadores y redactando el informe..."):
            clave = elegido[1]
            k = ag.calcular_kpis(datos, tipo, clave)
            quiere_ia = usar_ia and hay_ia
            if quiere_ia and st.session_state["ia_usos"] >= max_ia:
                st.warning(f"Se alcanzó el máximo de {max_ia} informes con IA por sesión. Se usa el redactor automático.")
                quiere_ia = False
            a, origen, info = ag.generar_analisis(k, usar_ia=quiere_ia)
            if origen == "ia":
                st.session_state["ia_usos"] += 1
            if info.get("aviso"):
                st.warning(info["aviso"])
            carpeta = tempfile.mkdtemp()
            graficos = ag.generar_graficos(k, carpeta)
            ruta_pdf = os.path.join(carpeta, "informe.pdf")
            logo = ag.buscar_logo(CARPETA_ASSETS)
            if escudo is not None:  # el escudo subido en la sesión tiene prioridad
                logo = os.path.join(carpeta, "escudo_subido" + os.path.splitext(escudo.name)[1].lower())
                with open(logo, "wb") as f:
                    f.write(escudo.getvalue())
            ag.construir_pdf(k, a, ruta_pdf, logo, origen, graficos)
            with open(ruta_pdf, "rb") as f:
                contenido_pdf = f.read()
            ag.registrar_corrida(RUTA_REGISTRO, st.session_state["fuente_txt"], k, a, origen, info)
        st.session_state["informe"] = {"k": k, "a": a, "origen": origen, "info": info, "graficos": graficos,
                                       "pdf": contenido_pdf, "nombre": f"Reporte_{tipo}_{clave}.pdf"}
    except Exception as e:
        st.error(f"No se pudo generar el informe: {e}")

# ----------------------------------------------------------------------------
# Resultado
# ----------------------------------------------------------------------------
inf = st.session_state.get("informe")
tab_inf, tab_datos, tab_reg = st.tabs(["Informe", "Datos y calidad", "Registro de corridas"])

with tab_inf:
    if not inf:
        st.info("Elegí el tipo de informe y el período en la barra lateral y presioná **Generar informe**.")
    else:
        k, a = inf["k"], inf["a"]
        t, comp = k["totales"], k["comparacion"]
        c1, c2 = st.columns([3, 1])
        c1.subheader(f"{k['periodo']['tipo'].capitalize()}: {k['periodo']['etiqueta']}")
        c1.caption("Análisis redactado por IA (Claude)." if inf["origen"] == "ia" else "Análisis redactado con reglas automáticas.")
        c2.download_button("Descargar PDF", data=inf["pdf"], file_name=inf["nombre"], mime="application/pdf")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Ingresos", ag.ars_m(t["ingresos"]),
                  f"{ag.signo_pct(comp['var_ingresos'])} vs {comp['etiqueta']}" if comp else None)
        m2.metric("Egresos", ag.ars_m(t["egresos"]),
                  f"{ag.signo_pct(comp['var_egresos'])} vs {comp['etiqueta']}" if comp else None, delta_color="inverse")
        m3.metric("Resultado", ag.ars_m(t["resultado"]))
        m4.metric("Margen", ag.pct(t["margen"]))

        st.markdown("### 1. Introducción")
        st.write(a["introduccion"])
        st.markdown("### 2. Desarrollo")
        st.markdown("**2.1 Ingresos**")
        st.write(a["ingresos"])
        st.markdown("**2.2 Egresos**")
        st.write(a["egresos"])
        mostrar_imagen(inf["graficos"]["composicion"][0])
        st.markdown("**2.3 Resultado y evolución**")
        st.write(a["resultado"])
        mostrar_imagen(inf["graficos"]["evolucion"][0])
        st.markdown("**2.4 Socios y cuota social**")
        st.write(a["socios"])
        mostrar_imagen(inf["graficos"]["socios"][0])
        st.markdown("**2.5 Puntos de atención**")
        for x in a["alertas"]:
            st.markdown(f"- {x}")
        st.markdown("### 3. Cierre")
        st.write(a["cierre"])
        st.markdown("**Recomendaciones**")
        for x in a["recomendaciones"]:
            st.markdown(f"- {x}")

with tab_datos:
    if not inf:
        st.info("Generá un informe para ver el detalle.")
    else:
        k = inf["k"]
        st.markdown("#### Detalle mensual")
        df_m = pd.DataFrame([{"Mes": x["nombre"].capitalize(), "Ingresos": x["ingresos"], "Egresos": x["egresos"],
                              "Resultado": x["resultado"], "Margen": round(x["margen"] * 100, 1),
                              "Socios que pagaron": x["pagantes"], "Tasa de cobro %": round(x["tasa_cobro"] * 100, 1)}
                             for x in k["mensual"]])
        st.dataframe(df_m, hide_index=True)
        cc1, cc2 = st.columns(2)
        cc1.markdown("#### Ingresos por rubro")
        cc1.dataframe(pd.DataFrame({"Rubro": list(k["ingresos_cat"]), "Monto": list(k["ingresos_cat"].values())}), hide_index=True)
        cc2.markdown("#### Egresos por rubro")
        cc2.dataframe(pd.DataFrame({"Rubro": list(k["egresos_cat"]), "Monto": list(k["egresos_cat"].values())}), hide_index=True)
        st.markdown("#### Calidad de los datos")
        q = k["calidad_datos"]
        if q["cuota_no_esta_en_ingresos"]:
            st.write("- La cuota social no figura en la hoja Ingresos: se suma como rubro adicional.")
        if q["morosos_con_pagos"]:
            st.write(f"- {q['morosos_con_pagos']} socios con estado Moroso registran pagos en el período.")
        if q["pagos_previos_al_alta"]:
            st.write(f"- {q['pagos_previos_al_alta']} pagos tienen fecha anterior al alta del socio.")
        if k["periodo"]["parcial"]:
            st.write("- El período es parcial: solo hay datos de una parte de los meses.")
        with st.expander("Indicadores calculados (JSON que recibe la IA)"):
            st.json({c: v for c, v in k.items() if c != "movimientos"})

with tab_reg:
    st.caption("Cada informe generado queda registrado. En Streamlit Community Cloud el archivo se reinicia al reiniciar "
               "la app, por eso conviene descargarlo.")
    reg = ag.leer_registro(RUTA_REGISTRO)
    if reg.empty:
        st.info("Todavía no se generó ningún informe.")
    else:
        st.dataframe(reg.iloc[::-1], hide_index=True)
        st.download_button("Descargar registro (CSV)", data=reg.to_csv(index=False).encode("utf-8"),
                           file_name="registro_corridas.csv", mime="text/csv")
