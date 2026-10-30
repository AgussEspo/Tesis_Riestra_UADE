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
import consultas as co

AQUI = os.path.dirname(os.path.abspath(__file__))
RUTA_EJEMPLO = os.path.join(AQUI, "datos_ejemplo", "Riestra_Tesis_Finanzas_V2.xlsx")
RUTA_REGISTRO = os.path.join(AQUI, "registro_corridas.csv")
RUTA_CONSULTAS = os.path.join(AQUI, "registro_consultas.csv")
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
def leer_archivos(items):
    return ag.cargar_archivos(list(items))


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
max_consultas = int(secreto("MAX_CONSULTAS_POR_SESION", 15))
st.session_state.setdefault("ia_usos", 0)

# ----------------------------------------------------------------------------
# Barra lateral: fuente de datos y período
# ----------------------------------------------------------------------------
with st.sidebar:
    st.subheader("1. Datos")
    fuente = st.radio("Fuente", ["Datos de ejemplo", "Enlace (Google Sheets / OneDrive)", "Subir archivo (Excel o CSV)"])
    url = archivo = None
    if fuente.startswith("Enlace"):
        url = st.text_input("Enlace del archivo", value=str(secreto("URL_DATOS_DEFAULT")),
                            help="Debe estar compartido como 'Cualquier persona con el enlace puede ver'.")
    elif fuente.startswith("Subir"):
        archivo = st.file_uploader("Archivos", type=["xlsx", "csv"], accept_multiple_files=True,
                                   help="Un Excel con las 4 pestañas, o 4 archivos CSV: socios, pagos, ingresos y egresos "
                                        "(el nombre de cada archivo debe incluir esa palabra).")

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
                if not archivo:
                    raise ValueError("Subí un Excel (.xlsx) o los 4 archivos CSV.")
                datos = leer_archivos(tuple((a.name, a.getvalue()) for a in archivo))
                etiqueta = "archivo: " + ", ".join(a.name for a in archivo)
            st.session_state["datos"], st.session_state["fuente_txt"] = datos, etiqueta
            st.session_state.pop("informe", None)
            st.session_state.pop("chat", None)
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
tab_inf, tab_socios, tab_cons, tab_datos, tab_reg = st.tabs(["Informe", "Socios", "Consultas", "Datos y calidad", "Registro"])

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
        st.markdown("**2.4 Movimientos extraordinarios**")
        st.write(a["extraordinarios"])
        xs = k["extraordinarios"]
        if xs["cantidad"]:
            st.dataframe(pd.DataFrame(xs["ingresos"] + xs["egresos"])[["fecha", "tipo", "rubro", "descripcion", "monto", "motivo"]]
                         .rename(columns=str.capitalize), hide_index=True)
            e1, e2, e3 = st.columns(3)
            e1.metric("Resultado del período", ag.ars_m(t["resultado"]))
            e2.metric("Efecto de extraordinarios", ag.ars_m(xs["neto"]))
            e3.metric("Resultado sin extraordinarios", ag.ars_m(xs["resultado_ordinario"]))
        st.markdown("**2.5 Socios y cuota social**")
        st.write(a["socios"])
        mostrar_imagen(inf["graficos"]["socios"][0])
        dd = k["deudores"]
        st.markdown(f"**2.6 Los {len(dd['top'])} mayores deudores**")
        st.write(a["deudores"])
        if dd["top"]:
            st.dataframe(pd.DataFrame([{"#": r["puesto"], "Socio": f"{r['socio']} (N° {r['id']})", "Estado": r["estado"],
                                        "Meses impagos": r["meses_impagos"], "Deuda estimada": r["deuda"],
                                        "Último pago": r["ultimo_pago"]} for r in dd["top"]]), hide_index=True)
        st.markdown("**2.7 Puntos de atención**")
        for x in a["alertas"]:
            st.markdown(f"- {x}")
        st.markdown("### 3. Cierre")
        st.write(a["cierre"])
        st.markdown("**Recomendaciones**")
        for x in a["recomendaciones"]:
            st.markdown(f"- {x}")

with tab_socios:
    if datos is None:
        st.info("Cargá los datos para ver el estado de cuenta de los socios.")
    else:
        hasta_s = pd.Timestamp(inf["k"]["periodo"]["hasta"]) if inf else ag.rango_datos(datos)[1]
        cuenta = ag.estado_cuenta_socios(datos, hasta_s)
        con_deuda = cuenta[cuenta["Deuda_estimada"] > 0]
        st.caption(f"Estado de cuenta de **todos** los socios al {hasta_s:%d/%m/%Y}. Deuda estimada = meses sin pago por la "
                   "cuota promedio del club ese mes (no es una deuda contable).")
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Socios", len(cuenta))
        s2.metric("Con cuotas impagas", len(con_deuda))
        s3.metric("Deuda estimada total", ag.ars(int(cuenta["Deuda_estimada"].sum())))
        tot_s = cuenta["Deuda_estimada"].sum()
        s4.metric("Peso de los 5 mayores", ag.pct(con_deuda.head(5)["Deuda_estimada"].sum() / tot_s) if tot_s else "-")

        def vista(df):
            v = df.copy()
            v["Fecha_Alta"] = v["Fecha_Alta"].dt.strftime("%d/%m/%Y")
            v["Ultimo_pago"] = v["Ultimo_pago"].dt.strftime("%d/%m/%Y").fillna("sin pagos")
            return v.rename(columns={"Fecha_Alta": "Alta", "Meses_esperados": "Meses esperados", "Meses_pagados": "Meses pagados",
                                     "Meses_impagos": "Meses impagos", "Deuda_estimada": "Deuda estimada", "Detalle_impagos": "Detalle impagos",
                                     "Ultimo_pago": "Último pago", "Total_pagado": "Total pagado", "ID_Socio": "N°"})
        st.markdown("#### Los 5 mayores deudores")
        st.dataframe(vista(con_deuda.head(5)), hide_index=True)
        st.markdown("#### Todos los socios")
        filtro = st.selectbox("Mostrar", ["Todos", "Solo con deuda", "Solo morosos", "Solo activos"])
        base = {"Todos": cuenta, "Solo con deuda": con_deuda, "Solo morosos": cuenta[cuenta["Estado"] == "Moroso"],
                "Solo activos": cuenta[cuenta["Estado"] == "Activo"]}[filtro]
        st.dataframe(vista(base), hide_index=True)
        st.download_button("Descargar estado de cuenta (CSV)", data=vista(cuenta).to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"estado_cuenta_socios_{hasta_s:%Y-%m-%d}.csv", mime="text/csv")

with tab_cons:
    st.caption("Preguntale a los datos del club. Cada cifra la calcula el código sobre los datos cargados; "
               "la IA solo elige qué consultar y redacta la respuesta.")
    if datos is None:
        st.info("Cargá los datos para poder hacer consultas.")
    elif not hay_ia:
        st.info("Las consultas necesitan una API key de Anthropic (`ANTHROPIC_API_KEY` en los secretos de la app).")
    else:
        st.session_state.setdefault("chat", [])
        st.session_state.setdefault("consultas_usadas", 0)
        with st.expander("Ejemplos de preguntas"):
            st.markdown("- ¿Cuánto gastamos en seguridad en el tercer trimestre?\n"
                        "- ¿Cuáles fueron los ingresos extraordinarios de este año y qué parte del resultado representan?\n"
                        "- ¿Quiénes son los 10 mayores deudores y cuánto deben?\n"
                        "- ¿Cómo viene la cuota promedio mes a mes? ¿Y la tasa de cobro?\n"
                        "- Comparame septiembre contra agosto en ingresos y egresos.\n"
                        "- ¿Cuánto pagó en cuotas Kevin Acosta y qué meses debe?")
        for m in st.session_state["chat"]:
            with st.chat_message(m["rol"]):
                st.markdown(m["texto"])
                if m.get("trazas"):
                    with st.expander("Cómo se calculó"):
                        for tz in m["trazas"]:
                            st.code(f"{tz['herramienta']}({tz['argumentos']})" + (f"  ->  ERROR: {tz['error']}" if tz["error"] else ""))
        pregunta = st.chat_input("Escribí tu consulta sobre los datos del club")
        if pregunta:
            if st.session_state["consultas_usadas"] >= max_consultas:
                st.warning(f"Se alcanzó el máximo de {max_consultas} consultas por sesión.")
            else:
                previo = [{"rol": m["rol"], "texto": m["texto"]} for m in st.session_state["chat"]]
                st.session_state["chat"].append({"rol": "user", "texto": pregunta})
                with st.chat_message("user"):
                    st.markdown(pregunta)
                with st.chat_message("assistant"):
                    with st.spinner("Consultando los datos..."):
                        try:
                            texto, trazas, uso = co.Consultor(datos).preguntar(pregunta, previo)
                            co.registrar_consulta(RUTA_CONSULTAS, pregunta, texto, trazas, uso)
                        except Exception as e:
                            texto, trazas = f"No pude responder la consulta ({type(e).__name__}). Revisá la API key y el crédito de la cuenta.", []
                    st.markdown(texto)
                st.session_state["chat"].append({"rol": "assistant", "texto": texto, "trazas": trazas})
                st.session_state["consultas_usadas"] += 1

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
    st.markdown("#### Consultas realizadas")
    regc = co.leer_registro_consultas(RUTA_CONSULTAS)
    if regc.empty:
        st.info("Todavía no se hizo ninguna consulta.")
    else:
        st.dataframe(regc.iloc[::-1], hide_index=True)
        st.download_button("Descargar consultas (CSV)", data=regc.to_csv(index=False).encode("utf-8"),
                           file_name="registro_consultas.csv", mime="text/csv")
