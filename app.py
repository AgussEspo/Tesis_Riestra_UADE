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
import correo

AQUI = os.path.dirname(os.path.abspath(__file__))
RUTA_EJEMPLO = os.path.join(AQUI, "datos_ejemplo", "Riestra_Tesis_Finanzas_V3.xlsx")
RUTA_REGISTRO = os.path.join(AQUI, "registro_corridas.csv")
RUTA_ENVIOS = os.path.join(AQUI, "registro_envios.csv")
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




def mostrar_proyeccion_tabla(pr):
    st.dataframe(pd.DataFrame([{"Horizonte": f"{h['meses']} meses", "Hasta": h["hasta"].capitalize(), "Ingresos": ag.ars_m(h["ingresos"]),
                                "Egresos": ag.ars_m(h["egresos"]), "Resultado": ag.ars_m(h["resultado"]),
                                "Pesimista": ag.ars_m(h["res_pesimista"]), "Optimista": ag.ars_m(h["res_optimista"]),
                                "Confiabilidad": h["confiabilidad"]} for h in pr["horizontes"]]), hide_index=True)


def mostrar_caja(k, graficos):
    """Flujo de caja histórico del período."""
    rp = k["caja"]["periodo"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Saldo inicial", ag.ars_m(rp["saldo_inicial"]))
    c2.metric("Variación neta", ag.ars_m(rp["neto"]), "operativo + inversión + financiación", delta_color="off")
    c3.metric("Saldo final", ag.ars_m(rp["saldo_final"]), f"mínimo {ag.ars_m(rp['minimo']['saldo'])}" if rp["meses"] > 1 else None, delta_color="off")
    cob = rp.get("cobertura")
    c4.metric("Cobertura de caja", (f"{cob:.1f}".replace(".", ",") + " meses") if cob is not None else "n/c", "de egresos y cuotas de deuda", delta_color="off")
    mostrar_imagen(graficos["caja"][0])
    filas = [{"Mes": f["nombre"].capitalize() + ("" if f["en_periodo"] else " (contexto)"), "Saldo inicial": ag.ars_m(f["saldo_inicial"]),
              "Operativo": ag.ars_m(f["operativo"]), "Inversión": ag.ars_m(f["inversion"]), "Financiación": ag.ars_m(f["financiacion"]),
              "Variación": ag.ars_m(f["neto"]), "Saldo final": ag.ars_m(f["saldo_final"])} for f in rp["filas"]]
    if rp["meses"] > 1:
        filas.append({"Mes": "Total del período", "Saldo inicial": ag.ars_m(rp["saldo_inicial"]), "Operativo": ag.ars_m(rp["operativo"]),
                      "Inversión": ag.ars_m(rp["inversion"]), "Financiación": ag.ars_m(rp["financiacion"]),
                      "Variación": ag.ars_m(rp["neto"]), "Saldo final": ag.ars_m(rp["saldo_final"])})
    st.dataframe(pd.DataFrame(filas), hide_index=True)
    if len(rp["cuentas_final"]) > 1:
        st.caption("Saldo final por cuenta: " + " · ".join(f"{c}: {ag.ars_m(v)}" for c, v in rp["cuentas_final"].items()))


def mostrar_caja_proyectada(k, graficos):
    """Flujo de caja proyectado: base, escenarios y compromisos considerados."""
    pj = k["caja"]["proyeccion"]
    mostrar_imagen(graficos["caja_proyeccion"][0])
    st.markdown("**Saldo de caja proyectado por horizonte**")
    st.dataframe(pd.DataFrame([{"Horizonte": f"{h['meses']} meses", "Hasta": h["hasta"].capitalize(), "Flujo neto": ag.ars_m(h["flujo_neto"]),
                                "Saldo (base)": ag.ars_m(h["saldo"]), "Pesimista": ag.ars_m(h["saldo_pesimista"]),
                                "Optimista": ag.ars_m(h["saldo_optimista"]), "Saldo mínimo": ag.ars_m(h["saldo_minimo"])}
                               for h in pj["horizontes"]]), hide_index=True)
    st.markdown("**Detalle mes a mes (escenario base)**")
    st.dataframe(pd.DataFrame([{"Mes": f["nombre"].capitalize(), "Operativo": ag.ars_m(f["operativo"]), "Inversión": ag.ars_m(f["inversion"]),
                                "Financiación": ag.ars_m(f["financiacion"]), "Variación": ag.ars_m(f["neto"]), "Saldo (base)": ag.ars_m(f["saldo"]),
                                "Pesimista": ag.ars_m(f["saldo_pesimista"]), "Optimista": ag.ars_m(f["saldo_optimista"])}
                               for f in pj["serie"][:12]]), hide_index=True)
    if pj["compromisos"]:
        with st.expander(f"Compromisos considerados ({len(pj['compromisos'])})"):
            st.dataframe(pd.DataFrame([{"Fecha": r["fecha"], "Tipo": r["tipo"], "Concepto": r["concepto"], "Actividad": r["actividad"],
                                        "Estado": r["estado"], "Monto": ag.ars_m(r["monto"])} for r in pj["compromisos"]]), hide_index=True)
    st.caption("Escenario base: proyección ordinaria más todos los compromisos. Pesimista: sin los cobros probables; optimista: sin los pagos probables.")


def mostrar_presupuesto(pp, ej):
    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"Ingresos ({pp['meses']} meses)", ag.ars_m(pp["total_ingresos"]))
    c2.metric("Egresos", ag.ars_m(pp["total_egresos"]))
    c3.metric("Reserva de contingencia", ag.ars_m(pp["reserva"]))
    c4.metric("Resultado presupuestado", ag.ars_m(pp["resultado_final"]))
    tabla = lambda filas: pd.DataFrame([{"Rubro": f["rubro"], "Actual (prom. mensual)": ag.ars_m(f["actual_mensual"]),
                                         "Presupuesto mensual": ag.ars_m(f["mensual"]), f"Total {pp['meses']} meses": ag.ars_m(f["total"]),
                                         "Variación": ag.signo_pct(f["variacion"], 0)} for f in filas])
    t1, t2 = st.columns(2)
    t1.markdown("**Ingresos por rubro**")
    t1.dataframe(tabla(pp["filas_ingresos"]), hide_index=True)
    t2.markdown("**Egresos por rubro**")
    t2.dataframe(tabla(pp["filas_egresos"]), hide_index=True)
    if pp.get("trimestres"):
        st.markdown("**Presupuesto por trimestre**")
        st.dataframe(pd.DataFrame([{"Trimestre": t["nombre"], "Ingresos": ag.ars_m(t["ingresos"]), "Egresos": ag.ars_m(t["egresos"]),
                                    "Reserva": ag.ars_m(t["reserva"]), "Resultado": ag.ars_m(t["resultado"])} for t in pp["trimestres"]]), hide_index=True)
    if ej:
        st.markdown("**Ejecución: presupuesto armado con datos previos vs. real del período**")
        filas = [{"Bloque": b, "Rubro": f["rubro"], "Presupuesto": ag.ars_m(f["presupuesto"]), "Real": ag.ars_m(f["real"]),
                  "Desvío": ag.ars_m(f["desvio"]), "Desvío %": ag.signo_pct(f["desvio_pct"], 0) if f["presupuesto"] else "n/c"}
                 for b, lista in (("Ingresos", ej["ingresos"]), ("Egresos", ej["egresos"])) for f in lista if max(f["presupuesto"], f["real"]) > 0]
        filas.append({"Bloque": "Resultado", "Rubro": "Total", "Presupuesto": ag.ars_m(ej["presupuesto_resultado"]), "Real": ag.ars_m(ej["real_resultado"]),
                      "Desvío": ag.ars_m(ej["desvio_resultado"]), "Desvío %": ag.signo_pct(ej["desvio_resultado_pct"], 0)})
        st.dataframe(pd.DataFrame(filas), hide_index=True)
    st.caption("Presupuesto prudente: ingresos -5%, egresos +5% y reserva del 5% sobre la tendencia de cada rubro. "
               "Es una propuesta elaborada a partir de los datos, no un presupuesto aprobado.")

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

cfg_correo = correo.leer_configuracion(secreto)
st.session_state.setdefault("envios", 0)

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
        detalle = st.radio("Nivel de detalle del PDF", ["completo", "resumido"], horizontal=True, format_func=str.capitalize,
                           help="El resumido son 2 o 3 páginas con lo esencial; el completo incluye todo el detalle.")
        opciones = ag.periodos_disponibles(datos)[tipo]
        elegido = st.selectbox("Período", opciones, format_func=lambda o: o[0])
        st.subheader("3. Proyección y presupuesto")
        proy_extra = st.number_input("Proyección a mi gusto (meses, 0 = ninguna)", min_value=0, max_value=36, value=0, step=1,
                                     help="Se suma a las proyecciones fijas de 3, 6 y 12 meses.")
        metodo = st.selectbox("Método de proyección", list(ag.METODOS), format_func=lambda m: ag.METODOS[m].capitalize())
        meses_ppto = st.selectbox("Meses del presupuesto", [12, 6, 3])
        generar = st.button("Generar informe", type="primary")
    else:
        proy_extra, metodo, meses_ppto, detalle = 0, "tendencia", 12, "completo"
        generar = False

# ----------------------------------------------------------------------------
# Generación
# ----------------------------------------------------------------------------
if generar and datos is not None:
    try:
        with st.spinner("Calculando indicadores y redactando el informe..."):
            clave = elegido[1]
            k = ag.calcular_kpis(datos, tipo, clave, int(proy_extra) or None, metodo, int(meses_ppto))
            a = ag.generar_analisis(k)
            carpeta = tempfile.mkdtemp()
            graficos = ag.generar_graficos(k, carpeta)
            ruta_pdf = os.path.join(carpeta, "informe.pdf")
            logo = ag.buscar_logo(CARPETA_ASSETS)
            ag.construir_pdf(k, a, ruta_pdf, logo, graficos, detalle)
            with open(ruta_pdf, "rb") as f:
                contenido_pdf = f.read()
            ag.registrar_corrida(RUTA_REGISTRO, st.session_state["fuente_txt"], k, a)
        st.session_state["informe"] = {"k": k, "a": a, "graficos": graficos,
                                       "pdf": contenido_pdf, "nombre": f"Reporte_{tipo}_{clave}" + ("_resumen" if detalle == "resumido" else "") + ".pdf"}
    except Exception as e:
        st.error(f"No se pudo generar el informe: {e}")

# ----------------------------------------------------------------------------
# Resultado
# ----------------------------------------------------------------------------
inf = st.session_state.get("informe")
tab_inf, tab_plan, tab_socios, tab_datos, tab_reg = st.tabs(["Informe", "Planificación", "Socios", "Datos y calidad", "Registro"])

with tab_inf:
    if not inf:
        st.info("Elegí el tipo de informe y el período en la barra lateral y presioná **Generar informe**.")
    else:
        k, a = inf["k"], inf["a"]
        t, comp = k["totales"], k["comparacion"]
        c1, c2 = st.columns([3, 1])
        c1.subheader(f"{k['periodo']['tipo'].capitalize()}: {k['periodo']['etiqueta']}")
        c2.download_button("Descargar PDF", data=inf["pdf"], file_name=inf["nombre"], mime="application/pdf")
        with st.expander("Enviar el PDF por correo electrónico"):
            if cfg_correo is None:
                st.info("El envío por correo no está configurado. Hay que cargar `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER` y `SMTP_PASSWORD` "
                        "en los secretos de la app (ver `.streamlit/secrets.toml.example`).")
            else:
                destino = st.text_input("Destinatario(s)", placeholder="nombre@dominio.com, otra@dominio.com",
                                        help=f"Hasta {correo.MAX_DESTINATARIOS} direcciones separadas por coma.", key="mail_dest")
                asunto = st.text_input("Asunto", value=f"Reporte financiero Riestra - {k['periodo']['etiqueta']}", key="mail_asunto")
                if st.button("Enviar PDF", key="mail_enviar"):
                    validos, invalidos = correo.separar_destinatarios(destino)
                    if invalidos:
                        st.error("Direcciones no válidas: " + ", ".join(invalidos))
                    elif st.session_state["envios"] >= cfg_correo["max_sesion"]:
                        st.warning(f"Se alcanzó el máximo de {cfg_correo['max_sesion']} envíos por sesión.")
                    else:
                        try:
                            with st.spinner("Enviando..."):
                                correo.enviar_pdf(cfg_correo, validos, asunto.strip() or "Reporte financiero", correo.cuerpo_por_defecto(k, a["introduccion"]),
                                                  inf["pdf"], inf["nombre"])
                            st.session_state["envios"] += 1
                            correo.registrar_envio(RUTA_ENVIOS, validos, inf["nombre"], "enviado")
                            st.success("Informe enviado a: " + ", ".join(validos))
                        except correo.ErrorCorreo as e:
                            correo.registrar_envio(RUTA_ENVIOS, validos, inf["nombre"], "error", str(e))
                            st.error(str(e))
                        except Exception as e:
                            correo.registrar_envio(RUTA_ENVIOS, validos, inf["nombre"], "error", type(e).__name__)
                            st.error(f"No se pudo enviar el correo ({type(e).__name__}).")

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
        _n = iter(range(1, 100))
        sec = lambda titulo: st.markdown(f"**2.{next(_n)} {titulo}**")
        sec("Ingresos")
        st.write(a["ingresos"])
        sec("Egresos")
        st.write(a["egresos"])
        mostrar_imagen(inf["graficos"]["composicion"][0])
        sec("Resultado y evolución")
        st.write(a["resultado"])
        mostrar_imagen(inf["graficos"]["evolucion"][0])
        if k.get("caja"):
            sec("Flujo de caja")
            st.write(a["caja"])
            mostrar_caja(k, inf["graficos"])
        sec("Movimientos extraordinarios")
        st.write(a["extraordinarios"])
        xs = k["extraordinarios"]
        if xs["cantidad"]:
            st.dataframe(pd.DataFrame(xs["ingresos"] + xs["egresos"])[["fecha", "tipo", "rubro", "descripcion", "monto", "motivo"]]
                         .rename(columns=str.capitalize), hide_index=True)
            e1, e2, e3 = st.columns(3)
            e1.metric("Resultado del período", ag.ars_m(t["resultado"]))
            e2.metric("Efecto de extraordinarios", ag.ars_m(xs["neto"]))
            e3.metric("Resultado sin extraordinarios", ag.ars_m(xs["resultado_ordinario"]))
        sec("Socios y cuota social")
        st.write(a["socios"])
        mostrar_imagen(inf["graficos"]["socios"][0])
        dd = k["deudores"]
        sec(f"Los {len(dd['top'])} mayores deudores")
        st.write(a["deudores"])
        if dd["top"]:
            st.dataframe(pd.DataFrame([{"#": r["puesto"], "Socio": f"{r['socio']} (N° {r['id']})", "Estado": r["estado"],
                                        "Meses impagos": r["meses_impagos"], "Deuda estimada": r["deuda"],
                                        "Último pago": r["ultimo_pago"]} for r in dd["top"]]), hide_index=True)
        sec("Presupuesto")
        st.write(a["presupuesto"])
        if k.get("presupuesto"):
            mostrar_presupuesto(k["presupuesto"], k.get("ejecucion"))
        sec("Puntos de atención")
        for x in a["alertas"]:
            st.markdown(f"- {x}")
        sm = k["semaforo"]
        circ = {"verde": "🟢", "amarillo": "🟡", "rojo": "🔴"}
        sec("Semáforo de indicadores")
        st.write(a["semaforo"])
        st.markdown(f"**Situación general:** {circ[sm['general']]} &nbsp;&nbsp;&nbsp; **Indicadores:** "
                    f"🟢 {sm['conteo']['verde']} &nbsp;&nbsp; 🟡 {sm['conteo']['amarillo']} &nbsp;&nbsp; 🔴 {sm['conteo']['rojo']}")
        st.dataframe(pd.DataFrame([{"Estado": circ[i["estado"]], "Indicador": i["indicador"], "Valor": i["valor"], "Criterio": i["criterio"]}
                                   for i in sm["items"]]), hide_index=True)
        sec("Proyección")
        st.write(a["proyeccion"])
        if k.get("proyeccion"):
            mostrar_imagen(inf["graficos"]["proyeccion"][0])
            mostrar_proyeccion_tabla(k["proyeccion"])
        if k.get("caja") and k["caja"].get("proyeccion"):
            st.markdown("**Flujo de caja proyectado**")
            st.write(a["caja_proyeccion"])
            mostrar_caja_proyectada(k, inf["graficos"])
        st.markdown("### 3. Cierre")
        st.write(a["cierre"])
        st.markdown("**Recomendaciones**")
        for x in a["recomendaciones"]:
            st.markdown(f"- {x}")

with tab_plan:
    if datos is None:
        st.info("Cargá los datos para proyectar y presupuestar.")
    else:
        st.caption("Explorá proyecciones y presupuestos al instante, sin generar el PDF. Se calculan solo con movimientos "
                   "ordinarios (sin extraordinarios), en pesos nominales y sin estacionalidad.")
        fmin_, fmax_ = ag.rango_datos(datos)
        p1, p2, p3 = st.columns(3)
        h_user = p1.number_input("Meses a proyectar", min_value=1, max_value=36, value=6, step=1, key="plan_h")
        m_user = p2.selectbox("Método", list(ag.METODOS), format_func=lambda m: ag.METODOS[m].capitalize(), key="plan_m")
        b_user = p3.selectbox("Meses del presupuesto", [12, 6, 3], key="plan_b")
        res_plan = ag.calcular_planificacion(datos, fmax_, None, int(h_user), m_user, int(b_user))
        if not res_plan["proyeccion"]:
            st.warning("No hay historial suficiente: se necesitan al menos 3 meses con datos.")
        else:
            pr_ = res_plan["proyeccion"]
            if res_plan["nota_mes_incompleto"]:
                st.caption(res_plan["nota_mes_incompleto"])
            hx = next(h for h in pr_["horizontes"] if h["meses"] == int(h_user))
            q1, q2, q3, q4 = st.columns(4)
            q1.metric(f"Ingresos a {hx['meses']} meses", ag.ars_m(hx["ingresos"]))
            q2.metric("Egresos", ag.ars_m(hx["egresos"]))
            q3.metric("Resultado", ag.ars_m(hx["resultado"]), f"Pesimista {ag.ars_m(hx['res_pesimista'])}", delta_color="off")
            q4.metric("Confiabilidad", hx["confiabilidad"], f"{pr_['meses_historial']} meses de historial", delta_color="off")
            ruta_g = os.path.join(tempfile.mkdtemp(), "proy_live.png")
            ag.grafico_proyeccion(pr_, ruta_g, max_meses=max(int(h_user), 6))
            mostrar_imagen(ruta_g)
            mostrar_proyeccion_tabla(pr_)
            st.markdown(f"#### Presupuesto propuesto a {int(b_user)} meses")
            if res_plan["presupuesto"]:
                mostrar_presupuesto(res_plan["presupuesto"], None)

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
        for act, v in q.get("no_operativos", {}).items():
            det = "" if act == "Transferencia interna" else f" (ingresos {ag.ars_m(v['ingresos'])}, egresos {ag.ars_m(v['egresos'])})"
            st.write(f"- {v['n']} movimientos de {act.lower()}{det} quedaron fuera del resultado y se informan solo en el flujo de caja.")
        avisos = list(q.get("avisos_carga", [])) + (k["caja"]["avisos"] if k.get("caja") else [])
        for av in dict.fromkeys(avisos):
            st.warning(av)
        if not k.get("caja"):
            st.info("No se muestra el flujo de caja: falta la hoja **Saldos** (fecha, cuenta y saldo inicial). Para el flujo proyectado, "
                    "sumá también la hoja **Compromisos**.")
        elif k["caja"]["proyeccion"] and not k["caja"].get("hay_compromisos"):
            st.info("El flujo de caja proyectado no incluye compromisos: falta la hoja **Compromisos** (vencimientos futuros).")
        with st.expander("Indicadores calculados"):
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
    st.markdown("#### Envíos por correo")
    regm = correo.leer_envios(RUTA_ENVIOS)
    if regm.empty:
        st.info("Todavía no se envió ningún informe por correo.")
    else:
        st.dataframe(regm.iloc[::-1], hide_index=True)
        st.download_button("Descargar envíos (CSV)", data=regm.to_csv(index=False).encode("utf-8"),
                           file_name="registro_envios.csv", mime="text/csv")
