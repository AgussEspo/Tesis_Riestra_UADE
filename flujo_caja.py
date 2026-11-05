# -*- coding: utf-8 -*-
"""
Flujo de caja - Club Atlético Deportivo Riestra
===============================================

Matemática pura (sin acceso a archivos ni a la interfaz). Método DIRECTO: se suman los cobros y pagos reales
del club, mes a mes, a partir de un saldo inicial.

CRITERIOS
  - La columna Fecha de Ingresos / Egresos es la fecha EFECTIVA en que se movió el dinero (criterio de caja).
  - Cada movimiento tiene una Actividad:
        Operativa              actividad habitual (cuotas, sponsors, TV, entradas, sueldos, servicios...)
        Inversión              obras, equipamiento, compra de pases
        Financiación           préstamos recibidos y sus cuotas
        Transferencia interna  pase de plata entre cuentas del propio club (no es ingreso ni gasto)
    Solo la actividad Operativa forma parte del RESULTADO del período; las otras tres se ven únicamente acá.
  - Las cuotas sociales (hoja Pago_Cuota_Social) son cobros operativos en la cuenta Banco.
  - Saldo del mes = saldo inicial + flujo operativo + inversión + financiación (+ transferencias, que suman cero).

FLUJO PROYECTADO
  - Parte del saldo al cierre del último mes histórico.
  - Suma, mes a mes, la proyección ORDINARIA de ingresos y egresos (la de planificacion.py) y los COMPROMISOS
    conocidos (hoja Compromisos: vencimientos futuros que no se repiten todos los meses).
  - Escenario base: todos los compromisos. Pesimista: piso de ingresos, techo de egresos, sin los cobros
    "Probables" y con todos los pagos. Optimista: lo contrario.
"""
import re
import unicodedata

import pandas as pd

import planificacion as pl

ACT_OPERATIVA, ACT_INVERSION = "Operativa", "Inversión"
ACT_FINANCIACION, ACT_TRANSFERENCIA = "Financiación", "Transferencia interna"
ACTIVIDADES = (ACT_OPERATIVA, ACT_INVERSION, ACT_FINANCIACION, ACT_TRANSFERENCIA)
CUENTA_DEFECTO = "Banco"
MESES_REFERENCIA = 3                      # meses usados para el egreso mensual de referencia
COBERTURA_VERDE, COBERTURA_ROJO = 3.0, 1.0


def _norm(txt):
    t = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "_", t.lower()).strip("_")


def normalizar_actividad(valor):
    """Devuelve una de ACTIVIDADES; vacío = Operativa; None si el valor no se reconoce."""
    t = _norm(valor)
    if t in ("", "nan", "none", "nat"):
        return ACT_OPERATIVA
    if t.startswith(("operativ", "operacion", "explotacion")):
        return ACT_OPERATIVA
    if t.startswith("inversi"):
        return ACT_INVERSION
    if t.startswith("financ"):
        return ACT_FINANCIACION
    if t.startswith("transfer") or "interna" in t:
        return ACT_TRANSFERENCIA
    return None


def normalizar_cuentas(serie):
    """Cuentas sin espacios y con una sola grafía por nombre (Banco / banco / BANCO); vacío = Banco."""
    vistas, out = {}, []
    for v in serie.astype(str):
        s = v.strip()
        if _norm(s) in ("", "nan", "none"):
            s = CUENTA_DEFECTO
        out.append(vistas.setdefault(_norm(s), s))
    return pd.Series(out, index=serie.index)


def normalizar_tipo_compromiso(valor):
    t = _norm(valor)
    if t.startswith(("cobr", "ingres")):
        return "Cobro"
    if t.startswith(("pag", "egres")):
        return "Pago"
    return None


def normalizar_estado_compromiso(valor):
    t = _norm(valor)
    return "Probable" if t.startswith(("probab", "estimad", "tentativ", "pendiente_de_confirmar")) else "Confirmado"


# ----------------------------------------------------------------------------
# HISTÓRICO
# ----------------------------------------------------------------------------
def _movimientos(ing, egr, pagos):
    """Un solo DataFrame con Fecha, Cuenta, Actividad y Monto con signo (cobro +, pago -)."""
    partes = []
    for df, signo in ((ing, 1), (egr, -1)):
        partes.append(pd.DataFrame({"Fecha": df["Fecha"].values, "Cuenta": df["Cuenta"].values,
                                    "Actividad": df["Actividad"].values, "Monto": signo * df["Monto"].astype(float).values}))
    partes.append(pd.DataFrame({"Fecha": pagos["Fecha"].values, "Cuenta": CUENTA_DEFECTO,
                                "Actividad": ACT_OPERATIVA, "Monto": pagos["Monto"].astype(float).values}))
    return pd.concat(partes, ignore_index=True)


def historico(ing, egr, pagos, saldos, hasta):
    """Flujo de caja mensual desde el último saldo conocido hasta `hasta`. None si no hay saldos.

    Devuelve {"filas": [...], "cuentas": [...], "avisos": [...], "referencia": {...}}."""
    if saldos is None or len(saldos) == 0:
        return None
    hasta = pd.Timestamp(hasta).normalize()
    s = saldos.sort_values("Fecha").groupby("Cuenta").tail(1)
    fecha_saldo = {r.Cuenta: pd.Timestamp(r.Fecha).normalize() for r in s.itertuples()}
    inicial = {r.Cuenta: float(r.Saldo) for r in s.itertuples()}
    f0 = min(fecha_saldo.values())
    avisos = []
    mov = _movimientos(ing, egr, pagos)
    for c in sorted(set(mov["Cuenta"]) - set(inicial)):
        inicial[c], fecha_saldo[c] = 0.0, f0
        avisos.append(f"La cuenta '{c}' tiene movimientos pero no tiene saldo inicial en la hoja Saldos: se asumió saldo 0.")
    mov = mov[(mov["Fecha"] <= hasta + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))].copy()
    mov = mov[mov["Fecha"] > mov["Cuenta"].map(fecha_saldo)]
    if hasta <= f0:
        return None
    meses = pd.period_range((f0 + pd.Timedelta(days=1)).to_period("M"), hasta.to_period("M"), freq="M")
    mov["mes"] = mov["Fecha"].dt.to_period("M")
    cuentas = sorted(inicial)
    saldo_c = dict(inicial)
    filas = []
    for m in meses:
        mm = mov[mov["mes"] == m]
        por_act = mm.groupby("Actividad")["Monto"].sum()
        op, inv, fin, tr = (float(por_act.get(a, 0.0)) for a in ACTIVIDADES)
        ini_tot = sum(saldo_c.values())
        for c, v in mm.groupby("Cuenta")["Monto"].sum().items():
            saldo_c[c] += float(v)
        filas.append({"mes": str(m), "etiqueta": pl.etiqueta_mes(m), "nombre": pl.nombre_mes(m),
                      "saldo_inicial": ini_tot, "operativo": op, "inversion": inv, "financiacion": fin, "transferencias": tr,
                      "neto": op + inv + fin + tr, "saldo_final": sum(saldo_c.values()), "cuentas": dict(saldo_c)})
    if not filas:
        return None
    if any(abs(f["transferencias"]) > 1 for f in filas):
        avisos.append("Las transferencias internas no suman cero: revisá que cada traspaso tenga su salida y su entrada.")
    if any(f["saldo_final"] < 0 for f in filas):
        avisos.append("Hay meses con saldo de caja negativo: revisá el saldo inicial y los movimientos.")
    # egreso mensual de referencia: egresos operativos y cuotas de deuda de los últimos meses
    sal = mov[(mov["Monto"] < 0) & mov["Actividad"].isin([ACT_OPERATIVA, ACT_FINANCIACION])]
    por_mes = (-sal.groupby("mes")["Monto"].sum()).reindex(meses, fill_value=0.0)
    ref = float(por_mes.iloc[-MESES_REFERENCIA:].mean()) if len(por_mes) else 0.0
    return {"filas": filas, "cuentas": cuentas, "avisos": avisos, "egreso_referencia": ref}


def cobertura_meses(saldo, egreso_mensual):
    return saldo / egreso_mensual if egreso_mensual and egreso_mensual > 0 else None


def resumen_periodo(h, desde, hasta):
    """Recorta el histórico al período del informe. None si el período no tiene meses con saldo."""
    if not h:
        return None
    p0, p1 = pd.Timestamp(desde).to_period("M"), pd.Timestamp(hasta).to_period("M")
    filas = h["filas"]
    sel = [f for f in filas if p0 <= pd.Period(f["mes"]) <= p1]
    if not sel:
        return None
    n = len(sel)
    out = {"saldo_inicial": sel[0]["saldo_inicial"], "saldo_final": sel[-1]["saldo_final"],
           "operativo": sum(f["operativo"] for f in sel), "inversion": sum(f["inversion"] for f in sel),
           "financiacion": sum(f["financiacion"] for f in sel), "neto": sum(f["neto"] for f in sel),
           "cuentas_final": sel[-1]["cuentas"], "meses": n}
    pmin = min(sel, key=lambda f: f["saldo_final"])
    out["minimo"] = {"saldo": pmin["saldo_final"], "mes": pmin["nombre"]}
    out["cobertura"] = cobertura_meses(out["saldo_final"], h["egreso_referencia"])
    out["egreso_referencia"] = h["egreso_referencia"]
    # tabla: los meses del período más contexto previo (como en los gráficos del informe)
    n_tab = max(6, n)
    idx0 = filas.index(sel[0])
    ult = filas.index(sel[-1])
    ini = max(0, ult - (n_tab - 1))
    ctx = filas[ini:ult + 1]
    out["filas"] = [dict(f, en_periodo=f in sel) for f in ctx]
    # período anterior de igual duración
    previas = filas[max(0, idx0 - n):idx0]
    if len(previas) == n:
        out["anterior"] = {"saldo_final": previas[-1]["saldo_final"], "neto": sum(f["neto"] for f in previas),
                           "operativo": sum(f["operativo"] for f in previas)}
    else:
        out["anterior"] = None
    return out


# ----------------------------------------------------------------------------
# PROYECTADO
# ----------------------------------------------------------------------------
def proyectar(saldo0, ultimo, serie, horizontes, compromisos=None):
    """Flujo de caja de los próximos meses.

    saldo0: saldo al cierre de `ultimo` (pd.Period). serie: lista de planificacion.proyectar()['serie'].
    horizontes: lista de meses (3, 6, 12 y el opcional). compromisos: DataFrame ya normalizado o None."""
    n = len(serie)
    if not n:
        return None
    meses = [ultimo + (j + 1) for j in range(n)]
    corte_inicio = meses[0].start_time
    corte_fin = meses[-1].end_time
    comp = pd.DataFrame()
    ignorados = fuera = 0
    if compromisos is not None and len(compromisos):
        c = compromisos[compromisos["Actividad"] != ACT_TRANSFERENCIA].copy()
        ignorados = int((c["Fecha_Vencimiento"] < corte_inicio).sum())
        fuera = int((c["Fecha_Vencimiento"] > corte_fin).sum())
        comp = c[(c["Fecha_Vencimiento"] >= corte_inicio) & (c["Fecha_Vencimiento"] <= corte_fin)].copy()
        comp["mes"] = comp["Fecha_Vencimiento"].dt.to_period("M")
        comp["firmado"] = comp["Monto"].astype(float) * comp["Tipo"].map({"Cobro": 1.0, "Pago": -1.0})
        prob = comp["Estado"] == "Probable"
        comp["w_pes"] = ((comp["Tipo"] == "Cobro") & prob).map({True: 0.0, False: 1.0})
        comp["w_opt"] = ((comp["Tipo"] == "Pago") & prob).map({True: 0.0, False: 1.0})

    base, pes, opt = saldo0, saldo0, saldo0
    filas = []
    for j, m in enumerate(meses):
        sj = serie[j]
        cm = comp[comp["mes"] == m] if len(comp) else comp
        por_act = {a: float(cm.loc[cm["Actividad"] == a, "firmado"].sum()) if len(cm) else 0.0 for a in ACTIVIDADES}
        c_pes = float((cm["firmado"] * cm["w_pes"]).sum()) if len(cm) else 0.0
        c_opt = float((cm["firmado"] * cm["w_opt"]).sum()) if len(cm) else 0.0
        c_all = float(cm["firmado"].sum()) if len(cm) else 0.0
        op = sj["resultado"] + por_act[ACT_OPERATIVA]
        neto = op + por_act[ACT_INVERSION] + por_act[ACT_FINANCIACION]
        base += neto
        pes += sj["res_pesimista"] + c_pes
        opt += sj["res_optimista"] + c_opt
        filas.append({"mes": str(m), "nombre": pl.nombre_mes(m), "etiqueta": pl.etiqueta_mes(m),
                      "operativo": op, "inversion": por_act[ACT_INVERSION], "financiacion": por_act[ACT_FINANCIACION],
                      "compromisos": c_all, "neto": neto, "saldo": base, "saldo_pesimista": pes, "saldo_optimista": opt})
    hz = []
    for h in sorted({int(x) for x in horizontes if 1 <= int(x) <= n}):
        f = filas[:h]
        pmin = min(f, key=lambda r: r["saldo"])
        negativos = [r["nombre"] for r in f if r["saldo_pesimista"] < 0]
        hz.append({"meses": h, "hasta": f[-1]["nombre"], "flujo_neto": sum(r["neto"] for r in f), "saldo": f[-1]["saldo"],
                   "saldo_pesimista": f[-1]["saldo_pesimista"], "saldo_optimista": f[-1]["saldo_optimista"],
                   "saldo_minimo": pmin["saldo"], "mes_minimo": pmin["nombre"],
                   "saldo_minimo_pesimista": min(r["saldo_pesimista"] for r in f),
                   "meses_negativos_pesimista": len(negativos), "primer_mes_negativo_pesimista": negativos[0] if negativos else None})
    detalle = []
    if len(comp):
        for r in comp.sort_values("Fecha_Vencimiento").itertuples():
            detalle.append({"fecha": r.Fecha_Vencimiento.strftime("%d/%m/%Y"), "tipo": r.Tipo, "monto": float(r.Monto),
                            "concepto": str(getattr(r, "Descripcion", "") or ""), "categoria": str(getattr(r, "Categoria", "") or ""),
                            "actividad": r.Actividad, "estado": r.Estado})
    return {"saldo_inicial": saldo0, "desde": filas[0]["nombre"], "ultimo_mes_historial": pl.nombre_mes(ultimo),
            "serie": filas, "horizontes": hz, "compromisos": detalle,
            "n_confirmados": sum(1 for d in detalle if d["estado"] == "Confirmado"),
            "n_probables": sum(1 for d in detalle if d["estado"] == "Probable"),
            "ignorados_vencidos": ignorados, "fuera_de_horizonte": fuera}
