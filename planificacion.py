# -*- coding: utf-8 -*-
"""
Proyecciones y presupuesto - Club Atlético Deportivo Riestra
============================================================

Matemática pura (sin acceso a archivos ni a la IA). Todo se calcula sobre series mensuales
ORDINARIAS: se excluyen los movimientos extraordinarios, porque no se repiten.

PROYECCIÓN
  - Método "tendencia": recta de mínimos cuadrados sobre los meses con datos, con intervalo de
    predicción del 80 %. La tendencia se limita a +40 % / -30 % sobre el promedio de los últimos
    3 meses para evitar extrapolaciones extremas.
  - Método "promedio3": promedio de los últimos 3 meses, sin crecimiento (más conservador).
  - Escenario pesimista = piso de ingresos menos techo de egresos; optimista = lo contrario.
  - Confiabilidad según cuánto se proyecta respecto del historial disponible.
  - Los valores son nominales (incluyen la inflación implícita en la tendencia). No hay
    estacionalidad: se necesitarían al menos 12 meses de historia para estimarla.

PRESUPUESTO
  - Por rubro: tendencia lineal (si hay 4 meses o más; si no, promedio de los últimos 3), con el mismo
    tope de variación (+15 % / -20 % sobre el promedio reciente de cada rubro, más estricto que la proyección).
    Criterio prudente: ingresos -5 %, egresos +5 % y reserva de contingencia del 5 % sobre los egresos.
"""
import numpy as np

Z80 = 1.2816                    # z del intervalo de predicción del 80 %
CAP_BAJA, CAP_ALTA = -0.30, 0.40
MIN_MESES = 3
HORIZONTES = (3, 6, 12)
MAX_HORIZONTE = 36
PRUDENCIA_ING, PRUDENCIA_EGR, RESERVA = 0.05, 0.05, 0.05
PPTO_CAP_BAJA, PPTO_CAP_ALTA = -0.20, 0.15   # el presupuesto es más estricto que la proyección

_MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre",
          "octubre", "noviembre", "diciembre"]
_ABR = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def nombre_mes(p):
    return f"{_MESES[p.month - 1]} {p.year}"


def etiqueta_mes(p):
    return f"{_ABR[p.month - 1]}-{str(p.year)[2:]}"


def _ols(y):
    y = np.asarray(y, float)
    n = len(y)
    x = np.arange(n)
    xm, ym = x.mean(), y.mean()
    sxx = ((x - xm) ** 2).sum()
    b = ((x - xm) * (y - ym)).sum() / sxx
    a = ym - b * xm
    resid = y - (a + b * x)
    s = float(np.sqrt((resid ** 2).sum() / (n - 2))) if n > 2 else 0.0
    return a, b, s, xm, sxx, n


def pronostico(y, h, metodo="tendencia"):
    """Devuelve (base, piso, techo): tres arreglos de largo h, nunca negativos."""
    y = np.asarray(y, float)
    prom3 = float(y[-3:].mean())
    if metodo == "promedio3":
        base = np.full(h, prom3)
        ult = y[-6:]
        s = float(ult.std(ddof=1)) if len(ult) > 1 else 0.0
        lo, hi = base - Z80 * s, base + Z80 * s
    else:
        a, b, s, xm, sxx, n = _ols(y)
        x0 = np.arange(n, n + h)
        base = a + b * x0
        se = s * np.sqrt(1 + 1 / n + ((x0 - xm) ** 2) / sxx)
        lo, hi = base - Z80 * se, base + Z80 * se
        base = np.clip(base, prom3 * (1 + CAP_BAJA), prom3 * (1 + CAP_ALTA))
    lo = np.clip(np.minimum(lo, base), 0, None)
    hi = np.maximum(hi, base)
    return np.clip(base, 0, None), lo, hi


def confiabilidad(h, n_hist):
    r = h / max(n_hist, 1)
    return "Aceptable" if r <= 0.4 else ("Baja" if r <= 1 else "Muy baja")


def proyectar(ing_ord, egr_ord, horizontes, metodo="tendencia"):
    """`ing_ord` / `egr_ord`: pandas Series mensuales (índice Period) de ingresos y egresos ordinarios."""
    n = len(ing_ord)
    if n < MIN_MESES:
        return None
    horizontes = sorted({int(h) for h in horizontes if 1 <= int(h) <= MAX_HORIZONTE})
    hmax = max(horizontes)
    bi, li, hi_ = pronostico(ing_ord.values, hmax, metodo)
    be, le, he = pronostico(egr_ord.values, hmax, metodo)
    ultimo = ing_ord.index[-1]
    serie = []
    for j in range(hmax):
        m = ultimo + (j + 1)
        serie.append({"mes": str(m), "nombre": nombre_mes(m), "etiqueta": etiqueta_mes(m),
                      "ingresos": float(bi[j]), "egresos": float(be[j]), "resultado": float(bi[j] - be[j]),
                      "ingresos_piso": float(li[j]), "ingresos_techo": float(hi_[j]),
                      "egresos_piso": float(le[j]), "egresos_techo": float(he[j]),
                      "res_pesimista": float(li[j] - he[j]), "res_optimista": float(hi_[j] - le[j])})
    resumen = []
    for h in horizontes:
        s = serie[:h]
        ing, egr = sum(x["ingresos"] for x in s), sum(x["egresos"] for x in s)
        resumen.append({"meses": h, "hasta": s[-1]["nombre"], "ingresos": ing, "egresos": egr, "resultado": ing - egr,
                        "margen": (ing - egr) / ing if ing else 0.0,
                        "res_pesimista": sum(x["res_pesimista"] for x in s),
                        "res_optimista": sum(x["res_optimista"] for x in s),
                        "confiabilidad": confiabilidad(h, n)})
    return {"metodo": metodo, "meses_historial": n, "ultimo_mes_historial": nombre_mes(ultimo),
            "desde": serie[0]["nombre"], "serie": serie, "horizontes": resumen}


def _base_rubro(y, h):
    """Valores mensuales base de un rubro para los próximos h meses y promedio reciente."""
    y = np.asarray(y, float)
    prom3 = float(y[-3:].mean())
    if prom3 <= 0:
        return np.zeros(h), 0.0
    if len(y) >= 4:
        a, b, *_ = _ols(y)
        pred = a + b * np.arange(len(y), len(y) + h)
    else:
        pred = np.full(h, prom3)
    return np.clip(pred, prom3 * (1 + PPTO_CAP_BAJA), prom3 * (1 + PPTO_CAP_ALTA)), prom3


def presupuesto(ing_r, egr_r, h, prud_ing=PRUDENCIA_ING, prud_egr=PRUDENCIA_EGR, reserva=RESERVA):
    """`ing_r` / `egr_r`: DataFrames mensuales (filas = meses, columnas = rubros) de movimientos ordinarios."""
    if len(ing_r) < MIN_MESES:
        return None

    def filas(df, factor):
        out = []
        for rubro in df.columns:
            pred, prom3 = _base_rubro(df[rubro].values, h)
            mensual = pred * factor
            out.append({"rubro": str(rubro), "actual_mensual": prom3, "mensual": float(mensual.mean()),
                        "total": float(mensual.sum()), "serie": [float(v) for v in mensual],
                        "variacion": float(mensual.mean() / prom3 - 1) if prom3 else 0.0})
        return sorted([f for f in out if f["total"] > 0 or f["actual_mensual"] > 0], key=lambda f: -f["total"])

    fi, fe = filas(ing_r, 1 - prud_ing), filas(egr_r, 1 + prud_egr)
    ti, te = sum(f["total"] for f in fi), sum(f["total"] for f in fe)
    res_monto = te * reserva
    serie_i = [sum(f["serie"][j] for f in fi) for j in range(h)]
    serie_e = [sum(f["serie"][j] for f in fe) for j in range(h)]
    return {"meses": h, "filas_ingresos": fi, "filas_egresos": fe, "total_ingresos": ti, "total_egresos": te,
            "serie_ingresos": serie_i, "serie_egresos": serie_e,
            "resultado_operativo": ti - te, "reserva": res_monto, "resultado_final": ti - te - res_monto,
            "margen": (ti - te - res_monto) / ti if ti else 0.0,
            "criterio": {"prudencia_ingresos": prud_ing, "prudencia_egresos": prud_egr, "reserva": reserva}}


def ejecucion(pres, real_ing, real_egr):
    """Compara un presupuesto con lo real. `real_*`: dict rubro -> monto del período."""
    def cmp(filas, real):
        rubros = {f["rubro"]: f["total"] for f in filas}
        for r in real:
            rubros.setdefault(r, 0.0)
        out = []
        for r, p in rubros.items():
            v = float(real.get(r, 0.0))
            out.append({"rubro": r, "presupuesto": p, "real": v, "desvio": v - p, "desvio_pct": (v / p - 1) if p else 0.0})
        return sorted(out, key=lambda f: -max(f["presupuesto"], f["real"]))

    ci, ce = cmp(pres["filas_ingresos"], real_ing), cmp(pres["filas_egresos"], real_egr)
    ri, re_ = sum(real_ing.values()), sum(real_egr.values())
    pres_res, real_res = pres["resultado_operativo"], ri - re_
    return {"meses": pres["meses"], "ingresos": ci, "egresos": ce,
            "presupuesto_ingresos": pres["total_ingresos"], "real_ingresos": ri,
            "presupuesto_egresos": pres["total_egresos"], "real_egresos": re_,
            "presupuesto_resultado": pres_res, "real_resultado": real_res,
            "desvio_resultado": real_res - pres_res,
            "desvio_resultado_pct": (real_res / pres_res - 1) if pres_res else 0.0}
