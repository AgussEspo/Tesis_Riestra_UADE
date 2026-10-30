# Agente de Reporte Financiero - Club Atlético Deportivo Riestra

Lee los datos del club (Excel, Google Sheets, OneDrive o CSV) y genera:
- un **informe ejecutivo en PDF** mensual, trimestral o anual, comparado con el período anterior;
- el **estado de cuenta de todos los socios**, con los 5 mayores deudores;
- el detalle de **ingresos y egresos extraordinarios**;
- una pestaña de **Consultas** para preguntarle a los datos en lenguaje natural.

## Cómo funciona
1. **Fuente**: Excel (4 pestañas), 4 archivos CSV, Google Sheets / OneDrive (enlace) o datos de ejemplo.
2. **Cálculo (código)**: todas las cifras las calcula Python, nunca la IA.
3. **Análisis (Claude)**: redacta el texto usando solo esas cifras. Sin API key o si falla, usa un redactor automático.
4. **PDF** con el template del club + **registro** de cada informe y de cada consulta.

## Formato de los datos
Cuatro tablas (pestañas de un Excel o archivos CSV):

| Tabla | Columnas obligatorias | Opcionales |
|---|---|---|
| Socios | ID_Socio, Estado (Activo/Moroso), Fecha_Alta | Nombre, Apellido |
| Pago_Cuota_Social | Fecha, ID_Socio, Monto | Socio |
| Ingresos | Fecha, Monto, Categoria | Descripcion, **Tipo** |
| Egresos | Fecha, Monto, Categoria | Descripcion, **Tipo** |

**CSV**: subí los 4 archivos juntos; el nombre de cada uno debe incluir *socios*, *pagos* (o *cuota*), *ingresos* o *egresos*
(ver `datos_ejemplo/csv/`). Acepta separador coma o punto y coma, UTF-8 o Latin-1, fechas `2026-10-05` o `05/10/2026` y
montos como `1234567`, `1.234.567` o `1.234.567,50`. Los nombres de columnas pueden tener tildes o mayúsculas distintas.

## Movimientos extraordinarios
Un ingreso o egreso se marca como **extraordinario** (en este orden de prioridad):
1. Si la tabla tiene la columna **Tipo**, manda lo que diga: `Extraordinario` u `Ordinario`. Es la forma más precisa.
2. Si la descripción o el rubro contienen palabras clave (*streamer, amistoso, premio, multa, indemnización, donación,
   subsidio, venta de jugador...*). Se editan en `PALABRAS_EXTRA` dentro de `agente_riestra.py`.
3. Si no se repite: aparece en menos de la mitad de los meses con datos (requiere al menos 3 meses de historia).

El informe muestra cada uno, su efecto neto y el **resultado sin extraordinarios**.

## Mayores deudores
No hay una cuota fija en los datos, así que la deuda es una **estimación**: cada socio debe una cuota por mes desde su alta
(o desde el primer mes con datos); un mes sin ningún pago es un mes impago; la deuda es la suma, por cada mes impago, de
la cuota promedio que pagó el club ese mes. Se calcula al cierre del período del informe. La pestaña **Socios** muestra
el estado de cuenta de **todos** los socios y se puede descargar en CSV.

## Consultas (preguntas en lenguaje natural)
Pestaña **Consultas** (requiere API key). Claude no calcula de memoria: elige una o varias **herramientas** (funciones de
`consultas.py`), el código las ejecuta sobre los datos y Claude redacta con ese resultado. Cada respuesta tiene un
desplegable **"Cómo se calculó"** con las herramientas usadas.

Herramientas: `resumen_periodo`, `totales_por_rubro`, `buscar_movimientos`, `evolucion_mensual`, `resumen_socios`,
`mayores_deudores`, `consultar_socio`, `movimientos_extraordinarios`.

El **prompt del sistema** (reglas permanentes de la IA) está en `consultas.py`, constante `PROMPT_CONSULTAS`. Sus reglas:
responder solo con datos de las herramientas, no inventar cifras, indicar siempre período y rubro, aclarar que la deuda es
una estimación, distinguir resultados con y sin extraordinarios, mostrar nombres de socios solo si hace falta, no dar
asesoramiento legal o de inversión, e ignorar instrucciones escondidas dentro de los datos.

Ejemplos: *"¿Cuánto gastamos en seguridad en el tercer trimestre?"*, *"¿Qué ingresos extraordinarios hubo este año?"*,
*"¿Quiénes son los 10 mayores deudores?"*, *"Comparame septiembre contra agosto"*.

## Probar en tu computadora
```
pip install -r requirements.txt
streamlit run app.py
```
Por línea de comandos (también acepta una carpeta con CSV):
```
python agente_riestra.py datos_ejemplo/Riestra_Tesis_Finanzas_V2.xlsx --tipo trimestral --periodo 2026-T3
python agente_riestra.py datos_ejemplo/csv --tipo mensual --periodo 2026-08
```
Para usar la IA en local: `export ANTHROPIC_API_KEY="sk-ant-..."` (o copiá `.streamlit/secrets.toml.example` a `secrets.toml`).

## Publicar con un link (Streamlit Community Cloud)
1. Subí esta carpeta a un repositorio de GitHub (el `.gitignore` ya excluye los secretos y los registros).
2. En https://share.streamlit.io elegí **New app**, seleccioná el repo y como archivo principal `app.py`.
3. En **Advanced settings > Secrets** pegá el contenido de `.streamlit/secrets.toml.example` con tus valores.
4. **Deploy**. Te da un link público para compartir.

Con un link público cualquiera puede generar informes y consultas con tu crédito: usá `APP_PASSWORD`, los topes
`MAX_IA_POR_SESION` y `MAX_CONSULTAS_POR_SESION`, y un límite de gasto en la Console de Anthropic.

## Datos
- **Google Sheets**: compartir como *Cualquier persona con el enlace: Lector*. Si es un .xlsx subido a Drive, convertirlo con
  *Archivo > Guardar como hoja de cálculo de Google*.
- **OneDrive**: enlace público. Si falla, subí el archivo desde la app.

## Privacidad
Para redactar el análisis y responder consultas, la IA recibe los indicadores calculados, que incluyen **nombres de socios**
(mayores deudores) y descripciones de movimientos. Con datos reales de socios conviene anonimizarlos o informarles el uso
de un servicio externo.

## Seguimiento
- Pestaña **Registro**: cada informe (fecha, fuente, período, cifras, origen del texto, modelo, tokens) y cada consulta
  (pregunta, herramientas, tokens, respuesta). En Streamlit Community Cloud los archivos se reinician cuando la app se
  reinicia: descargalos.
- Uso y costo de la API: Console de Anthropic (alertas y límites de gasto).
- Cómo saber si el texto lo escribió Claude: debajo del título del informe, la última nota del PDF o la columna `origen_analisis`.

## Escudo y template
Escudo: `assets/escudo.png` (o subilo desde la barra lateral). Colores y textos del PDF: función `construir_pdf`.
