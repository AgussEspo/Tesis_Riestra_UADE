# Agente de Reporte Financiero - Club Atlético Deportivo Riestra

Lee los datos del club (Excel, Google Sheets, OneDrive o CSV) y genera:
- un **informe ejecutivo en PDF** mensual, trimestral o anual, comparado con el período anterior;
- el **estado de cuenta de todos los socios**, con los 5 mayores deudores;
- el detalle de **ingresos y egresos extraordinarios**;
- un **semáforo** de 13 indicadores (verde = bien, amarillo = regular, rojo = mal);
- una **proyección** a 3, 6 y 12 meses (más un horizonte opcional a elección, de 1 a 36);
- un **presupuesto** propuesto por rubro, con su ejecución contra lo real;
- el **envío del PDF por correo electrónico** directo desde la app (además de descargarlo).

## Cómo funciona
1. **Fuente**: Excel (4 pestañas), 4 archivos CSV, Google Sheets / OneDrive (enlace) o datos de ejemplo.
2. **Cálculo (código)**: todas las cifras las calcula Python.
3. **Análisis (redacción automática)**: un redactor local arma Introducción, Desarrollo y Cierre solo con esas cifras. No usa ningún servicio externo ni API key.
4. **PDF** con el template del club + **registro** de cada informe y de cada envío por correo.

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

## Semáforo
13 indicadores en 4 grupos, con el criterio visible en el informe. Los umbrales son criterios de gestión orientativos
(editables en `calcular_semaforo`, en `agente_riestra.py`):

| Grupo | Indicador | Verde | Rojo |
|---|---|---|---|
| Resultados | Margen del período | ≥ 20% | < 5% |
| | Cobertura de egresos (ingresos / egresos) | ≥ 1,2 | < 1,0 |
| | Ingresos vs período anterior | ≥ 0% | < -5% |
| | Egresos vs ingresos (ritmo de crecimiento) | egresos crecen igual o menos | > 5 pts por encima |
| Estructura | Concentración de ingresos (rubro mayor) | < 30% | > 50% |
| | Concentración de egresos (rubro mayor) | < 40% | > 60% |
| | Dependencia de extraordinarios | ≤ 10% del resultado | > 25% |
| Socios | Morosidad del padrón | ≤ 10% | > 20% |
| | Tasa de cobro de cuotas | ≥ 90% | < 75% |
| | Cuota social sobre egresos | ≥ 25% | < 10% |
| | Deuda estimada sobre lo cobrado | ≤ 10% | > 25% |
| Futuro | Resultado proyectado a 3 meses | positivo aun en el escenario pesimista | negativo |
| | Ejecución presupuestaria | ≥ -5% | < -15% |

Lo que está entre verde y rojo es amarillo. Si falta información (por ejemplo, sin período anterior) el indicador **no aparece**
en el semáforo ni se menciona. Se muestra con círculos de colores (🟢 🟡 🔴). La situación general se resuelve por el puntaje promedio y es **crítica** si el margen o la
cobertura están en rojo.

## Proyección
- Se calcula solo con movimientos **ordinarios** (sin extraordinarios) y desde el cierre del período del informe.
- Métodos: **tendencia lineal** (por defecto) o **promedio de los últimos 3 meses**. Banda del 80% (escenarios pesimista y optimista).
  La variación mensual se limita a +40% / -30% del promedio reciente para evitar extrapolaciones extremas.
- **Confiabilidad** según cuánto se proyecta respecto del historial: aceptable (hasta 40% del historial), baja (hasta 100%), muy baja.
  Con 9 meses de datos, 12 meses de proyección es *muy baja*. Valores nominales, sin estacionalidad (haría falta 12+ meses).
- **Horizonte a elección**: campo en la barra lateral (informe y PDF), o la pestaña **Planificación** para explorarlo al instante.
  Por línea de comandos: `--proyeccion 9 --metodo promedio3`.

## Presupuesto
Se **inventa a partir de los datos**: tendencia de cada rubro (tope +15% / -20% sobre su promedio reciente) con criterio prudente:
ingresos -5%, egresos +5% y reserva de contingencia del 5% de los egresos. Se elige de 3, 6 o 12 meses (`--presupuesto-meses`) y se
muestra por rubro y por trimestre. La **ejecución** arma el mismo presupuesto con los datos anteriores al período y lo compara con lo real
(solo si hay al menos 3 meses previos). Es una propuesta, no un presupuesto aprobado.

## Orden del informe
Introducción · Desarrollo (2.1 a 2.10) · Cierre. Hacia el final del desarrollo van el **2.9 Semáforo de indicadores** y el
**2.10 Proyección**, justo antes del **Cierre**.

## Envío del PDF por correo
En la pestaña **Informe**, desplegable **Enviar el PDF por correo electrónico**: se escribe una o más direcciones (hasta 5, separadas
por coma), se puede editar el asunto y se envía el PDF adjunto con un breve texto. Usa SMTP (librería estándar de Python): hace falta
una cuenta de correo desde la que salgan los mensajes y sus datos cargados en los secretos (ver `.streamlit/secrets.toml.example`).
Sin esos datos, el desplegable solo muestra cómo configurarlo.

- **Gmail**: activar la verificación en dos pasos y crear una *contraseña de aplicación*; usar `smtp.gmail.com`, puerto 587.
- **Outlook / Hotmail**: `smtp-mail.outlook.com`, puerto 587 (puede requerir contraseña de aplicación).
- **Protecciones**: `APP_PASSWORD` para entrar a la app, tope de envíos por visitante (`EMAIL_MAX_POR_SESION`, 5 por defecto),
  máximo 5 destinatarios por envío y, opcionalmente, `EMAIL_DOMINIOS_PERMITIDOS` para enviar solo a ciertos dominios.
- Cada envío (fecha, destinatarios, informe, resultado) queda en la pestaña **Registro**.
- Los mensajes salen desde la cuenta configurada: conviene una cuenta del club o una dedicada, no una personal.

## Probar en tu computadora
```
pip install -r requirements.txt
streamlit run app.py
```
Por línea de comandos (también acepta una carpeta con CSV):
```
python agente_riestra.py datos_ejemplo/Riestra_Tesis_Finanzas_V2.xlsx --tipo trimestral --periodo 2026-T3
python agente_riestra.py datos_ejemplo/csv --tipo mensual --periodo 2026-08
python agente_riestra.py datos_ejemplo/Riestra_Tesis_Finanzas_V2.xlsx --tipo trimestral --periodo 2026-T3 --proyeccion 9 --presupuesto-meses 6
```
Para probar el envío por correo en local, copiá `.streamlit/secrets.toml.example` a `.streamlit/secrets.toml` y completá los datos SMTP.

## Publicar con un link (Streamlit Community Cloud)
1. Subí esta carpeta a un repositorio de GitHub (el `.gitignore` ya excluye los secretos y los registros).
2. En https://share.streamlit.io elegí **New app**, seleccioná el repo y como archivo principal `app.py`.
3. En **Advanced settings > Secrets** pegá el contenido de `.streamlit/secrets.toml.example` con tus valores.
4. **Deploy**. Te da un link público para compartir.

Con un link público cualquiera podría usar el envío de correo desde tu cuenta: usá `APP_PASSWORD`, el tope `EMAIL_MAX_POR_SESION`
y, si podés, `EMAIL_DOMINIOS_PERMITIDOS`.

## Datos
- **Google Sheets**: compartir como *Cualquier persona con el enlace: Lector*. Si es un .xlsx subido a Drive, convertirlo con
  *Archivo > Guardar como hoja de cálculo de Google*.
- **OneDrive**: enlace público. Si falla, subí el archivo desde la app.

## Privacidad
El análisis se redacta localmente: los datos no se envían a ningún servicio de IA. Si se envía el PDF por correo, el informe
(que incluye nombres de socios en el detalle de deudores) pasa por el servidor de correo configurado y llega a los destinatarios
indicados. Con datos reales conviene anonimizarlos o informar a los socios.

## Seguimiento
- Pestaña **Registro**: cada informe generado (fecha, fuente, período, cifras, resumen) y cada envío por correo. En Streamlit Community
  Cloud los archivos se reinician cuando la app se reinicia: descargalos.

## Estructura del código
`agente_riestra.py` (carga, cálculo, textos, PDF) · `planificacion.py` (matemática de proyección y presupuesto, sin dependencias de la app) ·
`correo.py` (envío del PDF por correo) · `app.py` (interfaz Streamlit).

## Escudo y template
Escudo: `assets/escudo.png` (o subilo desde la barra lateral). Colores y textos del PDF: función `construir_pdf`.
