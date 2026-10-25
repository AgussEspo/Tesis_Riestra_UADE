# Agente de Reporte Financiero - Club Atlético Deportivo Riestra

Lee un Excel / Google Sheets con 4 pestañas (`Socios`, `Pago_Cuota_Social`, `Ingresos`, `Egresos`) y genera un
informe ejecutivo en PDF **mensual, trimestral o anual**, comparado con el período anterior.

## Cómo funciona
1. **Fuente**: Google Sheets / OneDrive (enlace), archivo subido o datos de ejemplo.
2. **Cálculo (código)**: todas las cifras las calcula Python, no la IA.
3. **Análisis (Claude)**: redacta resumen, puntos clave, alertas y recomendaciones usando solo esas cifras.
   Si no hay API key o falla la llamada, usa un redactor automático (la app nunca se cae por esto).
4. **PDF** con el template del club + **registro** de cada corrida (`registro_corridas.csv`).

## Probar en tu computadora
```
pip install -r requirements.txt
streamlit run app.py
```
Por línea de comandos (sin interfaz):
```
python agente_riestra.py datos_ejemplo/Riestra_Tesis_Finanzas_V2.xlsx --tipo trimestral --periodo 2026-T3
```
Para usar la IA en local: `export ANTHROPIC_API_KEY="sk-ant-..."` (o copiá `.streamlit/secrets.toml.example` a `secrets.toml`).

## Publicar con un link (Streamlit Community Cloud)
1. Subí esta carpeta a un repositorio de GitHub (el `.gitignore` ya excluye los secretos).
2. Entrá a https://share.streamlit.io, elegí **New app**, seleccioná el repo y como archivo principal `app.py`.
3. En **Advanced settings > Secrets** pegá el contenido de `.streamlit/secrets.toml.example` con tus valores.
4. **Deploy**. Te da un link público para compartir.

Recomendado: poné `APP_PASSWORD` y un límite de gasto en la Console de Anthropic, porque con el link público
cualquiera que entre puede generar informes con tu crédito.

## Datos
- **Google Sheets**: compartir como *Cualquier persona con el enlace: Lector*. Si es un .xlsx subido a Drive,
  convertirlo con *Archivo > Guardar como hoja de cálculo de Google*.
- **OneDrive**: enlace público (*Cualquier persona con el enlace puede ver*). Si falla, subí el archivo desde la app.
- Los nombres de pestañas y columnas deben respetar el formato de `datos_ejemplo/`.

## Seguimiento
- **Pestaña "Registro de corridas"** de la app: fecha, fuente, período, cifras, origen del texto (IA o automático),
  modelo y tokens. En Streamlit Community Cloud ese archivo se reinicia cuando la app se reinicia: descargalo.
- **Uso y costo de la API**: Console de Anthropic (podés fijar alertas y límites de gasto).

## Modificar el template
Colores y textos del PDF: función `construir_pdf` en `agente_riestra.py`. Escudo: `assets/escudo.png`.
