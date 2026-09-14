import os
import time
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# 1. Configuración de página (siempre al inicio)
st.set_page_config(
    page_title="Diagnóstico SIL - GADPI",
    page_icon="📊",
    layout="wide"
)

# Archivo local de respaldo
EXCEL_DIAGNOSTICO = "diagnostico_sil.xlsx"

# Permisos para la API de Google Sheets y Drive
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# --- FUNCIÓN DE CONEXIÓN A GOOGLE SHEETS ---
@st.cache_resource
def get_sheet_connection():
    credentials_info = st.secrets["connections"]["gsheets"]
    creds = Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
    gc = gspread.authorize(creds)
    
    url_or_id = st.secrets["connections"]["gsheets"]["spreadsheet"].strip()
    if "docs.google.com" in url_or_id:
        sh = gc.open_by_url(url_or_id)
    else:
        sh = gc.open_by_key(url_or_id)
        
    return sh.sheet1

# --- FUNCIÓN DE GUARDADO EN NUBE Y LOCAL ---
def guardar_datos_nube(registro_dicc):
    try:
        sheet = get_sheet_connection()
        
        # Obtener encabezados actuales de la Fila 1
        headers = sheet.row_values(1)
        
        # Si la Fila 1 está totalmente limpia, genera automáticamente las 41 columnas
        if not headers:
            headers = list(registro_dicc.keys())
            sheet.append_row(headers)
        
        # Construye la fila respetando la posición de los encabezados
        fila_valores = [str(registro_dicc.get(col, "")) for col in headers]
        
        # Guarda la fila de respuestas
        sheet.append_row(fila_valores)

        # Respaldo en archivo Excel local
        df_nuevo = pd.DataFrame([registro_dicc])
        if os.path.exists(EXCEL_DIAGNOSTICO):
            df_existente = pd.read_excel(EXCEL_DIAGNOSTICO)
            df_consolidado = pd.concat([df_existente, df_nuevo], ignore_index=True)
        else:
            df_consolidado = df_nuevo
        df_consolidado.to_excel(EXCEL_DIAGNOSTICO, index=False)
        
        st.balloons()
        st.success("¡Ficha guardada exitosamente en Google Sheets!")
        time.sleep(1.5)
        
        if "contador_guardado" in st.session_state:
            st.session_state.contador_guardado += 1
            
        st.rerun()

    except Exception as e:
        st.error(f"Error técnico al guardar los datos en Google Sheets: {e}")

# --- INTERFAZ PRINCIPAL CON LAS 41 PREGUNTAS ---
st.title("📋 Diagnóstico de Gestión de Información del SIL - GADPI")
st.markdown("---")

if "contador_guardado" not in st.session_state:
    st.session_state.contador_guardado = 0

with st.form("form_diagnostico_completo", clear_on_submit=True):
    
    st.header("1. Datos Generales de la Dependencia")
    col1, col2 = st.columns(2)
    with col1:
        q1 = st.text_input("1. Dirección / Dirección General")
        q2 = st.text_input("2. Subunidad / Departamento / Unidad")
        q3 = st.text_input("3. Nombre del Responsable de la Unidad")
        q4 = st.text_input("4. Cargo del Responsable")
    with col2:
        q5 = st.text_input("5. Nombre del Técnico que llena la ficha")
        q6 = st.text_input("6. Cargo del Técnico")
        q7 = st.text_input("7. Correo Electrónico de contacto")
        q8 = st.text_input("8. Teléfono / Extensión de contacto")
        
    st.header("2. Identificación del Producto / Insumo Institucional")
    q9 = st.text_input("9. Nombre del Producto o Insumo principal")
    q10 = st.text_area("10. Descripción breve del Producto o Insumo")
    q11 = st.selectbox("11. ¿Aplica a la Gestión de Información Institucional?", ["Sí", "No", "En proceso"])
    q12 = st.selectbox("12. Tipo de Información principal", ["Estadística", "Geográfica / Cartográfica", "Documental / Textual", "Alfanumérica / Tabular", "Multimedial", "Otra"])
    q13 = st.selectbox("13. Estado actual de la Información", ["Finalizada", "En proceso de producción", "En revisión / Validación", "Desactualizada"])
    
    st.header("3. Insumos Estadísticos (Alfanuméricos)")
    q14 = st.text_input("14. Nombre del Insumo Estadístico")
    q15 = st.text_input("15. Fuente origen de los datos estadísticos")
    q16 = st.selectbox("16. Nivel de Desagregación", ["Provincial", "Cantonal", "Parroquial", "Sector / Barrio", "Predial / Finca", "Sin desagregación"])
    q17 = st.text_input("17. Cobertura Territorial de los datos")
    q18 = st.selectbox("18. Periodicidad de recopilación", ["Diaria", "Semanal", "Mensual", "Trimestral", "Semestral", "Anual", "Eventual"])
    
    st.header("4. Insumos Cartográficos y Geográficos")
    q19 = st.text_input("19. Nombre del Insumo Cartográfico / Capa")
    q20 = st.selectbox("20. ¿La unidad genera Cartografía propia?", ["Sí", "No"])
    q21 = st.text_input("21. Escala de trabajo o representación (ej. 1:5000)")
    q22 = st.selectbox("22. Tipo de Geometría", ["Punto", "Línea", "Polígono", "Raster / Imagen", "Alfanumérico (Sin geometría)", "N/A"])
    q23 = st.text_input("23. Sistema de Referencia Espacial (ej. WGS84 UTM Zone 17S)")
    q24 = st.text_input("24. Formato de Entrega / Almacenamiento (Shapefile, GeoJSON, GDB, TIFF, Excel)")
    q25 = st.selectbox("25. Frecuencia de Actualización Cartográfica", ["Diaria", "Mensual", "Trimestral", "Anual", "Según demanda"])
    
    st.header("5. Almacenamiento, Infraestructura y Software")
    q26 = st.selectbox("26. ¿Dónde se almacena principalmente la información?", ["Servidor Institucional", "Disco Duro Local / PC", "Nube (Google Drive, OneDrive)", "Base de Datos Centralizada", "Archivo Físico / Impreso"])
    q27 = st.text_input("27. Nombre del Servidor o Base de Datos utilizada (si aplica)")
    q28 = st.text_input("28. Software o herramientas utilizadas (ArcGIS Pro, QGIS, Excel, PostgreSQL, etc.)")
    q29 = st.selectbox("29. ¿Cuenta con Metadato / Ficha Técnica estandarizada?", ["Sí", "No", "En proceso"])
    q30 = st.selectbox("30. ¿Existe respaldo (Backup) de la información?", ["Sí, automatizado", "Sí, manual", "No cuenta con respaldo"])

    st.header("6. Accesibilidad, Seguridad y Normativa")
    q31 = st.selectbox("31. Nivel de Acceso a la información", ["Público", "Uso Interno de la Dirección", "Uso Interinstitucional", "Restringido / Confidencial"])
    q32 = st.text_area("32. Normativa Legal o Respaldo Institucional que regula la información")
    q33 = st.text_area("33. Restricciones de Uso, Licencia o Difusión")
    q34 = st.selectbox("34. ¿Dispone de un catálogo de servicios web (WMS/WFS)?", ["Sí", "No", "En planes"])

    st.header("7. Requerimientos, Capacitación y Fortalecimiento")
    q35 = st.selectbox("35. ¿El equipo requiere Capacitación técnica?", ["Sí", "No"])
    q36 = st.text_area("36. Temas específicos requeridos para Capacitación")
    q37 = st.selectbox("37. ¿Requiere Equipamiento informático o Software adicional?", ["Sí", "No"])
    q38 = st.text_area("38. Detalle del Equipamiento o Licencias requeridas")
    q39 = st.selectbox("39. ¿Requiere asesoría técnica del equipo SIL?", ["Sí", "No"])
    q40 = st.text_area("40. Principales nudos críticos / Dificultades detectadas")
    q41 = st.text_area("41. Observaciones y Sugerencias Adicionales")

    boton_guardar = st.form_submit_button("💾 Guardar Ficha de Diagnóstico")

if boton_guardar:
    # Mapeo exacto de las 41 preguntas en las 41 columnas
    registro_dicc = {
        "1. Dirección": q1,
        "2. Subunidad": q2,
        "3. Responsable Unidad": q3,
        "4. Cargo Responsable": q4,
        "5. Técnico Resp.": q5,
        "6. Cargo Técnico": q6,
        "7. Correo Contacto": q7,
        "8. Teléfono Contacto": q8,
        "9. Nombre Producto": q9,
        "10. Descripción Producto": q10,
        "11. Aplica Gestión Info": q11,
        "12. Tipo Información": q12,
        "13. Estado Información": q13,
        "14. Nombre Insumo Estad": q14,
        "15. Fuente Origen Estad": q15,
        "16. Nivel Desagregación": q16,
        "17. Cobertura Territorial": q17,
        "18. Periodicidad Recopilación": q18,
        "19. Nombre Insumo Cartog": q19,
        "20. Genera Cartografía": q20,
        "21. Escala Trabajo": q21,
        "22. Tipo Geometría": q22,
        "23. Sistema Referencia": q23,
        "24. Formato Entrega": q24,
        "25. Frecuencia Act Cartog": q25,
        "26. Dónde Almacena": q26,
        "27. Servidor / Base Datos": q27,
        "28. Software Utilizado": q28,
        "29. Cuenta Metadato": q29,
        "30. Cuenta Backup": q30,
        "31. Nivel Acceso": q31,
        "32. Normativa Legal": q32,
        "33. Restricciones Uso": q33,
        "34. Servicios Web Cartog": q34,
        "35. Requiere Capacitación": q35,
        "36. Temas Capacitación": q36,
        "37. Requiere Equipamiento": q37,
        "38. Detalle Equipamiento": q38,
        "39. Requiere Asesoría SIL": q39,
        "40. Nudos Críticos": q40,
        "41. Observaciones": q41
    }
    
    guardar_datos_nube(registro_dicc)
