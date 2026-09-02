import os
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

EXCEL_MATRIZ = "matriz_gad.xlsx"


@st.cache_data
def cargar_matriz_limpia():
    if not os.path.exists(EXCEL_MATRIZ):
        st.error(f"No se encontró el archivo '{EXCEL_MATRIZ}'.")
        return pd.DataFrame()
    df = pd.read_excel(EXCEL_MATRIZ, header=0)
    df.columns = df.columns.astype(str).str.strip()
    for a, l in [
        ("é", "e"),
        ("ó", "o"),
        ("í", "i"),
        ("á", "a"),
        ("ú", "u"),
        ("É", "E"),
        ("Ó", "O"),
        ("Í", "I"),
        ("Á", "A"),
        ("Ú", "U"),
    ]:
        df.columns = df.columns.str.replace(a, l)
    return df


df_matriz = cargar_matriz_limpia()

if not df_matriz.empty:
    st.set_page_config(
        page_title="Ficha Diagnóstico GADPI - SIL", layout="centered"
    )
    st.title("DIRECCIÓN GENERAL DE PLANIFICACIÓN Y COOPERACIÓN")
    st.title("🏛️ Diagnóstico de LA Gestión de Información - GADPI")
    st.write(
        "Ficha técnica oficial para el levantamiento de información, bases de datos y productos del SIL Geo-Imbabura."
    )
    st.info(
        "✉️ **¿Preguntas o información adicional?** lvega@imbabura.gob.ec"
    )
    st.markdown("---")

    columnas = list(df_matriz.columns)
    col_dir = next((c for c in columnas if "dir" in c.lower()), "Direccion")
    col_sub = next(
        (
            c
            for c in columnas
            if "sub" in c.lower() or "jef" in c.lower() or "dep" in c.lower()
        ),
        "Subunidad",
    )
    col_prod = next(
        (c for c in columnas if "prod" in c.lower() or "est" in c.lower()),
        "Producto",
    )

    try:
        st.header("Sección 1: Identificación del Informante")
        dir_opcion = st.selectbox(
            "1.1 Direcciones o Áreas:",
            sorted(df_matriz[col_dir].dropna().unique()),
        )
        df_f_sub = df_matriz[df_matriz[col_dir] == dir_opcion]
        sub_opcion = st.selectbox(
            "1.2 Subdirección / Jefatura / Unidad:",
            sorted(df_f_sub[col_sub].dropna().unique()),
        )
        tecnico_resp = st.text_input(
            "1.3 Nombre del Técnico Responsable del Llenado:",
            placeholder="Nombres y Apellidos completos",
        )
        correo_ext = st.text_input(
            "1.4 Correo Institucional, Extensión Telefónica, cell:",
            placeholder="ejemplo@imbabura.gob.ec - Ext. 0000",
        )

        st.markdown("---")
        st.header("Sección 2: Producto según Estatuto 2026")
        df_f_prod = df_f_sub[df_f_sub[col_sub] == sub_opcion]
        prod_opcion = st.selectbox(
            "2.1 Seleccione el Producto Institucional del Estatuto Orgánico:",
            sorted(df_f_prod[col_prod].dropna().unique()),
        )
        aplica_info = st.radio(
            "2.2 ¿Aplica generación o manejo de información en este producto?",
            ["Sí", "No"],
        )

        conn = st.connection("gsheets", type=GSheetsConnection)

        def guardar_datos_nube(registro_dicc):
            try:
                df_existente = conn.read(ttl=0)
                df_nuevo_registro = pd.DataFrame([registro_dicc])
                df_consolidado = pd.concat(
                    [df_existente, df_nuevo_registro], ignore_index=True
                )
                conn.update(data=df_consolidado)
                st.success("¡Ficha de diagnóstico guardada con éxito!")
                st.toast("¡Registro guardado con éxito! 👍", icon="👍")
            except Exception as e:
                st.error(f"Error de conexión con Google Sheets: {e}")

        if aplica_info == "No":
            st.warning(
                "Ha seleccionado que NO aplica información para este producto. Guarde el registro para finalizar."
            )
            if st.button("💾 Guardar Producto (No Aplica)", type="secondary"):
                reg = {
                    "Direccion": dir_opcion,
                    "Subunidad": sub_opcion,
                    "Tecnico": tecnico_resp,
                    "Contacto": correo_ext,
                    "Producto": prod_opcion,
                    "Aplica Info": "No",
                }
                guardar_datos_nube(reg)
        else:
            st.markdown("---")
            st.header("Sección 3: Datos Estadísticos")
            gen_est = st.radio(
                "3.1 ¿Genera o posee Datos Estadísticos alfanuméricos?",
                ["Sí", "No"],
            )
            if gen_est == "Sí":
                desag_est = st.multiselect(
                    "3.2 Nivel de Desagregación Estadística:",
                    [
                        "Provincial",
                        "Cantonal",
                        "Parroquial",
                        "Sector / Comunidad",
                        "Predio / Proyecto",
                    ],
                )
                cobertura_est = st.text_input(
                    "3.3 Cobertura Temporal de Datos Estadísticos:",
                    placeholder="Ejemplo: 2018 - 2026",
                )
            else:
                desag_est, cobertura_est = ["No aplica"], "No aplica"

            st.markdown("---")
            st.header("Sección 4: Datos Geográficos (GIS)")
            gen_gis = st.radio(
                "4.1 ¿Genera o posee Datos Geográficos / Espaciales (GIS)?",
                ["Sí", "No"],
            )
            if gen_gis == "Sí":
                desag_gis = st.multiselect(
                    "4.2 Nivel de Desagregación Geográfica:",
                    [
                        "Provincial",
                        "Cantonal",
                        "Parroquial",
                        "Sector / Comunidad",
                        "Predio / Proyecto",
                    ],
                )
                anio_gis = st.text_input(
                    "4.3 Año de Datos Geográficos:",
                    placeholder="Ejemplo: 2020 - 2026",
                )
                escala_gis = st.selectbox(
                    "4.4 Escala de Captura / Levantamiento GIS:",
                    ["1:5.000", "1:25.000", "1:50.000", "1:100.000", "No"],
                )
                formato_gis = st.multiselect(
                    "4.5 Formato de Datos Geográficos Disponibles:",
                    [
                        "File Geodatabase (.gdb)",
                        "Shapefile (.shp)",
                        "GeoJSON / KML",
                        "Tabla XY (Excel / CSV)",
                        "Servicio Web (WMS/WFS)",
                    ],
                )
            else:
                desag_gis, anio_gis, escala_gis, formato_gis = (
                    ["No aplica"],
                    "No aplica",
                    "No aplica",
                    ["No aplica"],
                )

            st.markdown("---")
            st.header("Sección 5: Fuentes y Origen del Dato")
            unidad_medida = st.selectbox(
                "5.1 Unidad de Medida del Dato / Indicador:",
                [
                    "Kilómetros",
                    "Hectáreas",
                    "Porcentaje",
                    "Número de usuarios",
                    "Unidades",
                    "No aplica",
                ],
            )
            fuente_origen = st.selectbox(
                "5.2 Fuente de Origen del Dato:",
                ["Interno GADPI", "Entidad Externa", "Mixto"],
            )
            nombre_fuente = st.text_input(
                "5.3 Nombre Específico de la Fuente / Proveedor:"
            )
            unidad_prov = st.text_input(
                "5.4 Unidad / Dirección Interna Proveedora (si aplica):"
            )
            inst_ext_prov = st.text_input(
                "5.5 Institución Externa Proveedora (si aplica):"
            )

            st.markdown("---")
            st.header("Sección 6: Medios de Verificación y Flujos")
            medio_verif = st.multiselect(
                "6.1 Medio de Verificación Disponible:",
                [
                    "Físico (Archivo)",
                    "Digital (Servidor/PC)",
                    "Base de Datos",
                    "Sistema Web",
                ],
            )
            ruta_archivo = st.text_input(
                "6.2 Nombre de archivo, BD o Enlace del medio de verificación:"
            )
            difunde_terceros = st.radio(
                "6.3 ¿Entrega o difunde este producto a terceros?",
                ["Sí", "No"],
            )
            destinatarios = st.multiselect(
                "6.4 Destinatarios de la Información (si aplica):",
                [
                    "Otras Direcciones GADPI",
                    "GADs Cantonales / Parroquiales",
                    "Ministerios",
                    "Público en general",
                ],
            )

            st.markdown("---")
            st.header("Sección 7: Gobernanza y Calidad")
            frec_act = st.selectbox(
                "7.1 Frecuencia de Actualización:",
                [
                    "Continuo",
                    "Mensual",
                    "Trimestral",
                    "Semestral",
                    "Anual",
                    "Por demanda",
                    "No se actualizan",
                ],
            )
            fecha_ultima = st.text_input(
                "7.2 Fecha de Última Actualización de la información (AAAA/MM):"
            )
            limitaciones = st.multiselect(
                "7.3 Principales Limitaciones para la Actualización:",
                [
                    "Falta personal técnico",
                    "Restricciones presupuestarias",
                    "Software obsoleto",
                    "Equipamiento insuficiente",
                    "Falta normativa",
                ],
            )
            planificacion = st.multiselect(
                "7.4 Alineación Marco de Planificación:",
                [
                    "PDOT Imbabura",
                    "POA Institucional",
                    "ODS",
                    "Competencias Ley / COOTAD",
                ],
            )
            ficha_met = st.radio(
                "7.5 ¿Cuenta con Ficha Metodológica Formalizada?",
                ["Sí", "No", "En proceso"],
            )
            uni_resp_calcul = st.text_input(
                "7.6 Unidad Responsable de la Ficha / Cálculo:"
            )
            riesgos_preserv = st.multiselect(
                "7.7 Identificación de Riesgos de Preservación:",
                [
                    "Dependencia una persona",
                    "Ausencia respaldos",
                    "Virus/Fallos",
                    "Rotación personal",
                    "Deterioro papel",
                ],
            )

            st.markdown("---")
            st.header("Sección 8: Uso de la Información")
            uso_interno = st.text_area(
                "8.1 Uso Interno Actual de la Información:"
            )
            uso_sil = st.text_area(
                "8.2 Potencial Uso:"
            )
            nivel_acceso = st.radio(
                "8.3 Nivel de Acceso de la Información:",
                ["Público", "Restringido", "Uso Interno únicamente"],
            )
            url_publicacion = st.text_input(
                "8.4 Plataforma / Enlace Web de Publicación (si aplica):"
            )

            st.markdown("---")
            if st.button("💾 Guardar Ficha de Diagnóstico", type="primary"):
                if not tecnico_resp or not correo_ext:
                    st.warning(
                        "Complete los datos del Técnico Responsable en la Sección 1."
                    )
                else:
                    reg = {
                        "Direccion": dir_opcion,
                        "Subunidad": sub_opcion,
                        "Tecnico": tecnico_resp,
                        "Contacto": correo_ext,
                        "Producto": prod_opcion,
                        "Aplica Info": aplica_info,
                        "Datos Estadisticos": gen_est,
                        "Desagregacion Est": ", ".join(desag_est),
                        "Cobertura Temporal": cobertura_est,
                        "Datos GIS": gen_gis,
                        "Desagregacion GIS": ", ".join(desag_gis),
                        "Anio GIS": anio_gis,
                        "Escala GIS": escala_gis,
                        "Formato GIS": ", ".join(formato_gis),
                        "Unidad Medida": unidad_medida,
                        "Fuente Origen": fuente_origen,
                        "Nombre Fuente": nombre_fuente,
                        "Unidad Prov": unidad_prov,
                        "Inst Ext Prov": inst_ext_prov,
                        "Medio Verificacion": ", ".join(medio_verif),
                        "Ruta/Enlace": ruta_archivo,
                        "Difunde Terceros": difunde_terceros,
                        "Destinatarios": ", ".join(destinatarios),
                        "Frecuencia Act": frec_act,
                        "Fecha Ultima Act": fecha_ultima,
                        "Limitaciones": ", ".join(limitaciones),
                        "Alineacion Planif": ", ".join(planificacion),
                        "Ficha Metodologica": ficha_met,
                        "Unidad Resp Cálculo": uni_resp_calcul,
                        "Riesgos Preservacion": ", ".join(riesgos_preserv),
                        "Uso Interno": uso_interno,
                        "Integracion SIL": uso_sil,
                        "Nivel Acceso": nivel_acceso,
                        "URL Publicacion": url_publicacion,
                        "Fecha de Registro": pd.Timestamp.now().strftime(
                            "%Y/%m/%d"
                        ),
                    }
                    guardar_datos_nube(reg)
    except KeyError as e:
        st.error(f"Error al acoplar las columnas: {columnas}")
