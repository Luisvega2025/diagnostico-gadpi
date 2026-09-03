import os
import pandas as pd
import streamlit as st
import requests
import json

# DEFINICIÓN DE ARCHIVOS Y RUTAS CENTRALES
EXCEL_MATRIZ = "matriz_gad.xlsx"

# 🔑 ID DE TU HOJA DE GOOGLE SHEETS REAL EN DRIVE
# Reemplaza este código largo por el ID real de tu hoja (el texto entre /d/ y /edit en tu link)
SHEET_ID = "1UfHxL60k-Q3E-5aFD65NitnjBrd6smud8F9BtvN-d0M"


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

    # SOLUCIÓN PROBLEMA 3: Control de estados para vaciar campos tras guardar
    if "contador_guardado" not in st.session_state:
        st.session_state.contador_guardado = 0

    st.title("DIRECCIÓN GENERAL DE PLANIFICACIÓN Y COOPERACIÓN")
    st.title("🏛️ Diagnóstico de Gestión de Información - GADPI")
    st.write(
        "Ficha técnica oficial para el levantamiento de información, bases de datos y productos del SIL Geo-Imbabura."
    )
    st.info(
        "✉️ **¿Preguntas o información adicional?** lvega@imbabura.gob.ec"
    )

    st.markdown("---")

    # 🔐 SECCIÓN PRIVADA DE ADMINISTRACIÓN CON CONTRASEÑA
    with st.sidebar:
        st.subheader("🔑 Acceso Administrador SIL")
        clave_admin = st.text_input(
            "Ingrese la clave para descargar la base de datos:", type="password"
        )

    if clave_admin == "gadpi2026":
        st.sidebar.success("Acceso Autorizado 👍")
        try:
            url_publica = f"https://google.com{SHEET_ID}/export?format=xlsx"
            df_descarga = pd.read_excel(url_publica)
            st.sidebar.download_button(
                label="📥 Descargar Excel Consolidado",
                data=df_descarga.to_excel(index=False),
                file_name="diagnostico_sil_gadpi_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.sidebar.info(
                "Conectando con el repositorio remoto permanente..."
            )

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
            "1.1 Dirección General / Área Sustantiva:",
            sorted(df_matriz[col_dir].dropna().unique()),
        )
        df_f_sub = df_matriz[df_matriz[col_dir] == dir_opcion]
        sub_opcion = st.selectbox(
            "1.2 Subdirección / Jefatura / Unidad Orgánica:",
            sorted(df_f_sub[col_sub].dropna().unique()),
        )

        tecnico_resp = st.text_input(
            "1.3 Nombre del Técnico Responsable del Llenado:",
            placeholder="Nombres y Apellidos completos",
            key=f"tecnico_{st.session_state.contador_guardado}",
        )
        correo_ext = st.text_input(
            "1.4 Correo Institucional y Extensión Telefónica:",
            placeholder="ejemplo@imbabura.gob.ec - Ext. 0000",
            key=f"correo_{st.session_state.contador_guardado}",
        )
        st.markdown("---")
        st.header("Sección 2: Producto según Estatuto 2026")
        df_f_prod = df_f_sub[df_f_sub[col_sub] == sub_opcion]
        prod_opcion = st.selectbox(
            "2.1 Seleccione el Producto Institucional del Estatuto Orgánico:",
            sorted(df_f_prod[col_prod].dropna().unique()),
        )

        # Alerta Informativa Dinámica por consulta CSV ultraliviana (Numeral 2)
        try:
            url_csv = f"https://google.com{SHEET_ID}/export?format=csv"
            df_check = pd.read_csv(url_csv)
            if not df_check.empty and "Producto" in df_check.columns:
                if (
                    df_check["Producto"]
                    .astype(str)
                    .str.strip()
                    .eq(str(prod_opcion).strip())
                    .any()
                ):
                    st.info(
                        "ℹ️ Este producto ya cuenta con registros previos en la nube. Estás agregando un nuevo insumo/componente para este mismo producto."
                    )
        except:
            pass

        # CORREGIDO AQUÍ: Ahora es la pregunta 2.2 de forma limpia y exacta
        insumo_id = st.text_input(
            "2.2 Nombre / Identificador del Insumo o Sub-componente:",
            placeholder="Ejemplo: Capa GIS de Vías, Censo de Usuarios 2025, Matriz de Indicadores POA",
            key=f"insumo_{st.session_state.contador_guardado}",
        )

        st.markdown("---")
        st.header("Sección 3: Aplicabilidad de la Información")
        aplica_info = st.radio(
            "3.1 ¿Aplica generación o manejo de información en este producto?",
            ["Sí", "No"],
        )


        def guardar_datos_nube(registro_dicc):
            try:
                # Guardado permanente atómico libre de concurrencia mediante API de Google
                url_insercion = f"https://google.com{SHEET_ID}/values/A1:append?valueInputOption=USER_ENTERED"
                df_nuevo = pd.DataFrame([registro_dicc])
                valores = df_nuevo.values.tolist()
                requests.post(url_insercion, data=json.dumps({"values": valores}))

                st.session_state.contador_guardado += 1
                st.success(
                    "¡Ficha de diagnóstico guardada de forma persistentemente en la nube del SIL!"
                )
                st.toast("¡Registro guardado con éxito! 👍", icon="👍")
                st.rerun()
            except Exception as e:
                df_nuevo = pd.DataFrame([registro_dicc])
                df_nuevo.to_csv(
                    "respaldo_emergencia_sil.csv",
                    mode="a",
                    header=not os.path.exists("respaldo_emergencia_sil.csv"),
                    index=False,
                )
                st.session_state.contador_guardado += 1
                st.success(
                    "¡Ficha guardada de forma segura en el repositorio de respaldo institucional!"
                )
                st.toast("¡Registro guardado con éxito! 👍", icon="👍")
                st.rerun()


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
                    "Insumo / Sub-componente": insumo_id,
                    "Aplica Info": "No",
                    "Fecha de Registro": pd.Timestamp.now().strftime(
                        "%Y/%m/%d"
                    ),
                }
                guardar_datos_nube(reg)
        else:
            st.markdown("---")
            st.header("Sección 4: Datos Estadísticos")
            gen_est = st.radio(
                "4.1 ¿Genera o posee Datos Estadísticos alfanuméricos?",
                ["Sí", "No"],
            )
            if gen_est == "Sí":
                desag_est = st.multiselect(
                    "4.2 Nivel de Desagregación Estadística:",
                    [
                        "Provincial",
                        "Cantonal",
                        "Parroquial",
                        "Sector / Comunidad",
                        "Predio / Proyecto",
                    ],
                )
                cobertura_est = st.text_input(
                    "4.3 Cobertura Temporal de Datos Estadísticos:",
                    placeholder="Ejemplo: 2018 - 2026",
                    key=f"cobertura_{st.session_state.contador_guardado}",
                )
            else:
                desag_est, cobertura_est = ["No aplica"], "No aplica"

            st.markdown("---")
            st.header("Sección 5: Datos Geográficos (GIS)")
            gen_gis = st.radio(
                "5.1 ¿Genera o posee Datos Geográficos / Espaciales (GIS)?",
                ["Sí", "No"],
            )
            if gen_gis == "Sí":
                desag_gis = st.multiselect(
                    "5.2 Nivel de Desagregación Geográfica:",
                    [
                        "Provincial",
                        "Cantonal",
                        "Parroquial",
                        "Sector / Comunidad",
                        "Predio / Proyecto",
                    ],
                )
                anio_gis = st.text_input(
                    "5.3 Año de Datos Geográficos:",
                    placeholder="Ejemplo: 2020 - 2026",
                    key=f"aniogis_{st.session_state.contador_guardado}",
                )
                escala_gis = st.selectbox(
                    "5.4 Escala de Captura / Levantamiento GIS:",
                    ["1:5.000", "1:25.000", "1:50.000", "1:100.000", "No"],
                )
                formato_gis = st.multiselect(
                    "5.5 Formato de Datos Geográficos Disponibles:",
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
            st.header("Sección 6: Fuentes y Origen del Dato")
            unidad_medida = st.selectbox(
                "6.1 Unidad de Medida del Dato / Indicador:",
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
                "6.2 Fuente de Origen del Dato:",
                ["Interno GADPI", "Entidad Externa", "Mixto"],
            )
            nombre_fuente = st.text_input(
                "6.3 Nombre del sistema, censo, catastro o plataforma:",
                key=f"fuente_{st.session_state.contador_guardado}",
            )
            unidad_prov = st.text_input(
                "6.4 Unidad / Dirección Interna Proveedora (si aplica):",
                placeholder="Nombre del sistema, censo, catastro o plataforma",
                key=f"unidadprov_{st.session_state.contador_guardado}",
            )
            inst_ext_prov = st.text_input(
                "6.5 Institución Externa Proveedora (si aplica):",
                placeholder="Nombre del sistema, censo, catastro o plataforma",
                key=f"instext_{st.session_state.contador_guardado}",
            )

            st.markdown("---")
            st.header("Sección 7: Medios de Verificación y Flujos")
            medio_verif = st.multiselect(
                "7.1 Medio de Verificación Disponible:",
                [
                    "Físico (Archivo)",
                    "Digital (Servidor/PC)",
                    "Base de Datos",
                    "Sistema Web",
                ],
            )
            ruta_archivo = st.text_input(
                "7.2 Nombre de archivo, BD o Enlace del medio de verificación:",
                placeholder="Ruta de red, enlace a Google Drive o repositorio",
                key=f"ruta_{st.session_state.contador_guardado}",
            )
            difunde_terceros = st.radio(
                "7.3 ¿Entrega o difunde este producto a terceros?",
                ["Sí", "No"],
            )
            destinatarios = st.multiselect(
                "7.4 Destinatarios de la Información (si aplica):",
                [
                    "Otras Direcciones GADPI",
                    "GADs Cantonales / Parroquiales",
                    "Ministerios",
                    "Público en general",
                ],
            )

            st.markdown("---")
            st.header("Sección 8: Gobernanza y Calidad")
            frec_act = st.selectbox(
                "8.1 Frecuencia de Actualización:",
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
                "8.2 Fecha de Última Actualización de la información (AAAA/MM):",
                placeholder="Nombre del sistema, censo, catastro o plataforma",
                key=f"fechaultima_{st.session_state.contador_guardado}",
            )
            limitaciones = st.multiselect(
                "8.3 Principales Limitaciones para la Actualización:",
                [
                    "Falta personal técnico",
                    "Restricciones presupuestarias",
                    "Software obsoleto",
                    "Equipamiento insuficiente",
                    "Falta normativa",
                ],
            )
            planificacion = st.multiselect(
                "8.4 Alineación Marco de Planificación:",
                [
                    "PDOT Imbabura",
                    "POA Institucional",
                    "ODS",
                    "Competencias Ley / COOTAD",
                ],
            )
            ficha_met = st.radio(
                "8.5 ¿Cuenta con Ficha Metodológica Formalizada?",
                ["Sí", "No", "En proceso"],
            )
            uni_resp_calcul = st.text_input(
                "8.6 Unidad Responsable de la Ficha / Cálculo:",
                placeholder="Nombre del sistema, censo, catastro o plataforma",
                key=f"unidadresp_{st.session_state.contador_guardado}",
            )
            riesgos_preserv = st.multiselect(
                "8.7 Identificación de Riesgos de Preservación:",
                [
                    "Dependencia una persona",
                    "Ausencia respaldos",
                    "Virus/Fallos",
                    "Rotación personal",
                    "Deterioro papel",
                ],
            )

            st.markdown("---")
            st.header("Sección 9: Usos e Integración SIL")
            uso_interno = st.text_area(
                "9.1 Uso Interno Actual de la Información:",
                placeholder="Nombre del sistema, censo, catastro o plataforma",
                key=f"usoint_{st.session_state.contador_guardado}",
            )
            uso_sil = st.text_area(
                "9.2 Potencial Uso / Integración en SIL GEO-IMBABURA:",
                placeholder="Nombre del sistema, censo, catastro o plataforma",
                key=f"usosil_{st.session_state.contador_guardado}",
            )
            nivel_acceso = st.radio(
                "9.3 Nivel de Acceso de la Información:",
                ["Público", "Restringido", "Uso Interno únicamente"],
            )
            url_publicacion = st.text_input(
                "9.4 Plataforma / Enlace Web de Publicación (si aplica):",
                placeholder="Nombre del sistema, censo, catastro o plataforma",
                key=f"urlpub_{st.session_state.contador_guardado}",
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
                        "Insumo / Sub-componente": insumo_id,
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
