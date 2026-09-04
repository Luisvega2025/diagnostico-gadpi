import os
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

# DEFINICIÓN DE ARCHIVOS INSTITUCIONALES DEL SIL
EXCEL_MATRIZ = "matriz_gad.xlsx"

@st.cache_data
def cargar_matriz_limpia():
    if not os.path.exists(EXCEL_MATRIZ):
        st.error(f"No se encontró el archivo '{EXCEL_MATRIZ}'.")
        return pd.DataFrame()
    df = pd.read_excel(EXCEL_MATRIZ, header=0)
    df.columns = df.columns.astype(str).str.strip()
    for a, l in [
        ("é", "e"), ("ó", "o"), ("í", "i"), ("á", "a"), ("ú", "u"),
        ("É", "E"), ("Ó", "O"), ("Í", "I"), ("Á", "A"), ("Ú", "U"),
    ]:
        df.columns = df.columns.str.replace(a, l)
    return df

df_matriz = cargar_matriz_limpia()

if not df_matriz.empty:
    st.set_page_config(
        page_title="Ficha Diagnóstico GADPI - SIL", layout="centered"
    )

    # Conexión directa y segura a Google Sheets usando tus Secrets
    conn = st.connection("gsheets", type=GSheetsConnection)

    # Control de estados para vaciar campos de texto tras guardar exitosamente
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
            df_descarga = conn.read(ttl="5s")
            import io
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_descarga.to_excel(writer, index=False)
            buffer.seek(0)

            st.sidebar.download_button(
                label="📥 Descargar Excel Consolidado (Desde Google Sheets)",
                data=buffer,
                file_name="diagnostico_sil_gadpi_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.sidebar.info("Aún no se registran fichas técnicas o configure sus Secrets.")

    columnas = list(df_matriz.columns)
    col_dir = next((c for c in columnas if "dir" in c.lower()), "Direccion")
    col_sub = next((c for c in columnas if "sub" in c.lower() or "jef" in c.lower() or "dep" in c.lower()), "Subunidad")
    col_prod = next((c for c in columnas if "prod" in c.lower() or "est" in c.lower()), "Producto")

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
            "1.4 Correo Institucional, Extensión Telefónica, cell:",
            placeholder="ejemplo@imbabura.gob.ec - Ext. 0000 , cell",
            key=f"correo_{st.session_state.contador_guardado}",
        )
        st.markdown("---")
        
        st.header("Sección 2: Producto e Insumo según Estatuto 2026")
        df_f_prod = df_f_sub[df_f_sub[col_sub] == sub_opcion]
        prod_opcion = st.selectbox(
            "2.1 Seleccione el Producto Institucional del Estatuto Orgánico:",
            sorted(df_f_prod[col_prod].dropna().unique()),
        )

        # Alerta Informativa leyendo directamente de Google Sheets en tiempo real
        try:
            df_check = conn.read(ttl="5s")
            if (
                not df_check.empty
                and "Producto" in df_check.columns
                and (df_check["Producto"].astype(str).str.strip().eq(str(prod_opcion).strip()).any())
            ):
                st.info(
                    "ℹ️ Este producto ya cuenta con registros previos en la nube. Estás agregando un nuevo insumo/componente para este mismo producto."
                )
        except:
            pass

        # NUEVO CAMPO REQUERIDO: Identificador del Insumo o Sub-componente
        insumo_identificador = st.text_input(
            "2.2 Nombre / Identificador del Insumo o Sub-componente del Producto:",
            placeholder="Ejemplo: Base de datos de predios, Capa Shapefile de vías, Registro alfanumérico",
            key=f"insumo_{st.session_state.contador_guardado}",
        )

        aplica_info = st.radio(
            "2.3 ¿Aplica o genera información para cumplir con este producto?",
            ["Sí", "No"],
        )

        def guardar_datos_nube(registro_dicc):
            try:
                df_nuevo = pd.DataFrame([registro_dicc])
                try:
                    df_existente = conn.read(ttl="0s")
                    df_consolidado = pd.concat([df_existente, df_nuevo], ignore_index=True)
                except:
                    df_consolidado = df_nuevo

                # Envío directo e inmediato a la hoja de Google Sheets para evitar pérdidas de datos
                conn.update(data=df_consolidado)
                
                # Incrementa el contador para resetear y vaciar de forma automática todos los campos
                st.session_state.contador_guardado += 1
                st.success("¡Ficha de diagnóstico guardada de forma persistente en Google Sheets!")
                st.toast("¡Registro guardado con éxito! 👍", icon="👍")
                st.rerun()
            except Exception as e:
                st.error(f"Error de consistencia al escribir en Google Sheets: {e}")

        if aplica_info == "No":
            st.warning("Ha seleccionado que NO aplica información para este producto. Guarde el registro para finalizar.")
            if st.button("💾 Guardar Producto (No Aplica)", type="secondary"):
                reg = {
                    "Direccion": dir_opcion,
                    "Subunidad": sub_opcion,
                    "Tecnico": tecnico_resp,
                    "Contacto": correo_ext,
                    "Producto": prod_opcion,
                    "Insumo Identificador": insumo_identificador,
                    "Aplica Info": "No",
                    "Fecha de Registro": pd.Timestamp.now().strftime("%Y/%m/%d"),
                }
                guardar_datos_nube(reg)
        else:
            st.markdown("---")
            st.header("Sección 3: Datos Alfanuméricos/Estadísticos")
            gen_est = st.radio("3.1 ¿Genera o posee Datos Estadísticos/alfanuméricos?", ["Sí", "No"])
            if gen_est == "Sí":
                desag_est = st.multiselect(
                    "3.2 Nivel de Desagregación estadistica:",
                    ["Provincial", "Cantonal", "Parroquial", "Sector / Comunidad", "Predio / Proyecto"],
                )
                cobertura_est = st.text_input(
                    "3.3 Temporalidad de Datos Estadísticos:",
                    placeholder="Ejemplo: 2018 - 2026",
                    key=f"cobertura_{st.session_state.contador_guardado}",
                )
            else:
                desag_est, cobertura_est = ["No aplica"], "No aplica"

            st.markdown("---")
            st.header("Sección 4: Datos Geográficos (GIS)")
            gen_gis = st.radio("4.1 ¿Genera o posee Datos Geográficos / Espaciales (GIS)?", ["Sí", "No"])
            if gen_gis == "Sí":
                desag_gis = st.multiselect(
                    "4.2 Nivel de Desagregación Geográfica:",
                    ["Provincial", "Cantonal", "Parroquial", "Sector / Comunidad", "Predio / Proyecto"],
                )
                anio_gis = st.text_input(
                    "4.3 Año de Datos Geográficos:",
                    placeholder="Ejemplo: 2020 - 2026",
                    key=f"aniogis_{st.session_state.contador_guardado}",
                )
                escala_gis = st.selectbox(
                    "4.4 Escala de la cartografia:",
                    ["1:5.000", "1:25.000", "1:50.000", "1:100.000", "No"],
                )
                formato_gis = st.multiselect(
                    "4.5 Formato de Datos Geográficos Disponibles:",
                    ["File Geodatabase (.gdb)", "Shapefile (.shp)", "GeoJSON / KML", "Tabla XY (Excel / CSV)", "Servicio Web (WMS/WFS)"],
                )
            else:
                desag_gis, anio_gis, escala_gis, formato_gis = ["No aplica"], "No aplica", "No aplica", ["No aplica"]
            
            st.markdown("---")
            st.header("Sección 5: Fuentes y Origen del Dato")
            unidad_medida = st.selectbox(
                "5.1 Unidad de Medida del Dato / Indicador:",
                ["Kilómetros", "Hectáreas", "Porcentaje", "Número de usuarios", "Unidades", "No aplica"],
            )
            fuente_origen = st.selectbox(
                "5.2 Fuente de Origen del Dato:",
                ["Interno GADPI", "Entidad Externa", "Mixto"],
            )
            nombre_fuente = st.text_input(
                "5.3 Nombre de la fuente/proveedor:",
                placeholder="Nombre del sistema, censo, catastro o plataforma",
                key=f"fuente_{st.session_state.contador_guardado}",
            )
            unidad_prov = st.text_input(
                "5.4 Unidad / Dirección Interna Proveedora (si aplica):",
                placeholder="Nombre de la unidad interna proveedora",
                key=f"unidadprov_{st.session_state.contador_guardado}",
            )
            inst_ext_prov = st.text_input(
                "5.5 Institución Externa Proveedora (si aplica):",
                placeholder="Ejemplo: INEC, MAATE, MTOP, MAG, INAMHI",
                key=f"instext_{st.session_state.contador_guardado}",
            )

            st.markdown("---")
            st.header("Sección 6: Medios de Verificación y Flujos")
            medio_verif = st.multiselect(
                "6.1 Medio de Verificación Disponible:",
                ["Físico (Archivo)", "Digital (Servidor/PC)", "Base de Datos", "Sistema Web"],
            )
            ruta_archivo = st.text_input(
                "6.2 Nombre de archivo, BD o Enlace del medio de verificación:",
                placeholder="Ruta de red, enlace a Google Drive o repositorio",
                key=f"ruta_{st.session_state.contador_guardado}",
            )
            difunde_terceros = st.radio("6.3 ¿Entrega o difunde este producto a terceros?", ["Sí", "No"])
            destinatarios = st.multiselect(
                "6.4 Destinatarios de la Información (si aplica):",
                ["Otras Direcciones GADPI", "GADs Cantonales / Parroquiales", "Ministerios", "Público en general"],
            )

            st.markdown("---")
            st.header("Sección 7: Gobernanza y Calidad")
            frec_act = st.selectbox(
                "7.1 Frecuencia de Actualización:",
                ["Continuo", "Mensual", "Trimestral", "Semestral", "Anual", "Por demanda", "No se actualizan"],
            )
            fecha_ultima = st.text_input(
                "7.2 Fecha de Última Actualización de la información (AAAA/MM):",
                placeholder="AAAA/MM",
                key=f"fechaultima_{st.session_state.contador_guardado}",
            )
            limitaciones = st.multiselect(
                "7.3 Principales Limitaciones para la Actualización:",
                ["Falta personal técnico", "Restricciones presupuestarias", "Software obsoleto", "Equipamiento insuficiente", "Falta normativa"],
            )
            planificacion = st.multiselect(
                "7.4 Alineación Marco de Planificación:",
                ["PDOT Imbabura", "POA Institucional", "ODS", "Competencias Ley / COOTAD"],
            )
            ficha_met = st.radio("7.5 ¿Cuenta con Ficha Metodológica Formalizada?", ["Sí", "No", "En proceso"])
            uni_resp_calcul = st.text_input(
                "7.6 Unidad Responsable de la Ficha / Cálculo:",
                placeholder="Nombre del departamento o perfil técnico",
                key=f"uniresp_{st.session_state.contador_guardado}",
            )
            riesgos_preserv = st.multiselect(
                "7.7 Identificación de Riesgos de Preservación dela Información:",
                ["Dependencia una persona", "Ausencia respaldos", "Virus/Fallos", "Rotación personal", "Deterioro papel"],
            )

            st.markdown("---")
            st.header("Sección 8: Usos de la Información")
            uso_interno = st.text_area(
                "8.1 Uso Interno Actual de la Información:",
                placeholder="Mencione quien hace uso de la información generada",
                key=f"usoint_{st.session_state.contador_guardado}",
            )
            uso_sil = st.text_area(
                "8.2 Potencial Uso / Integración en SIL GEO-IMBABURA:",
                placeholder="Como puede aprovecharse la información",
                key=f"usosil_{st.session_state.contador_guardado}",
            )
            nivel_acceso = st.radio("8.3 Nivel de Acceso de la Información:", ["Público", "Restringido", "Uso Interno únicamente"])
            url_publicacion = st.text_input(
                "8.4 Plataforma / Enlace Web de Publicación (si aplica):",
                placeholder="URL pública del geoportal o visor web",
                key=f"urlpub_{st.session_state.contador_guardado}",
            )

            st.markdown("---")
            if st.button("💾 Guardar Ficha de Diagnóstico", type="primary"):
                if not tecnico_resp or not correo_ext:
                    st.warning("Complete los datos del Técnico Responsable en la Sección 1.")
                else:
                    reg = {
                        "Direccion": dir_opcion,
                        "Subunidad": sub_opcion,
                        "Tecnico": tecnico_resp,
                        "Contacto": correo_ext,
                        "Producto": prod_opcion,
                        "Insumo Identificador": insumo_identificador,
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
                        "Fecha de Registro": pd.Timestamp.now().strftime("%Y/%m/%d"),
                    }
                    guardar_datos_nube(reg)
    except KeyError as e:
        st.error(f"Error al acoplar las columnas: {columnas}")
