import os
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

# DEFINICIÓN DE ARCHIVOS INSTITUCIONALES DEL SIL
EXCEL_MATRIZ = "matriz_gad.xlsx"
EXCEL_DIAGNOSTICO = "diagnostico_sil_gadpi_2026.xlsx"

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

    # Conexión para lectura de datos en tiempo real
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
        st.sidebar.success("Acceso Autorizado 🎈")
        if os.path.exists(EXCEL_DIAGNOSTICO):
            try:
                df_descarga = pd.read_excel(EXCEL_DIAGNOSTICO)
                import io

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_descarga.to_excel(writer, index=False)
                buffer.seek(0)

                st.sidebar.download_button(
                    label="📥 Descargar Excel Consolidado",
                    data=buffer,
                    file_name="diagnostico_sil_gadpi_2026.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.sidebar.error(f"Error al procesar el archivo: {e}")
        else:
            st.sidebar.info("Aún no se registran fichas técnicas en la nube.")

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
            "1.4 Correo Institucional / Contacto:",
            placeholder="correo; celular",
            key=f"correo_{st.session_state.contador_guardado}",
        )
        st.markdown("---")
        st.header("Sección 2: Producto e Insumo según Estatuto 2026")
        df_f_prod = df_f_sub[df_f_sub[col_sub] == sub_opcion]
        prod_opcion = st.selectbox(
            "2.1 Seleccione el Producto Institucional del Estatuto Orgánico:",
            sorted(df_f_prod[col_prod].dropna().unique()),
        )

        # Alerta Informativa leyendo directamente de Google Sheets para control de duplicados
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

        aplica_info = st.radio(
            "2.2 ¿Aplica o genera información para cumplir con este producto?",
            ["Sí", "No"],
            key=f"aplica_{st.session_state.contador_guardado}"
        )

        tipo_informacion = st.radio(
            "2.3 Tipo de información que genera/utiliza?:",
            ["Alfanumérica / Estadística", "Geográfica"],
            key=f"tipo_info_{st.session_state.contador_guardado}"
        )

        st.info("ℹ️ **SI TIENE LOS DOS TIPOS DE INFORMACIÓN SE DEBE LLENAR UN REGISTRO A LA VEZ POR INSUMO (alfanumérico o Cartográfico)**")

        def guardar_datos_nube(registro_dicc):
            try:
                import requests
                import json
                
                # URL Conector oficial de tu Google Apps Script
                url_google_script = "https://google.com"
                
                # 1. Enviar los datos en tiempo real de forma externa a Google Sheets
                payload = json.dumps(registro_dicc)
                headers = {'Content-Type': 'application/json'}
                requests.post(url_google_script, data=payload, headers=headers, timeout=10)
                
                # 2. Respaldo local doble en el servidor por seguridad de la sesión
                df_nuevo = pd.DataFrame([registro_dicc])
                if os.path.exists(EXCEL_DIAGNOSTICO):
                    df_existente = pd.read_excel(EXCEL_DIAGNOSTICO)
                    df_consolidado = pd.concat([df_existente, df_nuevo], ignore_index=True)
                else:
                    df_consolidado = df_nuevo
                df_consolidado.to_excel(EXCEL_DIAGNOSTICO, index=False)
                
                # 🎈 Animación de globos ascendentes
                st.balloons()
                
                import time
                time.sleep(1.5)
                
                # Incrementa el contador para resetear y vaciar los campos de texto
                st.session_state.contador_guardado += 1
                st.rerun()
            except Exception as e:
                st.error(f"Error al enviar datos al sistema central: {e}")
        if aplica_info == "No":
            st.warning("Ha seleccionado que NO aplica información para este producto. Guarde el registro para finalizar.")
            if st.button("💾 Guardar Producto (No Aplica)", type="secondary"):
                reg = {
                    "Direccion": dir_opcion,
                    "Subunidad": sub_opcion,
                    "Tecnico": tecnico_resp,
                    "Contacto": correo_ext,
                    "Producto": prod_opcion,
                    "Aplica Info": aplica_info,
                    "Tipo Informacion": tipo_informacion,
                    "nombre Ins Estad": "No aplica",
                    "Desagregacion Est": "No aplica",
                    "Cobertura Temporal": "No aplica",
                    "nombre Insu Carto": "No aplica",
                    "genera cart": "No aplica",
                    "Desagregacion GIS": "No aplica",
                    "Anio GIS": "No aplica",
                    "Escala GIS": "No aplica",
                    "Formato GIS": "No aplica",
                    "Otro Formato GIS": "No aplica",
                    "Genera Info Georreferenciada": "No aplica",
                    "Otras Fuentes GIS": "No aplica",
                    "Tiene Metadatos": "No aplica",
                    "Unidad Medida": "No aplica",
                    "Fuente Origen": "No aplica",
                    "Nombre Fuente": "No aplica",
                    "Unidad Prov": "No aplica",
                    "Inst Ext Prov": "No aplica",
                    "Medio Verificacion": "No aplica",
                    "Ruta/Enlace": "No aplica",
                    "Difunde Terceros": "No aplica",
                    "Destinatarios": "No aplica",
                    "Frecuencia Act": "No aplica",
                    "Fecha Ultima Act": "No aplica",
                    "Limitaciones": "No aplica",
                    "Otra Razon Limitacion": "No aplica",
                    "Alineacion Planif": "No aplica",
                    "Ficha Metodologica": "No aplica",
                    "Unidad Resp Calculo": "No aplica",
                    "Riesgos Preservacion": "No aplica",
                    "Uso Interno": "No aplica",
                    "Integracion SIL": "No aplica",
                    "Nivel Acceso": "No aplica",
                    "URL Publicacion": "No aplica",
                    "Fecha de Registro": pd.Timestamp.now().strftime("%Y/%m/%d"),
                }
                guardar_datos_nube(reg)
        else:
            # CONDICIONAL PRINCIPAL SEGÚN EL NUMERAL 2.3
            if tipo_informacion == "Alfanumérica / Estadística":
                st.markdown("---")
                st.header("Sección 3: Datos Alfanuméricos/Estadísticos")
                
                nombre_ins_estad = st.text_input(
                    "3.1 ¿Nombre del insumo estadístico/alfanumérico que aporta a este producto?",
                    placeholder="ejem: usuarios_canal_riego.*/doc/pdf/xls/",
                    key=f"insumo_est_{st.session_state.contador_guardado}"
                )
                
                desag_est = st.multiselect(
                    "3.2 Nivel de Desagregación estadística / Desagregacion Est:",
                    ["Provincial", "Cantonal", "Parroquial", "Sector / Comunidad", "Predio / Proyecto"],
                )
                
                cobertura_est = st.text_input(
                    "3.3 Temporalidad de Datos Estadísticos / Cobertura Temporal:",
                    placeholder="Ejemplo: 2018 - 2026",
                    key=f"cobertura_{st.session_state.contador_guardado}",
                )
                
                # Contingencias GIS estables en "No aplica" para este flujo alfanumérico
                nombre_insu_carto, genera_cart, desag_gis, anio_gis, escala_gis = "No aplica", "No aplica", ["No aplica"], "No aplica", "No aplica"
                formato_gis, otro_formato_gis, genera_info_georref, otras_fuentes_gis, tiene_metadatos = ["No aplica"], "No aplica", "No aplica", "No aplica", "No aplica"

            else:  # Caso: "Geográfica"
                st.markdown("---")
                st.header("Sección 4: Datos Geográficos (GIS)")
                
                nombre_insu_carto = st.text_input(
                    "4.1 ¿Nombre del insumo cartográfico que aporta a este producto?",
                    placeholder="ejem: vias.shp/*nombre.mxd/nombre.gdb",
                    key=f"insumo_carto_{st.session_state.contador_guardado}"
                )
                
                genera_cart = st.radio(
                    "4.2 ¿Genera o posee Datos Geográficos / Espaciales (GIS)? / genera cart:", 
                    ["Sí", "No"],
                    key=f"genera_cart_{st.session_state.contador_guardado}"
                )
                
                desag_gis = st.multiselect(
                    "4.3 Nivel de Desagregación Geográfica / Desagregacion GIS:",
                    ["Provincial", "Cantonal", "Parroquial", "Sector / Comunidad", "Predio / Proyecto"],
                )
                
                anio_gis = st.text_input(
                    "4.4 Año de Datos Geográficos / Anio GIS:",
                    placeholder="Ejemplo: 2020 - 2026",
                    key=f"aniogis_{st.session_state.contador_guardado}",
                )
                
                escala_gis = st.selectbox(
                    "4.5 Escala de la cartografía / Escala GIS:",
                    ["1:5.000", "1:25.000", "1:50.000", "1:100.000", "No"],
                )
                
                formato_gis = st.multiselect(
                    "4.6 Formato de Datos Geográficos Disponibles / Formato GIS:",
                    ["File Geodatabase (.gdb)", "Shapefile (.shp)", "GeoJSON / KML", "Tabla XY (Excel / CSV)", "Servicio Web (WMS/WFS)"],
                )
                
                otro_formato_gis = st.text_input(
                    "4.7 ¿Otro formato? / Otro Formato GIS:",
                    key=f"otro_formato_{st.session_state.contador_guardado}"
                )
                
                genera_info_georref = st.text_input(
                    "4.8 ¿Genera información georreferenciada - Cartografía? / Genera Info Georreferenciada:",
                    placeholder="ejem: si, archivo shp, maps mxd, etc",
                    key=f"genera_georref_{st.session_state.contador_guardado}"
                )
                
                otras_fuentes_gis = st.text_input(
                    "4.9 ¿Obtiene de otras fuentes? Cuáles? / Otras Fuentes GIS:",
                    placeholder="ejem: IGM, INEC, MAG, etc",
                    key=f"otras_fuentes_{st.session_state.contador_guardado}"
                )
                
                tiene_metadatos = st.text_input(
                    "4.10 ¿Tiene metadatos, catálogo de objetos? / Tiene Metadatos:",
                    placeholder="si/no",
                    key=f"metadatos_{st.session_state.contador_guardado}"
                )
                
                # Valores por defecto para el bloque estadístico que se ocultó en este flujo
                nombre_ins_estad, desag_est, cobertura_est = "No aplica", ["No aplica"], "No aplica"

            # El flujo alfanumérico salta linealmente aquí (Sección 5)
            st.markdown("---")
            st.header("Sección 5: Fuentes y Origen del Dato")
            unidad_medida = st.selectbox(
                "5.1 Unidad de Medida del Dato / Indicador / Unidad Medida:",
                ["Kilómetros", "Hectáreas", "Porcentaje", "Número de usuarios", "Unidades", "No aplica"],
            )
            fuente_origen = st.selectbox(
                "5.2 Fuente de Origen del Dato / Fuente Origen:",
                ["Interno GADPI", "Entidad Externa", "Mixto"],
            )
            nombre_fuente = st.text_input(
                "5.3 Nombre de la fuente/proveedor / Nombre Fuente:",
                placeholder="Nombre del sistema, censo, catastro o plataforma",
                key=f"fuente_{st.session_state.contador_guardado}",
            )
            unidad_prov = st.text_input(
                "5.4 Unidad / Dirección Interna Proveedora (si aplica) / Unidad Prov:",
                placeholder="Nombre de la unidad interna proveedora",
                key=f"unidadprov_{st.session_state.contador_guardado}",
            )
            inst_ext_prov = st.text_input(
                "5.5 Institución Externa Proveedora (si aplica) / Inst Ext Prov:",
                placeholder="Ejemplo: INEC, MAATE, MTOP, MAG, INAMHI",
                key=f"instext_{st.session_state.contador_guardado}",
            )

            st.markdown("---")
            st.header("Sección 6: Medios de Verificación y Flujos")
            medio_verif = st.multiselect(
                "6.1 Medio de Verificación Disponible / Medio Verificacion:",
                ["Físico (Archivo)", "Digital (Servidor/PC)", "Base de Datos", "Sistema Web"],
            )
            ruta_archivo = st.text_input(
                "6.2 Nombre de archivo, BD o Enlace del medio de verificación / Ruta/Enlace:",
                placeholder="Ruta de red, enlace a Google Drive o repositorio",
                key=f"ruta_{st.session_state.contador_guardado}",
            )
            difunde_terceros = st.radio("6.3 ¿Entrega o difunde este producto a terceros? / Difunde Terceros:", ["Sí", "No"])
            destinatarios = st.multiselect(
                "6.4 Destinatarios de la Información (si aplica) / Destinatarios:",
                ["Otras Direcciones GADPI", "GADs Cantonales / Parroquiales", "Ministerios", "Público en general"],
            )
            st.markdown("---")
            st.header("Sección 7: Gobernanza y Calidad")
            frec_act = st.selectbox(
                "7.1 Frecuencia de Actualización General / Frecuencia Act:",
                ["Continuo", "Mensual", "Trimestral", "Semestral", "Anual", "Por demanda", "No se actualizan"],
            )
            fecha_ultima = st.text_input(
                "7.2 Fecha de Última Actualización de la información (AAAA/MM) / Fecha Ultima Act:",
                placeholder="AAAA/MM",
                key=f"fechaultima_{st.session_state.contador_guardado}",
            )
            limitaciones = st.multiselect(
                "7.3 Principales Limitaciones para la Actualización / Limitaciones:",
                ["Falta personal técnico", "Restricciones presupuestarias", "Software obsoleto", "Equipamiento insuficiente", "Falta normativa"],
            )
            planificacion = st.multiselect(
                "7.4 Alineación Marco de Planificación / Alineacion Planif:",
                ["PDOT Imbabura", "POA Institucional", "ODS", "Competencias Ley / COOTAD"],
            )
            ficha_met = st.radio("7.5 ¿Cuenta con Ficha Metodológica Formalizada? / Ficha Metodologica:", ["Sí", "No", "En proceso"])
            uni_resp_calcul = st.text_input(
                "7.6 Unidad Responsable de la Ficha / Cálculo / Unidad Resp Calculo:",
                placeholder="Nombre del departamento o perfil técnico",
                key=f"uniresp_{st.session_state.contador_guardado}",
            )
            riesgos_preserv = st.multiselect(
                "7.7 Identificación de Riesgos de Preservación de la Información / Riesgos Preservacion:",
                ["Dependencia una persona", "Ausencia respaldos", "Virus/Fallos", "Rotación personal", "Deterioro papel"],
            )
            
            otra_razon_limitacion = st.text_input(
                "7.8 ¿Otra razón? / Otra Razon Limitacion:",
                placeholder="describa",
                key=f"otra_razon_{st.session_state.contador_guardado}"
            )

            st.markdown("---")
            st.header("Sección 8: Usos de la Información")
            uso_interno = st.text_area(
                "8.1 Uso Interno Actual de la Información / Uso Interno:",
                placeholder="Mencione quien hace uso de la información generada",
                key=f"usoint_{st.session_state.contador_guardado}",
            )
            uso_sil = st.text_area(
                "8.2 Potencial Uso / Integración en SIL GEO-IMBABURA / Integracion SIL:",
                placeholder="Como puede aprovecharse la información",
                key=f"usosil_{st.session_state.contador_guardado}",
            )
            nivel_acceso = st.radio("8.3 Nivel de Acceso de la Información / Nivel Acceso:", ["Público", "Restringido", "Uso Interno únicamente"])
            url_publicacion = st.text_input(
                "8.4 Plataforma / Enlace Web de Publicación (si aplica) / URL Publicacion:",
                placeholder="URL pública del geoportal o visor web",
                key=f"urlpub_{st.session_state.contador_guardado}",
            )

            st.markdown("---")
            
            if st.button("💾 Guardar Ficha de Diagnóstico", type="primary"):
                if not tecnico_resp:
                    st.warning("Complete el Nombre del Técnico Responsable en la Sección 1.")
                else:
                    # 🔄 ORDENAMIENTO EN SECUENCIA MATEMÁTICA EXACTA CON TU GOOGLE SHEET REAL
                    reg = {
                        "Direccion": dir_opcion,
                        "Subunidad": sub_opcion,
                        "Tecnico": tecnico_resp,
                        "Contacto": correo_ext,
                        "Producto": prod_opcion,
                        "Aplica Info": aplica_info,
                        "Tipo Informacion": tipo_informacion,
                        "nombre Ins Estad": nombre_ins_estad,
                        "Desagregacion Est": ", ".join(desag_est) if isinstance(desag_est, list) else desag_est,
                        "Cobertura Temporal": cobertura_est,
                        "nombre Insu Carto": nombre_insu_carto,
                        "genera cart": genera_cart,
                        "Desagregacion GIS": ", ".join(desag_gis) if isinstance(desag_gis, list) else desag_gis,
                        "Anio GIS": anio_gis,
                        "Escala GIS": escala_gis,
                        "Formato GIS": ", ".join(formato_gis) if isinstance(formato_gis, list) else formato_gis,
                        "Otro Formato GIS": otro_formato_gis,
                        "Genera Info Georreferenciada": genera_info_georref,
                        "Otras Fuentes GIS": otras_fuentes_gis,
                        "Tiene Metadatos": tiene_metadatos,
                        "Unidad Medida": unidad_medida,
                        "Fuente Origen": fuente_origen,
                        "Nombre Fuente": nombre_fuente,
                        "Unidad Prov": unidad_prov,
                        "Inst Ext Prov": inst_ext_prov,
                        "Medio Verificacion": ", ".join(medio_verif) if isinstance(medio_verif, list) else medio_verif,
                        "Ruta/Enlace": ruta_archivo,
                        "Difunde Terceros": difunde_terceros,
                        "Destinatarios": ", ".join(destinatarios) if isinstance(destinatarios, list) else destinatarios,
                        "Frecuencia Act": frec_act,
                        "Fecha Ultima Act": fecha_ultima,
                        "Limitaciones": ", ".join(limitaciones) if isinstance(limitaciones, list) else limitaciones,
                        "Otra Razon Limitacion": otra_razon_limitacion,
                        "Alineacion Planif": ", ".join(planificacion) if isinstance(planificacion, list) else planificacion,
                        "Ficha Metodologica": ficha_met,
                        "Unidad Resp Calculo": uni_resp_calcul,
                        "Riesgos Preservacion": ", ".join(riesgos_preserv) if isinstance(riesgos_preserv, list) else riesgos_preserv,
                        "Uso Interno": uso_interno,
                        "Integracion SIL": uso_sil,
                        "Nivel Acceso": nivel_acceso,
                        "URL Publicacion": url_publicacion,
                        "Fecha de Registro": pd.Timestamp.now().strftime("%Y/%m/%d"),
                    }
                    guardar_datos_nube(reg)
    except KeyError as e:
        st.error(f"Error al acoplar las columnas: {columnas}")
