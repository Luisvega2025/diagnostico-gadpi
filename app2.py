import os
import io
import time
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Ficha Diagnóstico GADPI - SIL",
    page_icon="🏛️",
    layout="centered"
)

EXCEL_MATRIZ = "matriz_gad.xlsx"
EXCEL_DIAGNOSTICO = "diagnostico_sil_gadpi_2026.xlsx"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]


# ============================================================
# CONEXIÓN A GOOGLE SHEETS
# ============================================================

@st.cache_resource
def get_sheet_connection():

    # --------------------------------------------------------
    # Leer configuración desde Streamlit Secrets
    # --------------------------------------------------------

    try:
        gsheets_config = st.secrets["connections"]["gsheets"]
    except Exception as e:
        raise RuntimeError(
            "No se encontró la configuración [connections.gsheets] "
            "en Streamlit Secrets."
        ) from e

    # Convertir Secrets a diccionario normal
    credentials_info = dict(gsheets_config)

    # --------------------------------------------------------
    # Obtener identificación de la hoja
    # --------------------------------------------------------

    spreadsheet = credentials_info.get("spreadsheet", "")
    worksheet_name = credentials_info.get("worksheet", "")

    if not spreadsheet:
        raise RuntimeError(
            "Falta 'spreadsheet' dentro de [connections.gsheets] "
            "en Streamlit Secrets."
        )

    # --------------------------------------------------------
    # Crear credenciales
    # --------------------------------------------------------

    campos_requeridos = [
        "type",
        "project_id",
        "private_key_id",
        "private_key",
        "client_email",
        "client_id",
        "auth_uri",
        "token_uri",
        "auth_provider_x509_cert_url",
        "client_x509_cert_url"
    ]

    faltantes = [
        campo
        for campo in campos_requeridos
        if campo not in credentials_info
    ]

    if faltantes:
        raise RuntimeError(
            "Faltan estos campos de la cuenta de servicio en Secrets: "
            + ", ".join(faltantes)
        )

    # Asegurar que la clave privada tenga saltos de línea correctos
    private_key = credentials_info["private_key"]

    if "\\n" in private_key:
        private_key = private_key.replace("\\n", "\n")

    credentials_info["private_key"] = private_key

    # --------------------------------------------------------
    # Autenticar
    # --------------------------------------------------------

    creds = Credentials.from_service_account_info(
        credentials_info,
        scopes=SCOPES
    )

    gc = gspread.authorize(creds)

    # --------------------------------------------------------
    # Abrir archivo de Google Sheets
    # --------------------------------------------------------

    if "docs.google.com" in spreadsheet:
        sh = gc.open_by_url(spreadsheet)
    else:
        sh = gc.open_by_key(spreadsheet)

    # --------------------------------------------------------
    # Seleccionar hoja
    # --------------------------------------------------------

    if worksheet_name:
        try:
            worksheet = sh.worksheet(worksheet_name)
        except Exception as e:
            raise RuntimeError(
                f"No se encontró la pestaña '{worksheet_name}' "
                f"dentro del archivo de Google Sheets."
            ) from e
    else:
        worksheet = sh.sheet1

    return worksheet


# ============================================================
# CARGAR MATRIZ EXCEL
# ============================================================

@st.cache_data
def cargar_matriz_limpia():

    if not os.path.exists(EXCEL_MATRIZ):
        return pd.DataFrame()

    df = pd.read_excel(EXCEL_MATRIZ, header=0)

    # Limpiar nombres de columnas
    df.columns = df.columns.astype(str).str.strip()

    # Quitar tildes de nombres de columnas
    reemplazos = {
        "é": "e",
        "ó": "o",
        "í": "i",
        "á": "a",
        "ú": "u",
        "É": "E",
        "Ó": "O",
        "Í": "I",
        "Á": "A",
        "Ú": "U",
    }

    for origen, destino in reemplazos.items():
        df.columns = df.columns.str.replace(
            origen,
            destino,
            regex=False
        )

    return df


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def convertir_lista(valor):

    if isinstance(valor, list):
        if not valor:
            return ""
        return ", ".join(str(x) for x in valor)

    return str(valor)


def guardar_respaldo_local(registro):

    try:
        df_nuevo = pd.DataFrame([registro])

        if os.path.exists(EXCEL_DIAGNOSTICO):
            df_existente = pd.read_excel(EXCEL_DIAGNOSTICO)

            df_consolidado = pd.concat(
                [df_existente, df_nuevo],
                ignore_index=True
            )
        else:
            df_consolidado = df_nuevo

        df_consolidado.to_excel(
            EXCEL_DIAGNOSTICO,
            index=False
        )

        return True

    except Exception as e:
        st.warning(
            "⚠️ El registro se guardó en Google Sheets, "
            "pero no se pudo actualizar el respaldo Excel local.\n\n"
            f"Detalle: {type(e).__name__}: {str(e)}"
        )
        return False


def guardar_datos_nube(registro):

    # --------------------------------------------------------
    # Intentar conectar
    # --------------------------------------------------------

    try:
        sheet = get_sheet_connection()

    except Exception as e:

        st.error(
            "❌ NO SE PUDO CONECTAR CON GOOGLE SHEETS"
        )

        st.code(
            f"{type(e).__name__}: {str(e)}",
            language="text"
        )

        st.warning(
            "Revisa la configuración de Secrets en "
            "Streamlit Cloud. El formulario NO se ha guardado."
        )

        return False

    # --------------------------------------------------------
    # Obtener encabezados
    # --------------------------------------------------------

    try:
        headers = sheet.row_values(1)

    except Exception as e:

        st.error(
            "❌ Se pudo autenticar con Google, "
            "pero no se pudo leer la primera fila de la hoja."
        )

        st.code(
            f"{type(e).__name__}: {str(e)}",
            language="text"
        )

        return False

    # --------------------------------------------------------
    # Si la hoja está completamente vacía
    # --------------------------------------------------------

    if not headers:

        headers = list(registro.keys())

        try:
            sheet.append_row(
                headers,
                value_input_option="USER_ENTERED"
            )

        except Exception as e:

            st.error(
                "❌ No se pudieron crear los encabezados "
                "en Google Sheets."
            )

            st.code(
                f"{type(e).__name__}: {str(e)}",
                language="text"
            )

            return False

    # --------------------------------------------------------
    # Preparar fila
    # --------------------------------------------------------

    fila_valores = []

    for columna in headers:

        valor = registro.get(columna, "")

        if isinstance(valor, list):
            valor = convertir_lista(valor)

        elif valor is None:
            valor = ""

        else:
            valor = str(valor)

        fila_valores.append(valor)

    # --------------------------------------------------------
    # Guardar fila en Google Sheets
    # --------------------------------------------------------

    try:

        sheet.append_row(
            fila_valores,
            value_input_option="USER_ENTERED"
        )

    except Exception as e:

        st.error(
            "❌ LA CONEXIÓN CON GOOGLE FUNCIONÓ, "
            "PERO NO SE PUDO ESCRIBIR LA FILA."
        )

        st.code(
            f"{type(e).__name__}: {str(e)}",
            language="text"
        )

        st.warning(
            "Esto normalmente indica un problema de permisos "
            "de la cuenta de servicio sobre el archivo de Google Sheets."
        )

        return False

    # --------------------------------------------------------
    # Respaldo local
    # --------------------------------------------------------

    guardar_respaldo_local(registro)

    # --------------------------------------------------------
    # Confirmación
    # --------------------------------------------------------

    st.success(
        "✅ ¡Datos guardados correctamente en Google Sheets!"
    )

    st.balloons()

    time.sleep(1)

    st.session_state.contador_guardado += 1

    st.rerun()

    return True


# ============================================================
# CARGAR MATRIZ
# ============================================================

df_matriz = cargar_matriz_limpia()


if df_matriz.empty:

    st.error(
        f"❌ No se encontró o no se pudo leer el archivo "
        f"'{EXCEL_MATRIZ}'."
    )

    st.info(
        "Verifica que el archivo matriz_gad.xlsx esté en el "
        "mismo repositorio que app2.py."
    )

    st.stop()


# ============================================================
# ESTADO DE SESIÓN
# ============================================================

if "contador_guardado" not in st.session_state:
    st.session_state.contador_guardado = 0


# ============================================================
# ENCABEZADO
# ============================================================

st.title(
    "DIRECCIÓN GENERAL DE PLANIFICACIÓN Y COOPERACIÓN"
)

st.title(
    "🏛️ Diagnóstico de Gestión de Información - GADPI"
)

st.write(
    "Ficha técnica oficial para el levantamiento de información, "
    "bases de datos y productos del SIL Geo-Imbabura."
)

st.info(
    "✉️ **¿Preguntas o información adicional?** "
    "lvega@imbabura.gob.ec"
)

st.markdown("---")


# ============================================================
# BARRA LATERAL
# ============================================================

with st.sidebar:

    st.subheader("🔑 Acceso Administrador SIL")

    clave_admin = st.text_input(
        "Ingrese la clave para descargar la base de datos:",
        type="password"
    )

    st.markdown("---")

    st.subheader("🔌 Estado de Google Sheets")

    if st.button("🔄 Probar conexión con Google Sheets"):

        try:

            hoja_prueba = get_sheet_connection()

            nombre_hoja = hoja_prueba.title

            st.success(
                f"✅ Conexión correcta.\n\n"
                f"Pestaña: {nombre_hoja}"
            )

        except Exception as e:

            st.error(
                "❌ Error de conexión"
            )

            st.code(
                f"{type(e).__name__}: {str(e)}",
                language="text"
            )


# ============================================================
# DESCARGA ADMINISTRATIVA
# ============================================================

if clave_admin == "gadpi2026":

    st.sidebar.success("Acceso Autorizado 🎈")

    if os.path.exists(EXCEL_DIAGNOSTICO):

        try:

            df_descarga = pd.read_excel(
                EXCEL_DIAGNOSTICO
            )

            buffer = io.BytesIO()

            with pd.ExcelWriter(
                buffer,
                engine="openpyxl"
            ) as writer:

                df_descarga.to_excel(
                    writer,
                    index=False
                )

            buffer.seek(0)

            st.sidebar.download_button(
                label="📥 Descargar Excel Consolidado",
                data=buffer,
                file_name="diagnostico_sil_gadpi_2026.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                )
            )

        except Exception as e:

            st.sidebar.error(
                f"Error al procesar el archivo: {e}"
            )

    else:

        st.sidebar.info(
            "Aún no se registran fichas técnicas "
            "en el respaldo local."
        )


# ============================================================
# IDENTIFICAR COLUMNAS DE LA MATRIZ
# ============================================================

columnas = list(df_matriz.columns)

col_dir = next(
    (
        c for c in columnas
        if "dir" in c.lower()
    ),
    None
)

col_sub = next(
    (
        c for c in columnas
        if (
            "sub" in c.lower()
            or "jef" in c.lower()
            or "dep" in c.lower()
        )
    ),
    None
)

col_prod = next(
    (
        c for c in columnas
        if (
            "prod" in c.lower()
            or "est" in c.lower()
        )
    ),
    None
)


if not col_dir or not col_sub or not col_prod:

    st.error(
        "❌ No se pudieron identificar correctamente "
        "las columnas Dirección, Subunidad o Producto "
        "en matriz_gad.xlsx."
    )

    st.write("Columnas encontradas:")

    st.write(columnas)

    st.stop()


# ============================================================
# FORMULARIO
# ============================================================

try:

    # ========================================================
    # SECCIÓN 1
    # ========================================================

    st.header(
        "Sección 1: Identificación del Informante"
    )

    direcciones = sorted(
        df_matriz[col_dir]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    dir_opcion = st.selectbox(
        "1.1 Seleccione Dirección:",
        direcciones
    )

    df_f_sub = df_matriz[
        df_matriz[col_dir] == dir_opcion
    ]

    subunidades = sorted(
        df_f_sub[col_sub]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    sub_opcion = st.selectbox(
        "1.2 Seleccione Subdirección / Jefatura / Unidad:",
        subunidades
    )

    tecnico_resp = st.text_input(
        "1.3 Nombre del Técnico Responsable del Llenado:",
        placeholder="Nombres y Apellidos completos",
        key=f"tecnico_{st.session_state.contador_guardado}"
    )

    correo_ext = st.text_input(
        "1.4 Correo Institucional / Contacto:",
        placeholder="correo; celular",
        key=f"correo_{st.session_state.contador_guardado}"
    )

    st.markdown("---")


    # ========================================================
    # SECCIÓN 2
    # ========================================================

    st.header(
        "Sección 2: Producto e Insumos según Estatuto 2026"
    )

    df_f_prod = df_f_sub[
        df_f_sub[col_sub] == sub_opcion
    ]

    productos = sorted(
        df_f_prod[col_prod]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    prod_opcion = st.selectbox(
        "2.1 Seleccione el Producto Institucional "
        "del Estatuto Orgánico:",
        productos
    )


    # --------------------------------------------------------
    # Aviso si ya existe el producto
    # --------------------------------------------------------

    try:

        sheet_check = get_sheet_connection()

        records = sheet_check.get_all_records()

        if records:

            df_check = pd.DataFrame(records)

            if "Producto" in df_check.columns:

                producto_existe = (
                    df_check["Producto"]
                    .astype(str)
                    .str.strip()
                    .eq(
                        str(prod_opcion).strip()
                    )
                    .any()
                )

                if producto_existe:

                    st.info(
                        "ℹ️ Este producto ya cuenta con "
                        "registros previos en la nube. "
                        "Estás agregando un nuevo "
                        "insumo/componente para este mismo producto."
                    )

    except Exception:

        # No detener el formulario si falla solamente
        # esta comprobación.
        pass


    aplica_info = st.radio(
        "2.2 ¿Aplica o genera información "
        "para cumplir con este producto?",
        ["Sí", "No"],
        key=f"aplica_{st.session_state.contador_guardado}"
    )

    tipo_informacion = st.radio(
        "2.3 Tipo de información que genera/utiliza?:",
        [
            "Alfanumérica / Estadística",
            "Geográfica"
        ],
        key=f"tipo_info_{st.session_state.contador_guardado}"
    )

    st.info(
        "ℹ️ **SI TIENE LOS DOS TIPOS DE INFORMACIÓN "
        "SE DEBE LLENAR UN REGISTRO A LA VEZ POR INSUMO "
        "(alfanumérico o Cartográfico)**"
    )


    # ========================================================
    # CASO: NO APLICA
    # ========================================================

    if aplica_info == "No":

        st.warning(
            "Ha seleccionado que NO aplica información "
            "para este producto. Guarde el registro para finalizar."
        )

        if st.button(
            "💾 Guardar Producto (No Aplica)",
            type="secondary"
        ):

            if not tecnico_resp:

                st.warning(
                    "Complete el Nombre del Técnico Responsable "
                    "en la Sección 1."
                )

            else:

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

                    "Fecha de Registro":
                        pd.Timestamp.now().strftime(
                            "%Y/%m/%d %H:%M:%S"
                        )
                }

                guardar_datos_nube(reg)


    # ========================================================
    # CASO: SÍ APLICA
    # ========================================================

    else:

        # Variables por defecto
        nombre_ins_estad = "No aplica"
        desag_est = ["No aplica"]
        cobertura_est = "No aplica"

        nombre_insu_carto = "No aplica"
        genera_cart = "No aplica"
        desag_gis = ["No aplica"]
        anio_gis = "No aplica"
        escala_gis = "No aplica"
        formato_gis = ["No aplica"]
        otro_formato_gis = "No aplica"
        genera_info_georref = "No aplica"
        otras_fuentes_gis = "No aplica"
        tiene_metadatos = "No aplica"


        # ====================================================
        # SECCIÓN 3 - ESTADÍSTICA
        # ====================================================

        if tipo_informacion == "Alfanumérica / Estadística":

            st.markdown("---")

            st.header(
                "Sección 3: Datos Alfanuméricos/Estadísticos"
            )

            nombre_ins_estad = st.text_input(
                "3.1 ¿Nombre del insumo de información "
                "estadística / alfanumérica que aporta a este producto?",
                placeholder="ejem: usuarios_canal_riego.*/doc/pdf/xls/",
                key=f"insumo_est_v3_{st.session_state.contador_guardado}"
            )

            desag_est = st.multiselect(
                "3.2 Seleccione el nivel de desagregación "
                "de la información estadística:",
                [
                    "Provincial",
                    "Cantonal",
                    "Parroquial",
                    "Sector / Comunidad",
                    "Predio / Proyecto"
                ],
                key=f"desag_est_v3_{st.session_state.contador_guardado}"
            )

            cobertura_est = st.text_input(
                "3.3 Temporalidad de los Datos Estadísticos:",
                placeholder="Ejemplo: 2018 - 2026",
                key=f"cobertura_v3_{st.session_state.contador_guardado}"
            )


        # ====================================================
        # SECCIÓN 4 - GIS
        # ====================================================

        else:

            st.markdown("---")

            st.header(
                "Sección 4: Datos Geográficos (GIS)"
            )

            nombre_insu_carto = st.text_input(
                "4.1 ¿Nombre del insumo cartográfico "
                "que aporta a este producto?",
                placeholder="ejem: vias.shp/*nombre.mxd/nombre.gdb",
                key=f"insumo_carto_v3_{st.session_state.contador_guardado}"
            )

            genera_cart = st.radio(
                "4.2 ¿El insumo es generado "
                "por la unidad o jefatura?",
                ["Sí", "No"],
                key=f"genera_cart_v3_{st.session_state.contador_guardado}"
            )

            desag_gis = st.multiselect(
                "4.3 Seleccione el nivel de desagregación "
                "de la información Geográfica:",
                [
                    "Provincial",
                    "Cantonal",
                    "Parroquial",
                    "Sector / Comunidad",
                    "Predio / Proyecto"
                ],
                key=f"desag_gis_v3_{st.session_state.contador_guardado}"
            )

            anio_gis = st.text_input(
                "4.4 Año de la información cartográfica:",
                placeholder="Ejemplo: 2020 - 2026",
                key=f"aniogis_v3_{st.session_state.contador_guardado}"
            )

            escala_gis = st.selectbox(
                "4.5 Escala de la cartografía:",
                [
                    "1:5.000",
                    "1:25.000",
                    "1:50.000",
                    "1:100.000",
                    "No"
                ],
                key=f"escala_gis_v3_{st.session_state.contador_guardado}"
            )

            formato_gis = st.multiselect(
                "4.6 Formato de los Datos Geográficos:",
                [
                    "File Geodatabase (.gdb)",
                    "Shapefile (.shp)",
                    "GeoJSON / KML",
                    "Tabla XY (Excel / CSV)",
                    "Servicio Web (WMS/WFS)"
                ],
                key=f"formato_gis_v3_{st.session_state.contador_guardado}"
            )

            otro_formato_gis = st.text_input(
                "4.7 ¿Otro formato?:",
                key=f"otro_formato_v3_{st.session_state.contador_guardado}"
            )

            genera_info_georref = st.text_input(
                "4.8 ¿Qué metodología utiliza para "
                "generar la información cartográfica?",
                placeholder="ejem: mediante topografía, vuelo con drones, etc.",
                key=f"genera_georref_v3_{st.session_state.contador_guardado}"
            )

            otras_fuentes_gis = st.text_input(
                "4.9 ¿Obtiene información espacial "
                "de otras fuentes? ¿Cuáles?",
                placeholder="ejem: IGM, INEC, MAG, etc.",
                key=f"otras_fuentes_v3_{st.session_state.contador_guardado}"
            )

            tiene_metadatos = st.text_input(
                "4.10 ¿La cartografía utilizada tiene "
                "metadatos, catálogo de objetos?",
                placeholder="sí/no",
                key=f"metadatos_v3_{st.session_state.contador_guardado}"
            )


        # ====================================================
        # SECCIÓN 5
        # ====================================================

        st.markdown("---")

        st.header(
            "Sección 5: Fuentes y Origen del Dato"
        )

        if tipo_informacion == "Alfanumérica / Estadística":

            unidad_medida = st.selectbox(
                "5.1 Unidad de Medida del Dato / Indicador:",
                [
                    "Kilómetros",
                    "Hectáreas",
                    "Porcentaje",
                    "Número de usuarios",
                    "Unidades",
                    "No aplica"
                ],
                key=f"unidad_medida_v3_{st.session_state.contador_guardado}"
            )

        else:

            unidad_medida = "No aplica"


        fuente_origen = st.selectbox(
            "5.2 Fuente de Origen del Dato:",
            [
                "Interno GADPI",
                "Entidad Externa",
                "Mixto"
            ],
            key=f"fuente_origen_v3_{st.session_state.contador_guardado}"
        )

        nombre_fuente = st.text_input(
            "5.3 Nombre del producto / insumo:",
            placeholder="Nombre del sistema, censo, catastro o plataforma",
            key=f"fuente_v3_{st.session_state.contador_guardado}"
        )

        unidad_prov = st.text_input(
            "5.4 En caso de que el proveedor sea interno "
            "nombre de la Unidad / Dirección:",
            placeholder="Nombre de la unidad interna proveedora",
            key=f"unidadprov_v3_{st.session_state.contador_guardado}"
        )

        inst_ext_prov = st.text_input(
            "5.5 En caso de proveedor externo - "
            "nombre de la Institución:",
            placeholder="Ejemplo: INEC, MAATE, MTOP, MAG, INAMHI",
            key=f"instext_v3_{st.session_state.contador_guardado}"
        )


        # ====================================================
        # SECCIÓN 6
        # ====================================================

        st.markdown("---")

        st.header(
            "Sección 6: Medios de Verificación y Flujos"
        )

        medio_verif = st.multiselect(
            "6.1 Medio de Verificación Disponible:",
            [
                "Físico (Archivo)",
                "Digital (Servidor/PC)",
                "Base de Datos",
                "Sistema Web",
                "No existen"
            ],
            key=f"medio_verif_v3_{st.session_state.contador_guardado}"
        )

        ruta_archivo = st.text_input(
            "6.2 Nombre de archivo, BD o Enlace "
            "del medio de verificación:",
            placeholder="almacenamiento local, Ruta de red, enlace a Google Drive o repositorio",
            key=f"ruta_v3_{st.session_state.contador_guardado}"
        )

        difunde_terceros = st.radio(
            "6.3 ¿Entrega o difunde este producto a terceros?",
            ["Sí", "No"],
            key=f"difunde_v3_{st.session_state.contador_guardado}"
        )

        destinatarios = st.multiselect(
            "6.4 Destinatarios de la Información (si aplica):",
            [
                "Otras Direcciones GADPI",
                "GADs Cantonales / Parroquiales",
                "Ministerios",
                "Público en general",
                "Ciudadanía"
            ],
            key=f"destinatarios_v3_{st.session_state.contador_guardado}"
        )


        # ====================================================
        # SECCIÓN 7
        # ====================================================

        st.markdown("---")

        st.header(
            "Sección 7: Gobernanza y Calidad"
        )

        frec_act = st.selectbox(
            "7.1 Frecuencia de Actualización General:",
            [
                "Continuo",
                "Mensual",
                "Trimestral",
                "Semestral",
                "Anual",
                "Por demanda",
                "No se actualizan"
            ],
            key=f"frec_act_v3_{st.session_state.contador_guardado}"
        )

        fecha_ultima = st.text_input(
            "7.2 Fecha de Última Actualización "
            "de la información (AAAA/MM):",
            placeholder="AAAA/MM",
            key=f"fechaultima_v3_{st.session_state.contador_guardado}"
        )

        limitaciones = st.multiselect(
            "7.3 Principales Limitaciones para la Actualización:",
            [
                "Falta personal técnico",
                "Restricciones presupuestarias",
                "Problemas tecnológicos/conectividad",
                "Equipamiento insuficiente",
                "Falta normativa",
                "No hay acceso a fuentes primarias de información"
            ],
            key=f"limitaciones_v3_{st.session_state.contador_guardado}"
        )

        planificacion = st.multiselect(
            "7.4 ¿La información gestionada está alineada "
            "con instrumentos de Planificación?:",
            [
                "Indicadores institucionales",
                "PDOT Imbabura",
                "POA Institucional",
                "ODS",
                "Competencias Ley / COOTAD"
            ],
            key=f"planificacion_v3_{st.session_state.contador_guardado}"
        )

        ficha_met = st.radio(
            "7.5 Si la información gestionada está consolidada "
            "en un indicador, ¿Cuenta con Ficha Metodológica?:",
            [
                "Sí",
                "No",
                "En proceso"
            ],
            key=f"ficha_met_v3_{st.session_state.contador_guardado}"
        )

        uni_resp_calcul = st.text_input(
            "7.6 Unidad Responsable de la Ficha / Cálculo "
            "(Si aplica):",
            placeholder="Nombre del departamento o perfil técnico",
            key=f"uniresp_v3_{st.session_state.contador_guardado}"
        )

        riesgos_preserv = st.multiselect(
            "7.7 Identifique los Riesgos para Preservar "
            "la Información gestionada por su unidad:",
            [
                "Dependencia una persona",
                "Ausencia respaldos",
                "Virus/Fallos",
                "Rotación de personal",
                "No aplica"
            ],
            key=f"riesgos_v3_{st.session_state.contador_guardado}"
        )

        otra_razon_limitacion = st.text_input(
            "7.8 ¿Otra razón?:",
            placeholder="describa",
            key=f"otra_razon_v3_{st.session_state.contador_guardado}"
        )


        # ====================================================
        # SECCIÓN 8
        # ====================================================

        st.markdown("---")

        st.header(
            "Sección 8: Usos de la Información"
        )

        uso_interno = st.text_area(
            "8.1 ¿Mencione otros usos de la Información "
            "gestionada por su unidad?:",
            placeholder=(
                "Ej: en planificación, informes de gestión, "
                "atención ciudadana, etc."
            ),
            key=f"usoint_v3_{st.session_state.contador_guardado}"
        )

        uso_sil = st.text_area(
            "8.2 Observaciones / Recomendaciones:",
            placeholder=(
                "compartir con otras instituciones, "
                "para visualización pública, generar alertas, etc."
            ),
            key=f"usosil_v3_{st.session_state.contador_guardado}"
        )

        nivel_acceso = st.radio(
            "8.3 Nivel de Acceso de la Información:",
            [
                "Público",
                "Restringido",
                "Uso Interno únicamente"
            ],
            key=f"nivel_acceso_v3_{st.session_state.contador_guardado}"
        )

        url_publicacion = st.text_input(
            "8.4 Plataforma / Enlace Web de Publicación (si aplica):",
            placeholder="donde publica, URL del geoportal o visor web",
            key=f"urlpub_v3_{st.session_state.contador_guardado}"
        )


        # ====================================================
        # BOTÓN GUARDAR
        # ====================================================

        st.markdown("---")

        if st.button(
            "💾 Guardar Ficha de Diagnóstico",
            type="primary"
        ):

            if not tecnico_resp.strip():

                st.warning(
                    "Complete el Nombre del Técnico Responsable "
                    "en la Sección 1."
                )

            else:

                reg = {

                    "Direccion": dir_opcion,
                    "Subunidad": sub_opcion,
                    "Tecnico": tecnico_resp,
                    "Contacto": correo_ext,
                    "Producto": prod_opcion,
                    "Aplica Info": aplica_info,
                    "Tipo Informacion": tipo_informacion,

                    "nombre Ins Estad":
                        nombre_ins_estad,

                    "Desagregacion Est":
                        convertir_lista(desag_est),

                    "Cobertura Temporal":
                        cobertura_est,

                    "nombre Insu Carto":
                        nombre_insu_carto,

                    "genera cart":
                        genera_cart,

                    "Desagregacion GIS":
                        convertir_lista(desag_gis),

                    "Anio GIS":
                        anio_gis,

                    "Escala GIS":
                        escala_gis,

                    "Formato GIS":
                        convertir_lista(formato_gis),

                    "Otro Formato GIS":
                        otro_formato_gis,

                    "Genera Info Georreferenciada":
                        genera_info_georref,

                    "Otras Fuentes GIS":
                        otras_fuentes_gis,

                    "Tiene Metadatos":
                        tiene_metadatos,

                    "Unidad Medida":
                        unidad_medida,

                    "Fuente Origen":
                        fuente_origen,

                    "Nombre Fuente":
                        nombre_fuente,

                    "Unidad Prov":
                        unidad_prov,

                    "Inst Ext Prov":
                        inst_ext_prov,

                    "Medio Verificacion":
                        convertir_lista(medio_verif),

                    "Ruta/Enlace":
                        ruta_archivo,

                    "Difunde Terceros":
                        difunde_terceros,

                    "Destinatarios":
                        convertir_lista(destinatarios),

                    "Frecuencia Act":
                        frec_act,

                    "Fecha Ultima Act":
                        fecha_ultima,

                    "Limitaciones":
                        convertir_lista(limitaciones),

                    "Otra Razon Limitacion":
                        otra_razon_limitacion,

                    "Alineacion Planif":
                        convertir_lista(planificacion),

                    "Ficha Metodologica":
                        ficha_met,

                    "Unidad Resp Calculo":
                        uni_resp_calcul,

                    "Riesgos Preservacion":
                        convertir_lista(riesgos_preserv),

                    "Uso Interno":
                        uso_interno,

                    "Integracion SIL":
                        uso_sil,

                    "Nivel Acceso":
                        nivel_acceso,

                    "URL Publicacion":
                        url_publicacion,

                    "Fecha de Registro":
                        pd.Timestamp.now().strftime(
                            "%Y/%m/%d %H:%M:%S"
                        )
                }

                guardar_datos_nube(reg)


except Exception as e:

    st.error(
        "❌ Se produjo un error inesperado en el formulario."
    )

    st.code(
        f"{type(e).__name__}: {str(e)}",
        language="text"
    )
