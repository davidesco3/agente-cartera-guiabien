import io
from pathlib import Path
import streamlit as st
import pandas as pd

from config import (
    BRAND_PRIMARY,
    BRAND_SECONDARY,
    BRAND_ACCENT,
    BRAND_DARK,
    BRAND_LIGHT_BG,
    CLIENTES_ELITE_PATH,
    LOGO_PATH,
    GOOGLE_SHEETS_NO_CONTACTAR_URL,
)
from cartera_processor import (
    CarteraPipeline,
    ClientesEliteManager,
    NoContactarService,
    read_uploaded_file,
)
from excel_exporter import export_cartera_to_excel
from file_auditor import FileAuditor

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="Depuración de Cartera | GuíaBien Asesores",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ==============================================================================
# ESTILOS CORPORATIVOS GUIABIEN (LOOK & FEEL GUIABIEN.COM)
# ==============================================================================
CUSTOM_CSS = f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Segoe+UI:wght@400;600;700&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
    }}

    .stApp {{
        background-color: {BRAND_LIGHT_BG};
    }}

    .gb-header-container {{
        background: linear-gradient(135deg, {BRAND_PRIMARY} 0%, #101366 100%);
        border-radius: 12px;
        padding: 24px 30px;
        color: #FFFFFF;
        box-shadow: 0 4px 14px rgba(30, 34, 170, 0.15);
        margin-bottom: 24px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }}
    .gb-header-title {{
        font-size: 26px;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.5px;
        color: #FFFFFF;
    }}
    .gb-header-subtitle {{
        font-size: 14px;
        color: {BRAND_ACCENT};
        margin-top: 4px;
        font-weight: 400;
    }}

    .gb-metric-card {{
        background: #FFFFFF;
        border-radius: 10px;
        padding: 16px 20px;
        border-left: 5px solid {BRAND_PRIMARY};
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        margin-bottom: 12px;
        transition: transform 0.15s ease-in-out;
    }}
    .gb-metric-card:hover {{
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(0,0,0,0.08);
    }}
    .gb-metric-num {{
        font-size: 28px;
        font-weight: 700;
        color: {BRAND_PRIMARY};
        margin: 0;
    }}
    .gb-metric-label {{
        font-size: 13px;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}

    /* Botón Primario GuíaBien */
    div.stButton > button:first-child {{
        background: linear-gradient(135deg, {BRAND_PRIMARY} 0%, #2930D9 100%);
        color: #FFFFFF;
        font-weight: 700;
        font-size: 16px;
        padding: 14px 28px;
        border-radius: 8px;
        border: none;
        box-shadow: 0 4px 12px rgba(30, 34, 170, 0.25);
        transition: all 0.2s ease-in-out;
        width: 100%;
    }}
    div.stButton > button:first-child:hover {{
        background: linear-gradient(135deg, #16198C 0%, {BRAND_PRIMARY} 100%);
        box-shadow: 0 6px 16px rgba(30, 34, 170, 0.35);
        color: #FFFFFF;
    }}

    /* Botón de Descarga Excel */
    div.stDownloadButton > button:first-child {{
        background: linear-gradient(135deg, {BRAND_SECONDARY} 0%, #059669 100%);
        color: #FFFFFF;
        font-weight: 700;
        font-size: 16px;
        padding: 14px 28px;
        border-radius: 8px;
        border: none;
        box-shadow: 0 4px 12px rgba(0, 208, 132, 0.3);
        width: 100%;
        transition: all 0.2s ease-in-out;
    }}
    div.stDownloadButton > button:first-child:hover {{
        background: linear-gradient(135deg, #047857 0%, {BRAND_SECONDARY} 100%);
        color: #FFFFFF;
        box-shadow: 0 6px 16px rgba(0, 208, 132, 0.4);
    }}

    /* Sidebar Styling */
    section[data-testid="stSidebar"] {{
        background-color: #FFFFFF;
        border-right: 1px solid #E2E8F0;
    }}

    .gb-alert-badge {{
        background-color: #FEE2E2;
        color: #991B1B;
        padding: 4px 8px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 700;
        display: inline-block;
    }}

    .gb-upload-box {{
        background: #FFFFFF;
        border-radius: 10px;
        padding: 16px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 4px rgba(0,0,0,0.03);
        height: 100%;
    }}

    /* Tarjeta de Auditoría / Pre-vuelo */
    .gb-audit-card {{
        background: #F8FAFC;
        border-radius: 10px;
        padding: 18px 24px;
        border: 1px solid #CBD5E1;
        margin-top: 15px;
        margin-bottom: 20px;
    }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ==============================================================================
# INICIALIZACIÓN DE SERVICIOS
# ==============================================================================
def get_services():
    elite_mgr = ClientesEliteManager(CLIENTES_ELITE_PATH)
    no_contactar_srv = NoContactarService(GOOGLE_SHEETS_NO_CONTACTAR_URL)
    pipeline = CarteraPipeline(elite_mgr, no_contactar_srv)
    return elite_mgr, no_contactar_srv, pipeline


elite_manager, no_contactar_service, pipeline = get_services()


# ==============================================================================
# MODAL DISCRETO: GESTIÓN DE CLIENTES ÉLITE (OCULTO EN BOTÓN DE CONFIGURACIÓN)
# ==============================================================================
@st.dialog("👑 Gestión de Clientes Élite", width="large")
def modal_gestion_elite():
    st.caption(
        "Administra la lista oficial de clientes élite de GuíaBien. "
        "Puedes modificar la tabla en vivo, agregar nuevos documentos o dar de baja registros."
    )
    df_elite = elite_manager.get_dataframe()

    t_tabla, t_agregar, t_eliminar = st.tabs(["📋 Lista Editable en Vivo", "➕ Agregar Cliente", "🗑️ Eliminar Cliente"])

    with t_tabla:
        st.info("💡 Haz doble clic en cualquier celda para editar, o usa la última fila vacía para añadir nuevos:")
        edited_df = st.data_editor(
            df_elite,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Nº ident.fis.1": st.column_config.TextColumn("Cédula / NIT", required=True),
                "Nombre Tomador": st.column_config.TextColumn("Nombre Tomador"),
                "Segmento": st.column_config.TextColumn("Segmento", default="Élite"),
            },
            key="editor_elite_modal",
        )
        if st.button("💾 Guardar Cambios en la Base", type="primary", use_container_width=True, key="btn_save_table_elite"):
            success, msg, _ = elite_manager.save_dataframe(edited_df)
            if success:
                st.success(f"✅ {msg}")
                st.rerun()
            else:
                st.error(f"❌ {msg}")

    with t_agregar:
        st.markdown("##### Registrar nuevo cliente al segmento prioritario")
        c_dni, c_nom = st.columns([1, 2])
        with c_dni:
            nuevo_dni = st.text_input("Número de Documento (Cédula/NIT)", key="modal_input_dni")
        with c_nom:
            nuevo_nombre = st.text_input("Nombre Completo del Tomador", key="modal_input_nombre")

        if st.button("➕ Guardar Nuevo Cliente", type="primary", key="btn_save_single_elite"):
            if nuevo_dni.strip():
                success, msg = elite_manager.add_client(nuevo_dni, nuevo_nombre)
                if success:
                    st.success(f"✅ {msg}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")
            else:
                st.warning("Por favor ingresa un número de documento válido.")

    with t_eliminar:
        st.markdown("##### Retirar cliente de la lista de clientes élite")
        if not df_elite.empty:
            opciones = [f"{row['Nº ident.fis.1']} - {row['Nombre Tomador']}" for _, row in df_elite.iterrows()]
            seleccion = st.selectbox("Selecciona el cliente a retirar:", opciones, key="select_del_elite_modal")
            dni_del = seleccion.split(" - ")[0].strip() if seleccion else ""

            if st.button("🗑️ Eliminar de Clientes Élite", type="secondary", key="btn_confirm_del_elite"):
                success, msg = elite_manager.remove_client(dni_del)
                if success:
                    st.success(f"✅ {msg}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")
        else:
            st.info("No hay clientes registrados en la lista élite.")


# ==============================================================================
# MODAL DISCRETO: CONSULTA DE CLIENTES A NO CONTACTAR (DESDE GOOGLE SHEETS)
# ==============================================================================
@st.dialog("🚫 Clientes a No Contactar (Google Sheets)", width="large")
def modal_no_contactar():
    st.caption(
        "Listado oficial de clientes que han solicitado no ser contactados o poseen restricciones operativas. "
        "Esta lista se sincroniza automáticamente desde la pestaña **Cartera_Excluidos** en Google Sheets y no se edita en la app para mantener la integridad centralizada."
    )

    with st.spinner("Consultando lista en Google Sheets..."):
        df_nc, fuente_nc, estado_nc = no_contactar_service.fetch_dataframe()

    unique_clients = df_nc["_id_clean"].nunique() if "_id_clean" in df_nc.columns else len(df_nc)
    df_display = df_nc.drop(columns=["_id_clean"]) if "_id_clean" in df_nc.columns else df_nc

    c_m1, c_m2, c_m3, c_m4 = st.columns([1, 1, 1.4, 1.5])
    with c_m1:
        st.metric("Pólizas", f"{len(df_display):,}")
    with c_m2:
        st.metric("Tomadores Únicos", f"{unique_clients:,}")
    with c_m3:
        st.metric("Origen de Datos", fuente_nc)
    with c_m4:
        st.markdown("<div style='height: 18px;'></div>", unsafe_allow_html=True)
        st.link_button(
            "🔗 Abrir y Editar en Sheets",
            no_contactar_service.view_url,
            type="primary",
            use_container_width=True,
            help="Abre la pestaña Cartera_Excluidos en Google Sheets para agregar, editar o retirar documentos.",
        )

    if estado_nc != "OK":
        st.warning(f"⚠️ **Estado de sincronización:** {estado_nc}")
    else:
        st.success("🟢 **Sincronización en vivo activa:** Todo cambio realizado en el Google Sheet se aplicará automáticamente.")

    st.markdown("##### 📋 Pólizas y Clientes Excluidos (Solo Lectura)")
    if not df_display.empty:
        col_configs = {}
        if "Nº ident" in df_display.columns:
            col_configs["Nº ident"] = st.column_config.TextColumn("Cédula / NIT")
        elif "Nº ident.fis.1" in df_display.columns:
            col_configs["Nº ident.fis.1"] = st.column_config.TextColumn("Cédula / NIT")

        if "Póliza" in df_display.columns:
            col_configs["Póliza"] = st.column_config.TextColumn("Póliza")

        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True,
            column_config=col_configs,
        )
    else:
        st.info("No se encontraron registros de clientes a no contactar.")

    if st.button("🔄 Recargar datos desde Google Sheets", key="btn_refresh_sheets_modal", use_container_width=True):
        st.rerun()


# ==============================================================================
# SIDEBAR MINIMALISTA (SIN DISTRACTORES PARA EL OPERADOR)
# ==============================================================================
with st.sidebar:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=180)
    else:
        st.markdown(f"<h3 style='color:{BRAND_PRIMARY}; font-weight:700;'>GuíaBien</h3>", unsafe_allow_html=True)

    st.markdown("<hr style='margin: 12px 0;'>", unsafe_allow_html=True)
    st.markdown("#### ℹ️ Estado del Sistema")
    st.caption("Sistema de depuración de cartera semanal integrado con Seguros SURA.")
    
    total_elite_count = elite_manager.count()
    total_nc_count = no_contactar_service.count()
    st.markdown(
        f"""
        <div style="font-size: 13px; color: #475569; margin-bottom: 8px;">
            👑 <strong>Clientes Élite:</strong> {total_elite_count:,} registrados
        </div>
        <div style="font-size: 13px; color: #475569;">
            🚫 <strong>No Contactar:</strong> {total_nc_count:,} sincronizados (Google Sheets)
        </div>
        """,
        unsafe_allow_html=True
    )
    st.markdown("<br>", unsafe_allow_html=True)
    st.caption("GuíaBien Asesores en Seguros S.A.S. — Medellín, Colombia")


# ==============================================================================
# ENCABEZADO PRINCIPAL (LOOK & FEEL GUIABIEN.COM)
# ==============================================================================
col_head_1, col_head_2, col_head_3 = st.columns([3.4, 1.3, 1.3])
with col_head_1:
    st.markdown(
        f"""
        <div class="gb-header-container">
            <div>
                <h1 class="gb-header-title">🛡️ Depuración Operativa de Cartera</h1>
                <div class="gb-header-subtitle">GuíaBien Asesores en Seguros | Intermediario SURA</div>
            </div>
            <div style="text-align: right; font-size: 13px; opacity: 0.9;">
                <div><strong>Ambiente Operativo</strong></div>
                <div>Versión 2.2 Productiva</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col_head_2:
    st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
    if st.button(
        f"⚙️ Clientes Élite ({elite_manager.count()})",
        key="btn_open_elite_modal",
        help="Abrir panel administrativo discreto para agregar o eliminar clientes élite",
        use_container_width=True,
    ):
        modal_gestion_elite()

with col_head_3:
    st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
    if st.button(
        f"🚫 No Contactar ({no_contactar_service.count()})",
        key="btn_open_no_contactar_modal",
        help="Ver listado oficial de clientes excluidos de contacto sincronizado desde Google Sheets",
        use_container_width=True,
    ):
        modal_no_contactar()

st.markdown(
    """
    Bienvenido al depurador inteligente de cartera. Carga los tres informes operativos de la semana.
    El sistema cuenta con **auditoría automática antierrores**: valida hashes antiduplicados, verifica esquemas de datos
    y comprueba la coherencia cronológica para evitar inversiones o datos desactualizados.
    """
)

# ==============================================================================
# SECCIÓN 3: INTERFAZ DE CARGA CON BLINDAJE ANTIERRORES
# ==============================================================================
st.markdown("### 📂 Carga de Archivos Operativos Semanales")

tabs_upload = st.tabs(["📁 Carga Individual Guiada (Recomendada)", "⚡ Buzón Inteligente (Arrastra los 3 juntos)"])

file_actual = None
file_anterior = None
file_recaudos = None

with tabs_upload[0]:
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(
            """
            <div class="gb-upload-box">
                <h4 style="margin-top:0; color:#1E293B;">1. Cartera Actual (Semana en curso)</h4>
                <p style="font-size:12px; color:#64748B;">Informe de cartera emitido esta semana.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        file_actual = st.file_uploader(
            "Cargar Cartera Actual",
            type=["xlsx", "xls", "csv"],
            key="file_actual",
            label_visibility="collapsed",
        )

    with col2:
        st.markdown(
            """
            <div class="gb-upload-box">
                <h4 style="margin-top:0; color:#1E293B;">2. Cartera Semana Anterior</h4>
                <p style="font-size:12px; color:#64748B;">Corte de la semana pasada (más antiguo).</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        file_anterior = st.file_uploader(
            "Cargar Cartera Anterior",
            type=["xlsx", "xls", "csv"],
            key="file_anterior",
            label_visibility="collapsed",
        )

    with col3:
        st.markdown(
            """
            <div class="gb-upload-box">
                <h4 style="margin-top:0; color:#1E293B;">3. Recaudos de la Semana</h4>
                <p style="font-size:12px; color:#64748B;">Pólizas pagadas durante el ciclo reciente.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        file_recaudos = st.file_uploader(
            "Cargar Informe Recaudos",
            type=["xlsx", "xls", "csv"],
            key="file_recaudos",
            label_visibility="collapsed",
        )

with tabs_upload[1]:
    st.info("💡 **Buzón Inteligente:** Selecciona o arrastra los 3 archivos al mismo tiempo. El sistema analizará las columnas internas y las fechas de corte para identificar automáticamente cuál es Cartera Actual, cuál es Anterior y cuál es Recaudos.")
    batch_files = st.file_uploader(
        "Arrastra aquí tus 3 archivos simultáneamente",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="batch_uploader",
    )
    if batch_files:
        if len(batch_files) >= 2:
            files_dict = {}
            for bf in batch_files:
                bf_bytes = bf.getvalue()
                df_temp = read_uploaded_file(bf)
                files_dict[bf.name] = (bf_bytes, df_temp)

            classification = FileAuditor.auto_classify_batch(files_dict)
            if classification["cartera_actual"]:
                st.success(f"✅ **Cartera Actual detectada:** `{classification['cartera_actual']['name']}` (Fecha: {classification['cartera_actual']['date'] or 'No especificada'})")
                file_actual = [bf for bf in batch_files if bf.name == classification['cartera_actual']['name']][0]
            if classification["cartera_anterior"]:
                st.success(f"✅ **Cartera Anterior detectada:** `{classification['cartera_anterior']['name']}` (Fecha: {classification['cartera_anterior']['date'] or 'No especificada'})")
                file_anterior = [bf for bf in batch_files if bf.name == classification['cartera_anterior']['name']][0]
            if classification["recaudos"]:
                st.success(f"✅ **Recaudos detectados:** `{classification['recaudos']['name']}`")
                file_recaudos = [bf for bf in batch_files if bf.name == classification['recaudos']['name']][0]
        else:
            st.warning("Por favor sube al menos 2 o 3 archivos para que el clasificador automático pueda discernir entre ellos.")


# ==============================================================================
# AUDITORÍA DE PRE-VUELO (BLINDAJE DE INTEGRIDAD)
# ==============================================================================
audit_result = None
can_proceed_processing = False

if file_actual is not None:
    try:
        b_act = file_actual.getvalue()
        df_act = read_uploaded_file(file_actual)

        b_ant = file_anterior.getvalue() if file_anterior is not None else None
        df_ant = read_uploaded_file(file_anterior) if file_anterior is not None else None

        b_rec = file_recaudos.getvalue() if file_recaudos is not None else None
        df_rec = read_uploaded_file(file_recaudos) if file_recaudos is not None else None

        audit_result = FileAuditor.audit_upload(
            file_actual_bytes=b_act,
            file_actual_name=file_actual.name,
            df_actual=df_act,
            file_anterior_bytes=b_ant,
            file_anterior_name=file_anterior.name if file_anterior else None,
            df_anterior=df_ant,
            file_recaudos_bytes=b_rec,
            file_recaudos_name=file_recaudos.name if file_recaudos else None,
            df_recaudos=df_rec,
        )

        can_proceed_processing = audit_result["can_process"]

        # Mostrar tarjeta de Pre-vuelo
        st.markdown(
            f"""
            <div class="gb-audit-card">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <h4 style="margin:0; color:{BRAND_PRIMARY};">🛡️ Semáforo de Pre-vuelo y Coherencia de Archivos</h4>
                    <div>{'<span style="background:#DCFCE7; color:#166534; font-weight:700; padding:4px 10px; border-radius:6px; font-size:12px;">COHERENCIA VERIFICADA ✅</span>' if can_proceed_processing else '<span style="background:#FEE2E2; color:#991B1B; font-weight:700; padding:4px 10px; border-radius:6px; font-size:12px;">BLOQUEADO POR INCOHERENCIA ⛔</span>'}</div>
                </div>
                <div style="font-size:12px; color:#64748B; margin-top:6px;">
                    Fechas de corte detectadas: 
                    <strong>Cartera Actual:</strong> {audit_result['dates']['actual']} | 
                    <strong>Cartera Anterior:</strong> {audit_result['dates']['anterior']} | 
                    <strong>Recaudos:</strong> {audit_result['dates']['recaudos']}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # Si hay errores bloqueantes
        if audit_result["errors"]:
            for err in audit_result["errors"]:
                st.error(err)

            if audit_result["auto_swap_recommended"]:
                st.info("💡 **Consejo:** Para corregir la inversión, simplemente carga el archivo correspondiente en cada casilla o utiliza el 'Buzón Inteligente'.")

        # Si hay advertencias
        if audit_result["warnings"]:
            for warn in audit_result["warnings"]:
                st.warning(warn)

        # Si hay confirmaciones
        if audit_result["info"] and not audit_result["errors"]:
            for inf in audit_result["info"]:
                st.success(inf)

    except Exception as e:
        st.error(f"Error al auditar los archivos cargados: {str(e)}")
        can_proceed_processing = False

st.markdown("<br>", unsafe_allow_html=True)

# Botón de acción principal (se inhabilita si falla la auditoría)
btn_procesar = st.button(
    "🚀 Procesar Cartera",
    disabled=not can_proceed_processing,
    help="El botón solo se activa si los archivos superan las reglas de no duplicidad, esquema y cronología.",
    use_container_width=True,
)

# ==============================================================================
# EJECUCIÓN DEL PIPELINE Y RESULTADOS
# ==============================================================================
if btn_procesar:
    with st.spinner("Ejecutando depuración de cartera y cruzando bases de datos..."):
        try:
            # 1. Leer DataFrames
            df_act = read_uploaded_file(file_actual)
            df_ant = read_uploaded_file(file_anterior) if file_anterior is not None else pd.DataFrame(columns=df_act.columns)
            df_rec = read_uploaded_file(file_recaudos) if file_recaudos is not None else pd.DataFrame(columns=["Póliza", "Nº ident.fis.1"])

            # 2. Ejecutar Pipeline (sincroniza automáticamente la lista No Contactar desde Google Sheets)
            resultados = pipeline.process(
                df_cartera_actual_raw=df_act,
                df_cartera_anterior_raw=df_ant,
                df_recaudos_raw=df_rec,
            )

            st.session_state["resultados_cartera"] = resultados
            st.toast("¡Cartera procesada con éxito!", icon="✅")

        except Exception as e:
            st.error(f"❌ Ocurrió un error inesperado al procesar los archivos: {str(e)}")
            st.exception(e)

# ==============================================================================
# VISUALIZACIÓN DE RESULTADOS
# ==============================================================================
if "resultados_cartera" in st.session_state:
    res = st.session_state["resultados_cartera"]
    dfs = res["dfs"]
    metrics = res["metrics"]
    nc_info = res["no_contactar_info"]

    st.markdown("<hr style='margin: 25px 0;'>", unsafe_allow_html=True)
    st.markdown("### 📊 Tablero de Resultados Operativos")

    # Alerta sobre la fuente de No Contactar
    if "Google Sheets" in nc_info["fuente"]:
        st.success(f"🌐 **Lista No Contactar:** Sincronizada directamente en vivo desde Google Sheets ({nc_info['total_excluidos']} clientes excluidos).")
    else:
        st.warning(f"⚠️ **Lista No Contactar:** {nc_info['estado']}")

    # Cuadrícula de Métricas
    m_col1, m_col2, m_col3, m_col4, m_col5, m_col6 = st.columns(6)

    with m_col1:
        st.markdown(
            f"""
            <div class="gb-metric-card" style="border-left-color: #3B82F6;">
                <div class="gb-metric-label">Base Depurada</div>
                <div class="gb-metric-num" style="color: #1E40AF;">{metrics['base_depurada']:,}</div>
                <div style="font-size:11px; color:#64748B;">Cobros regulares</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with m_col2:
        st.markdown(
            f"""
            <div class="gb-metric-card" style="border-left-color: #EF4444;">
                <div class="gb-metric-label">Alto Impacto</div>
                <div class="gb-metric-num" style="color: #B91C1C;">{metrics['alto_impacto']:,}</div>
                <div style="font-size:11px; color:#64748B;">Élite / &gt;$3M / ≥45d / Mensual&gt;$1M</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with m_col3:
        st.markdown(
            f"""
            <div class="gb-metric-card" style="border-left-color: #F59E0B;">
                <div class="gb-metric-label">Débitos Fallidos</div>
                <div class="gb-metric-num" style="color: #D97706;">{metrics['debitos_fallidos']:,}</div>
                <div style="font-size:11px; color:#64748B;">Mora &gt; 16 días</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with m_col4:
        st.markdown(
            f"""
            <div class="gb-metric-card" style="border-left-color: #8B5CF6;">
                <div class="gb-metric-label">Devoluciones</div>
                <div class="gb-metric-num" style="color: #6D28D9;">{metrics['devoluciones']:,}</div>
                <div style="font-size:11px; color:#64748B;">Trámites a favor</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with m_col5:
        st.markdown(
            f"""
            <div class="gb-metric-card" style="border-left-color: #EC4899;">
                <div class="gb-metric-label">Por Cancelar</div>
                <div class="gb-metric-num" style="color: #BE185D;">{metrics['revisar_cancelacion']:,}</div>
                <div style="font-size:11px; color:#64748B;">Sin actual ni recaudo</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with m_col6:
        st.markdown(
            f"""
            <div class="gb-metric-card" style="border-left-color: #64748B;">
                <div class="gb-metric-label">No Contactar</div>
                <div class="gb-metric-num" style="color: #334155;">{metrics['total_no_contactar']:,}</div>
                <div style="font-size:11px; color:#64748B;">Excluidos de gestión</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    exitosos_cnt = metrics.get('debitos_exitosos_excluidos', 0)
    st.caption(f"ℹ️ Se excluyeron **{metrics['prematuros_excluidos']}** débitos prematuros (≤ 16 días) y **{exitosos_cnt}** débitos exitosos (no fallaron), de un total de **{metrics['total_actual']}** registros iniciales.")

    st.markdown("<br>", unsafe_allow_html=True)

    # ==============================================================================
    # SECCIÓN 5: DESCARGA DEL EXCEL FORMATEADO
    # ==============================================================================
    st.markdown("### 📥 Descarga de Archivo Depurado")
    excel_bytes = export_cartera_to_excel(dfs)

    col_btn_1, col_btn_2 = st.columns([2, 1])
    with col_btn_1:
        st.download_button(
            label="📊 Descargar Cartera Depurada (.xlsx con Formato Condicional)",
            data=excel_bytes,
            file_name="Cartera_Depurada_GuiaBien.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_btn_2:
        st.markdown(
            """
            <div style="font-size:12px; color:#475569; padding-top:6px;">
                Incluye 5 hojas independientes y <strong>resaltado rojo automático</strong>
                (#FFC7CE) en cualquier fila marcada con <code>No contactar</code>.
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ==============================================================================
    # EXPLORADOR INTERACTIVO DE PESTAÑAS
    # ==============================================================================
    st.markdown("### 🔍 Explorador Interactivo de Hojas")

    tabs = st.tabs([
        f"Base depurada ({len(dfs['Base depurada'])})",
        f"Alto impacto ({len(dfs['Alto impacto'])})",
        f"Débitos fallidos ({len(dfs['Débitos fallidos'])})",
        f"Devoluciones ({len(dfs['Devoluciones'])})",
        f"Revisar cancelación ({len(dfs['Revisar cancelación'])})",
    ])

    # 1. Base Depurada
    with tabs[0]:
        st.caption("Registros estándar de cobro listos para asignación y gestión habitual.")
        st.dataframe(dfs["Base depurada"], use_container_width=True, height=350)

    # 2. Alto Impacto
    with tabs[1]:
        st.caption("Pólizas con alta prioridad comercial o riesgo financiero. Solo incluye registros con días de cartera estrictamente positivos (> 0): Clientes Élite, saldo > $3M, mora ≥ 45 días o periodicidad mensual con saldo > $1M.")
        df_ai = dfs["Alto impacto"]
        st.dataframe(df_ai, use_container_width=True, height=350)

        diag = res.get("diagnostico_alto_impacto", {})
        if diag:
            with st.expander("🔍 Ver Auditoría de Columnas y Detección de Alto Impacto"):
                c_d1, c_d2, c_d3 = st.columns(3)
                with c_d1:
                    st.markdown(f"**Periodicidad detectada:** `{diag.get('col_periodicidad', 'N/A')}`")
                    st.caption(f"Pólizas mensuales en archivo: **{diag.get('total_mensuales', 0):,}**")
                with c_d2:
                    st.markdown(f"**Saldo detectado:** `{diag.get('col_saldo', 'N/A')}`")
                    st.caption(f"Pólizas saldo > $1M: **{diag.get('total_saldo_gt_1m', 0):,}**")
                with c_d3:
                    st.markdown(f"**Días Cartera detectado:** `{diag.get('col_dias_cartera', 'N/A')}`")
                    st.caption(f"Pólizas con mora > 0: **{diag.get('total_dias_gt_0', 0):,}**")

        if not df_ai.empty and "Mensaje_Contacto" in df_ai.columns:
            debitos_en_ai = df_ai[df_ai["Mensaje_Contacto"].astype(str).str.strip() != ""]
            if not debitos_en_ai.empty:
                with st.expander(f"💬 Mensajes de Contacto para Débitos de Alto Impacto ({len(debitos_en_ai)})"):
                    st.caption("Mensajes listos para WhatsApp de pólizas prioritarias asociadas a débito:")
                    for i, (_, row) in enumerate(debitos_en_ai.head(10).iterrows(), start=1):
                        tomador = row.get("Nombre Tomador", "Cliente")
                        pol = row.get("Póliza", "")
                        msg = row.get("Mensaje_Contacto", "")
                        motivo = row.get("Motivo_Alto_Impacto", "")
                        st.markdown(f"**{i}. {tomador} (Póliza: {pol})** — *{motivo}*", unsafe_allow_html=True)
                        st.text_area(label=f"Mensaje Alto Impacto #{i}", value=msg, height=75, key=f"msg_ai_txt_{i}")

    # 3. Débitos Fallidos
    with tabs[2]:
        st.caption("Débitos automáticos rechazados con más de 16 días de cartera (excluye débitos exitosos). Incluye mensaje personalizado con el primer nombre.")
        df_deb = dfs["Débitos fallidos"]
        st.dataframe(df_deb, use_container_width=True, height=300)

        if not df_deb.empty and "Mensaje_Contacto" in df_deb.columns:
            with st.expander("💬 Visualizar y Copiar Mensajes de Contacto"):
                st.caption("Mensajes preformateados para envío por WhatsApp o SMS:")
                for i, (_, row) in enumerate(df_deb.head(10).iterrows(), start=1):
                    tomador = row.get("Nombre Tomador", "Cliente")
                    pol = row.get("Póliza", "")
                    msg = row.get("Mensaje_Contacto", "")
                    estado = row.get("Estado_Contacto", "")

                    badge = "<span class='gb-alert-badge'>NO CONTACTAR</span>" if estado == "No contactar" else ""
                    st.markdown(f"**{i}. {tomador} (Póliza: {pol})** {badge}", unsafe_allow_html=True)
                    st.text_area(label=f"Mensaje #{i}", value=msg, height=75, key=f"msg_txt_{i}")

    # 4. Devoluciones
    with tabs[3]:
        st.caption("Operaciones de devolución a favor de los tomadores o asegurados.")
        st.dataframe(dfs["Devoluciones"], use_container_width=True, height=350)

    # 5. Revisar Cancelación
    with tabs[4]:
        st.caption("Pólizas que aparecían activas en la cartera anterior, pero no figuran en la cartera actual ni en recaudos.")
        st.dataframe(dfs["Revisar cancelación"], use_container_width=True, height=350)
