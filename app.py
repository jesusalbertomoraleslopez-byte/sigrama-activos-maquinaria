"""
app.py
APP ACTIVOS - MAQUINARIA Y HERRAMIENTAS
Diseñado para SIGRAMA - Industria 4.0 & Sistemas ERP/MES
"""

import os
import streamlit as st
import pandas as pd
from datetime import date, datetime
import subprocess

import core_logic as core

# -----------------------------------------------------------------------------
# Configuración Inicial de la Página
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="APP ACTIVOS - MAQUINARIA Y HERRAMIENTAS | SIGRAMA",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inyección de estilos CSS Corporativos
def load_css():
    css_path = os.path.join(os.path.dirname(__file__), "styles.css")
    if os.path.exists(css_path):
        with open(css_path, "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

load_css()

# Inicialización de la base de datos Excel si no existe
core.init_excel_db()

# -----------------------------------------------------------------------------
# Sidebar: Logotipo e Identidad SIGRAMA + Menú Secuencial
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("""
        <div class="sidebar-brand-box">
            <h2>🏭 SIGRAMA</h2>
            <p>Manufactura & Industria 4.0</p>
            <div style="font-size: 0.68rem; color: #CBD5E1; margin-top: 6px;">
                Control de Activos y Herramental
            </div>
        </div>
    """, unsafe_allow_html=True)

    st.markdown("### 📋 Navegación")
    menu = st.radio(
        label="Seleccione un módulo:",
        options=[
            "📌 1. Inicio / Dashboard",
            "➕ 2. Registro de Activos",
            "🔍 3. Consulta y Ficha Técnica",
            "⚙️ 4. Configuración y Respaldos"
        ],
        index=0,
        label_visibility="collapsed"
    )

    st.markdown("---")
    
    # Resumen Rápido en Sidebar
    df_sidebar = core.init_excel_db()
    total_reg = len(df_sidebar)
    st.markdown(f"""
        <div style="background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 10px; font-size: 0.8rem;">
            <div style="font-weight: 700; color: #0B4F8A; margin-bottom: 4px;">ESTADO DEL SISTEMA</div>
            <div>Activos Registrados: <b>{total_reg}</b></div>
            <div>Persistencia: <b>Excel (.xlsx)</b></div>
            <div>Sede: <b>Planta Principal</b></div>
        </div>
    """, unsafe_allow_html=True)


# =============================================================================
# 📌 1. INICIO / DASHBOARD
# =============================================================================
if menu == "📌 1. Inicio / Dashboard":
    st.markdown("""
        <div style="margin-bottom: 1.2rem;">
            <h1 style="color: #0B4F8A; margin: 0; font-size: 2rem;">Dashboard General de Planta</h1>
            <p style="color: #64748B; margin: 0.2rem 0 0 0; font-size: 0.95rem;">
                Supervisión ejecutiva de maquinaria, herramentales y flujo de trabajo (WIP) - SIGRAMA
            </p>
        </div>
    """, unsafe_allow_html=True)

    df = core.init_excel_db()

    # Si está vacío, dar la opción de cargar datos demo con 1 clic
    if df.empty:
        st.info("ℹ️ El inventario se encuentra actualmente vacío. Puede registrar un nuevo activo o cargar datos demo de prueba.")
        if st.button("🚀 Cargar Activos de Demostración Industrial para SIGRAMA"):
            core.create_demo_assets_if_empty()
            st.success("Activos de demostración cargados exitosamente.")
            st.rerun()

    # --- KPIs Principales ---
    total_activos = len(df)
    total_inversion = df["Inversion_Total"].sum() if not df.empty and "Inversion_Total" in df.columns else 0.0
    operativos = len(df[df["Estatus_Operativo"] == "Operativo"]) if not df.empty else 0
    en_mant = len(df[df["Estatus_Operativo"] == "En Mantenimiento"]) if not df.empty else 0
    criticos = len(df[df["Estatus_Operativo"] == "Crítico"]) if not df.empty else 0

    pct_operativo = (operativos / total_activos * 100) if total_activos > 0 else 0

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)

    with kpi1:
        st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">Total de Activos</div>
                <div class="kpi-value">{total_activos}</div>
                <div class="kpi-subtitle">Equipos y herramentales</div>
            </div>
        """, unsafe_allow_html=True)

    with kpi2:
        st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">Valor del Inventario (MOI)</div>
                <div class="kpi-value">${total_inversion:,.0f}</div>
                <div class="kpi-subtitle">MXN Neto + Gastos</div>
            </div>
        """, unsafe_allow_html=True)

    with kpi3:
        st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">Disponibilidad Operativa</div>
                <div class="kpi-value">{pct_operativo:.1f}%</div>
                <div class="kpi-subtitle">{operativos} equipos activos</div>
            </div>
        """, unsafe_allow_html=True)

    with kpi4:
        st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-title">Mantenimiento / Crítico</div>
                <div class="kpi-value" style="color: {'#EF4444' if criticos > 0 else '#F59E0B'};">{en_mant + criticos}</div>
                <div class="kpi-subtitle">{criticos} con alerta crítica</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<div style='height: 1.5rem;'></div>", unsafe_allow_html=True)

    # --- MÓDULO VISUAL KANBAN ESTILO ODOO (WIP POR ÁREAS) ---
    st.markdown("### 🗂️ Flujo de Trabajo y Distribución por Áreas (Tablero Kanban WIP)")
    st.markdown("<p style='font-size: 0.85rem; color: #64748B;'>Visualice la distribución de maquinaria y herramental en cada estación del proceso productivo.</p>", unsafe_allow_html=True)

    # Filtros rápidos para el Kanban
    col_fil1, col_fil2 = st.columns([2, 2])
    with col_fil1:
        cat_filtro = st.selectbox("Filtrar Kanban por Categoría:", ["Todas"] + list(core.CATEGORIAS_DICT.keys()))
    with col_fil2:
        est_filtro = st.selectbox("Filtrar Kanban por Estatus:", ["Todos"] + core.ESTATUS_OPCIONES)

    df_kanban = df.copy()
    if cat_filtro != "Todas" and not df_kanban.empty:
        df_kanban = df_kanban[df_kanban["Categoria"] == cat_filtro]
    if est_filtro != "Todos" and not df_kanban.empty:
        df_kanban = df_kanban[df_kanban["Estatus_Operativo"] == est_filtro]

    # Renderizar columnas del Kanban
    kanban_cols = st.columns(len(core.AREAS_PRODUCCION))

    for idx, area in enumerate(core.AREAS_PRODUCCION):
        with kanban_cols[idx]:
            sub_df = df_kanban[df_kanban["Area_Produccion"] == area] if not df_kanban.empty else pd.DataFrame()
            count_area = len(sub_df)
            
            st.markdown(f"""
                <div style="background: #E2E8F0; padding: 8px 12px; border-radius: 8px 8px 0 0; border-top: 3px solid #0B4F8A; display: flex; justify-content: space-between; align-items: center;">
                    <span style="font-weight: 700; font-size: 0.9rem; color: #072A4A;">{area}</span>
                    <span style="background: #0B4F8A; color: white; border-radius: 10px; padding: 2px 7px; font-size: 0.75rem; font-weight: bold;">{count_area}</span>
                </div>
            """, unsafe_allow_html=True)

            if sub_df.empty:
                st.markdown("""
                    <div style="background: #F8FAFC; border: 1px dashed #CBD5E1; border-radius: 0 0 8px 8px; padding: 18px 10px; text-align: center; color: #94A3B8; font-size: 0.78rem;">
                        Sin equipos asignados
                    </div>
                """, unsafe_allow_html=True)
            else:
                for _, row in sub_df.iterrows():
                    estatus = str(row.get("Estatus_Operativo", "Operativo"))
                    badge_class = "status-operativo" if estatus == "Operativo" else ("status-mantenimiento" if estatus == "En Mantenimiento" else "status-critico")
                    
                    st.markdown(f"""
                        <div class="kanban-card">
                            <div style="display: flex; justify-content: space-between; align-items: baseline;">
                                <span class="kanban-card-id">{row.get('ID_Activo')}</span>
                                <span class="status-pill {badge_class}">{estatus}</span>
                            </div>
                            <div class="kanban-card-title">{row.get('Nombre_Equipo')}</div>
                            <div style="font-size: 0.75rem; color: #64748B;">{row.get('Marca', '')} {row.get('Modelo', '')}</div>
                            <div class="kanban-card-meta">
                                <span>👤 {str(row.get('Responsable', 'N/A'))[:15]}</span>
                                <span>💵 ${float(row.get('MOI_Neto', 0) or 0)/1000:,.0f}k</span>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)


# =============================================================================
# ➕ 2. REGISTRO DE ACTIVOS (FORMULARIO CONTINUO CON VISTA PRELIMINAR EN VIVO)
# =============================================================================
elif menu == "➕ 2. Registro de Activos":
    st.markdown("""
        <div style="margin-bottom: 1.2rem;">
            <h1 style="color: #0B4F8A; margin: 0; font-size: 2rem;">Alta y Registro de Activos en Planta</h1>
            <p style="color: #64748B; margin: 0.2rem 0 0 0; font-size: 0.95rem;">
                Capture la información en una sola hoja continua. Visualice en la parte superior la fotografía del equipo y a la derecha la vista preliminar con su código QR en tiempo real.
            </p>
        </div>
    """, unsafe_allow_html=True)

    df_current = core.init_excel_db()

    # Layout de 2 Columnas: Izquierda (Formulario continuo con scroll) | Derecha (Vista Preliminar Fija)
    col_form, col_preview = st.columns([1.65, 1.15], gap="large")

    # ==================== COLUMNA IZQUIERDA: FORMULARIO CONTINUO ====================
    with col_form:
        # --- PARTE SUPERIOR: CARGA DE LA IMAGEN DEL EQUIPO ---
        st.markdown("""
            <div class="form-block-card" style="border-left: 5px solid #0B4F8A;">
                <div class="form-block-title">
                    📸 1. EVIDENCIA FOTOGRÁFICA DEL EQUIPO (REFERENCIA VISUAL)
                </div>
                <p style="font-size: 0.82rem; color: #64748B; margin-top: -8px; margin-bottom: 10px;">
                    Cargue la fotografía del activo para no perderla de vista durante el llenado de datos técnicos y fiscales.
                </p>
            </div>
        """, unsafe_allow_html=True)

        foto_archivo = st.file_uploader(
            "Seleccionar fotografía del activo (JPG, JPEG, PNG)",
            type=["jpg", "jpeg", "png"],
            key="foto_uploader_single",
            help="Suba una fotografía clara del equipo, placa del fabricante o herramental."
        )

        if foto_archivo is not None:
            col_img1, col_img2 = st.columns([1.2, 1])
            with col_img1:
                st.image(foto_archivo, caption=f"Fotografía cargada: {foto_archivo.name}", use_container_width=True)
            with col_img2:
                st.success("✅ Imagen cargada y lista para vincular al activo.")
                st.info("💡 La imagen se encuentra visible en el panel de vista preliminar a la derecha.")

        st.markdown("<div style='height: 0.8rem;'></div>", unsafe_allow_html=True)

        # --- SECCIÓN A: IDENTIFICACIÓN FÍSICA Y MAQUINARIA ---
        st.markdown("""
            <div class="form-block-card">
                <div class="form-block-title">
                    🏷️ 2. DATOS DE IDENTIFICACIÓN FÍSICA Y TÉCNICA
                </div>
            </div>
        """, unsafe_allow_html=True)

        col_a1, col_a2 = st.columns(2)
        with col_a1:
            nombre_equipo = st.text_input(
                "Nombre del Equipo / Máquina *",
                placeholder="Ej. Centro de Corte Láser Fibra 6kW",
                key="reg_nombre"
            )
            marca = st.text_input("Marca del Fabricante *", placeholder="Ej. Bystronic, Trumpf, Haas, Amada", key="reg_marca")
            modelo = st.text_input("Modelo *", placeholder="Ej. ByStar Fiber 3015", key="reg_modelo")

        with col_a2:
            numero_serie = st.text_input("Número de Serie *", placeholder="Ej. BY-60291-MX", key="reg_serie")
            categoria = st.selectbox(
                "Categoría Principal *",
                list(core.CATEGORIAS_DICT.keys()),
                key="reg_categoria"
            )
            subcategorias_disp = core.CATEGORIAS_DICT.get(categoria, [])
            subcategoria = st.selectbox(
                "Subcategoría Dinámica *",
                subcategorias_disp,
                key="reg_subcategoria"
            )

        st.markdown("<div style='height: 0.8rem;'></div>", unsafe_allow_html=True)

        # --- SECCIÓN B: CUMPLIMIENTO FISCAL Y LEGAL (SAT) ---
        st.markdown("""
            <div class="form-block-card">
                <div class="form-block-title">
                    ⚖️ 3. CUMPLIMIENTO FISCAL Y AUDITORÍAS (SAT MÉXICO)
                </div>
            </div>
        """, unsafe_allow_html=True)

        col_b1, col_b2 = st.columns(2)
        with col_b1:
            uuid_cfdi = st.text_input(
                "UUID CFDI (Folio Fiscal de 36 caracteres) *",
                placeholder="Ej. 4A1E8D23-67BC-44F1-92EA-1A8B3567D001",
                key="reg_uuid",
                help="Folio fiscal alfanumérico emitido por el SAT en la factura digital mexicana."
            )
            rfc_proveedor = st.text_input(
                "RFC del Proveedor *",
                placeholder="Ej. BYM980315LK2",
                max_chars=13,
                key="reg_rfc"
            )
        with col_b2:
            uso_cfdi = st.selectbox(
                "Uso de CFDI (Catálogo Oficial SAT) *",
                [
                    "I02 Maquinaria y equipo",
                    "I05 Dados, troqueles, moldes, matrices y herramental",
                    "I04 Equipo de cómputo y accesorios",
                    "I08 Otra maquinaria y equipo",
                    "G03 Gastos en general"
                ],
                key="reg_uso_cfdi"
            )
            es_importado = st.checkbox(
                "¿Es equipo de procedencia extranjera / importado?",
                value=False,
                key="reg_importado"
            )
            numero_pedimento = st.text_input(
                "Número de Pedimento Aduanal",
                placeholder="Obligatorio si es importado (ej. 23 16 3840 7001923)",
                disabled=not es_importado,
                key="reg_pedimento"
            )

        st.markdown("<div style='height: 0.8rem;'></div>", unsafe_allow_html=True)

        # --- SECCIÓN C: INFORMACIÓN FINANCIERA E INVERSIÓN (LISR) ---
        st.markdown("""
            <div class="form-block-card">
                <div class="form-block-title">
                    💰 4. INFORMACIÓN FINANCIERA E INVERSIÓN (ART. 31-38 LISR)
                </div>
            </div>
        """, unsafe_allow_html=True)

        col_c1, col_c2 = st.columns(2)
        with col_c1:
            moi_neto = st.number_input(
                "MOI Neto (Monto Original de Inversión sin IVA) [MXN] *",
                min_value=0.0,
                value=150000.0,
                step=10000.0,
                format="%.2f",
                key="reg_moi"
            )
            gastos_inherentes = st.number_input(
                "Gastos Inherentes / Instalación (fletes, seguros, cimentación) [MXN]",
                min_value=0.0,
                value=12000.0,
                step=5000.0,
                format="%.2f",
                key="reg_gastos"
            )
            inversion_calc = float(moi_neto) + float(gastos_inherentes)
            st.info(f"💵 **Inversión Total Capitalizable:** ${inversion_calc:,.2f} MXN")

        with col_c2:
            tasa_default = 35.0 if "Herramental" in categoria else (25.0 if "Equipo Móvil" in categoria else 10.0)
            tasa_depreciacion = st.number_input(
                "Tasa de Depreciación Anual (%) [LISR] *",
                min_value=1.0,
                max_value=100.0,
                value=tasa_default,
                step=1.0,
                key="reg_tasa",
                help="Porcentaje de depreciación fiscal: 10% Maquinaria, 35% Troqueles/Herramentales, 25% Equipo de Transporte."
            )
            fecha_adquisicion = st.date_input("Fecha de Adquisición *", value=date.today(), key="reg_f_adq")
            fecha_inicio_uso = st.date_input("Fecha de Inicio de Uso *", value=date.today(), key="reg_f_uso")

        st.markdown("<div style='height: 0.8rem;'></div>", unsafe_allow_html=True)

        # --- SECCIÓN D: ASIGNACIÓN Y CUSTODIA OPERATIVA ---
        st.markdown("""
            <div class="form-block-card">
                <div class="form-block-title">
                    📍 5. ASIGNACIÓN OPERATIVA Y CUSTODIA EN PLANTA
                </div>
            </div>
        """, unsafe_allow_html=True)

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            area_produccion = st.selectbox("Área de Producción *", core.AREAS_PRODUCCION, key="reg_area")
            responsable = st.text_input("Responsable / Custodio del Activo *", placeholder="Ej. Ing. Carlos Mendoza", key="reg_resp")
        with col_d2:
            estatus_operativo = st.selectbox("Estatus Operativo Inicial *", core.ESTATUS_OPCIONES, key="reg_estatus")

        st.markdown("<div style='height: 1.2rem;'></div>", unsafe_allow_html=True)

        # --- BOTÓN DE GUARDADO ---
        guardar_btn = st.button("💾 Guardar y Registrar Activo en Inventario", use_container_width=True)

        if guardar_btn:
            errores = []
            if not nombre_equipo.strip():
                errores.append("El Nombre del Equipo es obligatorio.")
            if not marca.strip() or not modelo.strip():
                errores.append("La Marca y Modelo son obligatorios.")
            if not numero_serie.strip():
                errores.append("El Número de Serie es obligatorio.")
            if not uuid_cfdi.strip() or len(uuid_cfdi.strip()) < 32:
                errores.append("El UUID CFDI debe ser un folio fiscal válido (32 a 36 caracteres).")
            if not rfc_proveedor.strip() or len(rfc_proveedor.strip()) < 12:
                errores.append("El RFC del Proveedor debe tener entre 12 y 13 caracteres.")
            if es_importado and not numero_pedimento.strip():
                errores.append("Para equipos importados, el Número de Pedimento es obligatorio ante el SAT.")
            if not responsable.strip():
                errores.append("Debe especificar el Responsable del Activo.")

            if errores:
                for err in errores:
                    st.error(f"⚠️ {err}")
            else:
                with st.spinner("Registrando activo, guardando fotografía física y generando código QR..."):
                    datos_activo = {
                        "Nombre_Equipo": nombre_equipo.strip(),
                        "Marca": marca.strip(),
                        "Modelo": modelo.strip(),
                        "Numero_Serie": numero_serie.strip(),
                        "Categoria": categoria,
                        "Subcategoria": subcategoria,
                        "UUID_CFDI": uuid_cfdi.strip().upper(),
                        "RFC_Proveedor": rfc_proveedor.strip().upper(),
                        "Uso_CFDI": uso_cfdi,
                        "Es_Importado": es_importado,
                        "Numero_Pedimento": numero_pedimento.strip() if es_importado else "",
                        "MOI_Neto": float(moi_neto),
                        "Gastos_Inherentes": float(gastos_inherentes),
                        "Tasa_Depreciacion_Anual": float(tasa_depreciacion),
                        "Fecha_Adquisicion": str(fecha_adquisicion),
                        "Fecha_Inicio_Uso": str(fecha_inicio_uso),
                        "Area_Produccion": area_produccion,
                        "Responsable": responsable.strip(),
                        "Estatus_Operativo": estatus_operativo,
                    }

                    exito, nuevo_id, res = core.register_new_asset(datos_activo, foto_archivo)

                    if exito:
                        st.balloons()
                        st.success(f"🎉 ¡Activo registrado exitosamente con clave: **{nuevo_id}**!")
                        st.info(f"📁 Se almacenó en `inventario_activos.xlsx` con su foto física en `media/fotos_activos/` y código QR en `media/qrs/`.")
                    else:
                        st.error(f"Error al guardar: {nuevo_id}")


    # ==================== COLUMNA DERECHA: VISTA PRELIMINAR EN TIEMPO REAL ====================
    with col_preview:
        # Calcular clave proyectada en vivo
        clave_proyectada = core.generate_unique_id(df_current, categoria, subcategoria)

        st.markdown(f"""
            <div class="live-preview-container">
                <div class="live-preview-header">
                    <span>📋 VISTA PRELIMINAR EN VIVO</span>
                    <span class="live-preview-pulse" title="Sincronizado en tiempo real"></span>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # 1. Foto en Tiempo Real
        if foto_archivo is not None:
            st.image(foto_archivo, caption="Fotografía del Equipo (Cargada)", use_container_width=True)
        else:
            st.markdown("""
                <div style="background: #0F172A; border: 2px dashed #38BDF8; border-radius: 10px; padding: 35px 15px; text-align: center; color: #94A3B8; margin-bottom: 12px;">
                    <div style="font-size: 2.2rem; margin-bottom: 6px;">📷</div>
                    <div style="font-weight: 700; color: #F8FAFC; font-size: 0.9rem;">Sin Fotografía Adjunta</div>
                    <div style="font-size: 0.75rem; color: #64748B;">Cargue una foto en la parte superior para verla aquí.</div>
                </div>
            """, unsafe_allow_html=True)

        # 2. Tarjeta Resumen con Datos Dinámicos
        nombre_display = nombre_equipo.strip() if nombre_equipo.strip() else "Nombre del Equipo pendiente..."
        marca_modelo = f"{marca.strip()} {modelo.strip()}".strip()
        if not marca_modelo:
            marca_modelo = "Marca / Modelo pendiente"

        serie_display = numero_serie.strip() if numero_serie.strip() else "SN: Pendiente"
        
        status_color = "#10B981" if estatus_operativo == "Operativo" else ("#F59E0B" if estatus_operativo == "En Mantenimiento" else "#EF4444")
        origen_label = "IMPORTADO (CON PEDIMENTO)" if es_importado else "NACIONAL"

        st.markdown(f"""
            <div style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 14px; margin-bottom: 12px;">
                <div style="display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 6px;">
                    <span style="font-weight: 800; color: #0B4F8A; font-size: 0.95rem;">{clave_proyectada}</span>
                    <span style="background: {status_color}; color: white; padding: 2px 8px; border-radius: 10px; font-size: 0.72rem; font-weight: bold;">
                        {estatus_operativo.upper()}
                    </span>
                </div>
                <div style="font-weight: 700; font-size: 1.05rem; color: #1E293B; line-height: 1.25; margin-bottom: 4px;">
                    {nombre_display}
                </div>
                <div style="font-size: 0.8rem; color: #64748B; margin-bottom: 10px;">
                    {marca_modelo} • <code>{serie_display}</code>
                </div>
                <div style="border-top: 1px solid #E2E8F0; padding-top: 8px; font-size: 0.8rem; display: flex; justify-content: space-between;">
                    <span style="color: #64748B;">Área Asignada:</span>
                    <span style="font-weight: 700; color: #072A4A;">{area_produccion}</span>
                </div>
                <div style="font-size: 0.8rem; display: flex; justify-content: space-between; margin-top: 4px;">
                    <span style="color: #64748B;">Custodio:</span>
                    <span style="font-weight: 600; color: #1E293B;">{responsable if responsable.strip() else 'No asignado'}</span>
                </div>
                <div style="font-size: 0.8rem; display: flex; justify-content: space-between; margin-top: 4px;">
                    <span style="color: #64748B;">Inversión Total:</span>
                    <span style="font-weight: 800; color: #0B4F8A;">${inversion_calc:,.2f} MXN</span>
                </div>
                <div style="font-size: 0.8rem; display: flex; justify-content: space-between; margin-top: 4px;">
                    <span style="color: #64748B;">Depreciación LISR:</span>
                    <span style="font-weight: 600; color: #047857;">{tasa_depreciacion}% anual</span>
                </div>
                <div style="font-size: 0.75rem; display: flex; justify-content: space-between; margin-top: 4px; color: #64748B;">
                    <span>Origen Fiscal:</span>
                    <span><b>{origen_label}</b></span>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # 3. Código QR en Tiempo Real
        st.markdown("<div style='font-size: 0.82rem; font-weight: 700; color: #0B4F8A; margin-bottom: 4px;'>CÓDIGO QR PROYECTADO:</div>", unsafe_allow_html=True)
        try:
            extra_qr = {
                "Nombre_Equipo": nombre_display,
                "Numero_Serie": serie_display,
                "Area_Produccion": area_produccion
            }
            qr_buffer = core.generate_qr_image_bytes(clave_proyectada, extra_qr)
            col_q1, col_q2 = st.columns([1, 1.2])
            with col_q1:
                st.image(qr_buffer, width=130)
            with col_q2:
                st.markdown(f"""
                    <div style="font-size: 0.76rem; color: #64748B; margin-top: 10px; line-height: 1.3;">
                        <b>Etiqueta SIGRAMA</b><br/>
                        ID: <code>{clave_proyectada}</code><br/>
                        Generada al vuelo para verificación en planta.
                    </div>
                """, unsafe_allow_html=True)
        except Exception as e:
            st.caption(f"QR en preparación: {e}")


# =============================================================================
# 🔍 3. CONSULTA Y FICHA TÉCNICA
# =============================================================================
elif menu == "🔍 3. Consulta y Ficha Técnica":
    st.markdown("""
        <div style="margin-bottom: 1.2rem;">
            <h1 style="color: #0B4F8A; margin: 0; font-size: 2rem;">Buscador y Ficha Técnica Ejecutiva</h1>
            <p style="color: #64748B; margin: 0.2rem 0 0 0; font-size: 0.95rem;">
                Consulte las especificaciones completas de cualquier activo, inspeccione evidencia fotográfica y descargue el reporte en PDF.
            </p>
        </div>
    """, unsafe_allow_html=True)

    df = core.init_excel_db()

    if df.empty:
        st.warning("⚠️ No hay activos registrados aún en la base de datos.")
    else:
        # Filtros Avanzados de Búsqueda
        col_s1, col_s2, col_s3, col_s4 = st.columns([2, 1.5, 1.5, 1.5])
        with col_s1:
            query = st.text_input("🔍 Buscar por ID, Nombre o Serie:", placeholder="Ej. CTR-LAS, Bystronic, BY-60291")
        with col_s2:
            filtro_cat = st.selectbox("Categoría:", ["Todas"] + sorted(list(df["Categoria"].dropna().unique())))
        with col_s3:
            filtro_area = st.selectbox("Área:", ["Todas"] + sorted(list(df["Area_Produccion"].dropna().unique())))
        with col_s4:
            filtro_est = st.selectbox("Estatus:", ["Todos"] + sorted(list(df["Estatus_Operativo"].dropna().unique())))

        # Aplicación de Filtros
        df_filtered = df.copy()
        if query.strip():
            q = query.strip().lower()
            df_filtered = df_filtered[
                df_filtered["ID_Activo"].astype(str).str.lower().str.contains(q) |
                df_filtered["Nombre_Equipo"].astype(str).str.lower().str.contains(q) |
                df_filtered["Marca"].astype(str).str.lower().str.contains(q) |
                df_filtered["Numero_Serie"].astype(str).str.lower().str.contains(q)
            ]
        if filtro_cat != "Todas":
            df_filtered = df_filtered[df_filtered["Categoria"] == filtro_cat]
        if filtro_area != "Todas":
            df_filtered = df_filtered[df_filtered["Area_Produccion"] == filtro_area]
        if filtro_est != "Todos":
            df_filtered = df_filtered[df_filtered["Estatus_Operativo"] == filtro_est]

        st.markdown(f"**Resultados encontrados:** {len(df_filtered)} activo(s)")

        # Tabla Resumen de Activos Encontrados
        st.dataframe(
            df_filtered[["ID_Activo", "Nombre_Equipo", "Marca", "Modelo", "Area_Produccion", "Estatus_Operativo", "Inversion_Total"]],
            use_container_width=True,
            hide_index=True
        )

        st.markdown("---")

        # Selector de Ficha Técnica Individual
        if not df_filtered.empty:
            opciones_activos = df_filtered["ID_Activo"].tolist()
            formato_mostrar = {row["ID_Activo"]: f"{row['ID_Activo']} - {row['Nombre_Equipo']}" for _, row in df_filtered.iterrows()}
            
            seleccionado = st.selectbox(
                "Seleccione un activo para desplegar la Ficha Técnica Ejecutiva:",
                options=opciones_activos,
                format_func=lambda x: formato_mostrar.get(x, x)
            )

            activo_row = df_filtered[df_filtered["ID_Activo"] == seleccionado].iloc[0].to_dict()

            # Despliegue de Ficha Técnica Ejecutiva (2 Columnas)
            st.markdown(f"""
                <div style="background: linear-gradient(135deg, #0B4F8A 0%, #072A4A 100%); color: white; padding: 12px 20px; border-radius: 8px; margin: 15px 0;">
                    <div style="font-size: 1.3rem; font-weight: 800;">{activo_row.get('Nombre_Equipo')}</div>
                    <div style="font-size: 0.85rem; color: #93C5FD;">ID: {activo_row.get('ID_Activo')} | Categoría: {activo_row.get('Categoria')} ({activo_row.get('Subcategoria')})</div>
                </div>
            """, unsafe_allow_html=True)

            col_ficha_izq, col_ficha_der = st.columns([1.6, 1.2])

            with col_ficha_izq:
                # 1. Datos Físicos
                st.markdown("<div class='section-badge'>DATOS FÍSICOS Y TÉCNICOS</div>", unsafe_allow_html=True)
                c_f1, c_f2 = st.columns(2)
                with c_f1:
                    st.write(f"**Marca:** {activo_row.get('Marca', 'N/A')}")
                    st.write(f"**Modelo:** {activo_row.get('Modelo', 'N/A')}")
                    st.write(f"**Número de Serie:** `{activo_row.get('Numero_Serie', 'N/A')}`")
                with c_f2:
                    st.write(f"**Área de Producción:** {activo_row.get('Area_Produccion', 'N/A')}")
                    st.write(f"**Responsable:** {activo_row.get('Responsable', 'N/A')}")
                    est = activo_row.get('Estatus_Operativo', 'Operativo')
                    st.write(f"**Estatus:** :{ 'green' if est=='Operativo' else ('orange' if est=='En Mantenimiento' else 'red') }[{est}]")

                # 2. Cumplimiento Fiscal SAT
                st.markdown("<div class='section-badge'>CUMPLIMIENTO FISCAL (SAT MÉXICO)</div>", unsafe_allow_html=True)
                st.write(f"**UUID CFDI:** `{activo_row.get('UUID_CFDI', 'N/A')}`")
                st.write(f"**RFC Proveedor:** `{activo_row.get('RFC_Proveedor', 'N/A')}`")
                st.write(f"**Uso de CFDI:** {activo_row.get('Uso_CFDI', 'N/A')}")
                if activo_row.get('Es_Importado'):
                    st.write(f"**Pedimento Aduanal:** `{activo_row.get('Numero_Pedimento', 'N/A')}` (Equipo Importado)")
                else:
                    st.write("**Origen:** Nacional")

                # 3. Datos Financieros
                st.markdown("<div class='section-badge'>INFORMACIÓN FINANCIERA (LISR)</div>", unsafe_allow_html=True)
                c_fin1, c_fin2 = st.columns(2)
                with c_fin1:
                    st.write(f"**MOI Neto:** ${float(activo_row.get('MOI_Neto', 0) or 0):,.2f} MXN")
                    st.write(f"**Gastos Inherentes:** ${float(activo_row.get('Gastos_Inherentes', 0) or 0):,.2f} MXN")
                    st.write(f"**Inversión Total:** **${float(activo_row.get('Inversion_Total', 0) or 0):,.2f} MXN**")
                with c_fin2:
                    st.write(f"**Tasa Depreciación:** {activo_row.get('Tasa_Depreciacion_Anual', 10)}% anual")
                    st.write(f"**Fecha Adquisición:** {activo_row.get('Fecha_Adquisicion', 'N/A')}")
                    st.write(f"**Fecha Inicio Uso:** {activo_row.get('Fecha_Inicio_Uso', 'N/A')}")

            with col_ficha_der:
                st.markdown("<div class='section-badge'>EVIDENCIA FOTOGRÁFICA Y CÓDIGO QR</div>", unsafe_allow_html=True)
                
                # Fotografía
                foto_path = str(activo_row.get("Ruta_Foto", ""))
                if foto_path and os.path.exists(foto_path):
                    st.image(foto_path, caption=f"Fotografía: {activo_row.get('Nombre_Equipo')}", use_container_width=True)
                else:
                    st.info("📷 Sin fotografía adjunta para este activo.")

                # QR Code
                qr_path = str(activo_row.get("Ruta_QR", ""))
                if qr_path and os.path.exists(qr_path):
                    col_q1, col_q2 = st.columns([1, 1.2])
                    with col_q1:
                        st.image(qr_path, caption=f"QR: {activo_row.get('ID_Activo')}", width=140)
                    with col_q2:
                        st.markdown(f"""
                            <div style="font-size: 0.8rem; color: #64748B; margin-top: 10px;">
                                <b>Etiqueta Oficial SIGRAMA</b><br/>
                                Escanee para verificación física en línea de ensamble o auditoría de planta.
                            </div>
                        """, unsafe_allow_html=True)

                st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)

                # Generar y Descargar PDF
                try:
                    pdf_buffer = core.generate_asset_pdf(activo_row)
                    st.download_button(
                        label="📥 Descargar Ficha Técnica en PDF",
                        data=pdf_buffer,
                        file_name=f"Ficha_Tecnica_{activo_row.get('ID_Activo')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error al preparar PDF: {e}")


# =============================================================================
# ⚙️ 4. CONFIGURACIÓN Y RESPALDOS
# =============================================================================
elif menu == "⚙️ 4. Configuración y Respaldos":
    st.markdown("""
        <div style="margin-bottom: 1.2rem;">
            <h1 style="color: #0B4F8A; margin: 0; font-size: 2rem;">Configuración, Respaldos y Git</h1>
            <p style="color: #64748B; margin: 0.2rem 0 0 0; font-size: 0.95rem;">
                Administración del repositorio de datos en Excel y sincronización segura con GitHub.
            </p>
        </div>
    """, unsafe_allow_html=True)

    col_conf1, col_conf2 = st.columns(2)

    with col_conf1:
        st.markdown("<div class='section-badge'>ADMINISTRACIÓN DE ARCHIVO EXCEL (.XLSX)</div>", unsafe_allow_html=True)
        st.markdown("""
            El sistema almacena todos los registros fiscales y técnicos en el archivo local `inventario_activos.xlsx`.
        """)

        # Descarga de Excel
        if os.path.exists(core.EXCEL_FILE):
            with open(core.EXCEL_FILE, "rb") as f:
                bytes_excel = f.read()
            st.download_button(
                label="📥 Descargar Archivo Excel Actual",
                data=bytes_excel,
                file_name=f"inventario_activos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.warning("El archivo Excel aún no ha sido inicializado.")

        st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)

        # Carga / Restauración de Excel
        st.markdown("#### Restaurar o Importar Excel")
        archivo_excel_subido = st.file_uploader(
            "Cargar archivo Excel (.xlsx) para reemplazar o actualizar:",
            type=["xlsx"],
            key="excel_restore_uploader"
        )
        if archivo_excel_subido is not None:
            if st.button("⚠️ Confirmar Sobreescritura de Inventario"):
                try:
                    df_nuevo = pd.read_excel(archivo_excel_subido, engine="openpyxl")
                    core.save_inventory(df_nuevo)
                    st.success("✅ Base de datos restaurada correctamente a partir del archivo subido.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al leer el archivo Excel: {e}")

        st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)
        st.markdown("#### Datos de Demostración")
        if st.button("🔄 Generar / Recargar Datos de Demostración Industrial"):
            core.create_demo_assets_if_empty()
            st.success("Datos demo verificados y cargados.")
            st.rerun()

    with col_conf2:
        st.markdown("<div class='section-badge'>SINCRONIZACIÓN Y CONTROL DE VERSIONES (GIT / GITHUB)</div>", unsafe_allow_html=True)
        st.markdown("""
            Para mantener el código y la arquitectura respaldados de manera segura en un repositorio de GitHub,
            ejecute los siguientes pasos en su terminal:
        """)

        st.code("""# 1. Inicializar repositorio local
git init

# 2. Agregar archivos respetando .gitignore
git add .

# 3. Realizar primer commit
git commit -m "feat: APP ACTIVOS MAQUINARIA Y HERRAMIENTAS - SIGRAMA v1.0"

# 4. Establecer rama principal
git branch -M main

# 5. Vincular repositorio remoto en GitHub (reemplace con su URL)
git remote add origin https://github.com/TU-USUARIO/APP-ACTIVOS-SIGRAMA.git

# 6. Subir cambios a GitHub
git push -u origin main
        """, language="bash")

        # Comprobación de estado Git en el directorio local
        st.markdown("#### Estado Local de Git:")
        try:
            res_git = subprocess.run(["git", "status"], capture_output=True, text=True, cwd=os.path.dirname(__file__))
            if res_git.returncode == 0:
                st.success("Repositorio Git inicializado localmente.")
                with st.expander("Ver salida de 'git status'"):
                    st.text(res_git.stdout)
            else:
                st.info("Directorio aún no inicializado como repositorio Git. Puede ejecutar `git init` en la consola.")
        except Exception:
            st.info("Git no detectado o no disponible en el PATH del sistema.")
