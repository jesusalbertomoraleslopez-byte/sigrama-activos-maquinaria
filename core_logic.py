"""
core_logic.py
Módulo central de lógica de negocio, persistencia en Excel,
procesamiento de imágenes, códigos QR y generación de reportes PDF para SIGRAMA.
"""

import os
import json
import re
from datetime import datetime
from io import BytesIO
from typing import Optional, Tuple, Dict, Any

import pandas as pd
from PIL import Image
import qrcode
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, KeepTogether, HRFlowable
)

# Constantes de Rutas
EXCEL_FILE = "inventario_activos.xlsx"
MEDIA_DIR = "media"
FOTOS_DIR = os.path.join(MEDIA_DIR, "fotos_activos")
QRS_DIR = os.path.join(MEDIA_DIR, "qrs")

# Esquema de Columnas del Inventario
COLUMNS = [
    "ID_Activo",
    "Nombre_Equipo",
    "Marca",
    "Modelo",
    "Numero_Serie",
    "Categoria",
    "Subcategoria",
    "UUID_CFDI",
    "RFC_Proveedor",
    "Uso_CFDI",
    "Es_Importado",
    "Numero_Pedimento",
    "MOI_Neto",
    "Gastos_Inherentes",
    "Inversion_Total",
    "Tasa_Depreciacion_Anual",
    "Fecha_Adquisicion",
    "Fecha_Inicio_Uso",
    "Area_Produccion",
    "Responsable",
    "Estatus_Operativo",
    "Ruta_Foto",
    "Ruta_QR",
    "Fecha_Registro"
]

# Mapa de Prefijos según Categoría y Subcategoría
PREFIX_MAP = {
    ("Maquinaria Principal", "Corte Láser"): "CTR-LAS",
    ("Maquinaria Principal", "Dobladora"): "DBL-PRE",
    ("Maquinaria Principal", "Lijadora/Rebabeadora"): "LIJ-REB",
    ("Maquinaria Principal", "*"): "MAQ-PRN",
    ("Línea de Pintura", "Pintura Batch"): "PNT-BAT",
    ("Línea de Pintura", "Pintura Continua"): "PNT-CNT",
    ("Línea de Pintura", "*"): "PNT-LIN",
    ("Equipo Móvil", "Montacargas"): "EQM-MTC",
    ("Equipo Móvil", "Patín Hidráulico"): "EQM-PTH",
    ("Equipo Móvil", "*"): "EQM-MOV",
    ("Estaciones/Mesas", "Mesa"): "EST-MES",
    ("Estaciones/Mesas", "*"): "EST-EST",
    ("Herramental de Doblez", "Dado/Punzón"): "HRR-PUN",
    ("Herramental de Doblez", "*"): "HRR-DBL",
    ("Herramienta de Mano", "*"): "HRR-MAN",
}

# Opciones de Categorías y Subcategorías Dinámicas
CATEGORIAS_DICT = {
    "Maquinaria Principal": [
        "Corte Láser",
        "Dobladora",
        "Lijadora/Rebabeadora",
        "Punzonadora CNC",
        "Torno / Fresa",
        "Otra Maquinaria Principal"
    ],
    "Línea de Pintura": [
        "Pintura Batch",
        "Pintura Continua",
        "Cabina de Aplicación",
        "Horno de Curado",
        "Línea de Lavado / Fosfatizado",
        "Otra Línea de Pintura"
    ],
    "Equipo Móvil": [
        "Montacargas",
        "Patín Hidráulico",
        "Grúa Viajera / Polipasto",
        "Carro de Arrastre",
        "Otro Equipo Móvil"
    ],
    "Estaciones/Mesas": [
        "Mesa",
        "Estación de Soldadura",
        "Mesa de Armado",
        "Mesa de Inspección QC",
        "Otra Estación"
    ],
    "Herramental de Doblez": [
        "Dado/Punzón",
        "Matriz Multivía",
        "Adaptador de Doblez",
        "Herramental Especial",
        "Otro Herramental"
    ],
    "Herramienta de Mano": [
        "Esmeriladora Angular",
        "Taladro / Rotomartillo",
        "Remachadora Neumática",
        "Pistola de Torque",
        "Otra Herramienta"
    ]
}

AREAS_PRODUCCION = ["Corte", "Doblez", "Pintura", "Ensamblado", "Logística"]
ESTATUS_OPCIONES = ["Operativo", "En Mantenimiento", "Crítico"]


def ensure_directories():
    """Asegura la existencia de directorios de almacenamiento local."""
    os.makedirs(FOTOS_DIR, exist_ok=True)
    os.makedirs(QRS_DIR, exist_ok=True)


def init_excel_db() -> pd.DataFrame:
    """Inicializa o carga el archivo Excel con las columnas requeridas."""
    ensure_directories()
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")
        return df
    try:
        df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
        # Asegurar todas las columnas
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = ""
        # Asegurar tipo texto para columnas descriptivas y de rutas
        text_cols = [
            "ID_Activo", "Nombre_Equipo", "Marca", "Modelo", "Numero_Serie",
            "Categoria", "Subcategoria", "UUID_CFDI", "RFC_Proveedor",
            "Uso_CFDI", "Numero_Pedimento", "Area_Produccion", "Responsable",
            "Estatus_Operativo", "Ruta_Foto", "Ruta_QR", "Fecha_Registro"
        ]
        for c in text_cols:
            if c in df.columns:
                df[c] = df[c].fillna("").astype(object)
        return df
    except Exception as e:
        print(f"Error al cargar Excel: {e}")
        return pd.DataFrame(columns=COLUMNS)


def save_inventory(df: pd.DataFrame) -> bool:
    """Guarda el DataFrame en el archivo Excel principal."""
    try:
        ensure_directories()
        df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")
        return True
    except Exception as e:
        print(f"Error al guardar Excel: {e}")
        return False


def get_prefix(categoria: str, subcategoria: str) -> str:
    """Calcula el prefijo alfanumérico según categoría y subcategoría."""
    key = (categoria, subcategoria)
    if key in PREFIX_MAP:
        return PREFIX_MAP[key]
    wildcard_key = (categoria, "*")
    if wildcard_key in PREFIX_MAP:
        return PREFIX_MAP[wildcard_key]
    return "ACT-IND"


def generate_unique_id(df: pd.DataFrame, categoria: str, subcategoria: str) -> str:
    """
    Genera una clave alfanumérica secuencial basada en la categoría
    (ej. CTR-LAS-001, DBL-PRE-002).
    """
    prefix = get_prefix(categoria, subcategoria)
    pattern = re.compile(rf"^{re.escape(prefix)}-(\d{{3,}})$")

    max_seq = 0
    if not df.empty and "ID_Activo" in df.columns:
        for val in df["ID_Activo"].dropna():
            match = pattern.match(str(val).strip())
            if match:
                seq = int(match.group(1))
                if seq > max_seq:
                    max_seq = seq

    new_seq = max_seq + 1
    return f"{prefix}-{new_seq:03d}"


def save_asset_image(uploaded_file, asset_id: str) -> Optional[str]:
    """
    Guarda la fotografía físicamente en `media/fotos_activos/`
    renombrándola automáticamente con el ID del activo.
    """
    if uploaded_file is None:
        return None
    ensure_directories()
    # Estandarizar a formato JPG
    filename = f"{asset_id}.jpg"
    dest_path = os.path.join(FOTOS_DIR, filename)

    try:
        img = Image.open(uploaded_file)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        # Redimensionar si es excesivamente grande manteniendo relación de aspecto
        max_dim = 1600
        if img.width > max_dim or img.height > max_dim:
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
        img.save(dest_path, "JPEG", quality=85)
        # Retorna ruta relativa estandarizada
        return os.path.join("media", "fotos_activos", filename).replace("\\", "/")
    except Exception as e:
        print(f"Error al guardar imagen: {e}")
        return None


def generate_qr_image_bytes(asset_id: str, extra_data: Optional[Dict[str, Any]] = None) -> BytesIO:
    """Genera el código QR en memoria (BytesIO) para vistas previas en tiempo real."""
    qr_payload = f"SIGRAMA | ACTIVO INDUSTRIAL\nID: {asset_id}"
    if extra_data:
        nombre = extra_data.get("Nombre_Equipo", "")
        serie = extra_data.get("Numero_Serie", "")
        area = extra_data.get("Area_Produccion", "")
        qr_payload += f"\nEquipo: {nombre}\nSerie: {serie}\nÁrea: {area}"

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=8,
        border=3,
    )
    qr.add_data(qr_payload)
    qr.make(fit=True)

    img_qr = qr.make_image(fill_color="#0B4F8A", back_color="white")
    buf = BytesIO()
    img_qr.save(buf, format="PNG")
    buf.seek(0)
    return buf


def generate_qr_code(asset_id: str, extra_data: Optional[Dict[str, Any]] = None) -> str:
    """
    Genera el código QR con qrcode, lo guarda en `media/qrs/QR_{asset_id}.png`
    y almacena la información identificadora del activo.
    """
    ensure_directories()
    filename = f"QR_{asset_id}.png"
    dest_path = os.path.join(QRS_DIR, filename)

    # Contenido del QR: Datos esenciales legibles localmente
    qr_payload = f"SIGRAMA | ACTIVO INDUSTRIAL\nID: {asset_id}"
    if extra_data:
        nombre = extra_data.get("Nombre_Equipo", "")
        serie = extra_data.get("Numero_Serie", "")
        area = extra_data.get("Area_Produccion", "")
        qr_payload += f"\nEquipo: {nombre}\nSerie: {serie}\nÁrea: {area}"

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=4,
    )
    qr.add_data(qr_payload)
    qr.make(fit=True)

    img_qr = qr.make_image(fill_color="#0B4F8A", back_color="white")
    img_qr.save(dest_path)

    return os.path.join("media", "qrs", filename).replace("\\", "/")


def register_new_asset(asset_data: Dict[str, Any], uploaded_file=None) -> Tuple[bool, str, Dict[str, Any]]:
    """
    Orquesta el flujo completo de backend:
    1. Genera ID único.
    2. Procesa y guarda la foto.
    3. Genera código QR.
    4. Escribe la fila en el Excel `inventario_activos.xlsx`.
    """
    df = init_excel_db()

    # 1. Generar ID único
    cat = asset_data.get("Categoria", "Maquinaria Principal")
    subcat = asset_data.get("Subcategoria", "")
    asset_id = generate_unique_id(df, cat, subcat)
    asset_data["ID_Activo"] = asset_id

    # 2. Procesamiento de Imagen
    ruta_foto = save_asset_image(uploaded_file, asset_id) if uploaded_file else ""
    asset_data["Ruta_Foto"] = ruta_foto or ""

    # 3. Generación de Código QR
    ruta_qr = generate_qr_code(asset_id, asset_data)
    asset_data["Ruta_QR"] = ruta_qr

    # Cálculos Financieros
    moi = float(asset_data.get("MOI_Neto", 0.0) or 0.0)
    gastos = float(asset_data.get("Gastos_Inherentes", 0.0) or 0.0)
    asset_data["Inversion_Total"] = moi + gastos
    asset_data["Fecha_Registro"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 4. Escritura en Excel
    new_row = {col: asset_data.get(col, None) for col in COLUMNS}
    new_df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    success = save_inventory(new_df)
    if success:
        return True, asset_id, asset_data
    return False, "Error al guardar en el archivo Excel", asset_data


def generate_asset_pdf(asset: Dict[str, Any]) -> BytesIO:
    """
    Genera la Ficha Técnica Ejecutiva del activo en formato PDF
    utilizando ReportLab, con el logotipo oficial de SIGRAMA,
    diseño ejecutivo a dos columnas y cumplimiento normativo SAT.
    """
    buffer = BytesIO()
    # Márgenes calibrados para diseño ejecutivo en 1 sola página
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=30,
        rightMargin=30,
        topMargin=26,
        bottomMargin=26
    )

    styles = getSampleStyleSheet()
    color_red = colors.HexColor("#EC2024")       # PANTONE 485 C
    color_black = colors.HexColor("#111111")     # PANTONE Black 7 C
    color_gray = colors.HexColor("#475569")
    color_light = colors.HexColor("#F8FAFC")
    border_color = colors.HexColor("#CBD5E1")

    # Estilos tipográficos
    title_corp = ParagraphStyle(
        'TitleCorp',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=12.5,
        leading=15,
        textColor=color_black
    )
    subtitle_corp = ParagraphStyle(
        'SubCorp',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=10,
        textColor=color_gray
    )
    meta_box_style = ParagraphStyle(
        'MetaBox',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=10.5,
        textColor=color_black,
        alignment=2 # Right
    )
    section_title_style = ParagraphStyle(
        'SecTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10,
        textColor=colors.white
    )
    cell_label = ParagraphStyle(
        'CellLabel',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7.5,
        leading=9.5,
        textColor=color_black
    )
    cell_val = ParagraphStyle(
        'CellVal',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=9.5,
        textColor=color_gray
    )
    cell_val_bold = ParagraphStyle(
        'CellValBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10,
        textColor=color_black
    )

    story = []

    # =========================================================================
    # 1. ENCABEZADO CON LOGOTIPO OFICIAL SIGRAMA
    # =========================================================================
    logo_file = "logo_sigrama.png"
    if os.path.exists(logo_file):
        try:
            # Dimensiones proporcionales exactas para el logotipo
            rl_logo = RLImage(logo_file, width=150, height=45)
            left_header_content = [
                rl_logo,
                Spacer(1, 2),
                Paragraph("<b>INDUSTRIA SIGRAMA S.A. DE C.V.</b>", title_corp),
                Paragraph("División de Manufactura 4.0 &bull; Control Central de Planta", subtitle_corp)
            ]
        except Exception:
            left_header_content = [
                Paragraph("<b>INDUSTRIA SIGRAMA S.A. DE C.V.</b>", title_corp),
                Paragraph("División de Manufactura 4.0 &bull; Control Central de Planta", subtitle_corp)
            ]
    else:
        left_header_content = [
            Paragraph("<font color='#EC2024' size=16><b>SIGRAMA</b></font><br/><b>INDUSTRIA SIGRAMA S.A. DE C.V.</b>", title_corp),
            Paragraph("División de Manufactura 4.0 &bull; Control Central de Planta", subtitle_corp)
        ]

    meta_text = f"""
    <b>DOCUMENTO CONTROLADO</b><br/>
    Código: <b>SIG-AF-2026</b> &bull; Rev: <b>02</b><br/>
    Emisión: <b>{datetime.now().strftime('%d/%m/%Y %H:%M')}</b><br/>
    Clave Activo: <font color='#EC2024' size=9><b>{asset.get('ID_Activo', 'N/A')}</b></font><br/>
    Planta: <b>Planta Principal México</b>
    """
    right_header_content = Paragraph(meta_text, meta_box_style)

    header_table = Table([[left_header_content, right_header_content]], colWidths=[360, 192])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 4))

    # Línea decorativa corporativa Rojo Pantone 485 C
    story.append(HRFlowable(width="100%", thickness=2.5, color=color_red, spaceBefore=2, spaceAfter=6))

    # =========================================================================
    # 2. BANNER PRINCIPAL DEL ACTIVO
    # =========================================================================
    nombre_eq = str(asset.get('Nombre_Equipo', 'Sin Nombre')).upper()
    estatus_eq = str(asset.get('Estatus_Operativo', 'Operativo')).upper()
    cat_sub = f"{asset.get('Categoria', '')} &bull; {asset.get('Subcategoria', '')}"
    area_eq = str(asset.get('Area_Produccion', 'N/A')).upper()

    banner_p1 = Paragraph(f"<b>{nombre_eq}</b><br/><font size=7.5 color='#CBD5E1'>{cat_sub} &bull; ÁREA: {area_eq}</font>", ParagraphStyle('BnrP1', parent=styles['Normal'], textColor=colors.white, fontName='Helvetica'))
    banner_p2 = Paragraph(f"ESTATUS:<br/><b>{estatus_eq}</b>", ParagraphStyle('BnrP2', parent=styles['Normal'], textColor=colors.white, alignment=2, fontName='Helvetica-Bold', fontSize=8.5))

    banner_table = Table([[banner_p1, banner_p2]], colWidths=[420, 132])
    banner_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), color_black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('LINELEFT', (1, 0), (1, 0), 2, color_red),
    ]))
    story.append(banner_table)
    story.append(Spacer(1, 7))

    # =========================================================================
    # 3. CUERPO A DOS COLUMNAS
    # =========================================================================
    # --- COLUMNA IZQUIERDA: ESPECIFICACIONES TÉCNICAS, SAT Y FINANZAS ---
    sec_hdr_style = ParagraphStyle('SecHdr', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.white)

    # Bloque 1: Identificación Técnica
    t1_hdr = [Paragraph("1. IDENTIFICACIÓN TÉCNICA Y DE PLANTA", sec_hdr_style), ""]
    t1_rows = [
        t1_hdr,
        [Paragraph("Fabricante / Marca:", cell_label), Paragraph(str(asset.get('Marca', 'N/A')), cell_val)],
        [Paragraph("Modelo Oficial:", cell_label), Paragraph(str(asset.get('Modelo', 'N/A')), cell_val)],
        [Paragraph("Número de Serie:", cell_label), Paragraph(f"<code>{asset.get('Numero_Serie', 'N/A')}</code>", cell_val_bold)],
        [Paragraph("Estación / Área:", cell_label), Paragraph(f"<b>{asset.get('Area_Produccion', 'N/A')}</b>", cell_val_bold)],
        [Paragraph("Custodio Asignado:", cell_label), Paragraph(str(asset.get('Responsable', 'N/A')), cell_val)],
    ]
    table_tec = Table(t1_rows, colWidths=[105, 215])
    table_tec.setStyle(TableStyle([
        ('SPAN', (0, 0), (1, 0)),
        ('BACKGROUND', (0, 0), (1, 0), color_black),
        ('LINELEFT', (0, 0), (0, 0), 3, color_red),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 1.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1.5),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('LINEBELOW', (0, 1), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
    ]))

    # Bloque 2: Cumplimiento Fiscal SAT
    es_imp = asset.get('Es_Importado', False)
    pedimento_val = str(asset.get('Numero_Pedimento', 'N/A')) if es_imp else "No aplica (Adquisición Nacional)"
    origen_val = "Extranjera (Importado)" if es_imp else "Nacional Mexicana"

    t2_hdr = [Paragraph("2. CUMPLIMIENTO FISCAL Y LEGAL (AUDITORÍAS SAT)", sec_hdr_style), ""]
    t2_rows = [
        t2_hdr,
        [Paragraph("UUID Folio Fiscal:", cell_label), Paragraph(f"<font size=6.5><code>{asset.get('UUID_CFDI', 'N/A')}</code></font>", cell_val)],
        [Paragraph("RFC Proveedor Emisor:", cell_label), Paragraph(f"<b>{asset.get('RFC_Proveedor', 'N/A')}</b>", cell_val_bold)],
        [Paragraph("Uso de CFDI (Catálogo):", cell_label), Paragraph(str(asset.get('Uso_CFDI', 'N/A')), cell_val)],
        [Paragraph("Procedencia del Bien:", cell_label), Paragraph(origen_val, cell_val)],
        [Paragraph("Pedimento Aduanal:", cell_label), Paragraph(pedimento_val, cell_val_bold)],
    ]
    table_sat = Table(t2_rows, colWidths=[105, 215])
    table_sat.setStyle(TableStyle([
        ('SPAN', (0, 0), (1, 0)),
        ('BACKGROUND', (0, 0), (1, 0), color_black),
        ('LINELEFT', (0, 0), (0, 0), 3, color_red),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 1.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1.5),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('LINEBELOW', (0, 1), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
    ]))

    # Bloque 3: Información Financiera LISR
    moi_f = float(asset.get('MOI_Neto', 0) or 0)
    gastos_f = float(asset.get('Gastos_Inherentes', 0) or 0)
    tot_f = float(asset.get('Inversion_Total', 0) or 0)
    depr_rate = float(asset.get('Tasa_Depreciacion_Anual', 10) or 10)

    t3_hdr = [Paragraph("3. INFORMACIÓN FINANCIERA (ART. 31-38 LISR)", sec_hdr_style), ""]
    t3_rows = [
        t3_hdr,
        [Paragraph("MOI Neto (sin IVA):", cell_label), Paragraph(f"${moi_f:,.2f} MXN", cell_val)],
        [Paragraph("Gastos Inherentes:", cell_label), Paragraph(f"${gastos_f:,.2f} MXN", cell_val)],
        [Paragraph("Inversión Total:", cell_label), Paragraph(f"<font color='#EC2024'><b>${tot_f:,.2f} MXN</b></font>", cell_val_bold)],
        [Paragraph("Tasa Depreciación Fiscal:", cell_label), Paragraph(f"<b>{depr_rate}% anual (LISR)</b>", cell_val_bold)],
        [Paragraph("Fecha Adquisición:", cell_label), Paragraph(str(asset.get('Fecha_Adquisicion', 'N/A')), cell_val)],
        [Paragraph("Inicio de Puesta en Uso:", cell_label), Paragraph(str(asset.get('Fecha_Inicio_Uso', 'N/A')), cell_val)],
    ]
    table_fin = Table(t3_rows, colWidths=[105, 215])
    table_fin.setStyle(TableStyle([
        ('SPAN', (0, 0), (1, 0)),
        ('BACKGROUND', (0, 0), (1, 0), color_black),
        ('LINELEFT', (0, 0), (0, 0), 3, color_red),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 1.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1.5),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('LINEBELOW', (0, 1), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
    ]))

    left_column_flowables = [
        table_tec,
        Spacer(1, 5),
        table_sat,
        Spacer(1, 5),
        table_fin
    ]

    # --- COLUMNA DERECHA: FOTOGRAFÍA, QR Y SELLO DIGITAL ---
    right_column_flowables = []

    # 1. Fotografía Real
    foto_rel = str(asset.get('Ruta_Foto', ''))
    foto_loaded = False
    if foto_rel and os.path.exists(foto_rel):
        try:
            rl_img = RLImage(foto_rel, width=220, height=140)
            pic_table = Table([[Paragraph("<b>REGISTRO FOTOGRÁFICO EN PLANTA</b>", sec_hdr_style)], [rl_img]], colWidths=[224])
            pic_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), color_black),
                ('LINELEFT', (0, 0), (-1, 0), 3, color_red),
                ('ALIGN', (0, 1), (0, 1), 'CENTER'),
                ('VALIGN', (0, 1), (0, 1), 'MIDDLE'),
                ('BOX', (0, 0), (-1, -1), 1, border_color),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ]))
            right_column_flowables.append(pic_table)
            foto_loaded = True
        except Exception:
            pass

    if not foto_loaded:
        no_pic = Table([
            [Paragraph("<b>REGISTRO FOTOGRÁFICO EN PLANTA</b>", sec_hdr_style)],
            [Paragraph("<br/><br/><i>[Sin fotografía adjunta registrada]</i><br/><br/>", ParagraphStyle('NP', parent=styles['Normal'], alignment=1, fontSize=8, textColor=color_gray))]
        ], colWidths=[224])
        no_pic.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), color_black),
            ('BOX', (0, 0), (-1, -1), 1, border_color),
        ]))
        right_column_flowables.append(no_pic)

    right_column_flowables.append(Spacer(1, 6))

    # 2. Código QR y Tarjeta de Inspección
    qr_rel = str(asset.get('Ruta_QR', ''))
    if qr_rel and os.path.exists(qr_rel):
        try:
            rl_qr = RLImage(qr_rel, width=88, height=88)
            qr_desc = Paragraph(
                f"""<b>ETIQUETA OFICIAL SIGRAMA</b><br/>
                ID: <font color='#EC2024'><b>{asset.get('ID_Activo')}</b></font><br/>
                <font size=6.5 color='#64748B'>Escanee con terminal móvil o escáner de códigos en piso de planta para validar asignación y trazabilidad en el ERP.</font>
                """,
                cell_val
            )
            qr_inner = Table([[rl_qr, qr_desc]], colWidths=[94, 126])
            qr_inner.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 2),
                ('RIGHTPADDING', (0, 0), (-1, -1), 2),
                ('TOPPADDING', (0, 0), (-1, -1), 1),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
            ]))

            qr_block = Table([
                [Paragraph("<b>CÓDIGO QR DE INSPECCIÓN FÍSICA</b>", sec_hdr_style)],
                [qr_inner]
            ], colWidths=[224])
            qr_block.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), color_black),
                ('LINELEFT', (0, 0), (-1, 0), 3, color_red),
                ('BOX', (0, 0), (-1, -1), 1, border_color),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]))
            right_column_flowables.append(qr_block)
        except Exception:
            pass

    right_column_flowables.append(Spacer(1, 5))

    # 3. Sello Digital de Control Interno
    sello_text = f"""
    <b>SELLO DIGITAL DE AUDITORÍA INTERNA:</b><br/>
    <font size=5.5 color='#64748B'>
    SIGRAMA-VERIFY-SHA256:{abs(hash(str(asset.get('ID_Activo')) + str(asset.get('UUID_CFDI'))))}<br/>
    EXPEDIENTE DE ACTIVO RESGUARDADO EN BASE MAESTRA EXCEL Y SERVIDORES LOCALES.
    </font>
    """
    sello_table = Table([[Paragraph(sello_text, ParagraphStyle('St', parent=styles['Normal'], fontSize=6.5, leading=8))]], colWidths=[224])
    sello_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), color_light),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#E2E8F0")),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    right_column_flowables.append(sello_table)

    # Tabla General del Layout
    main_layout = Table([[left_column_flowables, right_column_flowables]], colWidths=[324, 228])
    main_layout.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('LEFTPADDING', (1, 0), (1, 0), 8),
    ]))
    story.append(main_layout)

    # =========================================================================
    # 4. SECCIÓN DE FIRMAS Y VALIDEZ
    # =========================================================================
    story.append(Spacer(1, 14))

    firmas_data = [
        [
            Paragraph("___________________________________<br/><b>Ing. Responsable de Custodia</b><br/><font size=6.5 color='#64748B'>Recepción y Operación en Planta</font>", ParagraphStyle('F1', parent=styles['Normal'], alignment=1, fontSize=7, leading=9)),
            Paragraph("___________________________________<br/><b>Control Contable y Fiscal</b><br/><font size=6.5 color='#64748B'>Validación CFDI / LISR / SAT</font>", ParagraphStyle('F2', parent=styles['Normal'], alignment=1, fontSize=7, leading=9)),
            Paragraph("___________________________________<br/><b>Dirección de Operaciones</b><br/><font size=6.5 color='#64748B'>Aprobación Industria 4.0 SIGRAMA</font>", ParagraphStyle('F3', parent=styles['Normal'], alignment=1, fontSize=7, leading=9))
        ]
    ]
    firmas_table = Table(firmas_data, colWidths=[184, 184, 184])
    firmas_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(KeepTogether(firmas_table))

    # Pie de página institucional
    story.append(Spacer(1, 8))
    pie_institucional = Paragraph(
        "<b>Industria Sigrama S.A. de C.V.</b> &bull; División de Manufactura e Industria 4.0 &bull; <i>Ingeniería que da resultados!!</i> &bull; Documento controlado",
        ParagraphStyle('Pie', parent=styles['Normal'], alignment=1, fontName='Helvetica', fontSize=6.8, textColor=color_gray)
    )
    story.append(pie_institucional)

    doc.build(story)
    buffer.seek(0)
    return buffer


def _generate_demo_photo(asset_id: str, title: str, model: str, area: str) -> str:
    """Genera una imagen gráfica representativa del activo para demostración."""
    ensure_directories()
    dest_path = os.path.join(FOTOS_DIR, f"{asset_id}.jpg")
    try:
        from PIL import ImageDraw
        img = Image.new('RGB', (700, 480), color='#072A4A')
        draw = ImageDraw.Draw(img)
        # Borde y marco industrial
        draw.rectangle([15, 15, 685, 465], outline='#0072CE', width=3)
        draw.rectangle([40, 40, 660, 440], fill='#0B4F8A', outline='#38BDF8', width=2)
        draw.text((60, 65), "SIGRAMA - REGISTRO DE MAQUINARIA", fill='#FFFFFF')
        draw.text((60, 110), title[:40], fill='#38BDF8')
        draw.text((60, 150), f"Modelo: {model}", fill='#E2E8F0')
        draw.text((60, 190), f"Área de Trabajo: {area}", fill='#CBD5E1')
        draw.text((60, 390), f"CONTROL DE ACTIVO: {asset_id}", fill='#A7F3D0')
        img.save(dest_path, "JPEG", quality=90)
        return os.path.join("media", "fotos_activos", f"{asset_id}.jpg").replace("\\", "/")
    except Exception as e:
        print(f"Error al generar foto demo: {e}")
        return ""


def create_demo_assets_if_empty():
    """Crea activos de demostración industriales realistas para SIGRAMA si la base de datos está vacía."""
    df = init_excel_db()
    if not df.empty:
        return

    demos = [
        {
            "Nombre_Equipo": "Centro de Corte Láser Fibra Óptica 6kW",
            "Marca": "Bystronic",
            "Modelo": "ByStar Fiber 3015",
            "Numero_Serie": "BY-60291-MX",
            "Categoria": "Maquinaria Principal",
            "Subcategoria": "Corte Láser",
            "UUID_CFDI": "4A1E8D23-67BC-44F1-92EA-1A8B3567D001",
            "RFC_Proveedor": "BYM980315LK2",
            "Uso_CFDI": "I02 Maquinaria y equipo",
            "Es_Importado": True,
            "Numero_Pedimento": "23 16 3840 7001923",
            "MOI_Neto": 5850000.0,
            "Gastos_Inherentes": 320000.0,
            "Tasa_Depreciacion_Anual": 10.0,
            "Fecha_Adquisicion": "2023-03-15",
            "Fecha_Inicio_Uso": "2023-04-01",
            "Area_Produccion": "Corte",
            "Responsable": "Ing. Carlos Mendoza",
            "Estatus_Operativo": "Operativo",
        },
        {
            "Nombre_Equipo": "Plegadora Dobladora CNC Hidráulica 175T",
            "Marca": "Trumpf",
            "Modelo": "TruBend 5170",
            "Numero_Serie": "TB-5170-8834",
            "Categoria": "Maquinaria Principal",
            "Subcategoria": "Dobladora",
            "UUID_CFDI": "B7C91142-990A-4E32-B871-33C29910D442",
            "RFC_Proveedor": "TRU020511RT8",
            "Uso_CFDI": "I02 Maquinaria y equipo",
            "Es_Importado": True,
            "Numero_Pedimento": "23 16 3840 7002044",
            "MOI_Neto": 3420000.0,
            "Gastos_Inherentes": 145000.0,
            "Tasa_Depreciacion_Anual": 10.0,
            "Fecha_Adquisicion": "2023-06-20",
            "Fecha_Inicio_Uso": "2023-07-05",
            "Area_Produccion": "Doblez",
            "Responsable": "Téc. Roberto Garza",
            "Estatus_Operativo": "Operativo",
        },
        {
            "Nombre_Equipo": "Cabina de Pintura Electrostática en Polvo",
            "Marca": "Wagner",
            "Modelo": "SuperCube Quick-Clean",
            "Numero_Serie": "WG-SC-2022-09",
            "Categoria": "Línea de Pintura",
            "Subcategoria": "Pintura Batch",
            "UUID_CFDI": "9F10DE32-3401-44B8-B590-AA382109CC11",
            "RFC_Proveedor": "WAG881120AB9",
            "Uso_CFDI": "I02 Maquinaria y equipo",
            "Es_Importado": False,
            "Numero_Pedimento": "",
            "MOI_Neto": 1890000.0,
            "Gastos_Inherentes": 98000.0,
            "Tasa_Depreciacion_Anual": 10.0,
            "Fecha_Adquisicion": "2022-11-10",
            "Fecha_Inicio_Uso": "2022-12-01",
            "Area_Produccion": "Pintura",
            "Responsable": "Ing. Sofía Valdés",
            "Estatus_Operativo": "En Mantenimiento",
        },
        {
            "Nombre_Equipo": "Montacargas Eléctrico Hombre Sentado 2.5T",
            "Marca": "Toyota / Raymond",
            "Modelo": "8FBE20",
            "Numero_Serie": "TY-8FBE-99120",
            "Categoria": "Equipo Móvil",
            "Subcategoria": "Montacargas",
            "UUID_CFDI": "88A3CD19-4500-4762-AC81-6543209EE890",
            "RFC_Proveedor": "TMX950812GH4",
            "Uso_CFDI": "I02 Maquinaria y equipo",
            "Es_Importado": False,
            "Numero_Pedimento": "",
            "MOI_Neto": 850000.0,
            "Gastos_Inherentes": 25000.0,
            "Tasa_Depreciacion_Anual": 25.0,
            "Fecha_Adquisicion": "2024-01-15",
            "Fecha_Inicio_Uso": "2024-01-20",
            "Area_Produccion": "Logística",
            "Responsable": "Miguel Ángel Rivas",
            "Estatus_Operativo": "Operativo",
        },
        {
            "Nombre_Equipo": "Juego de Punzón y Dados de Alta Precisión R1",
            "Marca": "Wila",
            "Modelo": "New Standard Pro Tooling",
            "Numero_Serie": "WL-PUN-0918-X",
            "Categoria": "Herramental de Doblez",
            "Subcategoria": "Dado/Punzón",
            "UUID_CFDI": "3C871A02-B891-45FE-8711-2290CDA10988",
            "RFC_Proveedor": "WIL120405NM1",
            "Uso_CFDI": "I05 Dados, troqueles, moldes, matrices y herramental",
            "Es_Importado": True,
            "Numero_Pedimento": "24 16 3840 8000412",
            "MOI_Neto": 260000.0,
            "Gastos_Inherentes": 15000.0,
            "Tasa_Depreciacion_Anual": 35.0,
            "Fecha_Adquisicion": "2024-02-10",
            "Fecha_Inicio_Uso": "2024-02-15",
            "Area_Produccion": "Doblez",
            "Responsable": "Téc. Roberto Garza",
            "Estatus_Operativo": "Crítico",
        },
        {
            "Nombre_Equipo": "Estación Ergonomica de Ensamble y Remachado",
            "Marca": "Bosch Rexroth",
            "Modelo": "EcoSafe Modular Workstation",
            "Numero_Serie": "BR-WS-1142",
            "Categoria": "Estaciones/Mesas",
            "Subcategoria": "Mesa",
            "UUID_CFDI": "77D89C11-0022-498A-9877-1122AA990022",
            "RFC_Proveedor": "BRE010915TR3",
            "Uso_CFDI": "I02 Maquinaria y equipo",
            "Es_Importado": False,
            "Numero_Pedimento": "",
            "MOI_Neto": 145000.0,
            "Gastos_Inherentes": 8000.0,
            "Tasa_Depreciacion_Anual": 10.0,
            "Fecha_Adquisicion": "2024-04-05",
            "Fecha_Inicio_Uso": "2024-04-10",
            "Area_Produccion": "Ensamblado",
            "Responsable": "Téc. Laura Domínguez",
            "Estatus_Operativo": "Operativo",
        }
    ]

    for demo in demos:
        exito, asset_id, res = register_new_asset(demo)
        if exito:
            foto_path = _generate_demo_photo(
                asset_id, demo["Nombre_Equipo"], demo["Modelo"], demo["Area_Produccion"]
            )
            if foto_path:
                df_curr = init_excel_db()
                df_curr.loc[df_curr["ID_Activo"] == asset_id, "Ruta_Foto"] = foto_path
                save_inventory(df_curr)
