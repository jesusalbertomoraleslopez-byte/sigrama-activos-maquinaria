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
    utilizando ReportLab, con diseño a dos columnas y branding SIGRAMA.
    """
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()
    primary_color = colors.HexColor("#EC2024")     # PANTONE 485 C
    secondary_color = colors.HexColor("#111111")   # PANTONE Black 7 C
    bg_light = colors.HexColor("#F8FAFC")
    dark_gray = colors.HexColor("#2D3748")

    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=15,
        leading=18,
        textColor=secondary_color
    )
    subtitle_style = ParagraphStyle(
        'DocSubTitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=13,
        textColor=colors.HexColor("#4A5568")
    )
    section_heading = ParagraphStyle(
        'SectionHeading',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=10.5,
        leading=14,
        textColor=primary_color
    )
    cell_bold = ParagraphStyle(
        'CellBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=secondary_color
    )
    cell_normal = ParagraphStyle(
        'CellNormal',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        textColor=dark_gray
    )

    story = []

    # Encabezado Corporativo con Logotipo Oficial
    logo_file = "logo_sigrama.png"
    if os.path.exists(logo_file):
        try:
            rl_logo = RLImage(logo_file, width=135, height=38)
            header_cell_left = [
                rl_logo,
                Spacer(1, 4),
                Paragraph("<b>INDUSTRIA SIGRAMA S.A. DE C.V.</b><br/><font size=8 color='#64748B'>División de Manufactura 4.0 & Control de Activos</font>", title_style)
            ]
        except Exception:
            header_cell_left = Paragraph("<b>INDUSTRIA SIGRAMA S.A. DE C.V.</b><br/>División de Manufactura e Industria 4.0", title_style)
    else:
        header_cell_left = Paragraph("<b>INDUSTRIA SIGRAMA S.A. DE C.V.</b><br/>División de Manufactura e Industria 4.0", title_style)

    header_data = [
        [
            header_cell_left,
            Paragraph(f"<b>FICHA TÉCNICA OFICIAL</b><br/>Emisión: {datetime.now().strftime('%d/%m/%Y')}<br/>Clave: <b><font color='#EC2024'>{asset.get('ID_Activo', 'N/A')}</font></b>", subtitle_style)
        ]
    ]
    header_table = Table(header_data, colWidths=[350, 190])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=2.5, color=primary_color, spaceBefore=4, spaceAfter=10))

    # Banner del Activo
    nombre_eq = asset.get('Nombre_Equipo', 'Sin Nombre')
    estatus_eq = asset.get('Estatus_Operativo', 'Operativo')
    banner_text = f"<b>{nombre_eq}</b> | Categoría: {asset.get('Categoria', '')} - {asset.get('Subcategoria', '')} | Estatus: <b>{estatus_eq}</b>"
    banner_p = Paragraph(banner_text, ParagraphStyle('Banner', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9.5, textColor=colors.white))
    banner_table = Table([[banner_p]], colWidths=[540])
    banner_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), secondary_color),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('LINEBELOW', (0, 0), (-1, -1), 2, primary_color),
    ]))
    story.append(banner_table)
    story.append(Spacer(1, 10))

    # Columna Izquierda: Datos Técnicos, Fiscales, Financieros
    tech_data = [
        [Paragraph("SECCIÓN A: IDENTIFICACIÓN TÉCNICA", section_heading), ""],
        [Paragraph("Marca:", cell_bold), Paragraph(str(asset.get('Marca', '')), cell_normal)],
        [Paragraph("Modelo:", cell_bold), Paragraph(str(asset.get('Modelo', '')), cell_normal)],
        [Paragraph("Número de Serie:", cell_bold), Paragraph(str(asset.get('Numero_Serie', '')), cell_normal)],
        [Paragraph("Área de Producción:", cell_bold), Paragraph(str(asset.get('Area_Produccion', '')), cell_normal)],
        [Paragraph("Responsable Asignado:", cell_bold), Paragraph(str(asset.get('Responsable', '')), cell_normal)],
        
        [Paragraph("SECCIÓN B: CUMPLIMIENTO FISCAL (SAT)", section_heading), ""],
        [Paragraph("UUID CFDI:", cell_bold), Paragraph(str(asset.get('UUID_CFDI', '')), cell_normal)],
        [Paragraph("RFC Proveedor:", cell_bold), Paragraph(str(asset.get('RFC_Proveedor', '')), cell_normal)],
        [Paragraph("Uso de CFDI:", cell_bold), Paragraph(str(asset.get('Uso_CFDI', '')), cell_normal)],
        [Paragraph("Equipo Importado:", cell_bold), Paragraph("SÍ" if asset.get('Es_Importado') else "NO", cell_normal)],
        [Paragraph("No. Pedimento:", cell_bold), Paragraph(str(asset.get('Numero_Pedimento', 'N/A')), cell_normal)],

        [Paragraph("SECCIÓN C: FINANCIERA (Art. 31-38 LISR)", section_heading), ""],
        [Paragraph("MOI Neto:", cell_bold), Paragraph(f"${float(asset.get('MOI_Neto', 0) or 0):,.2f} MXN", cell_normal)],
        [Paragraph("Gastos Inherentes:", cell_bold), Paragraph(f"${float(asset.get('Gastos_Inherentes', 0) or 0):,.2f} MXN", cell_normal)],
        [Paragraph("Inversión Total:", cell_bold), Paragraph(f"<b>${float(asset.get('Inversion_Total', 0) or 0):,.2f} MXN</b>", cell_normal)],
        [Paragraph("Tasa Depr. Anual:", cell_bold), Paragraph(f"{asset.get('Tasa_Depreciacion_Anual', 10)}%", cell_normal)],
        [Paragraph("Fecha Adquisición:", cell_bold), Paragraph(str(asset.get('Fecha_Adquisicion', '')), cell_normal)],
        [Paragraph("Inicio de Uso:", cell_bold), Paragraph(str(asset.get('Fecha_Inicio_Uso', '')), cell_normal)],
    ]

    left_table = Table(tech_data, colWidths=[110, 200])
    left_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('SPAN', (0, 0), (1, 0)),
        ('SPAN', (0, 6), (1, 6)),
        ('SPAN', (0, 12), (1, 12)),
        ('LINEBELOW', (0, 0), (1, 0), 1, primary_color),
        ('LINEBELOW', (0, 6), (1, 6), 1, primary_color),
        ('LINEBELOW', (0, 12), (1, 12), 1, primary_color),
        ('TOPPADDING', (0, 0), (-1, -1), 2.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5),
    ]))

    # Columna Derecha: Imagen y Código QR
    right_elements = []
    
    # Imagen del activo
    foto_rel = str(asset.get('Ruta_Foto', ''))
    if foto_rel and os.path.exists(foto_rel):
        try:
            rl_img = RLImage(foto_rel, width=190, height=140)
            right_elements.append(Paragraph("<b>REGISTRO FOTOGRÁFICO</b>", section_heading))
            right_elements.append(Spacer(1, 4))
            right_elements.append(rl_img)
            right_elements.append(Spacer(1, 8))
        except Exception:
            right_elements.append(Paragraph("<i>[Error al cargar imagen fotográfica]</i>", cell_normal))
    else:
        right_elements.append(Paragraph("<b>REGISTRO FOTOGRÁFICO</b>", section_heading))
        right_elements.append(Spacer(1, 4))
        right_elements.append(Paragraph("<i>[Sin fotografía adjunta]</i>", cell_normal))
        right_elements.append(Spacer(1, 40))

    # Código QR
    qr_rel = str(asset.get('Ruta_QR', ''))
    if qr_rel and os.path.exists(qr_rel):
        try:
            rl_qr = RLImage(qr_rel, width=120, height=120)
            right_elements.append(Paragraph("<b>CÓDIGO QR OFICIAL</b>", section_heading))
            right_elements.append(Spacer(1, 4))
            right_elements.append(rl_qr)
            right_elements.append(Paragraph(f"<font size=7>ID: {asset.get('ID_Activo')}</font>", cell_bold))
        except Exception:
            right_elements.append(Paragraph("<i>[QR no disponible]</i>", cell_normal))
    
    # Layout 2 Columnas
    main_layout_data = [[left_table, right_elements]]
    main_layout = Table(main_layout_data, colWidths=[320, 220])
    main_layout.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (1, 0), (1, 0), 15),
    ]))
    story.append(main_layout)

    # Pie de página con firma y validez
    story.append(Spacer(1, 20))
    footer_data = [
        [
            Paragraph("____________________________<br/><b>Responsable de Planta</b><br/>SIGRAMA", cell_normal),
            Paragraph("____________________________<br/><b>Control Contable y Activos</b><br/>Auditoría SAT", cell_normal),
            Paragraph("____________________________<br/><b>Supervisión Mantenimiento</b><br/>Industria 4.0", cell_normal)
        ]
    ]
    footer_table = Table(footer_data, colWidths=[180, 180, 180])
    footer_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    story.append(KeepTogether(footer_table))

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
