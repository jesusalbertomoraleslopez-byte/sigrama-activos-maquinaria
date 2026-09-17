"""
seed_data.py
Script generador de datos maestros industriales, fotografías técnicas y códigos QR
para SIGRAMA - APP ACTIVOS MAQUINARIA Y HERRAMIENTAS.
"""

import os
import core_logic as core
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import qrcode
from datetime import datetime

def setup_rich_dataset():
    # Asegurar directorios
    core.ensure_directories()

    # Reiniciar inventario para tener datos limpios y organizados
    df_empty = pd.DataFrame(columns=core.COLUMNS)
    core.save_inventory(df_empty)

    activos = [
        # --- ÁREA: CORTE ---
        {
            'Nombre_Equipo': 'Centro de Corte Láser Fibra Óptica 6kW',
            'Marca': 'Bystronic',
            'Modelo': 'ByStar Fiber 3015',
            'Numero_Serie': 'BY-60291-MX',
            'Categoria': 'Maquinaria Principal',
            'Subcategoria': 'Corte Láser',
            'UUID_CFDI': '4A1E8D23-67BC-44F1-92EA-1A8B3567D001',
            'RFC_Proveedor': 'BYM980315LK2',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': True,
            'Numero_Pedimento': '23 16 3840 7001923',
            'MOI_Neto': 5850000.0,
            'Gastos_Inherentes': 320000.0,
            'Tasa_Depreciacion_Anual': 10.0,
            'Fecha_Adquisicion': '2023-03-15',
            'Fecha_Inicio_Uso': '2023-04-01',
            'Area_Produccion': 'Corte',
            'Responsable': 'Ing. Carlos Mendoza',
            'Estatus_Operativo': 'Operativo',
            'color_theme': '#0B4F8A',
            'subtitulo': 'Corte Láser de Alta Velocidad para Acero al Carbón e Inox'
        },
        {
            'Nombre_Equipo': 'Punzonadora CNC de Alta Frecuencia 300kN',
            'Marca': 'Trumpf',
            'Modelo': 'TruPunch 5000 S12',
            'Numero_Serie': 'TP-5000-4491',
            'Categoria': 'Maquinaria Principal',
            'Subcategoria': 'Punzonadora CNC',
            'UUID_CFDI': '5B2E9D34-78CD-55A2-A3FB-2B9C4678E112',
            'RFC_Proveedor': 'TRU020511RT8',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': True,
            'Numero_Pedimento': '23 16 3840 7002554',
            'MOI_Neto': 4150000.0,
            'Gastos_Inherentes': 180000.0,
            'Tasa_Depreciacion_Anual': 10.0,
            'Fecha_Adquisicion': '2023-05-10',
            'Fecha_Inicio_Uso': '2023-05-25',
            'Area_Produccion': 'Corte',
            'Responsable': 'Ing. Carlos Mendoza',
            'Estatus_Operativo': 'Operativo',
            'color_theme': '#1E3A8A',
            'subtitulo': 'Punzonado y Conformado Rápido con Cabezal Hidráulico'
        },
        {
            'Nombre_Equipo': 'Lijadora y Rebabeadora Automática de Bordes',
            'Marca': 'Timesavers',
            'Modelo': '42 Series Rotary Brush',
            'Numero_Serie': 'TS-42-8821-B',
            'Categoria': 'Maquinaria Principal',
            'Subcategoria': 'Lijadora/Rebabeadora',
            'UUID_CFDI': '8C3F0E45-89DE-66B3-B4AC-3C0D5789F223',
            'RFC_Proveedor': 'TSA100214KJ3',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': True,
            'Numero_Pedimento': '23 16 3840 7003118',
            'MOI_Neto': 1280000.0,
            'Gastos_Inherentes': 65000.0,
            'Tasa_Depreciacion_Anual': 10.0,
            'Fecha_Adquisicion': '2023-08-14',
            'Fecha_Inicio_Uso': '2023-09-01',
            'Area_Produccion': 'Corte',
            'Responsable': 'Téc. Javier Ortiz',
            'Estatus_Operativo': 'En Mantenimiento',
            'color_theme': '#B45309',
            'subtitulo': 'Desbaste de Escoria y Redondeo de Aristas Homogéneo'
        },

        # --- ÁREA: DOBLEZ ---
        {
            'Nombre_Equipo': 'Plegadora Dobladora CNC Hidráulica 175T',
            'Marca': 'Trumpf',
            'Modelo': 'TruBend 5170',
            'Numero_Serie': 'TB-5170-8834',
            'Categoria': 'Maquinaria Principal',
            'Subcategoria': 'Dobladora',
            'UUID_CFDI': 'B7C91142-990A-4E32-B871-33C29910D442',
            'RFC_Proveedor': 'TRU020511RT8',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': True,
            'Numero_Pedimento': '23 16 3840 7002044',
            'MOI_Neto': 3420000.0,
            'Gastos_Inherentes': 145000.0,
            'Tasa_Depreciacion_Anual': 10.0,
            'Fecha_Adquisicion': '2023-06-20',
            'Fecha_Inicio_Uso': '2023-07-05',
            'Area_Produccion': 'Doblez',
            'Responsable': 'Téc. Roberto Garza',
            'Estatus_Operativo': 'Operativo',
            'color_theme': '#065F46',
            'subtitulo': 'Plegado de Precisión 6 Ejes con Compensación Hidráulica'
        },
        {
            'Nombre_Equipo': 'Dobladora Eléctrica Servodrive 40T',
            'Marca': 'Amada',
            'Modelo': 'EG 4010 Compact',
            'Numero_Serie': 'AM-EG-4010-09',
            'Categoria': 'Maquinaria Principal',
            'Subcategoria': 'Dobladora',
            'UUID_CFDI': '6D4A1F56-90EF-77C4-C5BD-4D1E6890A334',
            'RFC_Proveedor': 'AMD910620LK9',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': True,
            'Numero_Pedimento': '24 16 3840 8001102',
            'MOI_Neto': 1950000.0,
            'Gastos_Inherentes': 82000.0,
            'Tasa_Depreciacion_Anual': 10.0,
            'Fecha_Adquisicion': '2024-01-18',
            'Fecha_Inicio_Uso': '2024-02-01',
            'Area_Produccion': 'Doblez',
            'Responsable': 'Téc. Roberto Garza',
            'Estatus_Operativo': 'Operativo',
            'color_theme': '#047857',
            'subtitulo': 'Servomotores Duales para Piezas Pequeñas de Alta Tolerancia'
        },
        {
            'Nombre_Equipo': 'Juego de Punzón y Dados de Alta Precisión R1',
            'Marca': 'Wila',
            'Modelo': 'New Standard Pro Tooling',
            'Numero_Serie': 'WL-PUN-0918-X',
            'Categoria': 'Herramental de Doblez',
            'Subcategoria': 'Dado/Punzón',
            'UUID_CFDI': '3C871A02-B891-45FE-8711-2290CDA10988',
            'RFC_Proveedor': 'WIL120405NM1',
            'Uso_CFDI': 'I05 Dados, troqueles, moldes, matrices y herramental',
            'Es_Importado': True,
            'Numero_Pedimento': '24 16 3840 8000412',
            'MOI_Neto': 260000.0,
            'Gastos_Inherentes': 15000.0,
            'Tasa_Depreciacion_Anual': 35.0,
            'Fecha_Adquisicion': '2024-02-10',
            'Fecha_Inicio_Uso': '2024-02-15',
            'Area_Produccion': 'Doblez',
            'Responsable': 'Téc. Roberto Garza',
            'Estatus_Operativo': 'Crítico',
            'color_theme': '#991B1B',
            'subtitulo': 'Acero Templado CNC-Deephardened para Cargas de Alto Impacto'
        },
        {
            'Nombre_Equipo': 'Matriz Multivía de Cambio Rápido V8-V24',
            'Marca': 'Rolleri',
            'Modelo': 'R1 Multi-V Die Type',
            'Numero_Serie': 'ROL-MV-2024',
            'Categoria': 'Herramental de Doblez',
            'Subcategoria': 'Matriz Multivía',
            'UUID_CFDI': '7E5B2A67-01FA-88D5-D6CE-5E2F7901B445',
            'RFC_Proveedor': 'ROL080312GH8',
            'Uso_CFDI': 'I05 Dados, troqueles, moldes, matrices y herramental',
            'Es_Importado': True,
            'Numero_Pedimento': '24 16 3840 8001990',
            'MOI_Neto': 185000.0,
            'Gastos_Inherentes': 12000.0,
            'Tasa_Depreciacion_Anual': 35.0,
            'Fecha_Adquisicion': '2024-03-05',
            'Fecha_Inicio_Uso': '2024-03-10',
            'Area_Produccion': 'Doblez',
            'Responsable': 'Téc. Roberto Garza',
            'Estatus_Operativo': 'Operativo',
            'color_theme': '#0284C7',
            'subtitulo': 'Ranuras V ajustables de 8mm a 24mm para Chapa Calibre 12-22'
        },

        # --- ÁREA: PINTURA ---
        {
            'Nombre_Equipo': 'Cabina de Pintura Electrostática en Polvo',
            'Marca': 'Wagner',
            'Modelo': 'SuperCube Quick-Clean',
            'Numero_Serie': 'WG-SC-2022-09',
            'Categoria': 'Línea de Pintura',
            'Subcategoria': 'Pintura Batch',
            'UUID_CFDI': '9F10DE32-3401-44B8-B590-AA382109CC11',
            'RFC_Proveedor': 'WAG881120AB9',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': False,
            'Numero_Pedimento': '',
            'MOI_Neto': 1890000.0,
            'Gastos_Inherentes': 98000.0,
            'Tasa_Depreciacion_Anual': 10.0,
            'Fecha_Adquisicion': '2022-11-10',
            'Fecha_Inicio_Uso': '2022-12-01',
            'Area_Produccion': 'Pintura',
            'Responsable': 'Ing. Sofía Valdés',
            'Estatus_Operativo': 'Operativo',
            'color_theme': '#D97706',
            'subtitulo': 'Recuperación Ciclónica de Polvo con Sistema de Cambio de Color'
        },
        {
            'Nombre_Equipo': 'Horno Convectivo de Curado de Pintura 220°C',
            'Marca': 'Blowtherm Industrial',
            'Modelo': 'ThermoCure 6000 Gas',
            'Numero_Serie': 'BLW-TC-220-41',
            'Categoria': 'Línea de Pintura',
            'Subcategoria': 'Horno de Curado',
            'UUID_CFDI': '8F6C3B78-12AB-99E6-E7DF-6F3A8012C556',
            'RFC_Proveedor': 'BLW050720RT2',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': True,
            'Numero_Pedimento': '23 16 3840 7004012',
            'MOI_Neto': 2450000.0,
            'Gastos_Inherentes': 110000.0,
            'Tasa_Depreciacion_Anual': 10.0,
            'Fecha_Adquisicion': '2023-02-22',
            'Fecha_Inicio_Uso': '2023-03-10',
            'Area_Produccion': 'Pintura',
            'Responsable': 'Ing. Sofía Valdés',
            'Estatus_Operativo': 'Operativo',
            'color_theme': '#EA580C',
            'subtitulo': 'Distribución Térmica Homogénea por Quemador Modulante de Gas'
        },
        {
            'Nombre_Equipo': 'Pistolas Electrostáticas de Aplicación Manual',
            'Marca': 'Nordson',
            'Modelo': 'Encore HD Spray System',
            'Numero_Serie': 'ND-ENC-2024-03',
            'Categoria': 'Herramienta de Mano',
            'Subcategoria': 'Pistola de Torque',
            'UUID_CFDI': '9A7D4C89-23BC-00F7-F8EA-7A4B9123D667',
            'RFC_Proveedor': 'NDS990412MK1',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': True,
            'Numero_Pedimento': '24 16 3840 8002230',
            'MOI_Neto': 165000.0,
            'Gastos_Inherentes': 8500.0,
            'Tasa_Depreciacion_Anual': 10.0,
            'Fecha_Adquisicion': '2024-03-12',
            'Fecha_Inicio_Uso': '2024-03-15',
            'Area_Produccion': 'Pintura',
            'Responsable': 'Téc. Andrés Lozano',
            'Estatus_Operativo': 'En Mantenimiento',
            'color_theme': '#CA8A04',
            'subtitulo': 'Bomba de Fase Densa para Espesor Constante sin Efecto Faraday'
        },

        # --- ÁREA: ENSAMBLADO ---
        {
            'Nombre_Equipo': 'Estación Ergonómica de Ensamble y Remachado',
            'Marca': 'Bosch Rexroth',
            'Modelo': 'EcoSafe Modular Workstation',
            'Numero_Serie': 'BR-WS-1142',
            'Categoria': 'Estaciones/Mesas',
            'Subcategoria': 'Mesa',
            'UUID_CFDI': '77D89C11-0022-498A-9877-1122AA990022',
            'RFC_Proveedor': 'BRE010915TR3',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': False,
            'Numero_Pedimento': '',
            'MOI_Neto': 145000.0,
            'Gastos_Inherentes': 8000.0,
            'Tasa_Depreciacion_Anual': 10.0,
            'Fecha_Adquisicion': '2024-04-05',
            'Fecha_Inicio_Uso': '2024-04-10',
            'Area_Produccion': 'Ensamblado',
            'Responsable': 'Téc. Laura Domínguez',
            'Estatus_Operativo': 'Operativo',
            'color_theme': '#0284C7',
            'subtitulo': 'Banco Antiestático ESD con Balanceador de Carga y Rieles Aéreos'
        },
        {
            'Nombre_Equipo': 'Mesa 3D de Sujeción y Soldadura 2000x1000mm',
            'Marca': 'Siegmund',
            'Modelo': 'System 28 Professional 750',
            'Numero_Serie': 'SG-28-200100',
            'Categoria': 'Estaciones/Mesas',
            'Subcategoria': 'Estación de Soldadura',
            'UUID_CFDI': '0B8E5D90-34CD-11A8-A9FB-8B5C0234E778',
            'RFC_Proveedor': 'SGM110825GH3',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': True,
            'Numero_Pedimento': '24 16 3840 8003112',
            'MOI_Neto': 210000.0,
            'Gastos_Inherentes': 14000.0,
            'Tasa_Depreciacion_Anual': 10.0,
            'Fecha_Adquisicion': '2024-04-12',
            'Fecha_Inicio_Uso': '2024-04-18',
            'Area_Produccion': 'Ensamblado',
            'Responsable': 'Téc. Laura Domínguez',
            'Estatus_Operativo': 'Operativo',
            'color_theme': '#0D9488',
            'subtitulo': 'Tablero Cuadriculado Plasma-Nitrurado Anti-Salpicaduras'
        },
        {
            'Nombre_Equipo': 'Celda Robotizada de Soldadura Mig/Mag',
            'Marca': 'KUKA / Fronius',
            'Modelo': 'KR CYBERTECH nano + TPS 400i',
            'Numero_Serie': 'KK-CYB-2023-11',
            'Categoria': 'Maquinaria Principal',
            'Subcategoria': 'Otra Maquinaria Principal',
            'UUID_CFDI': '1C9F6E01-45DE-22B9-B0AC-9C6D1345F889',
            'RFC_Proveedor': 'KUK040319TY5',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': True,
            'Numero_Pedimento': '23 16 3840 7005882',
            'MOI_Neto': 3650000.0,
            'Gastos_Inherentes': 195000.0,
            'Tasa_Depreciacion_Anual': 10.0,
            'Fecha_Adquisicion': '2023-10-02',
            'Fecha_Inicio_Uso': '2023-11-15',
            'Area_Produccion': 'Ensamblado',
            'Responsable': 'Ing. Carlos Mendoza',
            'Estatus_Operativo': 'Operativo',
            'color_theme': '#4338CA',
            'subtitulo': 'Brazo Articulado 6 Grados de Libertad con Posicionador Dual'
        },
        {
            'Nombre_Equipo': 'Remachadora Neumática Industrial de Alta Fuerza',
            'Marca': 'Lobster',
            'Modelo': 'R1A1 Heavy Duty',
            'Numero_Serie': 'LOB-R1A1-901',
            'Categoria': 'Herramienta de Mano',
            'Subcategoria': 'Remachadora Neumática',
            'UUID_CFDI': '2D0A7F12-56EF-33CA-C1BD-0D7E2456A990',
            'RFC_Proveedor': 'LOB140210VB7',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': False,
            'Numero_Pedimento': '',
            'MOI_Neto': 48000.0,
            'Gastos_Inherentes': 2500.0,
            'Tasa_Depreciacion_Anual': 10.0,
            'Fecha_Adquisicion': '2024-05-10',
            'Fecha_Inicio_Uso': '2024-05-12',
            'Area_Produccion': 'Ensamblado',
            'Responsable': 'Téc. Laura Domínguez',
            'Estatus_Operativo': 'Operativo',
            'color_theme': '#6D28D9',
            'subtitulo': 'Fijación Rápida de Remaches Estructurales hasta 1/4 pulgada'
        },

        # --- ÁREA: LOGÍSTICA ---
        {
            'Nombre_Equipo': 'Montacargas Eléctrico Hombre Sentado 2.5T',
            'Marca': 'Toyota / Raymond',
            'Modelo': '8FBE20',
            'Numero_Serie': 'TY-8FBE-99120',
            'Categoria': 'Equipo Móvil',
            'Subcategoria': 'Montacargas',
            'UUID_CFDI': '88A3CD19-4500-4762-AC81-6543209EE890',
            'RFC_Proveedor': 'TMX950812GH4',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': False,
            'Numero_Pedimento': '',
            'MOI_Neto': 850000.0,
            'Gastos_Inherentes': 25000.0,
            'Tasa_Depreciacion_Anual': 25.0,
            'Fecha_Adquisicion': '2024-01-15',
            'Fecha_Inicio_Uso': '2024-01-20',
            'Area_Produccion': 'Logística',
            'Responsable': 'Miguel Ángel Rivas',
            'Estatus_Operativo': 'Operativo',
            'color_theme': '#C2410C',
            'subtitulo': 'Mástil Triple Etapa 4.8m con Batería de Litio de Carga Rápida'
        },
        {
            'Nombre_Equipo': 'Patín Hidráulico de Carga Pesada 3T',
            'Marca': 'Crown',
            'Modelo': 'PTH 50 Series 6000lbs',
            'Numero_Serie': 'CRW-PTH-3301',
            'Categoria': 'Equipo Móvil',
            'Subcategoria': 'Patín Hidráulico',
            'UUID_CFDI': '3E1B8A23-67FA-44DB-D2CE-1E8F3567B001',
            'RFC_Proveedor': 'CRW030615NM4',
            'Uso_CFDI': 'I02 Maquinaria y equipo',
            'Es_Importado': False,
            'Numero_Pedimento': '',
            'MOI_Neto': 32000.0,
            'Gastos_Inherentes': 1800.0,
            'Tasa_Depreciacion_Anual': 25.0,
            'Fecha_Adquisicion': '2024-02-28',
            'Fecha_Inicio_Uso': '2024-03-01',
            'Area_Produccion': 'Logística',
            'Responsable': 'Miguel Ángel Rivas',
            'Estatus_Operativo': 'Operativo',
            'color_theme': '#B91C1C',
            'subtitulo': 'Chasis de Acero Reforzado y Ruedas de Poliuretano Anti-Huella'
        }
    ]

    def create_card_image(asset_id, item):
        W, H = 840, 560
        dest_path = os.path.join(core.FOTOS_DIR, f"{asset_id}.jpg")
        theme = item.get('color_theme', '#0B4F8A')
        
        img = Image.new('RGB', (W, H), color='#0A192F')
        draw = ImageDraw.Draw(img)

        # Marco tecnológico exterior
        draw.rectangle([10, 10, W - 10, H - 10], outline='#38BDF8', width=3)
        draw.rectangle([22, 22, W - 22, H - 22], outline='#1E293B', width=2)

        # Header Corporativo
        draw.rectangle([30, 30, W - 30, 115], fill=theme)
        draw.text((50, 42), "SIGRAMA | INDUSTRIA 4.0 - ACTIVO INDUSTRIAL REGISTRADO", fill='#FFFFFF')
        draw.text((50, 72), f"ID CONTROL: {asset_id}   •   ÁREA: {item['Area_Produccion'].upper()}", fill='#FEF08A')

        # Recuadro central
        draw.rectangle([30, 130, W - 30, 440], fill='#0F172A', outline='#1E3A8A', width=2)

        # Esquema visual de la máquina
        draw.rectangle([50, 150, 500, 420], fill='#1E293B', outline='#0284C7', width=2)
        draw.rectangle([70, 170, 480, 225], fill='#0F172A')
        draw.text((85, 188), f"[ {item['Subcategoria'].upper()} ]", fill='#38BDF8')

        draw.rectangle([70, 240, 480, 400], fill='#020617')
        draw.text((85, 255), f"FABRICANTE: {item['Marca']}", fill='#F8FAFC')
        draw.text((85, 285), f"MODELO:     {item['Modelo']}", fill='#93C5FD')
        draw.text((85, 315), f"SERIE:      {item['Numero_Serie']}", fill='#E2E8F0')
        draw.text((85, 345), f"INVERSIÓN:  ${item['MOI_Neto']:,.2f} MXN", fill='#4ADE80')
        draw.text((85, 372), f"SUBTÍTULO:  {item['subtitulo'][:36]}...", fill='#94A3B8')

        # Panel lateral técnico y fiscal
        draw.rectangle([520, 150, W - 50, 420], fill='#020617', outline='#334155', width=1)
        draw.text((535, 168), "TRAZABILIDAD SAT", fill='#38BDF8')
        draw.text((535, 195), f"Uso CFDI: {item['Uso_CFDI'][:18]}", fill='#94A3B8')
        draw.text((535, 220), f"Depreciación: {item['Tasa_Depreciacion_Anual']}% LISR", fill='#CBD5E1')
        
        orig_txt = "IMPORTADO" if item['Es_Importado'] else "NACIONAL"
        draw.text((535, 245), f"Origen: {orig_txt}", fill='#E2E8F0')

        # Badge de estatus
        est = item['Estatus_Operativo']
        color_est = '#10B981' if est == 'Operativo' else ('#F59E0B' if est == 'En Mantenimiento' else '#EF4444')
        draw.rectangle([535, 280, W - 65, 320], fill=color_est)
        draw.text((550, 292), f"ESTATUS: {est.upper()}", fill='#FFFFFF')

        draw.text((535, 345), f"Custodia: {item['Responsable'][:17]}", fill='#F1F5F9')
        draw.text((535, 375), f"Ingreso:  {item['Fecha_Adquisicion']}", fill='#94A3B8')

        # Banner Inferior
        draw.rectangle([30, 455, W - 30, 530], fill='#020617')
        draw.text((50, 470), f"EQUIPO: {item['Nombre_Equipo']}", fill='#FFFFFF')
        draw.text((50, 498), f"CFDI UUID: {item['UUID_CFDI']}  |  RFC: {item['RFC_Proveedor']}", fill='#64748B')

        img.save(dest_path, "JPEG", quality=95)
        return os.path.join("media", "fotos_activos", f"{asset_id}.jpg").replace("\\", "/")

    print(f"--- Registrando {len(activos)} activos industriales para SIGRAMA ---")
    for act in activos:
        exito, asset_id, res = core.register_new_asset(act)
        if exito:
            foto_path = create_card_image(asset_id, act)
            df_curr = core.init_excel_db()
            df_curr.loc[df_curr["ID_Activo"] == asset_id, "Ruta_Foto"] = foto_path
            core.save_inventory(df_curr)
            print(f" -> Creado {asset_id}: {act['Nombre_Equipo']} ({act['Area_Produccion']})")

    df_final = core.init_excel_db()
    print(f"\nProceso finalizado. Total de activos en base de datos: {len(df_final)}")

if __name__ == "__main__":
    setup_rich_dataset()
