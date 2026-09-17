"""
download_real_photos.py
Descarga y procesa fotografías reales de maquinaria industrial desde Wikimedia Commons
optimizadas para el inventario de SIGRAMA.
"""

import os
import json
import time
import urllib.request
import urllib.parse
from PIL import Image, ImageDraw
import io
import core_logic as core

# Mapeo exacto de los 16 activos a fotografías reales de maquinaria
PHOTO_MAP = {
    'CTR-LAS-001': 'CNC_Laser_Cutting_Machine.jpg',
    'MAQ-PRN-001': 'Burr-free_Punching_With_Low_Clearance_by_Ehrt.jpg',
    'LIJ-REB-001': 'Laser-and-CNC-Control.jpg',
    'DBL-PRE-001': 'High-tonnage_press_brake.jpg',
    'DBL-PRE-002': 'Bystronic_Bending.jpg',
    'HRR-PUN-001': 'Quick_Change_Tool_System.jpg',
    'HRR-DBL-001': 'TC625_2.jpg',
    'PNT-BAT-001': 'Powder_Coating_Aluminium_Extrusions.jpg',
    'PNT-LIN-001': 'Industrial_oven.jpg',
    'HRR-MAN-001': 'Pistolenvergleich.jpg',
    'EST-MES-001': 'Electronics_workbench.jpg',
    'EST-EST-001': 'G502_Overhead.jpg',
    'MAQ-PRN-002': 'FANUC_6-axis_welding_robots.jpg',
    'HRR-MAN-002': 'SpiralformMachine.jpg',
    'EQM-MTC-001': 'Toyota_L&F_Geneo_005.JPG',
    'EQM-PTH-001': 'Pompwagen.jpg'
}

def get_wikimedia_thumb_url(filename, width=960):
    """Obtiene la URL oficial de la miniatura de alta resolución de Wikimedia."""
    clean_title = filename.replace('File:', '').strip()
    api = f"https://commons.wikimedia.org/w/api.php?action=query&titles=File:{urllib.parse.quote(clean_title)}&prop=imageinfo&iiprop=url&iiurlwidth={width}&format=json"
    headers = {
        'User-Agent': 'SigramaAssetManagerApp/2.0 (admin@sigrama.com.mx)'
    }
    req = urllib.request.Request(api, headers=headers)
    with urllib.request.urlopen(req, timeout=15) as resp:
        d = json.loads(resp.read().decode())
        pages = d.get('query', {}).get('pages', {})
        for k, v in pages.items():
            info = v.get('imageinfo', [])
            if info:
                return info[0].get('thumburl') or info[0].get('url')
    return None

def download_and_inject():
    core.ensure_directories()
    df = core.init_excel_db()

    print(f"Descargando {len(PHOTO_MAP)} fotografias reales de maquinaria para SIGRAMA...")
    headers = {
        'User-Agent': 'SigramaAssetManagerApp/2.0 (admin@sigrama.com.mx)'
    }

    for asset_id, filename in PHOTO_MAP.items():
        dest_path = os.path.join(core.FOTOS_DIR, f"{asset_id}.jpg")
        try:
            # 1. Resolver URL de imagen
            thumb_url = get_wikimedia_thumb_url(filename, width=1024)
            if not thumb_url:
                print(f"[SKIP] No se pudo obtener URL para {asset_id} ({filename})")
                continue

            time.sleep(0.6)  # Pausa cortés para evitar rate-limiting

            # 2. Descargar bytes
            req = urllib.request.Request(thumb_url, headers=headers)
            with urllib.request.urlopen(req, timeout=20) as resp:
                raw_bytes = resp.read()

            # 3. Procesar y optimizar con Pillow
            img = Image.open(io.BytesIO(raw_bytes))
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")

            # Normalizar a máximo 1200px
            max_dim = 1200
            if img.width > max_dim or img.height > max_dim:
                img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)

            # 4. Cintillo corporativo de inspección oficial SIGRAMA
            draw = ImageDraw.Draw(img)
            w, h = img.size
            banner_h = 32
            draw.rectangle([0, h - banner_h, w, h], fill=(17, 17, 17))
            draw.rectangle([0, h - banner_h, 8, h], fill=(236, 32, 36))  # Rojo Pantone 485 C
            draw.text((16, h - banner_h + 8), f"INDUSTRIA SIGRAMA  |  ACTIVO REGISTRADO: {asset_id}", fill=(255, 255, 255))

            # 5. Guardar físicamente
            img.save(dest_path, "JPEG", quality=90)
            rel_path = os.path.join("media", "fotos_activos", f"{asset_id}.jpg").replace("\\", "/")

            if not df.empty and asset_id in df["ID_Activo"].values:
                df.loc[df["ID_Activo"] == asset_id, "Ruta_Foto"] = rel_path

            print(f"[OK] {asset_id} ({filename}) procesado exitosamente -> {img.size[0]}x{img.size[1]}px")

        except Exception as e:
            print(f"[ERROR] en {asset_id} ({filename}): {e}")

        time.sleep(0.5)

    core.save_inventory(df)
    print("\nInventario actualizado con fotografias reales en inventario_activos.xlsx!")

if __name__ == "__main__":
    download_and_inject()
