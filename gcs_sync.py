"""
=============================================================================
SIGRAMA APP ACTIVOS - INDUSTRIA SIGRAMA S.A. DE C.V.
Módulo de Persistencia y Sincronización con Google Cloud Storage (GCS)
=============================================================================
"""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
GCS_BUCKET = os.environ.get("GCS_BUCKET", "sigrama-activos-storage").strip()

_client = None
_bucket = None
_synced = False
_last_sync_ts = 0.0
SYNC_TTL = 300.0


def get_bucket():
    """Retorna el bucket de GCS autenticado."""
    global _client, _bucket
    if _bucket is not None:
        return _bucket
    try:
        from google.cloud import storage
        _client = storage.Client()
        _bucket = _client.bucket(GCS_BUCKET)
        return _bucket
    except Exception as e:
        print(f"[GCS_SYNC] Aviso: No se pudo conectar a GCS ({GCS_BUCKET}): {e}")
        return None


def is_gcs_available() -> bool:
    """Verifica si el bucket está accesible."""
    b = get_bucket()
    if b is None:
        return False
    try:
        return b.exists()
    except Exception:
        try:
            list(b.list_blobs(max_results=1))
            return True
        except Exception:
            return False


def sync_from_gcs(force: bool = False) -> bool:
    """
    Descarga bases de datos Excel, fotos, QRs y documentos desde GCS al arrancar.
    Garantiza que la instancia de Cloud Run tenga siempre el estado persistente más reciente.
    """
    global _synced
    if _synced and not force:
        return True
    b = get_bucket()
    if not b:
        return False
    try:
        count = 0
        for blob in b.list_blobs():
            if blob.name.endswith("/"):
                continue
            local_path = BASE_DIR / blob.name
            local_path.parent.mkdir(parents=True, exist_ok=True)
            if not local_path.exists() or local_path.stat().st_size != blob.size:
                blob.download_to_filename(str(local_path))
                count += 1
        _synced = True
        print(f"[GCS_SYNC] Sincronizados exitosamente {count} archivos desde gs://{GCS_BUCKET}")
        return True
    except Exception as e:
        print(f"[GCS_SYNC] Error al sincronizar desde gs://{GCS_BUCKET}: {e}")
        return False


def push_file_to_gcs(file_path: str) -> bool:
    """Sube un archivo específico a GCS conservando su ruta relativa respecto a BASE_DIR."""
    b = get_bucket()
    if not b:
        return False
    try:
        p = Path(file_path)
        if not p.is_absolute():
            p = BASE_DIR / file_path
        if not p.exists() or not p.is_file():
            return False
        rel_path = p.relative_to(BASE_DIR).as_posix()
        blob = b.blob(rel_path)
        blob.upload_from_filename(str(p))
        print(f"[GCS_SYNC] Subido exitosamente a GCS: {rel_path}")
        return True
    except Exception as e:
        print(f"[GCS_SYNC] Error al subir {file_path} a GCS: {e}")
        return False


def delete_file_from_gcs(file_path: str) -> bool:
    """Elimina un archivo del bucket de GCS."""
    b = get_bucket()
    if not b:
        return False
    try:
        p = Path(file_path)
        if p.is_absolute():
            rel_path = p.relative_to(BASE_DIR).as_posix()
        else:
            rel_path = file_path.replace("\\", "/")
        blob = b.blob(rel_path)
        if blob.exists():
            blob.delete()
            print(f"[GCS_SYNC] Eliminado exitosamente de GCS: {rel_path}")
        return True
    except Exception as e:
        print(f"[GCS_SYNC] Error al eliminar {file_path} de GCS: {e}")
        return False


def push_all_to_gcs() -> int:
    """Sube todo el inventario, fotos, QRs y repositorio documental a GCS."""
    b = get_bucket()
    if not b:
        return 0
    uploaded = 0
    # 1. Excel maestro
    excel_p = BASE_DIR / "inventario_activos.xlsx"
    if excel_p.exists():
        if push_file_to_gcs(str(excel_p)):
            uploaded += 1
    # 2. Índice de documentos
    idx_p = BASE_DIR / "media" / "documentos" / "_index.json"
    if idx_p.exists():
        if push_file_to_gcs(str(idx_p)):
            uploaded += 1
    # 3. Fotos principales y QRs
    for sub in ["fotos_activos", "qrs"]:
        d = BASE_DIR / "media" / sub
        if d.exists():
            for f in d.iterdir():
                if f.is_file():
                    if push_file_to_gcs(str(f)):
                        uploaded += 1
    # 4. Documentos de expedientes
    docs_dir = BASE_DIR / "media" / "documentos"
    if docs_dir.exists():
        for f in docs_dir.rglob("*"):
            if f.is_file() and f.name != "_index.json":
                if push_file_to_gcs(str(f)):
                    uploaded += 1
    return uploaded
