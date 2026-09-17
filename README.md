# 🏭 APP ACTIVOS - MAQUINARIA Y HERRAMIENTAS (SIGRAMA)
> **Sistema de Control de Activos Fijos, Herramental y Maquinaria para Industria 4.0 / MES**

Aplicación empresarial desarrollada en **Python** con **Streamlit**, persistencia en **Microsoft Excel (.xlsx)** mediante **pandas**, almacenamiento físico local de fotografías de alta resolución, generador dinámico de **códigos QR** para etiquetado en planta y emisión ejecutiva de **Fichas Técnicas en PDF**.

Diseñada especialmente bajo la identidad visual corporativa de **SIGRAMA** y adaptada al estricto cumplimiento fiscal de auditorías ante el **SAT (México)** y la **Ley del Impuesto sobre la Renta (LISR Art. 31-38)**.

---

## 🚀 1. Estructura del Proyecto

```text
app_activos_sigrama/
│
├── app.py                     # Interfaz principal Streamlit con menús secuenciales y vistas
├── core_logic.py              # Lógica de negocio, persistencia Excel, generación de QR y PDF
├── styles.css                 # Identidad visual corporativa SIGRAMA (Azul Corporativo, Kanban Odoo)
├── requirements.txt           # Dependencias del proyecto
├── .gitignore                 # Exclusiones seguras de control de versiones
├── README.md                  # Manual de instalación, despliegue y guía Git
├── inventario_activos.xlsx    # Archivo maestro de persistencia (creado automáticamente)
│
└── media/
    ├── fotos_activos/         # Fotografías físicas de maquinaria nombradas con el ID único
    └── qrs/                   # Códigos QR generados (QR_{ID}.png)
```

---

## 🛠️ 2. Instalación y Ejecución Local

### Paso 1: Clonar o navegar al directorio del proyecto
Abre una terminal (PowerShell o CMD en Windows) y dirígete a la carpeta del proyecto:
```powershell
cd C:\Users\albertol\.gemini\antigravity\scratch\app_activos_sigrama
```

### Paso 2: Crear el Entorno Virtual de Python
Se recomienda el uso de un entorno virtual aislado para evitar conflictos de dependencias:

**En Windows (PowerShell):**
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

*(Si PowerShell bloquea la ejecución de scripts, habilita temporalmente los permisos ejecutando: `Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned`)*

**En Linux / macOS:**
```bash
python3 -m venv venv
source venv/bin/activate
```

### Paso 3: Instalar las Dependencias
Con el entorno virtual activo, instala las librerías necesarias:
```bash
pip install -r requirements.txt
```

### Paso 4: Iniciar la Aplicación
Ejecuta el servidor local de Streamlit:
```bash
streamlit run app.py
```
La aplicación se abrirá automáticamente en tu navegador web en la dirección:
`http://localhost:8501`

---

## 📌 3. Módulos y Arquitectura de Navegación

La barra lateral implementa un menú de navegación estrictamente secuencial:

### 1. 📌 Inicio / Dashboard
*   **KPIs en Tiempo Real:** Total de activos, Inversión total en planta (MOI Neto + Gastos Inherentes), Porcentaje de Disponibilidad Operativa, y Conteo de Activos Críticos o en Mantenimiento.
*   **Tablero Kanban Estilo Odoo (WIP):** Visualización modular por áreas de producción (*Corte, Doblez, Pintura, Ensamblado, Logística*), con badges de estado (Operativo, En Mantenimiento, Crítico).

### 2. ➕ Registro de Activos (Formulario Secuencial en 4 Pasos)
*   **Paso A: Identificación Física:** Nombre del equipo, Marca, Modelo, Número de Serie, Categoría y Subcategoría Dinámica con prefijos automáticos (ej. `CTR-LAS-001`, `DBL-PRE-002`, `EQM-MTC-001`).
*   **Paso B: Cumplimiento Fiscal y Legal (SAT México):** Folio Fiscal UUID CFDI (36 caracteres), RFC del Proveedor, Catálogo de Uso de CFDI (`I02 Maquinaria en general`, `I05 Herramental/Dados`), y Número de Pedimento Aduanal (obligatorio si el equipo es importado).
*   **Paso C: Información Financiera e Inversión (Art. 31-38 LISR):** MOI Neto sin IVA, Gastos Inherentes de Instalación/Cimentación/Fletes, Inversión Total calculada, Tasa de Depreciación Anual %, Fecha de Adquisición y Fecha de Inicio de Uso.
*   **Paso D: Asignación y Evidencia Visual:** Área operativa, Responsable de custodia, Estatus del equipo y carga de imagen (JPG, PNG).

### 3. 🔍 Consulta y Ficha Técnica
*   Buscador rápido multivariable (por ID, Nombre, Marca o Serie) y filtros por Categoría, Área y Estatus.
*   **Ficha Técnica Ejecutiva a 2 Columnas:**
    *   **Columna Izquierda:** Especificaciones técnicas, trazabilidad fiscal ante el SAT, y desglose financiero de inversión y depreciación.
    *   **Columna Derecha:** Fotografía real del equipo, código QR para escaneo en planta y botón para **exportar a PDF oficial**.

### 4. ⚙️ Configuración y Respaldos
*   Descarga directa del archivo de Excel `inventario_activos.xlsx`.
*   Módulo para restaurar/importar archivos Excel externos.
*   Generador con 1 clic de datos de demostración industrial con especificaciones de maquinaria real (Bystronic Láser 6kW, Trumpf Plegadora 175T, Cabina Wagner, Montacargas Toyota, Punzones Wila, etc.).
*   Consola de comandos Git integrada.

---

## 🔒 4. Guía de Control de Versiones con Git y GitHub

Para inicializar el repositorio y respaldar el proyecto en GitHub:

```bash
# 1. Inicializar Git en el directorio
git init

# 2. Agregar todos los archivos (respetando el .gitignore)
git add .

# 3. Crear el primer commit
git commit -m "feat: APP ACTIVOS MAQUINARIA Y HERRAMIENTAS - SIGRAMA v1.0"

# 4. Establecer la rama principal como main
git branch -M main

# 5. Conectar con tu repositorio en GitHub (reemplaza con tu URL)
git remote add origin https://github.com/TU-USUARIO/APP-ACTIVOS-SIGRAMA.git

# 6. Subir los cambios a GitHub
git push -u origin main
```

---

## 📋 5. Cumplimiento Fiscal y Normativo (Auditorías SAT México)
El sistema ha sido estructurado para satisfacer los requisitos del Artículo 31 al 38 de la **Ley del Impuesto sobre la Renta (LISR)** y las reglas de auditoría de comercio exterior y activo fijo del SAT:
1. **Acreditación de Propiedad y Deducibilidad:** Registro del UUID del Comprobante Fiscal Digital por Internet (CFDI) versión 4.0.
2. **Legal Estancia de Maquinaria Extranjera:** Verificación del Número de Pedimento de importación para evitar embargo precautorio en revisiones de planta.
3. **Monto Original de la Inversión (MOI):** Registro desagregado de costos directos de compra más gastos contingentes e indispensables (fletes, seguros y maniobras).
4. **Tasas de Depreciación Lineal:** Flexibilidad de porcentajes según el tipo de bien (ej. 10% Maquinaria, 25% Equipo de Transporte, 35% Troqueles y Matrices).

---

*Desarrollado para SIGRAMA - Excelencia Operativa en Industria 4.0.*
