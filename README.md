# 🏥 Auditor Inteligente de Facturas y Soportes 🚀

Un sistema automatizado y veloz desarrollado en Python para conciliar grandes listados de facturación médica en Excel con estructuras masivas de carpetas de soportes físicos y digitales. 

Este script escanea miles de directorios en segundos, analiza el contenido de cada carpeta, valida las reglas de negocio, y actualiza visualmente el Excel original inyectando un código de colores auditivos según el estado final de cada soporte.

## 🌟 Características Principales

*   **🖥️ GUI Intuitiva y Persistente:** Interfaz gráfica limpia basada en `Tkinter` con 3 pasos (Listado Excel, Directorio de Soportes Automático, y Ruta de Guardado). Posee **Memoria Automática** (`auditor_config.json`) que recuerda las últimas rutas configuradas al reabrir la app.
*   **⚡ Motor de Escaneo Nativo (Motor HSV 3):** Algoritmo optimizado utilizando recursividad nativa que indexa árboles de carpetas y subcarpetas completos en milisegundos.
*   **👁️ Detección Inteligente de Filtros de Excel:** Integración algorítmica con OpenPyXL que respeta matemáticamente los filtros de Excel (`ws.row_dimensions[i].hidden`). Si el usuario oculta 500 filas en su software, Python solo auditará las 10 filas visibles, previniendo daños o sobrescrituras en el archivo original.
*   **🧠 Resolución de Duplicados:** Prioriza carpetas pendientes con texto; cuando encuentra varias carpetas exactas para una factura, permite elegir la ubicación que se auditará.
*   **🔄 Consolidación de Excel:** Desde el botón **“UNIR EXCEL NUEVO + EXISTENTE”**, conserva las filas del archivo base y agrega únicamente facturas que aún no estén registradas, usando `SFANUMFAC` como clave. Los duplicados conservan los datos antiguos.
*   **🎨 Lógica de Colorimetría Inyectada a Excel:** Verifica los `.pdf` en vivo y clasifica directamente en la hoja de cálculo:
    *   🟩 **Verde ("SIN RADICAR"):** Éxito. Carpeta hallada y validada con **4 o más PDFs** en su interior.
    *   🟨 **Amarillo ("FALTAN SOPORTES"):** Alerta. Carpeta hallada pero contiene **menos de 4 PDFs**.
    *   🟦 **Azul ("[Texto Extraído]"):** Estado Pendiente. Facturas ubicadas en carpetas anómalas o bajo subcarpetas de contingencia. El software elimina el número de la carpeta y **escribe la descripción literal de las letras** directo al Excel.
    *   🟥 **Rojo ("NO CARPETA"):** Alerta Crítica. La factura no existe físicamente en ningún lugar del sistema auditado.
    *   🟪 **Morado ("DUPLICADO (n): ruta1 | ruta2"):** El mismo número de factura aparece en más de una carpeta o ZIP. Se listan las ubicaciones para corregirlas.
    *   🟫 **Marrón claro ("UBICACIÓN EQUIVOCADA"):** Las carpetas numéricas (incluida POLICÍA) deben estar dentro de `CONSULTA EXTERNA`; las carpetas cuyo nombre tiene letras deben estar dentro de `PENDIENTE`.
    *   🟧 **Naranja ("FACTURA NO CORRESPONDE (PDF: n)"):** El número impreso en el FEV (`FACTURA ELECTRONICA DE VENTA No. …`) no coincide con el número de la carpeta. Si el FEV no tiene texto legible (p. ej. escaneado), esta verificación se omite.
*   **Nueva EPS:** Cada PDF debe usar el formato `TIPO_NIT_FACTURA.pdf` (por ejemplo, `FEV_800218979_2197025.pdf`) y corresponder a un tipo permitido en el reglamento. FEV, OPF y PDE son obligatorios; LDP y CRC (comprobante recibido del usuario, Circular N°001 de 2026) también son obligatorios, pero no suman al mínimo. En las demás reglas LDP y CRC tampoco cuentan como soporte. Los demás tipos permitidos pueden variar y se requieren al menos 4 tipos distintos que sí cuentan. Archivos con tipo, formato o factura incorrectos se marcan como **MAL SOPORTADO**; soportes obligatorios faltantes o menos de 4 tipos contables se marcan como **FALTAN SOPORTES**. Se mantienen las reglas de ubicación: carpetas numéricas en `CONSULTA EXTERNA` y carpetas pendientes con texto en `PENDIENTE`; las pendientes conservan su estado azul y las ubicaciones incorrectas se reportan aparte.
*   **Circular N°001 de 2026 y manual de radicación:**
    *   La fecha de la factura se lee del FEV (`Fecha: D/MM/AAAA`). En Nueva EPS el **CRC** se exige solo a facturas con fecha **desde el 11/11/2026**; si la fecha no se puede leer, se exige.
    *   Debe existir **HEV o PDX** (al menos uno). En Nueva EPS aparece como `faltan HEV/PDX`; en General, COOSALUD / Sanidad Militar, Seguros Bolívar ARL y SURA una carpeta completa sin ninguno de los dos se marca 🟨 **"FALTA HEV/PDX"**. No aplica a las aseguradoras que apilan los soportes en pocos PDF (la HC va dentro de otro archivo) ni a Positiva (usa `EPI` de Conexia).
*   **Sanitas:** Cada factura debe estar en un ZIP llamado `NIT_FACTURA.zip` dentro de `SANITAS\CONSULTA EXTERNA`. El ZIP debe contener uno o más archivos PDF directamente en su raíz; subcarpetas, entradas de directorio y archivos de otro tipo se marcan como **MAL SOPORTADO**. Esta validación es independiente de las reglas de Nueva EPS.
*   **📦 Portable (Stand-alone):** Arquitectura optimizada para ser compilada en un ejecutable cerrado (`.exe`) vía PyInstaller, apto para distribución sin requerir la instalación de Python en su destino.

## 🛠️ Stack Tecnológico
*   **Lenguaje:** Python 3.x
*   **Carga y Estructura:** `Pandas`
*   **Manipulación de Libro de Trabajo:** `OpenPyXL` (Cell Mapping, PatternFill)
*   **Interfaz Gráfica:** `Tkinter` / `ttk`
*   **OS Engine:** `os.walk`, `Pathlib`

## 📁 Estructura del proyecto

```text
.
├── auditor_facturas.py       # Aplicación principal
├── auditor_config.json       # Preferencias locales (no versionado)
├── abrir_auditor.bat         # Lanzador de escritorio
├── build_exe.bat             # Constructor del ejecutable
├── auditor_facturas.spec     # Configuración de PyInstaller
├── requirements.txt          # Dependencias (app y compilación)
├── requirements-dev.txt      # Dependencias de desarrollo
├── scripts/                  # Utilidades y tareas auxiliares
├── tests/                    # Pruebas y datos de prueba
├── docs/
│   ├── capturas/              # Imágenes de uso
│   └── referencias/           # Manuales y reglamentos
├── web/                      # Interfaz web independiente
├── Tesseract-OCR/             # Herramientas OCR incluidas
├── build/                     # Archivos temporales de compilación
└── dist/                      # Ejecutables generados
```

## ▶️ Ejecución y pruebas

- Para abrir la aplicación en Windows, ejecuta `abrir_auditor.bat`.
- La interfaz web está en `web/index.html`; ábrela con un navegador moderno y conexión a Internet para cargar ExcelJS y JSZip. Permite auditar con las reglas empresariales, filtrar cada columna con casillas como en Excel, revisar los colores reglamentarios (verde `99FF99`, azul `99CCFF`, rojo `FF9999`, amarillo `FFFF99`, naranja `FFB347`, morado `DDA0DD`, marrón `D2B48C` y rosado `FFC0CB`), leer nombres de PDF y archivos dentro de ZIP, descargar el Excel con filtros y colores, y consolidar dos listados conservando las facturas del archivo base. Los datos se procesan localmente en el navegador; la web no renombra archivos ni extrae texto interno de los PDF.
- Para publicar la web en una URL privada de tu cuenta de Google con Apps Script:
  1. Ejecuta `node scripts\build_apps_script.mjs` para empaquetar CSS y JavaScript actualizados en `web/apps_script/Index.html`.
  2. Instala una vez la herramienta oficial con `npm install --global @google/clasp`, activa la API de Apps Script en `https://script.google.com/home/usersettings` y autorízala con `clasp login`.
  3. Crea el proyecto una sola vez con `clasp create --type webapp --title "Auditor de Facturas" --rootDir web/apps_script`.
  4. Sube los archivos con `clasp push --force` y crea la implementación con `clasp deploy --description "Auditor privado"`. Copia la URL `/exec` que corresponde al ID de implementación.
  5. El manifiesto `web/apps_script/appsscript.json` define el acceso `MYSELF` (solo propietario) y la ejecución como propietario. Para publicar cambios posteriores, ejecuta otra vez el generador y `clasp push --force`; actualiza la misma implementación con `clasp deploy --deploymentId ID_DE_IMPLEMENTACION --description "Actualización"`.

  Google pedirá iniciar sesión para implementar y para abrir el enlace. Los Excel y soportes seleccionados permanecen en el navegador de cada equipo; no se guardan en Apps Script.
- Instalación del entorno (desde la raíz del proyecto):

  ```powershell
  python -m venv .venv
  .venv\Scripts\python.exe -m pip install -r requirements.txt
  ```

  `requirements.txt` contiene las dependencias de la aplicación y de la compilación (customtkinter, PyPDF2, openpyxl, pandas, pyinstaller). `requirements-dev.txt` agrega herramientas opcionales de desarrollo (pytest).

- Para correr toda la suite de pruebas automatizadas desde la raíz:

  ```powershell
  python -m unittest discover -s tests -v
  ```

  Para un solo módulo, por ejemplo Nueva EPS o Sanitas: `python -m unittest discover -s tests -p "test_nueva_eps.py"` o `-p "test_sanitas.py"`. Las pruebas de integración (`test_auditor.py`) crean sus datos y archivos de salida en una carpeta temporal.

- Para compilar el ejecutable ejecuta `build_exe.bat`: usa `.venv\Scripts\python.exe` si existe (si no, `python` del PATH), instala `requirements.txt`, cierra cualquier `auditor_facturas*.exe` abierto, incrementa la versión de parche con `scripts\incrementar_version.py` (actualiza `auditor_facturas.py`, `auditor_facturas.spec` y `abrir_auditor.bat`) y construye con PyInstaller.
- El ejecutable compilado se guarda en `dist/`; los archivos intermedios de PyInstaller se guardan en `build/`.
