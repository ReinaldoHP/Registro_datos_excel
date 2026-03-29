# 🏥 Auditor Inteligente de Facturas y Soportes 🚀

Un sistema automatizado y veloz desarrollado en Python para conciliar grandes listados de facturación médica en Excel con estructuras masivas de carpetas de soportes físicos y digitales. 

Este script escanea miles de directorios en segundos, analiza el contenido de cada carpeta, valida las reglas de negocio, y actualiza visualmente el Excel original inyectando un código de colores auditivos según el estado final de cada soporte.

## 🌟 Características Principales

*   **🖥️ GUI Intuitiva y Persistente:** Interfaz gráfica limpia basada en `Tkinter` con 3 pasos (Listado Excel, Directorio de Soportes Automático, y Ruta de Guardado). Posee **Memoria Automática** (`auditor_config.json`) que recuerda las últimas rutas configuradas al reabrir la app.
*   **⚡ Motor de Escaneo Nativo (Motor HSV 3):** Algoritmo optimizado utilizando recursividad nativa que indexa árboles de carpetas y subcarpetas completos en milisegundos.
*   **👁️ Detección Inteligente de Filtros de Excel:** Integración algorítmica con OpenPyXL que respeta matemáticamente los filtros de Excel (`ws.row_dimensions[i].hidden`). Si el usuario oculta 500 filas en su software, Python solo auditará las 10 filas visibles, previniendo daños o sobrescrituras en el archivo original.
*   **🧠 Resolución de Colisiones (Smart Duplicates):** Si el sistema encuentra dos carpetas apuntando al mismo número de factura (Ej. `2128` y `2128 P HC`), el programa bypassa la rigidez numérica y eleva la carpeta anómala o con caracteres alfabéticos a **Prioridad Máxima 0**, asumiendo el contexto de una "Carpeta Pendiente".
*   **🎨 Lógica de Colorimetría Inyectada a Excel:** Verifica los `.pdf` en vivo y clasifica directamente en la hoja de cálculo:
    *   🟩 **Verde ("SIN RADICAR"):** Éxito. Carpeta hallada y validada con **4 o más PDFs** en su interior.
    *   🟨 **Amarillo ("FALTAN SOPORTES"):** Alerta. Carpeta hallada pero contiene **menos de 4 PDFs**.
    *   🟦 **Azul ("[Texto Extraído]"):** Estado Pendiente. Facturas ubicadas en carpetas anómalas o bajo subcarpetas de contingencia. El software elimina el número de la carpeta y **escribe la descripción literal de las letras** directo al Excel.
    *   🟥 **Rojo ("NO CARPETA"):** Alerta Crítica. La factura no existe físicamente en ningún lugar del sistema auditado.
*   **📦 Portable (Stand-alone):** Arquitectura optimizada para ser compilada en un ejecutable cerrado (`.exe`) vía PyInstaller, apto para distribución sin requerir la instalación de Python en su destino.

## 🛠️ Stack Tecnológico
*   **Lenguaje:** Python 3.x
*   **Carga y Estructura:** `Pandas`
*   **Manipulación de Libro de Trabajo:** `OpenPyXL` (Cell Mapping, PatternFill)
*   **Interfaz Gráfica:** `Tkinter` / `ttk`
*   **OS Engine:** `os.walk`, `Pathlib`
