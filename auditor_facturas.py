import os
import re
import json
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import subprocess
import zipfile

# Constante de Windows para ocultar ventana de consola al usar subprocess
CREATE_NO_WINDOW = 0x08000000

class DuplicateSelector(tk.Toplevel):
    """Interfaz gráfica para resolver colisiones de facturas duplicadas en el sistema."""
    def __init__(self, parent, factura_id, options):
        super().__init__(parent)
        self.title(f"Duplicados - Fac {factura_id}")
        self.geometry("750x400")
        self.result = None
        
        # Etiqueta informativa
        tk.Label(
            self, 
            text=f"Se han detectado {len(options)} posibles ubicaciones para la factura {factura_id}.\nSelecciona la carpeta correcta para la auditoría:",
            font=("Segoe UI", 10, "bold"),
            pady=15
        ).pack()

        # Marco con lista y barra de desplazamiento
        frame = tk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=5)
        
        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, font=("Consolas", 10), selectmode=tk.SINGLE)
        for opt in options:
            # Convertir Path a cadena para mostrarla
            self.listbox.insert(tk.END, str(opt))
        
        self.listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)

        # Botón de confirmación con color
        tk.Button(
            self, 
            text="✅ Seleccionar Carpeta", 
            command=self.on_select, 
            bg="#A0C4FF", 
            width=25,
            pady=8
        ).pack(pady=20)
        
        # Mantener el foco
        self.transient(parent)
        self.grab_set()
        parent.wait_window(self)

    def on_select(self):
        selection = self.listbox.curselection()
        if selection:
            self.result = Path(self.listbox.get(selection[0]))
            self.destroy()

class InvoiceAuditor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Auditor de Facturas - Escaneo y Mapeo")
        self.root.geometry("650x380")
        self.root.configure(padx=20, pady=20)
        
        self.fills = {
            'VERDE': PatternFill(start_color="99FF99", end_color="99FF99", fill_type="solid"),
            'AZUL': PatternFill(start_color="99CCFF", end_color="99CCFF", fill_type="solid"),
            'ROJO': PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid"),
            'AMARILLO': PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        }
        
        # Variables de las rutas
        self.xl_file_var = tk.StringVar()
        self.search_root_var = tk.StringVar()
        self.save_path_var = tk.StringVar()

        self.config_file = Path(__file__).parent / "auditor_config.json"
        self.load_config()

        self.build_ui()

    def build_ui(self):
        # 1. Archivo Excel
        tk.Label(self.root, text="1. Archivo/Carpeta Raíz (Listado Excel de Facturas):", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 5))
        frame1 = tk.Frame(self.root)
        frame1.pack(fill="x", pady=(0, 15))
        tk.Entry(frame1, textvariable=self.xl_file_var, width=65).pack(side="left", padx=(0, 10))
        tk.Button(frame1, text="Examinar", command=self.sel_excel, bg="#E8F0FE").pack(side="left")

        # 2. Carpeta de Búsqueda
        tk.Label(self.root, text="2. Carpeta de Destino (Donde se buscarán los soportes):", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 5))
        frame2 = tk.Frame(self.root)
        frame2.pack(fill="x", pady=(0, 15))
        tk.Entry(frame2, textvariable=self.search_root_var, width=65).pack(side="left", padx=(0, 10))
        tk.Button(frame2, text="Examinar", command=self.sel_search, bg="#E8F0FE").pack(side="left")

        # 3. Ruta de Guardado
        tk.Label(self.root, text="3. Ruta de Guardado (Dónde se guardará el Excel final):", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 5))
        frame3 = tk.Frame(self.root)
        frame3.pack(fill="x", pady=(0, 25))
        tk.Entry(frame3, textvariable=self.save_path_var, width=65).pack(side="left", padx=(0, 10))
        tk.Button(frame3, text="Examinar", command=self.sel_save, bg="#E8F0FE").pack(side="left")

        # Botón de Procesamiento
        self.btn_run = tk.Button(self.root, text="🚀 INICIAR AUDITORÍA", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", pady=10, command=self.audit_process)
        self.btn_run.pack(fill="x", padx=50)

    def load_config(self):
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.xl_file_var.set(config.get('xl_file', ''))
                    self.search_root_var.set(config.get('search_root', ''))
                    self.save_path_var.set(config.get('save_path', ''))
            except: pass

    def save_config(self):
        try:
            config = {
                'xl_file': self.xl_file_var.get(),
                'search_root': self.search_root_var.get(),
                'save_path': self.save_path_var.get()
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
        except: pass

    def sel_excel(self):
        path = filedialog.askopenfilename(title="Selecciona el Listado Excel", filetypes=[("Archivos Excel", "*.xlsx")])
        if path: 
            self.xl_file_var.set(path)
            self.save_config()

    def sel_search(self):
        path = filedialog.askdirectory(title="Selecciona la Carpeta de Soportes")
        if path: 
            self.search_root_var.set(path)
            self.save_config()

    def sel_save(self):
        path = filedialog.asksaveasfilename(title="Guardar como", defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
        if path: 
            self.save_path_var.set(path)
            self.save_config()

    def _extract_id(self, text):
        """Detección inteligente de números eliminando letras y ceros a la izquierda."""
        if pd.isna(text) or str(text).strip() == "" or str(text).strip().lower() == 'nan': 
            return ""
        val = str(text).strip()
        match = re.search(r'(\d+)$', val)
        if match:
            return match.group(1).lstrip('0') or '0'
        return val

    def build_native_index(self, base_path, ids_to_search):
        """Motor HSV 3: Indexación en una sola pasada iterando el sistema (Python nativo).
        Asegura que SOLO se busquen los IDs indicados en el Excel de forma rápida y sin cmds extra."""
        index = {fid: [] for fid in ids_to_search if fid}
        ids_list = [fid for fid in ids_to_search if fid]
        
        if not ids_list: return {}
        print(f"🚀 Generando mapeo directo para {len(ids_list)} facturas del Excel...")
        
        # Iterar el directorio una sola vez y buscar coincidencias
        for root, dirs, files in os.walk(base_path):
            current_dir = Path(root)
            
            # Buscar en carpetas
            for d in dirs:
                # Limpiar el nombre de la carpeta por si tiene espacios invisibles
                d_upper = d.upper().strip()
                for fid in ids_list:
                    # Emparejamiento generoso pero con preferencia al exacto más adelante
                    if fid in d_upper:
                        index[fid].append(current_dir / d)
                        
            # Buscar en zips
            for f in files:
                if f.lower().endswith('.zip'):
                    f_upper = f.upper().strip()
                    f_clean = Path(f).stem.upper().strip()
                    for fid in ids_list:
                        if fid in f_clean or fid in f_upper:
                            index[fid].append(current_dir / f)

        # Eliminar duplicados si los hubiera
        found_count = 0
        for fid in index:
            index[fid] = list(set(index[fid]))
            if index[fid]:
                found_count += 1
                if fid in ["2127662", "2127695", "2127758"]:
                    print(f"✅ DEBUG - Carpeta encontrada para {fid}: {index[fid]}")
                    
        print(f"✅ DEBUG - Total de IDs mapeados exitosamente con al menos una ruta: {found_count} de {len(ids_list)}")
        return index

    def audit_process(self):
        # Extraer variables de la Interfaz
        xl_file = self.xl_file_var.get()
        search_root = self.search_root_var.get()
        save_path = self.save_path_var.get()

        if not xl_file or not search_root or not save_path:
            messagebox.showwarning("Faltan Rutas", "Por favor, completa las 3 rutas arriba seleccionando un archivo o carpeta en cada una.")
            return

        self.btn_run.config(text="Procesando...", state=tk.DISABLED)
        self.root.update()

        # 2. Carga ID de Excel (Asegurando leer la hoja correcta y filas visibles)
        try:
            print(f"✅ DEBUG - Archivo de Excel seleccionado: {xl_file}")
            wb = load_workbook(xl_file)
            ws = wb.active
            df = pd.read_excel(xl_file, sheet_name=ws.title)
        except Exception as e:
            messagebox.showerror("Error de Lectura", f"No se pudo abrir el Excel:\n{e}")
            return

        col = next((c for c in df.columns if 'SFANUMFAC' in c.upper()), None)
        if not col:
            messagebox.showerror("Error", "No se encontró la columna 'SFANUMFAC'.")
            return

        print("🔍 Extrayendo identificadores únicos (Ignorando filas ocultas/filtradas)...")
        # Extraer IDs respetando si el usuario filtró u ocultó filas en Excel
        excel_ids = set()
        for i, val in enumerate(df[col], start=2):
            if pd.notnull(val) and str(val).strip() != "" and str(val).strip().lower() != 'nan':
                if not ws.row_dimensions[i].hidden:
                    fid = self._extract_id(val)
                    if fid: excel_ids.add(fid)
        
        # 3. Escaneo del Sistema (HSV motor)
        dir_map = self.build_native_index(search_root, excel_ids)

        # 4. Auditoría de Filas
        idx_obs = ws.max_column + 1
        ws.cell(row=1, column=idx_obs).value = "RESULTADO_AUDITORIA"

        print("🚦 Iniciando auditoría visual y de soportes...")
        for i, val in enumerate(df[col], start=2):
            # Ignorar celdas filtradas, ocultas o en blanco
            if ws.row_dimensions[i].hidden:
                continue
            if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == 'nan':
                continue
                
            fid = self._extract_id(val)
            if not fid: continue
            
            matches = dir_map.get(fid, [])
            
            # Gestión Inteligente de Duplicados
            final_path = None
            if matches:
                # 0. Prioridad Máxima: Si la factura tiene copia "PENDIENTE" (por ruta o por letras en el nombre)
                pendientes = [p for p in matches if p.is_dir() and ("PENDIENTES" in str(p).upper() or any(c.isalpha() for c in p.stem))]
                
                exact_folders = [p for p in matches if p.is_dir() and p.name.upper().strip() == fid]
                prefix_folders = [p for p in matches if p.is_dir() and (p.name.upper().strip().startswith(f"{fid}_") or p.name.upper().strip().startswith(f"{fid} "))]
                
                if pendientes:
                    final_path = pendientes[0]
                elif exact_folders:
                    final_path = exact_folders[0]
                elif len(prefix_folders) == 1:
                    final_path = prefix_folders[0]
                elif len([p for p in matches if p.is_dir()]) == 1:
                    final_path = [p for p in matches if p.is_dir()][0]
                elif len(matches) > 1:
                    # Si al final la ambigüedad no se rompe silenciosamente
                    final_path = DuplicateSelector(self.root, fid, matches).result
                elif len(matches) == 1:
                    final_path = matches[0]
                        
            if fid in ["2127662", "2127695", "2127758"]:
                print(f"✅ DEBUG - Procesando Excel fila {i} | Valor_celda: '{val}' -> Fid: '{fid}' -> final_path: {final_path}")

            # Clasificación por Defecto (Rojo)
            fill = self.fills['ROJO']
            msg = "NO CARPETA"

            if final_path:
                count = 0
                try:
                    if final_path.suffix.lower() == '.zip':
                        with zipfile.ZipFile(final_path, 'r') as zf:
                            count = len([f for f in zf.namelist() if f.lower().endswith('.pdf')])
                    else:
                        count = len(list(final_path.glob("*.pdf")))
                except: pass

                # Lógica de Pendientes (Azul): Si el nombre contiene letras/descripción
                stem_upper = final_path.stem.upper()
                has_letters = any(c.isalpha() for c in stem_upper)
                
                if has_letters:
                    # Extraer la descripción limpiando el número de ID y guiones
                    desc = stem_upper.replace(fid, "").replace("_", " ").replace("-", " ").strip()
                    msg = desc if desc else "PENDIENTE"
                    fill = self.fills['AZUL']
                elif count >= 4:
                    msg = "SIN RADICAR"
                    fill = self.fills['VERDE']
                else:
                    msg = f"FALTAN SOPORTES ({count}/4)"
                    fill = self.fills['AMARILLO']

            # Aplicar a la fila del Excel (openpyxl)
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=i, column=col_idx).fill = fill
            ws.cell(row=i, column=idx_obs).value = msg

        # 5. Guardado Final
        if save_path:
            try:
                print(f"✅ DEBUG - Intentando guardar archivo en: {save_path}")
                wb.save(save_path)
                messagebox.showinfo("Completado", "¡Proceso terminado!\nArchivo guardado exitosamente.")
                print(f"✅ DEBUG - GUARDADO CONFIRMADO Y REALIZADO CON ÉXITO.")
            except Exception as e:
                messagebox.showerror("Error al Guardar", f"No se pudo guardar el archivo:\n{e}")
                print(f"❌ DEBUG - ERROR AL GUARDAR (probable archivo abierto o bloqueado por Excel): {e}")
                
        self.btn_run.config(text="🚀 INICIAR AUDITORÍA", state=tk.NORMAL)

if __name__ == "__main__":
    app = InvoiceAuditor()
    app.root.mainloop()
