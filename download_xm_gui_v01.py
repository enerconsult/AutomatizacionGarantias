import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import threading
import sys
import os
import queue
import concurrent.futures

# Importar lógica del script existente
# Asegúrate de que download_xm_file.py esté en la misma carpeta
try:
    import download_xm_file
except ImportError:
    messagebox.showerror("Error", "No se encontró el archivo 'download_xm_file.py'.\nAsegúrese de que esté en la misma carpeta.")
    sys.exit(1)

# --- CLASE CALENDARIO ---
class CalendarDialog(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback = callback
        self.title("Seleccionar Fecha")
        self.geometry("300x250")
        self.resizable(False, False)
        
        self.current_date = datetime.now()
        self.selected_date = None
        
        self.create_widgets()
        
        # Modal
        self.transient(parent)
        self.grab_set()
        self.focus_set()

    def create_widgets(self):
        # Header (Mes y Año)
        header_frame = ttk.Frame(self)
        header_frame.pack(fill='x', padx=5, pady=5)
        
        self.lbl_month_year = ttk.Label(header_frame, text="", font=("Arial", 10, "bold"))
        self.lbl_month_year.pack(side='left', padx=10)
        
        btn_next = ttk.Button(header_frame, text=">", command=self.next_month, width=3)
        btn_next.pack(side='right')
        
        btn_prev = ttk.Button(header_frame, text="<", command=self.prev_month, width=3)
        btn_prev.pack(side='right', padx=2)
        
        # Días
        self.days_frame = ttk.Frame(self)
        self.days_frame.pack(fill='both', expand=True, padx=5)
        
        self.update_calendar()

    def update_calendar(self):
        # Limpiar frame
        for widget in self.days_frame.winfo_children():
            widget.destroy()
            
        year = self.current_date.year
        month = self.current_date.month
        
        self.lbl_month_year.config(text=self.current_date.strftime("%B %Y"))
        
        # Cabecera días
        days = ["Do", "Lu", "Ma", "Mi", "Ju", "Vi", "Sa"]
        for i, day in enumerate(days):
            lbl = ttk.Label(self.days_frame, text=day, font=("Arial", 8))
            lbl.grid(row=0, column=i, padx=2, pady=2)
            
        # Días del mes
        first_day_weekday, num_days = self.get_month_info(year, month)
        # Ajuste para que domingo sea 0
        first_day_weekday = (first_day_weekday + 1) % 7 
        
        row = 1
        col = first_day_weekday
        for day in range(1, num_days + 1):
            btn = ttk.Button(self.days_frame, text=str(day), width=3,
                             command=lambda d=day: self.select_day(d))
            btn.grid(row=row, column=col, padx=2, pady=2)
            
            col += 1
            if col > 6:
                col = 0
                row += 1

    def get_month_info(self, year, month):
        import calendar
        return calendar.monthrange(year, month)

    def next_month(self):
        month = self.current_date.month + 1
        year = self.current_date.year
        if month > 12:
            month = 1
            year += 1
        self.current_date = self.current_date.replace(year=year, month=month, day=1)
        self.update_calendar()

    def prev_month(self):
        month = self.current_date.month - 1
        year = self.current_date.year
        if month < 1:
            month = 12
            year -= 1
        self.current_date = self.current_date.replace(year=year, month=month, day=1)
        self.update_calendar()

    def select_day(self, day):
        self.selected_date = self.current_date.replace(day=day)
        if self.callback:
            self.callback(self.selected_date)
        self.destroy()

# --- APP PRINCIPAL ---
class XMDownloaderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestor de Descargas XM - Por Esquema")
        self.geometry("600x600")
        
        self.esquemas_data = download_xm_file.ESQUEMAS
        
        # Cola para comunicación entre hilos
        self.msg_queue = queue.Queue()

        self.create_widgets()
        # Inicializar combo de esquemas
        self.cmb_scheme['values'] = list(self.esquemas_data.keys())
        self.cmb_scheme.current(0)
        self.on_scheme_change()

        # Iniciar polling de la cola
        self.after(100, self.check_queue)

    def check_queue(self):
        """Revisa la cola de mensajes para actualizar la GUI desde el hilo principal."""
        try:
            while True:
                msg_type, data = self.msg_queue.get_nowait()

                if msg_type == "LOG":
                    self.txt_log.insert(tk.END, data + "\n")
                    self.txt_log.see(tk.END)
                elif msg_type == "ENABLE_BTN":
                    self.btn_run.config(state='normal')
                elif msg_type == "MSGBOX":
                    title, msg = data
                    messagebox.showinfo(title, msg)
                elif msg_type == "ERRORBOX":
                    title, msg = data
                    messagebox.showerror(title, msg)

                self.msg_queue.task_done()
        except queue.Empty:
            pass
        finally:
            self.after(100, self.check_queue)

    def create_widgets(self):
        style = ttk.Style()
        style.configure("Bold.TLabel", font=("Arial", 10, "bold"))
        
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill='both', expand=True)
        
        # --- FECHAS ---
        dates_frame = ttk.LabelFrame(main_frame, text="Rango de Fechas", padding="10")
        dates_frame.pack(fill='x', pady=5)
        
        # Inicio
        ttk.Label(dates_frame, text="Fecha Inicio:").grid(row=0, column=0, padx=5, pady=5)
        self.ent_start_date = ttk.Entry(dates_frame, width=15)
        self.ent_start_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.ent_start_date.grid(row=0, column=1, padx=5)
        ttk.Button(dates_frame, text="📅", width=3, 
                   command=lambda: self.open_calendar(self.ent_start_date)).grid(row=0, column=2)
        
        # Fin
        ttk.Label(dates_frame, text="Fecha Fin:").grid(row=0, column=3, padx=5, pady=5)
        self.ent_end_date = ttk.Entry(dates_frame, width=15)
        self.ent_end_date.insert(0, (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%d"))
        self.ent_end_date.grid(row=0, column=4, padx=5)
        ttk.Button(dates_frame, text="📅", width=3,
                   command=lambda: self.open_calendar(self.ent_end_date)).grid(row=0, column=5)

        # --- CONFIGURACIÓN DE ARCHIVO ---
        config_frame = ttk.LabelFrame(main_frame, text="Configuración de Descarga", padding="10")
        config_frame.pack(fill='x', pady=5)
        
        # Selector de Esquema
        ttk.Label(config_frame, text="Esquema (Carpeta Local):", style="Bold.TLabel").pack(anchor='w')
        self.cmb_scheme = ttk.Combobox(config_frame, state="readonly", width=50)
        self.cmb_scheme.pack(fill='x', pady=(0, 10))
        self.cmb_scheme.bind("<<ComboboxSelected>>", self.on_scheme_change)
        
        # Selector de Archivo
        ttk.Label(config_frame, text="Tipo de Archivo:").pack(anchor='w')
        self.cmb_file_type = ttk.Combobox(config_frame, state="readonly", width=50)
        self.cmb_file_type.pack(fill='x', pady=5)
        
        # --- CARPETA ---
        folder_frame = ttk.Frame(main_frame, padding="0 5")
        folder_frame.pack(fill='x', pady=5)
        
        ttk.Label(folder_frame, text="Carpeta Raíz Local:").pack(anchor='w')
        self.ent_folder = ttk.Entry(folder_frame)
        self.ent_folder.pack(side='left', fill='x', expand=True)
        self.ent_folder.insert(0, os.path.join(os.getcwd(), "Descargas_XM"))
        
        ttk.Button(folder_frame, text="📂", width=3, command=self.select_folder).pack(side='right', padx=5)

        #Info
        lbl_info = ttk.Label(main_frame, text="Nota: Se creará una subcarpeta con el nombre del esquema (ej: /Mensual)", 
                             font=("Arial", 8, "italic"), foreground="gray")
        lbl_info.pack(anchor='w', padx=5)

        # --- BOTÓN EJECUTAR ---
        self.btn_run = ttk.Button(main_frame, text="INICIAR DESCARGA", command=self.start_download_thread)
        self.btn_run.pack(fill='x', pady=20)

        # --- LOGS ---
        log_frame = ttk.LabelFrame(main_frame, text="Registro de Actividad", padding="5")
        log_frame.pack(fill='both', expand=True)
        
        self.txt_log = tk.Text(log_frame, height=10, font=("Consolas", 9))
        self.txt_log.pack(fill='both', expand=True, side='left')
        
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.txt_log.yview)
        scrollbar.pack(side='right', fill='y')
        self.txt_log.config(yscrollcommand=scrollbar.set)

    def on_scheme_change(self, event=None):
        scheme = self.cmb_scheme.get()
        if scheme in self.esquemas_data:
            files = self.esquemas_data[scheme]["archivos"]
            # Agregar opción "TODOS" al principio
            options = ["--- TODOS LOS DEL ESQUEMA ---"] + files
            self.cmb_file_type['values'] = options
            self.cmb_file_type.current(0)

    def open_calendar(self, entry_widget):
        def set_date(date):
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, date.strftime("%Y-%m-%d"))
            
        CalendarDialog(self, set_date)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.ent_folder.delete(0, tk.END)
            self.ent_folder.insert(0, folder)

    def log(self, message):
        """Envia el mensaje a la cola para ser procesado por el hilo principal."""
        self.msg_queue.put(("LOG", message))

    def start_download_thread(self):
        self.btn_run.config(state='disabled')
        self.log("--- Iniciando Proceso ---")
        
        # Validar fechas
        try:
            start_date = datetime.strptime(self.ent_start_date.get(), "%Y-%m-%d")
            end_date = datetime.strptime(self.ent_end_date.get(), "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Error", "Formato de fecha inválido. Use YYYY-MM-DD.")
            self.btn_run.config(state='normal')
            return

        scheme = self.cmb_scheme.get()
        selected_file = self.cmb_file_type.get()
        root_dir = self.ent_folder.get()

        # Ejecutar en hilo separado
        threading.Thread(target=self.run_download_process, 
                         args=(start_date, end_date, scheme, selected_file, root_dir), 
                         daemon=True).start()

    def download_worker(self, url, filename, scheme_folder, scheme):
        """Helper function to run in a thread worker."""
        try:
            success = download_xm_file.download_file(url, filename, scheme_folder)

            if success:
                msg = f"[OK] Descargado: {filename}"
                # Post-procesamiento para TIE
                if scheme == "TIE":
                    full_path = os.path.join(scheme_folder, filename)
                    # Llamar a la función de limpieza
                    new_path, error_msg = download_xm_file.clean_tie_file(full_path)

                    if new_path:
                        new_filename = os.path.basename(new_path)
                        if new_filename != filename:
                             msg += f"\n[INFO] Convertido y Limpiado: {new_filename}"
                        else:
                             msg += "\n[INFO] Archivo TIE Limpiado."
                    else:
                        msg += f"\n[WARN] No se pudo limpiar TIE: {error_msg}"
                return True, msg
            return False, None
        except Exception as e:
            return False, f"[ERROR] {filename}: {str(e)}"

    def run_download_process(self, start_date, end_date, scheme, selected_file, root_dir):
        try:
            current_date = start_date
            delta = timedelta(days=1)
            found_count = 0
            days_count = 0

            # Crear subcarpeta del esquema
            scheme_folder = os.path.join(root_dir, scheme)
            if not os.path.exists(scheme_folder):
                try:
                    os.makedirs(scheme_folder)
                    self.log(f"Creada carpeta: {scheme_folder}")
                except OSError as e:
                    self.log(f"[ERROR] No se pudo crear carpeta: {e}")
                    self.msg_queue.put(("ENABLE_BTN", None))
                    return

            # Determinar lista de archivos a probar
            files_to_try = []
            if selected_file.startswith("--- TODOS"):
                files_to_try = self.esquemas_data[scheme]["archivos"]
                self.log(f"Modo: Escaneando TODOS los archivos del esquema '{scheme}'...")
            else:
                files_to_try = [selected_file]
                self.log(f"Modo: Buscando solo '{selected_file}'...")

            self.log(f"Guardando en: {scheme_folder}")

            # Generar todas las tareas
            tasks = []
            versions = ["", "_V2"]
            extensions = [".xlsx", ".XLSX", ".xls", ".XLS"]

            # Loop dates
            temp_date = start_date
            while temp_date <= end_date:
                days_count += 1
                
                for file_base in files_to_try:
                    variations = [
                        file_base,
                        file_base + " ",
                        file_base.replace(" ", "  ")
                    ]
                    variations = list(set(variations))

                    for variant in variations:
                        for ver in versions:
                            for ext in extensions:
                                url, filename = download_xm_file.get_xm_url(variant, temp_date, esquema_nombre=scheme, version_suffix=ver, extension=ext)
                                tasks.append((url, filename, scheme_folder, scheme))
                
                temp_date += delta

            self.log(f"Generadas {len(tasks)} combinaciones posibles. Iniciando descarga concurrente...")

            # Ejecutar con ThreadPoolExecutor
            # Limitamos workers para no saturar
            with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
                # Map futures to tasks
                future_to_task = {executor.submit(self.download_worker, *task): task for task in tasks}

                for future in concurrent.futures.as_completed(future_to_task):
                    try:
                        success, message = future.result()
                        if success:
                            found_count += 1
                            if message:
                                self.log(message)
                        elif message and "[ERROR]" in message:
                             self.log(message)
                    except Exception as exc:
                        self.log(f"[EXCEPTION] Worker generó excepción: {exc}")

            self.log(f"\n--- Finalizado ---")
            self.log(f"Días escaneados: {days_count}")
            self.log(f"Archivos descargados: {found_count}")
            
            self.msg_queue.put(("ENABLE_BTN", None))
            self.msg_queue.put(("MSGBOX", ("Proceso Terminado", f"Se descargaron {found_count} archivos en '{scheme}'.")))

        except Exception as e:
            self.log(f"[CRITICAL ERROR] {e}")
            self.msg_queue.put(("ENABLE_BTN", None))
            self.msg_queue.put(("ERRORBOX", ("Error Crítico", str(e))))

if __name__ == "__main__":
    app = XMDownloaderApp()
    app.mainloop()
