# InfoSystem_GUI.py (CORREGIDO)

import tkinter as tk
from tkinter import ttk, messagebox
import wmi
import socket
import datetime
import os
import threading
import pythoncom  

# Importa las funciones de los otros archivos
import InfoSystem_backend as backend
import report_generator as reporter

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Generador de Informacion del Sistema")
        self.geometry("450x550")

        # --- Variables de control ---
        self.info_vars = {}
        self.format_vars = {
            "HTML": tk.BooleanVar(value=True),
            "PDF": tk.BooleanVar(value=False),
            "Excel": tk.BooleanVar(value=False)
        }

        # --- Secciones de información ---
        self.info_options = {
            "General y Red": ("system", "network"),
            "BIOS y Hardware": ("bios",),
            "Procesador (CPU)": ("cpu",),
            "Memoria RAM": ("ram",),
            "Discos": ("disks",),
            "Sistema Operativo": ("os",),
            "Impresoras": ("printers",),
            "Programas Instalados (Lento)": ("software",)
        }
        
        self.backend_functions = {
            "system": backend.get_system_info,
            "network": backend.get_network_info,
            "bios": backend.get_bios_info,
            "cpu": backend.get_cpu_info,
            "ram": backend.get_ram_info,
            "disks": backend.get_disk_info,
            "os": backend.get_os_info,
            "printers": backend.get_installed_printers,
            "software": backend.get_installed_software
        }

        # --- Diseño de la Interfaz ---
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.LabelFrame(main_frame, text="1. Seleccione la información a incluir", padding="10")
        info_frame.pack(fill=tk.X, pady=5)
        
        for name in self.info_options.keys():
            self.info_vars[name] = tk.BooleanVar(value=True)
            cb = ttk.Checkbutton(info_frame, text=name, variable=self.info_vars[name])
            cb.pack(anchor=tk.W, padx=5)

        format_frame = ttk.LabelFrame(main_frame, text="2. Seleccione el formato de salida", padding="10")
        format_frame.pack(fill=tk.X, pady=5)
        
        for fmt, var in self.format_vars.items():
            cb = ttk.Checkbutton(format_frame, text=fmt, variable=var)
            cb.pack(anchor=tk.W, padx=5)
            
        action_frame = ttk.Frame(main_frame, padding="10")
        action_frame.pack(fill=tk.X, pady=10)
        
        self.generate_button = ttk.Button(action_frame, text="Generar Reporte", command=self.start_report_generation)
        self.generate_button.pack(fill=tk.X, ipady=5)
        
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.pack(fill=tk.X, pady=5)

        self.status_label = ttk.Label(main_frame, text="Listo para generar el reporte.")
        self.status_label.pack(fill=tk.X, pady=5)

        # --- Firma inferior ---
        credit_label = ttk.Label(main_frame, text="Created by: Ez07", anchor="center", font=("Arial", 9, "italic"))
        credit_label.pack(fill=tk.X, side=tk.BOTTOM, pady=5)

    def start_report_generation(self):
        if not any(v.get() for v in self.info_vars.values()):
            messagebox.showerror("Error", "Debe seleccionar al menos una categoría de información.")
            return
        if not any(v.get() for v in self.format_vars.values()):
            messagebox.showerror("Error", "Debe seleccionar al menos un formato de salida.")
            return

        self.generate_button.config(state=tk.DISABLED)
        self.status_label.config(text="Iniciando recolección de datos...")
        self.progress_bar.start()

        thread = threading.Thread(target=self.run_report_logic)
        thread.start()

    def run_report_logic(self):
       
        pythoncom.CoInitialize()
        try:
            self.status_label.config(text="Conectando con WMI...")
            c = wmi.WMI()
            
            all_data = {}
            for name, var in self.info_vars.items():
                if var.get():
                    self.status_label.config(text=f"Obteniendo: {name}...")
                    if name == "Discos":
                        all_data['disks'] = backend.get_disk_info()
                    else:
                        for key in self.info_options[name]:
                            all_data[key] = self.backend_functions[key](c)

            self.status_label.config(text="Generando archivos de reporte...")
            hostname = socket.gethostname()
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            base_filename = f"reporte_sistema_{hostname}_{timestamp}"
            
            generated_files = []
            if self.format_vars["HTML"].get():
                filename = f"{base_filename}.html"
                reporter.generate_html(all_data, filename)
                generated_files.append(filename)
            if self.format_vars["PDF"].get():
                filename = f"{base_filename}.pdf"
                reporter.generate_pdf(all_data, filename)
                generated_files.append(filename)
            if self.format_vars["Excel"].get():
                filename = f"{base_filename}.xlsx"
                reporter.generate_excel(all_data, filename)
                generated_files.append(filename)

            messagebox.showinfo("Éxito", f"Reporte(s) generado(s) exitosamente:\n\n" + "\n".join(generated_files))
        
        except Exception as e:
            messagebox.showerror("Error Crítico", f"Ocurrió un error inesperado:\n{e}")
        
        finally:
           
            self.after(100, self.reset_ui)
            pythoncom.CoUninitialize()

    def reset_ui(self):
        self.progress_bar.stop()
        self.status_label.config(text="Listo para generar el reporte.")
        self.generate_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    app = App()
    app.mainloop()