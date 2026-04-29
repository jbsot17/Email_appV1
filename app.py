"""Email Masivo Sender - Interfaz Gráfica"""
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path
from threading import Thread
import time

sys.path.insert(0, str(Path(__file__).parent))

from src.reader import leer_archivo_datos, validar_datos
from src.auth import (
    obtener_config, guardar_credenciales, esta_configurado,
    agregar_cuenta, lista_cuentas, seleccionar_cuenta,
    obtener_cuenta_activa, obtener_cuenta_por_nombre
)
from src.templates import listar_templates, obtener_template, obtener_subject_template
from src.gmail_draft import GmailBorrador


class EmailMasivoApp:
    """Aplicación de escritorio para crear borradores de email."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Email Masivo - Creador de Borradores")
        self.root.geometry("700x680")
        self.root.resizable(False, False)
        
        self.datos = None
        self.stats = None
        self.template_seleccionado = None
        self.adjunto = None
        
        self.color_primary = "#667eea"
        self.color_success = "#27ae60"
        self.color_danger = "#e74c3c"
        
        self.crear_interfaz()
        self.cargar_config()
    
    def crear_interfaz(self):
        """Crea la interfaz gráfica."""
        title = tk.Label(self.root, text="Email Masivo - Creador de Borradores", 
                      font=("Arial", 16, "bold"), fg=self.color_primary)
        title.pack(pady=10)
        
        # ======================
        # Sección 1: Cuentas
        # ======================
        cuenta_frame = tk.LabelFrame(self.root, text="1. Cuenta Gmail", 
                             font=("Arial", 11, "bold"), padx=10, pady=10)
        cuenta_frame.pack(fill="x", padx=10, pady=5)
        
        # Dropdown para seleccionar cuenta
        self.combo_cuentas = ttk.Combobox(cuenta_frame, width=20, state="readonly")
        self.combo_cuentas.pack(side="left", padx=5)
        
        btn_seleccionar = tk.Button(cuenta_frame, text="Usar", 
                           command=self.seleccionar_cuenta_usar,
                           bg=self.color_primary, fg="white")
        btn_seleccionar.pack(side="left", padx=5)
        
        # Nombre de cuenta nueva
        tk.Label(cuenta_frame, text=" Nueva:").pack(side="left", padx=(20,0))
        self.entry_nombre_cuenta = tk.Entry(cuenta_frame, width=15)
        self.entry_nombre_cuenta.pack(side="left", padx=5)
        
        tk.Label(cuenta_frame, text="Email:").pack(side="left", padx=(10,0))
        self.entry_email = tk.Entry(cuenta_frame, width=18)
        self.entry_email.pack(side="left", padx=5)
        
        tk.Label(cuenta_frame, text="Password:").pack(side="left")
        self.entry_password = tk.Entry(cuenta_frame, width=12, show="*")
        self.entry_password.pack(side="left", padx=5)
        
        btn_agregar = tk.Button(cuenta_frame, text="+ Agregar", 
                          command=self.agregar_cuenta,
                          bg=self.color_success, fg="white")
        btn_agregar.pack(side="left", padx=5)
        
        # ======================
        # Sección 2: Cargar Datos
        # ======================
        datos_frame = tk.LabelFrame(self.root, text="2. Cargar Archivo de Datos", 
                             font=("Arial", 11, "bold"), padx=10, pady=10)
        datos_frame.pack(fill="x", padx=10, pady=5)
        
        btn_cargar = tk.Button(datos_frame, text="Cargar XLSX/CSV", 
                         command=self.cargar_datos,
                         bg=self.color_success, fg="white", width=15)
        btn_cargar.pack(side="left")
        
        self.lbl_datos = tk.Label(datos_frame, text="Sin archivo", fg="gray")
        self.lbl_datos.pack(side="left", padx=10)
        
        # ======================
        # Sección 3: Seleccionar Template
        # ======================
        template_frame = tk.LabelFrame(self.root, text="3. Seleccionar Template", 
                                           font=("Arial", 11, "bold"), padx=10, pady=10)
        template_frame.pack(fill="x", padx=10, pady=5)
        
        self.combo_template = ttk.Combobox(template_frame, 
                                       values=listar_templates(), 
                                       width=20, state="readonly")
        self.combo_template.pack(side="left", padx=5)
        self.combo_template.bind("<<ComboboxSelected>>", self.seleccionar_template)
        
        btn_preview = tk.Button(template_frame, text="Ver Preview", 
                          command=self.ver_preview)
        btn_preview.pack(side="left", padx=5)
        
        # Mostrar subject automático
        self.lbl_subject = tk.Label(template_frame, text="", fg="gray", font=("Arial", 9))
        self.lbl_subject.pack(side="left", padx=10)
        
        # ======================
        # Sección 4: Adjunto (Template 3)
        # ======================
        adj_frame = tk.LabelFrame(self.root, text="4. Adjunto (solo Template 3)", 
                                           font=("Arial", 11, "bold"), padx=10, pady=10)
        adj_frame.pack(fill="x", padx=10, pady=5)
        
        btn_adj = tk.Button(adj_frame, text="Seleccionar PDF", 
                         command=self.seleccionar_adjunto)
        btn_adj.pack(side="left")
        
        self.lbl_adj = tk.Label(adj_frame, text="Ninguno", fg="gray")
        self.lbl_adj.pack(side="left", padx=10)
        
        # ======================
        # Sección 5: Resumen
        # ======================
        resumen_frame = tk.LabelFrame(self.root, text="5. Resumen", 
                                           font=("Arial", 11, "bold"), padx=10, pady=10)
        resumen_frame.pack(fill="x", padx=10, pady=5)
        
        self.lbl_resumen = tk.Label(resumen_frame, text="Sin datos cargados", 
                                font=("Arial", 10), fg="gray")
        self.lbl_resumen.pack()
        
        # ======================
        # Botón Crear Borradores
        # ======================
        self.btn_crear = tk.Button(self.root, text="CREAR BORRADORES", 
                              command=self.crear_borradores,
                              bg=self.color_danger, fg="white", 
                              font=("Arial", 12, "bold"), height=2)
        self.btn_crear.pack(fill="x", padx=10, pady=10)
        
        # ======================
        # Log
        # ======================
        log_frame = tk.LabelFrame(self.root, text="Log", 
                                           font=("Arial", 11, "bold"), padx=10, pady=5)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, font=("Courier", 9))
        self.log_text.pack(fill="both", expand=True)
        
        # Función log para mostrar mensajes
        def log_msg(msg):
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.root.update()
        
        self.log = log_msg
    
    def cargar_config(self):
        """Carga la configuración guardada."""
        cuentas = lista_cuentas()
        if cuentas:
            self.combo_cuentas['values'] = cuentas
            self.combo_cuentas.current(0)
            
            cuenta = obtener_cuenta_activa()
            if cuenta:
                self.entry_email.insert(0, cuenta.get('email', ''))
                self.log(f"✓ Cuenta cargada: {cuenta.get('nombre')}")
    
    def agregar_cuenta(self):
        """Agrega una nueva cuenta."""
        nombre = self.entry_nombre_cuenta.get().strip()
        email = self.entry_email.get().strip()
        password = self.entry_password.get().strip()
        
        if not nombre or not email or not password:
            messagebox.showwarning("Falta", "Ingrese nombre, email y password")
            return
        
        agregar_cuenta(nombre, email, password)
        
        self.entry_nombre_cuenta.delete(0, tk.END)
        self.entry_email.delete(0, tk.END)
        self.entry_password.delete(0, tk.END)
        
        cuentas = lista_cuentas()
        self.combo_cuentas['values'] = cuentas
        
        messagebox.showinfo("Guardado", f"Cuenta '{nombre}' guardada")
        self.log(f"✓ Cuenta guardada: {nombre}")
    
    def seleccionar_cuenta_usar(self):
        """Selecciona una cuenta para usar."""
        nombre = self.combo_cuentas.get()
        if not nombre:
            return
        
        cuenta = obtener_cuenta_por_nombre(nombre)
        if cuenta:
            seleccionar_cuenta(nombre)
            self.entry_email.delete(0, tk.END)
            self.entry_email.insert(0, cuenta.get('email', ''))
            self.log(f"✓ Cuenta seleccionada: {nombre}")
    
    def cargar_datos(self):
        """Carga el archivo de datos."""
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo",
            filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv")]
        )
        
        if archivo:
            try:
                self.datos, self.stats = leer_archivo_datos(archivo)
                
                self.lbl_datos.config(
                    text=f"{self.stats['total']} reg.",
                    fg=self.color_success
                )
                
                self.actualizar_resumen()
                
                self.log(f"✓ Archivo: {self.stats['total']} reg.")
                self.log(f"  Válidos: {self.stats['validos']} | Inválidos: {self.stats['invalidos']}")
                
            except Exception as e:
                messagebox.showerror("Error", str(e))
    
    def seleccionar_template(self, event=None):
        """Selecciona un template."""
        self.template_seleccionado = self.combo_template.get()
        
        if self.template_seleccionado and self.datos:
            # Obtener el primer address para el subject
            address = self.datos[0].get('address', '')
            subject_auto = obtener_subject_template(self.template_seleccionado, address)
            self.lbl_subject.config(text=f"Subject: {subject_auto}", fg=self.color_primary)
        
        if self.template_seleccionado:
            self.log(f"✓ Template: {self.template_seleccionado}")
            self.actualizar_resumen()
    
    def ver_preview(self):
        """Muestra preview del template."""
        if not self.template_seleccionado:
            messagebox.showwarning("Warning", "Seleccione un template")
            return
        
        try:
            template = obtener_template(self.template_seleccionado)
            
            ejemplo = {
                'Folio Number': 'EXAMPLE-001',
                'Property Address': '123 Main Street, Miami, FL'
            }
            
            from src.templates import aplicar_variables
            preview = aplicar_variables(template, ejemplo)
            
            win = tk.Toplevel(self.root)
            win.title("Preview")
            win.geometry("600x500")
            
            text = scrolledtext.ScrolledText(win, wrap=tk.WORD)
            text.pack(fill="both", expand=True, padx=10, pady=10)
            text.insert(1.0, preview)
            text.config(state="disabled")
            
        except Exception as e:
            messagebox.showerror("Error", str(e))
    
    def seleccionar_adjunto(self):
        """Selecciona archivo adjunto."""
        archivo = filedialog.askopenfilename(
            title="Seleccionar adjunto",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")]
        )
        
        if archivo:
            self.adjunto = archivo
            self.lbl_adj.config(text=os.path.basename(archivo), fg="black")
            self.log(f"✓ Adjunto: {os.path.basename(archivo)}")
    
    def actualizar_resumen(self):
        """Actualiza el resumen."""
        cuenta = self.combo_cuentas.get()
        
        if self.datos and self.stats and self.template_seleccionado:
            adj = os.path.basename(self.adjunto) if self.adjunto else "Ningún"
            self.lbl_resumen.config(
                text=f"📧 {self.stats['total']} | 📧 {self.template_seleccionado} | 📎 {adj}",
                fg="black"
            )
        elif self.datos and self.stats:
            self.lbl_resumen.config(
                text=f"📧 {self.stats['total']} | Seleccione template",
                fg="orange"
            )
        else:
            self.lbl_resumen.config(text="Sin datos", fg="gray")
    
    def log(self, msg):
        """Agrega al log."""
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.root.update()
    
    def crear_borradores(self):
        """Crea los borradores."""
        if not self.datos:
            messagebox.showwarning("Warning", "Carge un archivo de datos")
            return
        
        if not self.template_seleccionado:
            messagebox.showwarning("Warning", "Seleccione un template")
            return
        
        cuenta = obtener_cuenta_activa()
        if not cuenta:
            messagebox.showwarning("Warning", "Seleccione una cuenta Gmail")
            return
        
        cantidad = len(self.datos)
        confirmar = messagebox.askyesno("Confirmar", f"¿Crear {cantidad} borradores?")
        
        if not confirmar:
            return
        
        try:
            template = obtener_template(self.template_seleccionado)
        except Exception as e:
            messagebox.showerror("Error", f"Template: {e}")
            return
        
        adjunto = None
        if self.template_seleccionado == 'template3.html' and self.adjunto:
            adjunto = self.adjunto
        
        # Obtener subject automático
        subject = obtener_subject_template(self.template_seleccionado, self.datos[0].get('address', ''))
        
        self.btn_crear.config(state="disabled")
        self.log("=" * 40)
        self.log(f"Creando {cantidad} borradores...")
        
        Gmail = GmailBorrador(cuenta['email'], cuenta['app_password'])
        
        def proceso():
            try:
                stats = Gmail.crear_borradores(
                    datos=self.datos,
                    template=template,
                    subject=subject,
                    adjunto=adjunto,
                    callback=self.log
                )
                
                self.root.after(0, lambda: messagebox.showinfo(
                    "Completado",
                    f"OK: {stats['creados']} | Errores: {stats['fallidos']}"
                ))
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            
            self.root.after(0, lambda: self.btn_crear.config(state="normal"))
        
        Thread(target=proceso, daemon=True).start()


def main():
    root = tk.Tk()
    app = EmailMasivoApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()