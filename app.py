"""Email Masivo Sender - Interfaz Gráfica"""
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path
from threading import Thread

sys.path.insert(0, str(Path(__file__).parent))

from src.reader import leer_archivo_datos, validar_datos
from src.auth import (
    obtener_config, guardar_credenciales, esta_configurado,
    agregar_cuenta, lista_cuentas, seleccionar_cuenta,
    obtener_cuenta_activa, obtener_cuenta_por_nombre, eliminar_cuenta
)
from src.templates import listar_templates, obtener_template, obtener_subject_template
from src.gmail_draft import GmailBorrador


class EmailMasivoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Masivo - Enviador de Emails")
        self.root.geometry("750x680")
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
        title = tk.Label(self.root, text="Email Masivo - Enviador de Emails", 
                      font=("Arial", 16, "bold"), fg=self.color_primary)
        title.pack(pady=10)
        
        # Cuenta Gmail
        cuenta_frame = tk.LabelFrame(self.root, text="1. Cuenta Gmail", 
                             font=("Arial", 11, "bold"), padx=10, pady=10)
        cuenta_frame.pack(fill="x", padx=10, pady=5)
        
        # Cuenta Gmail - Primera fila
        cuenta_top = tk.Frame(cuenta_frame)
        cuenta_top.pack(fill="x", pady=2)
        
        self.combo_cuentas = ttk.Combobox(cuenta_top, width=18, state="readonly")
        self.combo_cuentas.pack(side="left", padx=5)
        
        btn_usar = tk.Button(cuenta_top, text="Usar", 
                           command=self.seleccionar_cuenta_usar,
                           bg=self.color_primary, fg="white", width=8)
        btn_usar.pack(side="left", padx=3)
        
        btn_eliminar = tk.Button(cuenta_top, text="🗑 Eliminar", 
                          command=self.eliminar_cuenta,
                          bg=self.color_danger, fg="white", width=12)
        btn_eliminar.pack(side="left", padx=3)
        
# Cuenta Gmail - Segunda fila
        cuenta_bot = tk.Frame(cuenta_frame)
        cuenta_bot.pack(fill="x", pady=2)
        
        tk.Label(cuenta_bot, text="Nombre:").pack(side="left")
        self.entry_nombre_cuenta = tk.Entry(cuenta_bot, width=12)
        self.entry_nombre_cuenta.pack(side="left", padx=3)
        
        tk.Label(cuenta_bot, text="Email:").pack(side="left", padx=3)
        self.entry_email = tk.Entry(cuenta_bot, width=18)
        self.entry_email.pack(side="left", padx=3)
        
        tk.Label(cuenta_bot, text="Pass:").pack(side="left", padx=3)
        self.entry_password = tk.Entry(cuenta_bot, width=10, show="*")
        self.entry_password.pack(side="left", padx=3)
        
        # Cuenta Gmail - Tercera fila para Remitente
        cuenta_remitente = tk.Frame(cuenta_frame)
        cuenta_remitente.pack(fill="x", pady=2)
        
        tk.Label(cuenta_remitente, text="Remitente:").pack(side="left")
        self.entry_sender_name = tk.Entry(cuenta_remitente, width=25)
        self.entry_sender_name.pack(side="left", padx=3)
        tk.Label(cuenta_remitente, text="(como sale en email)", fg="gray", font=("Arial", 8)).pack(side="left")
        
        btn_agregar = tk.Button(cuenta_bot, text="+ Agregar", 
                           command=self.agregar_cuenta,
                           bg=self.color_success, fg="white", width=10)
        btn_agregar.pack(side="left", padx=5)
        
        # Datos
        datos_frame = tk.LabelFrame(self.root, text="2. Cargar Archivo de Datos", 
                             font=("Arial", 11, "bold"), padx=10, pady=10)
        datos_frame.pack(fill="x", padx=10, pady=5)
        
        btn_cargar = tk.Button(datos_frame, text="📁 Cargar XLSX/CSV", 
                         command=self.cargar_datos,
                         bg=self.color_success, fg="white", width=18)
        btn_cargar.pack(side="left")
        
        self.lbl_datos = tk.Label(datos_frame, text="Sin archivo", fg="gray")
        self.lbl_datos.pack(side="left", padx=10)
        
        # Template
        template_frame = tk.LabelFrame(self.root, text="3. Seleccionar Template", 
                                           font=("Arial", 11, "bold"), padx=10, pady=10)
        template_frame.pack(fill="x", padx=10, pady=5)
        
        from src.templates import listar_templates
        templates = listar_templates()
        
        self.combo_template = ttk.Combobox(template_frame, 
                                       values=templates, 
                                       width=20, state="readonly")
        self.combo_template.pack(side="left", padx=5)
        self.combo_template.bind("<<ComboboxSelected>>", self.seleccionar_template)
        
        btn_preview = tk.Button(template_frame, text="Ver Preview", 
                          command=self.ver_preview)
        btn_preview.pack(side="left", padx=5)
        
        self.lbl_subject = tk.Label(template_frame, text="", fg="gray", font=("Arial", 9))
        self.lbl_subject.pack(side="left", padx=10)
        
# Adjunto
        adj_frame = tk.LabelFrame(self.root, text="4. Adjuntar Archivo", 
                                           font=("Arial", 11, "bold"), padx=10, pady=10)
        adj_frame.pack(fill="x", padx=10, pady=5)
        
        btn_adj = tk.Button(adj_frame, text="📎 Seleccionar PDF", 
                          command=self.seleccionar_adjunto,
                          bg="#3498db", fg="white", width=18)
        btn_adj.pack(side="left", padx=5)
        
        self.lbl_adj = tk.Label(adj_frame, text="Ninguno", fg="gray", font=("Arial", 10))
        self.lbl_adj.pack(side="left", padx=10)
        
        # Resumen
        resumen_frame = tk.LabelFrame(self.root, text="5. Resumen", 
                                           font=("Arial", 11, "bold"), padx=10, pady=10)
        resumen_frame.pack(fill="x", padx=10, pady=5)
        
        self.lbl_resumen = tk.Label(resumen_frame, text="Sin datos cargados", 
                                font=("Arial", 10), fg="gray")
        self.lbl_resumen.pack()
        
        # Botón enviar
        self.btn_enviar = tk.Button(self.root, text="ENVIAR EMAILS", 
                              command=self.enviar_emails,
                              bg=self.color_danger, fg="white", 
                              font=("Arial", 12, "bold"), height=2)
        self.btn_enviar.pack(fill="x", padx=10, pady=10)
        
        # Log
        log_frame = tk.LabelFrame(self.root, text="Log", 
                                           font=("Arial", 11, "bold"), padx=10, pady=5)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, font=("Courier", 9))
        self.log_text.pack(fill="both", expand=True)
        
        def log_msg(msg):
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.root.update()
        
        self.log = log_msg
    
    def cargar_config(self):
        cuentas = lista_cuentas()
        if cuentas:
            self.combo_cuentas['values'] = cuentas
            self.combo_cuentas.current(0)
            
            cuenta = obtener_cuenta_activa()
            if cuenta:
                self.entry_email.insert(0, cuenta.get('email', ''))
                self.entry_sender_name.insert(0, cuenta.get('sender_name', ''))
                self.log(f"✓ Cuenta cargada: {cuenta.get('nombre')}")
    
    def agregar_cuenta(self):
        nombre = self.entry_nombre_cuenta.get().strip()
        email = self.entry_email.get().strip()
        password = self.entry_password.get().strip()
        sender_name = self.entry_sender_name.get().strip()
        
        if not nombre or not email or not password:
            messagebox.showwarning("Falta", "Ingrese nombre, email y password")
            return
        
        agregar_cuenta(nombre, email, password, sender_name)
        
        self.entry_nombre_cuenta.delete(0, tk.END)
        self.entry_email.delete(0, tk.END)
        self.entry_password.delete(0, tk.END)
        self.entry_sender_name.delete(0, tk.END)
        
        cuentas = lista_cuentas()
        self.combo_cuentas['values'] = cuentas
        
        messagebox.showinfo("Guardado", f"Cuenta '{nombre}' guardada")
        self.log(f"✓ Cuenta guardada: {nombre}")
    
    def seleccionar_cuenta_usar(self):
        nombre = self.combo_cuentas.get()
        if not nombre:
            return
        
        cuenta = obtener_cuenta_por_nombre(nombre)
        if cuenta:
            seleccionar_cuenta(nombre)
            self.entry_email.delete(0, tk.END)
            self.entry_email.insert(0, cuenta.get('email', ''))
            self.entry_sender_name.delete(0, tk.END)
            self.entry_sender_name.insert(0, cuenta.get('sender_name', ''))
            self.log(f"✓ Cuenta seleccionada: {nombre}")
    
    def eliminar_cuenta(self):
        nombre = self.combo_cuentas.get()
        if not nombre:
            messagebox.showwarning("Falta", "Seleccione una cuenta del dropdown")
            return
        
        confirmar = messagebox.askyesno("Eliminar", f"¿Eliminar la cuenta '{nombre}'?")
        if not confirmar:
            return
        
        eliminar_cuenta(nombre)
        self.entry_email.delete(0, tk.END)
        self.entry_nombre_cuenta.delete(0, tk.END)
        self.entry_sender_name.delete(0, tk.END)
        
        cuentas = lista_cuentas()
        self.combo_cuentas['values'] = cuentas
        if cuentas:
            self.combo_cuentas.current(0)
        
        messagebox.showinfo("Eliminado", f"Cuenta '{nombre}' eliminada")
        self.log(f"✓ Cuenta eliminada: {nombre}")
    
    def cargar_datos(self):
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
        self.template_seleccionado = self.combo_template.get()
        
        if self.template_seleccionado and self.datos:
            address = self.datos[0].get('address', '')
            subject_auto = obtener_subject_template(self.template_seleccionado, address)
            self.lbl_subject.config(text=f"Subject: {subject_auto}", fg=self.color_primary)
        
        if self.template_seleccionado:
            self.log(f"✓ Template: {self.template_seleccionado}")
            self.actualizar_resumen()
    
    def ver_preview(self):
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
            preview_html = aplicar_variables(template, ejemplo)
            
            cuenta = obtener_cuenta_activa()
            sender = cuenta.get('sender_name', cuenta.get('email', 'genesis@engsv.com')) if cuenta else 'genesis@engsv.com'
            subject = obtener_subject_template(self.template_seleccionado, ejemplo['Property Address'])
            
            import tempfile, webbrowser
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(preview_html)
                webbrowser.open('file://' + f.name)
            
            self.log(f"Preview abierta en navegador")
            
        except Exception as e:
            messagebox.showerror("Error", str(e))
    
    def abrir_en_navegador(self, html):
        import tempfile
        import webbrowser
        with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
            f.write(html)
            webbrowser.open(f.name)
    
    def seleccionar_adjunto(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar adjunto",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")]
        )
        
        if archivo:
            self.adjunto = archivo
            self.lbl_adj.config(text=os.path.basename(archivo), fg="black")
            self.log(f"✓ Adjunto: {os.path.basename(archivo)}")
    
    def actualizar_resumen(self):
        if self.datos and self.stats and self.template_seleccionado:
            cuenta = self.combo_cuentas.get()
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
    
    def enviar_emails(self):
        if not self.datos:
            messagebox.showwarning("Warning", "Cargue un archivo de datos")
            return
        
        if not self.template_seleccionado:
            messagebox.showwarning("Warning", "Seleccione un template")
            return
        
        cuenta = obtener_cuenta_activa()
        if not cuenta:
            messagebox.showwarning("Warning", "Seleccione una cuenta Gmail")
            return
        
        cantidad = len(self.datos)
        confirmar = messagebox.askyesno("Confirmar", f"¿Enviar {cantidad} emails?")
        if not confirmar:
            return
        
        try:
            template = obtener_template(self.template_seleccionado)
        except Exception as e:
            messagebox.showerror("Error", f"Template: {e}")
            return
        
        adjunto = self.adjunto if self.adjunto else None
        
        subject = obtener_subject_template(self.template_seleccionado, self.datos[0].get('address', ''))
        
        self.btn_enviar.config(state="disabled")
        self.log("=" * 40)
        self.log(f"Enviando {cantidad} emails...")
        
        Gmail = GmailBorrador(cuenta['email'], cuenta['app_password'])
        
        def proceso():
            try:
                sender_name = cuenta.get('sender_name', '')
                
                stats = Gmail.crear_borradores(
                    datos=self.datos,
                    template=template,
                    subject=subject,
                    adjunto=adjunto,
                    callback=self.log,
                    sender_name=sender_name
                )
                
                self.root.after(0, lambda: messagebox.showinfo(
                    "Completado",
                    f"OK: {stats['creados']} | Errores: {stats['fallidos']}"
                ))
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            
            self.root.after(0, lambda: self.btn_enviar.config(state="normal"))
        
        Thread(target=proceso, daemon=True).start()


def main():
    root = tk.Tk()
    app = EmailMasivoApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()