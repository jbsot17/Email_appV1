"""Email Masivo Sender - CustomTkinter Interface"""
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from threading import Thread
import customtkinter as ctk

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

sys.path.insert(0, str(Path(__file__).parent))

from src.reader import leer_archivo_datos
from src.auth import (
    agregar_cuenta, lista_cuentas, seleccionar_cuenta,
    obtener_cuenta_activa, obtener_cuenta_por_nombre, eliminar_cuenta
)
from src.templates import listar_templates, obtener_template, obtener_subject_template
from src.gmail_draft import GmailBorrador


class EmailMasivoApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Email Masivo - Enviador")
        self.root.geometry("800x750")
        
        self.color_primary = "#3B8ED0"
        self.color_success = "#2CC985"
        self.color_danger = "#FF5C5C"
        
        self.datos = None
        self.stats = None
        self.template_seleccionado = None
        self.adjunto = None
        self.log_text = None
        self.log = lambda m: None
        
        self.crear_interfaz()

    def crear_interfaz(self):
        ctk.CTkLabel(self.root, text="Email Masivo", font=("Roboto", 28, "bold")).pack(pady=15)
        ctk.CTkLabel(self.root, text="Enviador de Emails", font=("Roboto", 14)).pack()
        
        # Cuenta
        f1 = ctk.CTkFrame(self.root)
        f1.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(f1, text="1. Cuenta Gmail", font=("Roboto", 16, "bold")).pack(anchor="w", padx=10, pady=8)
        
        self.combo_cuentas = ctk.CTkComboBox(f1, values=lista_cuentas(), width=220)
        self.combo_cuentas.pack(padx=10, pady=5)
        
        f1a = ctk.CTkFrame(f1, fg_color="transparent")
        f1a.pack()
        ctk.CTkButton(f1a, text="Usar", command=self.seleccionar_cuenta_usar, width=80).pack(side="left", padx=2)
        ctk.CTkButton(f1a, text="Eliminar", command=self.eliminar_cuenta, fg_color=self.color_danger, width=80).pack(side="left", padx=2)
        
        f1b = ctk.CTkFrame(f1, fg_color="transparent")
        f1b.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(f1b, text="Nombre:").pack(side="left")
        self.entry_nombre = ctk.CTkEntry(f1b, width=100)
        self.entry_nombre.pack(side="left", padx=5)
        ctk.CTkLabel(f1b, text="Email:").pack(side="left", padx=5)
        self.entry_email = ctk.CTkEntry(f1b, width=150)
        self.entry_email.pack(side="left", padx=5)
        ctk.CTkLabel(f1b, text="Pass:").pack(side="left", padx=5)
        self.entry_password = ctk.CTkEntry(f1b, width=100, show="*")
        self.entry_password.pack(side="left", padx=5)
        ctk.CTkButton(f1b, text="+ Agregar", command=self.agregar_cuenta, fg_color=self.color_success).pack(side="left", padx=5)
        
        f1c = ctk.CTkFrame(f1, fg_color="transparent")
        f1c.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(f1c, text="Remitente:").pack(side="left")
        self.entry_sender = ctk.CTkEntry(f1c, width=200)
        self.entry_sender.pack(side="left", padx=5)
        
        # Archivo
        f2 = ctk.CTkFrame(self.root)
        f2.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(f2, text="2. Archivo", font=("Roboto", 16, "bold")).pack(anchor="w", padx=10, pady=8)
        ctk.CTkButton(f2, text="Cargar XLSX/CSV", command=self.cargar_datos, fg_color=self.color_success).pack(anchor="w", padx=10)
        self.lbl_datos = ctk.CTkLabel(f2, text="Sin archivo")
        self.lbl_datos.pack(anchor="w", padx=10, pady=5)
        
        # Template
        f3 = ctk.CTkFrame(self.root)
        f3.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(f3, text="3. Template", font=("Roboto", 16, "bold")).pack(anchor="w", padx=10, pady=8)
        
        templates = listar_templates()
        self.combo_template = ctk.CTkComboBox(f3, values=templates, width=280)
        self.combo_template.pack(anchor="w", padx=10, pady=5)
        self.combo_template.configure(command=lambda x: self.seleccionar_template())
        
        f3a = ctk.CTkFrame(f3, fg_color="transparent")
        f3a.pack(anchor="w", padx=10, pady=5)
        ctk.CTkButton(f3a, text="Ver Preview", command=self.ver_preview).pack(side="left", padx=2)
        self.lbl_subject = ctk.CTkLabel(f3, text="")
        self.lbl_subject.pack(anchor="w", padx=10)
        
        # Adjunto
        f4 = ctk.CTkFrame(self.root)
        f4.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(f4, text="4. Adjunto", font=("Roboto", 16, "bold")).pack(anchor="w", padx=10, pady=8)
        ctk.CTkButton(f4, text="Seleccionar PDF", command=self.seleccionar_adjunto).pack(anchor="w", padx=10)
        self.lbl_adj = ctk.CTkLabel(f4, text="Ninguno")
        self.lbl_adj.pack(anchor="w", padx=10, pady=5)
        
        self.lbl_resumen = ctk.CTkLabel(self.root, text="Sin datos", font=("Roboto", 14))
        self.lbl_resumen.pack(pady=10)
        
        self.btn_enviar = ctk.CTkButton(self.root, text="ENVIAR EMAILS", command=self.enviar_emails, fg_color=self.color_danger)
        self.btn_enviar.pack(fill="x", padx=20, pady=15)
        
        # Log
        f5 = ctk.CTkFrame(self.root)
        f5.pack(fill="both", expand=True, padx=20, pady=10)
        ctk.CTkLabel(f5, text="Log", font=("Roboto", 12, "bold")).pack(anchor="w")
        self.log_text = ctk.CTkTextbox(f5, height=120)
        self.log_text.pack(fill="both", expand=True, pady=5)
        self.log = lambda m: (self.log_text.insert("end", m + "\n"), self.log_text.see("end"))
        
        self.cargar_config()
        self.root.mainloop()

    def cargar_config(self):
        cuentas = lista_cuentas()
        self.combo_cuentas.configure(values=cuentas)
        if cuentas:
            self.combo_cuentas.set(cuentas[0])
            cuenta = obtener_cuenta_activa()
            if cuenta:
                self.entry_email.insert(0, cuenta.get('email', ''))
                self.entry_sender.insert(0, cuenta.get('sender_name', ''))
                self.log(f"Cuenta: {cuenta.get('nombre')}")

    def agregar_cuenta(self):
        nombre = self.entry_nombre.get().strip()
        email = self.entry_email.get().strip()
        password = self.entry_password.get().strip()
        sender = self.entry_sender.get().strip()
        if not nombre or not email or not password:
            messagebox.showwarning("Falta", "Datos incompletos")
            return
        agregar_cuenta(nombre, email, password, sender)
        self.entry_nombre.delete(0, tk.END)
        self.entry_email.delete(0, tk.END)
        self.entry_password.delete(0, tk.END)
        self.entry_sender.delete(0, tk.END)
        self.combo_cuentas.configure(values=lista_cuentas())
        self.log(f"Cuenta '{nombre}' guardada")

    def seleccionar_cuenta_usar(self):
        nombre = self.combo_cuentas.get()
        if not nombre:
            return
        cuenta = obtener_cuenta_por_nombre(nombre)
        if cuenta:
            seleccionar_cuenta(nombre)
            self.entry_email.delete(0, tk.END)
            self.entry_email.insert(0, cuenta.get('email', ''))
            self.entry_sender.delete(0, tk.END)
            self.entry_sender.insert(0, cuenta.get('sender_name', ''))
            self.log(f"Cuenta: {nombre}")

    def eliminar_cuenta(self):
        nombre = self.combo_cuentas.get()
        if not nombre:
            return
        if not messagebox.askyesno("Eliminar", f"¿Eliminar '{nombre}'?"):
            return
        eliminar_cuenta(nombre)
        self.entry_email.delete(0, tk.END)
        self.entry_nombre.delete(0, tk.END)
        self.entry_sender.delete(0, tk.END)
        self.combo_cuentas.configure(values=lista_cuentas())
        self.log(f"Cuenta eliminada")

    def cargar_datos(self):
        archivo = filedialog.askopenfilename(title="Archivo", filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv")])
        if archivo:
            try:
                self.datos, self.stats = leer_archivo_datos(archivo)
                self.lbl_datos.configure(text=f"{self.stats['total']} registros")
                self.log(f"Archivo: {self.stats['total']} registros")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def seleccionar_template(self):
        self.template_seleccionado = self.combo_template.get()
        if self.template_seleccionado:
            if self.datos:
                subject = obtener_subject_template(self.template_seleccionado, self.datos[0].get('address', ''))
                self.lbl_subject.configure(text=f"Subject: {subject}")
            self.log(f"Template: {self.template_seleccionado}")

    def ver_preview(self):
        if not self.template_seleccionado:
            messagebox.showwarning("!", "Seleccione un template")
            return
        try:
            from src.templates import aplicar_variables
            html = aplicar_variables(
                obtener_template(self.template_seleccionado),
                {'Folio Number': 'EX', 'Property Address': '123 St'}
            )
            import tempfile, webbrowser
            with tempfile.NamedTemporaryFile(suffix='.html', delete=False, mode='w', encoding='utf-8') as f:
                f.write(html)
                webbrowser.open(f.name)
            self.log("Preview abierta en navegador")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def seleccionar_adjunto(self):
        archivo = filedialog.askopenfilename(title="Adjunto", filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")])
        if archivo:
            self.adjunto = archivo
            self.lbl_adj.configure(text=os.path.basename(archivo))
            self.log(f"Adjunto: {os.path.basename(archivo)}")

    def enviar_emails(self):
        if not self.datos or not self.template_seleccionado or not obtener_cuenta_activa():
            messagebox.showwarning("!", "Complete todos los datos")
            return
        if not messagebox.askyesno("Confirmar", f"¿Enviar {len(self.datos)} emails?"):
            return
        try:
            template = obtener_template(self.template_seleccionado)
            subject = obtener_subject_template(self.template_seleccionado, self.datos[0].get('address', ''))
            cuenta = obtener_cuenta_activa()
            Gmail = GmailBorrador(cuenta['email'], cuenta['app_password'])
            self.btn_enviar.configure(state="disabled")
            self.log(f"Enviando {len(self.datos)}...")

            def proceso():
                try:
                    stats = Gmail.crear_borradores(
                        datos=self.datos,
                        template=template,
                        subject=subject,
                        adjunto=self.adjunto,
                        sender_name=cuenta.get('sender_name', ''),
                        callback=self.log
                    )
                    self.root.after(0, lambda: messagebox.showinfo("Completado", f"OK: {stats['creados']}"))
                except Exception as ex:
                    self.root.after(0, lambda: messagebox.showerror("Error", str(ex)))
                self.root.after(0, lambda: self.btn_enviar.configure(state="normal"))

            Thread(target=proceso, daemon=True).start()
        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == '__main__':
    EmailMasivoApp()