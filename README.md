# Email Masivo Sender

Aplicación de escritorio para envío masivo de emails personalizados desde Gmail usando templates HTML.

## Características

- 📁 Carga archivos XLSX/CSV con datos (email, folio, address)
- 📧 5 templates HTML profesionales incluidos
- 👤 Múltiples cuentas Gmail configurables (agregar/eliminar)
- 📎 Adjuntar archivos a los emails (template 3)
- 🎨 Asunto automático según template seleccionado
- 💾 Credenciales guardadas entre sesiones
- 🖥 Interfaz gráfica simple para Windows

## Requisitos

- Windows 10/11
- Python 3.9+
- Cuenta Gmail con **App Password** (no tu contraseña normal)

## Instalación

1. **Clonar o descargar** el proyecto
2. **Instalar dependencias**:
   ```cmd
   pip install -r requirements.txt
   ```
3. **Ejecutar**:
   ```cmd
   python app.py
   ```
   O haz doble click en `Ejecutar_app.bat`

## Cómo obtener App Password de Gmail

1. Ve a https://myaccount.google.com → Seguridad
2. Activa **Verificación en 2 pasos**
3. Busca **Contraseñas de aplicaciones**
4. Genera una (16 caracteres)
5. Usa esa contraseña en la app

## Uso

1. **Agregar cuenta**: 
   - Escribí un nombre (ej: "Trabajo")
   - Escribí tu email Gmail
   - Escribí el App Password
   - Click en **+ Agregar**

2. **Seleccionar cuenta**: 
   - Elegí del dropdown
   - Click en **Usar**

3. **Cargar datos**: Click "Cargar XLSX/CSV"

4. **Seleccionar template**: Elegí 1-5

5. **Adjunto** (solo template 3): Click "Seleccionar PDF"

6. **Enviar**: Click "ENVIAR EMAILS"

## Eliminar cuenta

1. Seleccioná la cuenta del dropdown
2. Click en **🗑 Eliminar**

## Estructura

```
EmailMasivo/
├── app.py              # Aplicación principal
├── src/
│   ├── auth.py        # Gestión de cuentas
│   ├── reader.py     # Lectura de archivos
│   ├── templates.py # Gestión de templates
│   └── gmail_draft.py # Envío de emails
├── templates/         # 5 templates HTML
├── data/             # Archivos de datos
├── config.json       # Cuentas guardadas
└── requirements.txt # Dependencias
```

## Templates

| # | Template | Subject |
|---|---------|---------|
| 1 | Inquiry | Inquiry about {address} Recertification Status |
| 2 | Warning | Why you should not wait on the Recertification of {address} |
| 3 | Brochure | How to make the Recertification of {address} easier |
| 4 | Consultant | When selecting a consultant for {address} Recertification |
| 5 | Follow-up | {address} - Should I close your file? |

## Licencia

MIT License