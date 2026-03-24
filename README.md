# 🖥️ Checklist Home Office — OTD Américas

Herramienta para la **validación técnica de equipos de trabajo remoto**, combinando una página web interactiva y un script de PowerShell para recopilar información automáticamente.

---

## 📌 ¿Qué hace este proyecto?

Permite:

- 📊 Recopilar información del equipo automáticamente (hardware, SO, red, programas, drivers)
- 📥 Importar esos datos en una interfaz web
- ✅ Completar un checklist técnico (periféricos, cables, estado del equipo)
- 📄 Generar un **reporte PDF profesional**

---

## 🧩 Estructura del proyecto

-📁 HomeOffice-Checklist
- checklist-otd.html # Interfaz principal (web)
- Recopilar-DatosEquipo.ps1 # Script de recolección automática
- README.md

---

## ⚙️ ¿Cómo usarlo?

### 1. Preparar entorno
Coloca ambos archivos en la misma carpeta:
- `checklist-otd.html`
- `Recopilar-DatosEquipo.ps1`

---

### 2. Ejecutar script

Desde la página:

1. Haz clic en **"Descargar lanzador"**
2. Ejecuta `lanzar-script.bat`
3. Se generará un archivo `.json` con la información del equipo

---

### 3. Cargar datos

1. En la página, haz clic en **"Cargar JSON"**
2. Selecciona el archivo generado
3. El formulario se autocompleta automáticamente

---

### 4. Completar checklist

- Validar periféricos
- Revisar cables
- Agregar observaciones

---

### 5. Generar reporte

Haz clic en: 
GENERAR REPORTE PDF

---

## 🛠️ Tecnologías usadas

- HTML + CSS + JavaScript
- PowerShell
- jsPDF (generación de PDF)

---

## ⚠️ Requisitos

- Windows (para ejecutar PowerShell)
- Permisos para ejecutar scripts (`ExecutionPolicy`)

Si tienes restricciones, el lanzador usa:

```powershell
ExecutionPolicy Bypass
🐛 Problemas conocidos
Error con [datetime]::TryParse

En algunas versiones de PowerShell aparece:

No se encuentra ninguna sobrecarga para "TryParse"
✅ Solución

Reemplazar:

[datetime]::TryParse($biosRaw, [ref]$parsed)

Por:
try {
    $parsed = [datetime]$biosRaw
} catch {
    $parsed = $null
}
🚀 Mejoras futuras
Exportación a Excel
Validaciones automáticas de requisitos mínimos
Integración con inventarios (CMDB)
Firma digital del reporte
👨‍💻 Autor

Proyecto desarrollado para procesos de alistamiento técnico de equipos Home Office.
