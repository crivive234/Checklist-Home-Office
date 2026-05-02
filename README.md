# Checklist OTD Américas — Guía de Uso

Herramienta web para validar, documentar y certificar equipos de cómputo (laptops y desktops) destinados a trabajo en casa o procesos de inventario. Genera un reporte PDF al finalizar.

---

## Archivos incluidos

| Archivo | Descripción |
|---|---|
| `checklist-otd.html` | La herramienta principal. Se abre en cualquier navegador web (Chrome recomendado). |
| `Recopilar-DatosEquipo.ps1` | Script de PowerShell que recopila la información técnica del equipo automáticamente. |

> Ambos archivos deben estar en la **misma carpeta** para que el lanzador funcione correctamente.

---

## ¿Cómo se usa?

### Paso 1 — Seleccionar el tipo de checklist

Al abrir el HTML, lo primero que verás es un selector con tres opciones:

| Tipo | Cuándo usarlo |
|---|---|
| **Alistamiento Entrega de equipo Home Office** | Cuando se va a entregar un equipo a un colaborador para trabajo remoto. Incluye sección de firmas. |
| **Paz y Salvo (Entrega de equipo)** | Cuando el colaborador devuelve el equipo. Incluye sección de firmas. |
| **Inventario (Certificación)** | Para registro y certificación interna de inventario. **No incluye** sección de firmas ni datos de entrega. |

---

### Paso 2 — Recopilar datos del equipo automáticamente

La herramienta puede llenarse automáticamente con la información del equipo usando el script de PowerShell.

1. Haz clic en **"⬇ Descargar lanzador"** — se descargará un archivo llamado `lanzar-script.bat`.
2. Mueve ese archivo `.bat` a la misma carpeta donde está el script `Recopilar-DatosEquipo.ps1`.
3. Haz doble clic en `lanzar-script.bat`. Se abrirá una ventana negra (consola) y el script se ejecutará.
4. Cuando termine, aparecerá un archivo `.json` en esa misma carpeta.
5. De vuelta en el checklist, haz clic en **"CARGAR JSON"** y selecciona ese archivo.

Los campos del equipo se llenarán solos: hostname, marca, modelo, procesador, RAM, discos, sistema operativo, drivers, programas instalados y datos de red.

> **¿Sin acceso para ejecutar el script?** Puedes llenar los campos manualmente uno por uno.

---

### Paso 3 — Revisar y completar cada sección

El checklist está dividido en 5 secciones. Recórrelas de arriba a abajo:

---

#### Sección 1 — Información del Equipo

Muestra los datos técnicos del equipo: hostname, fecha, marca, modelo, número de serie, sistema operativo, procesador, RAM, BIOS y discos.

- Los campos en **gris** (solo lectura) se llenan automáticamente desde el JSON.
- La **fecha de alistamiento** se puede editar manualmente.
- Se incluyen preguntas adicionales sobre **Windows Update**, **licencia de Windows** y **actualizaciones pendientes**, que deben responderse manualmente con los botones Sí / No.

---

#### Sección 2 — Drivers Instalados

Muestra la lista de drivers detectados en el equipo.

**Auto-selección automática:** Al cargar el JSON, el sistema marca automáticamente los 5 drivers requeridos:

| Driver | Cómo se identifica |
|---|---|
| Red (PCIe) | Clase NET + nombre contiene "PCIe" |
| Video (Graphics) | Clase DISPLAY + nombre contiene "Graphics" |
| Audio (Realtek) | Nombre contiene "Realtek" y "Audio" |
| Mouse | Nombre contiene "mouse" |
| Teclado | Nombre contiene "keyboard" |

> Los drivers de software como *Intel Graphics Command Center* o *Intel Graphics Control Panel* **no se seleccionan**, aunque contengan "Graphics".

Puedes marcar o desmarcar drivers adicionales usando las casillas de verificación. Solo los drivers marcados aparecerán en el reporte PDF.

Hay un campo de **observaciones** al final para anotar cualquier driver faltante o con problemas.

---

#### Sección 3 — Verificación Física de Periféricos

Se verifica el estado físico del equipo y sus accesorios.

**Limpieza del equipo:** Indica si el equipo fue limpiado internamente, con un campo de observaciones.

**Configuración de pantallas:** Usa los botones para indicar si el puesto tiene 1 o 2 monitores:

- **⬜ 1 Pantalla** — Solo se muestran los campos del monitor principal. (opción por defecto)
- **⬜⬜ 2 Pantallas** — Se activan campos adicionales para el segundo monitor: Monitor 2ª Pantalla, Cable de Video 2ª Pantalla y Cable Poder Monitor 2ª Pantalla.

> Los campos de la 2ª pantalla aparecen desactivados (en gris) hasta que se seleccione esa opción, y **solo se incluyen en el PDF si están activos**.

**Periféricos:** Para cada elemento (teclado, mouse, monitor, cámara, diadema, YubiKey) se indica si funciona con los botones **Sí / No / N/A** y se puede agregar una observación.

**Cables:** Se registra el tipo y estado de cada cable incluido:
- Cable de Video (con selector de tipo: HDMI, DisplayPort, VGA, etc.)
- Cable de Red UTP (con selector Cat 5 / Cat 6, etc.)
- Cable Poder Equipo
- Cable Poder Monitor

---

#### Sección 4 — Conectividad y Red

Muestra las tarjetas de red detectadas en el equipo (equivalente a ver `ncpa.cpl`) con nombre, descripción, MAC, IP, estado y velocidad.

También muestra el resultado del **Ping a 8.8.8.8** con tiempo de respuesta y pérdida de paquetes.

Esta información se llena automáticamente desde el JSON.

---

#### Sección 5 — Información de Entrega y Firmas

> ⚠️ **Esta sección no aparece cuando el tipo de checklist es "Inventario (Certificación)".**

Se registran los datos del colaborador que recibe el equipo y del técnico responsable:

- Nombre completo, cédula, cargo, teléfono y campaña/área del colaborador.
- Nombre y cargo del técnico.

**Firma digital:** Cada parte puede firmar directamente en el recuadro usando el mouse o el dedo (en pantallas táctiles). Si se prefiere firmar en papel, se deja en blanco y el PDF generará un espacio con línea para firma manual.

El botón **✕ Limpiar firma** borra la firma del recuadro correspondiente.

---

### Paso 4 — Generar el PDF

Al terminar de revisar todas las secciones, haz clic en el botón negro **"GENERAR PDF"** al final de la página.

El archivo se descargará automáticamente con el nombre: `Checklist_[Hostname]_[Fecha].pdf`

El PDF incluye todas las secciones activas y solo los datos que se hayan completado. Los campos vacíos se muestran como `—` o como líneas para llenar a mano.

---

## Preguntas frecuentes

**¿Necesita internet para funcionar?**
No. Una vez abierto el archivo HTML, todo funciona sin conexión a internet.

**¿Puedo usarlo en Mac o Linux?**
El checklist HTML funciona en cualquier sistema. El script de PowerShell (`Recopilar-DatosEquipo.ps1`) es exclusivo para **Windows**.

**¿Qué hago si el script no se ejecuta?**
Asegúrate de que el archivo `.bat` esté en la misma carpeta que el script `.ps1`. Si hay restricciones de seguridad en la empresa, es posible que necesites ejecutar el script manualmente desde PowerShell con permisos de administrador.

**¿Puedo llenar el checklist sin el JSON?**
Sí. Todos los campos se pueden completar manualmente. El JSON solo agiliza el proceso.

**¿Los datos se guardan si cierro la página?**
No. Al cerrar o recargar la página se pierden los datos. Genera el PDF antes de cerrar.

---

## Requisitos

- Navegador web moderno (Google Chrome recomendado)
- Windows 10 / 11 para ejecutar el script de PowerShell
- PowerShell 5.1 o superior (incluido por defecto en Windows 10/11)

---

*OTD Américas — Herramienta interna de alistamiento y certificación de equipos*
