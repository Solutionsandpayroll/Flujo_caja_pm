# Flujo de Caja - Punto Medical

Editor web de flujo de caja para **Punto Medical** desarrollado por **Solutions & Payroll (S&P)**. Permite abrir, editar, y gestionar archivos Excel de flujo de caja directamente desde el navegador, integrado con Google Sheets para importación automática de facturas.

---

## Funcionalidades

### Gestión del Flujo de Caja
- **Abrir Excel local** desde el navegador usando File System Access API (sin subir nada a la nube)
- **Guardar directamente** al mismo archivo en disco
- **Reconectar** automáticamente archivos abiertos previamente (persistencia en IndexedDB)
- **Vista estructurada por mes**: cada hoja se divide en:
  - **Sección Inicial**: Bancos y Clientes con TOTAL INGRESOS como resumen
  - **Cuentas por Pagar (CXP)**: dropdown de subsecciones con tabla filtrable de registros
- **Edición inline de celdas**: texto, números (COP con separadores de miles), fechas, estados con badges de color
- **Inserción de filas** en cualquier sección/subsección
- **Formato preservado**: colores, estilos, fórmulas, celdas fusionadas

### Sincronización con Google Sheets
- **Botón "Actualizar registros"**: consulta Google Sheets vía Apps Script y detecta facturas nuevas
- **Inserción automática** de facturas nuevas en la subsección "COMPRAS CON FACTURAS"
- **Filtro por mes visible**: solo sincroniza el mes que el usuario está viendo
- **Idempotente**: ejecutar múltiples veces no duplica registros
- **Identificación por N° factura** para evitar duplicados
- **Cálculo de vencimiento**: `fechaFact + PLAZO` (desde hoja Descuentos en Google Sheets)

### Abonos
- **Botón `$` en cada fila CXP**: abre modal para registrar un abono con fecha, valor y método de pago
- **Valor efectivo** mostrado en tiempo real (original - abonos)
- **Filas resaltadas en naranja** cuando tienen abonos
- **Hoja oculta `ABONOS`** en el Excel con columnas: Mes, Proveedor, Factura, Fecha, Valor Antes, Abono, Resultante, Método
- **Pestaña "Abonos"** en la web con historial completo del mes
- **Bloqueo de edición**: registros con abonos no permiten modificar su valor

### Modal de Cancelado
- Al cambiar un registro a **"Cancelado"**, aparece una modal pidiendo el método de pago (texto libre)
- Se guarda en la columna **DEUDA PENDIENTE**

### Crear Mes Siguiente
- **Botón "Crear mes siguiente"**: clona la hoja del mes actual preservando toda la estructura y formato
- **Filtrado inteligente**:
  - Filas "Cancelado" NO se pasan (excepto conceptos recurrentes)
  - Conceptos recurrentes (seguridad social, arriendos, nómina, etc.) siempre se mantienen con estado "Sin Pagar"
  - **Valores originales restaurados** para filas con abonos
- **Periodicidad de fechas**: 8 conceptos actualizan su fecha de vencimiento automáticamente:
  - IVA: +4 meses (cuatrimestral)
  - Industria y Comercio: +2 meses (bimestral)
  - Poliza de seguro PYME: +12 meses (anual)
  - Acueducto Bodega: +2 meses (bimestral)
  - HOSTING MANTIS: +12 meses (anual)
  - Dotación: +4 meses + rotación de nombre (Primera → Segunda → Tercera → siguiente año)
  - CESANTIAS: +12 meses (anual)
  - Intereses a las cesantias: +12 meses (anual)
- **Actualización de Nómina quincena**: quincena 2 = mes anterior, quincena 1 = mes actual

### Generar Resumen
- **Botón "Generar Resumen"**: descarga un Excel con los registros "Sin Pagar" del mes visible
- Columnas: Tipo (PROVEDOR/GASTOS), Nombre, Valor, Fecha, DESCUENTO
- **Fórmula VLOOKUP** para categorizar como OBLIGATORIO, MANEJABLE o IMPORTANTE
- **Hoja oculta `MAPEO_DESCUENTOS`** guarda las categorías asignadas por el usuario para no repetir preguntas
- **Modal de mapeo**: si un concepto no está categorizado, pregunta y guarda la respuesta

---

## Estructura del Proyecto

```
Flujo de caja - PM/
├── public/
│   ├── Logo syp.png
│   └── FLUJO DE CAJA RESUMEN.xlsx    # Plantilla para "Generar Resumen"
├── src/
│   ├── main.jsx                      # Entry point
│   ├── App.jsx                       # Componente raíz (header, footer, layout)
│   ├── App.css                       # Estilos (2500+ líneas)
│   ├── index.css                     # Reset global
│   ├── components/
│   │   ├── ExcelEditor.jsx           # Orquestador: slots, archivos, abonos, sync, resumen, modales
│   │   ├── MonthViewer.jsx           # Vista de mes (Sección Inicial + CXP)
│   │   ├── SectionInitial.jsx        # Sección Inicial (Bancos, Clientes, TOTAL INGRESOS)
│   │   ├── SectionCXP.jsx            # Cuentas por Pagar (subsecciones + abonos)
│   │   └── EditableCell.jsx          # Celda editable inline (texto, número, select, badge)
│   └── utils/
│       ├── excelParser.js            # Parseo de hojas de Excel a estructura de datos
│       ├── xlsxPatcher.js            # Parcheo quirúrgico de XML del .xlsx (lectura/escritura/inserción)
│       ├── sheetsSync.js             # Sincronización con Google Sheets vía Apps Script
│       ├── abonosStore.js            # Lectura/escritura de abonos desde hoja ABONOS
│       └── fileHandleStore.js        # Persistencia de FileSystemFileHandle en IndexedDB
├── package.json
├── vite.config.js
└── index.html
```

---

## Tecnologías

| Tecnología | Uso |
|---|---|
| **React 18** | Framework UI |
| **Vite 5** | Build tool y dev server |
| **XLSX (SheetJS) 0.18.5** | Lectura y escritura de archivos Excel |
| **JSZip 3.10.1** | Manipulación quirúrgica del ZIP interno de .xlsx |
| **File System Access API** | Abrir/guardar archivos del disco local (Chrome/Edge) |
| **IndexedDB** | Persistencia de handles de archivo |
| **Google Apps Script** | API para consultar Google Sheets |
| **CSS3 puro** | Estilos sin frameworks externos |

---

## Configuración

### Apps Script (Google Sheets)

1. Abrir la Google Sheet de facturas
2. Extensiones → Apps Script
3. Pegar el código del archivo `google-apps-script.js` (debe incluir `doGet` con `listarFacturas`, `listarFacturasPorMes`, `formatearFecha`, `perteneceAlMes`)
4. Implementar → Nueva implementación → Aplicación web → "Cualquier persona"
5. Copiar la URL generada
6. Pegarla en `src/utils/sheetsSync.js` línea 3 (`APPS_SCRIPT_URL`)

### Plantilla de Resumen

Colocar `FLUJO DE CAJA RESUMEN.xlsx` en la carpeta `public/` del proyecto. Este archivo contiene la Hoja2 con el mapeo de DESCUENTO usado por la fórmula VLOOKUP.

---

## Instalación y Uso

```bash
# Instalar dependencias
npm install

# Desarrollo
npm run dev

# Producción
npm run build
npm run preview
```

---

## Notas

- **Solo funciona en Chrome/Edge** (requiere File System Access API)
- El archivo Excel **nunca se sube a ningún servidor** — todo se procesa localmente
- La primera sincronización con Google Sheets puede tardar unos segundos (cold start de Apps Script)
- Las hojas ocultas (`ABONOS`, `MAPEO_DESCUENTOS`) se crean automáticamente al necesitarse

---

© 2026 Solutions & Payroll. Todos los derechos reservados.
