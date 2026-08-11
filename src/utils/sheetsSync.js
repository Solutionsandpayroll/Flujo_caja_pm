import { parseMonthSheet, MONTHS } from './excelParser'

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyKCWmQf8NRzQ2oqsGVov7N0J_i37K72jWJz8Pc4Ex4fb6zY6fJWtX1B5SrsEBVckaKKg/exec'

const TARGET_SUBSECTION = 'COMPRAS CON FACTURAS'

export function getMonthKey(sheetName) {
  const idx = MONTHS.findIndex(m =>
    sheetName.toUpperCase() === m ||
    sheetName.toUpperCase().includes(m)
  )
  if (idx === -1) return null
  const year = new Date().getFullYear()
  const month = String(idx + 1).padStart(2, '0')
  return `${year}-${month}`
}

export function buildSheetsUrl(mes) {
  const base = `${APPS_SCRIPT_URL}?action=listarFacturas`
  return mes ? `${base}&mes=${mes}` : base
}

export async function fetchFacturasFromSheets(mes) {
  const url = buildSheetsUrl(mes)
  const response = await fetch(url)
  if (!response.ok) {
    throw new Error(`Error HTTP ${response.status} al consultar Google Sheets`)
  }
  const data = await response.json()
  if (!data.success) {
    throw new Error(data.error || 'Respuesta inesperada de Google Sheets')
  }
  return data.facturas || []
}

function parseFechaGoogleSheets(val) {
  if (!val) return null

  if (val instanceof Date || (typeof val === 'object' && val.toISOString)) {
    return val
  }

  if (typeof val === 'string') {
    const isoMatch = /^(\d{4})-(\d{2})-(\d{2})/.exec(val)
    if (isoMatch) {
      return new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]))
    }
    const dmyMatch = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/.exec(val)
    if (dmyMatch) {
      return new Date(Number(dmyMatch[3]), Number(dmyMatch[2]) - 1, Number(dmyMatch[1]))
    }
    const num = parseFloat(val)
    if (!isNaN(num) && num > 40000 && num < 60000) {
      const d = new Date((num - 25569) * 86400 * 1000)
      if (!isNaN(d.getTime())) return d
    }
    return null
  }

  if (typeof val === 'number' && val > 40000 && val < 60000) {
    const utcMs = (val - 25569) * 86400 * 1000
    const utcDate = new Date(utcMs)
    if (!isNaN(utcDate.getTime())) {
      return new Date(utcDate.getUTCFullYear(), utcDate.getUTCMonth(), utcDate.getUTCDate())
    }
    return null
  }

  return null
}

function perteneceAlMes(fecha, mes) {
  const d = parseFechaGoogleSheets(fecha)
  if (!d) return false
  const año = d.getFullYear()
  const mesNum = d.getMonth() + 1
  const mesStr = año + '-' + String(mesNum).padStart(2, '0')
  return mesStr === mes
}

export function getExistingFacturaIds(rows) {
  const parsed = parseMonthSheet(rows)
  if (!parsed || !parsed.sectionCXP) return new Set()

  const subsection = parsed.sectionCXP.find(
    sub => sub.title.toUpperCase() === TARGET_SUBSECTION
  )
  if (!subsection) return new Set()

  return new Set(
    subsection.rows
      .filter(r => r.factura)
      .map(r => String(r.factura).trim())
  )
}

export function getSubsectionInsertRow(rows) {
  const parsed = parseMonthSheet(rows)
  if (!parsed || !parsed.sectionCXP) return -1

  const subsection = parsed.sectionCXP.find(
    sub => sub.title.toUpperCase() === TARGET_SUBSECTION
  )
  if (!subsection || subsection.rows.length === 0) return -1

  return subsection.rows[subsection.rows.length - 1]._row
}

function formatearFechaSheets(val) {
  if (!val) return ''
  const d = parseFechaGoogleSheets(val)
  if (d && !isNaN(d.getTime())) {
    const dia = String(d.getDate()).padStart(2, '0')
    const mes = String(d.getMonth() + 1).padStart(2, '0')
    const año = d.getFullYear()
    return `${dia}/${mes}/${año}`
  }
  return String(val)
}

export function buildSyncInsertions(newFacturas, insertAfterRow) {
  return newFacturas.map(factura => ({
    id: `sync-${factura.numFactura}-${Date.now()}-${Math.random().toString(36).slice(2, 5)}`,
    insertAfterRow,
    sectionKey: `cxp:${TARGET_SUBSECTION}`,
    cells: {
      0: 'Sin Pagar',
      1: '',
      2: String(factura.numFactura ?? '').trim(),
      3: String(factura.nombre ?? '').trim(),
      4: typeof factura.granTotal === 'number' ? factura.granTotal : (parseFloat(factura.granTotal) || 0),
      5: factura.fechaVencimiento || formatearFechaSheets(factura.fechaFact),
    }
  }))
}

export async function syncFacturas(rows, sheetName, existingInsertions = []) {
  const mes = getMonthKey(sheetName)
  if (!mes) return { inserted: 0, error: 'No se pudo determinar el mes de la hoja activa' }

  let facturasSheets = await fetchFacturasFromSheets()
  console.log(`[Sync] Google Sheets (sin filtro) devolvió ${facturasSheets.length} facturas`)

  if (!Array.isArray(facturasSheets) || facturasSheets.length === 0) {
    console.log(`[Sync] Reintentando con filtro de mes...`)
    facturasSheets = await fetchFacturasFromSheets(mes)
    console.log(`[Sync] Google Sheets (con filtro mes) devolvió ${facturasSheets.length} facturas`)
  }

  if (!Array.isArray(facturasSheets) || facturasSheets.length === 0) {
    return { inserted: 0, message: 'No hay facturas en Google Sheets' }
  }

  const filtradas = facturasSheets.filter(f => perteneceAlMes(f.fechaFact, mes))
  const fueraDeMes = facturasSheets.length - filtradas.length
  console.log(`[Sync] ${filtradas.length} pertenecen al mes ${mes}, ${fueraDeMes} fuera del mes (filtrado local)`)
  console.table(filtradas.map(f => ({ numFactura: f.numFactura, nombre: f.nombre, granTotal: f.granTotal, fechaFact: f.fechaFact, plazo: f.plazo, vencimiento: f.fechaVencimiento })))

  const existingIds = getExistingFacturaIds(rows)
  console.log(`[Sync] Facturas existentes en Excel (${TARGET_SUBSECTION}): ${existingIds.size}`)
  if (existingIds.size > 0) console.log([...existingIds])

  const pendingIds = new Set(
    existingInsertions
      .filter(ins => ins.sectionKey === `cxp:${TARGET_SUBSECTION}`)
      .map(ins => String(ins.cells[2] ?? '').trim())
      .filter(Boolean)
  )

  const sinNumFactura = []
  const yaEnExcel = []
  const yaPendientes = []
  const newFacturas = []

  for (const f of filtradas) {
    const id = String(f.numFactura ?? '').trim()
    if (!id) {
      sinNumFactura.push(f)
      continue
    }
    if (existingIds.has(id)) {
      yaEnExcel.push(id)
      continue
    }
    if (pendingIds.has(id)) {
      yaPendientes.push(id)
      continue
    }
    newFacturas.push(f)
  }

  console.log(`[Sync] Diagnóstico:`)
  console.log(`  - En Sheets (total):    ${facturasSheets.length}`)
  console.log(`  - Del mes ${mes}:       ${filtradas.length}`)
  console.log(`  - Fuera del mes:        ${fueraDeMes}`)
  console.log(`  - Sin numFactura:       ${sinNumFactura.length} ${sinNumFactura.length > 0 ? JSON.stringify(sinNumFactura.map(f => f.nombre)) : ''}`)
  console.log(`  - Ya en Excel:          ${yaEnExcel.length} ${yaEnExcel.length > 0 ? JSON.stringify(yaEnExcel) : ''}`)
  console.log(`  - Ya pendientes:        ${yaPendientes.length}`)
  console.log(`  - NUEVAS a insertar:    ${newFacturas.length}`)
  if (newFacturas.length > 0) console.table(newFacturas.map(f => ({ numFactura: f.numFactura, nombre: f.nombre, granTotal: f.granTotal, vencimiento: f.fechaVencimiento })))

  if (newFacturas.length === 0) {
    let msg = 'Todas las facturas de Google Sheets ya existen en el flujo de caja'
    if (sinNumFactura.length > 0) msg += ` (${sinNumFactura.length} sin numFactura)`
    if (yaEnExcel.length > 0) msg += ` (${yaEnExcel.length} ya en Excel)`
    return { inserted: 0, message: msg }
  }

  const insertRow = getSubsectionInsertRow(rows)
  if (insertRow === -1) {
    return { inserted: 0, error: `No se encontró la subsección "${TARGET_SUBSECTION}" en la hoja "${sheetName}"` }
  }

  const insertions = buildSyncInsertions(newFacturas, insertRow)

  let msg = `${newFacturas.length} factura(s) nueva(s) de Google Sheets`
  if (sinNumFactura.length > 0) msg += ` · ${sinNumFactura.length} sin N° factura ignorada(s)`
  if (yaEnExcel.length > 0) msg += ` · ${yaEnExcel.length} ya existente(s)`

  return { inserted: insertions.length, insertions, message: msg }
}
