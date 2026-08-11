import * as XLSX from 'xlsx'

const ABONOS_SHEET = 'ABONOS'

export function parseAbonosFromWorkbook(workbook) {
  if (!workbook || !workbook.Sheets[ABONOS_SHEET]) return {}
  
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[ABONOS_SHEET], { header: 1, defval: '' })
  const abonos = {}
  
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i]
    const mes = String(row[0] || '').trim()
    const proveedor = String(row[1] || '').trim()
    const factura = String(row[2] || '').trim()
    const fecha = String(row[3] || '').trim()
    // Col E(4) = Valor Antes, Col F(5) = Abono, Col G(6) = Resultante, Col H(7) = Metodo
    // Soporte backward-compat: si col 5 es 0/empty, usar col 4 como valor (formato viejo)
    let valor = typeof row[5] === 'number' ? row[5] : (parseFloat(row[5]) || 0)
    let valorAntes = typeof row[4] === 'number' ? row[4] : (parseFloat(row[4]) || 0)
    const metodo = String(row[7] || '').trim() || String(row[6] || '').trim()
    
    // Backward compat: si col 5 es 0 o vacío, formato viejo donde col 4 = abono, col 5 = valorAntes
    if (valor <= 0 && valorAntes > 0) {
      valor = valorAntes
      valorAntes = 0
      // En formato viejo: col 4 = abono, col 5 = valorAntes, col 6 = metodo
      valorAntes = row[5] !== undefined && row[5] !== '' ? (typeof row[5] === 'number' ? row[5] : parseFloat(row[5]) || 0) : null
      if (!metodo) {
        // En formato viejo metodo estaba en col 6
        valor = typeof row[4] === 'number' ? row[4] : (parseFloat(row[4]) || 0)
      }
    }
    
    if (!mes || !proveedor || valor <= 0) continue
    
    const key = buildKey(proveedor, factura)
    if (!abonos[mes]) abonos[mes] = {}
    if (!abonos[mes][key]) abonos[mes][key] = { proveedor, factura, items: [] }
    abonos[mes][key].items.push({ fecha, valor, valorAntes: valorAntes || null, metodo, _row: i })
  }
  
  return abonos
}

export function getAbonosForKey(abonos, sheetName, proveedor, factura) {
  const mesAbonos = abonos[sheetName] || {}
  const key = buildKey(proveedor, factura)
  if (mesAbonos[key]) return mesAbonos[key].items
  return []
}

export function getTotalAbonadoForKey(abonos, sheetName, proveedor, factura) {
  const items = getAbonosForKey(abonos, sheetName, proveedor, factura)
  return items.reduce((sum, a) => sum + a.valor, 0)
}

function buildKey(proveedor, factura) {
  const p = String(proveedor || '').trim().toLowerCase()
  const f = String(factura || '').trim()
  return f ? p + '||' + f : p
}

export function buildAbonoInsertion(mes, proveedor, factura, fecha, valor, metodo, valorAntes) {
  const valorDespues = valorAntes != null ? valorAntes - valor : null
  let fechaFmt = fecha
  const dm = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(fecha)
  if (dm) fechaFmt = dm[3] + '-' + dm[2] + '-' + dm[1]
  const cells = {
    0: mes,
    1: proveedor,
    2: factura || '',
    3: fechaFmt,
    4: valorAntes != null ? valorAntes : '',
    5: valor,
    6: valorDespues != null ? valorDespues : '',
    7: metodo || 'Efectivo'
  }
  return {
    id: 'abono-' + Date.now() + '-' + Math.random().toString(36).slice(2, 5),
    insertAfterRow: -1,
    sectionKey: ABONOS_SHEET,
    cells
  }
}

export function getAllAbonosForMonth(abonos, sheetName) {
  const mesAbonos = abonos[sheetName] || {}
  const all = []
  for (const key of Object.keys(mesAbonos)) {
    const entry = mesAbonos[key]
    for (const item of entry.items) {
      all.push({ proveedor: entry.proveedor, factura: entry.factura, ...item })
    }
  }
  all.sort((a, b) => String(a.fecha || '').localeCompare(String(b.fecha || '')))
  return all
}

export { ABONOS_SHEET }
