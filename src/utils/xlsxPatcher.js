/**
 * xlsxPatcher.js
 *
 * Parchea quirúrgicamente un archivo .xlsx sin pasar por SheetJS.
 * Un .xlsx es un ZIP que contiene archivos XML. Abrimos ese ZIP,
 * encontramos la celda exacta en el XML de la hoja, sustituimos
 * solo el valor y volvemos a comprimir. Todo lo demás (estilos,
 * celdas fusionadas, colores, fuentes, fórmulas no tocadas, macros…)
 * queda byte-a-byte idéntico al original.
 */

import JSZip from 'jszip'
import { MONTHS } from './excelParser'

function deepCloneStyle(obj) {
  if (!obj || typeof obj !== 'object') return obj
  const clone = {}
  for (const key of Object.keys(obj)) {
    const val = obj[key]
    if (val && typeof val === 'object' && !Array.isArray(val)) {
      clone[key] = deepCloneStyle(val)
    } else if (Array.isArray(val)) {
      clone[key] = val.map(v => deepCloneStyle(v))
    } else {
      clone[key] = val
    }
  }
  return clone
}

// ─────────────────────────────────────────────────────────────
// Helpers internos
// ─────────────────────────────────────────────────────────────

/** Convierte índices 0-based a referencia de celda tipo "D6". */
function encodeCell(r, c) {
  let col = ''
  let n   = c + 1
  while (n > 0) {
    col = String.fromCharCode(65 + ((n - 1) % 26)) + col
    n   = Math.floor((n - 1) / 26)
  }
  return col + (r + 1)
}

/** Convierte índice 0-based de columna a letras ("A", "Z", "AA", …). */
function colLetter(c) {
  let col = ''
  let n   = c + 1
  while (n > 0) {
    col = String.fromCharCode(65 + ((n - 1) % 26)) + col
    n   = Math.floor((n - 1) / 26)
  }
  return col
}

/**
 * Devuelve el índice (exclusive) del final de <row r="rowNum">…</row>.
 * rowNum es 1-based (como aparece en el XML de Excel).
 */
function findRowEnd(xml, rowNum) {
  const needle = `r="${rowNum}"`
  let pos = 0
  while (pos < xml.length) {
    const rIdx = xml.indexOf(needle, pos)
    if (rIdx === -1) return -1
    // Retroceder hasta '<'
    let tagStart = rIdx
    while (tagStart > 0 && xml[tagStart] !== '<') tagStart--
    // Confirmar que es <row
    if (xml.slice(tagStart, tagStart + 4) !== '<row') {
      pos = rIdx + needle.length
      continue
    }
    // Avanzar hasta '>' de cierre del tag de apertura
    let tagEnd = rIdx + needle.length
    while (tagEnd < xml.length && xml[tagEnd] !== '>') tagEnd++
    if (tagEnd >= xml.length) return -1
    if (xml[tagEnd - 1] === '/') return tagEnd + 1   // self-closing
    const closeIdx = xml.indexOf('</row>', tagEnd + 1)
    if (closeIdx === -1) return -1
    return closeIdx + 6
  }
  return -1
}

/** Convierte letras de columna ("A", "Z", "AA", …) a índice 0-based. */
function colLetterToIndex(col) {
  let n = 0
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + col.charCodeAt(i) - 64
  }
  return n - 1
}

/**
 * Ajusta referencias de fila relativas en una fórmula de Excel.
 * Cambia COLUMN+oldRow → COLUMN+newRow para referencias no absolutas ($).
 * Las referencias con $ antes del número de fila se mantienen intactas.
 */
function adjustFormulaRow(formula, oldRow, newRow) {
  return formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (match, colDollar, col, rowDollar, rowStr) => {
    if (rowDollar === '$') return match                          // fila absoluta → no tocar
    if (Number(rowStr) === oldRow) return colDollar + col + newRow  // fila relativa del template → actualizar
    return match
  })
}

/** Extrae el XML completo de la fila rowNum (1-based) para copiar estilos. */
function getRowXml(xml, rowNum) {
  const needle = `r="${rowNum}"`
  let pos = 0
  while (pos < xml.length) {
    const rIdx = xml.indexOf(needle, pos)
    if (rIdx === -1) return null
    let tagStart = rIdx
    while (tagStart > 0 && xml[tagStart] !== '<') tagStart--
    if (xml.slice(tagStart, tagStart + 4) !== '<row') {
      pos = rIdx + needle.length
      continue
    }
    const endIdx = findRowEnd(xml, rowNum)
    if (endIdx === -1) return null
    return xml.slice(tagStart, endIdx)
  }
  return null
}

/**
 * Construye el XML de una nueva fila.
 * @param {number}      rowNum         Número de fila 1-based.
 * @param {Object}      cells          { colIdx0based: value }
 * @param {string|null} templateRowXml XML de la fila plantilla para copiar estilos numéricos.
 */
function buildNewRowXml(rowNum, cells, templateRowXml) {
  // numericStyleMap: col letter → s= style id  (celdas numéricas sin fórmula)
  // formulaCellMap:  col letter → { style, formula }  (celdas con <f>)
  //
  // Las celdas de TEXTO no copian el estilo porque el estilo original puede tener
  // fuente blanca (el color lo gestiona conditional formatting), lo que haría el
  // texto invisible en Excel aunque el valor sí esté en el archivo.
  const numericStyleMap = {}
  const formulaCellMap  = {}
  const templateRowNum  = rowNum - 1    // fila plantilla (1-based)

  if (templateRowXml) {
    const process = (attrs, content = '') => {
      const colM = /\br="([A-Z]+)\d+"/.exec(attrs)
      const sM   = /\bs="(\d+)"/.exec(attrs)
      const tM   = /\bt="([^"]+)"/.exec(attrs)
      // Fórmula: extrae solo si hay texto real dentro de <f> (no shared slaves vacíos)
      const fM   = /<f[^>]*>([^<]+)<\/f>/.exec(content)
      if (!colM) return
      const col = colM[1]
      if (fM) {
        formulaCellMap[col] = { style: sM ? sM[1] : null, formula: fM[1] }
      } else if (sM && (!tM || tM[1] === 'n')) {
        numericStyleMap[col] = sM[1]
      }
    }
    let m
    // self-closing <c ... />
    const reSC = /<c\b([^>]*?)\/>/g
    while ((m = reSC.exec(templateRowXml)) !== null) process(m[1])
    // open-close <c ...>...</c>
    const reOC = /<c\b([^>]*)>([\s\S]*?)<\/c>/g
    while ((m = reOC.exec(templateRowXml)) !== null) process(m[1], m[2])
  }

  // Acumular celdas indexadas por columna (0-based) para emitirlas ordenadas
  const userColLetters = new Set(Object.keys(cells).map(k => colLetter(Number(k))))
  const cellParts = {}   // colIndex → xml string

  // 1. Fórmulas propagadas desde la fila plantilla (si el usuario no sobreescribió esa columna)
  for (const [col, { style, formula }] of Object.entries(formulaCellMap)) {
    if (userColLetters.has(col)) continue
    const colIdx = colLetterToIndex(col)
    const ref    = `${col}${rowNum}`
    const sAttr  = style ? ` s="${style}"` : ''
    const adj    = adjustFormulaRow(formula, templateRowNum, rowNum)
    cellParts[colIdx] = `<c r="${ref}"${sAttr}><f>${adj}</f></c>`
  }

  // 2. Valores proporcionados explícitamente (Estado, Descripción, Valor, Fecha, etc.)
  for (const [colIdxStr, value] of Object.entries(cells)) {
    if (value === '' || value === null || value === undefined) continue
    const colIdx    = Number(colIdxStr)
    const col       = colLetter(colIdx)
    const ref       = `${col}${rowNum}`
    const isNumeric = typeof value === 'number' || dateStringToSerial(String(value)) !== null
    const sAttr     = isNumeric && numericStyleMap[col] ? ` s="${numericStyleMap[col]}"` : ''
    cellParts[colIdx] = buildCellXml(`<c r="${ref}"${sAttr}>`, value)
  }

  // Emitir celdas en orden de columna (requisito de Excel)
  const cellsXml = Object.keys(cellParts)
    .map(Number).sort((a, b) => a - b)
    .map(idx => cellParts[idx])
    .join('')

  return `<row r="${rowNum}">${cellsXml}</row>`
}

/**
 * Incrementa en +inc todos los números de fila > afterRow (1-based) en el fragmento xml dado.
 * Actualiza: <row r="N">, <c r="XN">, texto de fórmulas <f>, atributos ref= (mergeCell,
 * fórmulas compartidas, dataValidation, dimension) y sqref=.
 */
function renumberAfter(xml, afterRow, inc) {
  const shift    = n => Number(n) >  afterRow ? Number(n) + inc : Number(n)
  // shiftEnd: para el extremo FINAL de un rango sqref, usamos >= para que el rango
  // se extienda cuando la nueva fila se inserta justo al final del rango.
  // Ejemplo: sqref="J33:J80" + insertar tras fila 80 → sqref="J33:J81"
  const shiftEnd = n => Number(n) >= afterRow ? Number(n) + inc : Number(n)

  // 1. Atributos r= en <row>
  xml = xml.replace(/<row\b[^>]*>/g, tag =>
    tag.replace(/\br="(\d+)"/, (m, n) => `r="${shift(n)}"`)
  )

  // 2. Atributos r= en <c> (aperturas y self-closing)
  xml = xml.replace(/<c\b[^>]*>/g, tag =>
    tag.replace(/\br="([A-Z]+)(\d+)"/, (m, col, n) => `r="${col}${shift(n)}"`)
  )

  // 3. Texto de fórmulas <f ...>FORMULA</f>
  //    Actualiza todas las referencias A1-style dentro de la expresión.
  //    Los < dentro de fórmulas son &lt; en XML, así que [^<]* es seguro.
  xml = xml.replace(/(<f\b[^>]*>)([^<]*)(<\/f>)/g, (m, open, text, close) =>
    open + text.replace(/\b([A-Z]+)(\d+)\b/g, (r, col, n) => `${col}${shift(n)}`) + close
  )

  // 4. Todos los atributos ref="..." (mergeCell, fórmulas compartidas, dimension,
  //    dataValidation, conditionalFormatting, etc.)
  //    Rangos  ref="A5:B20"
  xml = xml.replace(/\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g,
    (m, c1, r1, c2, r2) => `ref="${c1}${shift(r1)}:${c2}${shift(r2)}"`
  )
  //    Celdas simples  ref="A5"
  xml = xml.replace(/\bref="([A-Z]+)(\d+)"/g,
    (m, col, n) => `ref="${col}${shift(n)}"`
  )

  // Helper que procesa una lista de rangos sqref (separados por espacio):
  // - Extremo INICIAL: shift normal (>)
  // - Extremo FINAL:   shiftEnd (>=) para extender al insertar en el límite
  function processSqrefList(val) {
    const parts = val.split(/\s+/).filter(Boolean)
    return parts.map(part => {
      const colon = part.indexOf(':')
      if (colon !== -1) {
        const startFixed = part.slice(0, colon).replace(/([A-Z]+)(\d+)/, (r, c, n) => c + shift(n))
        const endFixed   = part.slice(colon + 1).replace(/([A-Z]+)(\d+)/, (r, c, n) => c + shiftEnd(n))
        return startFixed + ':' + endFixed
      }
      return part.replace(/([A-Z]+)(\d+)/, (r, c, n) => c + shift(n))
    }).join(' ')
  }

  // 5a. sqref="..." como atributo (formato tradicional de conditionalFormatting / dataValidation)
  xml = xml.replace(/\bsqref="([^"]*)"/g, (m, val) => `sqref="${processSqrefList(val)}"`)

  // 5b. <xm:sqref>J33:J80</xm:sqref> — formato extendido de CF en Excel 2013+
  //     Las reglas de colorScale, dataBar, iconSet usan este elemento en extLst.
  xml = xml.replace(/<xm:sqref>([^<]*)<\/xm:sqref>/g,
    (m, val) => `<xm:sqref>${processSqrefList(val)}</xm:sqref>`
  )

  return xml
}

/**
 * Inserta nuevas filas en el XML de una hoja.
 * @param {string} xml
 * @param {Array}  insertions  [{ insertAfterRow (0-based), cells: { colIdx: value } }]
 */
function insertRowsInSheetXml(xml, insertions) {
  if (!insertions || insertions.length === 0) return xml
  const sorted = [...insertions].sort((a, b) => a.insertAfterRow - b.insertAfterRow)
  let offset = 0
  for (const ins of sorted) {
    const targetXmlRow = ins.insertAfterRow + 1 + offset   // convertir a 1-based + offset acumulado
    const templateXml  = getRowXml(xml, targetXmlRow)
    const endIdx       = findRowEnd(xml, targetXmlRow)
    if (endIdx === -1) continue                            // fila no encontrada, saltar
    const newRowNum = targetXmlRow + 1
    const newRowXml = '\n    ' + buildNewRowXml(newRowNum, ins.cells, templateXml)
    // Renumerar sufijo (filas > targetXmlRow) antes de insertar para mantener orden
    const prefix        = xml.slice(0, endIdx)
    const renamedSuffix = renumberAfter(xml.slice(endIdx), targetXmlRow, 1)
    xml = prefix + newRowXml + renamedSuffix
    offset++
  }
  return xml
}

/** Intenta convertir "dd/mm/yyyy" a serial de Excel. */
function dateStringToSerial(str) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(str)
  if (!m) return null
  const [, d, mo, y] = m.map(Number)
  const date = new Date(Date.UTC(y, mo - 1, d))
  if (isNaN(date.getTime())) return null
  return Math.round(date.getTime() / 86400000) + 25569
}

/**
 * Construye el XML de una celda con el nuevo valor.
 * Preserva todos los atributos originales (s, r, etc.) salvo t.
 */
function buildCellXml(openTag, value) {
  // Extraer atributos: quitar < c inicial, > final y atributo t si existe
  let attrs = openTag
    .replace(/^<c\s*/, '')
    .replace(/>$/, '')
    .replace(/\bt="[^"]*"/, '')
    .replace(/\s+/g, ' ')
    .trim()

  if (typeof value === 'number') {
    return `<c ${attrs}><v>${value}</v></c>`
  }

  // Si parece una fecha "dd/mm/yyyy", convertir a serial (preserva formato de fecha de la celda)
  const serial = dateStringToSerial(String(value))
  if (serial) {
    return `<c ${attrs}><v>${serial}</v></c>`
  }

  // Texto plano → inline string (Excel acepta este formato; no requiere tocar sharedStrings.xml)
  const escaped = String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
  return `<c ${attrs} t="inlineStr"><is><t>${escaped}</t></is></c>`
}

/**
 * Encuentra y reemplaza la celda `cellRef` en el XML de la hoja.
 *
 * Estrategia robusta en dos pasos:
 * 1. Localizar el atributo r="cellRef" en el XML.
 * 2. Retroceder hasta el inicio del <c, avanzar hasta el cierre del tag
 *    de apertura (> o />) y determinar si es self-closing o tiene contenido.
 *
 * Esto evita falsas coincidencias de regex sobre self-closing tags con
 * atributos adicionales (p.ej. <c r="E5" s="12"/>) que hacen que
 * patrones basados en [^>]* capturen el />, rompiendo el XML.
 */
function patchCell(xml, cellRef, value) {
  const needle = `r="${cellRef}"`
  let searchFrom = 0

  while (true) {
    const attrIdx = xml.indexOf(needle, searchFrom)
    if (attrIdx === -1) return xml   // celda no encontrada

    // Retroceder hasta el > del elemento anterior, lo que sigue debe ser <c
    const prevGt  = xml.lastIndexOf('>', attrIdx)
    const segment = xml.slice(prevGt + 1, attrIdx) // texto entre > previo y r="..."

    if (!segment.trimStart().startsWith('<c')) {
      // El atributo r= pertenece a otro elemento (p.ej. <row r="5">), seguir buscando
      searchFrom = attrIdx + needle.length
      continue
    }

    // Inicio del elemento <c
    const tagStart = prevGt + 1 + segment.indexOf('<c')

    // Avanzar desde el atributo hasta el cierre del tag de apertura (> o />)
    const gtIdx = xml.indexOf('>', attrIdx + needle.length)
    if (gtIdx === -1) return xml  // XML malformado

    const isSelfClosing = xml[gtIdx - 1] === '/'
    const openTag       = xml.slice(tagStart, gtIdx + 1)   // incluye > o />

    if (isSelfClosing) {
      // <c r="E5" s="12"/>  →  reemplazar solo el elemento self-closing
      const fakeOpen = openTag.slice(0, -2) + '>'  // quitar />, añadir >
      const newCell  = buildCellXml(fakeOpen, value)
      return xml.slice(0, tagStart) + newCell + xml.slice(gtIdx + 1)
    } else {
      // <c r="E5" ...>...</c>  →  reemplazar hasta </c>
      const closeIdx = xml.indexOf('</c>', gtIdx + 1)
      if (closeIdx === -1) return xml  // XML malformado
      const newCell = buildCellXml(openTag, value)
      return xml.slice(0, tagStart) + newCell + xml.slice(closeIdx + 4)
    }
  }
}

// ─────────────────────────────────────────────────────────────
// Mapa hoja → archivo XML en el ZIP
// ─────────────────────────────────────────────────────────────

/**
 * Lee xl/workbook.xml y xl/_rels/workbook.xml.rels del ZIP para
 * construir un mapa { "ENERO": "xl/worksheets/sheet2.xml", … }.
 */
async function buildSheetFileMap(zip) {
  const relsText = await zip.file('xl/_rels/workbook.xml.rels')?.async('string')
  const wbText   = await zip.file('xl/workbook.xml')?.async('string')
  if (!relsText || !wbText) throw new Error('No se pudo leer la estructura interna del .xlsx')

  // rId → ruta relativa dentro de xl/
  const relMap = {}
  const relRe  = /Id="([^"]+)"[^>]+Target="([^"]+)"/g
  let m
  while ((m = relRe.exec(relsText)) !== null) {
    relMap[m[1]] = 'xl/' + m[2].replace(/^\//, '')
  }

  // nombre de hoja → ruta completa
  const nameToFile = {}
  const sheetRe    = /<sheet\s[^>]*\bname="([^"]+)"[^>]*\br:id="([^"]+)"/g
  while ((m = sheetRe.exec(wbText)) !== null) {
    if (relMap[m[2]]) nameToFile[m[1]] = relMap[m[2]]
  }

  return nameToFile
}

// ─────────────────────────────────────────────────────────────
// API pública
// ─────────────────────────────────────────────────────────────

/**
 * Aplica edits puntuales y/o inserciones de filas al buffer original de un .xlsx.
 *
 * @param {ArrayBuffer} rawBuffer        Buffer original del archivo (sin modificar).
 * @param {Object}      pendingEdits     { sheetName: { "row,col": value } }
 * @param {Object}      pendingInsertions { sheetName: [{ insertAfterRow (0-based), cells: { colIdx: value } }] }
 * @returns {Promise<ArrayBuffer>}       Nuevo buffer listo para escribir al disco.
 */
export async function patchXlsx(rawBuffer, pendingEdits, pendingInsertions = {}) {
  const zip          = await JSZip.loadAsync(rawBuffer)
  const sheetFileMap = await buildSheetFileMap(zip)

  const allSheets = new Set([
    ...Object.keys(pendingEdits),
    ...Object.keys(pendingInsertions).filter(k => (pendingInsertions[k] || []).length > 0)
  ])

  for (const sheetName of allSheets) {
    const sheetFile = sheetFileMap[sheetName]
    if (!sheetFile) continue
    const entry = zip.file(sheetFile)
    if (!entry) continue

    let xml = await entry.async('string')

    // 1. Insertar nuevas filas (modifica números de fila del XML)
    const insertions = pendingInsertions[sheetName] || []
    if (insertions.length > 0) {
      xml = insertRowsInSheetXml(xml, insertions)
    }

    // 2. Parchear celdas existentes, ajustando índices de fila por las inserciones previas
    const sheetEdits = pendingEdits[sheetName] || {}
    if (Object.keys(sheetEdits).length > 0) {
      const sortedIns = [...insertions].sort((a, b) => a.insertAfterRow - b.insertAfterRow)
      for (const [key, value] of Object.entries(sheetEdits)) {
        let [r, c] = key.split(',').map(Number)
        // Por cada inserción cuyo insertAfterRow (0-based) < r, la fila se desplazó +1
        const shift = sortedIns.filter(ins => ins.insertAfterRow < r).length
        xml = patchCell(xml, encodeCell(r + shift, c), value)
      }
    }

    zip.file(sheetFile, xml)
  }

  // Eliminar calcChain para evitar advertencias de Excel por cadenas de fórmulas desactualizadas
  zip.remove('xl/calcChain.xml')

  return zip.generateAsync({
    type:               'arraybuffer',
    compression:        'DEFLATE',
    compressionOptions: { level: 6 }
  })
}

/**
 * Elimina filas específicas del XML de una hoja (por número de fila 1-based).
 * Reenumera todas las filas posteriores.
 */
function removeRowsFromSheetXml(xml, rowNumbers1Based) {
  if (!rowNumbers1Based || rowNumbers1Based.length === 0) return xml
  const sorted = [...rowNumbers1Based].sort((a, b) => b - a) // procesar de abajo hacia arriba
  for (const rowNum of sorted) {
    const startIdx = xml.indexOf(`<row r="${rowNum}"`)
    if (startIdx === -1) continue
    const endIdx = findRowEnd(xml, rowNum)
    if (endIdx === -1) continue
    const prefix = xml.slice(0, startIdx)
    const suffix = xml.slice(endIdx)
    xml = prefix + renumberAfter(suffix, rowNum, -1)
  }
  return xml
}

/**
 * PASO 2: Solo workbook.xml (sin rels ni content_types).
 */
export async function cloneSheet(rawBuffer, sourceSheetName, targetSheetName, canceladoRows) {
  const zip = await JSZip.loadAsync(rawBuffer)
  const sheetFileMap = await buildSheetFileMap(zip)

  const sourceFile = sheetFileMap[sourceSheetName]
  if (!sourceFile) throw new Error(`No se encontró la hoja "${sourceSheetName}"`)

  let sheetXml = await zip.file(sourceFile).async('string')

  // Limpiar referencia a drawing (evita reparación de Excel)
  sheetXml = sheetXml.replace(/<legacyDrawing[^>]*\/>/g, '')

  // Eliminar filas Cancelado
  if (canceladoRows && canceladoRows.length > 0) {
    const rows1Based = canceladoRows.map(i => i + 1)
    sheetXml = removeRowsFromSheetXml(sheetXml, rows1Based)
    sheetXml = sheetXml.replace(/<f\b([^>]*?)>/g, (full, attrs) => {
      if (attrs.includes('t="shared"')) {
        return '<f' + attrs.replace(/\s*t="shared"/, '').replace(/\s*si="\d+"/, '').replace(/\s*ref="[^"]*"/, '') + '>'
      }
      return full
    })
  }

  // Poner todos los estados en "Sin Pagar" (solo columna G)
  sheetXml = sheetXml.replace(/(<c r="G\d+"[^>]*t="inlineStr"[^>]*><is><t>)[^<]*(<\/t><\/is><\/c>)/g, '$1Sin Pagar$2')
  sheetXml = sheetXml.replace(/<c r="(G\d+)"([^>]*)t="s"([^>]*)>(<v>\d+<\/v>)<\/c>/g, '<c r="$1"$2t="inlineStr"$3><is><t>Sin Pagar</t></is></c>')
  sheetXml = sheetXml.replace(/<c r="(G\d+)"([^>]*?)\/>/g, '<c r="$1"$2 t="inlineStr"><is><t>Sin Pagar</t></is></c>')

  // Actualizar referencias de mes en Nómina quincena
  const sheetIdx = targetSheetName ? MONTHS.findIndex(m => targetSheetName.toUpperCase().includes(m)) : -1
  if (sheetIdx >= 0) {
    const currentMonth = MONTHS[sheetIdx]
    const prevMonth = MONTHS[(sheetIdx + 11) % 12]
    const MESES_ABREV = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC']
    const currentAbrev = MESES_ABREV[sheetIdx]
    const prevAbrev = MESES_ABREV[(sheetIdx + 11) % 12]
    const allMonths = MONTHS.concat(MESES_ABREV).join('|')

    // Quincena 2 → mes anterior
    sheetXml = sheetXml.replace(
      new RegExp('(Nómina quincena 2[^<]*?\\()(' + allMonths + ')(\\))', 'gi'),
      (m, p1, p2, p3) => p1 + prevMonth + p3
    )
    // Quincena 1 → mes actual  
    sheetXml = sheetXml.replace(
      new RegExp('(Nómina quincena 1[^<]*?\\()(' + allMonths + ')(\\))', 'gi'),
      (m, p1, p2, p3) => p1 + currentMonth + p3
    )
  }

  const wbText = await zip.file('xl/workbook.xml').async('string')
  const relsText = await zip.file('xl/_rels/workbook.xml.rels').async('string')

  const existingNums = Object.keys(zip.files)
    .filter(f => /^xl\/worksheets\/sheet\d+\.xml$/.test(f))
    .map(f => parseInt(f.match(/sheet(\d+)/)[1]))
  const newSheetNum = (existingNums.length > 0 ? Math.max(...existingNums) : 0) + 1
  const newSheetFile = 'xl/worksheets/sheet' + newSheetNum + '.xml'

  const rids = (relsText.match(/Id="rId(\d+)"/g) || []).map(m => parseInt(m.match(/\d+/)[0]))
  const newRid = (rids.length > 0 ? Math.max(...rids) : 0) + 1
  const ids = (wbText.match(/sheetId="(\d+)"/g) || []).map(m => parseInt(m.match(/\d+/)[0]))
  const newSheetId = (ids.length > 0 ? Math.max(...ids) : 0) + 1
  const escName = targetSheetName.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

  const modWb = wbText.replace('</sheets>', '  <sheet name="' + escName + '" sheetId="' + newSheetId + '" r:id="rId' + newRid + '"/>\n  </sheets>')
  const modRels = relsText.replace('</Relationships>', '  <Relationship Id="rId' + newRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + newSheetNum + '.xml"/>\n</Relationships>')

  const enc = new TextEncoder()
  const newZip = new JSZip()
  const skipSet = new Set(['xl/workbook.xml', 'xl/_rels/workbook.xml.rels'])
  const allFiles = Object.keys(zip.files)
  for (const name of allFiles) {
    const entry = zip.files[name]
    if (entry.dir) continue
    if (skipSet.has(name)) continue
    const data = await entry.async('uint8array')
    newZip.file(name, data)
  }
  newZip.file(newSheetFile, enc.encode(sheetXml))
  newZip.file('xl/workbook.xml', enc.encode(modWb))
  newZip.file('xl/_rels/workbook.xml.rels', enc.encode(modRels))
  // Content_Types NO se modifica

  const blob = await newZip.generateAsync({ type: 'blob', compression: 'DEFLATE' })
  return await blob.arrayBuffer()
}

/**
 * Genera un archivo de resumen basado en la plantilla FLUJO DE CAJA RESUMEN.xlsx.
 * Reemplaza los datos de Hoja1 preservando TODO el formato.
 */
export async function generarResumenXlsx(templateBuf, dataRows, nuevosMapeos = {}) {
  const zip = await JSZip.loadAsync(templateBuf)
  const sheetFileMap = await buildSheetFileMap(zip)
  const sourceFile = sheetFileMap['Hoja1']
  if (!sourceFile) throw new Error('No se encontró Hoja1 en la plantilla')

  let sheetXml = await zip.file(sourceFile).async('string')

  // Eliminar todas las filas de datos existentes (row > 1)
  const dataRowNums = []
  const rowRe = /<row\s+r="(\d+)"/g
  let m
  while ((m = rowRe.exec(sheetXml)) !== null) {
    const rn = parseInt(m[1])
    if (rn > 1) dataRowNums.push(rn)
  }
  dataRowNums.sort((a, b) => b - a) // de abajo hacia arriba
  for (const rowNum of dataRowNums) {
    sheetXml = removeRowsFromSheetXml(sheetXml, [rowNum])
  }

  // Template de estilo: usar la fila 2 del XML original
  const origSheetXml = await zip.file(sourceFile).async('string')
  const templateRowXml = getRowXml(origSheetXml, 2)

  // Insertar nuevas filas después de la fila 1
  const enc = new TextEncoder()
  let currentAfter = 1
  for (let i = 0; i < dataRows.length; i++) {
    const [tipo, nombre, valor, fecha] = dataRows[i]
    const cells = {
      0: tipo || '',
      1: nombre || '',
      2: typeof valor === 'number' ? valor : (parseFloat(valor) || 0),
      3: typeof fecha === 'string' && fecha.includes('/') ? fecha : (fecha || '')
    }

    const endIdx = findRowEnd(sheetXml, currentAfter)
    if (endIdx === -1) break

    const newRowNum = currentAfter + 1
    const newRowXml = '\n    ' + buildNewRowXml(newRowNum, cells, templateRowXml)
    const prefix = sheetXml.slice(0, endIdx)
    const renamedSuffix = renumberAfter(sheetXml.slice(endIdx), currentAfter, 1)
    sheetXml = prefix + newRowXml + renamedSuffix
    currentAfter++
  }

  // Agregar fórmulas VLOOKUP en columna E
  for (let i = 0; i < dataRows.length; i++) {
    const rowNum = 2 + i
    const cellRef = encodeCell(rowNum - 1, 4)
    const formulaXml = '<c r="' + cellRef + '"><f>VLOOKUP(B' + rowNum + ',Hoja2!A:B,2,FALSE)</f><v></v></c>'
    // Buscar celda E existente y reemplazar
    const eCellStart = sheetXml.indexOf('r="' + cellRef + '"')
    if (eCellStart !== -1) {
      let tagStart = eCellStart
      while (tagStart > 0 && sheetXml[tagStart] !== '<') tagStart--
      const gtIdx = sheetXml.indexOf('>', eCellStart)
      if (sheetXml[gtIdx - 1] === '/') {
        sheetXml = sheetXml.slice(0, tagStart) + formulaXml + sheetXml.slice(gtIdx + 1)
      } else {
        const closeIdx = sheetXml.indexOf('</c>', gtIdx + 1)
        if (closeIdx !== -1) {
          sheetXml = sheetXml.slice(0, tagStart) + formulaXml + sheetXml.slice(closeIdx + 4)
        }
      }
    }
  }

  zip.file(sourceFile, sheetXml)

  // Forzar recálculo automático al abrir
  let wbXml = await zip.file('xl/workbook.xml').async('string')
  if (wbXml.includes('<calcPr')) {
    wbXml = wbXml.replace(/<calcPr[^>]*\/>/, '<calcPr calcMode="auto" fullCalcOnLoad="1"/>')
  } else {
    wbXml = wbXml.replace('<workbookPr', '<calcPr calcMode="auto" fullCalcOnLoad="1"/><workbookPr')
  }
  zip.file('xl/workbook.xml', wbXml)
  zip.remove('xl/calcChain.xml')

  // Agregar nuevos mapeos a Hoja2 si hay
  const hoja2Entries = Object.entries(nuevosMapeos || {})
  const hs2File = sheetFileMap['Hoja2']
  if (hs2File) {
    let hs2Xml = await zip.file(hs2File).async('string')
    // Recortar espacios al final de todos los textos en Hoja2
    hs2Xml = hs2Xml.replace(/(<t[^>]*>)([\s\S]*?)(\s+)<\/t>/g, (m, open, text, spaces) => open + text + '</t>')

    if (hoja2Entries.length > 0) {
      const rowNums = []
      const rowRe2 = /<row\s+r="(\d+)"/g
      let m2
      while ((m2 = rowRe2.exec(hs2Xml)) !== null) rowNums.push(parseInt(m2[1]))
      const lastRow = rowNums.length > 0 ? Math.max(...rowNums) : 1

      const hs2Template = getRowXml(hs2Xml, Math.min(2, lastRow))
      let afterRow = lastRow
      for (const [nombre, descuento] of hoja2Entries) {
        const endIdx = findRowEnd(hs2Xml, afterRow)
        if (endIdx === -1) break
        const newRowNum = afterRow + 1
        const cells = { 0: nombre, 1: descuento }
        const newRowXml = '\n    ' + buildNewRowXml(newRowNum, cells, hs2Template)
        hs2Xml = hs2Xml.slice(0, endIdx) + newRowXml + renumberAfter(hs2Xml.slice(endIdx), afterRow, 1)
        afterRow++
      }
    }
    zip.file(hs2File, hs2Xml)
  }

  const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' })
  return await blob.arrayBuffer()
}

/**
 * Agrega una hoja ABONOS a un workbook si no existe.
 * Retorna el nuevo buffer o el original si ya existe.
 */
export async function ensureAbonosSheet(rawBuffer) {
  const zip = await JSZip.loadAsync(rawBuffer)
  const sheetFileMap = await buildSheetFileMap(zip)
  
  if (sheetFileMap['ABONOS']) return rawBuffer // ya existe
  
  const existingNums = Object.keys(zip.files)
    .filter(f => /^xl\/worksheets\/sheet\d+\.xml$/.test(f))
    .map(f => parseInt(f.match(/sheet(\d+)/)[1]))
  const newSheetNum = (existingNums.length > 0 ? Math.max(...existingNums) : 0) + 1
  const newSheetFile = 'xl/worksheets/sheet' + newSheetNum + '.xml'
  
  let wbText = await zip.file('xl/workbook.xml').async('string')
  const ids = (wbText.match(/sheetId="(\d+)"/g) || []).map(m => parseInt(m.match(/\d+/)[0]))
  const newSheetId = (ids.length > 0 ? Math.max(...ids) : 0) + 1
  
  const relsText = await zip.file('xl/_rels/workbook.xml.rels').async('string')
  const rids = (relsText.match(/Id="rId(\d+)"/g) || []).map(m => parseInt(m.match(/\d+/)[0]))
  const newRid = (rids.length > 0 ? Math.max(...rids) : 0) + 1
  
  const headerXml = '<row r="1">' +
    '<c r="A1" t="inlineStr"><is><t>Mes</t></is></c>' +
    '<c r="B1" t="inlineStr"><is><t>Proveedor</t></is></c>' +
    '<c r="C1" t="inlineStr"><is><t>Factura</t></is></c>' +
    '<c r="D1" t="inlineStr"><is><t>Fecha</t></is></c>' +
    '<c r="E1" t="inlineStr"><is><t>Valor Antes</t></is></c>' +
    '<c r="F1" t="inlineStr"><is><t>Abono</t></is></c>' +
    '<c r="G1" t="inlineStr"><is><t>Resultante</t></is></c>' +
    '<c r="H1" t="inlineStr"><is><t>Metodo</t></is></c>' +
    '</row>'
  const sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData>' + headerXml + '</sheetData></worksheet>'
  
  const enc = new TextEncoder()
  const newZip = new JSZip()
  const skipSet = new Set(['xl/workbook.xml', 'xl/_rels/workbook.xml.rels'])
  const allFiles = Object.keys(zip.files)
  for (const name of allFiles) {
    const entry = zip.files[name]
    if (entry.dir) continue
    if (skipSet.has(name)) continue
    newZip.file(name, await entry.async('uint8array'))
  }
  newZip.file(newSheetFile, enc.encode(sheetXml))
  newZip.file('xl/workbook.xml', enc.encode(wbText.replace('</sheets>', '  <sheet name="ABONOS" sheetId="' + newSheetId + '" r:id="rId' + newRid + '" state="hidden"/>\n  </sheets>')))
  newZip.file('xl/_rels/workbook.xml.rels', enc.encode(relsText.replace('</Relationships>', '  <Relationship Id="rId' + newRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + newSheetNum + '.xml"/>\n</Relationships>')))
  
  const blob = await newZip.generateAsync({ type: 'blob', compression: 'DEFLATE' })
  return await blob.arrayBuffer()
}

/**
 * Agrega una hoja MAPEO_DESCUENTOS al workbook si no existe.
 */
export async function ensureMapeoSheet(rawBuffer) {
  const zip = await JSZip.loadAsync(rawBuffer)
  const sheetFileMap = await buildSheetFileMap(zip)
  if (sheetFileMap['MAPEO_DESCUENTOS']) return rawBuffer

  const existingNums = Object.keys(zip.files)
    .filter(f => /^xl\/worksheets\/sheet\d+\.xml$/.test(f))
    .map(f => parseInt(f.match(/sheet(\d+)/)[1]))
  const newSheetNum = (existingNums.length > 0 ? Math.max(...existingNums) : 0) + 1
  const newSheetFile = 'xl/worksheets/sheet' + newSheetNum + '.xml'

  let wbText = await zip.file('xl/workbook.xml').async('string')
  const ids = (wbText.match(/sheetId="(\d+)"/g) || []).map(m => parseInt(m.match(/\d+/)[0]))
  const newSheetId = (ids.length > 0 ? Math.max(...ids) : 0) + 1

  const relsText = await zip.file('xl/_rels/workbook.xml.rels').async('string')
  const rids = (relsText.match(/Id="rId(\d+)"/g) || []).map(m => parseInt(m.match(/\d+/)[0]))
  const newRid = (rids.length > 0 ? Math.max(...rids) : 0) + 1

  const headerXml = '<row r="1">' +
    '<c r="A1" t="inlineStr"><is><t>Concepto</t></is></c>' +
    '<c r="B1" t="inlineStr"><is><t>Categoria</t></is></c>' +
    '</row>'
  const sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData>' + headerXml + '</sheetData></worksheet>'

  const enc = new TextEncoder()
  const newZip = new JSZip()
  const skipSet = new Set(['xl/workbook.xml', 'xl/_rels/workbook.xml.rels'])
  const allFiles = Object.keys(zip.files)
  for (const name of allFiles) {
    const entry = zip.files[name]
    if (entry.dir) continue
    if (skipSet.has(name)) continue
    newZip.file(name, await entry.async('uint8array'))
  }
  newZip.file(newSheetFile, enc.encode(sheetXml))
  newZip.file('xl/workbook.xml', enc.encode(wbText.replace('</sheets>', '  <sheet name="MAPEO_DESCUENTOS" sheetId="' + newSheetId + '" r:id="rId' + newRid + '" state="hidden"/>\n  </sheets>')))
  newZip.file('xl/_rels/workbook.xml.rels', enc.encode(relsText.replace('</Relationships>', '  <Relationship Id="rId' + newRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + newSheetNum + '.xml"/>\n</Relationships>')))

  const blob = await newZip.generateAsync({ type: 'blob', compression: 'DEFLATE' })
  return await blob.arrayBuffer()
}
