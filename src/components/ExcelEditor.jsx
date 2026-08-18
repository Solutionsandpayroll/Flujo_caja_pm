import { useState, useMemo, useEffect, useRef } from 'react'
import * as XLSX from 'xlsx'
import MonthViewer from './MonthViewer'
import { isMonthSheet, monthSheetIndex, MONTHS, parseMonthSheet, excelDateToString, dateStringToSerial, buildSubtotalEdits } from '../utils/excelParser'
import { patchXlsx, cloneSheet, generarResumenXlsx, ensureAbonosSheet, ensureMapeoSheet } from '../utils/xlsxPatcher'
import { saveHandle, loadHandle, clearHandle, requestPermission } from '../utils/fileHandleStore'
import { syncFacturas, getMonthKey } from '../utils/sheetsSync'
import { parseAbonosFromWorkbook, getTotalAbonadoForKey, getAllAbonosForMonth, buildAbonoInsertion, ABONOS_SHEET } from '../utils/abonosStore'

const SLOT_KEYS   = ['colombia']
const SLOT_LABELS = { colombia: 'Flujo de Caja' }

function emptySlot(handle, fileName, buffer, wb) {
  const first = wb.SheetNames.find(isMonthSheet) || wb.SheetNames[0] || ''
  return { fileHandle: handle, rawBuffer: buffer, fileName, workbook: wb,
           selectedSheet: first, pendingEdits: {}, pendingInsertions: {} }
}

function parseMapeoFromWorkbook(workbook) {
  const sheet = workbook?.Sheets?.['MAPEO_DESCUENTOS']
  if (!sheet) return {}
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })
  const map = {}
  for (let i = 1; i < rows.length; i++) {
    const concepto = String(rows[i][0] || '').trim()
    const categoria = String(rows[i][1] || '').trim()
    if (concepto && categoria) map[concepto.toLowerCase()] = categoria
  }
  return map
}

function ExcelEditor() {
  const [slots,       setSlots]       = useState({ colombia: null })
  const [activeSlot,  setActiveSlot]  = useState(null)
  const [savedHandles, setSavedHandles] = useState({ colombia: null })
  const [error,        setError]        = useState('')
  const [saveStatus,   setSaveStatus]   = useState('idle')
  const [reconnecting, setReconnecting] = useState(false)
  const [syncStatus,   setSyncStatus]   = useState('idle')
  const [syncMessage,  setSyncMessage]  = useState('')
  const [resumenModal, setResumenModal] = useState(null)
  const [abonos, setAbonos] = useState({}) // { mes: { proveedor: [{fecha, valor, metodo, _row}] } }
  const [abonosModalOpen, setAbonosModalOpen] = useState(false)
  const [abonoTarget, setAbonoTarget] = useState(null)
  const [lastMonthSheet, setLastMonthSheet] = useState('')
  const [mapeoDescuentos, setMapeoDescuentos] = useState({})
  const [canceladoModal, setCanceladoModal] = useState(null) // { rowIdx, colIdx, onSave }
  const [nuevaSubseccionModal, setNuevaSubseccionModal] = useState(false)
  const rawBufferRef = useRef(null)
  const fileHandleRef = useRef(null)
  const abonosRef = useRef({}) // { proveedor, valorOriginal } // { sinMapeo: [{nombre, valor, fecha}], onConfirm: fn }

  const current           = activeSlot ? slots[activeSlot] : null
  const fileHandle        = current?.fileHandle        ?? null
  const rawBuffer         = current?.rawBuffer         ?? null
  const fileName          = current?.fileName          ?? ''
  const workbook          = current?.workbook          ?? null
  const selectedSheet     = current?.selectedSheet     ?? ''
  const pendingEdits      = current?.pendingEdits      ?? {}
  const pendingInsertions = current?.pendingInsertions ?? {}

  const totalEdits      = Object.values(pendingEdits).reduce((s, e) => s + Object.keys(e).length, 0)
  const totalInsertions = Object.values(pendingInsertions).reduce((s, a) => s + a.length, 0)
  const hasChanges      = totalEdits > 0 || totalInsertions > 0
  const hasAnyFile      = SLOT_KEYS.some(k => slots[k] !== null)

  // Mantener refs actualizadas
  useEffect(() => { rawBufferRef.current = rawBuffer }, [rawBuffer])
  useEffect(() => { fileHandleRef.current = fileHandle }, [fileHandle])
  useEffect(() => { abonosRef.current = abonos }, [abonos])

  // ──────────────────────────────────────────────
  // Al montar: verificar IndexedDB para ambos slots
  // FIX: usar handle.name (no requiere permiso) en vez de handle.getFile()
  // ──────────────────────────────────────────────
  useEffect(() => {
    SLOT_KEYS.forEach(key => {
      loadHandle(`slot-${key}`).then(handle => {
        if (!handle) return
        // handle.name es una propiedad del FileSystemHandle, sin permiso requerido
        setSavedHandles(prev => ({ ...prev, [key]: handle }))
      }).catch(() => {})
    })
  }, [])

  // ──────────────────────────────────────────────
  // Sincronización manual con Google Sheets
  // ──────────────────────────────────────────────
  const handleSyncNow = async () => {
    if (!workbook || !selectedSheet) return
    const mes = getMonthKey(selectedSheet)
    if (!mes) return

    const currentInsertions = slots.colombia?.pendingInsertions?.[selectedSheet] || []

    setSyncStatus('syncing')
    syncFacturas(sheetRows, selectedSheet, currentInsertions).then(result => {
      if (result.insertions && result.insertions.length > 0) {
        setSlots(prev => ({
          ...prev,
          colombia: { ...prev.colombia,
            pendingInsertions: { ...prev.colombia.pendingInsertions,
              [selectedSheet]: [...(prev.colombia.pendingInsertions[selectedSheet] || []), ...result.insertions]
            }
          }
        }))
        setSyncStatus('synced')
        setSyncMessage(result.message || `${result.inserted} factura(s) nueva(s) de Google Sheets`)
      } else if (result.error) {
        setSyncStatus('error')
        setSyncMessage(result.error)
      } else {
        setSyncStatus('synced')
        setSyncMessage(result.message || 'Sin facturas nuevas')
      }
    }).catch(err => {
      setSyncStatus('error')
      setSyncMessage('Error de conexión con Google Sheets')
      console.error('Sync error:', err)
    })
  }

  // ──────────────────────────────────────────────
  // Crear hoja del mes siguiente (clon sin Cancelados)
  // ──────────────────────────────────────────────
  const handleCreateNextMonth = async () => {
    if (!rawBuffer || !workbook || !selectedSheet) return
    const handle = fileHandle || slots.colombia?.fileHandle
    if (!handle) { setError('No se pudo acceder al archivo'); return }

    const sheetIdx = monthSheetIndex(selectedSheet)
    if (sheetIdx === -1) {
      setError('La hoja actual no es un mes. Seleccioná una hoja de mes primero.')
      return
    }

    const year = parseInt(selectedSheet.match(/\d{4}/)?.[0]) || new Date().getFullYear()
    const nextMonthIdx = (sheetIdx + 1) % 12
    const nextYear = sheetIdx === 11 ? year + 1 : year
    const nextName = MONTHS[nextMonthIdx] + ' ' + nextYear

    if (workbook.SheetNames.includes(nextName)) {
      setError(`La hoja "${nextName}" ya existe.`)
      return
    }

    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[selectedSheet], { header: 1, defval: '' })

    // Patrones de items recurrentes (se normalizan quitando mes/año para comparar)
    const RECURRENTES = [
      'seguridad social pm',
      'seguridad personal independiente - antonio bernal',
      'seguridad personal independiente - victor gil',
      'rte ica', 'rte iva', 'rte fte', 'iva',
      'arriendo bodega nueva',
      'internet claro',
      'regente de farmacia',
      'energia bodega nueva',
      'telefonia lineas corporativas',
      'mundo express 1',
      'mundo express 2',
      'jose florez saludia 1',
      'jose florez saludia 2',
      'victor gil 1',
      'victor gil 2',
      'honorarios contabilidad act&j',
      'revisoria fiscal',
      'comisiones jeimy paola ortiz duarte',
      'comisiones nancy rosalba mora moreno ventas',
      'comisiones antonio bernal cardona',
      'comisiones luz stella bernal cardona',
      'comision edward forero',
      'comisiones saludia santiago feliciano',
      'tarjeta de credito bancolombia',
      'intereses viejos (obligatorio)',
      'prestamo 150mm s&p (obligatorio)',
      'intereses prestamo maritza forero',
      'prestamo angie amaya',
      'prestamo dayan manjarres',
      'prestamo oscar cala',
      'prestamo helda doncel',
      'pestamo claudia forero',
      'pestamo deysy gonzalez',
      'pestamo ingrid olarte',
      'pestamo edward forero',
      'servicio de nomina y seguridad social - servicios de',
      'servicio de nómina y seguridad social - servicios de',
      'arriendo software mantis',
      'pleia', 'suit medimantis',
      'nomina quincena 2',
      'nomina quincena 1',
      'nómina quincena 2',
      'nómina quincena 1',
      'pago de intereses',
      'cesantias',
      'intereses a las cesantias',
      'primera dotacion', 'segunda dotacion',
      '1° prima de servicios',
      'poliza de seguro pyme',
      'industria y comercio',
      'acueducto bodega',
      'hosting mantis',
    ]

    function normalizar(texto) {
      return String(texto || '')
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .replace(/\([^)]*(?:ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic|enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)[^)]*\)/gi, '')
        .replace(/\b(?:ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic|enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\b/gi, '')
        .replace(/\s+/g, ' ')
        .trim()
    }

    function esRecurrente(descripcion) {
      const norm = normalizar(descripcion)
      return RECURRENTES.some(p => norm.includes(p))
    }

    const canceladoRows = []
    for (let i = 0; i < rows.length; i++) {
      const estadoA = String(rows[i][0] ?? '').trim().toLowerCase()
      const estadoG = String(rows[i][6] ?? '').trim().toLowerCase()
      const desc = String(rows[i][3] ?? '')
      if ((estadoA === 'cancelado' || estadoG === 'cancelado') && !esRecurrente(desc)) canceladoRows.push(i)
    }

    // Pedir permiso de escritura ANTES de cualquier operación async
    let writable
    try {
      writable = await handle.createWritable()
    } catch (e) {
      setError('Error al acceder al archivo para escritura: ' + e.message)
      return
    }

    try {
      const newBuffer = await cloneSheet(rawBuffer, selectedSheet, nextName, canceladoRows)

      const parsedSrc = parseMonthSheet(rows)
      let finalBuf = newBuffer
      let restaurados = false
      if (parsedSrc && parsedSrc.sectionCXP && abonos[selectedSheet]) {
        const edits = {}
        const canceladoSet = new Set(canceladoRows)
        const calcNewRow = (oldR) => oldR - [...canceladoSet].filter(x => x < oldR).length
        for (const sub of parsedSrc.sectionCXP) {
          for (const r of sub.rows) {
            const totalAbonado = getTotalAbonadoForKey(abonos, selectedSheet, r.proveedor, r.factura)
            if (totalAbonado > 0 && typeof r.valor === 'number') {
              const valorOriginal = r.valor + totalAbonado
              const newRow = calcNewRow(r._row)
              if (!edits[nextName]) edits[nextName] = {}
              edits[nextName][`${newRow},4`] = valorOriginal
            }
          }
        }
        if (Object.keys(edits).length > 0) {
          finalBuf = await patchXlsx(newBuffer, edits, {})
          restaurados = true
        }
      }

      // Actualizar fechas de conceptos con periodicidad (solo si estaban Cancelado en el mes origen)
      const PERIODICIDAD = {
        'iva': 4, 'industria y comercio': 2, 'poliza de seguro pyme': 12,
        'acueducto bodega': 2, 'hosting mantis': 12, 'segunda dotacion': 4,
        'tercera dotacion': 4, 'primera dotacion': 4,
        'cesantias': 12, 'intereses a las cesantias': 12
      }
      const fechaEdits = {}
      // Primero verificar en el mes ORIGEN cuáles estaban Cancelado
      const canceladoEnOrigen = new Map() // nombre → fechaVencimiento
      if (parsedSrc && parsedSrc.sectionCXP) {
        for (const sub of parsedSrc.sectionCXP) {
          for (const r of sub.rows) {
            const nombreNorm = String(r.proveedor || '').trim().toLowerCase().replace(/\s+/g, ' ')
            const estado = String(r.estado || '').trim().toLowerCase()
            const meses = PERIODICIDAD[nombreNorm] || Object.entries(PERIODICIDAD).find(([k]) => nombreNorm.includes(k))?.[1]
            if (meses && estado === 'cancelado') {
              canceladoEnOrigen.set(nombreNorm, r.fechaVencimiento)
            }
          }
        }
      }
      // Luego actualizar las fechas en la hoja NUEVA (solo match exacto o mismo PERIODICIDAD key)
      if (canceladoEnOrigen.size > 0) {
        const newWbTemp = XLSX.read(finalBuf, { type: 'array' })
        const newRows = XLSX.utils.sheet_to_json(newWbTemp.Sheets[nextName], { header: 1, defval: '' })
        const newParsed = parseMonthSheet(newRows)
        if (newParsed && newParsed.sectionCXP) {
          for (const sub of newParsed.sectionCXP) {
            for (const r of sub.rows) {
              const nombreNorm = String(r.proveedor || '').trim().toLowerCase().replace(/\s+/g, ' ')
              // Solo si el concepto en sí tiene periodicidad
              const meses = PERIODICIDAD[nombreNorm] || Object.entries(PERIODICIDAD).find(([k]) => nombreNorm.includes(k))?.[1]
              if (!meses) continue
              // Buscar si estaba Cancelado en el origen
              const fechaOrig = canceladoEnOrigen.get(nombreNorm)
              if (fechaOrig) {
                const partes = fechaOrig.split('/')
                if (partes.length === 3) {
                  const d = new Date(Number(partes[2]), Number(partes[1]) - 1 + meses, Number(partes[0]))
                  const nuevaFecha = String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + d.getFullYear()
                  const serial = dateStringToSerial(nuevaFecha)
                  if (serial) {
                    if (!fechaEdits[nextName]) fechaEdits[nextName] = {}
                    fechaEdits[nextName][`${r._row},5`] = serial
                    // Rotar nombre de dotaciones
                    const nombreRaw = String(r.proveedor || '').trim()
                    const dotMatch = nombreRaw.match(/^(primera|segunda|tercera)\s+dotacion\s+(\d{4})$/i)
                    if (dotMatch) {
                      const orden = ['primera', 'segunda', 'tercera']
                      const idx = orden.indexOf(dotMatch[1].toLowerCase())
                      const year = parseInt(dotMatch[2])
                      const nextIdx = (idx + 1) % 3
                      const nextYear = nextIdx === 0 ? year + 1 : year
                      const nuevoNombre = orden[nextIdx].charAt(0).toUpperCase() + orden[nextIdx].slice(1) + ' dotacion ' + nextYear
                      fechaEdits[nextName][`${r._row},3`] = nuevoNombre
                    }
                  }
                }
              }
            }
          }
        }
        if (Object.keys(fechaEdits).length > 0) {
          finalBuf = await patchXlsx(finalBuf, fechaEdits, {})
        }
      }

      await writable.write(finalBuf)
      await writable.close()
      const newWb = XLSX.read(finalBuf, { type: 'array' })
      setSlots(prev => ({
        ...prev,
        colombia: { ...prev.colombia, rawBuffer: finalBuf, workbook: newWb, selectedSheet: nextName, pendingEdits: {}, pendingInsertions: {} }
      }))
      setSyncStatus('synced')
      setSyncMessage(`Hoja "${nextName}" creada (${canceladoRows.length} filas omitidas${restaurados ? ', valores originales restaurados' : ''})`)
    } catch (err) {
      setError('Error al crear la hoja: ' + err.message)
    }
  }

  // ──────────────────────────────────────────────
  // Generar Excel de Resumen del mes actual
  // ──────────────────────────────────────────────
  const handleGenerarResumen = async () => {
    if (!workbook || !selectedSheet || !isMonthSheet(selectedSheet)) {
      setError('Seleccioná una hoja de mes primero.')
      return
    }
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[selectedSheet], { header: 1, defval: '' })
    const parsed = parseMonthSheet(rows)
    if (!parsed || !parsed.sectionCXP) {
      setError('No se pudo parsear la hoja actual.')
      return
    }

    // Construir datos filtrando filas no deseadas
    const comprasSub = parsed.sectionCXP.find(s => s.title.toUpperCase().includes('COMPRAS CON FACTURAS'))
    const dataRows = []
    const skipPatterns = /^(total egresos|saldo|%\s*$)/i

    for (const sub of parsed.sectionCXP) {
      if (skipPatterns.test(sub.title)) continue
      const isProveedor = sub === comprasSub
      for (const r of sub.rows) {
        const nombreOriginal = String(r.proveedor || '')
        const nombre = nombreOriginal.trim()
        if (!nombre || skipPatterns.test(nombre)) continue
        const estado = String(r.estado || '').trim().toLowerCase()
        if (estado !== 'sin pagar') continue
        const tipo = isProveedor ? 'PROVEDOR' : 'GASTOS'
        const valor = typeof r.valor === 'number' ? r.valor : 0
        const fecha = r.fechaVencimiento || ''
        dataRows.push({ tipo, nombre: nombreOriginal, valor, fecha })
      }
    }

    // Cargar Hoja2 del template para verificar mapeos (como fallback)
    let hoja2Names = new Set()
    try {
      const tmplResp = await fetch('/FLUJO DE CAJA RESUMEN.xlsx')
      if (tmplResp.ok) {
        const tmplBuf = await tmplResp.arrayBuffer()
        const tmplWb = XLSX.read(tmplBuf, { type: 'array' })
        const hs2 = tmplWb.Sheets['Hoja2']
        if (hs2) {
          const hs2Rows = XLSX.utils.sheet_to_json(hs2, { header: 1, defval: '' })
          for (let i = 0; i < hs2Rows.length; i++) {
            if (hs2Rows[i][0]) {
              hoja2Names.add(String(hs2Rows[i][0]).trim().toLowerCase())
            }
          }
        }
      }
    } catch (e) { /* continuar sin Hoja2 */ }

    // Verificar mapeos: primero MAPEO_DESCUENTOS local, luego Hoja2 del template
    const normalizar = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim()
    const sinMapeo = dataRows.filter(r => {
      const norm = normalizar(r.nombre)
      return !mapeoDescuentos[norm] && !hoja2Names.has(norm)
    })

    if (sinMapeo.length > 0) {
      setResumenModal({
        sinMapeo: sinMapeo.map(r => ({ nombre: r.nombre, valor: r.valor, fecha: r.fecha })),
        dataRows,
        hoja2Names,
        onConfirm: async (nuevosMapeos) => {
          setResumenModal(null)
          // Guardar nuevos mapeos en MAPEO_DESCUENTOS
          await guardarMapeos(nuevosMapeos)
          descargarResumen(dataRows, hoja2Names, nuevosMapeos)
        }
      })
      return
    }

    await descargarResumen(dataRows, hoja2Names, {})
  }

  async function descargarResumen(dataRows, hoja2Names, nuevosMapeos) {
    try {
      const tmplResp = await fetch('/FLUJO DE CAJA RESUMEN.xlsx')
      if (!tmplResp.ok) throw new Error('No se pudo cargar la plantilla')
      const tmplBuf = await tmplResp.arrayBuffer()

      const dataArray = dataRows.map(r => [r.tipo, r.nombre, r.valor, r.fecha])
      const newBuf = await generarResumenXlsx(tmplBuf, dataArray, nuevosMapeos)

      const blob = new Blob([newBuf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
      const a = document.createElement('a')
      a.href = URL.createObjectURL(blob)
      a.download = 'Flujo de Caja Resumen ' + selectedSheet + '.xlsx'
      a.click()
      URL.revokeObjectURL(a.href)
    } catch (err) {
      setError('Error al generar resumen: ' + err.message)
    }
  }

  async function guardarMapeos(nuevosMapeos) {
    const entries = Object.entries(nuevosMapeos)
    if (entries.length === 0 || !rawBufferRef.current) return

    try {
      let buf = await ensureMapeoSheet(rawBufferRef.current)
      const wb = XLSX.read(buf, { type: 'array' })
      let lastRow = 0
      const sheet = wb.Sheets['MAPEO_DESCUENTOS']
      if (sheet) {
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })
        lastRow = rows.length - 1
      }

      const insertions = []
      for (const [nombre, categoria] of entries) {
        insertions.push({
          id: 'mapeo-' + Date.now() + '-' + Math.random().toString(36).slice(2, 5),
          insertAfterRow: lastRow,
          sectionKey: 'MAPEO_DESCUENTOS',
          cells: { 0: nombre, 1: categoria }
        })
      }

      const pendingIns = { 'MAPEO_DESCUENTOS': insertions }
      const newBuf = await patchXlsx(buf, {}, pendingIns)

      const handle = fileHandleRef.current
      if (handle) {
        await requestPermission(handle)
        const writable = await handle.createWritable()
        await writable.write(newBuf)
        await writable.close()
      }

      const newWb = XLSX.read(newBuf, { type: 'array' })
      setMapeoDescuentos(prev => ({ ...prev, ...Object.fromEntries(entries.map(([k, v]) => [k.toLowerCase(), v])) }))
      setSlots(prev => ({ ...prev, colombia: { ...prev.colombia, rawBuffer: newBuf, workbook: newWb } }))
    } catch (e) {
      console.error('Error al guardar mapeos:', e)
    }
  }

  // ──────────────────────────────────────────────
  // Abonos
  // ──────────────────────────────────────────────
  const handleOpenAbonoModal = (proveedor, factura, valorOriginal) => {
    setAbonoTarget({ proveedor, factura, valorOriginal })
    setAbonosModalOpen(true)
  }

  const handleConfirmarAbono = async (fecha, valor, metodo) => {
    if (!abonoTarget || !selectedSheet) return
    setAbonosModalOpen(false)
    const target = { ...abonoTarget }
    setAbonoTarget(null)

    try {
      const buf = await ensureAbonosSheet(rawBufferRef.current)
      const wb = XLSX.read(buf, { type: 'array' })

      let lastRow = 0
      if (wb.Sheets[ABONOS_SHEET]) {
        const abRows = XLSX.utils.sheet_to_json(wb.Sheets[ABONOS_SHEET], { header: 1, defval: '' })
        lastRow = abRows.length - 1
      }

      const valorAntes = target.valorOriginal
      const insertion = buildAbonoInsertion(selectedSheet, target.proveedor, target.factura || '', fecha, valor, metodo, valorAntes)
      insertion.insertAfterRow = lastRow

      // Encontrar la fila y calcular nuevo valor
      const monthRows = XLSX.utils.sheet_to_json(wb.Sheets[selectedSheet], { header: 1, defval: '' })
      const parsed = parseMonthSheet(monthRows)
      const edits = {}
      const nuevoValor = valorAntes - valor
      if (parsed && parsed.sectionCXP) {
        for (const sub of parsed.sectionCXP) {
          for (const r of sub.rows) {
            if (String(r.proveedor || '').trim() === String(target.proveedor || '').trim() &&
                String(r.factura || '').trim() === String(target.factura || '').trim()) {
              edits[`${r._row},4`] = nuevoValor
              break
            }
          }
        }
      }

      const pendingIns = { [ABONOS_SHEET]: [insertion] }
      const sheetEdits = Object.keys(edits).length > 0 ? { [selectedSheet]: edits } : {}
      let newBuf = await patchXlsx(buf, sheetEdits, pendingIns)

      // Recalcular subtotales
      const wbCheck = XLSX.read(newBuf, { type: 'array' })
      const rowsCheck = XLSX.utils.sheet_to_json(wbCheck.Sheets[selectedSheet], { header: 1, defval: '' })
      const parsedCheck = parseMonthSheet(rowsCheck)
      if (parsedCheck && parsedCheck.sectionCXP) {
        const subEdits = buildSubtotalEdits(parsedCheck.sectionCXP)
        if (Object.keys(subEdits).length > 0) {
          newBuf = await patchXlsx(newBuf, { [selectedSheet]: subEdits }, {})
        }
      }

      const handle = fileHandleRef.current
      if (handle) {
        await requestPermission(handle)
        const writable = await handle.createWritable()
        await writable.write(newBuf)
        await writable.close()
      }

      const newWb = XLSX.read(newBuf, { type: 'array' })
      setAbonos(parseAbonosFromWorkbook(newWb))
      setSlots(prev => ({ ...prev, colombia: { ...prev.colombia, rawBuffer: newBuf, workbook: newWb, pendingEdits: {}, pendingInsertions: {} } }))
      setSyncStatus('synced')
      setSyncMessage('Abono registrado. Valor actualizado en el Excel.')
    } catch (err) {
      setError('Error al guardar abono: ' + err.message)
    }
  }

  // ── Cargar archivo en un slot (compartido entre Open y Reconnect) ──
  async function loadIntoSlot(handle, slotKey) {
    const file   = await handle.getFile()
    const buffer = await file.arrayBuffer()
    const wb     = XLSX.read(buffer, { type: 'array' })
    setAbonos(parseAbonosFromWorkbook(wb))
    setMapeoDescuentos(parseMapeoFromWorkbook(wb))
    setSlots(prev => ({
      ...prev,
      [slotKey]: emptySlot(handle, file.name, buffer, wb)
    }))
    setActiveSlot(slotKey)
  }

  // ──────────────────────────────────────────────
  // Abrir archivo para un slot específico
  // ──────────────────────────────────────────────
  const handleOpenFile = async (slotKey) => {
    setError('')
    if (!window.showOpenFilePicker) {
      setError('Tu navegador no soporta la File System Access API. Usa Chrome o Edge.')
      return
    }
    try {
      const [handle] = await window.showOpenFilePicker({
        types: [{ description: 'Archivo Excel',
          accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }],
        multiple: false
      })
      await loadIntoSlot(handle, slotKey)
      saveHandle(`slot-${slotKey}`, handle).catch(() => {})
      setSavedHandles(prev => ({ ...prev, [slotKey]: null }))
    } catch (err) {
      if (err.name !== 'AbortError') setError('Error al abrir el archivo: ' + err.message)
    }
  }

  // ──────────────────────────────────────────────
  // Reconectar desde handle guardado
  // ──────────────────────────────────────────────
  const handleReconnect = async (slotKey) => {
    const handle = savedHandles[slotKey]
    if (!handle) return
    setReconnecting(true)
    setError('')
    try {
      const granted = await requestPermission(handle)
      if (!granted) {
        setError('Permiso denegado. Usa el botón "Abrir" para buscar el archivo manualmente.')
        setReconnecting(false)
        return
      }
      await loadIntoSlot(handle, slotKey)
      setSavedHandles(prev => ({ ...prev, [slotKey]: null }))
    } catch (err) {
      setError('No se pudo reconectar: ' + err.message)
    }
    setReconnecting(false)
  }

  // ──────────────────────────────────────────────
  // Cerrar un slot
  // ──────────────────────────────────────────────
  const handleCloseFile = (slotKey) => {
    setSlots(prev => ({ ...prev, [slotKey]: null }))
    clearHandle(`slot-${slotKey}`).catch(() => {})
    setSavedHandles(prev => ({ ...prev, [slotKey]: null }))
    if (activeSlot === slotKey) {
      // Pasar al primer slot que aún tenga archivo cargado
      const next = SLOT_KEYS.find(k => k !== slotKey && slots[k] !== null) ?? null
      setActiveSlot(next)
    }
  }

  // ── setSelectedSheet en el slot activo ──
  const setSelectedSheet = (name) => {
    if (!activeSlot) return
    if (isMonthSheet(name)) setLastMonthSheet(name)
    setSlots(prev => ({ ...prev, [activeSlot]: { ...prev[activeSlot], selectedSheet: name } }))
  }

  // ──────────────────────────────────────────────
  // Edición de celdas
  // ──────────────────────────────────────────────
  const handleCellEdit = (rowIdx, colIdx, draft) => {
    if (!activeSlot) return
    const trimmed = typeof draft === 'string' ? draft.trim() : draft
    let value
    if (trimmed === '' || trimmed === null || trimmed === undefined) { value = '' }
    else { const n = Number(trimmed); value = isNaN(n) ? trimmed : n }

    // Si cambia estado a Cancelado, pedir método de pago
    if (colIdx === 6 && String(trimmed).toLowerCase() === 'cancelado') {
      setCanceladoModal({
        rowIdx,
        colIdx,
        value,
        onSave: (metodo) => {
          setCanceladoModal(null)
          setSlots(prev => {
            const s = prev[activeSlot]
            return { ...prev, [activeSlot]: { ...s,
              pendingEdits: { ...s.pendingEdits,
                [selectedSheet]: {
                  ...(s.pendingEdits[selectedSheet] || {}),
                  [`${rowIdx},6`]: value,
                  [`${rowIdx},10`]: metodo  // DEUDA PENDIENTE
                }
              }
            }}
          })
        }
      })
      return
    }

    setSlots(prev => {
      const s = prev[activeSlot]
      return { ...prev, [activeSlot]: { ...s,
        pendingEdits: { ...s.pendingEdits,
          [selectedSheet]: { ...(s.pendingEdits[selectedSheet] || {}), [`${rowIdx},${colIdx}`]: value }
        }
      }}
    })
  }

  // ──────────────────────────────────────────────
  // Inserción de filas
  // ──────────────────────────────────────────────
  const handleAddRow = (insertAfterRow, sectionKey) => {
    if (!activeSlot) return
    const id = `new-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`
    setSlots(prev => {
      const s = prev[activeSlot]
      return { ...prev, [activeSlot]: { ...s,
        pendingInsertions: { ...s.pendingInsertions,
          [selectedSheet]: [...(s.pendingInsertions[selectedSheet] || []), { id, insertAfterRow, cells: {}, sectionKey }]
        }
      }}
    })
  }

  const handleInsertedRowEdit = (id, colIdx, value) => {
    if (!activeSlot) return
    
    // Si cambia estado a Cancelado, pedir método de pago
    if (colIdx === 6 && String(value).toLowerCase() === 'cancelado') {
      setCanceladoModal({
        rowIdx: null,
        colIdx,
        value,
        insertionId: id,
        onSave: (metodo) => {
          setCanceladoModal(null)
          setSlots(prev => {
            const s = prev[activeSlot]
            return { ...prev, [activeSlot]: { ...s,
              pendingInsertions: { ...s.pendingInsertions,
                [selectedSheet]: (s.pendingInsertions[selectedSheet] || []).map(ins =>
                  ins.id === id
                    ? { ...ins, cells: { ...ins.cells, 6: value, 10: metodo } }
                    : ins
                )
              }
            }}
          })
        }
      })
      return
    }
    
    setSlots(prev => {
      const s = prev[activeSlot]
      return { ...prev, [activeSlot]: { ...s,
        pendingInsertions: { ...s.pendingInsertions,
          [selectedSheet]: (s.pendingInsertions[selectedSheet] || []).map(ins =>
            ins.id === id ? { ...ins, cells: { ...ins.cells, [colIdx]: value } } : ins
          )
        }
      }}
    })
  }

  const handleDeleteInsertedRow = (id) => {
    if (!activeSlot) return
    setSlots(prev => {
      const s = prev[activeSlot]
      return { ...prev, [activeSlot]: { ...s,
        pendingInsertions: { ...s.pendingInsertions,
          [selectedSheet]: (s.pendingInsertions[selectedSheet] || []).filter(ins => ins.id !== id)
        }
      }}
    })
  }

  const handleAddSubsection = (nombre) => {
    if (!activeSlot || !workbook || !selectedSheet) return
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[selectedSheet], { header: 1, defval: '' })
    const parsed = parseMonthSheet(rows)
    if (!parsed || !parsed.sectionCXP || parsed.sectionCXP.length === 0) return

    const lastSubsection = parsed.sectionCXP[parsed.sectionCXP.length - 1]
    const lastRow = lastSubsection.rows.length > 0
      ? lastSubsection.rows[lastSubsection.rows.length - 1]._row
      : lastSubsection._row

    const id = `subsec-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`
    const nombreUpper = nombre.toUpperCase()

    setSlots(prev => {
      const s = prev[activeSlot]
      return { ...prev, [activeSlot]: { ...s,
        pendingInsertions: { ...s.pendingInsertions,
          [selectedSheet]: [...(s.pendingInsertions[selectedSheet] || []), {
            id,
            insertAfterRow: lastRow,
            cells: { 3: nombreUpper, 4: 0 },
            sectionKey: `cxp:${nombreUpper}`,
            isSubsection: true
          }]
        }
      }}
    })
    setNuevaSubseccionModal(false)
  }

  // ──────────────────────────────────────────────
  // Guardar archivo
  // ──────────────────────────────────────────────
  const handleSave = async () => {
    if (!fileHandle || !rawBuffer || !activeSlot) return
    setSaveStatus('saving')
    setError('')
    try {
      let buf = await patchXlsx(rawBuffer, pendingEdits, pendingInsertions)

      // Actualizar fórmulas de subtotales en subsecciones CXP
      const monthSheets = Object.keys({ ...pendingEdits, ...pendingInsertions })
      const formulaEdits = {}
      for (const sheet of monthSheets) {
        if (!isMonthSheet(sheet)) continue
        const wbTemp = XLSX.read(buf, { type: 'array' })
        const rows = XLSX.utils.sheet_to_json(wbTemp.Sheets[sheet], { header: 1, defval: '' })
        const parsed = parseMonthSheet(rows)
        if (parsed && parsed.sectionCXP) {
          const subEdits = buildSubtotalEdits(parsed.sectionCXP)
          if (Object.keys(subEdits).length > 0) {
            formulaEdits[sheet] = subEdits
          }
        }
      }
      if (Object.keys(formulaEdits).length > 0) {
        buf = await patchXlsx(buf, formulaEdits, {})
      }

      const writable = await fileHandle.createWritable()
      await writable.write(buf)
      await writable.close()
      const newWb = XLSX.read(buf, { type: 'array' })
      setSlots(prev => ({ ...prev, [activeSlot]: {
        ...prev[activeSlot], rawBuffer: buf, workbook: newWb,
        pendingEdits: {}, pendingInsertions: {}
      }}))
      setAbonos(parseAbonosFromWorkbook(newWb))
      setSaveStatus('saved')
      setTimeout(() => setSaveStatus('idle'), 2500)
    } catch (err) {
      setError('Error al guardar: ' + err.message)
      setSaveStatus('idle')
    }
  }

  const handleDiscard = () => {
    if (!activeSlot) return
    setSlots(prev => ({ ...prev, [activeSlot]: {
      ...prev[activeSlot], pendingEdits: {}, pendingInsertions: {}
    }}))
  }

  // ──────────────────────────────────────────────
  // Helpers UI
  // ──────────────────────────────────────────────
  const sheetExists = (name) => workbook?.SheetNames.includes(name) ?? false

  const displayTabs = useMemo(() => {
    if (!workbook) return []
    const wbSheets = workbook.Workbook?.Sheets ?? []
    const hiddenSet = new Set(
      wbSheets.reduce((acc, s, i) => {
        if (s.Hidden && s.Hidden > 0) acc.push(workbook.SheetNames[i])
        return acc
      }, [])
    )
    const visibleNames = workbook.SheetNames.filter(n => !hiddenSet.has(n))
    const monthTabs = visibleNames.filter(isMonthSheet).sort((a, b) => monthSheetIndex(a) - monthSheetIndex(b))
    const otherTabs = visibleNames.filter(s => !isMonthSheet(s))
    return [...monthTabs, ...otherTabs]
  }, [workbook])

  const sheetRows = useMemo(() => {
    if (!workbook || !selectedSheet || !workbook.Sheets[selectedSheet]) return []
    return XLSX.utils.sheet_to_json(workbook.Sheets[selectedSheet], { header: 1, defval: '' })
  }, [workbook, selectedSheet])

  // ──────────────────────────────────────────────
  // Render
  // ──────────────────────────────────────────────
  return (
    <div className="excel-editor">

      {/* ── Barra de herramientas ── */}
      <div className="excel-toolbar">
        <div className="toolbar-left">

          {slots.colombia ? (
            <div className="country-tab active" title={slots.colombia.fileName}>
              <span className="country-tab-label">{SLOT_LABELS.colombia}</span>
              <span className="country-tab-filename">{slots.colombia.fileName}</span>
              <button
                className="btn-close-country"
                onClick={() => handleCloseFile('colombia')}
                title="Cerrar"
              >
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
              </button>
            </div>
          ) : savedHandles.colombia ? (
            <div className="slot-reconnect-group">
              <button
                className="btn-toolbar btn-slot-reconnect"
                onClick={() => handleReconnect('colombia')}
                disabled={reconnecting}
                title={`Reconectar: ${savedHandles.colombia.name}`}
              >
                <span>⚡</span>
                <span className="slot-reconnect-label">{SLOT_LABELS.colombia}</span>
                <span className="slot-reconnect-filename">{savedHandles.colombia.name}</span>
                {reconnecting && <span className="slot-reconnect-spinner">…</span>}
              </button>
              <button
                className="btn-slot-open-other"
                onClick={() => handleOpenFile('colombia')}
                title="Abrir otro archivo"
              >
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                  <polyline points="14 2 14 8 20 8"/>
                </svg>
              </button>
            </div>
          ) : (
            <button className="btn-toolbar btn-open" onClick={() => handleOpenFile('colombia')}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14 2 14 8 20 8"/>
              </svg>
              Abrir Excel
            </button>
          )}
        </div>

        {workbook && (
          <div className="toolbar-right">
            {isMonthSheet(selectedSheet) && (
              <button
                className="btn-toolbar btn-sync"
                onClick={handleSyncNow}
                disabled={syncStatus === 'syncing'}
                title="Consultar Google Sheets e importar facturas nuevas del mes visible"
              >
                {syncStatus === 'syncing' ? (
                  <><span className="sync-spinner-sm"></span> Sincronizando…</>
                ) : (
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ marginRight: '4px' }}>
                    <polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/>
                  </svg>
                )}
                Actualizar registros
              </button>
            )}
            {isMonthSheet(selectedSheet) && (
              <button
                className="btn-toolbar btn-next-month"
                onClick={handleCreateNextMonth}
                title="Crear hoja del mes siguiente copiando la estructura actual"
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ marginRight: '4px' }}>
                  <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M12 8v8M8 12h8"/>
                </svg>
                Crear mes siguiente
              </button>
            )}
            {isMonthSheet(selectedSheet) && (
              <button
                className="btn-toolbar btn-resumen"
                onClick={handleGenerarResumen}
                title="Generar Excel de resumen del mes actual"
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ marginRight: '4px' }}>
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                  <polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>
                </svg>
                Generar Resumen
              </button>
            )}
            {saveStatus === 'saved' && !hasChanges && (
              <span className="save-ok-badge">✓ Guardado</span>
            )}
          </div>
        )}

        {/* Botones de guardado */}
        {hasChanges && (
          <div className="toolbar-right">
            {totalEdits > 0 && (
              <span className="changes-badge">{totalEdits} cambio{totalEdits !== 1 ? 's' : ''}</span>
            )}
            {totalInsertions > 0 && (
              <span className="changes-badge changes-badge-new">{totalInsertions} fila{totalInsertions !== 1 ? 's' : ''} nueva{totalInsertions !== 1 ? 's' : ''}</span>
            )}
            <button className="btn-toolbar btn-discard" onClick={handleDiscard}>Descartar</button>
            <button
              className={`btn-toolbar btn-save ${saveStatus === 'saving' ? 'saving' : ''}`}
              onClick={handleSave}
              disabled={saveStatus === 'saving'}
            >
              {saveStatus === 'saving' ? 'Guardando…' : 'Guardar en Excel'}
            </button>
          </div>
        )}
        {saveStatus === 'saved' && !hasChanges && (
          <div className="toolbar-right"><span className="save-ok-badge">✓ Guardado</span></div>
        )}
      </div>

      {/* ── Sync status ── */}
      {syncStatus === 'syncing' && workbook && activeSlot === 'colombia' && (
        <div className="alert alert-sync">
          <span className="sync-spinner"></span>
          Sincronizando con Google Sheets…
        </div>
      )}
      {syncStatus === 'synced' && syncMessage && workbook && activeSlot === 'colombia' && (
        <div className="alert alert-sync-success">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>
          </svg>
          {syncMessage}
        </div>
      )}
      {syncStatus === 'error' && syncMessage && workbook && activeSlot === 'colombia' && (
        <div className="alert alert-sync-error">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
          </svg>
          {syncMessage}
        </div>
      )}

      {/* ── Error ── */}
      {error && (
        <div className="alert alert-error">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
          </svg>
          {error}
        </div>
      )}

      {/* ── Drop-zone (sin archivo cargado) ── */}
      {!hasAnyFile && (
        <div className="dual-dropzone">
          <div
            className="dropzone-card"
            onClick={() => handleOpenFile('colombia')}
          >
            <div className="dropzone-card-icon">
              <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/>
              </svg>
            </div>
            <span className="dropzone-card-label">Flujo de Caja</span>
            <span className="dropzone-card-hint">Haz clic para buscar el archivo .xlsx</span>
          </div>
        </div>
      )}

      {/* ── Contenido del slot activo ── */}
      {workbook && (
        <>
          <div className="sheet-selector-bar">
            {displayTabs.map(name => {
              const exists  = sheetExists(name)
              const isMonth = isMonthSheet(name)
              return (
                <button
                  key={name}
                  className={['sheet-tab', selectedSheet === name ? 'active' : '', !exists ? 'unavailable' : '', !isMonth ? 'tab-special' : ''].filter(Boolean).join(' ')}
                  onClick={() => exists && setSelectedSheet(name)}
                  disabled={!exists}
                  title={!exists ? `${name} — sin datos` : name}
                >
                  {name}
                </button>
              )
            })}
            <button
              className={'sheet-tab tab-special' + (selectedSheet === '__ABONOS__' ? ' active' : '')}
              onClick={() => setSelectedSheet('__ABONOS__')}
            >
              Abonos
            </button>
          </div>

          {selectedSheet === '__ABONOS__' && (
            <AbonosPanel abonos={abonos} sheetName={lastMonthSheet || selectedSheet} onOpenAbono={handleOpenAbonoModal} />
          )}

          {selectedSheet !== '__ABONOS__' && isMonthSheet(selectedSheet) && sheetExists(selectedSheet) && (
            <MonthViewer
              rows={sheetRows} sheetName={selectedSheet}
              edits={pendingEdits[selectedSheet] || {}}
              onCellEdit={handleCellEdit}
              insertions={pendingInsertions[selectedSheet] || []}
              onAddRow={handleAddRow}
              onInsertedRowEdit={handleInsertedRowEdit}
              onDeleteInsertedRow={handleDeleteInsertedRow}
              abonos={abonos}
              onOpenAbono={handleOpenAbonoModal}
              onAddSubsection={() => setNuevaSubseccionModal(true)}
            />
          )}

          {!isMonthSheet(selectedSheet) && sheetExists(selectedSheet) && (
            <div className="coming-soon-panel">
              <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                <circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>
              </svg>
              <p>La vista estructurada de <strong>{selectedSheet}</strong> estará disponible próximamente.</p>
            </div>
          )}
        </>
      )}

      {/* Modal: registros sin mapeo en Hoja2 */}
      {resumenModal && (
        <ResumenMapeoModal
          sinMapeo={resumenModal.sinMapeo}
          onCancel={() => setResumenModal(null)}
          onConfirm={(mapeos) => resumenModal.onConfirm(mapeos)}
        />
      )}

      {abonosModalOpen && abonoTarget && (
        <AbonoModal
          target={abonoTarget}
          onConfirm={handleConfirmarAbono}
          onCancel={() => { setAbonosModalOpen(false); setAbonoTarget(null) }}
        />
      )}

      {canceladoModal && (
        <CanceladoModal
          onConfirm={(metodo) => canceladoModal.onSave(metodo)}
          onCancel={() => setCanceladoModal(null)}
        />
      )}

      {nuevaSubseccionModal && (
        <NuevaSubseccionModal
          onConfirm={handleAddSubsection}
          onCancel={() => setNuevaSubseccionModal(false)}
        />
      )}
    </div>
  )
}

function ResumenMapeoModal({ sinMapeo, onCancel, onConfirm }) {
  const [selecciones, setSelecciones] = useState({})
  const categorias = ['OBLIGATORIO', 'MANEJABLE', 'IMPORTANTE']

  const handleConfirmar = () => {
    const mapeos = {}
    for (const item of sinMapeo) {
      const sel = selecciones[item.nombre] || 'OBLIGATORIO'
      mapeos[item.nombre] = sel
    }
    onConfirm(mapeos)
  }

  const todosSeleccionados = sinMapeo.every(item => selecciones[item.nombre])

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal-content modal-resumen" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h3>Registros sin categoría en Hoja2</h3>
          <p>{sinMapeo.length} registro(s) no tienen mapeo. Seleccioná una categoría para cada uno:</p>
        </div>
        <div className="modal-body">
          <div className="resumen-lista">
            {sinMapeo.map((item, i) => (
              <div key={i} className="resumen-item">
                <span className="resumen-item-nombre">{item.nombre}</span>
                <span className="resumen-item-valor">{typeof item.valor === 'number' ? '$' + item.valor.toLocaleString('es-CO') : ''}</span>
                <select
                  className="resumen-select"
                  value={selecciones[item.nombre] || ''}
                  onChange={e => setSelecciones(prev => ({ ...prev, [item.nombre]: e.target.value }))}
                >
                  <option value="">— Elegir —</option>
                  {categorias.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
            ))}
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn-toolbar btn-discard" onClick={onCancel}>Cancelar</button>
          <button
            className="btn-toolbar btn-save"
            onClick={handleConfirmar}
            disabled={!todosSeleccionados}
          >
            Confirmar y descargar
          </button>
        </div>
      </div>
    </div>
  )
}

function AbonosPanel({ abonos, sheetName, onOpenAbono }) {
  const items = sheetName ? getAllAbonosForMonth(abonos, sheetName) : []

  function formatFecha(val) {
    if (!val) return '—'
    const num = parseFloat(val)
    if (!isNaN(num) && num > 40000 && num < 60000) {
      return excelDateToString(num)
    }
    // yyyy-mm-dd → dd/mm/yyyy
    const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(val))
    if (iso) return iso[3] + '/' + iso[2] + '/' + iso[1]
    return val
  }

  if (!sheetName || items.length === 0) {
    return (
      <div className="coming-soon-panel">
        <p>No hay abonos registrados para este mes.</p>
      </div>
    )
  }

  return (
    <div className="viewer-section">
      <div className="viewer-section-header">
        <div>
          <h3 className="viewer-section-title">Abonos · {sheetName}</h3>
          <p className="viewer-section-sub">{items.length} abono{items.length !== 1 ? 's' : ''} registrado{items.length !== 1 ? 's' : ''}</p>
        </div>
      </div>
      <div className="table-scroll-wrapper">
        <table className="viewer-table">
          <thead>
            <tr>
              <th>Proveedor</th>
              <th>Factura</th>
              <th>Fecha</th>
              <th className="col-right">Valor Antes</th>
              <th className="col-right">Abono</th>
              <th className="col-right">Resultante</th>
              <th>Método</th>
            </tr>
          </thead>
          <tbody>
            {items.map((item, i) => {
              const valorResultante = item.valorAntes != null ? item.valorAntes - item.valor : null
              return (
              <tr key={i} className={i % 2 === 0 ? 'tr-even' : ''}>
                <td className="td-desc">{item.proveedor}</td>
                <td className="td-factura">{item.factura || '—'}</td>
                <td>{formatFecha(item.fecha)}</td>
                <td className="col-right td-valor">{item.valorAntes != null ? '$' + Number(item.valorAntes).toLocaleString('es-CO') : '—'}</td>
                <td className="col-right td-valor">-${Number(item.valor).toLocaleString('es-CO')}</td>
                <td className="col-right td-valor" style={{fontWeight:700}}>{valorResultante != null ? '$' + valorResultante.toLocaleString('es-CO') : '—'}</td>
                <td>{item.metodo}</td>
              </tr>
            )})}
          </tbody>
        </table>
      </div>
    </div>
  )
}

function CanceladoModal({ onConfirm, onCancel }) {
  const [metodo, setMetodo] = useState('')
  const metodos = ['Efectivo', 'Transferencia', 'Cruce de cuentas', 'Tarjeta de crédito']

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal-content modal-cancelado" onClick={e => e.stopPropagation()}>
        <h3>Método de Pago</h3>
        <p>¿Cómo se realizó el pago?</p>
        <select
          className="cancelado-input"
          value={metodo}
          onChange={e => setMetodo(e.target.value)}
          autoFocus
        >
          <option value="">Selecciona un método de pago</option>
          {metodos.map(m => <option key={m} value={m}>{m}</option>)}
        </select>
        <div className="modal-footer">
          <button className="btn-toolbar btn-discard" onClick={onCancel}>Cancelar</button>
          <button className="btn-toolbar btn-save" onClick={() => onConfirm(metodo)} disabled={!metodo}>Guardar</button>
        </div>
      </div>
    </div>
  )
}

function NuevaSubseccionModal({ onConfirm, onCancel }) {
  const [nombre, setNombre] = useState('')

  const handleConfirm = () => {
    if (!nombre.trim()) return
    onConfirm(nombre.trim())
  }

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal-content modal-nueva-subseccion" onClick={e => e.stopPropagation()}>
        <h3>Nueva Subsección</h3>
        <p>Ingresa el nombre de la nueva subsección:</p>
        <input
          type="text"
          className="subseccion-input"
          placeholder="Ej: GASTOS OPERATIVOS"
          value={nombre}
          onChange={e => setNombre(e.target.value)}
          autoFocus
          onKeyDown={e => { if (e.key === 'Enter' && nombre.trim()) handleConfirm() }}
        />
        <div className="modal-footer">
          <button className="btn-toolbar btn-discard" onClick={onCancel}>Cancelar</button>
          <button className="btn-toolbar btn-save" onClick={handleConfirm} disabled={!nombre.trim()}>Crear</button>
        </div>
      </div>
    </div>
  )
}

export default ExcelEditor

function AbonoModal({ target, onConfirm, onCancel }) {
  const [fecha, setFecha] = useState('')
  const [valor, setValor] = useState('')
  const [metodo, setMetodo] = useState('Efectivo')
  const metodos = ['Efectivo', 'Transferencia', 'Cruce de cuentas', 'Tarjeta de crédito']

  const handleConfirm = () => {
    const v = parseFloat(valor)
    if (!fecha || isNaN(v) || v <= 0) return
    onConfirm(fecha, v, metodo)
  }

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal-content modal-abono" onClick={e => e.stopPropagation()}>
        <h3>Nuevo Abono</h3>
        <p className="modal-abono-concepto">{target?.proveedor}</p>
        {target?.valorOriginal && (
          <p className="modal-abono-original">Valor original: ${typeof target.valorOriginal === 'number' ? target.valorOriginal.toLocaleString('es-CO') : target.valorOriginal}</p>
        )}
        <div className="modal-abono-fields">
          <label>Fecha</label>
          <input type="text" placeholder="dd/mm/aaaa" value={fecha} onChange={e => setFecha(e.target.value)} />
          <label>Valor</label>
          <input type="number" placeholder="0" value={valor} onChange={e => setValor(e.target.value)} />
          <label>Método</label>
          <select value={metodo} onChange={e => setMetodo(e.target.value)}>
            {metodos.map(m => <option key={m} value={m}>{m}</option>)}
          </select>
        </div>
        <div className="modal-footer">
          <button className="btn-toolbar btn-discard" onClick={onCancel}>Cancelar</button>
          <button className="btn-toolbar btn-save" onClick={handleConfirm} disabled={!fecha || !valor}>Guardar Abono</button>
        </div>
      </div>
    </div>
  )
}
