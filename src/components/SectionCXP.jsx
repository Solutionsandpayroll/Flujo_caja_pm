import { useState, useMemo } from 'react'
import { formatCOP } from '../utils/excelParser'
import { getTotalAbonadoForKey, getAllAbonosForMonth } from '../utils/abonosStore'
import EditableCell from './EditableCell'

const ESTADO_OPTIONS_CXP = ['Cancelado', 'Pagar', 'Sin Pagar']
const ESTADO_BADGE_CXP = {
  'cancelado': 'badge badge-success',
  'sin pagar': 'badge badge-warning',
  'pagar':     'badge badge-action',
}

function EstadoBadge({ estado }) {
  if (!estado) return null
  const lower = estado.toLowerCase()
  const cls =
    lower === 'cancelado' || lower === 'pagado'
      ? 'badge badge-success'
      : 'badge badge-warning'
  return <span className={cls}>{estado}</span>
}

function esSinPagar(estado) {
  if (!estado) return false
  const lower = String(estado).trim().toLowerCase()
  return lower === 'sin pagar'
}

function calcularTotalSinPagar(rows, insertions, edits) {
  let total = 0
  for (const row of rows) {
    const estado = String((edits[`${row._row},0`] ?? edits[`${row._row},6`] ?? row.estado) || '')
    if (!esSinPagar(estado)) continue
    const editKey = `${row._row},4`
    const valor = editKey in edits ? edits[editKey] : row.valor
    if (typeof valor === 'number') total += valor
  }
  if (insertions) {
    for (const ins of insertions) {
      const estado = String(ins.cells[6] ?? '')
      if (!esSinPagar(estado)) continue
      if (typeof ins.cells[4] === 'number') total += ins.cells[4]
    }
  }
  return total
}

function fechaAComparar(fechaStr) {
  if (!fechaStr) return Infinity
  const str = String(fechaStr).trim()
  if (!str) return Infinity
  // Serial de Excel
  if (typeof str === 'number' || /^\d{5,6}$/.test(str)) {
    const serial = Number(str)
    if (serial > 40000 && serial < 60000) {
      const date = new Date((serial - 25569) * 86400 * 1000)
      return date.getTime()
    }
  }
  // Formato dd/mm/yyyy
  const partes = str.split('/')
  if (partes.length === 3) {
    const [d, m, y] = partes.map(Number)
    if (d && m && y) {
      const fullYear = y < 100 ? 2000 + y : y
      return new Date(fullYear, m - 1, d).getTime()
    }
  }
  return Infinity
}

function SectionCXP({ subsections, sheetName, edits, onCellEdit, insertions, onAddRow, onInsertedRowEdit, onDeleteInsertedRow, abonos, onOpenAbono, onAddSubsection }) {
  const [activeIdx, setActiveIdx]       = useState(0)
  const [filterEstado, setFilterEstado] = useState('')
  const [search, setSearch]             = useState('')
  const [searchFactura, setSearchFactura] = useState('')
  const [filterFecha, setFilterFecha]   = useState('')

  const current = subsections[activeIdx] ?? null

  const switchSub = (idx) => {
    setActiveIdx(idx)
    setFilterEstado('')
    setSearch('')
    setSearchFactura('')
    setFilterFecha('')
  }

  const filteredRows = useMemo(() => {
    if (!current) return []
    const fechaFiltro = filterFecha ? new Date(filterFecha + 'T23:59:59').getTime() : null
    const rows = current.rows.filter(row => {
      const estado = String((filterEstado ? edits[`${row._row},0`] ?? row.estado : row.estado) || '')
      const prov   = String(row.proveedor || '').toLowerCase()
      const factura = String(row.factura || '').toLowerCase()

      if (filterEstado && estado.toLowerCase() !== filterEstado.toLowerCase()) return false
      if (search && !prov.includes(search.toLowerCase())) return false
      if (searchFactura && !factura.includes(searchFactura.toLowerCase())) return false
      if (fechaFiltro !== null) {
        const fechaRow = fechaAComparar(edits[`${row._row},5`] ?? row.fechaVencimiento)
        if (fechaRow > fechaFiltro) return false
      }
      return true
    })
    // Ordenar por fecha de vencimiento (menor a mayor)
    return [...rows].sort((a, b) => {
      const fechaA = fechaAComparar(edits[`${a._row},5`] ?? a.fechaVencimiento)
      const fechaB = fechaAComparar(edits[`${b._row},5`] ?? b.fechaVencimiento)
      return fechaA - fechaB
    })
    }, [current, filterEstado, search, searchFactura, filterFecha, edits])

  const subInsertions = useMemo(() => {
    if (!current) return []
    return (insertions || [])
      .filter(ins => ins.sectionKey === `cxp:${current.title}`)
      .sort((a, b) => fechaAComparar(a.cells[5]) - fechaAComparar(b.cells[5]))
  }, [current, insertions])

  const calculatedTotal = current
    ? calcularTotalSinPagar(current.rows, subInsertions, edits)
    : 0

  return (
    <div className="viewer-section">
      <div className="viewer-section-header">
        <div>
          <h3 className="viewer-section-title">Cuentas x Pagar Mes</h3>
          <p className="viewer-section-sub">Pagos y obligaciones · {sheetName}</p>
        </div>
        {current && (
          <div className="section-total-badge">
            <span className="section-total-label">Total por pagar</span>
            <span className="section-total-value">{formatCOP(calculatedTotal)}</span>
          </div>
        )}
      </div>

      <div className="cxp-dropdown-bar">
        <label className="cxp-dropdown-label" htmlFor="cxp-select">Subsección</label>
        <select
          id="cxp-select"
          className="cxp-dropdown"
          value={activeIdx}
          onChange={e => switchSub(Number(e.target.value))}
        >
          {subsections.map((sub, i) => {
            const subIns = (insertions || []).filter(ins => ins.sectionKey === `cxp:${sub.title}`)
            const subTotal = calcularTotalSinPagar(sub.rows, subIns, edits)
            return (
              <option key={i} value={i}>
                {sub.title} — {formatCOP(subTotal)}
              </option>
            )
          })}
        </select>
        {onAddSubsection && (
          <button
            className="btn-add-subsection"
            onClick={onAddSubsection}
            title="Crear nueva subsección"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M12 8v8M8 12h8"/>
            </svg>
            Nueva subsección
          </button>
        )}
      </div>

      {current && (
        <div className="filter-bar">
          <select
            className="filter-select"
            value={filterEstado}
            onChange={e => setFilterEstado(e.target.value)}
          >
            <option value="">Todos los estados</option>
            {ESTADO_OPTIONS_CXP.map(o => <option key={o} value={o}>{o}</option>)}
          </select>
          <div className="filter-search-wrap">
            <svg className="filter-search-icon" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
            </svg>
            <input
              type="text"
              className="filter-search"
              placeholder="Buscar Proveedor / Concepto..."
              value={search}
              onChange={e => setSearch(e.target.value)}
            />
            {search && (
              <button className="filter-clear" onClick={() => setSearch('')} title="Limpiar">×</button>
            )}
          </div>
          <div className="filter-search-wrap">
            <svg className="filter-search-icon" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <rect x="3" y="3" width="18" height="18" rx="2"/><line x1="8" y1="8" x2="16" y2="8"/><line x1="8" y1="12" x2="16" y2="12"/><line x1="8" y1="16" x2="12" y2="16"/>
            </svg>
            <input
              type="text"
              className="filter-search filter-search-factura"
              placeholder="Buscar Factura..."
              value={searchFactura}
              onChange={e => setSearchFactura(e.target.value)}
            />
            {searchFactura && (
              <button className="filter-clear" onClick={() => setSearchFactura('')} title="Limpiar">×</button>
            )}
          </div>
          <div className="filter-search-wrap">
            <svg className="filter-search-icon" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>
            </svg>
            <input
              type="date"
              className="filter-search filter-fecha"
              value={filterFecha}
              onChange={e => setFilterFecha(e.target.value)}
              title="Mostrar hasta esta fecha (inclusive)"
            />
            {filterFecha && (
              <button className="filter-clear" onClick={() => setFilterFecha('')} title="Limpiar">×</button>
            )}
          </div>
          {(filterEstado || search || searchFactura || filterFecha) && (
            <span className="filter-count">{filteredRows.length} resultado{filteredRows.length !== 1 ? 's' : ''}</span>
          )}
        </div>
      )}

      {current && (
        <div className="table-scroll-wrapper">
          <table className="viewer-table">
            <thead>
              <tr>
                <th style={{ width: '100px' }}>Estado</th>
                <th style={{ width: '50px' }} className="col-center">Sem.</th>
                <th style={{ width: '130px' }}>Factura</th>
                <th>Proveedor / Concepto</th>
                <th className="col-right" style={{ width: '160px' }}>Valor (COP)</th>
                <th style={{ width: '110px' }}>Vencimiento</th>
                <th>Observaciones</th>
                <th style={{ width: '32px' }}></th>
              </tr>
            </thead>
            <tbody>
              {filteredRows.length === 0 && subInsertions.length === 0 ? (
                <tr>
                  <td colSpan={8} className="td-empty">Sin resultados para este filtro</td>
                </tr>
              ) : (() => {
                const allRows = [
                  ...filteredRows.map(r => ({ type: 'existing', data: r, fecha: fechaAComparar(edits[`${r._row},5`] ?? r.fechaVencimiento) })),
                  ...subInsertions.map(ins => ({ type: 'new', data: ins, fecha: fechaAComparar(ins.cells[5]) }))
                ].sort((a, b) => a.fecha - b.fecha)

                return allRows.map((item, i) => {
                  if (item.type === 'new') {
                    const ins = item.data
                    const valorRaw = ins.cells[4]
                    const fechaRaw = ins.cells[5]
                    const valorFormateado = typeof valorRaw === 'number' && !isNaN(valorRaw)
                      ? new Intl.NumberFormat('es-CO', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(valorRaw)
                      : valorRaw || ''
                    const fechaFormateada = typeof fechaRaw === 'number' && fechaRaw > 40000 && fechaRaw < 60000
                      ? (() => {
                          const date = new Date((fechaRaw - 25569) * 86400 * 1000)
                          const d = String(date.getUTCDate()).padStart(2, '0')
                          const m = String(date.getUTCMonth() + 1).padStart(2, '0')
                          const y = date.getUTCFullYear()
                          return `${d}/${m}/${y}`
                        })()
                      : fechaRaw || ''

                    return (
                    <tr key={ins.id} className="tr-new">
                      <td className="td-new-cell">
                        <select className="new-row-select" value={ins.cells[6] ?? ''}
                          onChange={e => onInsertedRowEdit(ins.id, 6, e.target.value)}>
                          <option value="">— Estado —</option>
                          {ESTADO_OPTIONS_CXP.map(o => <option key={o}>{o}</option>)}
                        </select>
                      </td>
                      <td className="td-new-cell col-center">
                        <input className="new-row-input" type="text" placeholder="Sem."
                          value={ins.cells[1] ?? ''} onChange={e => onInsertedRowEdit(ins.id, 1, e.target.value)} />
                      </td>
                      <td className="td-new-cell">
                        <input className="new-row-input" type="text" placeholder="Factura"
                          value={ins.cells[2] ?? ''} onChange={e => onInsertedRowEdit(ins.id, 2, e.target.value)} autoFocus />
                      </td>
                      <td className="td-new-cell">
                        <input className="new-row-input" type="text" placeholder="Proveedor / Concepto"
                          value={ins.cells[3] ?? ''} onChange={e => onInsertedRowEdit(ins.id, 3, e.target.value)} />
                      </td>
                      <td className="td-new-cell col-right">
                        <input className="new-row-input" type="text" placeholder="0"
                          value={valorFormateado}
                          onChange={e => {
                            const val = e.target.value.replace(/[^\d]/g, '')
                            onInsertedRowEdit(ins.id, 4, val === '' ? '' : Number(val))
                          }} />
                      </td>
                      <td className="td-new-cell">
                        <input className="new-row-input" type="text" placeholder="dd/mm/yyyy"
                          value={fechaFormateada} onChange={e => onInsertedRowEdit(ins.id, 5, e.target.value)} />
                      </td>
                      <td className="td-new-cell">
                        <input className="new-row-input" type="text" placeholder="Observaciones"
                          value={ins.cells[10] ?? ''} onChange={e => onInsertedRowEdit(ins.id, 10, e.target.value)} />
                      </td>
                      <td className="td-action">
                        <button className="btn-del-new" onClick={() => onDeleteInsertedRow(ins.id)} title="Eliminar fila nueva">×</button>
                      </td>
                    </tr>
                    )
                  }

                  const row = item.data
                  const totalAbonado = getTotalAbonadoForKey(abonos || {}, sheetName, row.proveedor, row.factura)
                  const tieneAbonos = totalAbonado > 0
                  return (
                  <tr key={`row-${row._row}`} className={(i % 2 === 0 ? 'tr-even' : '') + (tieneAbonos ? ' tr-abonado' : '')}>
                    <EditableCell rowIdx={row._row} colIdx={6} raw={row.estado}           edits={edits} onCellEdit={onCellEdit} options={ESTADO_OPTIONS_CXP} badgeMap={ESTADO_BADGE_CXP} />
                    <EditableCell rowIdx={row._row} colIdx={1} raw={row.semana}           edits={edits} onCellEdit={onCellEdit} format="text" className="col-center" />
                    <EditableCell rowIdx={row._row} colIdx={2} raw={row.factura}          edits={edits} onCellEdit={onCellEdit} format="text" className="td-factura" />
                    <EditableCell rowIdx={row._row} colIdx={3} raw={row.proveedor}        edits={edits} onCellEdit={onCellEdit} format="text" className="td-desc" />
                    <td className="col-right td-valor">
                      {tieneAbonos ? (
                        <span className={tieneAbonos ? 'valor-efectivo' : ''}>
                          {formatCOP(row.valor)}
                        </span>
                      ) : (
                        <EditableCell rowIdx={row._row} colIdx={4} raw={row.valor} edits={edits} onCellEdit={onCellEdit} format="cop" className="col-right td-valor" />
                      )}
                      {tieneAbonos && (
                        <span className="valor-abonado-info" title={'Total abonado: ' + formatCOP(totalAbonado)}>
                          {' (-' + formatCOP(totalAbonado) + ')'}
                        </span>
                      )}
                    </td>
                    <EditableCell rowIdx={row._row} colIdx={5} raw={row.fechaVencimiento} edits={edits} onCellEdit={onCellEdit} format="text" className="td-fecha" />
                    <EditableCell rowIdx={row._row} colIdx={10} raw={row.observaciones} edits={edits} onCellEdit={onCellEdit} format="text" className="td-observaciones" />
                    <td className="td-action">
                      {onOpenAbono && row.valor > 0 && (
                        <button className="btn-abonar"
                          onClick={() => onOpenAbono(row.proveedor, row.factura, row.valor)}
                          title="Agregar abono">$</button>
                      )}
                    </td>
                  </tr>
                  )
                })
              })()}

              <tr className="tr-add-row">
                <td colSpan={8}>
                  <button
                    className="btn-add-row"
                    disabled={!current}
                    onClick={() => {
                      const lastRow = current.rows.length > 0
                        ? current.rows[current.rows.length - 1]._row
                        : current._row
                      onAddRow(lastRow, `cxp:${current.title}`)
                    }}
                  >
                    + Agregar fila
                  </button>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}

export default SectionCXP
