import { useState } from 'react'
import { formatCOP } from '../utils/excelParser'
import EditableCell from './EditableCell'

const TABS = [
  { key: 'bancos',   label: 'Bancos' },
  { key: 'clientes', label: 'Clientes' }
]

function SectionInitial({ data, sheetName, edits, onCellEdit, insertions, onAddRow, onInsertedRowEdit, onDeleteInsertedRow }) {
  const [activeTab, setActiveTab] = useState('bancos')
  const [search, setSearch]       = useState('')

  const current = data[activeTab]

  const switchTab = (key) => {
    setActiveTab(key)
    setSearch('')
  }

  // Buscar TOTAL INGRESOS en cualquier sección
  const totalIngresos = (() => {
    for (const key of ['bancos', 'clientes']) {
      const row = data[key]?.rows?.find(r => String(r.descripcion || '').trim().toUpperCase() === 'TOTAL INGRESOS')
      if (row?.valor != null) return row.valor
    }
    return null
  })()

  const displayRows = current.rows.filter(r => String(r.descripcion || '').trim().toUpperCase() !== 'TOTAL INGRESOS')

  const filteredRows = displayRows.filter(row => {
    const desc = String(row.descripcion || '').toLowerCase()
    if (search && !desc.includes(search.toLowerCase())) return false
    return true
  })

  const tabInsertions = (insertions || []).filter(ins => ins.sectionKey === `initial:${activeTab}`)

  return (
    <div className="viewer-section">
      <div className="viewer-section-header">
        <div>
          <h3 className="viewer-section-title">Sección Inicial</h3>
          <p className="viewer-section-sub">Ingresos y saldos · {sheetName}</p>
        </div>
        {totalIngresos != null && (
          <div className="section-total-badge">
            <span className="section-total-label">Total Ingresos</span>
            <span className="section-total-value">{formatCOP(totalIngresos)}</span>
          </div>
        )}
      </div>

      <div className="subsection-tabs">
        {TABS.map(tab => {
          const group = data[tab.key]
          return (
            <button
              key={tab.key}
              className={`subsection-tab ${activeTab === tab.key ? 'active' : ''}`}
              onClick={() => switchTab(tab.key)}
            >
              <span className="tab-label">{tab.label}</span>
              {group.total !== null && (
                <span className="tab-total">{formatCOP(group.total)}</span>
              )}
            </button>
          )
        })}
      </div>

      <div className="filter-bar">
        <div className="filter-search-wrap">
          <svg className="filter-search-icon" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
          </svg>
          <input
            type="text"
            className="filter-search"
            placeholder="Buscar en Descripción…"
            value={search}
            onChange={e => setSearch(e.target.value)}
          />
          {search && (
            <button className="filter-clear" onClick={() => setSearch('')} title="Limpiar">×</button>
          )}
        </div>
        {search && (
          <span className="filter-count">{filteredRows.length} resultado{filteredRows.length !== 1 ? 's' : ''}</span>
        )}
      </div>

      <div className="table-scroll-wrapper">
        <table className="viewer-table">
          <thead>
            <tr>
              <th>Descripción</th>
              <th className="col-right" style={{ width: '180px' }}>Valor (COP)</th>
              <th style={{ width: '32px' }}></th>
            </tr>
          </thead>
          <tbody>
            {filteredRows.length === 0 && tabInsertions.length === 0 ? (
              <tr><td colSpan={3} className="td-empty">Sin resultados</td></tr>
            ) : (
              filteredRows.map((row, i) => (
                <tr key={i} className={i % 2 === 0 ? 'tr-even' : ''}>
                  <EditableCell rowIdx={row._row} colIdx={3} raw={row.descripcion} edits={edits} onCellEdit={onCellEdit} format="text" className="td-desc" />
                  <EditableCell rowIdx={row._row} colIdx={4} raw={row.valor}  edits={edits} onCellEdit={onCellEdit} format="cop"  className="col-right td-valor" />
                  <td className="td-action"></td>
                </tr>
              ))
            )}

            {tabInsertions.map(ins => (
              <tr key={ins.id} className="tr-new">
                <td className="td-new-cell">
                  <input className="new-row-input" type="text" placeholder="Descripción" value={ins.cells[3] ?? ''} onChange={e => onInsertedRowEdit(ins.id, 3, e.target.value)} autoFocus />
                </td>
                <td className="td-new-cell col-right">
                  <input className="new-row-input" type="number" placeholder="0" value={ins.cells[4] ?? ''} onChange={e => onInsertedRowEdit(ins.id, 4, e.target.value === '' ? '' : Number(e.target.value))} />
                </td>
                <td className="td-action">
                  <button className="btn-del-new" onClick={() => onDeleteInsertedRow(ins.id)} title="Eliminar">×</button>
                </td>
              </tr>
            ))}

            <tr className="tr-add-row">
              <td colSpan={3}>
                <button className="btn-add-row" disabled={current.rows.length === 0} onClick={() => { const lastRow = current.rows[current.rows.length - 1]._row; onAddRow(lastRow, `initial:${activeTab}`) }}>
                  + Agregar fila
                </button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  )
}

export default SectionInitial
