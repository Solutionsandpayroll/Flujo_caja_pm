import { useState } from 'react'
import { isMonthSheet, isCurrentYearSheet, monthSheetIndex, MONTHS } from '../utils/excelParser'

export default function SheetSelectorModal({ sheetNames, onSelect, fileName }) {
  const [selected, setSelected] = useState(null)

  const monthSheets = sheetNames
    .filter(n => isMonthSheet(n) && isCurrentYearSheet(n))
    .sort((a, b) => monthSheetIndex(a) - monthSheetIndex(b))

  const handleConfirm = () => {
    if (selected) onSelect(selected)
  }

  return (
    <div className="modal-overlay">
      <div className="modal-content modal-sheet-selector">
        <div className="modal-header">
          <h3>Seleccioná la hoja a trabajar</h3>
          <p>Archivo: <strong>{fileName}</strong></p>
          <p className="modal-subtitle">Elegí el mes que querés revisar o editar</p>
        </div>

        <div className="modal-body">
          {monthSheets.length > 0 ? (
            <div className="sheet-grid">
              {monthSheets.map(name => {
                const monthName = MONTHS.find(m => name.toUpperCase().includes(m)) || name
                const isSelected = selected === name
                return (
                  <button
                    key={name}
                    className={`sheet-grid-item ${isSelected ? 'selected' : ''}`}
                    onClick={() => setSelected(name)}
                  >
                    <span className="sheet-grid-month">{monthName}</span>
                    {name !== monthName && (
                      <span className="sheet-grid-full">{name}</span>
                    )}
                  </button>
                )
              })}
            </div>
          ) : (
            <div className="sheet-grid-empty">
              <p>No se encontraron hojas de mes del año actual en este archivo.</p>
              <p className="sheet-grid-empty-hint">Verificá que el Excel contenga hojas con nombres de mes (ej: "ABRIL 2026")</p>
            </div>
          )}
        </div>

        <div className="modal-footer">
          <button
            className="btn-toolbar btn-save"
            onClick={handleConfirm}
            disabled={!selected}
          >
            {selected ? `Continuar con "${selected}"` : 'Seleccioná una hoja'}
          </button>
        </div>
      </div>
    </div>
  )
}
