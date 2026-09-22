import { useMemo } from 'react'
import { parseMonthSheet } from '../utils/excelParser'
import SectionInitial from './SectionInitial'
import SectionCXP from './SectionCXP'

/**
 * MonthViewer
 * Recibe las filas crudas de una hoja de mes y renderiza
 * las dos secciones estructuradas: Inicial y Cuentas x Pagar.
 */
function MonthViewer({ rows, sheetName, edits, onCellEdit, insertions, onAddRow, onInsertedRowEdit, onDeleteInsertedRow, abonos, onOpenAbono, onAddSubsection }) {
  const parsed = useMemo(() => parseMonthSheet(rows), [rows])

  if (!parsed) {
    return <div className="empty-sheet">Esta hoja no tiene datos para mostrar.</div>
  }

  const { sectionInitial, sectionCXP } = parsed

  const hasInitial =
    sectionInitial &&
    (sectionInitial.bancos.rows.length > 0 ||
      sectionInitial.clientes.rows.length > 0 ||
      sectionInitial.eor.rows.length > 0)

  const hasCXP = sectionCXP && sectionCXP.length > 0

  if (!hasInitial && !hasCXP) {
    return (
      <div className="empty-sheet">
        Esta hoja no tiene datos estructurados para mostrar.
      </div>
    )
  }

  return (
    <div className="month-viewer">
      {hasCXP && (
        <SectionCXP
          subsections={sectionCXP}
          sheetName={sheetName}
          edits={edits}
          onCellEdit={onCellEdit}
          insertions={insertions}
          onAddRow={onAddRow}
          onInsertedRowEdit={onInsertedRowEdit}
          onDeleteInsertedRow={onDeleteInsertedRow}
          abonos={abonos}
          onOpenAbono={onOpenAbono}
          onAddSubsection={onAddSubsection}
        />
      )}
    </div>
  )
}

export default MonthViewer
