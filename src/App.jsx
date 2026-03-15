import { useState, useRef, useCallback, useMemo, useEffect } from 'react'
import './App.css'
import { createEngine } from './engine/core.js'

const TOTAL_ROWS = 50
const TOTAL_COLS = 50
const STORAGE_KEY = 'spreadsheet_app_data'
const SAVE_DEBOUNCE_MS = 500

// ── LocalStorage helpers ──

function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return null
    const data = JSON.parse(raw)
    if (!data || typeof data !== 'object') return null
    return data
  } catch {
    // Corrupted data — remove it
    try { localStorage.removeItem(STORAGE_KEY) } catch { /* ignore */ }
    return null
  }
}

function saveToStorage(engineData, cellStyles) {
  try {
    const payload = { engine: engineData, styles: cellStyles }
    const json = JSON.stringify(payload)
    localStorage.setItem(STORAGE_KEY, json)
    return true
  } catch (e) {
    // Storage full or unavailable
    if (e.name === 'QuotaExceededError' || e.code === 22) {
      console.warn('LocalStorage quota exceeded. Unable to save.')
    }
    return false
  }
}

export default function App() {
  // Engine instance is created once and reused across renders
  const [engine] = useState(() => {
    const eng = createEngine(TOTAL_ROWS, TOTAL_COLS)
    // Attempt to restore from localStorage
    const stored = loadFromStorage()
    if (stored && stored.engine) {
      eng.deserialize(stored.engine)
    }
    return eng
  })

  const [version, setVersion] = useState(0)
  const [selectedCell, setSelectedCell] = useState(null)
  const [editingCell, setEditingCell] = useState(null)
  const [editValue, setEditValue] = useState('')

  // Cell styles persisted separately
  const [cellStyles, setCellStyles] = useState(() => {
    const stored = loadFromStorage()
    return (stored && stored.styles && typeof stored.styles === 'object') ? stored.styles : {}
  })

  // Multi-cell selection: { startR, startC, endR, endC } or null
  const [selectionRange, setSelectionRange] = useState(null)
  const isDragging = useRef(false)

  // Sort/filter UI state
  const [filterOpenCol, setFilterOpenCol] = useState(null)
  const [filterChecked, setFilterChecked] = useState({})

  const cellInputRef = useRef(null)
  const saveTimerRef = useRef(null)
  const gridRef = useRef(null)

  const forceRerender = useCallback(() => setVersion(v => v + 1), [])

  // ── Debounced auto-save ──

  const scheduleSave = useCallback(() => {
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current)
    saveTimerRef.current = setTimeout(() => {
      const engineData = engine.serialize()
      saveToStorage(engineData, cellStyles)
    }, SAVE_DEBOUNCE_MS)
  }, [engine, cellStyles])

  // Save whenever data changes (version tracks engine mutations, cellStyles tracks style changes)
  useEffect(() => {
    scheduleSave()
    return () => { if (saveTimerRef.current) clearTimeout(saveTimerRef.current) }
  }, [version, cellStyles, scheduleSave])

  // ── Cell style helpers ──

  const getCellStyle = useCallback((row, col) => {
    const key = `${row},${col}`
    return cellStyles[key] || {
      bold: false, italic: false, underline: false,
      bg: 'white', color: '#202124', align: 'left', fontSize: 13
    }
  }, [cellStyles])

  const updateCellStyle = useCallback((row, col, updates) => {
    const key = `${row},${col}`
    setCellStyles(prev => ({
      ...prev,
      [key]: { ...getCellStyle(row, col), ...updates }
    }))
  }, [getCellStyle])

  // ── Cell editing ──

  const startEditing = useCallback((row, col) => {
    setSelectedCell({ r: row, c: col })
    setEditingCell({ r: row, c: col })
    const cellData = engine.getCell(row, col)
    setEditValue(cellData.raw)
    setTimeout(() => cellInputRef.current?.focus(), 0)
  }, [engine])

  const commitEdit = useCallback((row, col) => {
    const currentCell = engine.getCell(row, col)
    if (currentCell.raw !== editValue) {
      engine.setCell(row, col, editValue)
      forceRerender()
    }
    setEditingCell(null)
  }, [engine, editValue, forceRerender])

  const handleCellClick = useCallback((row, col) => {
    if (editingCell && (editingCell.r !== row || editingCell.c !== col)) {
      commitEdit(editingCell.r, editingCell.c)
    }
    if (!editingCell || editingCell.r !== row || editingCell.c !== col) {
      startEditing(row, col)
    }
  }, [editingCell, commitEdit, startEditing])

  // ── Multi-cell selection ──

  const normalizedSelection = useMemo(() => {
    if (!selectionRange) return null
    return {
      r1: Math.min(selectionRange.startR, selectionRange.endR),
      c1: Math.min(selectionRange.startC, selectionRange.endC),
      r2: Math.max(selectionRange.startR, selectionRange.endR),
      c2: Math.max(selectionRange.startC, selectionRange.endC),
    }
  }, [selectionRange])

  const isCellInSelection = useCallback((row, col) => {
    if (!normalizedSelection) return false
    return row >= normalizedSelection.r1 && row <= normalizedSelection.r2 &&
      col >= normalizedSelection.c1 && col <= normalizedSelection.c2
  }, [normalizedSelection])

  const handleCellMouseDown = useCallback((e, row, col) => {
    e.preventDefault()
    if (editingCell && (editingCell.r !== row || editingCell.c !== col)) {
      commitEdit(editingCell.r, editingCell.c)
    }

    if (e.shiftKey && selectedCell) {
      // Extend selection
      setSelectionRange({
        startR: selectedCell.r, startC: selectedCell.c,
        endR: row, endC: col,
      })
      setEditingCell(null)
    } else {
      setSelectedCell({ r: row, c: col })
      setSelectionRange({ startR: row, startC: col, endR: row, endC: col })
      isDragging.current = true
      // Don't start editing on mousedown — wait for click (mouseup on same cell)
    }
  }, [editingCell, commitEdit, selectedCell])

  const handleCellMouseEnter = useCallback((row, col) => {
    if (!isDragging.current) return
    setSelectionRange(prev => prev ? { ...prev, endR: row, endC: col } : null)
  }, [])

  const handleMouseUp = useCallback(() => {
    isDragging.current = false
  }, [])

  const handleCellDoubleClick = useCallback((row, col) => {
    startEditing(row, col)
  }, [startEditing])

  useEffect(() => {
    window.addEventListener('mouseup', handleMouseUp)
    return () => window.removeEventListener('mouseup', handleMouseUp)
  }, [handleMouseUp])

  // ── Clipboard: Copy & Paste ──

  const handleCopy = useCallback((e) => {
    // Get the effective selection range
    const sel = normalizedSelection
    if (!sel && !selectedCell) return

    const r1 = sel ? sel.r1 : selectedCell.r
    const c1 = sel ? sel.c1 : selectedCell.c
    const r2 = sel ? sel.r2 : selectedCell.r
    const c2 = sel ? sel.c2 : selectedCell.c

    // Build tab-separated text from computed values
    const lines = []
    for (let r = r1; r <= r2; r++) {
      const cols = []
      for (let c = c1; c <= c2; c++) {
        const cellData = engine.getCell(r, c)
        const display = cellData.error
          ? cellData.error
          : (cellData.computed !== null && cellData.computed !== '' ? String(cellData.computed) : cellData.raw)
        cols.push(display)
      }
      lines.push(cols.join('\t'))
    }
    const text = lines.join('\n')

    if (e.clipboardData) {
      e.clipboardData.setData('text/plain', text)
      e.preventDefault()
    } else {
      navigator.clipboard.writeText(text).catch(() => {})
    }
  }, [normalizedSelection, selectedCell, engine])

  const handlePaste = useCallback((e) => {
    const target = selectedCell
    if (!target) return

    let text = ''
    if (e.clipboardData) {
      text = e.clipboardData.getData('text/plain')
      e.preventDefault()
    }

    if (!text) return

    // If currently editing, cancel the edit
    if (editingCell) {
      setEditingCell(null)
    }

    // Parse tab-separated data (Excel/Google Sheets format)
    const rows = text.split(/\r?\n/).filter((line, i, arr) => {
      // Remove trailing empty line (common in clipboard data)
      if (i === arr.length - 1 && line === '') return false
      return true
    })
    const grid = rows.map(row => row.split('\t'))

    // Build batch update
    const updates = []
    for (let r = 0; r < grid.length; r++) {
      for (let c = 0; c < grid[r].length; c++) {
        const targetRow = target.r + r
        const targetCol = target.c + c
        if (targetRow < engine.rows && targetCol < engine.cols) {
          updates.push({ row: targetRow, col: targetCol, value: grid[r][c] })
        }
      }
    }

    if (updates.length > 0) {
      engine.batchSetCells(updates)
      forceRerender()
      // Select the pasted range
      setSelectionRange({
        startR: target.r, startC: target.c,
        endR: Math.min(target.r + grid.length - 1, engine.rows - 1),
        endC: Math.min(target.c + (grid[0]?.length || 1) - 1, engine.cols - 1),
      })
    }
  }, [selectedCell, editingCell, engine, forceRerender])

  // Global clipboard handlers
  useEffect(() => {
    const onCopy = (e) => {
      // Only handle if no text input is focused (except our cell input)
      const active = document.activeElement
      const isFormInput = active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA')
      const isCellInput = active?.classList?.contains('cell-input')
      const isFormulaBar = active?.classList?.contains('formula-bar-input')
      if (isFormInput && !isCellInput && !isFormulaBar) return

      // If editing a cell, let default copy behavior work on the input
      if (editingCell) return

      handleCopy(e)
    }
    const onPaste = (e) => {
      const active = document.activeElement
      const isFormInput = active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA')
      const isCellInput = active?.classList?.contains('cell-input')
      const isFormulaBar = active?.classList?.contains('formula-bar-input')
      if (isFormInput && !isCellInput && !isFormulaBar) return

      // If editing, let default paste work on the single cell
      if (editingCell) return

      handlePaste(e)
    }

    document.addEventListener('copy', onCopy)
    document.addEventListener('paste', onPaste)
    return () => {
      document.removeEventListener('copy', onCopy)
      document.removeEventListener('paste', onPaste)
    }
  }, [handleCopy, handlePaste, editingCell])

  // ── Keyboard navigation + shortcuts ──

  const handleUndo = useCallback(() => { if (engine.undo()) forceRerender() }, [engine, forceRerender])
  const handleRedo = useCallback(() => { if (engine.redo()) forceRerender() }, [engine, forceRerender])

  const handleKeyDown = useCallback((event, row, col) => {
    if (event.key === 'Enter') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.min(row + 1, engine.rows - 1), col)
    } else if (event.key === 'Tab') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1))
    } else if (event.key === 'Escape') {
      setEditValue(engine.getCell(row, col).raw)
      setEditingCell(null)
    } else if (event.key === 'ArrowDown') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.min(row + 1, engine.rows - 1), col)
    } else if (event.key === 'ArrowUp') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.max(row - 1, 0), col)
    } else if (event.key === 'ArrowLeft') {
      event.preventDefault()
      commitEdit(row, col)
      if (col > 0) {
        startEditing(row, col - 1)
      } else if (row > 0) {
        startEditing(row - 1, engine.cols - 1)
      }
    } else if (event.key === 'ArrowRight') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1))
    }
  }, [engine, commitEdit, startEditing])

  // Global keyboard shortcuts (Ctrl+Z, Ctrl+Y, Delete)
  useEffect(() => {
    const onKeyDown = (e) => {
      // Don't intercept if a non-cell input/textarea is focused
      const active = document.activeElement
      const isFormInput = active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA')
      const isCellInput = active?.classList?.contains('cell-input')
      const isFormulaBar = active?.classList?.contains('formula-bar-input')

      if (e.ctrlKey || e.metaKey) {
        if (e.key === 'z' || e.key === 'Z') {
          if (editingCell) return // Let browser handle undo in text input
          if (isFormInput && !isCellInput && !isFormulaBar) return
          e.preventDefault()
          handleUndo()
        } else if (e.key === 'y' || e.key === 'Y') {
          if (editingCell) return
          if (isFormInput && !isCellInput && !isFormulaBar) return
          e.preventDefault()
          handleRedo()
        }
        return
      }

      // Delete/Backspace clears selected cells
      if ((e.key === 'Delete' || e.key === 'Backspace') && !editingCell) {
        if (isFormInput && !isCellInput && !isFormulaBar) return
        const sel = normalizedSelection
        if (sel) {
          e.preventDefault()
          const updates = []
          for (let r = sel.r1; r <= sel.r2; r++) {
            for (let c = sel.c1; c <= sel.c2; c++) {
              updates.push({ row: r, col: c, value: '' })
            }
          }
          engine.batchSetCells(updates)
          forceRerender()
        } else if (selectedCell) {
          e.preventDefault()
          engine.setCell(selectedCell.r, selectedCell.c, '')
          forceRerender()
        }
      }
    }

    document.addEventListener('keydown', onKeyDown)
    return () => document.removeEventListener('keydown', onKeyDown)
  }, [editingCell, handleUndo, handleRedo, normalizedSelection, selectedCell, engine, forceRerender])

  // ── Formula bar handlers ──

  const handleFormulaBarKeyDown = useCallback((event) => {
    if (!editingCell) return
    handleKeyDown(event, editingCell.r, editingCell.c)
  }, [editingCell, handleKeyDown])

  const handleFormulaBarFocus = useCallback(() => {
    if (selectedCell && !editingCell) {
      setEditingCell(selectedCell)
      setEditValue(engine.getCell(selectedCell.r, selectedCell.c).raw)
    }
  }, [selectedCell, editingCell, engine])

  const handleFormulaBarChange = useCallback((value) => {
    if (!editingCell && selectedCell) setEditingCell(selectedCell)
    setEditValue(value)
  }, [editingCell, selectedCell])

  // ── Formatting toggles ──

  const toggleBold = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { bold: !style.bold })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleItalic = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { italic: !style.italic })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleUnderline = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { underline: !style.underline })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const changeFontSize = useCallback((size) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { fontSize: size })
  }, [selectedCell, updateCellStyle])

  const changeAlignment = useCallback((align) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { align })
  }, [selectedCell, updateCellStyle])

  const changeFontColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { color })
  }, [selectedCell, updateCellStyle])

  const changeBackgroundColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { bg: color })
  }, [selectedCell, updateCellStyle])

  // ── Clear operations ──

  const clearSelectedCell = useCallback(() => {
    if (!selectedCell) return
    engine.setCell(selectedCell.r, selectedCell.c, '')
    forceRerender()
    const key = `${selectedCell.r},${selectedCell.c}`
    setCellStyles(prev => { const next = { ...prev }; delete next[key]; return next })
    setEditValue('')
  }, [selectedCell, engine, forceRerender])

  const clearAllCells = useCallback(() => {
    for (let r = 0; r < engine.rows; r++) {
      for (let c = 0; c < engine.cols; c++) {
        engine.setCell(r, c, '')
      }
    }
    forceRerender()
    setCellStyles({})
    setSelectedCell(null)
    setEditingCell(null)
    setEditValue('')
  }, [engine, forceRerender])

  // ── Row / Column operations ──

  const insertRow = useCallback(() => {
    if (!selectedCell) return
    engine.insertRow(selectedCell.r)
    forceRerender()
    setSelectedCell({ r: selectedCell.r + 1, c: selectedCell.c })
  }, [selectedCell, engine, forceRerender])

  const deleteRow = useCallback(() => {
    if (!selectedCell) return
    engine.deleteRow(selectedCell.r)
    forceRerender()
    if (selectedCell.r >= engine.rows) {
      setSelectedCell({ r: engine.rows - 1, c: selectedCell.c })
    }
  }, [selectedCell, engine, forceRerender])

  const insertColumn = useCallback(() => {
    if (!selectedCell) return
    engine.insertColumn(selectedCell.c)
    forceRerender()
    setSelectedCell({ r: selectedCell.r, c: selectedCell.c + 1 })
  }, [selectedCell, engine, forceRerender])

  const deleteColumn = useCallback(() => {
    if (!selectedCell) return
    engine.deleteColumn(selectedCell.c)
    forceRerender()
    if (selectedCell.c >= engine.cols) {
      setSelectedCell({ r: selectedCell.r, c: engine.cols - 1 })
    }
  }, [selectedCell, engine, forceRerender])

  // ── Sort & Filter ──

  const sortState = engine.getSortState()

  const handleSort = useCallback((col) => {
    const current = engine.getSortState()
    if (current && current.col === col) {
      if (current.direction === 'asc') {
        engine.setSortState(col, 'desc')
      } else {
        engine.setSortState(col, null)
      }
    } else {
      engine.setSortState(col, 'asc')
    }
    forceRerender()
  }, [engine, forceRerender])

  const openFilterDropdown = useCallback((col) => {
    if (filterOpenCol === col) {
      setFilterOpenCol(null)
      return
    }
    const uniqueValues = engine.getColumnUniqueValues(col)
    const currentFilter = engine.getFilterState()
    const currentAllowed = currentFilter[col]
    const checked = {}
    for (const val of uniqueValues) {
      checked[val] = currentAllowed ? currentAllowed.includes(val) : true
    }
    setFilterChecked(checked)
    setFilterOpenCol(col)
  }, [filterOpenCol, engine])

  const applyFilter = useCallback(() => {
    if (filterOpenCol === null) return
    const allowed = new Set()
    let allChecked = true
    for (const [val, isChecked] of Object.entries(filterChecked)) {
      if (isChecked) {
        allowed.add(val)
      } else {
        allChecked = false
      }
    }
    if (allChecked) {
      engine.clearColumnFilter(filterOpenCol)
    } else {
      engine.setColumnFilter(filterOpenCol, allowed)
    }
    setFilterOpenCol(null)
    forceRerender()
  }, [filterOpenCol, filterChecked, engine, forceRerender])

  const clearFilter = useCallback(() => {
    if (filterOpenCol === null) return
    engine.clearColumnFilter(filterOpenCol)
    setFilterOpenCol(null)
    forceRerender()
  }, [filterOpenCol, engine, forceRerender])

  // ── Derived state ──

  const viewRowOrder = useMemo(() => engine.getViewRowOrder(), [engine, version])

  const selectedCellStyle = useMemo(() => {
    return selectedCell ? getCellStyle(selectedCell.r, selectedCell.c) : null
  }, [selectedCell, getCellStyle])

  const getColumnLabel = useCallback((col) => {
    let label = ''
    let num = col + 1
    while (num > 0) {
      num--
      label = String.fromCharCode(65 + (num % 26)) + label
      num = Math.floor(num / 26)
    }
    return label
  }, [])

  const selectedCellLabel = selectedCell
    ? `${getColumnLabel(selectedCell.c)}${selectedCell.r + 1}`
    : 'No cell'

  const formulaBarValue = editingCell
    ? editValue
    : (selectedCell ? engine.getCell(selectedCell.r, selectedCell.c).raw : '')

  const activeFilters = engine.getFilterState()
  const hasAnyFilter = Object.keys(activeFilters).length > 0

  // ── Render ──

  return (
    <div className="app-wrapper" tabIndex={-1}>
      <div className="app-header">
        <h2 className="app-title">Spreadsheet App</h2>
      </div>

      <div className="main-content">

        {/* Toolbar */}
        <div className="toolbar">
          <div className="toolbar-group">
            <button className={`toolbar-btn bold-btn ${selectedCellStyle?.bold ? 'active' : ''}`} onClick={toggleBold} title="Bold">B</button>
            <button className={`toolbar-btn italic-btn ${selectedCellStyle?.italic ? 'active' : ''}`} onClick={toggleItalic} title="Italic">I</button>
            <button className={`toolbar-btn underline-btn ${selectedCellStyle?.underline ? 'active' : ''}`} onClick={toggleUnderline} title="Underline">U</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Size:</span>
            <select className="toolbar-select" value={selectedCellStyle?.fontSize || 13} onChange={(e) => changeFontSize(parseInt(e.target.value))}>
              {[8, 10, 11, 12, 13, 14, 16, 18, 20, 24].map(s => (
                <option key={s} value={s}>{s}</option>
              ))}
            </select>
          </div>

          <div className="toolbar-group">
            <button className={`align-btn ${selectedCellStyle?.align === 'left' ? 'active' : ''}`} onClick={() => changeAlignment('left')} title="Align Left">&#x2190;</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'center' ? 'active' : ''}`} onClick={() => changeAlignment('center')} title="Align Center">&#x2194;</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'right' ? 'active' : ''}`} onClick={() => changeAlignment('right')} title="Align Right">&#x2192;</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Text:</span>
            <input
              type="color"
              value={selectedCellStyle?.color || '#000000'}
              onChange={(e) => changeFontColor(e.target.value)}
              title="Font color"
              style={{ width: '32px', height: '32px', border: '1px solid #dadce0', cursor: 'pointer', borderRadius: '4px' }}
            />
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Fill:</span>
            <select className="toolbar-select" value={selectedCellStyle?.bg || 'white'} onChange={(e) => changeBackgroundColor(e.target.value)}>
              <option value="white">White</option>
              <option value="#ffff99">Yellow</option>
              <option value="#99ffcc">Green</option>
              <option value="#ffcccc">Red</option>
              <option value="#cce5ff">Blue</option>
              <option value="#e0ccff">Purple</option>
              <option value="#ffd9b3">Orange</option>
              <option value="#f0f0f0">Gray</option>
            </select>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={handleUndo} disabled={!engine.canUndo()} title="Undo (Ctrl+Z)">&#x21B6; Undo</button>
            <button className="toolbar-btn" onClick={handleRedo} disabled={!engine.canRedo()} title="Redo (Ctrl+Y)">&#x21B7; Redo</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={insertRow} title="Insert Row">+ Row</button>
            <button className="toolbar-btn" onClick={deleteRow} title="Delete Row">- Row</button>
            <button className="toolbar-btn" onClick={insertColumn} title="Insert Column">+ Col</button>
            <button className="toolbar-btn" onClick={deleteColumn} title="Delete Column">- Col</button>
          </div>

          <div className="toolbar-group">
            {hasAnyFilter && (
              <button className="toolbar-btn" onClick={() => { engine.clearAllFilters(); forceRerender() }} title="Clear all filters">Clear Filters</button>
            )}
            <button className="toolbar-btn danger" onClick={clearSelectedCell}>&#x2715; Cell</button>
            <button className="toolbar-btn danger" onClick={clearAllCells}>&#x2715; All</button>
          </div>
        </div>

        {/* Formula Bar */}
        <div className="formula-bar">
          <span className="formula-bar-label">{selectedCellLabel}</span>
          <input
            className="formula-bar-input"
            value={formulaBarValue}
            onChange={(e) => handleFormulaBarChange(e.target.value)}
            onKeyDown={handleFormulaBarKeyDown}
            onFocus={handleFormulaBarFocus}
            placeholder="Select a cell then type, or enter a formula like =SUM(A1:A5)"
          />
        </div>

        {/* Grid */}
        <div className="grid-scroll" ref={gridRef}>
          <table className="grid-table">
            <thead>
              <tr>
                <th className="col-header-blank"></th>
                {Array.from({ length: engine.cols }, (_, colIndex) => {
                  const colSorted = sortState && sortState.col === colIndex
                  const colFiltered = activeFilters[colIndex] != null
                  return (
                    <th key={colIndex} className={`col-header ${colFiltered ? 'filtered' : ''}`}>
                      <div className="col-header-content">
                        <span
                          className="col-header-label"
                          onClick={() => handleSort(colIndex)}
                          title="Click to sort"
                        >
                          {getColumnLabel(colIndex)}
                          {colSorted && (
                            <span className="sort-indicator">
                              {sortState.direction === 'asc' ? ' \u25B2' : ' \u25BC'}
                            </span>
                          )}
                        </span>
                        <button
                          className={`filter-btn ${colFiltered ? 'active' : ''}`}
                          onClick={(e) => { e.stopPropagation(); openFilterDropdown(colIndex) }}
                          title="Filter"
                        >
                          &#x25BC;
                        </button>
                      </div>
                      {filterOpenCol === colIndex && (
                        <div className="filter-dropdown" onClick={(e) => e.stopPropagation()}>
                          <div className="filter-dropdown-header">Filter: {getColumnLabel(colIndex)}</div>
                          <div className="filter-dropdown-actions">
                            <button
                              className="filter-action-btn"
                              onClick={() => {
                                const next = {}
                                for (const key of Object.keys(filterChecked)) next[key] = true
                                setFilterChecked(next)
                              }}
                            >Select All</button>
                            <button
                              className="filter-action-btn"
                              onClick={() => {
                                const next = {}
                                for (const key of Object.keys(filterChecked)) next[key] = false
                                setFilterChecked(next)
                              }}
                            >Clear All</button>
                          </div>
                          <div className="filter-dropdown-list">
                            {Object.entries(filterChecked).map(([val, checked]) => (
                              <label key={val} className="filter-item">
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  onChange={() => setFilterChecked(prev => ({ ...prev, [val]: !prev[val] }))}
                                />
                                <span className="filter-item-text">{val === '' ? '(Blank)' : val}</span>
                              </label>
                            ))}
                          </div>
                          <div className="filter-dropdown-footer">
                            <button className="filter-apply-btn" onClick={applyFilter}>Apply</button>
                            <button className="filter-clear-btn" onClick={clearFilter}>Clear</button>
                            <button className="filter-cancel-btn" onClick={() => setFilterOpenCol(null)}>Cancel</button>
                          </div>
                        </div>
                      )}
                    </th>
                  )
                })}
              </tr>
            </thead>
            <tbody>
              {viewRowOrder.map((rowIndex) => (
                <tr key={rowIndex}>
                  <td className="row-header">{rowIndex + 1}</td>
                  {Array.from({ length: engine.cols }, (_, colIndex) => {
                    const isSelected = selectedCell?.r === rowIndex && selectedCell?.c === colIndex
                    const isEditing = editingCell?.r === rowIndex && editingCell?.c === colIndex
                    const inRange = isCellInSelection(rowIndex, colIndex)
                    const cellData = engine.getCell(rowIndex, colIndex)
                    const style = cellStyles[`${rowIndex},${colIndex}`] || {}
                    const displayValue = cellData.error
                      ? cellData.error
                      : (cellData.computed !== null && cellData.computed !== '' ? String(cellData.computed) : cellData.raw)

                    return (
                      <td
                        key={colIndex}
                        className={`cell ${isSelected ? 'selected' : ''} ${inRange && !isSelected ? 'in-range' : ''}`}
                        style={{ background: style.bg || 'white' }}
                        onMouseDown={(e) => handleCellMouseDown(e, rowIndex, colIndex)}
                        onMouseEnter={() => handleCellMouseEnter(rowIndex, colIndex)}
                        onDoubleClick={() => handleCellDoubleClick(rowIndex, colIndex)}
                      >
                        {isEditing ? (
                          <input
                            autoFocus
                            className="cell-input"
                            value={editValue}
                            onChange={(e) => setEditValue(e.target.value)}
                            onBlur={() => commitEdit(rowIndex, colIndex)}
                            onKeyDown={(e) => handleKeyDown(e, rowIndex, colIndex)}
                            ref={isSelected ? cellInputRef : undefined}
                            style={{
                              fontWeight: style.bold ? 'bold' : 'normal',
                              fontStyle: style.italic ? 'italic' : 'normal',
                              textDecoration: style.underline ? 'underline' : 'none',
                              color: style.color || '#202124',
                              fontSize: (style.fontSize || 13) + 'px',
                              textAlign: style.align || 'left',
                              background: style.bg || 'white',
                            }}
                          />
                        ) : (
                          <div
                            className={`cell-display align-${style.align || 'left'} ${cellData.error ? 'error' : ''}`}
                            style={{
                              fontWeight: style.bold ? 'bold' : 'normal',
                              fontStyle: style.italic ? 'italic' : 'normal',
                              textDecoration: style.underline ? 'underline' : 'none',
                              color: cellData.error ? '#d93025' : (style.color || '#202124'),
                              fontSize: (style.fontSize || 13) + 'px',
                            }}
                          >
                            {displayValue}
                          </div>
                        )}
                      </td>
                    )
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <p className="footer-hint">
          Click a cell to edit &middot; Shift+Click or drag to select range &middot; Ctrl+C/V to copy/paste &middot; Ctrl+Z/Y to undo/redo &middot; Click column header to sort &middot; Formulas: =SUM(A1:A5)
        </p>
      </div>
    </div>
  )
}
