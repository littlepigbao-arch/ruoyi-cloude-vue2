/**
 * Univer 指令执行器
 * 纯逻辑模块（除调用 univer facade API 的副作用外无副作用）
 * 输入 actions + univerAPI → 输出 {results, summary, notes}
 */
import {
  parseA1,
  clampRange,
  validateAction,
  assertSafeFormula,
  MAX_ACTIONS,
  rangeToA1,
  toA1,
  buildCustomFilters,
  parseFormulaToCustomFilters
} from './ai-action-schema'

/**
 * 批量执行
 * @param {object} univerAPI  FUniver facade（仅用于 undo/redo 等命令级操作）
 * @param {object} workbook   FWorkbook 引用（createUnit 返回值，含 sheet/range 操作）
 * @param {Array} actions
 * @param {object} options { sheetName?, onProgress(i,total,action), continueOnError=true }
 * @returns {{results:Array, summary:{ok,failed,skipped}, notes:Array}}
 */
export function executeActions(univerAPI, workbook, actions, options = {}) {
  const results = []
  const notes = []
  let ok = 0
  let failed = 0
  let skipped = 0
  if (!Array.isArray(actions) || actions.length === 0) {
    return { results, summary: { ok, failed, skipped }, notes }
  }
  const total = Math.min(actions.length, MAX_ACTIONS)
  for (let i = 0; i < total; i++) {
    const action = actions[i]
    if (options.onProgress) {
      try {
        options.onProgress(i, total, action)
      } catch (e) {
        // ignore progress error
      }
    }
    const r = executeAction(univerAPI, workbook, action, options)
    results.push(r)
    if (r.ok) ok++
    else failed++
    if (r.note) notes.push(r.note)
    if (!r.ok && options.continueOnError === false) break
  }
  if (actions.length > MAX_ACTIONS) {
    notes.push('actions 数量超过上限 ' + MAX_ACTIONS + '，已截断')
  }
  return { results, summary: { ok, failed, skipped }, notes }
}

/**
 * 执行单个 action
 * @returns {{type,range,ok,error,result,note,durationMs}}
 */
export function executeAction(univerAPI, workbook, action, options = {}) {
  const t0 = typeof performance !== 'undefined' ? performance.now() : Date.now()
  const type = action && action.type
  const range = action && action.range
  // 1. handler 存在性
  const handler = handlers[type]
  if (!handler) {
    return fail(type, range, '未知指令类型: ' + type, t0)
  }
  // 2. schema 校验
  const validateError = validateAction(action)
  if (validateError) {
    return fail(type, range, validateError, t0)
  }
  // 3. 解析 sheet
  let sheet
  try {
    sheet = resolveSheet(workbook, action.sheetName || options.sheetName)
    if (!sheet) throw new Error('无法获取工作表')
  } catch (e) {
    return fail(type, range, e.message || String(e), t0)
  }
  // 4. 执行
  try {
    const ret = handler(univerAPI, sheet, action)
    return {
      type,
      range: range || null,
      ok: true,
      error: null,
      result: (ret && ret.result) || null,
      note: (ret && ret.note) || null,
      durationMs: elapsed(t0)
    }
  } catch (e) {
    return fail(type, range, e.message || String(e), t0)
  }
}

/**
 * 解析 sheet：优先 sheetName，否则活动 sheet，否则第一个
 * @param {object} workbook  FWorkbook 引用（createUnit 返回值）
 */
export function resolveSheet(workbook, sheetName) {
  if (!workbook) return null
  if (sheetName) {
    const byName = safeCall(() => workbook.getSheetByName && workbook.getSheetByName(sheetName))
    if (byName) return byName
  }
  const active = safeCall(() => workbook.getActiveSheet && workbook.getActiveSheet())
  if (active) return active
  const sheets = safeCall(() => workbook.getSheets && workbook.getSheets()) || []
  return sheets[0] || null
}

/**
 * 获取工作表行数（兼容 facade getMaxRows / 核心 getRowCount）
 */
function getRowCount(sheet) {
  if (typeof sheet.getMaxRows === 'function') return sheet.getMaxRows()
  if (typeof sheet.getRowCount === 'function') return sheet.getRowCount()
  return 1000
}

/**
 * 获取工作表列数（兼容 facade getMaxColumns / 核心 getColumnCount）
 */
function getColCount(sheet) {
  if (typeof sheet.getMaxColumns === 'function') return sheet.getMaxColumns()
  if (typeof sheet.getColumnCount === 'function') return sheet.getColumnCount()
  return 26
}

/**
 * 解析 range 字符串为 FRange，越界自动截断并返回 note
 * @returns {{fRange:object, range:object, note:string|null}}
 */
export function resolveRange(sheet, rangeStr) {
  const parsed = parseA1(rangeStr)
  if (!parsed) throw new Error('range 解析失败: ' + rangeStr)
  const rowCount = getRowCount(sheet)
  const colCount = getColCount(sheet)
  if (rowCount <= 0 || colCount <= 0) throw new Error('工作表行列数为 0')
  const { range: clamped, changed, original } = clampRange(parsed, 0, rowCount - 1, 0, colCount - 1)
  let note = null
  if (changed) {
    note = 'range 从 ' + rangeToA1(original) + ' 截断到 ' + rangeToA1(clamped)
  }
  const fRange = sheet.getRange(
    clamped.startRow,
    clamped.startColumn,
    clamped.endRow - clamped.startRow + 1,
    clamped.endColumn - clamped.startColumn + 1
  )
  return { fRange, range: clamped, note }
}

// ============ handlers ============

const handlers = {
  setCell: handleSetCell,
  setCellForCell: handleSetCellForCell,
  setValues: handleSetValues,
  clearRange: handleClearRange,
  setFormula: handleSetFormula,
  merge: handleMerge,
  breakApart: handleBreakApart,
  insertRows: handleInsertRows,
  deleteRows: handleDeleteRows,
  insertColumns: handleInsertColumns,
  deleteColumns: handleDeleteColumns,
  setRowHeight: handleSetRowHeight,
  setColumnWidth: handleSetColumnWidth,
  setStyle: handleSetStyle,
  undo: handleUndo,
  redo: handleRedo,
  getRangeValue: handleGetRangeValue,
  getRangeFormulas: handleGetRangeFormulas,
  getMerges: handleGetMerges,
  getSelection: handleGetSelection,
  sumToCell: handleSumToCell,
  setFilter: handleSetFilter,
  clearFilter: handleClearFilter,
  getFilter: handleGetFilter
}

function handleSetCell(univerAPI, sheet, action) {
  const { fRange, note } = resolveRange(sheet, action.range)
  let value = action.value
  // 公式判断：显式 isFormula 或字符串以 = 开头
  if (action.isFormula === true || (typeof value === 'string' && value.startsWith('='))) {
    let formula = String(value)
    if (!formula.startsWith('=')) formula = '=' + formula
    const r = assertSafeFormula(formula)
    if (r !== true) throw new Error(r)
    value = formula
  }
  fRange.setValue(value)
  return note ? { note } : null
}

function handleSetCellForCell(univerAPI, sheet, action) {
  const { fRange, note } = resolveRange(sheet, action.range)
  let value = action.value
  if (action.isFormula === true || (typeof value === 'string' && value.startsWith('='))) {
    let formula = String(value)
    if (!formula.startsWith('=')) formula = '=' + formula
    const r = assertSafeFormula(formula)
    if (r !== true) throw new Error(r)
    value = formula
  }
  fRange.setValueForCell(value)
  return note ? { note } : null
}

function handleSetValues(univerAPI, sheet, action) {
  const { fRange, range, note } = resolveRange(sheet, action.range)
  const expectedRows = range.endRow - range.startRow + 1
  const expectedCols = range.endColumn - range.startColumn + 1
  if (!Array.isArray(action.values)) throw new Error('values 必须为二维数组')
  if (action.values.length !== expectedRows) {
    throw new Error('values 行数 ' + action.values.length + ' 与区域 ' + expectedRows + ' 不匹配')
  }
  for (let i = 0; i < action.values.length; i++) {
    const row = action.values[i]
    if (!Array.isArray(row)) throw new Error('values[' + i + '] 不是数组')
    if (row.length !== expectedCols) {
      throw new Error('values[' + i + '] 列数 ' + row.length + ' 与区域 ' + expectedCols + ' 不匹配')
    }
  }
  // 公式安全校验（逐格）
  const cleaned = action.values.map((row) =>
    row.map((cell) => {
      if (cell !== null && typeof cell === 'object' && typeof cell.f === 'string') {
        const r = assertSafeFormula(cell.f)
        if (r !== true) throw new Error(r)
      }
      if (typeof cell === 'string' && cell.startsWith('=')) {
        const r = assertSafeFormula(cell)
        if (r !== true) throw new Error(r)
      }
      return cell
    })
  )
  fRange.setValues(cleaned)
  return note ? { note } : null
}

function handleClearRange(univerAPI, sheet, action) {
  const { fRange, note } = resolveRange(sheet, action.range)
  const what = action.clearWhat || 'content'
  // facade 提供 clearContent()/clearFormat()/clear()，比 setValue(null) 更可靠
  // setValue(null) 会被 covertCellValue 校验拒绝并抛 "Invalid value"
  if (what === 'content') {
    if (typeof fRange.clearContent === 'function') {
      fRange.clearContent()
    } else {
      fRange.setValue('')
    }
  } else if (what === 'format') {
    if (typeof fRange.clearFormat === 'function') {
      fRange.clearFormat()
    } else {
      try {
        fRange.setBackgroundColor && fRange.setBackgroundColor('#ffffff')
        fRange.setFontColor && fRange.setFontColor('#000000')
        fRange.setFontSize && fRange.setFontSize(11)
        fRange.setFontWeight && fRange.setFontWeight('normal')
        fRange.setHorizontalAlignment && fRange.setHorizontalAlignment('left')
        fRange.setVerticalAlignment && fRange.setVerticalAlignment('middle')
        fRange.setWrap && fRange.setWrap(false)
      } catch (e) {}
    }
  } else if (what === 'all') {
    if (typeof fRange.clear === 'function') {
      fRange.clear()
    } else {
      if (typeof fRange.clearContent === 'function') fRange.clearContent()
      if (typeof fRange.clearFormat === 'function') fRange.clearFormat()
    }
    if (fRange.isPartOfMerge && fRange.isPartOfMerge()) {
      try { fRange.breakApart() } catch (e) {}
    }
  }
  return note ? { note } : null
}

function handleSetFormula(univerAPI, sheet, action) {
  const { fRange, note } = resolveRange(sheet, action.range)
  let formula = String(action.formula)
  if (!formula.startsWith('=')) formula = '=' + formula
  const r = assertSafeFormula(formula)
  if (r !== true) throw new Error(r)
  fRange.setValue(formula)
  return note ? { note } : null
}

function handleMerge(univerAPI, sheet, action) {
  const { fRange, note } = resolveRange(sheet, action.range)
  const opts = { defaultMerge: true, isForceMerge: action.force === true }
  const mode = action.mode || 'default'
  if (mode === 'across') fRange.mergeAcross(opts)
  else if (mode === 'vertically') fRange.mergeVertically(opts)
  else fRange.merge(opts)
  return note ? { note } : null
}

function handleBreakApart(univerAPI, sheet, action) {
  const { fRange, note } = resolveRange(sheet, action.range)
  fRange.breakApart()
  return note ? { note } : null
}

function handleInsertRows(univerAPI, sheet, action) {
  const count = Math.max(1, Number(action.count) || 1)
  const pos = Number(action.rowIndex)
  const position = action.position || 'at'
  if (position === 'after') sheet.insertRowsAfter(pos, count)
  else if (position === 'before') sheet.insertRowsBefore(pos, count)
  else sheet.insertRows(pos, count)
  return null
}

function handleDeleteRows(univerAPI, sheet, action) {
  const count = Math.max(1, Number(action.count) || 1)
  const pos = Number(action.rowPosition)
  sheet.deleteRows(pos, count)
  return null
}

function handleInsertColumns(univerAPI, sheet, action) {
  const count = Math.max(1, Number(action.count) || 1)
  const pos = Number(action.columnIndex)
  const position = action.position || 'at'
  if (position === 'after') {
    sheet.insertColumnAfter(pos, count)
  } else if (sheet.insertColumns) {
    sheet.insertColumns(pos, count)
  } else {
    sheet.insertColumnAfter(pos, count)
  }
  return null
}

function handleDeleteColumns(univerAPI, sheet, action) {
  const count = Math.max(1, Number(action.count) || 1)
  const pos = Number(action.columnPosition)
  sheet.deleteColumns(pos, count)
  return null
}

function handleSetRowHeight(univerAPI, sheet, action) {
  sheet.setRowHeight(Number(action.rowPosition), Number(action.height))
  return null
}

function handleSetColumnWidth(univerAPI, sheet, action) {
  sheet.setColumnWidth(Number(action.columnPosition), Number(action.width))
  return null
}

function handleSetStyle(univerAPI, sheet, action) {
  const { fRange, note } = resolveRange(sheet, action.range)
  const s = action
  try {
    if (s.background) fRange.setBackgroundColor(s.background)
  } catch (e) {}
  try {
    if (s.fontColor) fRange.setFontColor(s.fontColor)
  } catch (e) {}
  try {
    if (s.fontSize !== undefined) fRange.setFontSize(Number(s.fontSize))
  } catch (e) {}
  try {
    if (s.fontFamily) fRange.setFontFamily(s.fontFamily)
  } catch (e) {}
  try {
    if (s.fontWeight) fRange.setFontWeight(s.fontWeight)
  } catch (e) {}
  try {
    if (s.hAlign) fRange.setHorizontalAlignment(s.hAlign)
  } catch (e) {}
  try {
    if (s.vAlign) fRange.setVerticalAlignment(s.vAlign)
  } catch (e) {}
  try {
    if (typeof s.wrap === 'boolean') fRange.setWrap(s.wrap)
  } catch (e) {}
  try {
    if (s.textRotation !== undefined) fRange.setTextRotation(Number(s.textRotation))
  } catch (e) {}
  return note ? { note } : null
}

function handleUndo(univerAPI, sheet, action) {
  if (typeof univerAPI.undo !== 'function') throw new Error('当前 univer 版本不支持 undo')
  univerAPI.undo()
  return null
}

function handleRedo(univerAPI, sheet, action) {
  if (typeof univerAPI.redo !== 'function') throw new Error('当前 univer 版本不支持 redo')
  univerAPI.redo()
  return null
}

function handleGetRangeValue(univerAPI, sheet, action) {
  const { fRange, note } = resolveRange(sheet, action.range)
  const grid = fRange.getValues() || []
  const clean = grid.map((row) =>
    (row || []).map((v) => {
      if (v === null || v === undefined) return null
      if (typeof v === 'object' && !Array.isArray(v)) return v.v !== undefined ? v.v : null
      return v
    })
  )
  return { result: { values: clean, a1: action.range }, note }
}

function handleGetRangeFormulas(univerAPI, sheet, action) {
  const { fRange, note } = resolveRange(sheet, action.range)
  const grid = fRange.getFormulas ? fRange.getFormulas() : []
  return { result: { formulas: grid, a1: action.range }, note }
}

function handleGetMerges(univerAPI, sheet, action) {
  const raw = sheet.getMergeData ? sheet.getMergeData() : []
  const list = (raw || []).map((m) => {
    const r = m._range || m
    let a1 = null
    try {
      a1 = m.getA1Notation ? m.getA1Notation() : null
    } catch (e) {}
    return {
      startRow: r.startRow,
      endRow: r.endRow,
      startColumn: r.startColumn,
      endColumn: r.endColumn,
      a1
    }
  })
  return { result: { merges: list } }
}

function handleGetSelection(univerAPI, sheet, action) {
  const ar = sheet.getActiveRange ? sheet.getActiveRange() : null
  const ac = sheet.getActiveCell ? sheet.getActiveCell() : null
  let selection = null
  let activeCell = null
  if (ar) {
    const r = ar._range || ar
    let a1 = null
    try {
      a1 = ar.getA1Notation ? ar.getA1Notation() : null
    } catch (e) {}
    selection = {
      a1,
      startRow: r.startRow,
      endRow: r.endRow,
      startColumn: r.startColumn,
      endColumn: r.endColumn
    }
  }
  if (ac) {
    const r = ac._range || ac
    let a1 = null
    try {
      a1 = ac.getA1Notation ? ac.getA1Notation() : null
    } catch (e) {}
    activeCell = { row: r.startRow, column: r.startColumn, a1 }
  }
  return { result: { selection, activeCell } }
}

function handleSumToCell(univerAPI, sheet, action) {
  // 语义糖：sourceRange 求和填到 targetRange（单格）
  const target = resolveRange(sheet, action.targetRange)
  // 剥 sheet 前缀
  const srcA1 = String(action.sourceRange).replace(/^.*!/, '')
  const formula = '=SUM(' + srcA1 + ')'
  target.fRange.setValueForCell(formula)
  return target.note ? { note: target.note } : null
}

// ============ 筛选 handlers ============

/**
 * 列字母（A/B/AA）→ 0-based 绝对列索引
 * A=0, Z=25, AA=26
 */
function colLetterToIndex(letters) {
  if (typeof letters !== 'string') return -1
  const s = letters.toUpperCase().replace(/[^A-Z]/g, '')
  if (!s) return -1
  let n = 0
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64)
  }
  return n - 1
}

/**
 * 解析 action.column（字母 'C'）或 action.columnIndex（数字）→ 绝对列索引
 */
function resolveColumnIndex(action) {
  if (typeof action.column === 'string' && action.column.trim()) {
    return colLetterToIndex(action.column)
  }
  if (action.columnIndex !== undefined && Number.isFinite(+action.columnIndex)) {
    return +action.columnIndex
  }
  return -1
}

/**
 * setFilter：创建或复用 filter，并对指定列设置条件
 * action: { range?, column/columnIndex, filters?: string[], condition?: {operator, value}, customFormula?: string }
 * range 缺失时自动探测 sheet 的 usedRange（扫描前 SCAN_ROWS×SCAN_COLS 找最后一个非空单元格）
 */
function handleSetFilter(univerAPI, sheet, action) {
  // range 缺失：自动探测 usedRange
  let rangeStr = action.range
  let autoRange = false
  if (!rangeStr) {
    const used = detectUsedRange(sheet)
    if (!used) throw new Error('未指定 range 且无法探测数据区域（工作表可能为空）')
    rangeStr = used
    autoRange = true
  }
  const { fRange, range, note } = resolveRange(sheet, rangeStr)
  const autoNote = autoRange ? '已自动应用 range ' + rangeStr : null
  // 1. 拿到或创建 filter
  let fFilter = null
  if (typeof fRange.getFilter === 'function') {
    fFilter = fRange.getFilter()
  }
  if (!fFilter && typeof fRange.createFilter === 'function') {
    fFilter = fRange.createFilter()
    // createFilter 返回 null 表示已存在但被锁？极少见，再次尝试 getFilter
    if (!fFilter) {
      fFilter = typeof fRange.getFilter === 'function' ? fRange.getFilter() : null
    }
  }
  if (!fFilter) {
    throw new Error('当前 Univer 版本不支持 filter facade API（需 @univerjs/sheets-filter）')
  }

  // 2. 解析目标列（绝对索引）
  const colIdx = resolveColumnIndex(action)
  if (colIdx < 0) throw new Error('column 解析失败: ' + action.column)
  // colId 是相对 filter range 起始列的偏移
  const colId = colIdx - range.startColumn
  if (colId < 0 || colId > range.endColumn - range.startColumn) {
    throw new Error('column ' + action.column + ' 不在 range ' + action.range + ' 内')
  }

  // 3. 构造 criteria
  let criteria
  if (Array.isArray(action.filters) && action.filters.length > 0) {
    // 按值筛选
    criteria = {
      colId,
      filters: { filters: action.filters.map(String) }
    }
  } else if (action.condition && typeof action.condition === 'object') {
    const built = buildCustomFilters(action.condition.operator, action.condition.value)
    if (!built) throw new Error('condition 参数非法')
    criteria = { colId, ...built }
  } else if (typeof action.customFormula === 'string' && action.customFormula.trim()) {
    const built = parseFormulaToCustomFilters(action.customFormula)
    if (!built) throw new Error('customFormula 解析失败')
    criteria = { colId, ...built }
  } else {
    throw new Error('setFilter 缺少筛选条件')
  }

  // 4. 设置（fFilter.setColumnFilterCriteria 接收绝对列索引 + criteria）
  if (typeof fFilter.setColumnFilterCriteria !== 'function') {
    throw new Error('fFilter.setColumnFilterCriteria 不可用')
  }
  fFilter.setColumnFilterCriteria(colIdx, criteria)

  // 拼装 note：自动 range 提示 + 截断提示 + 操作描述
  const parts = []
  if (autoNote) parts.push(autoNote)
  if (note) parts.push(note)
  parts.push('已对 ' + rangeStr + ' 的列 ' + action.column + ' 应用筛选')
  return { note: parts.join('；') }
}

/**
 * clearFilter：清筛选条件 / 移除 filter
 * action: { range?, column/columnIndex?, removeFilter?: boolean }
 */
function handleClearFilter(univerAPI, sheet, action) {
  // 拿 filter 实例（优先 range，否则活动 sheet）
  let fFilter = null
  if (action.range) {
    const { fRange, note } = resolveRange(sheet, action.range)
    if (typeof fRange.getFilter === 'function') {
      fFilter = fRange.getFilter()
    }
    if (!fFilter) return { note: note || '指定 range 上无 filter' }
  } else {
    if (typeof sheet.getFilter === 'function') {
      fFilter = sheet.getFilter()
    }
    if (!fFilter) return { note: '当前 sheet 无 filter' }
  }

  // removeFilter=true：整张 filter 框移除
  if (action.removeFilter === true) {
    if (typeof fFilter.remove !== 'function') throw new Error('fFilter.remove 不可用')
    fFilter.remove()
    return { note: '已移除 filter' }
  }

  // 指定 column：清该列条件
  if (action.column !== undefined || action.columnIndex !== undefined) {
    const colIdx = resolveColumnIndex(action)
    if (colIdx < 0) throw new Error('column 解析失败: ' + action.column)
    if (typeof fFilter.removeColumnFilterCriteria !== 'function') {
      throw new Error('fFilter.removeColumnFilterCriteria 不可用')
    }
    fFilter.removeColumnFilterCriteria(colIdx)
    return { note: '已清除列 ' + action.column + ' 的筛选条件' }
  }

  // 不指定 column：清所有列条件（保留 filter 框）
  if (typeof fFilter.removeFilterCriteria !== 'function') {
    throw new Error('fFilter.removeFilterCriteria 不可用')
  }
  fFilter.removeFilterCriteria()
  return { note: '已清除所有筛选条件' }
}

/**
 * getFilter：读取 filter 范围 + 各列条件 + 被过滤行
 */
function handleGetFilter(univerAPI, sheet, action) {
  let fFilter = null
  let rangeA1 = null
  if (action.range) {
    const { fRange } = resolveRange(sheet, action.range)
    if (typeof fRange.getFilter === 'function') fFilter = fRange.getFilter()
    rangeA1 = action.range
  } else {
    if (typeof sheet.getFilter === 'function') fFilter = sheet.getFilter()
  }
  if (!fFilter) {
    return { result: { hasFilter: false, range: rangeA1, columns: [], filteredOutRows: [] } }
  }
  // filter 自身的 range
  let filterRangeA1 = rangeA1
  try {
    if (typeof fFilter.getRange === 'function') {
      const fr = fFilter.getRange()
      if (fr && typeof fr.getA1Notation === 'function') filterRangeA1 = fr.getA1Notation()
    }
  } catch (e) {}
  // 被过滤的行
  let filteredOutRows = []
  try {
    if (typeof fFilter.getFilteredOutRows === 'function') {
      filteredOutRows = fFilter.getFilteredOutRows() || []
    }
  } catch (e) {}
  // 各列条件：靠遍历 filter range 的列调 getColumnFilterCriteria
  const columns = []
  try {
    if (typeof fFilter.getColumnFilterCriteria === 'function' && typeof fFilter.getRange === 'function') {
      const fr = fFilter.getRange()
      const r = fr._range || fr
      if (r && Number.isFinite(r.startColumn) && Number.isFinite(r.endColumn)) {
        for (let c = r.startColumn; c <= r.endColumn; c++) {
          try {
            const crit = fFilter.getColumnFilterCriteria(c)
            if (crit) columns.push({ column: c, criteria: crit })
          } catch (e) {}
        }
      }
    }
  } catch (e) {}
  return {
    result: {
      hasFilter: true,
      range: filterRangeA1,
      columns,
      filteredOutRows
    }
  }
}

// ============ 工具 ============

// 探测 usedRange 时的最大扫描范围（避免大表全扫，影响性能）
const SCAN_ROWS = 200
const SCAN_COLS = 50

/**
 * 探测 sheet 的有效数据区域，返回 'A1:EndCell' A1 字符串
 * 扫描前 SCAN_ROWS × SCAN_COLS 个单元格，找最后一个非空的位置
 * 探测失败（空表）返回 null
 */
function detectUsedRange(sheet) {
  const rowCount = getRowCount(sheet)
  const colCount = getColCount(sheet)
  if (rowCount <= 0 || colCount <= 0) return null
  const scanRows = Math.min(rowCount, SCAN_ROWS)
  const scanCols = Math.min(colCount, SCAN_COLS)
  let endRow = -1
  let endCol = -1
  try {
    const range = sheet.getRange(0, 0, scanRows, scanCols)
    const grid = range.getValues() || []
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r] || []
      for (let c = 0; c < row.length; c++) {
        const v = row[c]
        if (v === null || v === undefined || v === '') continue
        // 处理 ICellData 对象 { v: ... }
        let real = v
        if (v && typeof v === 'object' && !Array.isArray(v)) {
          real = v.v !== undefined ? v.v : null
        }
        if (real === null || real === undefined || real === '') continue
        if (r > endRow) endRow = r
        if (c > endCol) endCol = c
      }
    }
  } catch (e) {
    return null
  }
  if (endRow < 0 || endCol < 0) return null
  return 'A1:' + toA1(endRow, endCol)
}

function fail(type, range, error, t0) {
  return {
    type,
    range: range || null,
    ok: false,
    error,
    result: null,
    note: null,
    durationMs: elapsed(t0)
  }
}

function elapsed(t0) {
  const t1 = typeof performance !== 'undefined' ? performance.now() : Date.now()
  return Math.round(t1 - t0)
}

function safeCall(fn) {
  try {
    return fn()
  } catch (e) {
    return null
  }
}
