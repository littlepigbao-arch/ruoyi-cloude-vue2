/**
 * Univer 表格上下文采集
 * 把当前活动 sheet 的状态摘要成一个紧凑对象，发给 AI 后端（或 mock 解析器）
 * 控制 token 成本：最多 50 行 × 26 列，无公式时不发 formulas 字段
 */

const MAX_ROWS = 50
const MAX_COLS = 26
// 探测 usedRange 时的最大扫描范围（避免 1 万行全扫）
const SCAN_ROWS = 200
const SCAN_COLS = 50
// merges 最多采集多少个
const MAX_MERGES = 50

/**
 * 采集当前活动 sheet 的紧凑上下文
 * @param {object} univerAPI
 * @returns {object|null} context 对象；无活动工作簿/sheet 时返回 null
 */
export function collectContext(workbook) {
  if (!workbook) return null
  const sheet = safeCall(() => workbook.getActiveSheet()) || safeCall(() => (workbook.getSheets() || [])[0])
  if (!sheet) return null

  // 兼容 facade（getMaxRows/getMaxColumns）与核心（getRowCount/getColumnCount）
  const rowCount = safeCall(() => sheet.getMaxRows && sheet.getMaxRows()) || safeCall(() => sheet.getRowCount && sheet.getRowCount()) || 0
  const columnCount = safeCall(() => sheet.getMaxColumns && sheet.getMaxColumns()) || safeCall(() => sheet.getColumnCount && sheet.getColumnCount()) || 0

  // 1. 探测 usedRange（最后一个非空单元格的行列）
  const { endRow, endColumn } = detectUsedRange(sheet, rowCount, columnCount)

  // 2. 截断到 MAX_ROWS × MAX_COLS，逐格取 value/formula
  const capRows = Math.min(endRow + 1, MAX_ROWS)
  const capCols = Math.min(endColumn + 1, MAX_COLS)
  const values = []
  const formulas = []
  let hasFormula = false
  for (let r = 0; r < capRows; r++) {
    const vRow = []
    const fRow = []
    for (let c = 0; c < capCols; c++) {
      const cell = safeGetCell(sheet, r, c)
      vRow.push(cell.value)
      fRow.push(cell.formula)
      if (cell.formula) hasFormula = true
    }
    values.push(vRow)
    formulas.push(fRow)
  }

  // 3. merges
  let merges = []
  const rawMerges = safeCall(() => sheet.getMergeData()) || []
  if (Array.isArray(rawMerges)) {
    merges = rawMerges.slice(0, MAX_MERGES).map(extractMerge)
    merges = merges.filter((m) => m !== null)
  }

  // 4. selection + activeCell
  let selection = null
  let activeCell = null
  const ar = safeCall(() => sheet.getActiveRange && sheet.getActiveRange())
  if (ar) {
    const a1 = safeCall(() => ar.getA1Notation && ar.getA1Notation())
    const r = ar._range || ar
    selection = {
      a1: a1 || null,
      startRow: r.startRow,
      endRow: r.endRow,
      startColumn: r.startColumn,
      endColumn: r.endColumn
    }
  }
  const ac = safeCall(() => sheet.getActiveCell && sheet.getActiveCell())
  if (ac) {
    const a1 = safeCall(() => ac.getA1Notation && ac.getA1Notation())
    const r = ac._range || ac
    activeCell = { row: r.startRow, column: r.startColumn, a1: a1 || null }
  }

  const workbookName = safeCall(() => workbook.getName()) || '工作簿'
  const sheetName = safeCall(() => sheet.getSheetName && sheet.getSheetName()) ||
    safeCall(() => sheet.getName && sheet.getName()) || 'Sheet1'
  const sheetId = safeCall(() => sheet.getSheetId && sheet.getSheetId()) || null

  const ctx = {
    workbookName,
    sheetName,
    sheetId,
    rowCount,
    columnCount,
    usedRange: { startRow: 0, endRow, startColumn: 0, endColumn },
    values,
    merges,
    selection,
    activeCell,
    truncated: endRow + 1 > MAX_ROWS || endColumn + 1 > MAX_COLS
  }
  if (hasFormula) ctx.formulas = formulas
  return ctx
}

/**
 * 探测有效数据区域：扫描最多 SCAN_ROWS × SCAN_COLS，找最后一个非空单元格
 */
function detectUsedRange(sheet, rowCount, columnCount) {
  let endRow = 0
  let endColumn = 0
  const scanRows = Math.min(rowCount, SCAN_ROWS)
  const scanCols = Math.min(columnCount, SCAN_COLS)
  try {
    const range = sheet.getRange(0, 0, scanRows, scanCols)
    const grid = range.getValues() || []
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r] || []
      for (let c = 0; c < row.length; c++) {
        const v = row[c]
        if (v !== null && v !== undefined && v !== '') {
          if (r > endRow) endRow = r
          if (c > endColumn) endColumn = c
        }
      }
    }
  } catch (e) {
    // 扫描失败，退化到 0,0
  }
  return { endRow, endColumn }
}

/**
 * 安全读取单个单元格的 value 和 formula
 */
function safeGetCell(sheet, r, c) {
  try {
    const range = sheet.getRange(r, c, 1, 1)
    const v = range.getValue()
    // getValue 可能返回裸值或 ICellData 对象 {v, p, ...}
    let value = v
    if (v && typeof v === 'object' && !Array.isArray(v)) value = v.v !== undefined ? v.v : null
    let formula = null
    try {
      const f = range.getFormulas && range.getFormulas()
      if (f && f[0] && f[0][0]) formula = f[0][0]
    } catch (e) {
      // ignore
    }
    return { value, formula }
  } catch (e) {
    return { value: null, formula: null }
  }
}

/**
 * 从 mergeData 项提取 {startRow,endRow,startColumn,endColumn,a1}
 * 防御式取值：兼容多种返回结构
 */
function extractMerge(m) {
  if (!m) return null
  const a1 = safeCall(() => (m.getA1Notation && m.getA1Notation()) || null)
  const r = m._range || m
  if (r.startRow === undefined || r.startColumn === undefined) return null
  return {
    startRow: r.startRow,
    endRow: r.endRow,
    startColumn: r.startColumn,
    endColumn: r.endColumn,
    a1
  }
}

/**
 * 安全调用，抛错返回 null
 */
function safeCall(fn) {
  try {
    return fn()
  } catch (e) {
    return null
  }
}
