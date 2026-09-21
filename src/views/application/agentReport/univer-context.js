/**
 * Univer 表格上下文采集
 * 把当前活动 sheet 的状态摘要成一个紧凑对象，发给 AI 后端（或 mock 解析器）
 * 控制 token 成本：values 只发预览（maxRows × maxCols），完整范围由 usedRange/dataRange 承载
 * 无公式时不发 formulas 字段
 *
 * 各项上限均可按需覆盖：在页面加载前设置 window.__AI_CONTEXT_CONFIG__ = { maxRows, maxCols, scanRows, scanCols, maxMerges }
 */

const DEFAULT_CONFIG = {
  // 最多向模型提供多少个 Sheet 摘要
  maxSheets: 20,
  // values 预览上限（token 节流）；range 判定不依赖它，模型应使用 usedRange/dataRange
  maxRows: 50,
  maxCols: 26,
  // 探测 usedRange 时的最大扫描范围（应覆盖全量数据，避免 range 判错）
  scanRows: 5000,
  scanCols: 200,
  // merges 最多采集多少个
  maxMerges: 50
}

export function getContextConfig() {
  const win = typeof window !== 'undefined' ? window : {}
  const overrides = (win && win.__AI_CONTEXT_CONFIG__) || {}
  return Object.assign({}, DEFAULT_CONFIG, overrides)
}

/**
 * 采集当前活动 sheet 的紧凑上下文
 * @param {object} univerAPI
 * @returns {object|null} context 对象；无活动工作簿/sheet 时返回 null
 */
export function collectContext(workbook) {
  if (!workbook) return null
  const cfg = getContextConfig()
  const allSheets = safeCall(() => workbook.getSheets && workbook.getSheets()) || []
  const activeSheet = safeCall(() => workbook.getActiveSheet()) || allSheets[0]
  if (!activeSheet) return null
  const selectedSheets = allSheets.slice(0, cfg.maxSheets)
  const sheets = selectedSheets.map(sheet => collectSheetContext(workbook, sheet, cfg)).filter(Boolean)
  const activeSheetId = getSheetId(activeSheet)
  const activeSheetName = getSheetName(activeSheet)
  return {
    workbookName: safeCall(() => workbook.getName()) || '工作簿',
    activeSheetId,
    activeSheetName,
    sheets,
    truncated: allSheets.length > selectedSheets.length
  }
}

/** 采集单个 Sheet 的受控摘要。 */
function collectSheetContext(workbook, sheet, cfg) {
  if (!sheet) return null

  // 兼容 facade（getMaxRows/getMaxColumns）与核心（getRowCount/getColumnCount）
  const rowCount = safeCall(() => sheet.getMaxRows && sheet.getMaxRows()) || safeCall(() => sheet.getRowCount && sheet.getRowCount()) || 0
  const columnCount = safeCall(() => sheet.getMaxColumns && sheet.getMaxColumns()) || safeCall(() => sheet.getColumnCount && sheet.getColumnCount()) || 0

  // 1. 探测 usedRange（最后一个非空单元格的行列）
  const { endRow, endColumn } = detectUsedRange(sheet, rowCount, columnCount, cfg)

  // 2. 截断到 maxRows × maxCols，逐格取 value/formula
  const capRows = Math.min(endRow + 1, cfg.maxRows)
  const capCols = Math.min(endColumn + 1, cfg.maxCols)
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
    merges = rawMerges.slice(0, cfg.maxMerges).map(extractMerge)
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
  const sheetName = getSheetName(sheet)
  const sheetId = getSheetId(sheet)

  const ctx = {
    workbookName,
    sheetName,
    sheetId,
    rowCount,
    columnCount,
    usedRange: { startRow: 0, endRow, startColumn: 0, endColumn },
    dataRange: toA1(0, 0) + ':' + toA1(endRow, endColumn),
    values,
    merges,
    selection,
    activeCell,
    truncated: endRow + 1 > cfg.maxRows || endColumn + 1 > cfg.maxCols
  }
  if (hasFormula) ctx.formulas = formulas
  return ctx
}

function getSheetName(sheet) {
  return safeCall(() => sheet.getSheetName && sheet.getSheetName()) ||
    safeCall(() => sheet.getName && sheet.getName()) || 'Sheet1'
}

function getSheetId(sheet) {
  return safeCall(() => sheet.getSheetId && sheet.getSheetId()) ||
    safeCall(() => sheet.getId && sheet.getId()) || null
}

/**
 * 0-based 行列号 → A1 字符串（如 (0,0)='A1'、(3,1)='B4'）
 */
function toA1(row, col) {
  let c = col + 1
  let s = ''
  while (c > 0) {
    const m = (c - 1) % 26
    s = String.fromCharCode(65 + m) + s
    c = Math.floor((c - 1) / 26)
  }
  return s + (row + 1)
}

/**
 * 探测有效数据区域：扫描最多 scanRows × scanCols，找最后一个非空单元格
 */
function detectUsedRange(sheet, rowCount, columnCount, cfg) {
  let endRow = 0
  let endColumn = 0
  const scanRows = Math.min(rowCount, cfg.scanRows)
  const scanCols = Math.min(columnCount, cfg.scanCols)
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
