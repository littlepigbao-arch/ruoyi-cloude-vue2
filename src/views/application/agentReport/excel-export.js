import * as XLSX from 'xlsx'

const ROWS = 1048576, COLS = 16384, SCAN_BUDGET = 10000000, CELL_BUDGET = 1000000
const own = (object, key) => Object.prototype.hasOwnProperty.call(object, key)
function index(key, bound) {
  if (!/^(0|[1-9]\d*)$/.test(String(key)) || !Number.isSafeInteger(Number(key)) || Number(key) >= bound) throw new Error('数据坐标超出 Excel 容量，请下载完整 JSON 备份')
  return Number(key)
}
function cellForExcel(cell, styles) {
  if (!cell || typeof cell !== 'object' || Array.isArray(cell)) throw new Error('单元格数据无效，已停止导出')
  if (cell.si != null && cell.si !== '') throw new Error('共享公式需先展开后再导出，请下载完整 JSON 备份')
  let value = cell.v
  if (cell.p) {
    if (!cell.p.body || typeof cell.p.body.dataStream !== 'string') throw new Error('富文本内容不完整，请下载完整 JSON 备份')
    value = cell.p.body.dataStream.replace(/\r\n$/, '')
  }
  const formula = cell.f == null || cell.f === '' ? null : cell.f
  if (formula !== null && (typeof formula !== 'string' || !formula.replace(/^=/, '').trim())) throw new Error('公式无效，已停止导出')
  if (value != null && !['number', 'string', 'boolean'].includes(typeof value)) throw new Error('单元格包含无法转换的值，请下载完整 JSON 备份')
  if (typeof value === 'number' && !Number.isFinite(value)) throw new Error('单元格含非有限数字，已停止导出')
  if (value == null && !formula) return null
  const out = { t: typeof value === 'number' ? 'n' : typeof value === 'boolean' ? 'b' : value == null ? 'n' : 's' }
  if (value != null) out.v = value
  if (formula) out.f = formula.replace(/^=/, '')
  const errors = { '#NULL!': 0, '#DIV/0!': 7, '#VALUE!': 15, '#REF!': 23, '#NAME?': 29, '#NUM!': 36, '#N/A': 42, '#GETTING_DATA': 43 }
  if (formula && typeof value === 'string' && own(errors, value)) { out.t = 'e'; out.v = errors[value] }
  const style = typeof cell.s === 'string' ? styles[cell.s] : cell.s
  if (style && style.n && typeof style.n.pattern === 'string') out.z = style.n.pattern
  return out
}

/** Sparse conversion; configured blank grid capacity never creates data or truncates cells. */
export function exportSheet(sheet, styles = {}) {
  if (!sheet || (sheet.cellData != null && (typeof sheet.cellData !== 'object' || Array.isArray(sheet.cellData)))) throw new Error('工作表数据不完整')
  const cellData = sheet.cellData || {}
  const ws = {}, merges = [], rows = [], columns = []
  let minRow = ROWS, minCol = COLS, maxRow = -1, maxCol = -1, count = 0
  const include = (r, c) => { minRow = Math.min(minRow, r); minCol = Math.min(minCol, c); maxRow = Math.max(maxRow, r); maxCol = Math.max(maxCol, c) }
  Object.keys(cellData).forEach(rk => {
    const r = index(rk, ROWS), row = cellData[rk]
    if (!row || typeof row !== 'object' || Array.isArray(row)) throw new Error('单元格行数据无效')
    Object.keys(row).forEach(ck => {
      const c = index(ck, COLS)
      if (++count > CELL_BUDGET) throw new Error('单元格数量超出本次 Excel 导出预算，请下载完整 JSON 备份')
      const cell = cellForExcel(row[ck], styles)
      if (cell) { ws[XLSX.utils.encode_cell({ r, c })] = cell; include(r, c) }
    })
  })
  if (sheet.mergeData != null && !Array.isArray(sheet.mergeData)) throw new Error('合并范围无效')
  if ((sheet.mergeData || []).length > 10000) throw new Error('合并范围过多，请下载完整 JSON 备份')
  for (const merge of sheet.mergeData || []) {
    const r = index(merge.startRow, ROWS), er = index(merge.endRow, ROWS), c = index(merge.startColumn, COLS), ec = index(merge.endColumn, COLS)
    if (er < r || ec < c || merges.some(m => r <= m.e.r && er >= m.s.r && c <= m.e.c && ec >= m.s.c)) throw new Error('合并范围倒置或重叠，已停止导出')
    merges.push({ s: { r, c }, e: { r: er, c: ec } }); include(r, c); include(er, ec)
  }
  for (const [key, data] of Object.entries(sheet.rowData || {})) {
    const r = index(key, ROWS)
    if (!data || typeof data !== 'object') throw new Error('行尺寸数据无效')
    if (data.h != null && (!Number.isFinite(data.h) || data.h < 0)) throw new Error('行高无效')
    if (data.h != null || data.hd != null) rows[r] = { ...(data.h > 0 ? { hpx: data.h } : {}), hidden: data.hd === 1 || data.h === 0 }
  }
  for (const [key, data] of Object.entries(sheet.columnData || {})) {
    const c = index(key, COLS)
    if (!data || typeof data !== 'object') throw new Error('列尺寸数据无效')
    if (data.w != null && (!Number.isFinite(data.w) || data.w < 0)) throw new Error('列宽无效')
    if (data.w != null || data.hd != null) columns[c] = { ...(data.w > 0 ? { wpx: data.w } : {}), hidden: data.hd === 1 || data.w === 0 }
  }
  const scan = maxRow < 0 ? 1 : (maxRow - minRow + 1) * (maxCol - minCol + 1)
  if (scan + rows.length + columns.length > SCAN_BUDGET) throw new Error('Excel 序列化范围超出本次导出预算，请下载完整 JSON 备份')
  ws['!ref'] = maxRow < 0 ? 'A1' : XLSX.utils.encode_range({ s: { r: minRow, c: minCol }, e: { r: maxRow, c: maxCol } })
  if (merges.length) ws['!merges'] = merges
  if (rows.length) ws['!rows'] = rows
  if (columns.length) ws['!cols'] = columns
  return ws
}

export function exportWorkbook(snapshot) {
  if (!snapshot || !snapshot.sheets || typeof snapshot.sheets !== 'object' || Array.isArray(snapshot.sheets)) throw new Error('工作簿快照不完整')
  const order = snapshot.sheetOrder == null ? Object.keys(snapshot.sheets) : snapshot.sheetOrder
  if (!Array.isArray(order) || !order.length || new Set(order).size !== order.length || order.length !== Object.keys(snapshot.sheets).length) throw new Error('工作表顺序不完整或重复，已停止导出')
  const workbook = XLSX.utils.book_new(), names = new Set(), metadata = []
  let totalBudget = 0
  for (const id of order) {
    if (!own(snapshot.sheets, id)) throw new Error('工作表顺序引用不存在的表')
    const sheet = snapshot.sheets[id], name = sheet && sheet.name || 'Sheet'
    if (typeof name !== 'string' || !name.trim() || name.length > 31 || /[\[\]:*?/\\]/.test(name) || /^'|'$/.test(name) || names.has(name.toLowerCase())) throw new Error('工作表名称不符合 Excel 规则或重名，请先重命名或下载完整 JSON 备份')
    names.add(name.toLowerCase())
    const ws = exportSheet(sheet, snapshot.styles || {}), range = XLSX.utils.decode_range(ws['!ref'])
    totalBudget += (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1) + (ws['!rows'] || []).length + (ws['!cols'] || []).length
    if (totalBudget > SCAN_BUDGET) throw new Error('整个工作簿的 Excel 序列化范围超出预算，请下载完整 JSON 备份')
    XLSX.utils.book_append_sheet(workbook, ws, name)
    metadata.push({ name, Hidden: sheet.hidden === 1 ? 1 : 0 })
  }
  if (metadata.every(s => s.Hidden === 1)) throw new Error('Excel 至少需要一个可见工作表，请下载完整 JSON 备份')
  workbook.Workbook = { Sheets: metadata, CalcPr: { fullCalcOnLoad: true } }
  return workbook
}
