import { parseA1 } from './ai-action-schema'
import { executeActions } from './univer-action-executor'

// Offline B5 validation foundation. The server still rejects these action types.
// This function never submits a receipt or saves a workbook.
const copy = value => JSON.parse(JSON.stringify(value))
function sorted(value) {
  if (Array.isArray(value)) return value.map(sorted)
  if (value && typeof value === 'object') return Object.keys(value).sort().reduce((out, key) => { out[key] = sorted(value[key]); return out }, {})
  return value
}
const equal = (a, b) => JSON.stringify(sorted(a)) === JSON.stringify(sorted(b))
function requireValid(ok, message) { if (!ok) throw new Error('B5_FORMAT_REJECTED: ' + message) }
function uncertain(message) { throw new Error('RESULT_UNKNOWN: ' + message + '；本地结果需核实，不提交保存且不重试') }
function keysOnly(object, allowed) { return Object.keys(object).every(key => allowed.includes(key)) }
function cell(snapshot, sheetId, row, column) { return snapshot.sheets?.[sheetId]?.cellData?.[row]?.[column] }
function styleOf(snapshot, sheetId, row, column) {
  const ref = cell(snapshot, sheetId, row, column)?.s
  return typeof ref === 'string' ? snapshot.styles?.[ref] : ref
}
/** Preflight, execute through the existing action executor, then verify Facade and saved representation. */
export function probeControlledFormatAction(univerAPI, workbook, action) {
  requireValid(action && action.schemaVersion === '2.0' && ['setStyle', 'setRowHeight', 'setColumnWidth'].includes(action.type), '动作类型或版本无效')
  requireValid(action.target && keysOnly(action.target, action.type === 'setStyle' ? ['sheetId', 'range'] : ['sheetId']) && typeof action.target.sheetId === 'string' && action.target.sheetId.length > 0, '目标 Sheet 不明确')
  const style = action.type === 'setStyle'
  const row = action.type === 'setRowHeight'
  const permitted = style ? ['schemaVersion', 'type', 'target', 'background'] : row ? ['schemaVersion', 'type', 'target', 'rowPosition', 'height'] : ['schemaVersion', 'type', 'target', 'columnPosition', 'width']
  requireValid(keysOnly(action, permitted), '存在未开放的样式属性或冲突目标')
  const sheet = workbook.getSheetBySheetId(action.target.sheetId)
  requireValid(sheet && sheet.getSheetId() === action.target.sheetId && Number.isInteger(sheet.getMaxRows()) && Number.isInteger(sheet.getMaxColumns()), '目标 Sheet 无效')
  let index, desired, range
  if (style) {
    range = parseA1(action.target.range)
    requireValid(range && range.startRow >= 0 && range.startColumn >= 0 && range.startRow === range.endRow && range.startColumn === range.endColumn && range.startRow < sheet.getMaxRows() && range.startColumn < sheet.getMaxColumns(), '仅支持 Sheet 内单格样式')
    requireValid(typeof action.background === 'string' && /^#[0-9a-fA-F]{6}$/.test(action.background), 'background 仅支持六位十六进制颜色')
    desired = action.background.toLowerCase()
  } else {
    index = row ? action.rowPosition : action.columnPosition
    desired = row ? action.height : action.width
    requireValid(Number.isInteger(index) && index >= 1 && index <= (row ? sheet.getMaxRows() : sheet.getMaxColumns()), '行列位置越界或非整数')
    requireValid(Number.isInteger(desired) && desired >= (row ? 20 : 40) && desired <= (row ? 400 : 800), '尺寸超出已验证范围')
  }
  const target = style ? sheet.getRange(action.target.range) : sheet
  const setter = style ? 'setBackgroundColor' : row ? 'setRowHeight' : 'setColumnWidth'
  const getter = style ? 'getBackground' : row ? 'getRowHeight' : 'getColumnWidth'
  requireValid(target && typeof target[setter] === 'function' && typeof target[getter] === 'function' && typeof workbook.save === 'function', '缺少 setter、getter 或保存回读')
  if (row) requireValid(desired > sheet.getRowHeight(index - 1), '当前 SDK 较矮行高会切换自动行高命令，本批拒绝')
  const before = copy(workbook.save())
  requireValid(before.sheets?.[action.target.sheetId], '保存快照缺少目标 Sheet')
  if (style) {
    const priorStyle = cell(before, action.target.sheetId, range.startRow, range.startColumn)?.s
    requireValid(priorStyle == null || (typeof priorStyle === 'object' && !Array.isArray(priorStyle)), '样式注册表引用或无效样式未支持')
    requireValid(priorStyle?.bg == null || (typeof priorStyle.bg === 'object' && !Array.isArray(priorStyle.bg)), '背景样式结构未支持')
    requireValid(priorStyle?.bg == null || keysOnly(priorStyle.bg, ['rgb']), '现有背景样式包含 SDK 会丢弃的字段')
  }
  const priorDimension = !style && before.sheets[action.target.sheetId][row ? 'rowData' : 'columnData']?.[index - 1]
  if (priorDimension) {
    requireValid(priorDimension.hd == null || priorDimension.hd === 0 || priorDimension.hd === false, '隐藏行列未证明安全')
    if (row) requireValid((priorDimension.ah == null || priorDimension.ah === 0 || priorDimension.ah === false) && (priorDimension.ia == null || priorDimension.ia === 0), '自动行高未证明安全')
  }
  if (row) requireValid(Number.isInteger(before.sheets[action.target.sheetId].defaultRowHeight) && desired > before.sheets[action.target.sheetId].defaultRowHeight, '未能排除 SDK 自动行高分支')
  const executed = executeActions(univerAPI, workbook, [action], { continueOnError: false })
  if (executed.summary.ok !== 1 || executed.summary.failed || executed.summary.skipped || executed.notes.length) uncertain(executed.results[0]?.error || '样式或尺寸执行未完整成功')
  let actual, after
  try { actual = style ? target.getBackground() : target[getter](index - 1); after = copy(workbook.save()) } catch (_) { uncertain('实际回读失败') }
  if (style ? typeof actual !== 'string' || actual.toLowerCase() !== desired : actual !== desired) uncertain('Facade 实际回读与预期不符')
  const afterSheet = after.sheets?.[action.target.sheetId]
  if (!afterSheet) uncertain('保存快照缺少目标 Sheet')
  const expected = copy(before)
  if (style) {
    const saved = styleOf(after, action.target.sheetId, range.startRow, range.startColumn)
    if (!saved || saved.bg?.rgb?.toLowerCase() !== desired) uncertain('保存快照样式回读不符')
    const prior = cell(before, action.target.sheetId, range.startRow, range.startColumn)
    const expectedSheet = expected.sheets[action.target.sheetId]
    if (!expectedSheet.cellData) expectedSheet.cellData = {}
    if (!expectedSheet.cellData[range.startRow]) expectedSheet.cellData[range.startRow] = {}
    if (!expectedSheet.cellData[range.startRow][range.startColumn]) expectedSheet.cellData[range.startRow][range.startColumn] = {}
    const expectedCell = expectedSheet.cellData[range.startRow][range.startColumn]
    if (!expectedCell.s) expectedCell.s = {}
    if (!expectedCell.s.bg) expectedCell.s.bg = {}
    expectedCell.s.bg.rgb = desired
    // Univer 0.25.1 saves a style setter as a cell registry reference. Accept
    // only the exact expected style under that reference and no other change.
    const savedRef = cell(after, action.target.sheetId, range.startRow, range.startColumn)?.s
    if (!equal(saved, expectedCell.s)) uncertain('保存快照包含未授权的样式字段')
    if (typeof savedRef === 'string') {
      if (!/^[A-Za-z0-9_-]{1,64}$/.test(savedRef)) uncertain('样式注册表引用格式无效')
      if (!expected.styles) expected.styles = {}
      if (Object.prototype.hasOwnProperty.call(expected.styles, savedRef) && !equal(expected.styles[savedRef], expectedCell.s)) uncertain('样式注册表引用与原有样式冲突')
      expected.styles[savedRef] = copy(expectedCell.s)
      expectedCell.s = savedRef
    }
  } else {
    const key = String(index - 1), field = row ? 'h' : 'w', table = row ? 'rowData' : 'columnData'
    if (afterSheet[table]?.[key]?.[field] !== desired || (row && afterSheet.rowData?.[key]?.ah)) uncertain('保存快照尺寸回读不符')
    if (!expected.sheets[action.target.sheetId][table]) expected.sheets[action.target.sheetId][table] = {}
    if (!expected.sheets[action.target.sheetId][table][key]) expected.sheets[action.target.sheetId][table][key] = {}
    expected.sheets[action.target.sheetId][table][key][field] = desired
    // SetRowHeightCommand also writes the explicit non-auto-height flag.
    if (row && afterSheet.rowData?.[key]?.ia === 0) expected.sheets[action.target.sheetId].rowData[key].ia = 0
  }
  if (!equal(after, expected)) uncertain('目标外或未授权字段发生变化')
  const formatReadback = style
    ? { sheetId: action.target.sheetId, row: range.startRow, column: range.startColumn, value: desired }
    : { sheetId: action.target.sheetId, index: index - 1, value: desired }
  return { verified: true, type: action.type, sheetId: action.target.sheetId, formatReadback }
}
