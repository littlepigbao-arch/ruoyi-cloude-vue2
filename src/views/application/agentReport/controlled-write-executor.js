import { executeActions } from './univer-action-executor'
import { parseA1, toA1 } from './ai-action-schema'

const plain = value => JSON.parse(JSON.stringify(value))
const scalar = value => value === null || ['string', 'number', 'boolean'].includes(typeof value)
const stable = value => JSON.stringify(sort(value))
function sort(value) {
  if (Array.isArray(value)) return value.map(sort)
  if (value && typeof value === 'object') return Object.keys(value).sort().reduce((out, key) => { out[key] = sort(value[key]); return out }, {})
  return value
}
const key = cell => cell.sheetId + ':' + cell.row + ':' + cell.column
export const sameControlledSnapshot = (left, right) => stable(left) === stable(right)
function masked(snapshot, cells) {
  const copy = plain(snapshot)
  for (const cell of cells) {
    const data = copy.sheets[cell.sheetId].cellData || {}
    const row = data[cell.row]
    if (row) { delete row[cell.column]; if (!Object.keys(row).length) delete data[cell.row] }
    if (!Object.keys(data).length) delete copy.sheets[cell.sheetId].cellData
  }
  return copy
}

/** Thin transport adapter: all mutations go through the existing executeActions. Never recreates a unit. */
export async function executeControlledWrite(univerAPI, workbook, manifest) {
  if (!manifest || manifest.transportVersion !== 2 || !Array.isArray(manifest.actions) || !manifest.actions.length || manifest.actions.length > 30 || !Array.isArray(manifest.cells) || !manifest.cells.length || manifest.cells.length > 1000) throw new Error('受控清单无效')
  const expected = new Map()
  for (const cell of manifest.cells) {
    if (!Number.isInteger(cell.row) || !Number.isInteger(cell.column) || cell.row < 0 || cell.column < 0 || !scalar(cell.value) || expected.has(key(cell))) throw new Error('受控单元格无效或重复')
    expected.set(key(cell), cell)
  }
  const actionCells = new Set()
  for (const action of manifest.actions) {
    if (action.schemaVersion !== '2.0' || !action.target || !action.target.sheetId || !['setCell', 'setValues'].includes(action.type)) throw new Error('受控动作类型或目标无效')
    const sheet = workbook.getSheetBySheetId(action.target.sheetId)
    if (!sheet) throw new Error('目标 Sheet 不存在')
    const r = parseA1(action.target.range)
    if (!r || r.endRow >= sheet.getMaxRows() || r.endColumn >= sheet.getMaxColumns()) throw new Error('受控范围超出 Sheet，不允许截断')
    if (action.type === 'setCell' && (r.startRow !== r.endRow || r.startColumn !== r.endColumn)) throw new Error('setCell 只能写单格')
    for (let row = r.startRow; row <= r.endRow; row++) for (let column = r.startColumn; column <= r.endColumn; column++) {
      const id = key({ sheetId: action.target.sheetId, row, column })
      const value = action.type === 'setCell' ? action.value : action.values?.[row - r.startRow]?.[column - r.startColumn]
      if (!scalar(value) || (typeof value === 'string' && value.trim().startsWith('=')) || actionCells.has(id) || !expected.has(id) || stable(expected.get(id).value) !== stable(value)) throw new Error('动作与授权单元格不一致')
      actionCells.add(id)
    }
  }
  if (actionCells.size !== expected.size) throw new Error('受控清单范围不完整')
  const before = plain(workbook.save())
  for (const cell of manifest.cells) {
    const actual = before.sheets?.[cell.sheetId]?.cellData?.[cell.row]?.[cell.column]
    if (stable(actual?.v ?? null) !== stable(cell.before?.v ?? null) || (actual?.f || '') !== (cell.before?.f || '')) throw new Error('本地目标已变化，保持保护并核实')
  }
  const result = executeActions(univerAPI, workbook, manifest.actions, { continueOnError: false })
  if (result.summary.failed || result.summary.skipped || result.summary.ok !== manifest.actions.length || result.notes.length) throw new Error('受控执行失败或未完整执行，不能报告完成')
  const cells = []
  for (const expectedCell of manifest.cells) {
    const range = workbook.getSheetBySheetId(expectedCell.sheetId).getRange(toA1(expectedCell.row, expectedCell.column))
    if (typeof range.getFormulas !== 'function' || typeof range.getValues !== 'function') throw new Error('当前版本无法独立回读单元格')
    const value = range.getValues()?.[0]?.[0] ?? null
    const formula = range.getFormulas()?.[0]?.[0] || ''
    if (formula || stable(value) !== stable(expectedCell.value)) throw new Error('实际回读与授权值不符')
    cells.push({ sheetId: expectedCell.sheetId, row: expectedCell.row, column: expectedCell.column, value, formula })
  }
  const after = plain(workbook.save())
  if (stable(masked(before, manifest.cells)) !== stable(masked(after, manifest.cells))) throw new Error('发现授权范围外的本地变化，保持保护并核实')
  return { ok: true, cells }
}
