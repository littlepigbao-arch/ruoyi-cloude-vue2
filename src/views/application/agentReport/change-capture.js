import { parseA1, rangeToA1 } from './ai-action-schema'
import { normalizeAction } from './univer-action-executor'

const READ_ONLY_TYPES = new Set(['getRangeValue', 'getRangeFormulas', 'getMerges', 'getSelection'])

const KIND_BY_TYPE = {
  setCell: 'cell-value',
  setCellForCell: 'cell-value',
  setValues: 'cell-value',
  clearRange: 'cell-value',
  setFormula: 'formula',
  merge: 'merge',
  breakApart: 'merge',
  insertRows: 'row-insert',
  deleteRows: 'row-delete',
  insertColumns: 'column-insert',
  deleteColumns: 'column-delete',
  setRowHeight: 'style',
  setColumnWidth: 'style',
  setStyle: 'style',
  setFilter: 'filter',
  clearFilter: 'filter',
  createChart: 'object',
  undo: 'snapshot-replace',
  redo: 'snapshot-replace'
}

export function createOperationId() {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID()
  return 'op-' + Date.now() + '-' + Math.random().toString(16).slice(2)
}

/** 根据执行前后 Snapshot 生成可审计的实际变化。 */
export function buildExecutedChanges(actions, results, beforeSnapshot, afterSnapshot, meta = {}) {
  const groupId = meta.groupId || createOperationId().replace(/^op-/, 'grp-')
  const operationId = meta.operationId || createOperationId()
  const changes = []
  ;(actions || []).forEach((raw, index) => {
    const result = results && results[index]
    if (!result || !result.ok) return
    const action = normalizeAction(raw)
    if (!action || READ_ONLY_TYPES.has(action.type)) return
    const sheetId = result.sheetId || action.sheetId || null
    const sheetName = result.sheetName || action.sheetName || null
    const range = action.range || null
    const beforeSheet = findSheet(beforeSnapshot, sheetId, sheetName)
    const afterSheet = findSheet(afterSnapshot, sheetId, sheetName)
    const parsed = range ? parseA1(range) : null
    const rangeBefore = parsed ? readRange(beforeSheet, parsed) : sheetSummary(beforeSheet)
    const rangeAfter = parsed ? readRange(afterSheet, parsed) : sheetSummary(afterSheet)
    if (JSON.stringify(rangeBefore) === JSON.stringify(rangeAfter) && action.type !== 'createChart') return
    changes.push({
      schemaVersion: '1.0',
      changeId: createOperationId().replace(/^op-/, 'chg-'),
      groupId,
      operationId,
      source: meta.source || 'ai',
      kind: KIND_BY_TYPE[action.type] || 'snapshot-replace',
      actionType: action.type,
      sheetId,
      sheetName,
      range,
      before: rangeBefore,
      after: rangeAfter,
      risk: inferRisk(action),
      needsSnapshotDiff: !parsed,
      occurredAt: new Date().toISOString()
    })
  })
  return { operationId, groupId, changes }
}

/** 将一组人工顶层 Command 转换为统一的 executedChanges[]。 */
export function buildManualChanges(commands, beforeSnapshot, afterSnapshot, meta = {}) {
  const operationId = meta.operationId || createOperationId()
  const groupId = meta.groupId || createOperationId().replace(/^op-/, 'grp-')
  const changes = []
  ;(commands || []).forEach(command => {
    if (!command || !command.id) return
    const params = command.params || {}
    const sheetId = params.subUnitId || params.sheetId || null
    const beforeSheet = findSheet(beforeSnapshot, sheetId, null)
    const afterSheet = findSheet(afterSnapshot, sheetId, null)
    const sheetName = (afterSheet && afterSheet.name) || (beforeSheet && beforeSheet.name) || null
    const rawRange = normalizeCommandRange(params.range || params.targetRange)
    const range = rawRange ? rangeToA1(rawRange) : null
    const before = rawRange ? readRange(beforeSheet, rawRange) : sheetSummary(beforeSheet)
    const after = rawRange ? readRange(afterSheet, rawRange) : sheetSummary(afterSheet)
    if (rawRange && JSON.stringify(before) === JSON.stringify(after)) return
    changes.push({
      schemaVersion: '1.0',
      changeId: createOperationId().replace(/^op-/, 'chg-'),
      groupId,
      operationId,
      source: 'manual',
      kind: inferManualKind(command.id),
      actionType: command.id,
      sheetId,
      sheetName,
      range,
      before,
      after,
      risk: inferManualRisk(command.id),
      needsSnapshotDiff: !rawRange,
      occurredAt: new Date().toISOString()
    })
  })
  return { operationId, groupId, changes }
}

/**
 * 大数据模式下只记录命令元数据，不复制整本工作簿。
 * 最终保存/提交时仍会生成完整快照；这里只放弃逐次变更的 before/after，
 * 以避免一次编辑触发多次大对象拷贝。
 */
export function buildLightweightManualChanges(commands, meta = {}) {
  const operationId = meta.operationId || createOperationId()
  const groupId = meta.groupId || createOperationId().replace(/^op-/, 'grp-')
  const changes = []
  const seen = new Set()
  ;(commands || []).forEach(command => {
    if (!command || !command.id) return
    const params = command.params || {}
    const rawRange = normalizeCommandRange(params.range || params.targetRange)
    const range = rawRange ? rangeToA1(rawRange) : null
    const sheetId = params.subUnitId || params.sheetId || null
    const dedupeKey = String(command.id) + '|' + String(sheetId || '') + '|' + String(range || '')
    if (seen.has(dedupeKey)) return
    seen.add(dedupeKey)
    changes.push({
      schemaVersion: '1.0',
      changeId: createOperationId().replace(/^op-/, 'chg-'),
      groupId,
      operationId,
      source: 'manual',
      kind: inferManualKind(command.id),
      actionType: command.id,
      sheetId,
      sheetName: null,
      range,
      before: null,
      after: null,
      risk: inferManualRisk(command.id),
      needsSnapshotDiff: true,
      lightweight: true,
      occurredAt: new Date().toISOString()
    })
  })
  return { operationId, groupId, changes }
}

function normalizeCommandRange(range) {
  if (!range || range.startRow === undefined || range.startColumn === undefined) return null
  return {
    startRow: Number(range.startRow),
    endRow: Number(range.endRow === undefined ? range.startRow : range.endRow),
    startColumn: Number(range.startColumn),
    endColumn: Number(range.endColumn === undefined ? range.startColumn : range.endColumn)
  }
}

function inferManualKind(commandId) {
  const id = String(commandId).toLowerCase()
  if (id.includes('formula')) return 'formula'
  if (id.includes('style') || id.includes('format') || id.includes('height') || id.includes('width')) return 'style'
  if (id.includes('merge')) return 'merge'
  if (id.includes('insert-row')) return 'row-insert'
  if (id.includes('remove-row') || id.includes('delete-row')) return 'row-delete'
  if (id.includes('insert-col')) return 'column-insert'
  if (id.includes('remove-col') || id.includes('delete-col')) return 'column-delete'
  if (id.includes('sheet') && (id.includes('remove') || id.includes('delete'))) return 'sheet-delete'
  if (id.includes('sheet') && (id.includes('insert') || id.includes('add'))) return 'sheet-create'
  if (id.includes('sort')) return 'sort'
  if (id.includes('filter')) return 'filter'
  return 'cell-value'
}

function inferManualRisk(commandId) {
  const id = String(commandId).toLowerCase()
  if (id.includes('delete') || id.includes('remove') || id.includes('clear') || id.includes('sort')) return 'high'
  if (id.includes('insert') || id.includes('merge') || id.includes('filter')) return 'medium'
  return 'low'
}

function findSheet(snapshot, sheetId, sheetName) {
  if (!snapshot || !snapshot.sheets) return null
  if (sheetId && snapshot.sheets[sheetId]) return snapshot.sheets[sheetId]
  const ids = snapshot.sheetOrder || Object.keys(snapshot.sheets)
  for (let i = 0; i < ids.length; i++) {
    const sheet = snapshot.sheets[ids[i]]
    if (sheet && sheet.name === sheetName) return sheet
  }
  return null
}

function readRange(sheet, range) {
  if (!sheet) return null
  const rows = []
  const data = sheet.cellData || {}
  for (let r = range.startRow; r <= range.endRow; r++) {
    const row = []
    for (let c = range.startColumn; c <= range.endColumn; c++) {
      const cell = data[r] && data[r][c]
      row.push(cell ? compactCell(cell) : null)
    }
    rows.push(row)
  }
  return { cells: rows }
}

function compactCell(cell) {
  return {
    value: cell.v !== undefined ? cell.v : null,
    formula: cell.f || null,
    style: cell.s !== undefined ? cell.s : null
  }
}

function sheetSummary(sheet) {
  if (!sheet) return null
  return {
    id: sheet.id || null,
    name: sheet.name || null,
    rowCount: sheet.rowCount || null,
    columnCount: sheet.columnCount || null
  }
}

function inferRisk(action) {
  if (['deleteRows', 'deleteColumns', 'clearRange'].includes(action.type)) return 'high'
  if (['insertRows', 'insertColumns', 'merge', 'breakApart', 'setFilter', 'createChart'].includes(action.type)) return 'medium'
  return 'low'
}
