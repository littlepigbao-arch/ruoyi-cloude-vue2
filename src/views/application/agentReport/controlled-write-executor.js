import { executeActions } from './univer-action-executor'
import { parseA1, toA1 } from './ai-action-schema'
import { prepareControlledContentClear } from './controlled-content-clear'
import { probeControlledFormatAction } from './controlled-format-foundation'

const plain = value => JSON.parse(JSON.stringify(value))
const scalar = value => value === null || ['string', 'boolean'].includes(typeof value) || (typeof value === 'number' && Number.isFinite(value))
const stable = value => JSON.stringify(sort(value))
function sort(value) {
  if (Array.isArray(value)) return value.map(sort)
  if (value && typeof value === 'object') return Object.keys(value).sort().reduce((out, key) => { out[key] = sort(value[key]); return out }, {})
  return value
}
const key = cell => cell.sheetId + ':' + cell.row + ':' + cell.column
const contentKeys = ['v', 'p', 'f', 'si', 't']
function normalizedCell(cell) {
  const copy = plain(cell || {})
  if (typeof copy.v === 'boolean' && copy.t === 3) copy.v = copy.v ? 1 : 0
  for (const name of ['v', 'p', 'f', 'si', 's', 'custom']) if (copy[name] == null) delete copy[name]
  if (!Object.prototype.hasOwnProperty.call(copy, 'v') && !copy.p && !copy.f && !copy.si) delete copy.t
  return copy
}
function normalizedSnapshot(snapshot) {
  const copy = plain(snapshot)
  for (const sheet of Object.values(copy.sheets || {})) {
    for (const [row, cells] of Object.entries(sheet.cellData || {})) {
      for (const [column, cell] of Object.entries(cells)) { const next = normalizedCell(cell); if (Object.keys(next).length) cells[column] = next; else delete cells[column] }
      if (!Object.keys(cells).length) delete sheet.cellData[row]
    }
    if (sheet.cellData && !Object.keys(sheet.cellData).length) delete sheet.cellData
  }
  return copy
}
export const sameControlledSnapshot = (left, right) => stable(normalizedSnapshot(left)) === stable(normalizedSnapshot(right))
const inertFormatResources = Object.freeze({
  SHEET_RANGE_PROTECTION_PLUGIN: '',
  SHEET_AuthzIoMockService_PLUGIN: '{}',
  SHEET_WORKSHEET_PROTECTION_PLUGIN: '{}',
  SHEET_WORKSHEET_PROTECTION_POINT_PLUGIN: '{}',
  SHEET_CONDITIONAL_FORMATTING_PLUGIN: '',
  SHEET_DRAWING_PLUGIN: '{}',
  SHEET_DEFINED_NAME_PLUGIN: '{}',
  SHEET_RANGE_THEME_MODEL_PLUGIN: '{}',
  SHEET_FILTER_PLUGIN: '{}'
})
/** Loading a saved workbook may add empty SDK plugin resources; retain every other resource exactly. */
function normalizedFormatSnapshot(snapshot) {
  const copy = normalizedSnapshot(snapshot)
  if (copy.resources === undefined) return copy
  if (!Array.isArray(copy.resources)) return null
  const seen = new Set(), kept = []
  for (const resource of copy.resources) {
    if (!resource || typeof resource !== 'object' || Array.isArray(resource) || typeof resource.name !== 'string' || seen.has(resource.name)) return null
    seen.add(resource.name)
    if (Object.keys(resource).length === 2 && Object.prototype.hasOwnProperty.call(inertFormatResources, resource.name) && resource.data === inertFormatResources[resource.name]) continue
    kept.push(resource)
  }
  if (kept.length) copy.resources = kept
  else delete copy.resources
  return copy
}
/** Univer stores a newly set background under a registry id; the trusted B5 patch stores it inline. */
export function sameControlledFormatSnapshot(local, server, patch) {
  if (!patch || typeof patch.sheetId !== 'string') return false
  if (patch.type !== 'setStyle') {
    const row = patch.type === 'setRowHeight', column = patch.type === 'setColumnWidth'
    if ((!row && !column) || !Number.isInteger(patch.index) || !Number.isInteger(patch.value)) return false
    const table = row ? 'rowData' : 'columnData', field = row ? 'h' : 'w'
    const left = normalizedFormatSnapshot(local), right = normalizedFormatSnapshot(server)
    return !!left && !!right && local?.sheets?.[patch.sheetId]?.[table]?.[patch.index]?.[field] === patch.value &&
      server?.sheets?.[patch.sheetId]?.[table]?.[patch.index]?.[field] === patch.value && stable(left) === stable(right)
  }
  if (!Number.isInteger(patch.row) || !Number.isInteger(patch.column) || typeof patch.value !== 'string' || !/^#[0-9a-fA-F]{6}$/.test(patch.value)) return false
  const canonical = snapshot => {
    const copy = normalizedFormatSnapshot(snapshot)
    if (!copy) return null
    const target = copy.sheets?.[patch.sheetId]?.cellData?.[patch.row]?.[patch.column]
    if (!target) return null
    if (typeof target.s === 'string') {
      const ref = target.s
      if (!/^[A-Za-z0-9_-]{1,64}$/.test(ref) || !copy.styles || !copy.styles[ref] || typeof copy.styles[ref] !== 'object' || Array.isArray(copy.styles[ref])) return null
      for (const sheet of Object.values(copy.sheets || {})) for (const row of Object.values(sheet.cellData || {})) for (const cell of Object.values(row)) {
        if (cell !== target && cell?.s === ref) return null
      }
      target.s = plain(copy.styles[ref])
      delete copy.styles[ref]
    }
    if (target.s?.bg?.rgb?.toLowerCase() !== patch.value.toLowerCase()) return null
    if (copy.styles && !Object.keys(copy.styles).length) delete copy.styles
    return copy
  }
  const left = canonical(local), right = canonical(server)
  return !!left && !!right && stable(left) === stable(right)
}
/** Rebuild the only permitted format change from an unchanged second-page baseline. */
export function sameControlledFormatTransition(before, after, patch) {
  if (!before || !after || !patch || typeof patch.sheetId !== 'string' || !Object.prototype.hasOwnProperty.call(patch, 'before')) return false
  const expected = normalizedFormatSnapshot(before), sheet = expected?.sheets?.[patch.sheetId]
  if (!sheet) return false
  if (patch.type === 'setStyle') {
    if (!Number.isInteger(patch.row) || patch.row < 0 || !Number.isInteger(patch.column) || patch.column < 0 || typeof patch.value !== 'string' || !/^#[0-9a-fA-F]{6}$/.test(patch.value)) return false
    const prior = sheet.cellData?.[patch.row]?.[patch.column]
    if (typeof prior?.s === 'string') {
      const ref = prior.s, registered = expected.styles?.[ref]
      if (!/^[A-Za-z0-9_-]{1,64}$/.test(ref) || !registered || typeof registered !== 'object' || Array.isArray(registered)) return false
      for (const candidate of Object.values(expected.sheets)) for (const row of Object.values(candidate.cellData || {})) for (const cell of Object.values(row)) {
        if (cell !== prior && cell?.s === ref) return false
      }
      prior.s = plain(registered)
      delete expected.styles[ref]
    }
    if (stable(prior ?? null) !== stable(patch.before ?? null)) return false
    if (!sheet.cellData) sheet.cellData = {}
    if (!sheet.cellData[patch.row]) sheet.cellData[patch.row] = {}
    const cell = sheet.cellData[patch.row][patch.column] || (sheet.cellData[patch.row][patch.column] = {})
    if (cell.s == null) cell.s = {}
    if (typeof cell.s !== 'object' || Array.isArray(cell.s)) return false
    if (cell.s.bg == null) cell.s.bg = {}
    if (typeof cell.s.bg !== 'object' || Array.isArray(cell.s.bg)) return false
    cell.s.bg.rgb = patch.value.toLowerCase()
  } else {
    const row = patch.type === 'setRowHeight', column = patch.type === 'setColumnWidth'
    if ((!row && !column) || !Number.isInteger(patch.index) || patch.index < 0 || !Number.isInteger(patch.value)) return false
    const table = row ? 'rowData' : 'columnData', field = row ? 'h' : 'w'
    if (stable(sheet[table]?.[patch.index] ?? null) !== stable(patch.before ?? null)) return false
    if (!sheet[table]) sheet[table] = {}
    const metadata = sheet[table][patch.index] || (sheet[table][patch.index] = {})
    if (typeof metadata !== 'object' || Array.isArray(metadata)) return false
    metadata[field] = patch.value
    if (row) metadata.ia = 0
  }
  return sameControlledFormatSnapshot(expected, after, patch)
}
export const CONTROLLED_VALUE_ACTIONS = ['setCell', 'setValues', 'setCellForCell', 'clearRange']
export const CONTROLLED_FORMAT_ACTIONS = ['setStyle', 'setRowHeight', 'setColumnWidth']
export const CONTROLLED_FILTER_ACTIONS = ['setFilter', 'clearFilter']
export const CONTROLLED_FORMULA_ACTIONS = ['setFormula', 'sumToCell']
export function normalizeControlledFormulaAction(raw) {
  if (raw && raw.type === 'setFormula') return raw
  const allowed = ['schemaVersion', 'type', 'target', 'sourceRange', 'targetRange']
  if (!raw || raw.schemaVersion !== '2.0' || raw.type !== 'sumToCell' || Object.keys(raw).some(k => !allowed.includes(k)) ||
      !raw.target || Object.keys(raw.target).some(k => !['sheetId', 'range'].includes(k)) || typeof raw.target.sheetId !== 'string' || !raw.target.sheetId) throw new Error('B10 sumToCell 清单无效')
  const range = Object.prototype.hasOwnProperty.call(raw.target, 'range') ? raw.target.range : raw.targetRange
  const pattern = /^[A-Z]{1,3}[1-9][0-9]*(?::[A-Z]{1,3}[1-9][0-9]*)?$/
  if (typeof range !== 'string' || !pattern.test(range) || typeof raw.sourceRange !== 'string' || !pattern.test(raw.sourceRange) ||
      (raw.targetRange !== undefined && raw.targetRange !== range)) throw new Error('B10 sumToCell 来源或目标无效')
  const target = parseA1(range), source = parseA1(raw.sourceRange)
  const sourceParts = raw.sourceRange.split(':'), first = parseA1(sourceParts[0]), last = parseA1(sourceParts[1] || sourceParts[0])
  if (!first || !last || first.startRow > last.startRow || first.startColumn > last.startColumn) throw new Error('B10 sumToCell 来源范围次序无效')
  if (!target || !source || target.startRow !== target.endRow || target.startColumn !== target.endColumn ||
      source.endRow < source.startRow || source.endColumn < source.startColumn || source.endRow >= 1000000 || source.endColumn >= 16384 ||
      (source.endRow - source.startRow + 1) * (source.endColumn - source.startColumn + 1) > 1000) throw new Error('B10 sumToCell 范围无效或超限')
  if (target.startRow >= source.startRow && target.startRow <= source.endRow && target.startColumn >= source.startColumn && target.startColumn <= source.endColumn) throw new Error('B10 sumToCell 来源不能包含目标')
  const sourceRange = raw.sourceRange.includes(':') ? raw.sourceRange : raw.sourceRange + ':' + raw.sourceRange
  return { schemaVersion: '2.0', type: 'setFormula', target: { sheetId: raw.target.sheetId, range }, formula: '=SUM(' + sourceRange + ')' }
}
export const CONTROLLED_MERGE_ACTIONS = ['merge', 'breakApart']
export const CONTROLLED_STRUCTURE_ACTIONS = ['insertRows', 'deleteRows', 'insertColumns', 'deleteColumns']
export const CONTROLLED_SHEET_ACTIONS = ['createSheet', 'renameSheet', 'copySheet', 'deleteSheet']
function normalizedStructureSnapshot(snapshot) {
  const copy = normalizedFormatSnapshot(snapshot)
  if (!copy) return null
  for (const sheet of Object.values(copy.sheets || {})) {
    if (sheet.cellData) for (const [row, cells] of Object.entries(sheet.cellData)) if (!Object.keys(cells).length) delete sheet.cellData[row]
    for (const key of ['cellData', 'rowData', 'columnData']) if (sheet[key] && !Object.keys(sheet[key]).length) delete sheet[key]
    if (Array.isArray(sheet.mergeData) && !sheet.mergeData.length) delete sheet.mergeData
  }
  return copy
}
function sameStructureSheet(actual, expected) {
  if (!actual || !expected || typeof actual !== 'object' || typeof expected !== 'object') return false
  const left = normalizedStructureSnapshot({ sheets: { s: actual } }), right = normalizedStructureSnapshot({ sheets: { s: expected } })
  return !!left && !!right && stable(left) === stable(right)
}
export function sameControlledStructureSnapshot(local, server, patch) {
  if (patch?.beforeWorkbook) {
    const expected = normalizedStructureSnapshot(patch.afterWorkbook), left = normalizedStructureSnapshot(local), right = normalizedStructureSnapshot(server)
    // The installed SDK retains the initial activeSheet metadata after deleting it.
    if (patch.type === 'deleteSheet') for (const snapshot of [left, right]) {
      if (snapshot?.activeSheet === patch.sheetId && !snapshot.sheets?.[patch.sheetId] && snapshot.sheets?.[patch.fallbackSheetId]) snapshot.activeSheet = patch.fallbackSheetId
    }
    return !!expected && !!left && !!right && stable(left) === stable(expected) && stable(right) === stable(expected)
  }
  if (!patch || !patch.sheetId || !patch.afterSheet || !sameStructureSheet(local?.sheets?.[patch.sheetId], patch.afterSheet) || !sameStructureSheet(server?.sheets?.[patch.sheetId], patch.afterSheet)) return false
  const left = normalizedStructureSnapshot(local), right = normalizedStructureSnapshot(server)
  return !!left && !!right && stable(left) === stable(right)
}
export function sameControlledStructureTransition(before, after, patch) {
  if (patch?.beforeWorkbook) {
    const baseline = normalizedStructureSnapshot(patch.beforeWorkbook), actualBefore = normalizedStructureSnapshot(before)
    return !!baseline && !!actualBefore && stable(baseline) === stable(actualBefore) && sameControlledStructureSnapshot(patch.afterWorkbook, after, patch)
  }
  if (!patch || !patch.sheetId || !patch.beforeSheet || !patch.afterSheet || !sameStructureSheet(before?.sheets?.[patch.sheetId], patch.beforeSheet)) return false
  const expected = plain(before)
  expected.sheets[patch.sheetId] = plain(patch.afterSheet)
  return sameControlledStructureSnapshot(expected, after, patch)
}
const mergeGeometry = merge => ({ startRow: merge.startRow, startColumn: merge.startColumn, endRow: merge.endRow, endColumn: merge.endColumn })
function mergeGeometries(snapshot, sheetId) {
  const list = snapshot?.sheets?.[sheetId]?.mergeData || []
  if (!Array.isArray(list)) throw new Error('B8b 合并资源无效')
  return list.map(mergeGeometry)
}
function normalizedMergeSnapshot(snapshot) {
  const copy = normalizedFormatSnapshot(snapshot)
  if (!copy) return null
  for (const sheet of Object.values(copy.sheets || {})) sheet.mergeData = (sheet.mergeData || []).map(mergeGeometry)
  return copy
}
export function sameControlledMergeSnapshot(local, server, patch) {
  try {
    if (!patch || !patch.sheetId || !Array.isArray(patch.afterMerges) || !Array.isArray(patch.removedCells)) return false
    const expected = patch.afterMerges.map(mergeGeometry)
    if (stable(mergeGeometries(local, patch.sheetId)) !== stable(expected) || stable(mergeGeometries(server, patch.sheetId)) !== stable(expected)) return false
    const left = normalizedMergeSnapshot(local), right = normalizedMergeSnapshot(server)
    return !!left && !!right && stable(left) === stable(right)
  } catch (_) { return false }
}
export function sameControlledMergeTransition(before, after, patch) {
  try {
    if (!patch || !patch.sheetId || !Array.isArray(patch.beforeMerges) || !Array.isArray(patch.afterMerges) || !Array.isArray(patch.removedCells)) return false
    if (stable(mergeGeometries(before, patch.sheetId)) !== stable(patch.beforeMerges.map(mergeGeometry))) return false
    const expected = plain(before), sheet = expected.sheets?.[patch.sheetId]
    if (!sheet) return false
    sheet.mergeData = plain(patch.afterMerges)
    for (const cell of patch.removedCells) {
      const row = sheet.cellData?.[cell.row]
      if (!row || stable(row[cell.column]) !== stable(cell.before)) return false
      delete row[cell.column]
      if (!Object.keys(row).length) delete sheet.cellData[cell.row]
    }
    return sameControlledMergeSnapshot(expected, after, patch)
  } catch (_) { return false }
}
export function sameControlledFormulaSnapshot(local, server, patch) {
  if (!patch || typeof patch.sheetId !== 'string' || !Array.isArray(patch.cells) || !patch.cells.length) return false
  const left = normalizedFormatSnapshot(local), right = normalizedFormatSnapshot(server)
  if (!left || !right) return false
  for (const cell of patch.cells) {
    if (!Number.isInteger(cell.row) || !Number.isInteger(cell.column) || typeof cell.formula !== 'string' || !Number.isSafeInteger(cell.value)) return false
    for (const snapshot of [left, right]) {
      const actual = snapshot.sheets?.[patch.sheetId]?.cellData?.[cell.row]?.[cell.column]
      if (actual?.f !== cell.formula || actual?.v !== cell.value || actual?.t !== 2) return false
    }
  }
  return stable(left) === stable(right)
}
export function sameControlledFormulaTransition(before, after, patch) {
  try {
    const expected = plain(before), sheet = expected.sheets?.[patch.sheetId]
    if (!sheet || !Array.isArray(patch.cells) || !patch.cells.length) return false
    if (!sheet.cellData) sheet.cellData = {}
    for (const cell of patch.cells) {
      const prior = sheet.cellData?.[cell.row]?.[cell.column]
      if (prior && Object.keys(prior).length) return false
      if (!sheet.cellData[cell.row]) sheet.cellData[cell.row] = {}
      sheet.cellData[cell.row][cell.column] = { f: cell.formula, v: cell.value, t: 2 }
    }
    return sameControlledFormulaSnapshot(expected, after, patch)
  } catch (_) { return false }
}
function filterData(snapshot) {
  const entries = (snapshot.resources || []).filter(resource => resource?.name === 'SHEET_FILTER_PLUGIN')
  if (entries.length > 1 || (entries.length && typeof entries[0].data !== 'string')) throw new Error('B6 筛选资源重复或无效')
  const parsed = entries.length ? JSON.parse(entries[0].data) : {}
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) throw new Error('B6 筛选资源结构无效')
  return parsed
}
function normalizedFilterSnapshot(snapshot) {
  const copy = normalizedFormatSnapshot(snapshot)
  if (!copy) return null
  for (const resource of copy.resources || []) if (resource.name === 'SHEET_FILTER_PLUGIN') resource.data = stable(JSON.parse(resource.data))
  return copy
}
export function sameControlledFilterSnapshot(local, server, patch) {
  try {
    if (!patch || !patch.sheetId || stable(filterData(local)) !== stable(JSON.parse(patch.resourceData)) || stable(filterData(server)) !== stable(JSON.parse(patch.resourceData))) return false
    const left = normalizedFilterSnapshot(local), right = normalizedFilterSnapshot(server)
    return !!left && !!right && stable(left) === stable(right)
  } catch (_) { return false }
}
export function sameControlledFilterTransition(before, after, patch) {
  try {
    if (!patch || !patch.sheetId || stable(filterData(before)) !== stable(patch.beforeData ? JSON.parse(patch.beforeData) : {})) return false
    const expected = plain(before)
    if (!Array.isArray(expected.resources)) expected.resources = []
    const existing = expected.resources.find(resource => resource.name === 'SHEET_FILTER_PLUGIN')
    if (existing) existing.data = patch.resourceData
    else expected.resources.push({ name: 'SHEET_FILTER_PLUGIN', data: patch.resourceData })
    return sameControlledFilterSnapshot(expected, after, patch)
  } catch (_) { return false }
}
export const READ_DISPLAY_ACTIONS = ['getRangeValue', 'getRangeFormulas', 'getMerges', 'getSelection', 'getFilter', 'createChart']
export function classifyAiActions(actions, controlledActions = []) {
  if (!Array.isArray(actions) || actions.length > 30) throw new Error('AI_CAPABILITY_BLOCKED: 动作组无效或超过30项')
  if (actions.length === 1 && actions[0] && CONTROLLED_FORMAT_ACTIONS.includes(actions[0].type) && Array.isArray(controlledActions) && controlledActions.includes(actions[0].type)) {
    const action = actions[0], target = action.target
    if (action.schemaVersion !== '2.0' || !target || typeof target.sheetId !== 'string' || !target.sheetId.trim() ||
      (action.type === 'setStyle' ? typeof target.range !== 'string' || !target.range.trim() : Object.prototype.hasOwnProperty.call(target, 'range'))) throw new Error('AI_CAPABILITY_BLOCKED: B5 格式目标无效')
    return 'controlled-values'
  }
  if (actions.length === 1 && actions[0] && CONTROLLED_FILTER_ACTIONS.includes(actions[0].type) && Array.isArray(controlledActions) && controlledActions.includes(actions[0].type)) {
    const action = actions[0], target = action.target
    if (action.schemaVersion !== '2.0' || !target || typeof target.sheetId !== 'string' || !target.sheetId.trim() || typeof target.range !== 'string' || !parseA1(target.range)) throw new Error('AI_CAPABILITY_BLOCKED: B6 筛选目标无效')
    return 'controlled-values'
  }
  if (actions.length === 1 && actions[0] && CONTROLLED_FORMULA_ACTIONS.includes(actions[0].type) && Array.isArray(controlledActions) && controlledActions.includes(actions[0].type)) {
    const action = normalizeControlledFormulaAction(actions[0]), target = action.target
    if (action.schemaVersion !== '2.0' || !target || typeof target.sheetId !== 'string' || !target.sheetId.trim() || typeof target.range !== 'string' || !parseA1(target.range) || typeof action.formula !== 'string') throw new Error('AI_CAPABILITY_BLOCKED: B7 公式目标无效')
    return 'controlled-values'
  }
  if (actions.length === 1 && actions[0] && CONTROLLED_MERGE_ACTIONS.includes(actions[0].type) && Array.isArray(controlledActions) && controlledActions.includes(actions[0].type)) {
    const action = actions[0], target = action.target
    if (action.schemaVersion !== '2.0' || !target || typeof target.sheetId !== 'string' || !target.sheetId.trim() || typeof target.range !== 'string' || !parseA1(target.range)) throw new Error('AI_CAPABILITY_BLOCKED: B8b 合并目标无效')
    return 'controlled-values'
  }
  if (actions.length === 1 && actions[0] && CONTROLLED_STRUCTURE_ACTIONS.includes(actions[0].type) && Array.isArray(controlledActions) && controlledActions.includes(actions[0].type)) {
    const action = actions[0], target = action.target
    const axis = action.type.endsWith('Rows') ? 'row' : 'column', index = axis + (action.type.startsWith('insert') ? 'Index' : 'Position')
    if (action.schemaVersion !== '2.0' || !target || typeof target.sheetId !== 'string' || !target.sheetId.trim() || Object.prototype.hasOwnProperty.call(target, 'range') || !Number.isInteger(action[index]) || !Number.isInteger(action.count) || action[index] < 1 || action.count < 1 || action.count > 10) throw new Error('AI_CAPABILITY_BLOCKED: B8c 行列目标无效')
    return 'controlled-values'
  }
  if (actions.length === 1 && CONTROLLED_SHEET_ACTIONS.includes(actions[0]?.type) && Array.isArray(controlledActions) && controlledActions.includes(actions[0].type)) {
    const action = actions[0], target = action.target
    if (action.schemaVersion !== '2.0' || !target || typeof target.sheetId !== 'string' || !target.sheetId.trim() || Object.prototype.hasOwnProperty.call(target, 'range') || (action.type === 'deleteSheet' ? (action.force !== true || typeof action.fallbackSheetId !== 'string' || action.fallbackSheetId === target.sheetId) : (typeof action.newSheetName !== 'string' || !action.newSheetName.trim())) || (action.type === 'createSheet' && (!Number.isInteger(action.rowCount) || !Number.isInteger(action.columnCount)))) throw new Error('AI_CAPABILITY_BLOCKED: B8c Sheet 目标无效')
    return 'controlled-values'
  }
  let reads = false; let writes = false
  for (const action of actions) {
    if (action && CONTROLLED_VALUE_ACTIONS.includes(action.type)) {
      if (action.type === 'clearRange' && Object.prototype.hasOwnProperty.call(action, 'clearWhat') && action.clearWhat !== 'content') throw new Error('AI_CAPABILITY_BLOCKED: clearRange 仅开放content')
      if (action.type === 'setCellForCell' && (!scalar(action.value) || action.value === null || action.value === undefined || (typeof action.value === 'string' && action.value.trim().startsWith('=')))) throw new Error('AI_CAPABILITY_BLOCKED: setCellForCell 仅接受非null普通标量')
      writes = true
    }
    else if (action && READ_DISPLAY_ACTIONS.includes(action.type)) reads = true
    else throw new Error('AI_CAPABILITY_BLOCKED: 此动作尚未迁移至受控链路：' + (action && action.type))
  }
  if (reads && writes) throw new Error('AI_CAPABILITY_BLOCKED: 读取展示与写入混合动作组暂不执行，请拆分请求')
  return writes ? 'controlled-values' : 'read-display'
}

/** Both entrances retain one identity. An empty lookup is not evidence of failed creation. */
export async function resolveControlledCreation(kind, request, send) {
  let result
  try { result = await send('/runs/resolve-create', { kind, request }) } catch (_) {
    const error = new Error('RESULT_UNKNOWN: 创建身份核实暂不可用，保留原请求与保护状态')
    error.creationUncertain = true
    throw error
  }
  const run = result && result.run
  if (!result || result.creationVerified !== true || result.requestId !== request.requestId || !run || !run.runId || run.requestId !== request.requestId || String(run.workbookId) !== String(request.workbookId) || String(run.branchId) !== String(request.branchId)) {
    const error = new Error('RESULT_UNKNOWN: 创建身份尚未核实，保持保护；空查询或旧任务不能证明未创建')
    error.creationUncertain = true
    throw error
  }
  return run
}
export async function createOrResolveControlledRun(kind, request, send) {
  const endpoint = kind === 'CONTROLLED_WRITE' ? '/controlled-writes' : '/runs'
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const run = await send(endpoint, request)
      if (!run || !run.runId || run.requestId !== request.requestId || String(run.workbookId) !== String(request.workbookId) || String(run.branchId) !== String(request.branchId)) throw new Error('创建响应身份不符')
      return run
    } catch (e) { if (attempt === 0 && e.definitive) throw e }
  }
  try { return await resolveControlledCreation(kind, request, send) } catch (e) {
    const error = new Error('RESULT_UNKNOWN: 创建结果待核实，保持原请求身份与锁定；不得重新提交任务')
    error.creationUncertain = true
    throw error
  }
}

/** Shared receipt transport for both entry points. Never retries browser execution. */
export async function submitControlledCall(workspace, call, runId, request, isCurrent) {
  if (!isCurrent()) throw new Error('页面已切换，不执行旧调用')
  const identity = { callId: call.callId, claimToken: call.claimToken, argsHash: call.argsHash, fencingToken: call.fencingToken }
  if (call.effectId) identity.effectId = call.effectId
  const receipt = { ...identity, transportVersion: 2, ok: false }
  try { Object.assign(receipt, await workspace.applyControlled(call)) } catch (e) { receipt.error = e.message }
  const verifyFailureReply = run => {
    if (!receipt.ok && run && (run.status === 'COMPLETED' || run.persistenceStatus === 'SAVED' || run.lastControlledResult?.persistenceStatus === 'SAVED')) {
      const error = new Error('RESULT_UNKNOWN: 本地执行失败与服务器已保存状态冲突，保持保护并核实；不能重放')
      throw error
    }
    return run
  }
  if (!isCurrent()) throw new Error('页面已切换，执行结果待核实；不提交旧回执')
  for (let attempt = 0; attempt < 2; attempt++) {
    let acknowledged
    try { acknowledged = await request('/runs/' + runId + '/controlled-results', receipt) } catch (e) { continue /* identical receipt only */ }
    return verifyFailureReply(acknowledged)
  }
  let resolved
  try { resolved = await request('/runs/' + runId + '/controlled-uncertain', identity) } catch (e) { throw new Error('RESULT_UNKNOWN: 保存回执与核实请求未返回，保持保护；不能再次执行') }
  verifyFailureReply(resolved)
  if (resolved.lastControlledResult && resolved.lastControlledResult.callId === call.callId && resolved.lastControlledResult.persistenceStatus === 'SAVED') return resolved
  const error = new Error('RESULT_UNKNOWN: 保存结果未核实，保持保护，请核实或取消；不能重放')
  error.run = resolved
  throw error
}
function masked(snapshot, cells) {
  const copy = plain(snapshot)
  for (const cell of cells) {
    const data = copy.sheets[cell.sheetId].cellData || {}
    const row = data[cell.row]
    if (row && row[cell.column]) {
      const metadata = normalizedCell(row[cell.column])
      for (const field of contentKeys) delete metadata[field]
      if (Object.keys(metadata).length) row[cell.column] = metadata; else delete row[cell.column]
    }
    if (row && !Object.keys(row).length) delete data[cell.row]
    if (!Object.keys(data).length) delete copy.sheets[cell.sheetId].cellData
  }
  return copy
}

/** Thin transport adapter: all mutations go through the existing executeActions. Never recreates a unit. */
export async function executeControlledWrite(univerAPI, workbook, manifest) {
  if (manifest && manifest.structurePatch) {
    if (manifest.structurePatch.type === 'deleteSheet') {
      if (manifest.transportVersion !== 2 || !Array.isArray(manifest.actions) || manifest.actions.length !== 1 || manifest.cells || manifest.patch || manifest.filterPatch || manifest.formulaPatch || manifest.mergePatch || stable(manifest.action) !== stable(manifest.actions[0])) throw new Error('B8c Sheet 删除清单无效')
      const patch = manifest.structurePatch, action = manifest.actions[0], before = plain(workbook.save())
      if (action.type !== 'deleteSheet' || action.schemaVersion !== '2.0' || action.force !== true || patch.sheetId !== action.target?.sheetId || patch.fallbackSheetId !== action.fallbackSheetId || !workbook.getSheetBySheetId(patch.sheetId) || !workbook.getSheetBySheetId(patch.fallbackSheetId) || patch.sheetId === patch.fallbackSheetId || !sameControlledStructureTransition(before, patch.afterWorkbook, patch)) throw new Error('B8c Sheet 删除目标、确认或前置事实不一致')
      workbook.setActiveSheet(patch.fallbackSheetId)
      if (workbook.getActiveSheet()?.getSheetId() !== patch.fallbackSheetId || !workbook.deleteSheet(patch.sheetId)) throw new Error('RESULT_UNKNOWN: B8c Sheet 删除或保留表切换失败')
      const after = plain(workbook.save())
      if (workbook.getSheetBySheetId(patch.sheetId) || workbook.getActiveSheet()?.getSheetId() !== patch.fallbackSheetId || !sameControlledStructureTransition(before, after, patch)) throw new Error('RESULT_UNKNOWN: B8c Sheet 删除完整回读不一致')
      const readback = { sheetId: patch.sheetId, type: 'deleteSheet', deleted: true, activeSheetId: patch.fallbackSheetId, workbookDigest: patch.afterDigest }
      if (stable(readback) !== stable(manifest.structureReadback)) throw new Error('RESULT_UNKNOWN: B8c Sheet 删除回读不一致')
      return { ok: true, structureReadback: readback }
    }
    if (manifest.structurePatch.type === 'renameSheet') {
      if (manifest.transportVersion !== 2 || !Array.isArray(manifest.actions) || manifest.actions.length !== 1 || manifest.cells || manifest.patch || manifest.filterPatch || manifest.formulaPatch || manifest.mergePatch || stable(manifest.action) !== stable(manifest.actions[0])) throw new Error('B8c Sheet 重命名清单无效')
      const patch = manifest.structurePatch, action = manifest.actions[0], before = plain(workbook.save()), sheet = workbook.getSheetBySheetId(patch.sheetId)
      if (!sheet || action.type !== 'renameSheet' || action.schemaVersion !== '2.0' || patch.sheetId !== action.target?.sheetId || patch.newSheetName !== action.newSheetName || !sameControlledStructureTransition(before, patch.afterWorkbook, patch)) throw new Error('B8c Sheet 重命名目标或前置工作簿不一致')
      sheet.setName(patch.newSheetName)
      const after = plain(workbook.save())
      if (!sameControlledStructureTransition(before, after, patch)) throw new Error('RESULT_UNKNOWN: B8c Sheet 重命名完整工作簿回读不一致')
      const readback = { sheetId: patch.sheetId, type: 'renameSheet', newSheetName: patch.newSheetName, workbookDigest: patch.afterDigest }
      if (stable(readback) !== stable(manifest.structureReadback)) throw new Error('RESULT_UNKNOWN: B8c Sheet 重命名回读不一致')
      return { ok: true, structureReadback: readback }
    }
    if (['createSheet', 'copySheet'].includes(manifest.structurePatch.type)) {
      if (manifest.transportVersion !== 2 || !Array.isArray(manifest.actions) || manifest.actions.length !== 1 || manifest.cells || manifest.patch || manifest.filterPatch || manifest.formulaPatch || manifest.mergePatch || stable(manifest.action) !== stable(manifest.actions[0])) throw new Error('B8c Sheet 清单无效')
      const patch = manifest.structurePatch, action = manifest.actions[0], before = plain(workbook.save())
      const copy = action.type === 'copySheet'
      if (action.type !== patch.type || action.schemaVersion !== '2.0' || patch.sheetId !== action.target?.sheetId || patch.newSheetName !== action.newSheetName || (!copy && (patch.rowCount !== action.rowCount || patch.columnCount !== action.columnCount)) || !patch.newSheetId || workbook.getSheetBySheetId(patch.newSheetId) || !sameControlledStructureTransition(before, patch.afterWorkbook, patch)) throw new Error('B8c 新 Sheet 目标或前置工作簿不一致')
      if (manifest.resultSheet === true && (!manifest.resultSource || !patch.resultData || patch.type !== 'createSheet')) throw new Error('B9 新 Sheet 来源清单无效')
      const data = (copy || manifest.resultSheet === true) ? plain(patch.afterWorkbook.sheets[patch.newSheetId]) : { id: patch.newSheetId, rowCount: patch.rowCount, columnCount: patch.columnCount }
      const created = workbook.insertSheet(patch.newSheetName, { index: patch.index, sheet: data })
      if (!created || created.getSheetId() !== patch.newSheetId) throw new Error('RESULT_UNKNOWN: B8c 新 Sheet ID 不一致')
      const after = plain(workbook.save())
      if (!sameControlledStructureTransition(before, after, patch)) throw new Error('RESULT_UNKNOWN: B8c 新 Sheet 完整工作簿回读不一致')
      const readback = { sheetId: patch.newSheetId, type: patch.type, rowCount: patch.rowCount, columnCount: patch.columnCount, workbookDigest: patch.afterDigest }
      if (stable(readback) !== stable(manifest.structureReadback)) throw new Error('RESULT_UNKNOWN: B8c 新 Sheet 回读不一致')
      return { ok: true, structureReadback: readback }
    }
    if (manifest.transportVersion !== 2 || !Array.isArray(manifest.actions) || manifest.actions.length !== 1 || manifest.cells || manifest.patch || manifest.filterPatch || manifest.formulaPatch || manifest.mergePatch || stable(manifest.action) !== stable(manifest.actions[0])) throw new Error('B8c 受控清单无效')
    const patch = manifest.structurePatch, action = manifest.actions[0], sheet = workbook.getSheetBySheetId(patch.sheetId)
    const axis = action.type.endsWith('Rows') ? 'row' : 'column', indexField = axis + (action.type.startsWith('insert') ? 'Index' : 'Position')
    if (!sheet || !CONTROLLED_STRUCTURE_ACTIONS.includes(action.type) || patch.type !== action.type || patch.sheetId !== action.target?.sheetId || patch.count !== action.count || patch.index !== action[indexField] || patch.position !== (action.position || 'at') || (action.type.startsWith('delete') ? action.force !== true : action.force === true)) throw new Error('B8c 目标、位置、数量或删除确认不一致')
    const before = plain(workbook.save())
    if (!sameStructureSheet(before.sheets?.[patch.sheetId], patch.beforeSheet)) throw new Error('B8c 工作表结构已变化，保持保护')
    const outcome = executeActions(univerAPI, workbook, [action], { continueOnError: false })
    if (outcome.summary.failed || outcome.summary.skipped || outcome.summary.ok !== 1 || outcome.notes.length) throw new Error('RESULT_UNKNOWN: B8c 结构执行失败')
    const after = plain(workbook.save())
    if (!sameControlledStructureTransition(before, after, patch)) throw new Error('RESULT_UNKNOWN: B8c 结构影响区或邻区不一致')
    const readback = { sheetId: patch.sheetId, type: patch.type, rowCount: sheet.getMaxRows(), columnCount: sheet.getMaxColumns(), sheetDigest: patch.afterDigest }
    if (stable(readback) !== stable(manifest.structureReadback)) throw new Error('RESULT_UNKNOWN: B8c 结构回读不一致')
    return { ok: true, structureReadback: readback }
  }
  if (manifest && manifest.mergePatch) {
    if (manifest.transportVersion !== 2 || !Array.isArray(manifest.actions) || manifest.actions.length !== 1 || manifest.cells || manifest.patch || manifest.filterPatch || manifest.formulaPatch || stable(manifest.action) !== stable(manifest.actions[0])) throw new Error('B8b 受控清单无效')
    const patch = manifest.mergePatch, action = manifest.actions[0]
    if (!CONTROLLED_MERGE_ACTIONS.includes(action.type) || patch.type !== action.type || patch.sheetId !== action.target?.sheetId || patch.range !== action.target?.range || patch.force !== (action.force === true)) throw new Error('B8b 目标或授权不一致')
    const sheet = workbook.getSheetBySheetId(patch.sheetId), range = parseA1(patch.range)
    if (!sheet || !range || range.endRow >= sheet.getMaxRows() || range.endColumn >= sheet.getMaxColumns()) throw new Error('B8b 合并范围已变化')
    const before = plain(workbook.save())
    if (stable(mergeGeometries(before, patch.sheetId)) !== stable(patch.beforeMerges.map(mergeGeometry))) throw new Error('B8b 合并资源已变化，保持保护')
    if (patch.type === 'merge') {
      const removed = new Map((patch.removedCells || []).map(cell => [cell.row + ':' + cell.column, cell]))
      for (let row = range.startRow; row <= range.endRow; row++) for (let column = range.startColumn; column <= range.endColumn; column++) {
        if (row === range.startRow && column === range.startColumn) continue
        const cell = before.sheets?.[patch.sheetId]?.cellData?.[row]?.[column], recorded = removed.get(row + ':' + column)
        if (Object.keys(cell || {}).length && (!patch.force || !recorded || stable(cell) !== stable(recorded.before))) throw new Error('B8b 非主格内容未获授权，保持保护')
        if (!Object.keys(cell || {}).length && recorded) throw new Error('B8b 内容补丁与当前单元格不一致')
      }
    }
    for (const cell of patch.removedCells || []) if (stable(before.sheets?.[patch.sheetId]?.cellData?.[cell.row]?.[cell.column]) !== stable(cell.before)) throw new Error('B8b 非主格内容已变化，保持保护')
    const outcome = executeActions(univerAPI, workbook, [action], { continueOnError: false })
    if (outcome.summary.failed || outcome.summary.skipped || outcome.summary.ok !== 1 || outcome.notes.length) throw new Error('RESULT_UNKNOWN: B8b 合并执行失败')
    const observed = { sheetId: patch.sheetId, range: patch.range, merges: (sheet.getMergeData?.() || []).map(merge => mergeGeometry(merge._range || merge)) }
    const after = plain(workbook.save())
    if (!sameControlledMergeTransition(before, after, patch) || stable(observed) !== stable(manifest.mergeReadback)) throw new Error('RESULT_UNKNOWN: B8b 合并影响区或回读不一致')
    return { ok: true, mergeReadback: observed }
  }
  if (manifest && manifest.formulaPatch) {
    if (manifest.transportVersion !== 2 || !Array.isArray(manifest.actions) || manifest.actions.length !== 1 || manifest.cells || manifest.patch || manifest.filterPatch || stable(manifest.action) !== stable(manifest.actions[0])) throw new Error('B7 受控清单无效')
    const patch = manifest.formulaPatch, action = normalizeControlledFormulaAction(manifest.actions[0])
    if (action.type !== 'setFormula' || patch.sheetId !== action.target?.sheetId || patch.range !== action.target?.range) throw new Error('B7 公式目标不一致')
    const sheet = workbook.getSheetBySheetId(patch.sheetId), target = parseA1(patch.range)
    if (!sheet || !target || target.endRow >= sheet.getMaxRows() || target.endColumn >= sheet.getMaxColumns()) throw new Error('B7 公式范围已变化')
    const before = plain(workbook.save())
    for (const cell of patch.cells || []) if (Object.keys(before.sheets?.[patch.sheetId]?.cellData?.[cell.row]?.[cell.column] || {}).length) throw new Error('B7 公式目标不再为空')
    const outcome = executeActions(univerAPI, workbook, [action], { continueOnError: false })
    if (outcome.summary.failed || outcome.summary.skipped || outcome.summary.ok !== 1 || outcome.notes.length) throw new Error('RESULT_UNKNOWN: B7 公式执行失败')
    const observed = []
    for (const cell of patch.cells || []) {
      const range = sheet.getRange(toA1(cell.row, cell.column))
      let formula, value, stored
      for (let attempt = 0; attempt < 40; attempt++) {
        formula = range.getFormulas?.()?.[0]?.[0]; value = range.getValues?.()?.[0]?.[0]
        stored = workbook.save().sheets?.[patch.sheetId]?.cellData?.[cell.row]?.[cell.column]
        if (formula === cell.formula && value === cell.value && stored?.f === cell.formula && stored?.v === cell.value && stored?.t === 2) break
        await new Promise(resolve => setTimeout(resolve, 50))
      }
      if (formula !== cell.formula || value !== cell.value || stored?.f !== cell.formula || stored?.v !== cell.value || stored?.t !== 2) throw new Error('RESULT_UNKNOWN: B7 公式计算或缓存回读不符')
      observed.push({ formula, value, row: cell.row, column: cell.column })
    }
    const after = plain(workbook.save())
    if (!sameControlledFormulaTransition(before, after, patch) || stable(observed) !== stable(manifest.formulaReadback)) throw new Error('RESULT_UNKNOWN: B7 公式影响区或回读不一致')
    return { ok: true, formulaReadback: observed }
  }
  if (manifest && manifest.filterPatch) {
    if (manifest.transportVersion !== 2 || !Array.isArray(manifest.actions) || manifest.actions.length !== 1 || manifest.cells || manifest.patch || stable(manifest.action) !== stable(manifest.actions[0])) throw new Error('B6 受控清单无效')
    const patch = manifest.filterPatch, action = manifest.actions[0]
    if (!CONTROLLED_FILTER_ACTIONS.includes(action.type) || patch.type !== action.type || patch.sheetId !== action.target?.sheetId || patch.range !== action.target?.range) throw new Error('B6 受控目标不一致')
    const before = plain(workbook.save())
    if (stable(filterData(before)) !== stable(patch.beforeData ? JSON.parse(patch.beforeData) : {})) throw new Error('B6 筛选资源已变化，保持保护')
    const sheet = workbook.getSheetBySheetId(patch.sheetId), range = parseA1(patch.range)
    if (!sheet || !range || range.endRow >= sheet.getMaxRows() || range.endColumn >= sheet.getMaxColumns()) throw new Error('B6 筛选范围已变化')
    const outcome = executeActions(univerAPI, workbook, [action], { continueOnError: false })
    if (outcome.summary.failed || outcome.summary.skipped || outcome.summary.ok !== 1) throw new Error('RESULT_UNKNOWN: B6 筛选执行失败')
    const filter = typeof sheet.getFilter === 'function' ? sheet.getFilter() : null
    const expected = patch.filter
    if (!!filter !== !!expected) throw new Error('RESULT_UNKNOWN: B6 Filter 实际状态不符')
    if (filter) {
      if (filter.getRange?.().getA1Notation?.() !== patch.range) throw new Error('RESULT_UNKNOWN: B6 Filter 范围回读不符')
      const columns = expected.filterColumns || []
      for (const criterion of columns) if (stable(filter.getColumnFilterCriteria?.(range.startColumn + criterion.colId)) !== stable(criterion)) throw new Error('RESULT_UNKNOWN: B6 筛选条件回读不符')
      if (stable(filter.getFilteredOutRows?.()) !== stable(expected.cachedFilteredOut)) throw new Error('RESULT_UNKNOWN: B6 可见结果回读不符')
    }
    const after = plain(workbook.save())
    if (!sameControlledFilterTransition(before, after, patch)) throw new Error('RESULT_UNKNOWN: B6 目标外资源或内容发生变化')
    const readback = { sheetId: patch.sheetId, range: patch.range, hasFilter: !!filter, filter: expected }
    if (stable(readback) !== stable(manifest.filterReadback)) throw new Error('RESULT_UNKNOWN: B6 回读与可信清单不符')
    return { ok: true, filterReadback: readback }
  }
  if (manifest && manifest.patch) {
    if (manifest.transportVersion !== 2 || !Array.isArray(manifest.actions) || manifest.actions.length !== 1 || manifest.cells || stable(manifest.action) !== stable(manifest.actions[0])) throw new Error('B5受控清单无效')
    const patch = manifest.patch, action = manifest.actions[0]
    if (patch.type !== action.type || patch.sheetId !== action.target?.sheetId) throw new Error('B5受控目标不一致')
    const before = plain(workbook.save()), sheet = before.sheets?.[patch.sheetId]
    if (!sheet) throw new Error('B5目标 Sheet 不存在')
    const actualBefore = patch.type === 'setStyle' ? sheet.cellData?.[patch.row]?.[patch.column] : sheet[patch.type === 'setRowHeight' ? 'rowData' : 'columnData']?.[patch.index]
    if (stable(actualBefore ?? null) !== stable(patch.before ?? null)) throw new Error('B5目标格式已变化，保持保护')
    const proof = probeControlledFormatAction(univerAPI, workbook, action)
    const readback = proof.formatReadback
    if (!proof.verified || readback.sheetId !== patch.sheetId || stable(readback.value) !== stable(patch.value) || (patch.type === 'setStyle' ? readback.row !== patch.row || readback.column !== patch.column : readback.index !== patch.index)) throw new Error('RESULT_UNKNOWN: B5回读与服务器可信清单不符')
    return { ok: true, formatReadback: readback }
  }
  if (!manifest || manifest.transportVersion !== 2 || !Array.isArray(manifest.actions) || !manifest.actions.length || manifest.actions.length > 30 || !Array.isArray(manifest.cells) || !manifest.cells.length || manifest.cells.length > 1000) throw new Error('受控清单无效')
  const expected = new Map()
  for (const cell of manifest.cells) {
    if (!Number.isInteger(cell.row) || !Number.isInteger(cell.column) || cell.row < 0 || cell.column < 0 || !scalar(cell.value) || expected.has(key(cell))) throw new Error('受控单元格无效或重复')
    expected.set(key(cell), cell)
  }
  const actionCells = new Set()
  for (const action of manifest.actions) {
    if (action.schemaVersion !== '2.0' || !action.target || !action.target.sheetId || !CONTROLLED_VALUE_ACTIONS.includes(action.type)) throw new Error('受控动作类型或目标无效')
    classifyAiActions([action])
    const sheet = workbook.getSheetBySheetId(action.target.sheetId)
    if (!sheet) throw new Error('目标 Sheet 不存在')
    const r = parseA1(action.target.range)
    if (!r || r.endRow >= sheet.getMaxRows() || r.endColumn >= sheet.getMaxColumns()) throw new Error('受控范围超出 Sheet，不允许截断')
    if (['setCell', 'setCellForCell'].includes(action.type) && (r.startRow !== r.endRow || r.startColumn !== r.endColumn)) throw new Error('标量赋值只能写单格')
    if (action.type === 'clearRange' && action.clearWhat !== 'content') throw new Error('受控清除仅支持content')
    for (let row = r.startRow; row <= r.endRow; row++) for (let column = r.startColumn; column <= r.endColumn; column++) {
      const id = key({ sheetId: action.target.sheetId, row, column })
      const value = action.type === 'clearRange' ? null : action.type === 'setValues' ? action.values?.[row - r.startRow]?.[column - r.startColumn] : action.value
      if (!scalar(value) || (typeof value === 'string' && value.trim().startsWith('=')) || actionCells.has(id) || !expected.has(id) || stable(expected.get(id).value) !== stable(value)) throw new Error('动作与授权单元格不一致')
      actionCells.add(id)
      if ((action.type === 'clearRange') !== (expected.get(id).operation === 'clear-content')) throw new Error('清除意图与可信清单不一致')
    }
  }
  if (actionCells.size !== expected.size) throw new Error('受控清单范围不完整')
  const before = plain(workbook.save())
  for (const cell of manifest.cells) {
    const actual = before.sheets?.[cell.sheetId]?.cellData?.[cell.row]?.[cell.column]
    if (stable(normalizedCell(actual)) !== stable(normalizedCell(cell.before))) throw new Error('本地目标内容、富文本或属性已变化，保持保护并核实')
    if (cell.operation === 'clear-content') {
      if (actual?.custom != null || before.sheets[cell.sheetId].rowData?.[cell.row]?.hd || before.sheets[cell.sheetId].columnData?.[cell.column]?.hd) throw new Error('含custom或隐藏行列的清除尚未证明安全')
      if ((before.resources || []).some(r => /filter/i.test(r.name) && (typeof r.data !== 'string' || !['', '{}', 'null'].includes(r.data.trim())))) throw new Error('筛选状态下清除尚未证明安全')
    }
  }
  const controlledContentClears = new Map()
  for (const action of manifest.actions) if (action.type === 'clearRange') {
    const { sheetId, range } = action.target
    controlledContentClears.set(sheetId + ':' + range, prepareControlledContentClear(workbook.getSheetBySheetId(sheetId), parseA1(range)))
  }
  const result = executeActions(univerAPI, workbook, manifest.actions, { continueOnError: false, controlledValues: true, controlledContentClears })
  if (result.summary.failed || result.summary.skipped || result.summary.ok !== manifest.actions.length || result.notes.length) throw new Error('受控执行失败或未完整执行，不能报告完成')
  const after = plain(workbook.save())
  const cells = []
  for (const expectedCell of manifest.cells) {
    const range = workbook.getSheetBySheetId(expectedCell.sheetId).getRange(toA1(expectedCell.row, expectedCell.column))
    if (typeof range.getFormulas !== 'function' || typeof range.getValues !== 'function') throw new Error('当前版本无法独立回读单元格')
    const values = range.getValues()
    const formulas = range.getFormulas()
    const singleCell = grid => Array.isArray(grid) && grid.length === 1 && Array.isArray(grid[0]) && grid[0].length === 1
    // Facade returns [[null]] for an empty cell and [['']] for no formula.
    // A missing/partial response is unknown, never evidence that a clear succeeded.
    if (!singleCell(values) || !scalar(values[0][0]) || !singleCell(formulas) || typeof formulas[0][0] !== 'string') throw new Error('实际回读结构不完整，保持保护并核实')
    const raw = values[0][0]
    const actualCell = after.sheets?.[expectedCell.sheetId]?.cellData?.[expectedCell.row]?.[expectedCell.column]
    // Values alone do not prove scalar semantics: "001" tagged as a number,
    // or false tagged as numeric, disagrees with the server's trusted patch.
    const expectedType = typeof expectedCell.value === 'number' ? 2 : typeof expectedCell.value === 'boolean' ? 3 : 1
    if (expectedCell.value !== null && actualCell?.t !== expectedType) throw new Error('实际单元格类型与授权标量不符，保持保护并核实')
    // Univer stores booleans as 0/1 with t=3. Numeric zero (t=2) is distinct.
    const value = actualCell?.t === 3 && (raw === 0 || raw === 1) ? raw === 1 : raw
    const formula = formulas[0][0]
    if (formula || stable(value) !== stable(expectedCell.value)) throw new Error('实际回读与授权值不符')
    const stored = actualCell?.v ?? null
    const storedValue = actualCell?.t === 3 && (stored === 0 || stored === 1) ? stored === 1 : stored
    if (stable(storedValue) !== stable(value)) throw new Error('实际回读与工作簿内容不一致，拒绝保存')
    if (actualCell?.p != null || actualCell?.si || actualCell?.f) throw new Error('回读仍含富文本或未授权公式，拒绝保存')
    cells.push({ sheetId: expectedCell.sheetId, row: expectedCell.row, column: expectedCell.column, value, formula })
  }
  if (stable(masked(before, manifest.cells)) !== stable(masked(after, manifest.cells))) throw new Error('发现授权范围外的本地变化，保持保护并核实')
  return { ok: true, cells }
}
