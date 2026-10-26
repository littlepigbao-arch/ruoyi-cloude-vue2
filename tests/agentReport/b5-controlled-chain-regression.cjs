'use strict'
const assert = require('assert')
const fs = require('fs'), path = require('path'), vm = require('vm')
const root = path.resolve(__dirname, '../..')
const babel = require(path.join(root, 'node_modules/@babel/core'))
const plugin = require(path.join(root, 'node_modules/@babel/plugin-transform-modules-commonjs'))
const cache = new Map(), clone = v => JSON.parse(JSON.stringify(v))
function load(file) {
  if (cache.has(file)) return cache.get(file)
  const m = { exports: {} }
  vm.runInNewContext(babel.transformSync(fs.readFileSync(file, 'utf8'), { configFile: false, babelrc: false, plugins: [plugin] }).code,
    { module: m, exports: m.exports, console, performance, require: n => require('./clear-content-sdk-fixture.cjs').modules[n] || load(path.resolve(path.dirname(file), n + '.js')) })
  cache.set(file, m.exports); return m.exports
}
const api = load(path.join(root, 'src/views/application/agentReport/controlled-write-executor.js'))
function fixture(silent = false) {
  const snapshot = { sheets: { s: { name: 'Sheet', rowCount: 10, columnCount: 10, defaultRowHeight: 24, cellData: { 0: { 0: { v: 'keep', t: 1 } } }, rowData: { 0: { h: 24 } }, columnData: { 0: { w: 80 } } } } }
  const range = { getBackground: () => snapshot.sheets.s.cellData[0][0].s?.bg.rgb || '#ffffff', setBackgroundColor(color) { if (!silent) snapshot.sheets.s.cellData[0][0].s = { bg: { rgb: color } }; return this } }
  const sheet = { getSheetId: () => 's', getSheetName: () => 'Sheet', getMaxRows: () => 10, getMaxColumns: () => 10, getRange: () => range,
    getRowHeight: i => snapshot.sheets.s.rowData[i]?.h || 24, getColumnWidth: i => snapshot.sheets.s.columnData[i]?.w || 80,
    setRowHeight(i, value) { if (!silent) snapshot.sheets.s.rowData[i] = { h: value }; return this },
    setColumnWidth(i, value) { if (!silent) snapshot.sheets.s.columnData[i] = { w: value }; return this } }
  return { book: { getSheetBySheetId: () => sheet, getActiveSheet: () => sheet, save: () => clone(snapshot) }, snapshot }
}
const action = { schemaVersion: '2.0', type: 'setStyle', target: { sheetId: 's', range: 'A1' }, background: '#ff0000' }
function manifest() { return { transportVersion: 2, action, actions: [action], patch: { type: 'setStyle', sheetId: 's', row: 0, column: 0, value: '#ff0000', before: { v: 'keep', t: 1 } } } }
async function run() {
  assert.throws(() => api.classifyAiActions([action]), /AI_CAPABILITY_BLOCKED/)
  assert.strictEqual(api.classifyAiActions([action], ['setStyle']), 'controlled-values')
  for (const type of ['setRowHeight', 'setColumnWidth']) {
    const candidate = { schemaVersion: '2.0', type, target: { sheetId: 's' }, [type === 'setRowHeight' ? 'rowPosition' : 'columnPosition']: 1, [type === 'setRowHeight' ? 'height' : 'width']: 50 }
    assert.strictEqual(api.classifyAiActions([candidate], [type]), 'controlled-values')
    assert.throws(() => api.classifyAiActions([candidate], []), /AI_CAPABILITY_BLOCKED/)
  }
  assert.throws(() => api.classifyAiActions([action, action], ['setStyle']), /AI_CAPABILITY_BLOCKED/)
  assert.throws(() => api.classifyAiActions([{ ...action, target: {} }], ['setStyle']), /AI_CAPABILITY_BLOCKED/)
  assert.throws(() => api.classifyAiActions([action, { schemaVersion: '2.0', type: 'setCell', target: { sheetId: 's', range: 'B1' }, value: 1 }], ['setStyle']), /AI_CAPABILITY_BLOCKED/)
  const f = fixture(), proved = await api.executeControlledWrite({}, f.book, manifest())
  assert.strictEqual(proved.ok, true)
  assert.deepStrictEqual(JSON.parse(JSON.stringify(proved.formatReadback)), { sheetId: 's', row: 0, column: 0, value: '#ff0000' })
  await assert.rejects(() => api.executeControlledWrite({}, fixture(true).book, manifest()), /RESULT_UNKNOWN|回读/)
  const stale = manifest(); stale.patch.before.v = 'changed'
  await assert.rejects(() => api.executeControlledWrite({}, fixture().book, stale), /变化|无效/)
  for (const type of ['setRowHeight', 'setColumnWidth']) {
    const row = type === 'setRowHeight', field = row ? 'rowPosition' : 'columnPosition', size = row ? 'height' : 'width', value = row ? 50 : 120
    const a = { schemaVersion: '2.0', type, target: { sheetId: 's' }, [field]: 1, [size]: value }
    const m = { transportVersion: 2, action: a, actions: [a], patch: { type, sheetId: 's', index: 0, value, before: row ? { h: 24 } : { w: 80 } } }
    assert.strictEqual((await api.executeControlledWrite({}, fixture().book, m)).formatReadback.value, value)
    await assert.rejects(() => api.executeControlledWrite({}, fixture(true).book, m), /RESULT_UNKNOWN|回读/)
  }
  const patch = { type: 'setStyle', sheetId: 's', row: 0, column: 0, value: '#223344' }
  const server = { sheets: { s: { cellData: { 0: { 0: { v: 'keep', s: { bg: { rgb: '#223344' } } } }, 1: { 1: { v: 'outside' } } } } } }
  const local = { styles: { ref: { bg: { rgb: '#223344' } } }, sheets: { s: { cellData: { 0: { 0: { v: 'keep', s: 'ref' } }, 1: { 1: { v: 'outside' } } } } } }
  assert.strictEqual(api.sameControlledFormatSnapshot(local, server, patch), true)
  const changedOutside = clone(local); changedOutside.sheets.s.cellData[1][1].v = 'changed'
  assert.strictEqual(api.sameControlledFormatSnapshot(changedOutside, server, patch), false)
  const extraRegistry = clone(local); extraRegistry.styles.unapproved = { bg: { rgb: '#ffffff' } }
  assert.strictEqual(api.sameControlledFormatSnapshot(extraRegistry, server, patch), false)
  const missingRegistry = clone(local); delete missingRegistry.styles.ref
  assert.strictEqual(api.sameControlledFormatSnapshot(missingRegistry, server, patch), false)
  const sharedReference = clone(local); sharedReference.sheets.s.cellData[1][1].s = 'ref'
  assert.strictEqual(api.sameControlledFormatSnapshot(sharedReference, server, patch), false)
  assert.strictEqual(api.sameControlledFormatSnapshot(local, server, { ...patch, value: '#ffffff' }), false)
  const baselineResources = [
    { name: 'SHEET_RANGE_PROTECTION_PLUGIN', data: '' },
    { name: 'SHEET_AuthzIoMockService_PLUGIN', data: '{}' },
    { name: 'SHEET_WORKSHEET_PROTECTION_PLUGIN', data: '{}' },
    { name: 'SHEET_WORKSHEET_PROTECTION_POINT_PLUGIN', data: '{}' },
    { name: 'SHEET_DEFINED_NAME_PLUGIN', data: '{}' },
    { name: 'SHEET_RANGE_THEME_MODEL_PLUGIN', data: '{}' }
  ]
  const sdkResources = [...baselineResources,
    { name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN', data: '' },
    { name: 'SHEET_DRAWING_PLUGIN', data: '{}' },
    { name: 'SHEET_FILTER_PLUGIN', data: '{}' }]
  const serverWithResources = clone(server); serverWithResources.resources = baselineResources
  const localWithResources = clone(local); localWithResources.resources = sdkResources
  assert.strictEqual(api.sameControlledFormatSnapshot(localWithResources, serverWithResources, patch), true)
  const meaningfulFilter = clone(localWithResources); meaningfulFilter.resources.find(r => r.name === 'SHEET_FILTER_PLUGIN').data = '{"filters":[1]}'
  assert.strictEqual(api.sameControlledFormatSnapshot(meaningfulFilter, serverWithResources, patch), false)
  const duplicateResource = clone(localWithResources); duplicateResource.resources.push({ name: 'SHEET_FILTER_PLUGIN', data: '{}' })
  assert.strictEqual(api.sameControlledFormatSnapshot(duplicateResource, serverWithResources, patch), false)
  const rowSnapshot = { sheets: { s: { rowData: { 0: { h: 50, ia: 0 } } } } }
  assert.strictEqual(api.sameControlledFormatSnapshot(rowSnapshot, clone(rowSnapshot), { type: 'setRowHeight', sheetId: 's', index: 0, value: 50 }), true)
  assert.strictEqual(api.sameControlledFormatSnapshot(rowSnapshot, clone(rowSnapshot), { type: 'setRowHeight', sheetId: 's', index: 0, value: 60 }), false)
  const rowWithResources = clone(rowSnapshot); rowWithResources.resources = sdkResources
  const rowServerResources = clone(rowSnapshot); rowServerResources.resources = baselineResources
  assert.strictEqual(api.sameControlledFormatSnapshot(rowWithResources, rowServerResources, { type: 'setRowHeight', sheetId: 's', index: 0, value: 50 }), true)
  const transitionBefore = { sheets: { s: { cellData: { 0: { 0: { v: 'keep', t: 1 } }, 1: { 1: { v: 'outside', t: 1 } } }, rowData: { 0: { h: 24 } }, columnData: { 0: { w: 80 } } } } }
  for (const [type, patch, apply] of [
    ['setStyle', { sheetId: 's', row: 0, column: 0, value: '#223344', before: { v: 'keep', t: 1 } }, snapshot => { snapshot.sheets.s.cellData[0][0].s = { bg: { rgb: '#223344' } } }],
    ['setRowHeight', { sheetId: 's', index: 0, value: 50, before: { h: 24 } }, snapshot => { snapshot.sheets.s.rowData[0] = { h: 50, ia: 0 } }],
    ['setColumnWidth', { sheetId: 's', index: 0, value: 120, before: { w: 80 } }, snapshot => { snapshot.sheets.s.columnData[0].w = 120 }]
  ]) {
    const trusted = { type, ...patch }, after = clone(transitionBefore); apply(after)
    assert.strictEqual(api.sameControlledFormatTransition(transitionBefore, after, trusted), true, type)
    const outside = clone(after); outside.sheets.s.cellData[1][1].v = 'changed'
    assert.strictEqual(api.sameControlledFormatTransition(transitionBefore, outside, trusted), false, type + ' outside')
    assert.strictEqual(api.sameControlledFormatTransition(transitionBefore, after, { ...trusted, before: {} }), false, type + ' before')
  }
  const styled = clone(transitionBefore), styledPatch = { type: 'setStyle', sheetId: 's', row: 0, column: 0, value: '#223344', before: { v: 'keep', t: 1 } }
  styled.sheets.s.cellData[0][0].s = { bg: { rgb: '#223344' }, ff: 'unauthorized' }
  assert.strictEqual(api.sameControlledFormatTransition(transitionBefore, styled, styledPatch), false)
  const registryBefore = clone(transitionBefore), registryAfter = clone(transitionBefore)
  registryBefore.styles = { old: { ff: 'Arial', bg: { rgb: '#111111' } } }
  registryBefore.sheets.s.cellData[0][0].s = 'old'
  registryAfter.sheets.s.cellData[0][0].s = { ff: 'Arial', bg: { rgb: '#223344' } }
  const registryPatch = { ...styledPatch, before: { v: 'keep', t: 1, s: { ff: 'Arial', bg: { rgb: '#111111' } } } }
  assert.strictEqual(api.sameControlledFormatTransition(registryBefore, registryAfter, registryPatch), true)
  registryBefore.sheets.s.cellData[1][1].s = 'old'
  assert.strictEqual(api.sameControlledFormatTransition(registryBefore, registryAfter, registryPatch), false)
  console.log('PASS B5 server-manifest format path verifies actual readback while public capability stays closed')
}
run().catch(error => { console.error(error); process.exitCode = 1 })
