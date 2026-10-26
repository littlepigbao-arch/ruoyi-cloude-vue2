'use strict'
const assert = require('assert')
const fs = require('fs')
const path = require('path')
const vm = require('vm')
const root = path.resolve(__dirname, '../..')
const babel = require(path.join(root, 'node_modules/@babel/core'))
const plugin = require(path.join(root, 'node_modules/@babel/plugin-transform-modules-commonjs'))
const cache = new Map()
function load(file) {
  if (cache.has(file)) return cache.get(file)
  const m = { exports: {} }
  vm.runInNewContext(babel.transformSync(fs.readFileSync(file, 'utf8'), { configFile: false, babelrc: false, plugins: [plugin] }).code,
    { module: m, exports: m.exports, console, performance, require: name => load(path.resolve(path.dirname(file), name + '.js')) })
  cache.set(file, m.exports)
  return m.exports
}
const { probeControlledFormatAction } = load(path.join(root, 'src/views/application/agentReport/controlled-format-foundation.js'))
const { executeActions } = load(path.join(root, 'src/views/application/agentReport/univer-action-executor.js'))
const copy = x => JSON.parse(JSON.stringify(x))
function fixture({ missing = false, missingGetter = false, throws = false, result = undefined, silent = false, wrong = false, outside = false, facadeLie = false, snapshotFails = false, styleDamage = false, registryRef = false, registryTamper = false, existingStyle = false, richBackground = false } = {}) {
  const snapshot = { styles: {}, sheets: { s: { name: 'Sheet', rowCount: 10, columnCount: 10, defaultRowHeight: 24, cellData: { 0: { 0: { v: 'keep', t: 1 } } }, rowData: { 0: { h: 24 } }, columnData: { 0: { w: 80 } } } } }
  let calls = 0
  const sheetData = snapshot.sheets.s
  if (existingStyle || richBackground) sheetData.cellData[0][0].s = { bg: richBackground ? { rgb: '#ffffff', theme: 'kept' } : { rgb: '#ffffff' }, ff: 'Arial' }
  const range = { getBackground: () => { const ref = sheetData.cellData[0][0].s; return facadeLie ? '#ff0000' : (typeof ref === 'string' ? snapshot.styles[ref] : ref)?.bg?.rgb || '#ffffff' },
    setBackgroundColor(color) { calls++; if (throws) throw Error('setter failed'); if (!silent) { sheetData.cellData[0][0].s = { ...sheetData.cellData[0][0].s, bg: { ...sheetData.cellData[0][0].s?.bg, rgb: wrong ? '#000000' : color } }; if (styleDamage) sheetData.cellData[0][0].s.ff = 'Changed'; if (registryRef || registryTamper) { snapshot.styles.newStyle = sheetData.cellData[0][0].s; sheetData.cellData[0][0].s = 'newStyle' } if (registryTamper) snapshot.styles.unrelated = { bg: { rgb: '#222222' } }; if (outside) sheetData.cellData[0][1] = { v: 'damage' } } return result === undefined ? this : result } }
  if (missing) delete range.setBackgroundColor
  if (missingGetter) delete range.getBackground
  const sheet = { getSheetId: () => 's', getSheetName: () => 'Sheet', getMaxRows: () => 10, getMaxColumns: () => 10,
    getRange: () => range, getRowHeight: i => sheetData.rowData[i]?.h || 24, getColumnWidth: i => sheetData.columnData[i]?.w || 80,
    setRowHeight(i, height) { calls++; if (throws) throw Error('row failed'); if (!silent) sheetData.rowData[i] = { h: wrong ? height + 1 : height }; return result === undefined ? this : result },
    setColumnWidth(i, width) { calls++; if (throws) throw Error('column failed'); if (!silent) sheetData.columnData[i] = { w: wrong ? width + 1 : width }; return result === undefined ? this : result } }
  if (missing) delete sheet.setRowHeight, delete sheet.setColumnWidth
  if (missingGetter) delete sheet.getRowHeight, delete sheet.getColumnWidth
  let saves = 0
  const book = { getSheetBySheetId: id => id === 's' ? sheet : null, getActiveSheet: () => sheet, save: () => { if (snapshotFails && ++saves > 1) throw Error('save readback failed'); return copy(snapshot) } }
  return { book, calls: () => calls }
}
const base = { schemaVersion: '2.0', type: 'setStyle', target: { sheetId: 's', range: 'A1' }, background: '#ff0000' }
// The installed Facade returns its range even when its internal command returns false.
// The existing executor cannot infer that failure from the setter return alone.
assert.strictEqual(executeActions({}, fixture({ silent: true }).book, [base], { continueOnError: false }).summary.ok, 1)
function rejectBefore(action) { const f = fixture(); assert.throws(() => probeControlledFormatAction({}, f.book, action)); assert.strictEqual(f.calls(), 0) }
for (const action of [
  { ...base, fontColor: '#000000' }, { ...base, fontSize: 12 }, { ...base, background: 'red' },
  { ...base, background: '#ff0000', target: { sheetId: 'bad', range: 'A1' } },
  { ...base, target: { sheetId: 's', range: 'A1:B1' } }, { ...base, target: { sheetId: 's', range: 'K1' } },
  { ...base, sheetId: 'bad' }, { ...base, target: { sheetId: 's' } },
  { schemaVersion: '2.0', type: 'setRowHeight', target: { sheetId: 's' }, rowPosition: 0, height: 50 },
  { schemaVersion: '2.0', type: 'setRowHeight', target: { sheetId: 's' }, rowPosition: 11, height: 50 },
  { schemaVersion: '2.0', type: 'setRowHeight', target: { sheetId: 's' }, rowPosition: 1, height: 0 },
  { schemaVersion: '2.0', type: 'setColumnWidth', target: { sheetId: 's' }, columnPosition: 1.5, width: 100 },
  { schemaVersion: '2.0', type: 'setColumnWidth', target: { sheetId: 's' }, columnPosition: 1, width: 10000 }
]) rejectBefore(action)
for (const options of [{ missing: true }, { missingGetter: true }, { throws: true }, { result: false }, { result: { error: 'failed' } }, { silent: true }, { silent: true, facadeLie: true }, { wrong: true }, { outside: true }, { snapshotFails: true }]) {
  const f = fixture(options)
  assert.throws(() => probeControlledFormatAction({}, f.book, base), /核实|失败|缺少|回读/)
}
for (const options of [{ styleDamage: true }, { registryTamper: true }]) {
  const f = fixture(options)
  assert.throws(() => probeControlledFormatAction({}, f.book, base), /核实|拒绝/)
}
assert.strictEqual(probeControlledFormatAction({}, fixture({ existingStyle: true }).book, base).verified, true)
const unsupportedBackground = fixture({ richBackground: true })
assert.throws(() => probeControlledFormatAction({}, unsupportedBackground.book, base), /B5_FORMAT_REJECTED/)
assert.strictEqual(unsupportedBackground.calls(), 0)
assert.strictEqual(probeControlledFormatAction({}, fixture({ registryRef: true }).book, base).verified, true)
for (const type of ['setRowHeight', 'setColumnWidth']) {
  const action = { schemaVersion: '2.0', type, target: { sheetId: 's' }, ...(type === 'setRowHeight' ? { rowPosition: 1, height: 50 } : { columnPosition: 1, width: 100 }) }
  const f = fixture(); const proved = probeControlledFormatAction({}, f.book, action); assert.strictEqual(proved.verified, true); assert.strictEqual(proved.formatReadback.value, type === 'setRowHeight' ? 50 : 100)
  for (const options of [{ missing: true }, { missingGetter: true }, { throws: true }, { result: false }, { result: { error: 'failed' } }, { silent: true }, { wrong: true }, { snapshotFails: true }]) {
    const failed = fixture(options)
    assert.throws(() => probeControlledFormatAction({}, failed.book, action), /B5_FORMAT_REJECTED|RESULT_UNKNOWN/)
    if (options.missing || options.missingGetter) assert.strictEqual(failed.calls(), 0)
  }
}
const styled = probeControlledFormatAction({}, fixture().book, base)
assert.strictEqual(styled.verified, true)
assert.deepStrictEqual(JSON.parse(JSON.stringify(styled.formatReadback)), { sheetId: 's', row: 0, column: 0, value: '#ff0000' })
console.log('PASS B5 format validation, readback, and failure isolation')
