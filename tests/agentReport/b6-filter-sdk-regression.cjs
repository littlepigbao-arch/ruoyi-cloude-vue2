'use strict'
const assert = require('assert')
const fs = require('fs'), path = require('path'), { createRequire } = require('module')
const root = path.resolve(__dirname, '../..'), requireFront = createRequire(path.join(root, 'package.json'))
const babel = requireFront('@babel/core'), cache = new Map()
function load(file) {
  if (cache.has(file)) return cache.get(file).exports
  const module = { exports: {} }; cache.set(file, module)
  const source = babel.transformSync(fs.readFileSync(file, 'utf8'), { filename: file, babelrc: false, configFile: false, plugins: [requireFront.resolve('@babel/plugin-transform-modules-commonjs')] }).code
  new Function('require', 'module', 'exports', source)(name => name.startsWith('.') ? load(path.resolve(path.dirname(file), name + '.js')) : requireFront(name), module, module.exports)
  return module.exports
}
const controlled = load(path.join(root, 'src/views/application/agentReport/controlled-write-executor.js'))
const core = requireFront('@univerjs/core/lib/cjs/index.js'), sheets = requireFront('@univerjs/sheets/lib/cjs/index.js'), filterPlugin = requireFront('@univerjs/sheets-filter/lib/cjs/index.js')
requireFront('@univerjs/sheets/lib/cjs/facade.js'); requireFront('@univerjs/sheets-filter/lib/cjs/facade.js')
const { FUniver } = requireFront('@univerjs/core/lib/cjs/facade.js')
function setup(numeric = false) {
  const univer = new core.Univer({ locale: core.LocaleType.EN_US, locales: { [core.LocaleType.EN_US]: {} } })
  univer.registerPlugin(sheets.UniverSheetsPlugin); univer.registerPlugin(filterPlugin.UniverSheetsFilterPlugin)
  univer.createUnit(core.UniverInstanceType.UNIVER_SHEET, { id: 'b6', name: 'B6', sheetOrder: ['s'], sheets: { s: { id: 's', name: 'Sheet', rowCount: 8, columnCount: 3,
    cellData: { 0: { 0: { v: 'Name', t: 1 } }, 1: { 0: { v: numeric ? 10 : 'Alice', t: numeric ? 2 : 1 } }, 2: { 0: { v: numeric ? 20 : 'Bob', t: numeric ? 2 : 1 } }, 3: { 0: { v: numeric ? 10 : 'Alice', t: numeric ? 2 : 1 } } } } } })
  const api = FUniver.newAPI(univer), book = api.getActiveWorkbook()
  return { univer, api, book, sheet: book.getSheetBySheetId('s') }
}
const ref = { startRow: 0, startColumn: 0, endRow: 3, endColumn: 0 }
const criterion = { colId: 0, filters: { filters: ['Alice'] } }
const filtered = { ref, filterColumns: [criterion], cachedFilteredOut: [2] }
const empty = { ref, filterColumns: [], cachedFilteredOut: [] }
const action = { schemaVersion: '2.0', type: 'setFilter', target: { sheetId: 's', range: 'A1:A4' }, columnIndex: 0, filters: ['Alice'] }
const clearColumn = { schemaVersion: '2.0', type: 'clearFilter', target: { sheetId: 's', range: 'A1:A4' }, columnIndex: 0 }
const removeAll = { schemaVersion: '2.0', type: 'clearFilter', target: { sheetId: 's', range: 'A1:A4' }, removeFilter: true }
function resource(book) { return book.save().resources.find(x => x.name === 'SHEET_FILTER_PLUGIN').data }
function manifest(a, beforeData, filter) {
  const afterData = JSON.stringify(filter ? { s: filter } : {})
  const patch = { type: a.type, sheetId: 's', range: 'A1:A4', beforeData, resourceData: afterData, filter }
  return { transportVersion: 2, action: a, actions: [a], filterPatch: patch, filterReadback: { sheetId: 's', range: 'A1:A4', hasFilter: !!filter, filter } }
}
async function main() {
  assert.throws(() => controlled.classifyAiActions([action]), /AI_CAPABILITY_BLOCKED/)
  assert.strictEqual(controlled.classifyAiActions([action], ['setFilter']), 'controlled-values')
  assert.throws(() => controlled.classifyAiActions([action, action], ['setFilter']), /AI_CAPABILITY_BLOCKED/)
  const { univer, api, book } = setup()
  try {
    let prior = book.save(), m = manifest(action, resource(book), filtered)
    const result = await controlled.executeControlledWrite(api, book, m)
    assert.strictEqual(result.ok, true)
    assert.deepStrictEqual(JSON.parse(JSON.stringify(result.filterReadback)), m.filterReadback)
    assert.strictEqual(controlled.sameControlledFilterTransition(prior, book.save(), m.filterPatch), true)
    const readOnlyBefore = book.save(), observed = book.getSheetBySheetId('s').getFilter()
    assert.deepStrictEqual(observed.getFilteredOutRows(), [2])
    assert.deepStrictEqual(book.save(), readOnlyBefore)
    const bad = JSON.parse(JSON.stringify(book.save())); bad.resources.push({ name: 'UNRELATED', data: 'tamper' })
    assert.strictEqual(controlled.sameControlledFilterSnapshot(bad, book.save(), m.filterPatch), false)
    prior = book.save(); m = manifest(clearColumn, resource(book), empty)
    assert.strictEqual((await controlled.executeControlledWrite(api, book, m)).ok, true)
    assert.strictEqual(controlled.sameControlledFilterTransition(prior, book.save(), m.filterPatch), true)
    prior = book.save(); m = manifest(removeAll, resource(book), null)
    assert.strictEqual((await controlled.executeControlledWrite(api, book, m)).ok, true)
    assert.strictEqual(controlled.sameControlledFilterTransition(prior, book.save(), m.filterPatch), true)
    console.log('PASS B6 actual SDK set/clear/remove, facade readback and resource isolation')
  } finally { univer.dispose() }
  const numeric = setup(true)
  try {
    const numericAction = { ...action, filters: ['10'] }, numericFilter = { ...filtered, filterColumns: [{ colId: 0, filters: { filters: ['10'] } }] }
    const result = await controlled.executeControlledWrite(numeric.api, numeric.book, manifest(numericAction, resource(numeric.book), numericFilter))
    assert.strictEqual(result.ok, true)
    assert.deepStrictEqual(numeric.sheet.getFilter().getFilteredOutRows(), [2])
    console.log('PASS B6 numeric exact match')
  } finally { numeric.univer.dispose() }
}
main().catch(error => { console.error(error); process.exitCode = 1 })
