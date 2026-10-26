'use strict'
// Executes the production B5 probe against the installed Univer 0.25.1 packages.
// The workbook is in memory; this does not exercise HTTP, SQL, CAS, or the browser UI.
const assert = require('assert')
const fs = require('fs')
const path = require('path')
const { createRequire } = require('module')
const front = path.resolve(__dirname, '../..')
const requireFront = createRequire(path.join(front, 'package.json'))
const babel = requireFront('@babel/core')
const cache = new Map()
function load(filename) {
  if (cache.has(filename)) return cache.get(filename).exports
  const module = { exports: {} }
  cache.set(filename, module)
  const source = babel.transformSync(fs.readFileSync(filename, 'utf8'), {
    filename, babelrc: false, configFile: false,
    plugins: [requireFront.resolve('@babel/plugin-transform-modules-commonjs')]
  }).code
  const localRequire = name => name.startsWith('.')
    ? load(path.resolve(path.dirname(filename), name + '.js')) : requireFront(name)
  new Function('require', 'module', 'exports', source)(localRequire, module, module.exports)
  return module.exports
}
const core = requireFront('@univerjs/core/lib/cjs/index.js')
const sheets = requireFront('@univerjs/sheets/lib/cjs/index.js')
requireFront('@univerjs/sheets/lib/cjs/facade.js')
const { FUniver } = requireFront('@univerjs/core/lib/cjs/facade.js')
const { probeControlledFormatAction } = load(path.join(front, 'src/views/application/agentReport/controlled-format-foundation.js'))

function workbook(priorStyle, priorRow = { h: 24 }, priorColumn = { w: 80 }) {
  const univer = new core.Univer({ locale: core.LocaleType.EN_US, locales: { [core.LocaleType.EN_US]: {} } })
  univer.registerPlugin(sheets.UniverSheetsPlugin)
  const target = { v: 'keep', t: 1 }
  if (priorStyle) target.s = priorStyle
  univer.createUnit(core.UniverInstanceType.UNIVER_SHEET, {
    id: 'b5-sdk', name: 'B5 SDK contract', sheetOrder: ['s'],
    sheets: { s: { id: 's', name: 'Sheet', rowCount: 10, columnCount: 10,
      defaultRowHeight: 24, cellData: { 0: { 0: target } },
      rowData: { 0: priorRow }, columnData: { 0: priorColumn } } }
  })
  const api = FUniver.newAPI(univer)
  return { univer, api, book: api.getActiveWorkbook() }
}

// The saved representation must preserve unrelated existing metadata when the
// installed SDK runs the real setter, including fields in the same style node.
for (const [action, style, rowData, columnData, expected] of [
  [{ schemaVersion: '2.0', type: 'setRowHeight', target: { sheetId: 's' }, rowPosition: 1, height: 50 },
    null, { h: 24, customFlag: 'kept' }, { w: 80 }, 'row'],
  [{ schemaVersion: '2.0', type: 'setColumnWidth', target: { sheetId: 's' }, columnPosition: 1, width: 120 },
    null, { h: 24 }, { w: 80, customFlag: 'kept' }, 'column']
]) {
  const { univer, api, book } = workbook(style, rowData, columnData)
  try {
    assert.strictEqual(probeControlledFormatAction(api, book, action).verified, true, `${expected} metadata`)
    const saved = book.save()
    if (expected === 'row') assert.deepStrictEqual(saved.sheets.s.rowData[0], { h: 50, customFlag: 'kept', ia: 0 })
    if (expected === 'column') assert.deepStrictEqual(saved.sheets.s.columnData[0], { w: 120, customFlag: 'kept' })
    console.log(`${expected} metadata: preserved by installed Univer`)
  } finally { univer.dispose() }
}

// Univer 0.25.1 drops other keys in bg when setBackgroundColor runs. Reject
// that input before the setter can partly mutate the local workbook.
{
  const { univer, api, book } = workbook({ ff: 'Arial', bg: { rgb: '#111111', theme: 'kept' } })
  try {
    const before = book.save()
    assert.throws(() => probeControlledFormatAction(api, book,
      { schemaVersion: '2.0', type: 'setStyle', target: { sheetId: 's', range: 'A1' }, background: '#a1b2c3' }), /B5_FORMAT_REJECTED/)
    assert.deepStrictEqual(book.save(), before, 'unsupported background metadata was changed before rejection')
  } finally { univer.dispose() }
}

for (const [action, priorStyle] of [
  [{ schemaVersion: '2.0', type: 'setStyle', target: { sheetId: 's', range: 'A1' }, background: '#a1b2c3' }, null],
  [{ schemaVersion: '2.0', type: 'setStyle', target: { sheetId: 's', range: 'A1' }, background: '#a1b2c3' }, { ff: 'Arial', bg: { rgb: '#111111' } }],
  [{ schemaVersion: '2.0', type: 'setRowHeight', target: { sheetId: 's' }, rowPosition: 1, height: 50 }, null],
  [{ schemaVersion: '2.0', type: 'setColumnWidth', target: { sheetId: 's' }, columnPosition: 1, width: 120 }, null]
]) {
  const { univer, api, book } = workbook(priorStyle)
  try {
    const result = probeControlledFormatAction(api, book, action)
    assert.strictEqual(result.verified, true, action.type)
    assert.strictEqual(result.formatReadback.value, action.background || action.height || action.width)
    const saved = book.save()
    assert.strictEqual(saved.sheets.s.cellData[0][0].v, 'keep')
    if (action.type === 'setStyle') {
      const ref = saved.sheets.s.cellData[0][0].s
      assert.strictEqual(typeof ref, 'string')
      assert.strictEqual(saved.styles[ref].bg.rgb, '#a1b2c3')
      if (priorStyle) assert.strictEqual(saved.styles[ref].ff, 'Arial')
    }
    if (action.type === 'setRowHeight') assert.deepStrictEqual(saved.sheets.s.rowData[0], { h: 50, ia: 0 })
    if (action.type === 'setColumnWidth') assert.deepStrictEqual(saved.sheets.s.columnData[0], { w: 120 })
    console.log(`${action.type}: verified against installed Univer`)
  } finally { univer.dispose() }
}
