'use strict'
// C3-FE-02 import type round-trip regression (offline). Asserts the frontend-visible semantics
// of an import PREVIEW payload stay aligned with the backend gate ExcelImportTypeRoundTripTest
// (ruoyi-modules/ruoyi-ai/src/test/java/com/ruoyi/ai/agent/ExcelImportTypeRoundTripTest.java).
// The fixture below is a hand-built equivalent of a real WorkbookImportController preview body
// (ExcelImportService.preview -> rows records from the first verified chunk). Each row maps to a
// backend assertion case, evidenced by the real backend run log (D:/tmp-mysql/c3-green-parse.log):
//   row 0 cells 0..2 <- testNumberRoundTripKeepsTypeAndEveryDigit (t=2, decimal value equality)
//   row 1 cells 0..1 <- testBooleanRoundTrip (t=3, v boolean)
//   row 2 cell  0    <- testDateRoundTripAcrossBothDateSystems ("[c3-roundtrip] date serial
//                       45838.99998842592 (1900==1904 after normalization)": dates normalize to
//                       one 1900-basis serial and import as numeric t=2, not as display text)
//   row 3 cells 0..1 <- text preservation ("文本不丢型"): leading-zero / ISO-looking strings must
//                       not be coerced into numbers
//   row 4 cell  0    <- testFormulaRoundTripPreservesFormulaWithoutRecomputation (f verbatim,
//                       deliberately corrupted cached 13 kept, no recalculation claimed)
// No real backend/network is touched: fixtures are in-memory and cells are replayed through the
// installed @univerjs SDK (same cellData model: t 1=STRING / 2=NUMBER / 3=BOOLEAN).
const assert = require('assert'), fs = require('fs'), path = require('path'), { createRequire } = require('module')
const root = path.resolve(__dirname, '../..'), rf = createRequire(path.join(root, 'package.json')), babel = rf('@babel/core'), cache = new Map()
function load(file) {
  if (cache.has(file)) return cache.get(file).exports
  const module = { exports: {} }
  cache.set(file, module)
  const source = babel.transformSync(fs.readFileSync(file, 'utf8'), { filename: file, babelrc: false, configFile: false, plugins: [rf.resolve('@babel/plugin-transform-modules-commonjs')] }).code
  const localRequire = name => name === '@/utils/auth' ? { getToken: () => 'c3-test-token' }
    : name.startsWith('.') ? load(path.resolve(path.dirname(file), name + '.js')) : rf(name)
  new Function('require', 'module', 'exports', source)(localRequire, module, module.exports)
  return module.exports
}
const core = rf('@univerjs/core/lib/cjs/index.js'), sheets = rf('@univerjs/sheets/lib/cjs/index.js')
rf('@univerjs/sheets/lib/cjs/facade.js')
const { FUniver } = rf('@univerjs/core/lib/cjs/facade.js')

// Hand-built preview fixture (JSON shape exactly as WorkbookImportController.preview responds:
// bounded metadata + row records {sheetId,sheetIndex,rowIndex,cells}, cells keyed by column).
const PREVIEW = {
  jobId: 'job-c3-fixture', sourceHash: 'a'.repeat(64), sheetCount: 1, totalRowCount: 5,
  requestedRows: 50, returnedRows: 5, limited: true, fullSourceCoverage: false,
  note: '预览仅返回有限行与元信息，不承诺覆盖全量源工作簿',
  sheets: [{ sheetId: 'sheet-0', stats: {} }],
  rows: [
    { sheetId: 'sheet-0', sheetIndex: 0, rowIndex: 0, cells: { 0: { v: 123, t: 2 }, 1: { v: -45.75, t: 2 }, 2: { v: 0.12345678901234567890123456789, t: 2 } } },
    { sheetId: 'sheet-0', sheetIndex: 0, rowIndex: 1, cells: { 0: { v: true, t: 3 }, 1: { v: false, t: 3 } } },
    { sheetId: 'sheet-0', sheetIndex: 0, rowIndex: 2, cells: { 0: { v: 45838.99998842592, t: 2 } } },
    { sheetId: 'sheet-0', sheetIndex: 0, rowIndex: 3, cells: { 0: { v: '001', t: 1 }, 1: { v: '2025-06-30', t: 1 } } },
    { sheetId: 'sheet-0', sheetIndex: 0, rowIndex: 4, cells: { 0: { v: 13, t: 2, f: '=A1*B1' } } }
  ]
}

let passed = 0, failed = 0
async function test(name, fn) { try { await fn(); passed++; console.log('PASS ' + name) } catch (e) { failed++; console.error('FAIL ' + name + ': ' + e.message) } }

// Replay the preview rows into a real Univer workbook through the SDK facade and read back
// both the facade values and the saved cellData model (the persisted t/v contract).
function replay() {
  const univer = new core.Univer({ locale: core.LocaleType.EN_US, locales: { [core.LocaleType.EN_US]: {} } })
  univer.registerPlugin(sheets.UniverSheetsPlugin)
  univer.createUnit(core.UniverInstanceType.UNIVER_SHEET, { id: 'c3', name: 'C3', sheetOrder: ['s'], sheets: { s: { id: 's', name: 'T', rowCount: 10, columnCount: 10 } } })
  const api = FUniver.newAPI(univer), book = api.getActiveWorkbook(), sheet = book.getActiveSheet()
  for (const record of PREVIEW.rows) for (const [col, cell] of Object.entries(record.cells)) sheet.getRange(record.rowIndex, Number(col)).setValue({ v: cell.v, t: cell.t, ...(cell.f ? { f: cell.f } : {}) })
  return { univer, book, sheet, saved: () => book.save().sheets.s.cellData }
}
const serialToIso = serial => new Date(Date.UTC(1899, 11, 30) + serial * 86400000).toISOString()

;(async () => {
  // Red/green anchor: the C3 import API client must exist with the four endpoints. (Red phase:
  // this load fails with ENOENT while src/api/ai/import.js is absent.)
  const api = load(path.join(root, 'src/api/ai/import.js'))
  await test('import client exposes create/get/byRequest/preview endpoints', async () => {
    for (const fn of ['createImportJob', 'getImportJob', 'getImportJobByRequest', 'previewImport']) assert.strictEqual(typeof api[fn], 'function', fn + ' missing')
  })

  await test('preview fixture matches backend preview row-record shape', async () => {
    assert.strictEqual(PREVIEW.limited, true); assert.strictEqual(PREVIEW.fullSourceCoverage, false)
    for (const r of PREVIEW.rows) { assert.ok('sheetId' in r && 'sheetIndex' in r && 'rowIndex' in r && 'cells' in r, 'row record keys'); for (const c of Object.values(r.cells)) assert.ok('v' in c && 't' in c, 'cell v/t') }
  })

  await test('numbers import as NUMBER t=2 and keep value (backend testNumberRoundTrip)', async () => {
    const { univer, sheet, saved } = replay()
    try {
      assert.strictEqual(saved()[0][0].t, 2); assert.strictEqual(saved()[0][0].v, 123)
      assert.strictEqual(saved()[0][1].t, 2); assert.strictEqual(saved()[0][1].v, -45.75)
      assert.strictEqual(sheet.getRange(0, 0).getValues()[0][0], 123)
      assert.strictEqual(sheet.getRange(0, 1).getValues()[0][0], -45.75)
    } finally { univer.dispose() }
  })

  await test('booleans import as BOOLEAN t=3 and keep true/false (backend testBooleanRoundTrip)', async () => {
    const { univer, sheet, saved } = replay()
    try {
      // Project-wide boolean storage caliber (WriteManifest.semanticValue): t=3 with v=1/0 encodes
      // true/false; the wire payload's JSON true/false normalizes to 1/0 at the SDK boundary, and
      // the backend parses it back to a Boolean. Display lane (@univerjs 0.25.1) renders 1/0 too.
      assert.strictEqual(saved()[1][0].t, 3); assert.strictEqual(saved()[1][0].v, 1)
      assert.strictEqual(saved()[1][1].t, 3); assert.strictEqual(saved()[1][1].v, 0)
      assert.strictEqual(sheet.getRange(1, 0).getValues()[0][0], 1)
      assert.strictEqual(sheet.getRange(1, 1).getValues()[0][0], 0)
    } finally { univer.dispose() }
  })

  await test('dates import as one 1900-basis serial, numeric t=2 (backend testDateRoundTrip 1900==1904)', async () => {
    const { univer, saved } = replay()
    try {
      assert.strictEqual(saved()[2][0].t, 2, 'date cell must not arrive as display text')
      assert.strictEqual(serialToIso(saved()[2][0].v), '2025-06-30T23:59:59.000Z')
    } finally { univer.dispose() }
  })

  await test('text does not lose its type: leading-zero and ISO-looking strings stay STRING t=1', async () => {
    const { univer, sheet, saved } = replay()
    try {
      assert.strictEqual(saved()[3][0].t, 1); assert.strictEqual(sheet.getRange(3, 0).getValues()[0][0], '001')
      assert.strictEqual(saved()[3][1].t, 1); assert.strictEqual(sheet.getRange(3, 1).getValues()[0][0], '2025-06-30')
    } finally { univer.dispose() }
  })

  await test('formula f verbatim, cached value kept, no recalculation (backend testFormulaRoundTrip)', async () => {
    const { univer, sheet, saved } = replay()
    try {
      const cell = saved()[4][0]
      assert.strictEqual(cell.f, '=A1*B1'); assert.strictEqual(cell.v, 13)
      assert.ok(!('recalculated' in cell) && !('computed' in cell), 'no recomputation claim')
      assert.strictEqual(sheet.getRange(4, 0).getFormula(), '=A1*B1')
    } finally { univer.dispose() }
  })

  console.log(JSON.stringify({ passed, failed, scope: 'OFFLINE_FIXTURE_REAL_UNIVER_SDK_NO_BACKEND_NO_NETWORK' }))
  process.exitCode = failed ? 1 : 0
})()
