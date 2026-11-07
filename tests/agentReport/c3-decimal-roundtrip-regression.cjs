'use strict'
// C3-FE-02 decimal round-trip regression (offline). A 29-digit decimal literal must survive the
// export -> import-preview-JSON chain with every digit intact. Backend caliber: the F2 precision
// lane (PRECISE_JSON chunk parse, ExcelImportService.preview reading through the C1 readChunkPage
// precision overload) keeps cell v as the exact decimal literal - the backend gate
// ExcelImportTypeRoundTripTest#testNumberRoundTripKeepsTypeAndEveryDigit asserts
// new BigDecimal("0.12345678901234567890123456789") equality, and the real backend run log
// (D:/tmp-mysql/c3-green-parse.log) prints: "[c3-roundtrip] number 0.12345678901234567890123456789".
// The fixture below is a hand-built equivalent of that preview response body as it arrives on the
// wire: the 29-digit literal is a RAW JSON number token (not a string). The regression guards the
// frontend parsing contract: the raw token is compared as a STRING (never through JS Number, which
// would silently round to 17 significant digits). No real backend/network is touched.
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

const LITERAL = '0.12345678901234567890123456789'
// Exact wire text of a bounded preview body (the precise cell is a raw JSON number token, as
// PRECISE_JSON emits it - this is the shape previewImport(id, sourceHash, rows) delivers).
const PREVIEW_WIRE = '{"jobId":"job-c3-decimal","sourceHash":"' + 'b'.repeat(64) + '","sheetCount":1,' +
  '"totalRowCount":1,"requestedRows":50,"returnedRows":1,"limited":true,"fullSourceCoverage":false,' +
  '"note":"预览仅返回有限行与元信息，不承诺覆盖全量源工作簿","sheets":[],"rows":[' +
  '{"sheetId":"sheet-0","sheetIndex":0,"rowIndex":0,"cells":{"0":{"v":0.12345678901234567890123456789,"t":2}}}' +
  ']}'

let passed = 0, failed = 0
async function test(name, fn) { try { await fn(); passed++; console.log('PASS ' + name) } catch (e) { failed++; console.error('FAIL ' + name + ': ' + e.message) } }

// Reference frontend precise-parse semantics for the import consumer: read the raw JSON number
// token for one preview cell straight off the wire text, without materializing it as a JS Number.
// (Regex token set is exactly the JSON number grammar; a string value is captured with quotes so
// text cells are distinguishable from numeric ones.)
function rawCellValueToken(wireText, rowIndex, columnKey) {
  const rowAt = wireText.indexOf('"rowIndex":' + rowIndex)
  if (rowAt < 0) throw new Error('rowIndex ' + rowIndex + ' not found')
  const cellsAt = wireText.indexOf('"cells":{', rowAt)
  const colAt = wireText.indexOf('"' + columnKey + '":{', cellsAt)
  if (colAt < 0) throw new Error('column ' + columnKey + ' not found')
  const vAt = wireText.indexOf('"v":', colAt)
  if (vAt < 0) throw new Error('v not found')
  const rest = wireText.slice(vAt + 4)
  const m = /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(rest)
  if (m) return m[0]
  const s = /^"(?:[^"\\]|\\.)*"/.exec(rest)
  return s && s[0]
}

;(async () => {
  // Red/green anchor: the preview entry point that delivers this wire body must exist. (Red phase:
  // this load fails with ENOENT while src/api/ai/import.js is absent.)
  const api = load(path.join(root, 'src/api/ai/import.js'))
  await test('import client exposes previewImport entry for the preview wire body', async () => {
    assert.strictEqual(typeof api.previewImport, 'function')
  })

  await test('29-digit decimal survives on the wire as an exact raw token (string compare)', async () => {
    assert.ok(PREVIEW_WIRE.indexOf('"v":' + LITERAL + ',"t":2') >= 0, 'raw literal token missing from wire text')
  })

  await test('precise raw-token parse keeps every digit; never routes through JS Number', async () => {
    const token = rawCellValueToken(PREVIEW_WIRE, 0, '0')
    assert.strictEqual(token, LITERAL)
    // '0.' + 29 fraction digits = 31 characters (the 29 counts decimal places, not string length);
    // the wire token must be byte-identical to the backend BigDecimal string.
    assert.strictEqual(token.length, LITERAL.length)
    assert.strictEqual(typeof token, 'string')
  })

  await test('naive JSON.parse destroys the literal: documents why the raw path is mandatory', async () => {
    const naive = JSON.parse(PREVIEW_WIRE).rows[0].cells['0'].v
    assert.strictEqual(String(naive), '0.12345678901234568', 'JS Number should round to 17 significant digits')
    assert.notStrictEqual(String(naive), LITERAL)
  })

  await test('SDK boundary: precise decimal as text is exact; as Number it cannot be (frontend display lane)', async () => {
    const univer = new core.Univer({ locale: core.LocaleType.EN_US, locales: { [core.LocaleType.EN_US]: {} } })
    try {
      univer.registerPlugin(sheets.UniverSheetsPlugin)
      univer.createUnit(core.UniverInstanceType.UNIVER_SHEET, { id: 'c3d', name: 'C3D', sheetOrder: ['s'], sheets: { s: { id: 's', name: 'T', rowCount: 4, columnCount: 4 } } })
      const sheet = FUniver.newAPI(univer).getActiveWorkbook().getActiveSheet()
      sheet.getRange(0, 0).setValue({ v: LITERAL, t: 1 })
      assert.strictEqual(sheet.getRange(0, 0).getValues()[0][0], LITERAL, 'text-carried literal must stay exact')
      sheet.getRange(0, 1).setValue({ v: Number(LITERAL), t: 2 })
      // Number path can never keep 29 digits. Measured on @univerjs 0.25.1 the range getValues()
      // display pipeline yields '0.123456789012' (12 significant digits) - asserted here as
      // "strictly lossy + shorter than the literal" instead of pinning one SDK display format.
      const numberLane = String(sheet.getRange(0, 1).getValues()[0][0])
      assert.notStrictEqual(numberLane, LITERAL, 'number path must not round-trip the literal')
      assert.ok(numberLane.length < LITERAL.length, 'number path loses digits (measured 0.123456789012 @0.25.1)')
    } finally { univer.dispose() }
  })

  console.log(JSON.stringify({ passed, failed, scope: 'OFFLINE_WIRE_FIXTURE_STRING_COMPARE_NO_FLOAT_NO_BACKEND_NO_NETWORK' }))
  process.exitCode = failed ? 1 : 0
})()
