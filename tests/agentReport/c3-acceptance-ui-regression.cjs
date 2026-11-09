'use strict'
// C3 acceptance page-method regression (offline, no backend, no network).
// Second-review BLOCKED remediation (F1/F2/F4a/F4b): the four counter-examples that exposed
// false greens on ImportAcceptance.vue are solidified here as page-method regressions, run with
// the real .vue source loaded the b10-import-ui-regression way (babel <script> extraction + vm +
// data()/methods binding + require-proxy mocks; element-ui components live in the template and
// need no mock; $message is a stub; '@/api/ai/import' is a controllable stub).
//   F1 (P1) - assertPreviewPrecision only checked the FORM (length===31 && /^0\.\d{29}$/), so
//             swapping the correct literal for 0.99999999999999999999999999999 still PASSed.
//             Fix under test: an independent fixture-expectation table
//             (src/views/application/agentReport/c3-acceptance-fixture-expectations.js, values
//             derived by the real BigDecimal serialization rule) is compared cell by cell; a
//             preview that misses the fixture signature (another file uploaded) may only record
//             人工核对.
//   F2 (P2) - runProbe treated the Axios TRANSPORT code (err.code='ERR_BAD_REQUEST', a string,
//             with err.response present) as a business envelope code, so an HTTP 403 was judged
//             "connected + authed OK". Fix under test: merged verdict
//             isAuthDenied(Number(err.code)) || isAuthDenied(err.response.status); the business
//             branch only accepts a numeric finite code; Axios errors carrying err.response are
//             always classified by HTTP status.
//   F4a(P2) - assertFormulaCalc PASSed a mixed scenario (one decidable match + one undecidable).
//             Fix under test: any undecidable formula forces 人工核对; PASS requires ALL formulas
//             decidable + f verbatim + equal to the evalSimpleFormula expectation.
//   F4b(P2) - compareBackendCellData kept fDiffCount informational, so a backend reload that lost
//             the formula f (cached value unchanged) PASSed. Fix under test: fDiffCount>0 → FAIL
//             with the f-loss coordinates listed.
// Plus (三审 P2 收口): the forward full chain 「真实后端预览 JSON → 生产解析器 → 页面断言」 runs
// offline: the backend's actual preview envelope (raw JSON number tokens on the wire, E-notation
// grammar included) goes through the real import.js parser into the page assertion, and the
// expectation table itself is re-derived INDEPENDENTLY with BigInt from (significand, exponent)
// pairs per the real BigDecimal.toString wire rule (negative scale folds to an integer; positive
// scale uses E-notation when the adjusted exponent < -6, plain decimals otherwise) and must equal
// FIXTURE_EXPECTATIONS exactly - proving the table depends on no API response.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert')
const root = path.resolve(__dirname, '../..')
const babel = require(path.join(root, 'node_modules/@babel/core'))
const plugin = require(path.join(root, 'node_modules/@babel/plugin-transform-modules-commonjs'))
const PAGE_FILE = path.join(root, 'src/views/application/agentReport/ImportAcceptance.vue')
const EXPECTATIONS_FILE = path.join(root, 'src/views/application/agentReport/c3-acceptance-fixture-expectations.js')

let passed = 0, failed = 0
async function test(name, fn) {
  try { await fn(); passed++; console.log('PASS ' + name) } catch (e) { failed++; console.error('FAIL ' + name + ': ' + e.message) }
}

// ---------- .vue loading (b10-import-ui-regression pattern) ----------
const importApi = {}, workbookApi = {}
const cache = new Map()
function load(file) {
  if (cache.has(file)) return cache.get(file)
  let src = fs.readFileSync(file, 'utf8')
  if (file.endsWith('.vue')) src = src.match(/<script>([\s\S]*?)<\/script>/)[1]
  const module = { exports: {} }
  cache.set(file, module)
  const code = babel.transformSync(src, { configFile: false, babelrc: false, plugins: [plugin] }).code
  vm.runInNewContext(code, {
    module, exports: module.exports, console, performance, Blob,
    process: { env: {} }, // data() reads process.env.VUE_APP_BASE_API
    sessionStorage: { getItem: () => null, setItem() {}, removeItem() {} },
    setTimeout: () => 0, clearTimeout() {}, setInterval: () => 0, clearInterval() {},
    require: name => {
      if (name === '@/api/ai/import') return new Proxy({}, { get: (_, k) => importApi[k] })
      if (name === '@/api/ai/workbook') return new Proxy({}, { get: (_, k) => workbookApi[k] })
      if (name === '@/utils/auth') return { getToken: () => 'c3-ui-regression-stub-token' }
      if (name.startsWith('.')) {
        const target = path.resolve(path.dirname(file), name + (path.extname(name) ? '' : '.js'))
        return load(target)
      }
      return {} // @univerjs/* presets, js-cookie, css side-effect imports: loaded but never called
    }
  }, { filename: file })
  cache.set(file, module.exports)
  return module.exports
}

// Real page component: data() + methods bound onto a plain object (no Vue instance needed for
// method-level regression; reactive watchers are not under test here).
function page() {
  const c = load(PAGE_FILE).default
  const p = c.data()
  for (const [k, v] of Object.entries(c.methods)) p[k] = v.bind(p)
  p.$message = { warning() {}, success() {}, error() {}, info() {} }
  p.$nextTick = () => Promise.resolve()
  p.$refs = {}
  return p
}

// The expectation table under test. Loaded with red-phase tolerance: at red time the module does
// not exist yet, the derivation test must fail with that reason while the four counter-examples
// still run against the CURRENT page.
let expectations = null, expectationsError = null
try { expectations = load(EXPECTATIONS_FILE) } catch (e) { expectationsError = e }

// ---------- fixture-shaped preview built from INDEPENDENT literals ----------
// Literal values per the REAL backend serialization contract (ExcelImportService
// wholeFractionNormalized + java.math.BigDecimal.toString; Codex isolated-JVM measurement +
// production parser probe): 1e400 carries a NEGATIVE scale (-400) that folds to a 401-digit
// integer literal on the PRECISE_JSON wire; 1E-400 / 4e-324 / 1.23456789012345e-310 carry
// POSITIVE scales and keep their BigDecimal.toString E-notation text ('1E-400' / '4E-324' /
// '1.23456789012345E-310' - NOT expanded plain decimals); the 29-digit C1 literal (scale 29,
// adjusted exponent -1) stays plain. All five are parked as strings by import.js
// parsePreservingNumbers; Number-carriable values stay JS Numbers; booleans/text keep their
// JSON types. (The earlier expanded-decimal spellings were a wrong single-case generalization.)
const L29 = '0.12345678901234567890123456789'
const C1_SWAPPED = '0.99999999999999999999999999999' // form-legal (31 chars, 0.x29) but WRONG value
const D1_WIRE = '1' + '0'.repeat(400) // 1e400 (scale -400) folded -> 401-char integer literal
const E1_WIRE = '1E-400' // 1E-400 (scale 400) -> BigDecimal.toString E-notation original
const F1_WIRE = '4E-324' // 4e-324 (scale 324) -> BigDecimal.toString E-notation original
const G1_WIRE = '1.23456789012345E-310' // (scale 324) -> BigDecimal.toString E-notation original
function fixturePreviewRows(opts) {
  const o = opts || {}
  return [
    { sheetId: 'sheet-0', sheetIndex: 0, rowIndex: 0, cells: {
      '0': { t: 2, v: 123 },
      '1': { t: 2, v: -45.75 },
      '2': { t: 2, v: o.swapC1 ? C1_SWAPPED : L29 },
      '3': { t: 2, v: D1_WIRE },
      '4': { t: 2, v: E1_WIRE },
      '5': { t: 2, v: F1_WIRE },
      '6': { t: 2, v: G1_WIRE }
    } },
    { sheetId: 'sheet-0', sheetIndex: 0, rowIndex: 1, cells: { '0': { t: 3, v: true }, '1': { t: 3, v: false } } },
    { sheetId: 'sheet-0', sheetIndex: 0, rowIndex: 2, cells: { '0': { t: 2, v: 45838.99998842592 } } },
    { sheetId: 'sheet-0', sheetIndex: 0, rowIndex: 3, cells: { '0': { t: 1, v: '001' }, '1': { t: 1, v: '2025-06-30' } } },
    { sheetId: 'sheet-0', sheetIndex: 0, rowIndex: 4, cells: { '0': { t: 2, v: 13, f: 'A1*B1' } } }
  ]
}

// ---------- 三审 P2 收口: real wire → production parser → page assert (offline, no network) ----------
// 正向全链路回归不再手写预览 rows 假设线缆形态，而是构造后端 PRECISE_JSON 通道实际输出的预览封套
// JSON 文本（数值格以原始 JSON number token 在线缆出现；'1E-400' / '4E-324' / '1.23456789012345E-310'
// 是合法 JSON number 文法），交给真实 import.js 解析器，再把生产解析结果喂给页面 queryPreview →
// assertPreviewPrecision，锁定「真实后端预览 JSON → 生产解析器 → 页面断言」。
// 复用 c3-import-api-regression.cjs 的离线装载口径：装载 import.js 之前先按应用启动顺序污染全局
// axios 默认值（request.js:14 的 JSON Content-Type + 离线 recording adapter），axios.create() 快照到
// 的就是与生产一致的默认值链。全程无后端、无网络。
const wireAxios = require(path.join(root, 'node_modules/axios'))
wireAxios.defaults.headers['Content-Type'] = 'application/json;charset=utf-8' // request.js:14 boot order
const wireBodies = []
wireAxios.defaults.adapter = config =>
  Promise.resolve({ data: wireBodies.shift(), status: 200, statusText: 'OK', headers: {}, config, request: {} })
const apiCache = new Map()
let productionApi = null
function productionImportApi() {
  if (productionApi) return productionApi
  const file = path.join(root, 'src/api/ai/import.js')
  const module = { exports: {} }
  apiCache.set(file, module)
  const source = babel.transformSync(fs.readFileSync(file, 'utf8'), { filename: file, babelrc: false, configFile: false, plugins: [plugin] }).code
  const localRequire = name => name === '@/utils/auth' ? { getToken: () => 'c3-ui-regression-stub-token' }
    : name.startsWith('.') ? load(path.resolve(path.dirname(file), name + '.js')) : require(path.join(root, 'node_modules', name))
  new Function('require', 'module', 'exports', source)(localRequire, module, module.exports)
  productionApi = module.exports
  return productionApi
}

// 真实线缆（后端 BigDecimal 序列化契约，ExcelImportService wholeFractionNormalized + BigDecimal.toString，
// Codex 隔离库 + 生产解析器探针实测）：D1=1e400（负 scale -400）折叠为 401 位整数 token；
// E1=1E-400（正 scale 400）→ '1E-400'、F1=4e-324（正 scale 324）→ '4E-324'、
// G1=1.23456789012345e-310（正 scale 324）→ '1.23456789012345E-310'（E 记法原文）；
// C1=29 位字面量（scale 29）plain 原样。其余格同夹具矩阵。
function fixturePreviewWire(opts) {
  const o = opts || {}
  return '{"code":200,"msg":"ok","data":{"jobId":"j-fwd","sourceHash":"' + 'c'.repeat(64) + '","sheetCount":1,' +
    '"totalRowCount":5,"requestedRows":50,"returnedRows":5,"limited":false,"fullSourceCoverage":true,' +
    '"note":"c3-acceptance.xlsx 全量已知 5 行（正向回归离线线缆）","sheets":[],"rows":[' +
    '{"sheetId":"sheet-0","sheetIndex":0,"rowIndex":0,"cells":{' +
    '"0":{"t":2,"v":123},' +
    '"1":{"t":2,"v":-45.75},' +
    '"2":{"t":2,"v":' + (o.swapC1 ? C1_SWAPPED : L29) + '},' +
    '"3":{"t":2,"v":1' + '0'.repeat(400) + '},' +
    '"4":{"t":2,"v":1E-400},' +
    '"5":{"t":2,"v":4E-324},' +
    '"6":{"t":2,"v":1.23456789012345E-310}' +
    '}},' +
    '{"sheetId":"sheet-0","sheetIndex":0,"rowIndex":1,"cells":{"0":{"t":3,"v":true},"1":{"t":3,"v":false}}},' +
    '{"sheetId":"sheet-0","sheetIndex":0,"rowIndex":2,"cells":{"0":{"t":2,"v":45838.99998842592}}},' +
    '{"sheetId":"sheet-0","sheetIndex":0,"rowIndex":3,"cells":{"0":{"t":1,"v":"001"},"1":{"t":1,"v":"2025-06-30"}}},' +
    '{"sheetId":"sheet-0","sheetIndex":0,"rowIndex":4,"cells":{"0":{"t":2,"v":13,"f":"A1*B1"}}}' +
    ']}}'
}

// 生产解析器离线执行：线缆文本进真实 previewImport（axios → recording adapter → transformResponse
// 链），返回 parsePreservingNumbers 的输出对象。
async function parseWireOnProductionParser(wire) {
  wireBodies.length = 0
  wireBodies.push(wire)
  return productionImportApi().previewImport('j-fwd', 'c'.repeat(64), 50)
}

// 前向用例的页面驱动：mock '@/api/ai/import' 的 previewImport 返回生产解析结果，页面方法回归
// 同一装载器（queryPreview 置预览状态后自动执行 assertPreviewPrecision）。
async function feedParsedPreviewToPage(parsed) {
  const p = page()
  importApi.previewImport = async () => parsed
  p.jobId = 'j-fwd'
  p.sourceHash = 'c'.repeat(64)
  p.rowsRequested = 50
  await p.queryPreview()
  return p
}

;(async () => {
  // ================= F1: preview precision must verify VALUES, not form =================
  await test('F1 counter-example: fixture-shaped preview with C1 swapped to a form-legal wrong literal must FAIL', async () => {
    const p = page()
    p.preview = { jobId: 'j-swap', rows: fixturePreviewRows({ swapC1: true }) }
    p.assertPreviewPrecision()
    const a = p.assertions.previewPrecision
    assert.strictEqual(a.result, 'FAIL', 'expected FAIL, got: ' + a.result + ' | ' + a.evidence)
    assert.ok(String(a.evidence).indexOf('0:0:2') >= 0, 'FAIL evidence must name the coordinate 0:0:2, got: ' + a.evidence)
  })

  await test('F1 green control: full fixture preview passes the expectation-table check', async () => {
    const p = page()
    p.preview = { jobId: 'j-full', rows: fixturePreviewRows({}) }
    p.assertPreviewPrecision()
    const a = p.assertions.previewPrecision
    assert.strictEqual(a.result, 'PASS', 'expected PASS, got: ' + a.result + ' | ' + a.evidence)
  })

  await test('F1 non-fixture file: preview without the fixture signature must record 人工核对 (never PASS)', async () => {
    const p = page()
    p.preview = { jobId: 'j-other', rows: [
      { sheetId: 'sheet-0', sheetIndex: 0, rowIndex: 10, cells: { '3': { t: 2, v: C1_SWAPPED } } }
    ] }
    p.assertPreviewPrecision()
    const a = p.assertions.previewPrecision
    assert.notStrictEqual(a.result, 'PASS', 'non-fixture preview must never claim PASS, got: ' + a.evidence)
    assert.strictEqual(a.result, '人工核对', "expected '人工核对', got: " + a.result + ' | ' + a.evidence)
  })

  // ================= F2: Axios transport errors are not business envelopes =================
  await test('F2 counter-example: Axios ERR_BAD_REQUEST with response.status=403 must judge auth-failed and FAIL connectivity', async () => {
    const p = page()
    importApi.getImportJob = async () => {
      throw Object.assign(new Error('Request failed with status code 403'), { code: 'ERR_BAD_REQUEST', response: { status: 403 } })
    }
    await p.runProbe()
    assert.strictEqual(p.probeResult.kind, 'auth-failed', "expected probeResult.kind 'auth-failed', got: " + JSON.stringify(p.probeResult))
    assert.strictEqual(p.assertions.connectivity.result, 'FAIL', 'expected connectivity FAIL, got: ' + p.assertions.connectivity.result + ' | ' + p.assertions.connectivity.evidence)
  })

  await test('F2 control: business envelope code=404 (no response) still means connected+authed OK', async () => {
    const p = page()
    importApi.getImportJob = async () => {
      throw Object.assign(new Error('IMPORT_FORBIDDEN: 任务不存在或仅任务所有者可访问'), { code: 404 })
    }
    await p.runProbe()
    assert.strictEqual(p.probeResult.kind, 'business', "expected kind 'business', got: " + JSON.stringify(p.probeResult))
    assert.strictEqual(p.assertions.connectivity.result, 'PASS', 'expected connectivity PASS, got: ' + p.assertions.connectivity.result + ' | ' + p.assertions.connectivity.evidence)
    assert.ok(String(p.assertions.connectivity.evidence).indexOf('code=404') >= 0, 'evidence must carry code=404, got: ' + p.assertions.connectivity.evidence)
  })

  await test('F2 control: business envelope code=401 (no response) still judges auth-failed', async () => {
    const p = page()
    importApi.getImportJob = async () => {
      throw Object.assign(new Error('认证失败，无法访问系统资源'), { code: 401 })
    }
    await p.runProbe()
    assert.strictEqual(p.probeResult.kind, 'auth-failed', "expected kind 'auth-failed', got: " + JSON.stringify(p.probeResult))
    assert.strictEqual(p.assertions.connectivity.result, 'FAIL', 'expected connectivity FAIL, got: ' + p.assertions.connectivity.result)
  })

  // ================= F4a: an undecidable formula must not sail through in a mixed set =================
  // P2-c (final review): a decidable formula row must also carry the apply/storage evidence the
  // real page produces — recalcApplied + save() snapshot (f verbatim, v === computed). Rows here
  // default to "applied + storage carried the result"; cases that exercise the absence override it.
  const formulaRow = over => ({
    f: over.f, formulaReadback: over.readback === undefined ? over.f : over.readback,
    computed: over.computed, expected: over.expected === undefined ? null : over.expected,
    decidable: over.decidable, match: Boolean(over.match), verdict: over.verdict || '',
    recalcApplied: over.recalcApplied === undefined ? true : over.recalcApplied,
    saveSnapshot: over.saveSnapshot === undefined
      ? { f: over.f, v: over.computed, t: 2 }
      : over.saveSnapshot
  })

  await test('F4a counter-example: one decidable match + one undecidable must record 人工核对 (not PASS)', async () => {
    const p = page()
    p.formulaRows = [
      formulaRow({ f: '=A1*B1', computed: -5627.25, expected: -5627.25, decidable: true, match: true, verdict: 'PASS' }),
      formulaRow({ f: '=SUM(A1:B1)', computed: 7, decidable: false, verdict: '人工核对' })
    ]
    p.assertFormulaCalc()
    const a = p.assertions.formulaCalc
    assert.notStrictEqual(a.result, 'PASS', 'mixed decidable/undecidable set must never PASS, got: ' + a.evidence)
    assert.strictEqual(a.result, '人工核对', "expected '人工核对', got: " + a.result + ' | ' + a.evidence)
  })

  await test('F4a green control: all formulas decidable, f verbatim and equal must PASS', async () => {
    const p = page()
    p.formulaRows = [
      formulaRow({ f: '=A1*B1', computed: -5627.25, expected: -5627.25, decidable: true, match: true, verdict: 'PASS' })
    ]
    p.assertFormulaCalc()
    assert.strictEqual(p.assertions.formulaCalc.result, 'PASS', 'expected PASS, got: ' + p.assertions.formulaCalc.result + ' | ' + p.assertions.formulaCalc.evidence)
  })

  await test('F4a regression: decidable but unequal (or f not verbatim) stays FAIL', async () => {
    const p = page()
    p.formulaRows = [
      formulaRow({ f: '=A1*B1', computed: 13, expected: -5627.25, decidable: true, match: false, verdict: 'FAIL' })
    ]
    p.assertFormulaCalc()
    assert.strictEqual(p.assertions.formulaCalc.result, 'FAIL', 'expected FAIL for mismatch, got: ' + p.assertions.formulaCalc.result)
    const q = page()
    q.formulaRows = [
      formulaRow({ f: '=A1*B1', readback: '=A1 * B1', computed: -5627.25, expected: -5627.25, decidable: true, match: true })
    ]
    q.assertFormulaCalc()
    assert.strictEqual(q.assertions.formulaCalc.result, 'FAIL', 'expected FAIL for f not verbatim, got: ' + q.assertions.formulaCalc.result)
  })

  // ================= F4b: backend reload must verify f verbatim =================
  const cellDataWithF = { 'sheet-0': { '4': { '0': { t: 2, v: 13, f: 'A1*B1' } } } }
  const cellDataLostF = { 'sheet-0': { '4': { '0': { t: 2, v: 13 } } } }

  await test('F4b counter-example: backend reload losing f (cached value unchanged) must FAIL backendSaveReload', async () => {
    const p = page()
    p.backendWorkbookId = 9001
    p.backendBaseline = cellDataWithF
    p.compareBackendCellData(cellDataLostF)
    const a = p.assertions.backendSaveReload
    assert.strictEqual(a.result, 'FAIL', 'expected FAIL when f is lost on reload, got: ' + a.result + ' | ' + a.evidence)
    assert.ok(p.backendReloadSummary.fDiffCount >= 1, 'summary must count the f diff, got: ' + JSON.stringify(p.backendReloadSummary))
    assert.ok(String(a.evidence).indexOf('r4c0') >= 0, 'FAIL evidence must list the f-loss coordinate r4c0, got: ' + a.evidence)
  })

  await test('F4b green control: f preserved verbatim must PASS', async () => {
    const p = page()
    p.backendWorkbookId = 9002
    p.backendBaseline = cellDataWithF
    p.compareBackendCellData({ 'sheet-0': { '4': { '0': { t: 2, v: 13, f: 'A1*B1' } } } })
    const a = p.assertions.backendSaveReload
    assert.strictEqual(a.result, 'PASS', 'expected PASS with f verbatim, got: ' + a.result + ' | ' + a.evidence)
  })

  // ================= 三审 P2 收口: 真实线缆正向全链路 =================
  await test('forward real wire: backend preview envelope → real import.js parser → page assertPreviewPrecision PASS', async () => {
    const parsed = await parseWireOnProductionParser(fixturePreviewWire({}))
    // 生产解析器守卫证据：D1 401 位整数与三个 E 记法 token 必须以 typeof 'string' 的精确原文离开
    // 解析器（import.js 精度守卫对 E 记法 token 不物化 Number）。
    const cells = parsed.rows[0].cells
    const snapshot = JSON.stringify({ D1: cells['3'] && cells['3'].v, E1: cells['4'] && cells['4'].v, F1: cells['5'] && cells['5'].v, G1: cells['6'] && cells['6'].v })
    assert.strictEqual(typeof cells['3'].v, 'string', 'D1 token must stay a string, got: ' + snapshot)
    assert.strictEqual(cells['3'].v, '1' + '0'.repeat(400), 'D1 must be the exact 401-digit integer literal, got: ' + snapshot)
    assert.strictEqual(typeof cells['4'].v, 'string', 'E1 token must stay a string, got: ' + snapshot)
    assert.strictEqual(cells['4'].v, '1E-400', 'E1 must keep the exact E-notation wire text, got: ' + snapshot)
    assert.strictEqual(cells['5'].v, '4E-324', 'F1 must keep the exact E-notation wire text, got: ' + snapshot)
    assert.strictEqual(typeof cells['6'].v, 'string', 'G1 token must stay a string, got: ' + snapshot)
    assert.strictEqual(cells['6'].v, '1.23456789012345E-310', 'G1 must keep the exact E-notation wire text, got: ' + snapshot)
    // 页面断言：生产解析结果逐格与期望表全等（t 全等 + v 逐字符/数值全等）。
    const p = await feedParsedPreviewToPage(parsed)
    const a = p.assertions.previewPrecision
    assert.strictEqual(a.result, 'PASS', 'expected PASS against the fixture expectation table, got: ' + a.result + ' | ' + a.evidence)
  })

  await test('wire counter-example: same real wire with C1 token swapped to the form-legal wrong literal must FAIL on the page', async () => {
    const parsed = await parseWireOnProductionParser(fixturePreviewWire({ swapC1: true }))
    assert.strictEqual(parsed.rows[0].cells['2'].v, C1_SWAPPED, 'swapped C1 token must survive the production parser verbatim')
    const p = await feedParsedPreviewToPage(parsed)
    const a = p.assertions.previewPrecision
    assert.strictEqual(a.result, 'FAIL', 'expected FAIL for the swapped C1 token, got: ' + a.result + ' | ' + a.evidence)
    assert.ok(String(a.evidence).indexOf('0:0:2') >= 0, 'FAIL evidence must name the coordinate 0:0:2, got: ' + a.evidence)
  })

  await test('E-notation materialization guard: production parse keeps every high-precision t=2 cell as string; a Number-materialized copy FAILs', async () => {
    const parsed = await parseWireOnProductionParser(fixturePreviewWire({}))
    // 守卫口径：import.js 解析结果中，凡期望 v 为字符串的 t=2 高精度格（0:0:2..0:0:6），
    // typeof v !== 'string' 的格子列出即 FAIL——生产解析器必须为空列表。
    const PRECISION_COLS = { '2': '0:0:2', '3': '0:0:3', '4': '0:0:4', '5': '0:0:5', '6': '0:0:6' }
    const materialized = []
    for (const [col, coord] of Object.entries(PRECISION_COLS)) {
      const cell = parsed.rows[0].cells[col]
      if (!cell || cell.t !== 2 || typeof cell.v !== 'string') {
        materialized.push(coord + '(t=' + (cell && cell.t) + ', typeof v=' + (cell && typeof cell.v) + ')')
      }
    }
    assert.deepStrictEqual(materialized, [], 'production parser must park every high-precision token as a string; materialized: ' + materialized.join(', '))
    // 反向守卫：任何 token 若被物化为 Number（naive JSON.parse 的默认行为），页面断言必须 FAIL
    // 并列出全部物化坐标。
    const naive = JSON.parse(JSON.stringify(parsed))
    for (const col of Object.keys(PRECISION_COLS)) naive.rows[0].cells[col].v = Number(naive.rows[0].cells[col].v)
    const p = await feedParsedPreviewToPage(naive)
    const a = p.assertions.previewPrecision
    assert.strictEqual(a.result, 'FAIL', 'a Number-materialized preview must FAIL, got: ' + a.result + ' | ' + a.evidence)
    for (const coord of Object.values(PRECISION_COLS)) {
      assert.ok(String(a.evidence).indexOf(coord) >= 0, 'FAIL evidence must list materialized cell ' + coord + ', got: ' + a.evidence)
    }
  })

  // ================= Expectation table: independent BigInt re-derivation =================
  await test('expectation table shape: 13 known fixture cells, D1 literal 401 chars, formula cache recorded on 0:4:0', async () => {
    assert.ok(expectations, 'c3-acceptance-fixture-expectations.js must load, got: ' + (expectationsError && expectationsError.message))
    const table = expectations.FIXTURE_EXPECTATIONS
    // The module is loaded inside the page's vm realm (same instance the page imports), so its
    // object prototypes differ from this realm - JSON-normalize before deepStrictEqual.
    const norm = value => JSON.parse(JSON.stringify(value))
    assert.strictEqual(Object.keys(table).length, 13, 'expected 13 known cells, got: ' + Object.keys(table).length)
    assert.strictEqual(expectations.FIXTURE_D1_JSON_LITERAL_LENGTH, 401, 'D1 literal length constant must be 401')
    assert.deepStrictEqual(norm(table['0:0:0']), { t: 2, v: 123 })
    assert.deepStrictEqual(norm(table['0:1:0']), { t: 3, v: true })
    assert.deepStrictEqual(norm(table['0:1:1']), { t: 3, v: false })
    assert.deepStrictEqual(norm(table['0:3:0']), { t: 1, v: '001' })
    assert.deepStrictEqual(norm(table['0:3:1']), { t: 1, v: '2025-06-30' })
    assert.deepStrictEqual(norm(table['0:4:0']), { t: 2, v: 13, f: 'A1*B1' }, 'fixture formula cache must be recorded')
    assert.strictEqual(table['0:0:3'].v.length, expectations.FIXTURE_D1_JSON_LITERAL_LENGTH, 'D1 expectation must be the 401-char integer literal')
  })

  await test('expectation table: independent BigInt derivation per the real BigDecimal.toString wire rule equals every FIXTURE_EXPECTATIONS value', async () => {
    assert.ok(expectations, 'c3-acceptance-fixture-expectations.js must load, got: ' + (expectationsError && expectationsError.message))
    const table = expectations.FIXTURE_EXPECTATIONS
    // (significand, exponent) pairs read off the fixture cell spellings (value = significand * 10^exp).
    // Nothing here consults an API response: D1=1e400 takes the negative-scale integer-fold path
    // (1n*10n**400n).toString(); E1/F1/G1 take the java.math.BigDecimal.toString grammar
    // (scale≠0 && adjusted exponent < -6 -> 'first digit[.rest]E adjusted'); C1 (scale 29,
    // adjusted -1) and every Number-carriable value take the plain path.
    const DERIVATIONS = {
      '0:0:0': [123n, 0],
      '0:0:1': [-4575n, -2],
      '0:0:2': [12345678901234567890123456789n, -29],
      '0:0:3': [1n, 400],
      '0:0:4': [1n, -400],
      '0:0:5': [4n, -324],
      '0:0:6': [123456789012345n, -324],
      '0:2:0': [4583899998842592n, -11],
      '0:4:0': [13n, 0]
    }
    // Real wire derivation (wholeFractionNormalized, the form the backend actually serializes):
    //   exponent >= 0 (negative scale) -> fold to an integer: unscaled = S × 10^exponent, scale 0,
    //   toString prints the plain integer (D1 path).
    //   exponent < 0 (positive scale) -> unscaled = S, scale = |exponent|; adjusted exponent =
    //   precision - 1 - scale; scale≠0 && adjusted < -6 -> scientific notation with 'E' + adjusted
    //   (E1/F1/G1 path); otherwise the plain decimal split (identical to the former plain-expansion
    //   path within the adjusted >= -6 domain, e.g. C1 and -45.75).
    function deriveBigDecimalToString(significand, exponent) {
      if (exponent >= 0) return (significand * 10n ** BigInt(exponent)).toString()
      const negative = significand < 0n
      const digits = (negative ? -significand : significand).toString()
      const scale = -exponent
      const adjusted = digits.length - 1 - scale
      const sign = negative ? '-' : ''
      if (adjusted < -6) {
        return sign + digits[0] + (digits.length > 1 ? '.' + digits.slice(1) : '') + 'E' + adjusted
      }
      const intLen = digits.length - scale
      const plain = intLen > 0
        ? digits.slice(0, intLen) + (intLen < digits.length ? '.' + digits.slice(intLen) : '')
        : '0.' + '0'.repeat(-intLen) + digits
      return sign + plain
    }
    const problems = []
    for (const [key, pair] of Object.entries(DERIVATIONS)) {
      const derived = deriveBigDecimalToString(pair[0], pair[1])
      const expected = table[key]
      if (!expected) { problems.push(key + ': missing from FIXTURE_EXPECTATIONS'); continue }
      if (typeof expected.v === 'string') {
        if (expected.v !== derived) problems.push(key + ': string expectation ' + JSON.stringify(expected.v) + ' != derived ' + JSON.stringify(derived))
        // The string lane must be NECESSARY: a Number would corrupt the value.
        const n = Number(expected.v)
        if (Number.isFinite(n) && String(n) === expected.v) problems.push(key + ': string lane unnecessary, Number carries ' + expected.v + ' losslessly')
      } else if (typeof expected.v === 'number') {
        const n = Number(derived)
        if (!Number.isFinite(n) || n !== expected.v) problems.push(key + ': number expectation ' + expected.v + ' != derived ' + derived)
      } else {
        problems.push(key + ': unexpected expectation type ' + typeof expected.v)
      }
    }
    assert.strictEqual(problems.length, 0, 'derivation mismatches:\n' + problems.join('\n'))
  })

  console.log(JSON.stringify({ passed, failed, scope: 'OFFLINE_PAGE_METHOD_REGRESSION_NO_BACKEND_NO_NETWORK' }))
  process.exitCode = failed ? 1 : 0
})().catch(e => { console.error(e); process.exitCode = 1 })
