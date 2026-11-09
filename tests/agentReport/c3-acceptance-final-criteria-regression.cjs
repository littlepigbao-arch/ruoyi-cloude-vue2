'use strict'
// C3 acceptance page-method regression — FINAL review criteria (offline, no backend/network).
// Final-review BLOCKED (3 × P2): the acceptance device could still let failing scenarios through.
// Each counter-example below was reproduced against ImportAcceptance.vue before the fix and is
// solidified here so it can never silently regress. Loaded the b10-import-ui-regression way:
// babel <script> extraction + vm + data()/methods binding + require-proxy mocks.
//
//   P2-a BUDGET LEDGER: a budget refusal message alone (no job row) used to PASS. Required now:
//        the server-side ledger row must be FAILED and carry structured failDetail
//        {kind, limit, observed}. Message parsing is diagnostic evidence only.
//   P2-b CONNECTIVITY: a generic business code=500 (e.g. missing table / SQL error) used to PASS
//        as "connected + authed OK". Required now: only an EXPECTED probe rejection
//        (404/409/422 with an IMPORT_* message) proves the link; unknown codes FAIL. 401/403 are
//        auth-failed (already covered by the ui regression).
//   P2-c FORMULA SUMMARY: computed === expected used to PASS even when the recalculation never
//        applied and storage still held the preview cache (recalcApplied=false, saveSnapshot.v=13).
//        Required now: PASS needs f verbatim, recalcApplied, and saveSnapshot.v === computed ===
//        expected for every decidable formula.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert')
const root = path.resolve(__dirname, '../..')
const babel = require(path.join(root, 'node_modules/@babel/core'))
const plugin = require(path.join(root, 'node_modules/@babel/plugin-transform-modules-commonjs'))
const PAGE_FILE = path.join(root, 'src/views/application/agentReport/ImportAcceptance.vue')

let passed = 0, failed = 0
async function test(name, fn) {
  try { await fn(); passed++; console.log('PASS ' + name) } catch (e) { failed++; console.error('FAIL ' + name + ': ' + e.message) }
}

// ---------- controllable API + fetch stubs ----------
const importApi = {}
const fetchState = { ok: true, bytes: 2153 }
const cache = new Map()
function load(file) {
  if (cache.has(file)) return cache.get(file)
  let src = fs.readFileSync(file, 'utf8')
  if (file.endsWith('.vue')) src = src.match(/<script>([\s\S]*?)<\/script>/)[1]
  const module = { exports: {} }
  cache.set(file, module)
  const code = babel.transformSync(src, { configFile: false, babelrc: false, plugins: [plugin] }).code
  vm.runInNewContext(code, {
    module, exports: module.exports, console, performance,
    Blob,
    // the budget self-check fetches the static fixture and wraps it in a File
    File: class File { constructor(parts, name, opts) { this.name = name; this.type = (opts || {}).type; this.parts = parts } },
    fetch: async () => ({ ok: fetchState.ok, status: fetchState.ok ? 200 : 404, blob: async () => ({ size: fetchState.bytes }) }),
    process: { env: { BASE_URL: '/' } },
    sessionStorage: { getItem: () => null, setItem() {}, removeItem() {} },
    setTimeout: () => 0, clearTimeout() {}, setInterval: () => 0, clearInterval() {},
    require: name => {
      if (name === '@/api/ai/import') return new Proxy({}, { get: (_, k) => importApi[k] })
      if (name === '@/api/ai/workbook') return new Proxy({}, {})
      if (name === '@/utils/auth') return { getToken: () => 'c3-final-criteria-stub' }
      if (name.startsWith('.')) {
        const target = path.resolve(path.dirname(file), name + (path.extname(name) ? '' : '.js'))
        return load(target)
      }
      return {}
    }
  }, { filename: file })
  cache.set(file, module.exports)
  return module.exports
}

function page() {
  const c = load(PAGE_FILE).default
  const p = c.data()
  for (const [k, v] of Object.entries(c.methods)) p[k] = v.bind(p)
  p.$message = { warning() {}, success() {}, error() {}, info() {} }
  p.$nextTick = () => Promise.resolve()
  p.$refs = {}
  return p
}

const BUDGET_MSG = 'IMPORT_BUDGET_EXCEEDED: COMPRESSION_RATIO 上限 100，观察值 100'
function budgetError() {
  const e = new Error(BUDGET_MSG)
  e.code = 409
  e.definitive = true
  return e
}

;(async () => {
  // ================= P2-a BUDGET: ledger row is mandatory =================
  await test('BUDGET counter-example: refusal message without a FAILED ledger row must FAIL (message parsing is diagnostic only)', async () => {
    const p = page()
    importApi.createImportJob = async () => { throw budgetError() }
    // ledger lookup finds nothing (e.g. row pruned / wrong requestId / backend restarted)
    importApi.getImportJobByRequest = async () => { throw new Error('IMPORT_FORBIDDEN: 任务不存在') }
    await p.runBudgetRefusalCheck()
    const a = p.assertions.budgetRefusal
    assert.strictEqual(p.budgetResult.ledgerFailed, false, 'stub must not provide a FAILED ledger row')
    assert.strictEqual(a.result, 'FAIL', 'a message-only refusal must NOT pass; got: ' + a.result + ' | ' + a.evidence)
    assert.ok(/账本|ledger|FAILED/.test(a.evidence), 'evidence must explain the missing FAILED ledger row: ' + a.evidence)
  })

  await test('BUDGET positive control: real FAILED ledger row with structured failDetail PASSes', async () => {
    const p = page()
    importApi.createImportJob = async () => { throw budgetError() }
    importApi.getImportJobByRequest = async () => ({
      jobId: 'job-budget-1', status: 'FAILED', requestId: 'req-1',
      failMessage: BUDGET_MSG,
      failDetail: { kind: 'COMPRESSION_RATIO', limit: '100', observed: '100', message: BUDGET_MSG }
    })
    await p.runBudgetRefusalCheck()
    const a = p.assertions.budgetRefusal
    assert.strictEqual(p.budgetResult.status, 'FAILED')
    assert.strictEqual(p.budgetResult.ledgerFailed, true)
    assert.strictEqual(a.result, 'PASS', a.evidence)
  })

  await test('BUDGET negative control: FAILED row but wrong structured facts must FAIL', async () => {
    const p = page()
    importApi.createImportJob = async () => { throw budgetError() }
    importApi.getImportJobByRequest = async () => ({
      jobId: 'job-budget-2', status: 'FAILED',
      failDetail: { kind: 'SOURCE_BYTES', limit: '104857600', observed: '104857601' }
    })
    await p.runBudgetRefusalCheck()
    // Message says COMPRESSION_RATIO but the ledger fact wins; a message/ledger mismatch is not a pass.
    assert.notStrictEqual(p.assertions.budgetRefusal.result, 'PASS',
      'ledger facts must drive the verdict, got PASS: ' + p.assertions.budgetRefusal.evidence)
  })

  // ================= P2-b CONNECTIVITY: unknown 500 is not "authed OK" =================
  await test('CONNECTIVITY counter-example: generic business code=500 must FAIL (not prove connectivity/auth)', async () => {
    const p = page()
    importApi.getImportJob = async () => {
      const e = new Error("PreparedStatementCallback; bad SQL grammar [select * from ai_workbook_import_job where job_id=?]; nested exception is java.sql.SQLSyntaxErrorException: Table 'ry-cloud.ai_workbook_import_job' doesn't exist")
      e.code = 500
      e.definitive = false
      throw e
    }
    await p.runProbe()
    const a = p.assertions.connectivity
    assert.strictEqual(a.result, 'FAIL', 'an unknown server error must NOT pass as connected+authed; got: ' + a.result + ' | ' + a.evidence)
    assert.ok(/未预期|unknown|500/.test(a.evidence), 'evidence must call out the unexpected code: ' + a.evidence)
  })

  await test('CONNECTIVITY positive control: expected probe rejection (IMPORT_FORBIDDEN 409) PASSes', async () => {
    const p = page()
    importApi.getImportJob = async () => {
      const e = new Error('IMPORT_FORBIDDEN: 任务不存在或仅任务所有者可访问')
      e.code = 409
      e.definitive = true
      throw e
    }
    await p.runProbe()
    assert.strictEqual(p.assertions.connectivity.result, 'PASS', p.assertions.connectivity.evidence)
  })

  await test('CONNECTIVITY negative control: expected-shaped message with an unexpected code still FAILs', async () => {
    const p = page()
    importApi.getImportJob = async () => {
      const e = new Error('IMPORT_FORBIDDEN: 任务不存在或仅任务所有者可访问')
      e.code = 501
      throw e
    }
    await p.runProbe()
    assert.strictEqual(p.assertions.connectivity.result, 'FAIL', p.assertions.connectivity.evidence)
  })

  // ================= P2-c FORMULA SUMMARY: apply + storage evidence mandatory =================
  await test('FORMULA counter-example: computed===expected but recalcApplied=false and storage still 13 must FAIL', async () => {
    const p = page()
    p.formulaRows = [{
      sheetId: 'sheet-0', row: 4, col: 0, f: '=A1*B1',
      formulaReadback: '=A1*B1',
      cachedV: 13,
      computed: -5627.25,
      expected: -5627.25,
      decidable: true,
      match: true,
      recalcApplied: false,
      saveSnapshot: { f: '=A1*B1', v: 13, t: 2 }
    }]
    p.assertFormulaCalc()
    const a = p.assertions.formulaCalc
    assert.strictEqual(a.result, 'FAIL', 'missing apply/storage evidence must FAIL; got: ' + a.result + ' | ' + a.evidence)
  })

  await test('FORMULA negative control: recalcApplied but storage v !== computed must FAIL', async () => {
    const p = page()
    p.formulaRows = [{
      sheetId: 'sheet-0', row: 4, col: 0, f: '=A1*B1',
      formulaReadback: '=A1*B1', cachedV: 13, computed: -5627.25, expected: -5627.25,
      decidable: true, match: true, recalcApplied: true,
      saveSnapshot: { f: '=A1*B1', v: 13, t: 2 }
    }]
    p.assertFormulaCalc()
    assert.strictEqual(p.assertions.formulaCalc.result, 'FAIL', p.assertions.formulaCalc.evidence)
  })

  await test('FORMULA negative control: formula text not verbatim on readback must FAIL', async () => {
    const p = page()
    p.formulaRows = [{
      sheetId: 'sheet-0', row: 4, col: 0, f: '=A1*B1',
      formulaReadback: '(不可读)', cachedV: 13, computed: -5627.25, expected: -5627.25,
      decidable: true, match: true, recalcApplied: true,
      saveSnapshot: { f: '=A1*B1', v: -5627.25, t: 2 }
    }]
    p.assertFormulaCalc()
    assert.strictEqual(p.assertions.formulaCalc.result, 'FAIL', p.assertions.formulaCalc.evidence)
  })

  await test('FORMULA positive control: verbatim f + applied + storage===computed===expected PASSes', async () => {
    const p = page()
    p.formulaRows = [{
      sheetId: 'sheet-0', row: 4, col: 0, f: '=A1*B1',
      formulaReadback: '=A1*B1', cachedV: 13, computed: -5627.25, expected: -5627.25,
      decidable: true, match: true, recalcApplied: true,
      saveSnapshot: { f: '=A1*B1', v: -5627.25, t: 2 }
    }]
    p.assertFormulaCalc()
    assert.strictEqual(p.assertions.formulaCalc.result, 'PASS', p.assertions.formulaCalc.evidence)
  })

  console.log(JSON.stringify({ passed, failed, scope: 'OFFLINE_PAGE_METHOD_FINAL_CRITERIA_NO_BACKEND_NO_NETWORK' }))
  process.exitCode = failed ? 1 : 0
})().catch(e => { console.error(e); process.exitCode = 1 })
