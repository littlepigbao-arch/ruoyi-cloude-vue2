'use strict'
// C3 integration-acceptance formula-evidence regression (offline, no backend, no network).
// Guards the pure evaluator src/views/application/agentReport/c3-acceptance-formula.js that the
// ImportAcceptance.vue harness page must use for its formula calculation evidence:
//   - the formula verdict is an EXACT value match against evalSimpleFormula(formula, resolveCell),
//     not the old "computed result is some finite number" assertion that let ANY number (including
//     the deliberately stale cached 13 of fixture cell A5) pass as a false green;
//   - the fixture anchor =A1*B1 with A1=123, B1=-45.75 is exactly -5627.25 (the former case-doc
//     value -5622.75 was an arithmetic error) - both factors are exact doubles and so is the
//     product, so strictEqual on doubles is the intended caliber;
//   - anything outside the simple four-arithmetic whitelist (function names, strings, ranges,
//     missing/non-numeric references, non-finite results) must return null so the page records
//     人工核对 instead of ever guessing a verdict.
// Red phase (recorded before the fix): the module file did not exist (ENOENT) and the page had no
// evalSimpleFormula wiring - both tests below failed. Green phase: module + page wiring land.
const assert = require('assert'), fs = require('fs'), path = require('path'), { createRequire } = require('module')
const root = path.resolve(__dirname, '../..'), rf = createRequire(path.join(root, 'package.json')), babel = rf('@babel/core')
const MODULE_FILE = path.join(root, 'src/views/application/agentReport/c3-acceptance-formula.js')
const PAGE_FILE = path.join(root, 'src/views/application/agentReport/ImportAcceptance.vue')

let passed = 0, failed = 0
async function test(name, fn) { try { await fn(); passed++; console.log('PASS ' + name) } catch (e) { failed++; console.error('FAIL ' + name + ': ' + e.message) } }

function loadFormulaModule() {
  const source = fs.readFileSync(MODULE_FILE, 'utf8')
  const code = babel.transformSync(source, { filename: MODULE_FILE, babelrc: false, configFile: false, plugins: [rf.resolve('@babel/plugin-transform-modules-commonjs')] }).code
  const module = { exports: {} }
  new Function('require', 'module', 'exports', code)(rf, module, module.exports)
  return module.exports
}
// resolveCell stub over a plain ref->value map: unknown refs return null (missing reference).
function cells(map) {
  const seen = []
  const resolve = ref => { seen.push(ref); return Object.prototype.hasOwnProperty.call(map, ref) ? map[ref] : null }
  resolve.seen = seen
  return resolve
}

let formula = null
;(async () => {
  await test('module exposes evalSimpleFormula', async () => {
    formula = loadFormulaModule()
    assert.strictEqual(typeof formula.evalSimpleFormula, 'function', 'evalSimpleFormula must be exported')
  })

  // Red phase: with the module absent the remaining evaluator cases cannot run; the page-wiring
  // anchor is still executed (it must fail red against the unwired page), then the run stops red.
  if (!formula) {
    await test('page wiring: ImportAcceptance.vue drives formula evidence through evalSimpleFormula', async () => {
      const page = fs.readFileSync(PAGE_FILE, 'utf8')
      assert.ok(/evalSimpleFormula\s*\(/.test(page), 'page must call evalSimpleFormula for the formula verdict')
      assert.ok(/c3-acceptance-formula/.test(page), 'page must import the c3-acceptance-formula module')
      assert.ok(!/计算结果为数值/.test(page), 'old isNumber-only false-green wording must be gone')
    })
    console.log(JSON.stringify({ passed, failed, scope: 'OFFLINE_PURE_MODULE_NO_BACKEND_NO_NETWORK' }))
    process.exitCode = 1
    return
  }

  await test('fixture anchor =A1*B1 with {A1:123, B1:-45.75} evaluates to exactly -5627.25', async () => {
    const v = formula.evalSimpleFormula('=A1*B1', cells({ A1: 123, B1: -45.75 }))
    assert.strictEqual(v, -5627.25, `expected exactly -5627.25 (123 x -45.75), got ${v}`)
  })

  await test('whitespace / absolute refs / parentheses / unary minus variants evaluate identically', async () => {
    const map = { A1: 123, B1: -45.75 }
    assert.strictEqual(formula.evalSimpleFormula('= A1 * B1 ', cells(map)), -5627.25)
    assert.strictEqual(formula.evalSimpleFormula('A1*B1', cells(map)), -5627.25, 'missing leading = must still evaluate')
    assert.strictEqual(formula.evalSimpleFormula('=$A$1*$B$1', cells(map)), -5627.25)
    assert.strictEqual(formula.evalSimpleFormula('=A1*(B1+1)', cells(map)), -5504.25)
    assert.strictEqual(formula.evalSimpleFormula('=-A1*B1', cells(map)), 5627.25)
  })

  await test('unsafe or non-simple input returns null (never guessed)', async () => {
    const map = { A1: 123, B1: -45.75 }
    for (const bad of ['=SUM(A1:B1)', '=A1+alert(1)', '=A1:B1', '="x"', '=A1&A1', '=A1>0', '=1/0', '=()', '=A1+', '=A1 B1', '']) {
      assert.strictEqual(formula.evalSimpleFormula(bad, cells(map)), null, `${JSON.stringify(bad)} must be undecidable (null)`)
    }
  })

  await test('missing or non-numeric reference returns null', async () => {
    assert.strictEqual(formula.evalSimpleFormula('=A1*B1', cells({ A1: 123 })), null, 'missing B1 must be undecidable')
    assert.strictEqual(formula.evalSimpleFormula('=A1*B1', cells({ A1: 123, B1: '001' })), null, 'non-numeric reference must be undecidable')
    assert.strictEqual(formula.evalSimpleFormula('=A1*B1', cells({ A1: 123, B1: undefined })), null, 'undefined reference must be undecidable')
  })

  await test('resolver receives normalized refs (absolute $ stripped, uppercased)', async () => {
    const resolve = cells({ A1: 123, B1: -45.75 })
    formula.evalSimpleFormula('=($a$1)*($B$1)', resolve)
    assert.deepStrictEqual(resolve.seen, ['A1', 'B1'])
  })

  await test('page wiring: ImportAcceptance.vue drives formula evidence through evalSimpleFormula', async () => {
    const page = fs.readFileSync(PAGE_FILE, 'utf8')
    assert.ok(/evalSimpleFormula\s*\(/.test(page), 'page must call evalSimpleFormula for the formula verdict')
    assert.ok(/c3-acceptance-formula/.test(page), 'page must import the c3-acceptance-formula module')
    assert.ok(!/计算结果为数值/.test(page), 'old isNumber-only false-green wording must be gone')
  })

  console.log(JSON.stringify({ passed, failed, scope: 'OFFLINE_PURE_MODULE_NO_BACKEND_NO_NETWORK' }))
  process.exitCode = failed ? 1 : 0
})().catch(e => { console.error(e); process.exitCode = 1 })
