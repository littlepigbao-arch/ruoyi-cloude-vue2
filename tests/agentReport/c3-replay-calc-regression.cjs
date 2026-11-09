'use strict'
// C3 replay calc regression (offline, real @univerjs 0.25.1 SDK, no backend, no network).
// Solidifies the two C3 browser-integration BLOCKED items with real-SDK evidence (headless device,
// same engine-level plugin set the browser preset wires: sheets + engine-formula + sheets-formula):
//
//   RED (formula)   - the browser failure mode (2026-11-08-C3联调证据-c3-replay-blocked.json:
//                     viaFacade=true yet computed=13): reading getValues() BEFORE the triggered
//                     recalculation's result is applied returns the stale preview cache. The old
//                     wait (onCalculationResultApplied + numeric poll) is racy: the facade wait can
//                     resolve on the 500ms start-watchdog / first applied session without the
//                     triggered result being applied, and the numeric poll is satisfied instantly
//                     by the stale numeric cache. Reproduced here deterministically.
//   GREEN (formula) - the remediated page sequence (ImportAcceptance.vue collectFormulaEvidence):
//                     executeCalculation -> facade wait as best-effort quiescence gate -> bounded
//                     poll until the formula cell's display value LEAVES the preview cache -> read
//                     computed + save() storage snapshot. Must yield -5627.25 exactly (real engine
//                     calculation, not a filled-in value), f verbatim, and the storage snapshot of
//                     the formula cell must equal the computed value (engine really wrote the
//                     result back into the model).
//   CACHE GUARD     - import semantics unchanged: after page-style replay (setValue({v,t,f})) and
//                     BEFORE the explicit recalculation lane, save().cellData of the formula cell
//                     still carries the stale cache v=13 (import does not recalculate).
//   A3 THREE-POINT  - input 45838.99998842592 vs save().cellData vs getValues(): storage lane keeps
//                     the exact double; the display lane truncates to 15 significant digits. The
//                     readback verdict caliber (storage === preview as the hard criterion;
//                     displayTruncated recorded, never silently relaxed) is locked here.
//   WIRING ANCHORS  - string-feature assertions that ImportAcceptance.vue actually implements the
//                     evidenced sequence (collectFormulaEvidence robust posture + storage-lane
//                     readback + three-point display), in the existing anchor style.
const assert = require('assert'), fs = require('fs'), path = require('path'), { createRequire } = require('module')
const root = path.resolve(__dirname, '../..'), rf = createRequire(path.join(root, 'package.json'))
const core = rf('@univerjs/core/lib/cjs/index.js'), sheets = rf('@univerjs/sheets/lib/cjs/index.js')
rf('@univerjs/sheets/lib/cjs/facade.js')
const engineFormula = rf('@univerjs/engine-formula/lib/cjs/index.js')
const sheetsFormula = rf('@univerjs/sheets-formula/lib/cjs/index.js')
rf('@univerjs/engine-formula/lib/cjs/facade.js')
rf('@univerjs/sheets-formula/lib/cjs/facade.js')
const { FUniver } = rf('@univerjs/core/lib/cjs/facade.js')

const EXPECTED = -5627.25 // 123 * -45.75, the exact double both the engine and evalSimpleFormula produce
const A3_INPUT = 45838.99998842592

let passed = 0, failed = 0
async function test(name, fn) { try { await fn(); passed++; console.log('PASS ' + name) } catch (e) { failed++; console.error('FAIL ' + name + ': ' + e.message) } }
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms))

// Same replay shape as ImportAcceptance.vue replayToWorkbook for the fixture's formula-bearing rows.
function buildReplayed() {
  const univer = new core.Univer({ locale: core.LocaleType.EN_US, locales: { [core.LocaleType.EN_US]: {} } })
  univer.registerPlugin(sheets.UniverSheetsPlugin)
  univer.registerPlugin(engineFormula.UniverFormulaEnginePlugin)
  univer.registerPlugin(sheetsFormula.UniverSheetsFormulaPlugin)
  univer.createUnit(core.UniverInstanceType.UNIVER_SHEET, { id: 'c3', name: 'C3', sheetOrder: ['s'], sheets: { s: { id: 's', name: 'T', rowCount: 10, columnCount: 10 } } })
  const api = FUniver.newAPI(univer), book = api.getActiveWorkbook(), sheet = book.getActiveSheet()
  // page replay: A1=123 t2, B1=-45.75 t2, A3 serial t2, texts, A5 {v:13,t:2,f:'=A1*B1'}
  sheet.getRange(0, 0).setValue({ v: 123, t: 2 })
  sheet.getRange(0, 1).setValue({ v: -45.75, t: 2 })
  sheet.getRange(2, 0).setValue({ v: A3_INPUT, t: 2 })
  sheet.getRange(3, 0).setValue({ v: '001', t: 1 })
  sheet.getRange(4, 0).setValue({ v: 13, t: 2, f: '=A1*B1' })
  return { univer, api, book, sheet }
}
const snapCell = (book, row, col) => {
  const cd = book.save().sheets.s.cellData
  const r = cd && cd[row]
  return r ? r[col] : undefined
}

;(async () => {
  // ============ RED: stale cache is what a too-early read sees (browser failure mode) ============
  await test('RED formula: reading getValues() right after executeCalculation (before apply) returns the stale cache 13 — the browser FAIL mechanism', async () => {
    const { univer, api, sheet } = buildReplayed()
    try {
      api.getFormula().executeCalculation()
      const tooEarly = sheet.getRange(4, 0).getValues()[0][0]
      assert.strictEqual(tooEarly, 13, 'pre-apply read must expose the stale preview cache 13, got: ' + tooEarly)
      assert.strictEqual(tooEarly === EXPECTED, false, 'pre-apply read must NOT equal the real result yet')
    } finally { univer.dispose() }
  })

  // ============ GREEN: the remediated page sequence computes -5627.25 for real ============
  await test('GREEN formula: page remediated sequence yields -5627.25 exactly, f verbatim, storage snapshot equals computed (real engine write-back)', async () => {
    const { univer, api, book, sheet } = buildReplayed()
    try {
      const fe = api.getFormula()
      fe.executeCalculation()
      try { await fe.onCalculationResultApplied(5000) } catch (e) { /* quiescence gate is best-effort; value poll below is the real gate */ }
      // Page posture: bounded poll until the display value leaves the preview cache (<=5s).
      const deadline = Date.now() + 5000
      let computed = sheet.getRange(4, 0).getValues()[0][0]
      let recalcApplied = false
      while (Date.now() < deadline) {
        computed = sheet.getRange(4, 0).getValues()[0][0]
        if (computed !== 13) { recalcApplied = true; break }
        await sleep(100)
      }
      assert.strictEqual(recalcApplied, true, 'display value must leave the stale cache within 5s, still: ' + computed)
      assert.strictEqual(computed, EXPECTED, 'computed must be exactly -5627.25 (double ===), got: ' + computed)
      assert.strictEqual(sheet.getRange(4, 0).getFormula(), '=A1*B1', 'f must read back verbatim')
      const stored = snapCell(book, 4, 0)
      assert.ok(stored, 'formula cell must exist in save().cellData')
      assert.strictEqual(stored.f, '=A1*B1', 'stored f must be verbatim')
      assert.strictEqual(stored.v, EXPECTED, 'storage snapshot v must equal the computed result (engine wrote the model), got: ' + stored.v)
    } finally { univer.dispose() }
  })

  // ============ PAGE POSTURE (2026-11-09 browser-verified): register via setFormula, cache kept before ============
  await test('PAGE POSTURE: replay writes setValue({v,t,f}) keeping v=13, then the calc lane registers via setFormula and yields -5627.25', async () => {
    const { univer, api, book, sheet } = buildReplayed()
    try {
      // 1) Replay posture: the cell still carries the preview cache (13) BEFORE any registration.
      const preCalc = snapCell(book, 4, 0)
      assert.ok(preCalc, 'formula cell must exist after replay')
      assert.strictEqual(preCalc.v, 13, 'replay must keep the source cache 13 (import does not recalculate), got: ' + preCalc.v)
      assert.strictEqual(preCalc.f, '=A1*B1', 'replay must keep f verbatim')
      // 2) Calc-lane registration: range.setFormula(f) is the write path the UI plugin stack honors.
      const range = sheet.getRange(4, 0)
      assert.strictEqual(typeof range.setFormula, 'function', 'FRange.setFormula must exist')
      range.setFormula('=A1*B1')
      const fe = api.getFormula()
      fe.executeCalculation()
      try { await fe.onCalculationResultApplied(5000) } catch (e) { /* best-effort gate; value poll below is the real gate */ }
      const deadline = Date.now() + 5000
      let computed = range.getValues()[0][0]
      let applied = false
      while (Date.now() < deadline) {
        const v = range.getValues()[0][0]
        if (typeof v === 'number' && Number.isFinite(v) && v !== 13) { computed = v; applied = true; break }
        await sleep(100)
      }
      assert.strictEqual(applied, true, 'a finite result different from the cache must land within 5s, still: ' + computed)
      assert.strictEqual(computed, EXPECTED, 'computed must be exactly -5627.25, got: ' + computed)
      const stored = snapCell(book, 4, 0)
      assert.strictEqual(stored.v, EXPECTED, 'storage must carry the real result, got: ' + stored.v)
      assert.strictEqual(stored.f, '=A1*B1', 'stored f must stay verbatim')
    } finally { univer.dispose() }
  })

  // ============ CACHE GUARD: replay keeps the stale cache until the explicit recalc lane ============
  await test('CACHE GUARD: after page-style replay and BEFORE any recalculation, save().cellData of the formula cell still carries stale v=13 (import does not recalculate)', async () => {
    const { univer, book, sheet } = buildReplayed()
    try {
      const stored = snapCell(book, 4, 0)
      assert.ok(stored, 'formula cell must exist right after replay')
      assert.strictEqual(stored.v, 13, 'cached v must still be 13 before the explicit recalc lane, got: ' + stored.v)
      assert.strictEqual(stored.f, '=A1*B1', 'f must be stored verbatim')
      assert.strictEqual(sheet.getRange(4, 0).getValues()[0][0], 13, 'display lane must also still show the cache before recalc')
    } finally { univer.dispose() }
  })

  // ============ A3 THREE-POINT: storage lane lossless, display lane truncated (flagged) ============
  await test('A3 three-point: input === save() storage (exact double); getValues() display truncates to 15 significant digits and is flagged displayTruncated; storage lane intact', async () => {
    const { univer, book, sheet } = buildReplayed()
    try {
      const stored = snapCell(book, 2, 0)
      const display = sheet.getRange(2, 0).getValues()[0][0]
      assert.ok(stored, 'A3 must exist in save().cellData')
      assert.strictEqual(stored.t, 2, 'A3 must stay numeric t=2 in storage')
      assert.strictEqual(stored.v, A3_INPUT, 'storage must keep the exact double 45838.99998842592, got: ' + stored.v)
      assert.strictEqual(display, 45838.9999884259, 'display lane must show the 15-digit truncation 45838.9999884259, got: ' + display)
      assert.strictEqual(display === A3_INPUT, false, 'display must differ from the input (truncation observable)')
      // The readback-row evidence shape the page produces: displayTruncated is truthful labeling.
      const readbackRow = { v: A3_INPUT, readback: display, saveSnapshot: stored, displayTruncated: display !== A3_INPUT }
      assert.strictEqual(readbackRow.displayTruncated, true, 'displayTruncated flag must be set when display !== input')
      assert.strictEqual(readbackRow.saveSnapshot.v === readbackRow.v, true, 'hard criterion (storage === preview) must hold')
    } finally { univer.dispose() }
  })

  // ============ WIRING ANCHORS: the page really implements the evidenced sequence ============
  const pageSrc = fs.readFileSync(path.join(root, 'src/views/application/agentReport/ImportAcceptance.vue'), 'utf8')
  const collectFormulaEvidenceSrc = (() => {
    const m = pageSrc.match(/async collectFormulaEvidence\(targets, sheetOf\) \{([\s\S]*?)\n    \},\n/)
    return m ? m[1] : ''
  })()
  const collectReadbackEvidenceSrc = (() => {
    const m = pageSrc.match(/collectReadbackEvidence\(targets, sheetOf\) \{([\s\S]*?)\n    \},\n/)
    return m ? m[1] : ''
  })()

  await test('ANCHOR collectFormulaEvidence: register via setFormula + trigger + quiescence gate + leave-cache poll + storage snapshot', async () => {
    assert.ok(collectFormulaEvidenceSrc.length > 0, 'collectFormulaEvidence body must be locatable')
    // 2026-11-09 real-browser evidence: setValue({v,t,f}) does NOT register the formula in the UI
    // plugin stack (executeCalculation idles, readback stays at the cache); range.setFormula(f)
    // registers it and the engine really recalculates. The page must use setFormula.
    assert.ok(collectFormulaEvidenceSrc.includes('setFormula'), 'must register via range.setFormula (UI stack ignores setValue({f}))')
    assert.ok(collectFormulaEvidenceSrc.includes('preCalcSnapshot'), 'must capture the pre-recalc storage snapshot')
    assert.ok(collectFormulaEvidenceSrc.includes('executeCalculation'), 'must trigger via facade executeCalculation')
    assert.ok(collectFormulaEvidenceSrc.includes('onCalculationResultApplied'), 'must keep the facade wait as a quiescence gate')
    assert.ok(collectFormulaEvidenceSrc.includes('leftCache'), 'must poll until the display value leaves the preview cache')
    assert.ok(collectFormulaEvidenceSrc.includes('saveSnapshot'), 'must take a save() storage snapshot')
    assert.ok(collectFormulaEvidenceSrc.includes('recalcApplied'), 'must record recalcApplied evidence')
    assert.ok(collectFormulaEvidenceSrc.includes('cachedV'), 'must record the preview cached value')
    // The defeated numeric poll must be gone: no isNumeric-based immediate break remains.
    assert.ok(!collectFormulaEvidenceSrc.includes('isNumeric'), 'the defeated isNumeric poll must be removed')
    assert.ok(/val !== fc\.cell\.v/.test(collectFormulaEvidenceSrc), 'the leave-cache gate must compare against the preview cached value')
    // setFormula clears the cache before recalculating: an empty/transient value must not count as applied.
    assert.ok(/typeof val === 'number' && Number\.isFinite\(val\) && val !== fc\.cell\.v/.test(collectFormulaEvidenceSrc), 'leave-cache gate must require a finite numeric result')
  })

  await test('ANCHOR collectReadbackEvidence: numeric verdict is storage-lane (saveSnapshot === preview), display truncation recorded not FAILed', async () => {
    assert.ok(collectReadbackEvidenceSrc.length > 0, 'collectReadbackEvidence body must be locatable')
    assert.ok(collectReadbackEvidenceSrc.includes('saveSnapshot'), 'must capture the save() storage snapshot')
    assert.ok(collectReadbackEvidenceSrc.includes('displayTruncated'), 'must flag display truncation truthfully')
    assert.ok(collectReadbackEvidenceSrc.includes('saveSnapshot.t === 2 && saveSnapshot.v === target.cell.v'), 'numeric hard criterion must be storage === preview')
    // Never silently relaxed: a storage mismatch still produces FAIL (verdict ternary present).
    assert.ok(/'FAIL'/.test(collectReadbackEvidenceSrc), 'storage mismatch must still verdict FAIL')
  })

  await test('ANCHOR page header: empirical basis documented (evidence files + storage-vs-display caliber)', async () => {
    assert.ok(pageSrc.includes('evidence-formula-sequences.json'), 'formula wait posture must cite the sequence evidence')
    assert.ok(pageSrc.includes('evidence-a3-three-point.json'), 'readback caliber must cite the three-point evidence')
    assert.ok(pageSrc.includes('displayTruncated'), 'page must mention the displayTruncated labeling')
  })

  console.log(JSON.stringify({ passed, failed, scope: 'OFFLINE_REAL_UNIVER_SDK_NO_BACKEND_NO_NETWORK' }))
  process.exitCode = failed ? 1 : 0
})().catch(e => { console.error(e); process.exitCode = 1 })
