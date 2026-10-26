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
    { module: m, exports: m.exports, console, performance, require: name => require('./clear-content-sdk-fixture.cjs').modules[name] || load(path.resolve(path.dirname(file), name + '.js')) })
  cache.set(file, m.exports)
  return m.exports
}
const api = load(path.join(root, 'src/views/application/agentReport/controlled-write-executor.js'))
const call = { callId: 'b5', claimToken: 'claim', argsHash: 'hash', fencingToken: 9 }
async function run() {
  for (const type of ['setStyle', 'setRowHeight', 'setColumnWidth']) assert.throws(() => api.classifyAiActions([{ type }]), /AI_CAPABILITY_BLOCKED/)
  let attempts = 0
  const receipts = []
  const workspace = { applyControlled: async () => { attempts++; throw Error('RESULT_UNKNOWN: local readback failed') } }
  const paused = await api.submitControlledCall(workspace, call, 'run', async (_endpoint, receipt) => {
    receipts.push(receipt)
    return { status: 'PAUSED', persistenceStatus: 'NOT_SAVED' }
  }, () => true)
  assert.strictEqual(attempts, 1)
  assert.strictEqual(receipts[0].ok, false)
  assert.strictEqual(paused.status, 'PAUSED')
  assert.strictEqual(paused.persistenceStatus, 'NOT_SAVED')
  const falseSuccess = { status: 'COMPLETED', persistenceStatus: 'SAVED', lastControlledResult: { callId: 'b5', persistenceStatus: 'SAVED' } }
  await assert.rejects(() => api.submitControlledCall(workspace, call, 'run', async () => falseSuccess, () => true), error => {
    assert.match(error.message, /RESULT_UNKNOWN/)
    assert.strictEqual(error.run, undefined, 'a conflicting saved response must not promote the UI run')
    return true
  })
  assert.strictEqual(attempts, 2, 'failure protection must never re-execute within a call')
  let requests = 0
  await assert.rejects(() => api.submitControlledCall(workspace, call, 'run', async endpoint => {
    requests++
    if (endpoint.endsWith('/controlled-results')) throw Error('receipt response lost')
    return falseSuccess
  }, () => true), /RESULT_UNKNOWN/)
  assert.strictEqual(attempts, 3)
  assert.strictEqual(requests, 3, 'only the identical receipt may be retried before uncertain lookup')
  console.log('PASS B5 unknown failure receipt is never promoted to saved or replayed')
}
run().catch(error => { console.error(error); process.exitCode = 1 })
