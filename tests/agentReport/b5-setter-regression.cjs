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
const { executeActions } = load(path.join(root, 'src/views/application/agentReport/univer-action-executor.js'))
let calls = []
const range = { setBackgroundColor() { calls.push('background') }, setFontColor() { calls.push('fontColor'); throw Error('setter failed') } }
const sheet = { getSheetId: () => 's', getSheetName: () => 'Sheet', getMaxRows: () => 10, getMaxColumns: () => 10, getRange: () => range }
const book = { getSheetBySheetId: () => sheet, getActiveSheet: () => sheet }
const result = executeActions({}, book, [{ schemaVersion: '2.0', type: 'setStyle', target: { sheetId: 's', range: 'A1' }, background: '#ff0000', fontColor: '#000000' }], { continueOnError: false })
assert.deepStrictEqual(calls, ['background', 'fontColor'])
assert.strictEqual(result.summary.ok, 0, 'partial style application must not report success')
assert.strictEqual(result.summary.failed, 1)
assert.match(result.results[0].error, /setter failed/)
console.log('PASS B5 style setter exception is reported as failed')
