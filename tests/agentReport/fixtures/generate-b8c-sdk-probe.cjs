'use strict'
// 备源生成器：用真实安装版 Univer SDK 重新探测 createSheet 产物快照，
// 生成 tests/agentReport/fixtures/b8c-sheet-20261028-sdk-probe.json，
// 供 tests/agentReport/b8c-create-sheet-sdk-regression.cjs 与
// tests/agentReport/b9-result-sheet-sdk-regression.cjs 作为预期模板。
//
// 生成命令（在仓库根目录执行）:
//   node tests/agentReport/fixtures/generate-b8c-sdk-probe.cjs
//   （可选参数: 显式指定 ruoyi-cloude-vue2 仓库根目录，默认取本文件上三级目录）
//
// 依赖 SDK 版本: Univer 0.25.1（仓库 package.json dependencies 中安装的
//   @univerjs/core 与 @univerjs/sheets）
//
// 注意: 回归测试以仓库内静态 JSON 夹具为准，本生成器仅用于 SDK 升级时再校验。
//   生成过程直接调用 SDK 原生 FWorkbook.insertSheet（与 controlled-write-executor.js
//   的受控 createSheet 调用方式一致），绝不经过被测的 controlled-write-executor.js，
//   避免用被测物生成预期。
const fs = require('fs'), path = require('path'), { createRequire } = require('module')
const root = path.resolve(process.argv[2] || path.join(__dirname, '..', '..', '..'))
const rf = createRequire(path.join(root, 'package.json'))
const core = rf('@univerjs/core/lib/cjs/index.js'), sheets = rf('@univerjs/sheets/lib/cjs/index.js')
rf('@univerjs/sheets/lib/cjs/facade.js')
const { FUniver } = rf('@univerjs/core/lib/cjs/facade.js')

const u = new core.Univer({ locale: core.LocaleType.EN_US, locales: { [core.LocaleType.EN_US]: {} } })
u.registerPlugin(sheets.UniverSheetsPlugin)
u.createUnit(core.UniverInstanceType.UNIVER_SHEET, {
  id: 'p', name: 'P', sheetOrder: ['s'], activeSheet: 's',
  sheets: { s: { id: 's', name: 'Source', rowCount: 6, columnCount: 6, cellData: { 0: { 0: { v: 'a', t: 1 } } }, mergeData: [] } }
})
const api = FUniver.newAPI(u), book = api.getActiveWorkbook()
const before = book.save()
const created = book.insertSheet('Generated', { index: 1, sheet: { id: 'ai-sheet-fixed', rowCount: 6, columnCount: 6 } })
if (!created || created.getSheetId() !== 'ai-sheet-fixed') throw new Error('probe insertSheet id mismatch')
const after = book.save()
u.dispose()

fs.writeFileSync(path.join(__dirname, 'b8c-sheet-20261028-sdk-probe.json'), JSON.stringify([{ op: 'createSheet', before, after }], null, 2))
console.log('PASS probe written; new sheet keys:', JSON.stringify(Object.keys(after.sheets['ai-sheet-fixed'])))
