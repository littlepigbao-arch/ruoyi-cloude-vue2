'use strict'
// B-REVIEW-02: tests/agentReport 统一回归启动器（自动门禁入口）
// 用 child_process.spawnSync 顺序执行本目录全部 *-regression.cjs（不递归子目录，
// 因此 fixtures/ 等非回归文件天然跳过），逐行输出脚本名与退出码，末尾汇总 N/M 通过；
// 任一脚本退出码非 0，本进程退出码为 1；全部通过退出码为 0。
// 用法: node tests/agentReport/run-all.cjs  或  npm run test:agent:all
const fs = require('fs')
const path = require('path')
const { spawnSync } = require('child_process')

const dir = __dirname
const scripts = fs.readdirSync(dir)
  .filter(name => /-regression\.cjs$/.test(name))
  .sort()

const results = []
for (const name of scripts) {
  const r = spawnSync(process.execPath, [path.join(dir, name)], { stdio: 'inherit' })
  if (r.error) console.error(r.error)
  const code = r.status === null ? 1 : r.status
  const signal = r.signal ? ` signal=${r.signal}` : ''
  const ok = code === 0
  results.push(ok)
  console.log(`[${ok ? 'PASS' : 'FAIL'}] ${name} exit=${code}${signal}`)
}

const passed = results.filter(Boolean).length
console.log(`agentReport regression summary: ${passed}/${results.length} passed`)
if (results.length === 0) {
  console.error('no *-regression.cjs scripts found under tests/agentReport')
  process.exit(1)
}
process.exit(passed === results.length ? 0 : 1)
