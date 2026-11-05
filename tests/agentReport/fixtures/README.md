# tests/agentReport/fixtures

本目录存放 agentReport 回归测试使用的确定性静态夹具，保证干净检出仓库后测试可直接复现，
不依赖仓库外任何文件。

## b8c-sheet-20261028-sdk-probe.json

- 来源：真实 Univer SDK 0.25.1（仓库 `package.json` dependencies 安装的
  `@univerjs/core` / `@univerjs/sheets`）通过原生 `FWorkbook.insertSheet`
  （显式 sheet id `ai-sheet-fixed`，index 1）执行 createSheet 后的 workbook
  快照（`book.save()` 的 before/after）。
- 用途：`tests/agentReport/b8c-create-sheet-sdk-regression.cjs` 与
  `tests/agentReport/b9-result-sheet-sdk-regression.cjs` 以其中的
  `after.sheets['ai-sheet-fixed']` 作为新建 Sheet 的预期结构模板。
- 确定性：快照不含时间戳、随机 id 等易变字段，同一 SDK 版本下重复生成结果字节一致。
- 探测过程直接调用 SDK 原生 API，不经过被测的
  `src/views/application/agentReport/controlled-write-executor.js`（避免用被测物生成预期）。

## generate-b8c-sdk-probe.cjs（备源生成器）

生成命令（在仓库根目录执行）：

```
node tests/agentReport/fixtures/generate-b8c-sdk-probe.cjs
```

生成器会用当前安装版 Univer SDK 重新探测并覆写上述 JSON。回归测试以仓库内静态
JSON 为准，本生成器仅用于 SDK 升级时再校验产物结构是否变化：重新生成后请 diff
新旧 JSON，确认差异是否为 SDK 升级带来的预期变化，再提交更新后的夹具。
