'use strict'
// C3 acceptance fixture expectation table (pure data module, no Vue/SDK/HTTP dependencies).
//
// 为什么需要独立期望表（F1 整改）：验收专页 assertPreviewPrecision 旧实现只做形态检查
// （length===31 && /^0\.\d{29}$/），把正确字面量换成 0.99999999999999999999999999999 仍 PASS
// ——形态合法但值错误。整改后页面以本表对预览逐格核对；每个期望值按下列真实序列化规则
// 数学推导得出，独立于任何 API 响应，并在 tests/agentReport/c3-acceptance-ui-regression.cjs 中
// 用 BigInt 从 (有效数字, 指数) 现场重推互证，且由「真实后端预览 JSON → 生产解析器 → 页面断言」
// 正向用例全链路锁定（真实线缆 → 真实 import.js 解析 → 页面逐格比对 PASS）。
//
// 序列化规则（Codex 隔离库 + 生产解析器探针实测，ExcelImportService wholeFractionNormalized +
// java.math.BigDecimal.toString；禁止把单一实测推广为全局 toPlainString 展开小数）：
//   值 = 有效数字 S × 10^指数 E。
//   E >= 0（负 scale）：wholeFractionNormalized 先把负 scale 归零折叠为整数再 toString → plain 整数。
//           夹具 D1=1e400（scale -400）→ '1'+400 个'0'（401 位整数字面量，
//           FIXTURE_D1_JSON_LITERAL_LENGTH=401 实测）；import.js parsePreservingNumbers 将该
//           token 停放为字符串（Number 会溢出为 Infinity）。
//   E <  0（正 scale）：scale 保留，toString 按 adjusted exponent = precision-1-scale 判定文法：
//           scale≠0 且 adjusted < -6 → 科学记法 '首数字[.余数字]E adjusted'（E 记法原文逐字符保留）：
//           E1=1E-400 → '1E-400'、F1=4e-324 → '4E-324'、G1=1.23456789012345e-310 → '1.23456789012345E-310'；
//           否则 → plain 小数：C1=0.12345678901234567890123456789（scale 29、adjusted -1）逐字符
//           29 位原样，无 E 记法。
// Number 可无损承载的值（123 / -45.75 / 45838.99998842592 / 13）在预览中仍为 JS Number；
// Number 会改变值的字面量（29 位小数、上溢/下溢/亚正常）以字符串逐字符承载（F2 契约）。
// 布尔（t=3）与文本（t=1）不经数值通道，保持 JSON 原生类型。
//
// 坐标键 '{sheetIndex}:{row}:{col}'（0 基，对应预览 rows[].sheetIndex / rowIndex / cells 列键）；
// 值 { t, v, f? }：t 为预览单元格类型（1=STRING / 2=NUMBER / 3=BOOLEAN）；v 为期望值——
// 页面比对口径为 t 类型全等 + v 逐字符全等（字符串）或数值/布尔全等；f 仅作夹具公式缓存记录
// （SAX 不重算口径：A5 公式 =A1*B1 的缓存 13 原样保留，公式计算证据由 Univer 重放阶段另行判定）。

// 夹具 D1=1e400 的 JSON 字面量长度（负 scale -400 经 wholeFractionNormalized 归零折叠 →
// '1'+400 个'0'，Codex 隔离库实测）。
export const FIXTURE_D1_JSON_LITERAL_LENGTH = 401

// c3-acceptance.xlsx（单 sheet "T"，sheetIndex=0）全部已知单元格的期望表。
export const FIXTURE_EXPECTATIONS = {
  // 第 1 行：数字 + 29 位保精度字面量 + 上溢/下溢/亚正常边界
  '0:0:0': { t: 2, v: 123 },
  '0:0:1': { t: 2, v: -45.75 },
  '0:0:2': { t: 2, v: '0.12345678901234567890123456789' }, // C1=29 位字面量（scale 29、adjusted -1）plain 原样，Number 必损 → 字符串
  '0:0:3': { t: 2, v: '1' + '0'.repeat(400) }, // D1=1e400 → 负 scale -400 归零折叠：401 位整数字面量（字符串）
  '0:0:4': { t: 2, v: '1E-400' }, // E1=1E-400 → 正 scale 400、adjusted -400 < -6 → BigDecimal.toString E 记法原文（字符串）
  '0:0:5': { t: 2, v: '4E-324' }, // F1=4e-324 → 正 scale 324、adjusted -324 < -6 → E 记法原文（字符串）
  '0:0:6': { t: 2, v: '1.23456789012345E-310' }, // G1 → 正 scale 324、adjusted -310 < -6 → E 记法原文（字符串）
  // 第 2 行：布尔（t=3，JSON true/false 原生类型）
  '0:1:0': { t: 3, v: true },
  '0:1:1': { t: 3, v: false },
  // 第 3 行：日期序列值（2025-06-30 23:59:59，1900 基准，数值通道无损）
  '0:2:0': { t: 2, v: 45838.99998842592 },
  // 第 4 行：文本（t=1，前导零 / ISO 形态字符串不强转）
  '0:3:0': { t: 1, v: '001' },
  '0:3:1': { t: 1, v: '2025-06-30' },
  // 第 5 行：公式缓存（SAX 不重算：v=13 原样；f 仅记录，比对口径为 t/v）
  '0:4:0': { t: 2, v: 13, f: 'A1*B1' }
}
