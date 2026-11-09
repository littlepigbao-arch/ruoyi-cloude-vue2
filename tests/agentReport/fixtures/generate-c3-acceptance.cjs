'use strict'
// C3 集成验收夹具生成器（离线，无网络无后端）。产出两个 XLSX：
//   c3-acceptance.xlsx      —— 类型/精度/公式全矩阵正样（普通 Excel 无法构造 29 位字面量：
//                              Excel 数值单元格是 float64，必须像后端 ExcelImportTypeRoundTripTest
//                              的 getCTCell().setV(literal) 一样向 sheet XML 注入原始 <v> 文本）
//   c3-budget-refusal.xlsx  —— sheet1.xml 为约 200KB 高重复字节流（全 '0'，包成合法 XML 内联文本），
//                              DEFLATE 后逐条目压缩比 expanded/computed 必 > 100，预期结构化
//                              IMPORT_BUDGET_EXCEEDED kind=COMPRESSION_RATIO、limit=100（POI 默认
//                              minInflateRatio=0.01 → perEntryLimit=100，ExcelImportService.auditZipEntries
//                              在读入循环内逐条目拒绝，先于 POI）。
//                              注意：默认配置下不存在 65536 的条目大小上限——ZipSecureFile.getMaxEntrySize()
//                              默认 4GB，此前"展开 70049 字节触发 ENTRY_SIZE"的口径在默认配置下并不成立。
// 用法: node tests/agentReport/fixtures/generate-c3-acceptance.cjs
// 生成物由用户在 C3 验收专页（/agent-report-import-acceptance）上传，作为验收证据留档。
const fs = require('fs')
const path = require('path')
const { createRequire } = require('module')
const root = path.resolve(__dirname, '../..')
const rf = createRequire(path.join(root, 'package.json'))
const XLSX = rf('xlsx')
const JSZip = rf('jszip')

const LITERAL29 = '0.12345678901234567890123456789'
// 覆盖 Codex 二轮复审数值边界反例：上溢/下溢/亚正常/15 位亚正常舍入
const EDGE = [
  ['D1', '1e400', '上溢 → Infinity'],
  ['E1', '1E-400', '下溢 → 0'],
  ['F1', '4e-324', '亚正常最近邻舍入 → 5e-324'],
  ['G1', '1.23456789012345e-310', '15 位有效数字亚正常舍入']
]

async function buildAcceptance() {
  // 先用 xlsx 库产出结构合法的基座（数字/布尔/日期/文本/公式各就各位，边缘值先用占位数），
  // 再用 jszip 重写 sheet XML，把占位 <v> 换成保精度原始文本。
  const wb = XLSX.utils.book_new()
  const ws = {}
  const put = (addr, cell) => { ws[addr] = cell }
  put('A1', { t: 'n', v: 123 })
  put('B1', { t: 'n', v: -45.75 })
  put('C1', { t: 'n', v: 0.5 }) // 占位，随后注入 29 位原始字面量
  EDGE.forEach(([addr], i) => { put(addr, { t: 'n', v: i + 1 }) }) // 占位
  put('A2', { t: 'b', v: true })
  put('B2', { t: 'b', v: false })
  put('A3', { t: 'n', v: 45838.99998842592 }) // 2025-06-30 23:59:59 的 1900 基准序列值
  put('A4', { t: 's', v: '001' })
  put('B4', { t: 's', v: '2025-06-30' })
  put('A5', { t: 'n', v: 13, f: 'A1*B1' }) // 公式 + 缓存值（SAX 不重算口径：缓存 13 原样保留）
  ws['!ref'] = 'A1:G5'
  XLSX.utils.book_append_sheet(wb, ws, 'T')
  const base = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })

  const zip = await JSZip.loadAsync(base)
  const sheetName = Object.keys(zip.files).find(n => /^xl\/worksheets\/sheet1\.xml$/.test(n))
  if (!sheetName) throw new Error('sheet1.xml not found in generated base workbook')
  let xml = await zip.file(sheetName).async('string')
  const patches = [['C1', LITERAL29]].concat(EDGE.map(([addr, lit]) => [addr, lit]))
  for (const [addr, literal] of patches) {
    // 仅替换该单元格的 <v> 数值文本；t 属性保持数字形态（与后端测试 setV 注入口径一致）
    const re = new RegExp('(<c r="' + addr + '"[^>]*>)<v>[^<]*</v>')
    if (!re.test(xml)) throw new Error('cell ' + addr + ' <v> not found in sheet XML')
    xml = xml.replace(re, '$1<v>' + literal + '</v>')
  }
  zip.file(sheetName, xml)
  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  fs.writeFileSync(path.join(__dirname, 'c3-acceptance.xlsx'), out)
  console.log('[fixture] c3-acceptance.xlsx bytes=' + out.length + ' patches=' + patches.map(p => p[0]).join(','))
}

// Parse the zip central directory of a generated archive and return { uncompressedSize,
// compressedSize } for one entry name - the honest "实测" pair for the ratio printout (JSZip's
// internal _data fields are an implementation detail; the central directory is the on-disk truth).
function centralDirectoryEntrySize(buffer, entryName) {
  const SIGNATURE = 0x02014b50
  for (let i = buffer.length - 46; i >= 0; i--) {
    if (buffer.readUInt32LE(i) !== SIGNATURE) continue
    const nameLength = buffer.readUInt16LE(i + 28)
    if (i + 46 + nameLength > buffer.length) continue
    if (buffer.toString('utf8', i + 46, i + 46 + nameLength) !== entryName) continue
    return {
      uncompressedSize: buffer.readUInt32LE(i + 24),
      compressedSize: buffer.readUInt32LE(i + 20)
    }
  }
  throw new Error('central directory entry not found: ' + entryName)
}

async function buildBudgetRefusal() {
  // 逐条目压缩比触发器（整改 5）：审计 auditZipEntries 在读入循环内检查
  //   entryExpanded > perEntryLimit * compressed （perEntryLimit = 1/ZipSecureFile.getMinInflateRatio() = 100）
  // 命中即结构化拒绝 BudgetRefusal("COMPRESSION_RATIO", 100, entryExpanded/compressed)，
  // 后端消息形如 "IMPORT_BUDGET_EXCEEDED: COMPRESSION_RATIO 上限 100，观察值 N"。
  // sheet1.xml 用约 200KB 的全 '0' 字节流（包成合法 XML 的内联字符串单元格，其余 entry 保持合法），
  // DEFLATE 后压缩比数百倍，必然 > 100；整体展开 200KB 远低于 EXPANDED_BYTES(512MB)，
  // 整包压缩比（expanded/containerBytes ≈ 22）也低于 100，先命中的必是逐条目审计。
  const PAYLOAD_BYTES = 200 * 1024
  const zeros = '0'.repeat(PAYLOAD_BYTES)
  const sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>' + zeros + '</t></is></c></row></sheetData>' +
    '</worksheet>'

  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
  zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="T" sheetId="1" r:id="rId1"/></sheets></workbook>')
  zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
  zip.file('xl/worksheets/sheet1.xml', sheetXml)
  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  fs.writeFileSync(path.join(__dirname, 'c3-budget-refusal.xlsx'), out)
  // 实测压缩比证据：从产物 zip 的 central directory 读该条目的真实 expanded/compressed 字节数。
  // F3 口径修正（仅文案，不改行为/夹具）：本行打印的是该条目的【最终压缩比】（expanded/compressed
  // ≈ 505，触发器充分）；后端落库的 observed 是逐条目审计"entryExpanded > 100×compressed"在
  // 流式读入的 8192 字节缓冲边界跨过时点的 floor 取整（Codex 隔离库实测落库 observed=100，
  // 40960/406≈100.88→100）——最终压缩比 505 不等于落库观测值，二者口径不可混用。
  const sizes = centralDirectoryEntrySize(out, 'xl/worksheets/sheet1.xml')
  const ratio = Math.floor(sizes.uncompressedSize / Math.max(1, sizes.compressedSize))
  console.log('[fixture] c3-budget-refusal.xlsx bytes=' + out.length +
    ' sheet1.xml expanded=' + sizes.uncompressedSize + ' compressed=' + sizes.compressedSize +
    ' measuredFinalRatio=' + (sizes.uncompressedSize / Math.max(1, sizes.compressedSize)).toFixed(1) +
    ' (实测最终压缩比 ' + Math.round(sizes.uncompressedSize / Math.max(1, sizes.compressedSize)) +
    '，触发器充分；落库 observed 为流式触发时点 floor 取整——隔离库实测 100，最终压缩比不等于观测值)' +
    ' floorRatio=' + ratio + ' (perEntryLimit=100, require ratio>100 => COMPRESSION_RATIO refusal)')
  if (!(ratio > 100)) throw new Error('measured per-entry ratio ' + ratio + ' does not exceed 100 - fixture would not trigger the refusal')
}

;(async () => {
  await buildAcceptance()
  await buildBudgetRefusal()
  console.log('[fixture] done')
})().catch(e => { console.error(e); process.exitCode = 1 })
