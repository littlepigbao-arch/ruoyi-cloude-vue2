/**
 * AI 指令 Schema —— 白名单、A1 解析、值转换、颜色转换、校验、越界截断、公式安全校验
 * 纯函数模块，无副作用，不依赖 univer / Vue，可被 executor / mock-parser / chat 复用
 */

// 单次 actions 数量上限
export const MAX_ACTIONS = 30
// insert/delete 行列数量上限
export const MAX_INSERT_DELETE = 100
// mock 模式下 deleteRows 数量上限（更保守）
export const MAX_MOCK_DELETE_ROWS = 5

/**
 * 指令白名单：type 必须命中此列表，否则执行器跳过
 */
export const ACTION_WHITELIST = [
  'setCell',
  'setCellForCell',
  'setValues',
  'clearRange',
  'setFormula',
  'merge',
  'breakApart',
  'insertRows',
  'deleteRows',
  'insertColumns',
  'deleteColumns',
  'setRowHeight',
  'setColumnWidth',
  'setStyle',
  'undo',
  'redo',
  'getRangeValue',
  'getRangeFormulas',
  'getMerges',
  'getSelection',
  'sumToCell',
  'setFilter',
  'clearFilter',
  'getFilter',
  'createChart'
]

/**
 * 图表类型白名单
 */
export const CHART_TYPES = ['bar', 'line', 'pie']

/**
 * 筛选条件操作符（语义层 → Univer CustomFilterOperator）
 * contains（包含）走 textMatch 通配符路径，operator 留空，val 包成 *xxx*
 */
export const FILTER_OPERATORS = {
  greaterThan: 'greaterThan',
  lessThan: 'lessThan',
  equal: 'equal',
  notEqual: 'notEqual',
  greaterThanOrEqual: 'greaterThanOrEqual',
  lessThanOrEqual: 'lessThanOrEqual',
  contains: 'contains' // 特殊：转 customFilters[0].val = *xxx*，operator 不填
}

/**
 * 把语义 operator + value 转成 Univer customFilters 结构
 * 返回 { customFilters: { customFilters: [{val, operator?}] } }
 */
export function buildCustomFilters(operator, value) {
  const allowed = ['greaterThan', 'lessThan', 'equal', 'notEqual', 'greaterThanOrEqual', 'lessThanOrEqual', 'contains']
  if (!allowed.includes(operator)) return null
  if (value === undefined || value === null) return null
  // 数字优先保留 number 类型，Univer 的 greaterThan 等会校验 typeof value === 'number'
  let val = value
  if (typeof val === 'string') {
    const trimmed = val.trim()
    if (/^-?\d+(\.\d+)?$/.test(trimmed)) val = Number(trimmed)
  }
  if (operator === 'contains') {
    // 包含：走 textMatch 通配符，* → .*
    const v = typeof value === 'string' ? value : String(value)
    return { customFilters: { customFilters: [{ val: '*' + v + '*' }] } }
  }
  return { customFilters: { customFilters: [{ val, operator }] } }
}

/**
 * 把自定义公式表达式（如 "C1>100"、"=AND(C1>100,C1<200)"）拆解为 Univer customFilters 结构
 * 支持：
 *   - 单比较: >100, <100, =100, !=100, >=100, <=100, =张三, 包含张三
 *   - AND/OR 双比较: AND(>100, <200) → customFilters.and=1, customFilters=[{>,100},{<,200}]
 * 不支持嵌套或更复杂公式，返回 null 让上层报错
 */
export function parseFormulaToCustomFilters(formulaStr) {
  if (!formulaStr || typeof formulaStr !== 'string') return null
  let s = formulaStr.trim()
  if (s.startsWith('=')) s = s.slice(1).trim()
  // AND(...) / OR(...)
  const combo = s.match(/^(AND|OR)\s*\((.+)\)$/i)
  if (combo) {
    const isAnd = combo[1].toUpperCase() === 'AND'
    const parts = combo[2].split(',').map((x) => x.trim())
    if (parts.length !== 2) return null
    const a = parseSingleCompare(parts[0])
    const b = parseSingleCompare(parts[1])
    if (!a || !b) return null
    return {
      customFilters: {
        and: isAnd ? 1 : undefined,
        customFilters: [a, b]
      }
    }
  }
  const single = parseSingleCompare(s)
  if (!single) return null
  return { customFilters: { customFilters: [single] } }
}

function parseSingleCompare(expr) {
  if (!expr) return null
  const s = String(expr).trim()
  // 支持 包含xxx / 包含"xxx"
  const containsM = s.match(/^包含\s*"?([^"]+?)"?$/i)
  if (containsM) return { val: '*' + containsM[1] + '*' }
  // 比较运算符
  const m = s.match(/^(>=|<=|!=|>|<|=)\s*(.+)$/)
  if (!m) return null
  const opMap = { '>': 'greaterThan', '<': 'lessThan', '>=': 'greaterThanOrEqual', '<=': 'lessThanOrEqual', '!=': 'notEqual', '=': 'equal' }
  const operator = opMap[m[1]]
  if (!operator) return null
  let val = m[2].trim()
  // 去引号
  if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
    val = val.slice(1, -1)
  }
  // 数字优先
  if (/^-?\d+(\.\d+)?$/.test(val)) val = Number(val)
  return { val, operator }
}

/**
 * 危险公式模式（含外部链接/网络请求，一律拒绝）
 */
const DANGEROUS_PATTERNS = [
  /HYPERLINK\s*\(\s*["']?\s*https?:/i,
  /WEBSERVICE\s*\(/i,
  /IMPORTDATA\s*\(/i,
  /IMPORTXML\s*\(/i,
  /IMPORTHTML\s*\(/i,
  /IMAGE\s*\(/i,
  /FILTERXML\s*\(/i
]

/**
 * 公式函数白名单（首函数名匹配，命中即允许）
 */
const SAFE_FORMULA_FN =
  /^(=)?\s*(SUM|AVERAGE|MAX|MIN|COUNT|COUNTA|COUNTIF|COUNTIFS|SUMIF|SUMIFS|IF|IFS|IFS|VLOOKUP|HLOOKUP|INDEX|MATCH|CONCATENATE|CONCAT|TEXT|ROUND|ROUNDUP|ROUNDDOWN|LEFT|RIGHT|MID|LEN|TRIM|UPPER|LOWER|ABS|SQRT|POWER|MOD|INT|DATE|TODAY|NOW|YEAR|MONTH|DAY|WEEKDAY|EOMONTH|IFERROR|ISBLANK|ISNUMBER|ISTEXT|ROW|COLUMN|INDIRECT|OFFSET|SUMPRODUCT|AVERAGEIF|AVERAGEIFS|MAXIFS|MINIFS|RANK|LARGE|SMALL|MEDIAN|MODE|VAR|STDEV|AND|OR|NOT|TRUE|FALSE|NA|EXACT|FIND|SEARCH|REPLACE|SUBSTITUTE|REPT|VALUE|NUMBERVALUE|DAYS|NETWORKDAYS|WORKDAY)\b/i

/**
 * 颜色校验正则：#rrggbb 或命名色
 */
const COLOR_RE = /^#([0-9a-fA-F]{6})$|^(red|green|blue|yellow|black|white|gray|orange|pink|purple)$/i

/**
 * 中文颜色 → 英文
 */
const ZH_COLOR_MAP = {
  红色: 'red',
  绿色: 'green',
  蓝色: 'blue',
  黄色: 'yellow',
  黑色: 'black',
  白色: 'white',
  灰色: 'gray',
  橙色: 'orange',
  粉色: 'pink',
  紫色: 'purple'
}

/**
 * 把 A1 列字母（A/B/AA）转成 0-based 列号
 * A=0, Z=25, AA=26
 */
function colLetterToIndex(letters) {
  let n = 0
  const s = letters.toUpperCase()
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i)
    if (code < 65 || code > 90) return -1
    n = n * 26 + (code - 64)
  }
  return n - 1
}

/**
 * 解析 A1 字符串为 {startRow,endRow,startColumn,endColumn}（0-based，含端点）
 * 支持：A1、A1:B2、Sheet1!A1、Sheet1!A1:B2
 * 解析失败返回 null
 * @param {string} str
 * @returns {{startRow:number,endRow:number,startColumn:number,endColumn:number}|null}
 */
export function parseA1(str) {
  if (!str || typeof str !== 'string') return null
  let s = str.trim()
  // 剥 sheet 前缀 Sheet1!A1
  const bang = s.indexOf('!')
  if (bang >= 0) s = s.slice(bang + 1).trim()
  if (!s) return null
  // 匹配 单元格或区域：[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?
  const m = s.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i)
  if (!m) return null
  const startCol = colLetterToIndex(m[1])
  const startRow = parseInt(m[2], 10) - 1
  if (startCol < 0 || startRow < 0) return null
  let endCol = startCol
  let endRow = startRow
  if (m[3] && m[4]) {
    endCol = colLetterToIndex(m[3])
    endRow = parseInt(m[4], 10) - 1
    if (endCol < 0 || endRow < 0) return null
  }
  // 规整：保证 start <= end
  if (startRow > endRow) [startRow, endRow] = [endRow, startRow]
  if (startCol > endCol) [startCol, endCol] = [endCol, startCol]
  return { startRow, endRow, startColumn: startCol, endColumn: endCol }
}

/**
 * 把行列号转成 A1 字符串（0-based）
 */
export function toA1(row, col) {
  let c = col
  let letters = ''
  do {
    letters = String.fromCharCode(65 + (c % 26)) + letters
    c = Math.floor(c / 26) - 1
  } while (c >= 0)
  return letters + (row + 1)
}

/**
 * 把区域对象转 A1 字符串
 */
export function rangeToA1(range) {
  if (!range) return ''
  if (range.startRow === range.endRow && range.startColumn === range.endColumn) {
    return toA1(range.startRow, range.startColumn)
  }
  return toA1(range.startRow, range.startColumn) + ':' + toA1(range.endRow, range.endColumn)
}

/**
 * 值转换：字符串转 number/boolean，否则原样字符串
 * "100" → 100, "true"→true, "张三"→"张三"
 */
export function coerceValue(v) {
  if (v === null || v === undefined) return ''
  if (typeof v !== 'string') return v
  const s = v.trim()
  if (s === '') return ''
  // 布尔
  if (/^(true|false)$/i.test(s)) return s.toLowerCase() === 'true'
  // 数字（含负数、小数、百分号？这里只转纯数字）
  if (/^-?\d+(\.\d+)?$/.test(s)) {
    const n = Number(s)
    if (!isNaN(n)) return n
  }
  return s
}

/**
 * 中文颜色 → 英文（#rrggbb 原样返回）
 */
export function zhColorToEn(zh) {
  if (!zh || typeof zh !== 'string') return null
  const s = zh.trim()
  if (/^#([0-9a-fA-F]{6})$/.test(s)) return s
  if (ZH_COLOR_MAP[s]) return ZH_COLOR_MAP[s]
  // 英文色名直接返回
  if (/^(red|green|blue|yellow|black|white|gray|orange|pink|purple)$/i.test(s)) return s.toLowerCase()
  return null
}

/**
 * 颜色校验：通过返回 true，否则 false
 */
export function isValidColor(color) {
  return !!color && COLOR_RE.test(color)
}

/**
 * 公式安全校验
 * @param {string} formula
 * @returns {true|string} true 表示通过，字符串表示拒绝原因
 */
export function assertSafeFormula(formula) {
  if (typeof formula !== 'string') return '公式必须是字符串'
  let f = formula.trim()
  if (!f) return '公式不能为空'
  // 强制以 = 开头
  if (!f.startsWith('=')) f = '=' + f
  // 拒绝危险函数
  for (const p of DANGEROUS_PATTERNS) {
    if (p.test(f)) return '不允许的公式函数（含外部链接/网络）'
  }
  // 白名单函数匹配（取首个函数名）
  if (!SAFE_FORMULA_FN.test(f)) {
    const m = f.match(/^=\s*([A-Z]+)/i)
    return '函数不在白名单: ' + (m ? m[1] : '未知')
  }
  return true
}

/**
 * 越界截断：把 range 截到 [minR,maxR]×[minC,maxC] 合法边界
 * @returns {{range:{startRow,endRow,startColumn,endColumn}, changed:boolean, original:object}}
 */
export function clampRange(range, minR, maxR, minC, maxC) {
  const original = { ...range }
  let startRow = Math.max(minR, Math.min(maxR, range.startRow))
  let endRow = Math.max(minR, Math.min(maxR, range.endRow))
  let startColumn = Math.max(minC, Math.min(maxC, range.startColumn))
  let endColumn = Math.max(minC, Math.min(maxC, range.endColumn))
  if (startRow > endRow) [startRow, endRow] = [endRow, startRow]
  if (startColumn > endColumn) [startColumn, endColumn] = [endColumn, startColumn]
  const changed =
    startRow !== original.startRow ||
    endRow !== original.endRow ||
    startColumn !== original.startColumn ||
    endColumn !== original.endColumn
  return { range: { startRow, endRow, startColumn, endColumn }, changed, original }
}

/**
 * 校验单个 action 的参数
 * @param {object} action
 * @returns {string|null} null 表示通过，字符串表示错误信息
 */
export function validateAction(action) {
  if (!action || typeof action !== 'object') return 'action 必须是对象'
  if (!action.type || typeof action.type !== 'string') return 'action.type 缺失'
  if (!ACTION_WHITELIST.includes(action.type)) return '未知指令类型: ' + action.type

  switch (action.type) {
    case 'setCell':
    case 'setCellForCell': {
      if (!action.range) return 'range 必填'
      if (!parseA1(action.range)) return 'range 解析失败: ' + action.range
      if (action.value === undefined || action.value === null) return 'value 必填'
      if (typeof action.value === 'string' && action.value.startsWith('=')) {
        const r = assertSafeFormula(action.value)
        if (r !== true) return r
      }
      return null
    }
    case 'setValues': {
      if (!action.range) return 'range 必填'
      if (!parseA1(action.range)) return 'range 解析失败: ' + action.range
      if (!Array.isArray(action.values) || !Array.isArray(action.values[0])) {
        return 'values 必须为二维数组'
      }
      return null
    }
    case 'clearRange':
    case 'breakApart':
    case 'getRangeValue':
    case 'getRangeFormulas': {
      if (!action.range) return 'range 必填'
      if (!parseA1(action.range)) return 'range 解析失败: ' + action.range
      return null
    }
    case 'setFormula': {
      if (!action.range) return 'range 必填'
      if (!parseA1(action.range)) return 'range 解析失败: ' + action.range
      if (!action.formula || typeof action.formula !== 'string') return 'formula 必填且为字符串'
      const r = assertSafeFormula(action.formula)
      if (r !== true) return r
      return null
    }
    case 'merge': {
      if (!action.range) return 'range 必填'
      if (!parseA1(action.range)) return 'range 解析失败: ' + action.range
      if (action.mode && !['default', 'across', 'vertically'].includes(action.mode)) {
        return 'mode 必须为 default|across|vertically'
      }
      return null
    }
    case 'insertRows':
    case 'insertColumns': {
      const idx = action.type === 'insertRows' ? action.rowIndex : action.columnIndex
      if (!Number.isFinite(+idx) || +idx < 0) return (action.type === 'insertRows' ? 'rowIndex' : 'columnIndex') + ' 非法'
      if (action.count !== undefined && (!Number.isFinite(+action.count) || +action.count < 1)) {
        return 'count 必须为正数'
      }
      if (+action.count > MAX_INSERT_DELETE) return 'count 超过上限 ' + MAX_INSERT_DELETE
      return null
    }
    case 'deleteRows':
    case 'deleteColumns': {
      const pos = action.type === 'deleteRows' ? action.rowPosition : action.columnPosition
      if (!Number.isFinite(+pos) || +pos < 0) {
        return (action.type === 'deleteRows' ? 'rowPosition' : 'columnPosition') + ' 非法'
      }
      if (action.count !== undefined && (!Number.isFinite(+action.count) || +action.count < 1)) {
        return 'count 必须为正数'
      }
      if (+action.count > MAX_INSERT_DELETE) return 'count 超过上限 ' + MAX_INSERT_DELETE
      return null
    }
    case 'setRowHeight': {
      if (!Number.isFinite(+action.rowPosition) || +action.rowPosition < 0) return 'rowPosition 非法'
      if (!Number.isFinite(+action.height) || +action.height < 0) return 'height 非法'
      return null
    }
    case 'setColumnWidth': {
      if (!Number.isFinite(+action.columnPosition) || +action.columnPosition < 0) return 'columnPosition 非法'
      if (!Number.isFinite(+action.width) || +action.width < 0) return 'width 非法'
      return null
    }
    case 'setStyle': {
      if (!action.range) return 'range 必填'
      if (!parseA1(action.range)) return 'range 解析失败: ' + action.range
      if (action.background && !isValidColor(action.background)) return 'background 颜色非法: ' + action.background
      if (action.fontColor && !isValidColor(action.fontColor)) return 'fontColor 颜色非法: ' + action.fontColor
      if (action.fontWeight && !['bold', 'normal'].includes(action.fontWeight)) {
        return 'fontWeight 必须为 bold|normal'
      }
      if (action.hAlign && !['left', 'center', 'right', 'normal'].includes(action.hAlign)) {
        return 'hAlign 非法'
      }
      if (action.vAlign && !['top', 'middle', 'bottom'].includes(action.vAlign)) {
        return 'vAlign 非法'
      }
      return null
    }
    case 'sumToCell': {
      if (!action.sourceRange || !parseA1(action.sourceRange)) return 'sourceRange 非法'
      if (!action.targetRange || !parseA1(action.targetRange)) return 'targetRange 非法'
      return null
    }
    case 'setFilter': {
      // range 可选：未填时 executor 自动探测 sheet 的 usedRange
      if (action.range && !parseA1(action.range)) return 'range 解析失败: ' + action.range
      // column 必填，支持字母（'C'）或数字（绝对列索引）
      if (action.column === undefined && action.columnIndex === undefined) {
        return 'column 或 columnIndex 必填'
      }
      // 三选一：filters / condition / customFormula
      const hasFilters = Array.isArray(action.filters) && action.filters.length > 0
      const hasCondition = action.condition && typeof action.condition === 'object'
      const hasFormula = typeof action.customFormula === 'string' && action.customFormula.trim()
      if (!hasFilters && !hasCondition && !hasFormula) {
        return '必须提供 filters / condition / customFormula 之一'
      }
      if (hasCondition) {
        const op = action.condition.operator
        const allowedOps = Object.keys(FILTER_OPERATORS)
        if (!allowedOps.includes(op)) return 'condition.operator 非法: ' + op
        if (action.condition.value === undefined || action.condition.value === null) {
          return 'condition.value 必填'
        }
      }
      if (hasFormula) {
        const parsed = parseFormulaToCustomFilters(action.customFormula)
        if (!parsed) return 'customFormula 解析失败（支持 >100、=张三、包含张三、AND(>100,<200)）'
      }
      return null
    }
    case 'clearFilter': {
      // range 可选：未填则使用活动 sheet 上已存在的 filter
      if (action.range && !parseA1(action.range)) return 'range 解析失败: ' + action.range
      // column 可选：填则只清该列，不填则清所有列条件
      // removeFilter=true 时整张 filter 框移除
      if (action.removeFilter !== undefined && typeof action.removeFilter !== 'boolean') {
        return 'removeFilter 必须为布尔'
      }
      return null
    }
    case 'getFilter': {
      if (action.range && !parseA1(action.range)) return 'range 解析失败: ' + action.range
      return null
    }
    case 'createChart': {
      if (!CHART_TYPES.includes(action.chartType)) {
        return 'chartType 必须为 ' + CHART_TYPES.join('|')
      }
      if (action.aggregate && !['count', 'sum', 'avg'].includes(action.aggregate)) {
        return 'aggregate 必须为 count|sum|avg'
      }
      if (!action.categoryRange || !parseA1(action.categoryRange)) {
        return 'categoryRange 非法'
      }
      // count 统计分组次数时 seriesRange 可省略；sum/avg 必须提供数值列
      if (action.aggregate !== 'count' && (!action.seriesRange || !parseA1(action.seriesRange))) {
        return 'seriesRange 非法'
      }
      return null
    }
    case 'undo':
    case 'redo':
    case 'getMerges':
    case 'getSelection':
      return null
    default:
      return null
  }
}
