/**
 * 本地 mock 解析器：把中文自然语言指令解析为 actions
 * 纯函数，无网络无副作用。后端未就绪或用户开启"本地模式"时使用
 * 设计：规则按数组顺序匹配，命中即加入 actions（可多条同时命中）
 */
import { coerceValue, zhColorToEn, MAX_MOCK_DELETE_ROWS } from './ai-action-schema'

// 单格或区域 A1 片段（用于复用）
const CELL = '[A-Za-z]+\\d+'
const RANGE = '[A-Za-z]+\\d+(?::[A-Za-z]+\\d+)?'
const CELL_OR_RANGE = '[A-Za-z]+\\d+(?::[A-Za-z]+\\d+)?'
const ZH_COLORS = '红色|绿色|蓝色|黄色|黑色|白色|灰色|橙色|粉色|紫色'

const rules = [
  // —— 撤销/重做（优先级最高，先匹配）——
  {
    name: 'undo',
    re: /撤销|undo/i,
    make: () => ({ type: 'undo' })
  },
  {
    name: 'redo',
    re: /重做|恢复|redo/i,
    make: () => ({ type: 'redo' })
  },
  // —— 清空 ——
  {
    name: 'clearRange',
    re: new RegExp('清空\\s*(' + RANGE + ')'),
    make: (m) => ({ type: 'clearRange', range: up(m[1]), clearWhat: 'content' })
  },
  {
    name: 'clearRow',
    re: /清空\s*第\s*(\d+)\s*行/,
    make: (m) => ({ type: 'clearRange', range: rowRange(+m[1]), clearWhat: 'content' })
  },
  {
    name: 'clearCol',
    re: /清空\s*第\s*(\d+)\s*列/,
    make: (m) => ({ type: 'clearRange', range: colRange(+m[1]), clearWhat: 'content' })
  },
  // —— 合并/取消合并 ——
  {
    name: 'mergeAcross',
    re: new RegExp('横向合并\\s*(' + RANGE + ')'),
    make: (m) => ({ type: 'merge', range: up(m[1]), mode: 'across', force: true })
  },
  {
    name: 'mergeVertically',
    re: new RegExp('纵向合并\\s*(' + RANGE + ')'),
    make: (m) => ({ type: 'merge', range: up(m[1]), mode: 'vertically', force: true })
  },
  {
    name: 'merge',
    re: new RegExp('合并\\s*(' + RANGE + ')'),
    make: (m) => ({ type: 'merge', range: up(m[1]), mode: 'default', force: true })
  },
  {
    name: 'breakApart',
    re: new RegExp('(?:拆分|取消合并|取消合并|拆开)\\s*(' + RANGE + ')'),
    make: (m) => ({ type: 'breakApart', range: up(m[1]) })
  },
  // —— 求和/平均 填入 ——
  {
    name: 'sumToCell',
    re: new RegExp('(' + CELL + '(?::' + CELL + ')?)\\s*(?:列|区域)?\\s*求和[\\s，,。.]*填(?:到|入)?\\s*(' + CELL + ')'),
    make: (m) => {
      const src = up(m[1])
      return { type: 'sumToCell', sourceRange: src, targetRange: up(m[2]) }
    }
  },
  {
    name: 'sumColumnToCell',
    re: new RegExp('([A-Za-z]+)\\s*列\\s*求和[\\s，,。.]*填(?:到|入)?\\s*(' + CELL + ')'),
    make: (m) => {
      const col = up(m[1])
      return { type: 'sumToCell', sourceRange: col + ':' + col, targetRange: up(m[2]) }
    }
  },
  {
    name: 'avgToCell',
    re: new RegExp('(' + CELL + '(?::' + CELL + ')?)\\s*(?:列|区域)?\\s*平均[\\s，,。.]*填(?:到|入)?\\s*(' + CELL + ')'),
    make: (m) => ({
      type: 'setFormula',
      range: up(m[2]),
      formula: 'AVERAGE(' + up(m[1]) + ')'
    })
  },
  // —— 设值 ——
  {
    name: 'setCell',
    re: new RegExp('把\\s*(' + CELL + ')\\s*(?:改成?|改为|设置成?|设为|设置为)\\s*(.+)'),
    make: (m) => ({ type: 'setCell', range: up(m[1]), value: coerceValue(m[2].trim()) })
  },
  {
    name: 'setCellEqual',
    re: new RegExp('(?:^|[^A-Za-z0-9])(' + CELL + ')\\s*=\\s*(.+?)(?:$|[，。\\s])'),
    make: (m) => ({ type: 'setCell', range: up(m[1]), value: coerceValue(m[2].trim()) })
  },
  // —— 样式：背景色 ——
  {
    name: 'setBgColor',
    re: new RegExp('(?:把)?\\s*(' + CELL_OR_RANGE + ')\\s*(?:的)?\\s*背景[\\s，,。.色]*\\s*(?:设为?|改成?|改为?|设置为?|变|填充)?\\s*(' + ZH_COLORS + '|#?[0-9a-fA-F]{6})'),
    make: (m) => ({ type: 'setStyle', range: up(m[1]), background: normalizeColor(m[2]) })
  },
  // —— 样式：字体颜色 ——
  {
    name: 'setFontColor',
    re: new RegExp('(?:把)?\\s*(' + CELL_OR_RANGE + ')\\s*(?:的)?\\s*(?:字体|字)\\s*[\\s，,。.颜色色]*\\s*(?:设为?|改成?|改为?|设置为?)?\\s*(' + ZH_COLORS + '|#?[0-9a-fA-F]{6})'),
    make: (m) => ({ type: 'setStyle', range: up(m[1]), fontColor: normalizeColor(m[2]) })
  },
  // —— 样式：加粗 ——
  {
    name: 'bold',
    re: new RegExp('(?:把)?\\s*(' + CELL_OR_RANGE + ')\\s*(?:的)?\\s*(?:字体|文字)?\\s*加粗'),
    make: (m) => ({ type: 'setStyle', range: up(m[1]), fontWeight: 'bold' })
  },
  // —— 样式：居中 ——
  {
    name: 'center',
    re: new RegExp('(?:把)?\\s*(' + CELL_OR_RANGE + ')\\s*(?:的)?\\s*(?:水平)?居中'),
    make: (m) => ({ type: 'setStyle', range: up(m[1]), hAlign: 'center' })
  },
  // —— 样式：字号 ——
  {
    name: 'fontSize',
    re: new RegExp('(?:把)?\\s*(' + CELL_OR_RANGE + ')\\s*(?:的)?\\s*字号\\s*(?:设为?|改成?|改为?|设置为?)?\\s*(\\d+)'),
    make: (m) => ({ type: 'setStyle', range: up(m[1]), fontSize: +m[2] })
  },
  // —— 插入行 ——
  {
    name: 'insertRowsAfter',
    re: /在?\s*第\s*(\d+)\s*行后\s*插入\s*(\d+)?\s*行/,
    make: (m) => ({
      type: 'insertRows',
      rowIndex: +m[1],
      count: m[2] ? +m[2] : 1,
      position: 'after'
    })
  },
  {
    name: 'insertRowsBefore',
    re: /在?\s*第\s*(\d+)\s*行前\s*插入\s*(\d+)?\s*行/,
    make: (m) => ({
      type: 'insertRows',
      rowIndex: +m[1],
      count: m[2] ? +m[2] : 1,
      position: 'before'
    })
  },
  // —— 删除行 ——
  {
    name: 'deleteRowsRange',
    re: /删除\s*第\s*(\d+)\s*行\s*到\s*第\s*(\d+)\s*行/,
    make: (m) => ({
      type: 'deleteRows',
      rowPosition: +m[1],
      count: Math.min(Math.max(1, +m[2] - +m[1] + 1), MAX_MOCK_DELETE_ROWS)
    })
  },
  {
    name: 'deleteRows',
    re: /删除\s*第\s*(\d+)\s*行/,
    make: (m) => ({ type: 'deleteRows', rowPosition: +m[1], count: 1 })
  },
  // —— 读取 ——
  {
    name: 'getRangeValue',
    re: new RegExp('(?:读取|看一?下|查询|读一下)\\s*(' + RANGE + ')'),
    make: (m) => ({ type: 'getRangeValue', range: up(m[1]) })
  },
  {
    name: 'getSelection',
    re: /当前选(?:区|中)|选区(?:是什么|是啥)/,
    make: () => ({ type: 'getSelection' })
  },
  // —— 筛选：自定义公式（优先匹配，避免被其他规则吞掉）——
  // 形如：对 A1:D20 的 C 列筛选公式 =AND(C1>100, C1<200)
  // 或：对 C 列筛选公式 C1>100
  {
    name: 'filterCustomFormula',
    re: new RegExp('(?:对|让)?\\s*(' + CELL_OR_RANGE + ')?\\s*(?:的)?\\s*([A-Za-z]+)\\s*列\\s*筛选\\s*公式\\s*(=?[\\s\\S]+)'),
    make: (m) => {
      const range = m[1] ? up(m[1]) : null
      const column = up(m[2])
      const formula = m[3].trim().replace(/[，。.]$/, '').trim()
      const action = { type: 'setFilter', column, customFormula: formula }
      if (range) action.range = range
      return action
    }
  },
  // —— 筛选：自然句式 "筛选出/找出 A、B、C 列包含 X 的内容" ——
  // 支持单列或多列（顿号/逗号/空格分隔），值可裸或被引号包裹
  // 不填 range，由 executor 在缺失时自动探测 usedRange
  {
    name: 'filterOutContains',
    re: /(?:筛选|找出|搜出|过滤出?|查询出?|查)\s*出?\s*([A-Za-z]+(?:\s*[、,，\s]\s*[A-Za-z]+)+)\s*列\s*包含\s*["'"]?([^，。"'\s]+?)["'"]?(?:\s*(?:的)?\s*(?:内容|数据|行|记录|结果))?(?:$|[，。.])/,
    make: (m) => {
      const cols = m[1].split(/[、,，\s]+/).map((s) => s.trim().toUpperCase()).filter(Boolean)
      const value = m[2].trim()
      return cols.map((column) => ({
        type: 'setFilter',
        column,
        condition: { operator: 'contains', value }
      }))
    }
  },
  // —— 单列自然句式 "筛选出 C 列包含 X 的内容" ——
  // （多列规则不匹配单列，因为它的列组要求至少一个分隔符，所以单列要单独一条）
  {
    name: 'filterOutContainsSingle',
    re: /(?:筛选|找出|搜出|过滤出?|查询出?|查)\s*出?\s*([A-Za-z]+)\s*列\s*包含\s*["'"]?([^，。"'\s]+?)["'"]?(?:\s*(?:的)?\s*(?:内容|数据|行|记录|结果))?(?:$|[，。.])/,
    make: (m) => {
      const column = up(m[1])
      const value = m[2].trim()
      return {
        type: 'setFilter',
        column,
        condition: { operator: 'contains', value }
      }
    }
  },
  // —— 筛选：包含（必须在通用条件之前匹配，避免被"等于"吞）——
  // 形如：对 A1:D20 的 C 列筛选包含 张三 / 筛选 C 列包含张三
  {
    name: 'filterContains',
    re: new RegExp('(?:对|让)?\\s*(' + CELL_OR_RANGE + ')?\\s*(?:的)?\\s*([A-Za-z]+)\\s*列\\s*筛选\\s*包含\\s*([^，。.\\s]+)'),
    make: (m) => {
      const range = m[1] ? up(m[1]) : null
      const column = up(m[2])
      const value = m[3].trim()
      const action = {
        type: 'setFilter',
        column,
        condition: { operator: 'contains', value }
      }
      if (range) action.range = range
      return action
    }
  },
  // —— 筛选：比较条件（大于/小于/等于/不等于/大于等于/小于等于）——
  // 形如：对 A1:D20 的 C 列筛选大于 100 / 筛选 C 列大于100 / 筛选 C 列等于 张三
  {
    name: 'filterCondition',
    re: new RegExp('(?:对|让)?\\s*(' + CELL_OR_RANGE + ')?\\s*(?:的)?\\s*([A-Za-z]+)\\s*列\\s*筛选\\s*(大于等于|小于等于|大于|小于|等于|不等于)\\s*([^，。.\\s]+)'),
    make: (m) => {
      const range = m[1] ? up(m[1]) : null
      const column = up(m[2])
      const opMap = {
        '大于等于': 'greaterThanOrEqual',
        '小于等于': 'lessThanOrEqual',
        '大于': 'greaterThan',
        '小于': 'lessThan',
        '等于': 'equal',
        '不等于': 'notEqual'
      }
      const operator = opMap[m[3]]
      const value = m[4].trim()
      const action = {
        type: 'setFilter',
        column,
        condition: { operator, value: coerceValue(value) }
      }
      if (range) action.range = range
      return action
    }
  },
  // —— 筛选：按值枚举 ——
  // 形如：对 A1:D20 的 C 列筛选值为 1,5,9 / 筛选 C 列值为 1、5、9
  {
    name: 'filterValues',
    re: new RegExp('(?:对|让)?\\s*(' + CELL_OR_RANGE + ')?\\s*(?:的)?\\s*([A-Za-z]+)\\s*列\\s*筛选(?:值为?|值)\\s*([^，。.]+)'),
    make: (m) => {
      const range = m[1] ? up(m[1]) : null
      const column = up(m[2])
      const rawValues = m[3].trim()
      const filters = rawValues.split(/[、,，\s]+/).filter((s) => s).map((s) => s.trim())
      const action = { type: 'setFilter', column, filters }
      if (range) action.range = range
      return action
    }
  },
  // —— 清除筛选 ——
  // 形如：清除 A1:D20 的 C 列筛选 / 清除 C 列筛选 / 清除筛选 / 取消筛选
  {
    name: 'clearFilterColumn',
    re: new RegExp('(?:清除|取消)\\s*(' + CELL_OR_RANGE + ')?\\s*(?:的)?\\s*([A-Za-z]+)\\s*列\\s*筛选'),
    make: (m) => {
      const range = m[1] ? up(m[1]) : null
      const column = up(m[2])
      const action = { type: 'clearFilter', column }
      if (range) action.range = range
      return action
    }
  },
  {
    name: 'clearFilterAll',
    re: /(?:清除|取消)\s*筛选/,
    make: () => ({ type: 'clearFilter' })
  },
  {
    name: 'removeFilter',
    re: /(?:移除|删除)\s*筛选/,
    make: () => ({ type: 'clearFilter', removeFilter: true })
  },
  {
    name: 'getFilter',
    re: /(?:查看|读取|看一?下|查询)\s*(?:当前)?\s*筛选/,
    make: () => ({ type: 'getFilter' })
  },
  // —— 绘图：两个区域 + 图表类型（"A2:A5 B2:B5 画成柱状图"）——
  {
    name: 'chartByRange',
    re: new RegExp('(' + RANGE + ')[\\s，,、。]*(?:和|与|跟)?[\\s，,、。]*(' + RANGE + ')[\\s，,、。]*(?:画|生成|创建|绘制|制作|做|弄)(?:成|一个|一张|个)?\\s*(柱状图|条形图|折线图|曲线图|饼图|扇形图)'),
    make: (m) => ({
      type: 'createChart',
      chartType: chartTypeOf(m[3]),
      categoryRange: up(m[1]),
      seriesRange: up(m[2])
    })
  },
  // —— 绘图：图表类型在前（"画一个柱状图 X轴A2:A5 Y轴B2:B5"）——
  {
    name: 'chartVerbFirst',
    re: new RegExp('(?:画|生成|创建|绘制|制作|做|弄)\\s*(?:一个|一张|个)?\\s*(柱状图|条形图|折线图|曲线图|饼图|扇形图)[\\s，,、。]*[^A-Za-z]*(' + RANGE + ')[\\s，,、。]*[^A-Za-z]*(' + RANGE + ')'),
    make: (m) => ({
      type: 'createChart',
      chartType: chartTypeOf(m[1]),
      categoryRange: up(m[2]),
      seriesRange: up(m[3])
    })
  }
]

/**
 * 主入口：解析中文指令
 * @param {string} message
 * @param {object} context 表格上下文（mock 解析时基本不用，保留参数对齐后端接口）
 * @returns {{reply:string, actions:Array}}
 */
export function mockParse(message, context) {
  if (!message || !message.trim()) {
    return { reply: '请输入指令。', actions: [] }
  }
  const actions = []
  for (const rule of rules) {
    const m = rule.re.exec(message)
    if (m) {
      try {
        const ret = rule.make(m)
        // make 可以返回单个 action 或 action 数组（多列筛选等会一次产生多个 action）
        if (Array.isArray(ret)) {
          for (const a of ret) {
            if (a) actions.push(a)
          }
        } else if (ret) {
          actions.push(ret)
        }
      } catch (e) {
        // 单条规则异常忽略，继续下一条
      }
    }
  }
  if (actions.length === 0) {
    return {
      reply:
        '本地演示模式暂未识别该指令。可试试：\n' +
        '• 把A1改成100\n' +
        '• B2:B5求和填到D1\n' +
        '• 清空B2:D10\n' +
        '• 合并A1:C1\n' +
        '• 把B2:B5背景设为红色\n' +
        '• A1:A3加粗\n' +
        '• 在第2行后插入3行\n' +
        '• 删除第2行到第4行\n' +
        '• 对 A1:D20 的 C 列筛选大于 100\n' +
        '• 筛选 C 列包含 张三\n' +
        '• 筛选 C 列筛选公式 =AND(C1>100, C1<200)\n' +
        '• 清除 C 列筛选 / 移除筛选 / 查看筛选\n' +
        '• 用 A2:A5 和 B2:B5 画柱状图 / 画一个饼图 A2:A5 B2:B5\n' +
        '• 撤销 / 读取一下A1:C3',
      actions: []
    }
  }
  return {
    reply: '已识别 ' + actions.length + ' 个操作：' + actions.map((a) => a.type).join('、'),
    actions
  }
}

// ============ 工具 ============

function up(s) {
  return String(s || '').toUpperCase()
}

/**
 * 中文图表名 → chartType（bar/line/pie）
 */
function chartTypeOf(word) {
  if (!word) return 'bar'
  if (/柱|条形/.test(word)) return 'bar'
  if (/折|曲线|线图/.test(word)) return 'line'
  if (/饼|扇形/.test(word)) return 'pie'
  return 'bar'
}

/**
 * 颜色归一化：中文色名转英文，#rrggbb 补 #
 */
function normalizeColor(s) {
  const c = zhColorToEn(s)
  if (c) return c
  // 纯 hex 无 #，补 #
  const hex = /^([0-9a-fA-F]{6})$/.exec((s || '').trim())
  if (hex) return '#' + hex[1]
  return s
}

/**
 * 行号(1-based) → A1 区域字符串，表示该整行
 * 第2行 → A2:Z2（保守取到 26 列）
 */
function rowRange(rowNum) {
  return 'A' + rowNum + ':Z' + rowNum
}

/**
 * 列号(1-based) → A1 区域字符串，表示该整列
 * 第3列 → C1:C100（保守取到 100 行）
 */
function colRange(colNum) {
  const letter = colIndexToLetter(colNum - 1)
  return letter + '1:' + letter + '100'
}

function colIndexToLetter(idx) {
  let s = ''
  let c = idx
  do {
    s = String.fromCharCode(65 + (c % 26)) + s
    c = Math.floor(c / 26) - 1
  } while (c >= 0)
  return s
}
