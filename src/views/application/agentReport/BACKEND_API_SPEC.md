# AI 表格助手 后端接口对接说明

> 前端已就绪，后端只需实现 **一个** 接口。前端会把用户中文指令 + 当前表格上下文发给你，你用大模型解析成结构化 JSON actions 返回，前端执行器照单全收。

---

## 一、接口地址

| 项 | 值 |
|---|---|
| 方法 | `POST` |
| 路径 | `/ai/chat` |
| Content-Type | `application/json;charset=utf-8` |
| 鉴权 | `Authorization: Bearer {token}`（前端自动注入，与现有若依体系统一） |
| 超时 | 前端设 60s，建议大模型流式/异步时在此内返回 |

> 路径走若依网关代理（`VUE_APP_BASE_API` 前缀，dev 下 `/dev-api`），实际到后端就是 `/ai/chat`。

---

## 二、请求体（前端发什么）

```json
{
  "message": "把A1改成100，B列求和填到D1",
  "conversationId": "conv-abc123",
  "history": [
    { "role": "user", "content": "把A1改成100" },
    { "role": "assistant", "content": "已设置 A1=100" }
  ],
  "context": { /* 见第三节，当前表格快照 */ },
  "clientMeta": {
    "locale": "zh-CN",
    "univerVersion": "0.25.1"
  }
}
```

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `message` | string | 是 | 用户当前输入的自然语言指令（中文） |
| `conversationId` | string | 否 | 会话 ID，首轮可由后端生成并回传 |
| `history` | array | 否 | 近几轮对话（role: user/assistant），用于多轮上下文 |
| `context` | object | 否 | **当前表格状态快照**，见第三节。后端应据此理解"当前选区""B列有什么"等指代 |
| `clientMeta` | object | 否 | 客户端元信息 |

---

## 三、表格上下文 context 结构

前端会把当前活动 sheet 的数据摘要发给你（已做 token 节流，最多 50 行 × 26 列）：

```jsonc
{
  "workbookName": "新建工作簿",
  "sheetName": "Sheet1",
  "sheetId": "sheet-0",
  "rowCount": 200,          // sheet 总行数
  "columnCount": 52,        // sheet 总列数
  "usedRange": {            // 实际有数据的范围（0-based，含端点）
    "startRow": 0,
    "endRow": 4,
    "startColumn": 0,
    "endColumn": 3
  },
  "values": [               // 二维数组，[行][列]，0-based 对齐 usedRange
    [100, "张三", "男", 25],
    [null, "李四", "女", 30],
    [null, null, null, null]
  ],
  "formulas": [             // 仅当表格含公式时才出现此字段
    [null, null, null, null],
    [null, null, null, null]
  ],
  "merges": [               // 合并单元格列表
    { "startRow": 0, "endRow": 0, "startColumn": 0, "endColumn": 2, "a1": "A1:C1" }
  ],
  "selection": {            // 当前选区（用户高亮的区域），可能为 null
    "a1": "B2:B5",
    "startRow": 1, "endRow": 4, "startColumn": 1, "endColumn": 1
  },
  "activeCell": {           // 当前激活单元格
    "row": 0, "column": 0, "a1": "A1"
  },
  "truncated": false        // true 表示表格被截断（超 50 行或 26 列），后端应注意数据不全
}
```

**关键：行列号全是 0-based**（第 1 行 = row 0，A 列 = column 0）。但 actions 里的 `range` 用 A1 字符串（`"A1"`、`"B2:D5"`），不用算行列号，详见第五节。

---

## 四、响应体（后端返回什么）

```json
{
  "code": 200,
  "msg": "ok",
  "data": {
    "reply": "已把 A1 改成 100，B 列求和结果填入 D1",
    "conversationId": "conv-abc123",
    "actions": [
      { "type": "setCell", "range": "A1", "value": 100 },
      { "type": "sumToCell", "sourceRange": "B2:B5", "targetRange": "D1" }
    ],
    "needFeedback": false
  }
}
```

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `code` | number | 是 | `200`=成功；**`503`=服务不可用**，前端会自动回退本地 mock 演示 |
| `msg` | string | 是 | 状态描述 |
| `data.reply` | string | 是 | 给用户看的自然语言回复（会显示在对话气泡里） |
| `data.conversationId` | string | 是 | 回传会话 ID |
| `data.actions` | array | 是 | **结构化指令数组**，前端逐条执行。可为空数组（纯回复） |
| `data.needFeedback` | bool | 否 | 预留，默认 false |

> **code=503 的意义**：后端模型限流/降级时返回 503，前端静默回退到本地正则解析器，用户无感继续可用（功能降级）。比直接报错友好。

---

## 五、actions 结构定义（核心）

每个 action 是一个对象，`type` 必须命中下列白名单，字段必须符合校验。**range 统一用 A1 字符串**（`"A1"`、`"B2:D5"`、`"Sheet1!A1:B2"`）。

### 5.1 白名单（25 种）

```
setCell / setCellForCell / setValues / clearRange / setFormula
merge / breakApart / insertRows / deleteRows / insertColumns / deleteColumns
setRowHeight / setColumnWidth / setStyle / undo / redo
getRangeValue / getRangeFormulas / getMerges / getSelection / sumToCell
setFilter / clearFilter / getFilter / createChart
```

### 5.2 各类型字段

#### setCell —— 设单个区域值
```jsonc
{ "type": "setCell", "range": "A1", "value": 100 }
{ "type": "setCell", "range": "B2", "value": "张三" }
{ "type": "setCell", "range": "C1", "value": "=SUM(B2:B5)" }  // value 以 = 开头即公式
```
| 字段 | 必填 | 说明 |
|------|------|------|
| range | 是 | A1 区域 |
| value | 是 | number/string/boolean。以 `=` 开头视为公式 |

#### setCellForCell —— 仅写区域左上首格
```jsonc
{ "type": "setCellForCell", "range": "A1:C3", "value": "=B1*2" }
```

#### setValues —— 批量写二维值
```jsonc
{
  "type": "setValues",
  "range": "A1:B2",
  "values": [ [1, "甲"], [2, "乙"] ]
}
```
| 字段 | 必填 | 说明 |
|------|------|------|
| range | 是 | A1 区域，行列数须与 values 维度一致 |
| values | 是 | 二维数组，`[行][列]`。元素可为 `{v:值, f:"公式"}` 对象 |

#### clearRange —— 清空
```jsonc
{ "type": "clearRange", "range": "B2:D10", "clearWhat": "content" }
```
| 字段 | 必填 | 说明 |
|------|------|------|
| range | 是 | A1 区域 |
| clearWhat | 否 | `content`(默认)/`format`/`all` |

#### setFormula —— 写公式
```jsonc
{ "type": "setFormula", "range": "D1", "formula": "SUM(B2:B5)" }
```
| 字段 | 必填 | 说明 |
|------|------|------|
| range | 是 | A1 区域 |
| formula | 是 | 公式字符串，不带 `=` 也行（前端补） |

#### merge —— 合并 / breakApart —— 拆分
```jsonc
{ "type": "merge", "range": "A1:C1", "mode": "across", "force": true }
{ "type": "breakApart", "range": "A1:C1" }
```
| 字段 | 必填 | 说明 |
|------|------|------|
| range | 是 | A1 区域 |
| mode | 否 | `default`(默认)/`across`(横向)/`vertically`(纵向) |
| force | 否 | bool，是否强制合并 |

#### insertRows / deleteRows / insertColumns / deleteColumns
```jsonc
{ "type": "insertRows", "rowIndex": 2, "count": 3, "position": "after" }
{ "type": "deleteRows", "rowPosition": 2, "count": 3 }
{ "type": "insertColumns", "columnIndex": 1, "count": 2, "position": "after" }
{ "type": "deleteColumns", "columnPosition": 1, "count": 2 }
```
| 字段 | 必填 | 说明 |
|------|------|------|
| rowIndex/columnIndex | insert 必填 | 插入位置（1-based 行/列号） |
| rowPosition/columnPosition | delete 必填 | 删除起始位置（1-based） |
| count | 否 | 数量，默认 1，上限 100 |
| position | 否 | `after`/`before`/`at`，默认 `at` |

#### setRowHeight / setColumnWidth
```jsonc
{ "type": "setRowHeight", "rowPosition": 1, "height": 30 }
{ "type": "setColumnWidth", "columnPosition": 1, "width": 120 }
```

#### setStyle —— 样式（一次可设多个属性）
```jsonc
{
  "type": "setStyle",
  "range": "B2:B5",
  "background": "#ff0000",
  "fontColor": "red",
  "fontSize": 14,
  "fontWeight": "bold",
  "fontFamily": "Arial",
  "hAlign": "center",
  "vAlign": "middle",
  "wrap": true,
  "textRotation": 45
}
```
| 字段 | 必填 | 说明 |
|------|------|------|
| range | 是 | A1 区域 |
| background | 否 | 背景色：`#rrggbb` 或英文色名 |
| fontColor | 否 | 字体色 |
| fontSize | 否 | 字号 number |
| fontWeight | 否 | `bold`/`normal` |
| hAlign | 否 | `left`/`center`/`right`/`normal` |
| vAlign | 否 | `top`/`middle`/`bottom` |
| wrap | 否 | bool 是否换行 |

#### sumToCell —— 求和填入（语义糖）
```jsonc
{ "type": "sumToCell", "sourceRange": "B2:B5", "targetRange": "D1" }
```
等价于在 targetRange 写 `=SUM(sourceRange)`。

#### undo / redo —— 撤销重做
```jsonc
{ "type": "undo" }
{ "type": "redo" }
```

#### getRangeValue / getRangeFormulas / getMerges / getSelection —— 读取类
```jsonc
{ "type": "getRangeValue", "range": "A1:C3" }
{ "type": "getRangeFormulas", "range": "A1:C3" }
{ "type": "getMerges" }
{ "type": "getSelection" }
```
这类 action 执行后结果会回显在对话气泡里，后端 `reply` 里也可自行描述。

#### createChart —— 根据表格数据绘图
```jsonc
{
  "type": "createChart",
  "chartType": "bar",              // bar | line | pie
  "categoryRange": "A2:A5",        // 分类轴（单列）
  "seriesRange": "B2:B5",          // 数值（单列或多列）
  "title": "各月销售额",            // 可选，图表标题
  "seriesName": "销售额",           // 可选，单系列名
  "seriesNames": ["销售额", "成本"] // 可选，多系列名（多列 seriesRange 时）
}
```
| 字段 | 必填 | 说明 |
|------|------|------|
| chartType | 是 | `bar`(柱状图)/`line`(折线图)/`pie`(饼图) |
| categoryRange | 是 | 分类轴数据区域（A1 字符串，单列） |
| seriesRange | 是 | 数值数据区域（A1 字符串，单列或多列；饼图只取第一列） |
| title | 否 | 图表标题 |
| seriesName | 否 | 单系列名称，默认「数值」 |
| seriesNames | 否 | 多系列名称数组，与 seriesRange 列数对应 |

> 前端执行后会在对话气泡中渲染一张 ECharts 图表。**关键**：模型需依据 `context.values` 推断正确的 `categoryRange` 与 `seriesRange`（通常分类列与数值列相邻，且首行为表头时从第 2 行开始）。

---

## 六、关键约束（前端会强校验，不符合直接丢弃）

| 约束 | 值 | 说明 |
|------|----|------|
| 单次 actions 数量上限 | 30 | 超出截断 |
| insert/delete 数量上限 | 100 | 防误删全表 |
| **公式函数白名单** | 见下 | 非白名单函数直接拒绝 |
| 危险公式 | HYPERLINK(https/WEBSERVICE/IMPORTDATA/IMPORTXML/IMPORTHTML/IMAGE/FILTERXML 含 http | 一律拒绝 |
| 颜色格式 | `#rrggbb` 或 `red/green/blue/...` | 不合法颜色丢弃 |
| range 格式 | A1 字符串 | `A1`、`B2:D5`、`Sheet1!A1:B2` |

**公式函数白名单**（首函数名匹配）：
```
SUM AVERAGE MAX MIN COUNT COUNTA COUNTIF COUNTIFS SUMIF SUMIFS
IF IFS VLOOKUP HLOOKUP INDEX MATCH CONCATENATE CONCAT TEXT
ROUND ROUNDUP ROUNDDOWN LEFT RIGHT MID LEN TRIM UPPER LOWER
ABS SQRT POWER MOD INT DATE TODAY NOW YEAR MONTH DAY WEEKDAY EOMONTH
IFERROR ISBLANK ISNUMBER ISTEXT ROW COLUMN INDIRECT OFFSET
SUMPRODUCT AVERAGEIF AVERAGEIFS MAXIFS MINIFS RANK LARGE SMALL
MEDIAN MODE VAR STDEV AND OR NOT TRUE FALSE NA EXACT FIND SEARCH
REPLACE SUBSTITUTE REPT VALUE NUMBERVALUE DAYS NETWORKDAYS WORKDAY
```

---

## 七、完整示例

### 请求
```json
{
  "message": "把B列有数据的求和填到D1，标题行加粗居中",
  "conversationId": "conv-001",
  "history": [],
  "context": {
    "workbookName": "销售表",
    "sheetName": "Sheet1",
    "rowCount": 200,
    "columnCount": 52,
    "usedRange": { "startRow": 0, "endRow": 5, "startColumn": 0, "endColumn": 3 },
    "values": [
      ["产品", "数量", "单价", null],
      ["苹果", 10, 5, null],
      ["香蕉", 20, 3, null],
      ["橘子", 15, 4, null]
    ],
    "merges": [],
    "selection": null,
    "activeCell": { "row": 0, "column": 0, "a1": "A1" },
    "truncated": false
  },
  "clientMeta": { "locale": "zh-CN", "univerVersion": "0.25.1" }
}
```

### 响应
```json
{
  "code": 200,
  "msg": "ok",
  "data": {
    "reply": "已对 B2:B4 求和填入 D1（=SUM(B2:B4)），并将标题行 A1:D1 加粗居中。",
    "conversationId": "conv-001",
    "actions": [
      { "type": "sumToCell", "sourceRange": "B2:B4", "targetRange": "D1" },
      { "type": "setStyle", "range": "A1:D1", "fontWeight": "bold", "hAlign": "center" }
    ],
    "needFeedback": false
  }
}
```

---

## 八、后端实现建议

1. **Prompt 设计**：把第三节 context 结构 + 第五节 actions schema 作为 system prompt 喂给大模型，要求"只输出符合 schema 的 JSON actions 数组，不要输出其他"。
2. **few-shot**：放 3-5 组"用户指令→actions"样例（第七节就是一组），显著提升解析准确率。
3. **JSON 修复**：模型偶发输出非法 JSON，建议用 json5 或带 repair 的解析库兜底。
4. **多 action**：一条用户指令可拆成多个 action（如示例），前端按序执行。
5. **降级**：模型不可用/超限时返回 `code:503`，前端自动回退本地正则，不报错。
6. **安全**：前端已做公式白名单+危险函数拦截+行列数截断，但建议后端也做一层校验（range 越界、count 超限等）。
