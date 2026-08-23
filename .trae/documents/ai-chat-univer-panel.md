# AI 对话框操控 Univer 表格 — 实现方案

## Context（为什么做这个改动）

当前 `src/views/application/agentReport/index.vue` 已用 @univerjs/presets@0.25.1 实现了 Excel 导入/导出/新建/清空。用户希望在此基础上增加一个 **AI 对话框**，让用户用自然语言直接操纵线上表格（如"把 A1 改成 100""B 列求和填到 D1""合并 A1:C1"），把 univer 从"纯手工编辑表格"升级为"对话式表格"。

项目现状：RuoYi Cloud Vue2（Vue2 + element-ui 2.15.14 + Webpack4），API 走 `src/utils/request.js`（baseURL=/dev-api，timeout=10s，有全局 401/500/601 拦截弹窗），proxy 转发到 `localhost:8080`。**项目前后端均无任何 AI 集成**。

用户已拍板的 4 个关键决策：
1. **AI 来源**：后端网关代理（前端调 `POST /dev-api/ai/chat`，后端转发大模型，Key 存后端）；后端不在本仓库，前端给接口契约 + 本地 mock 兜底（后端未就绪也能演示）。
2. **操纵方式**：AI 返回结构化 JSON 指令，前端解析执行（安全可控）。
3. **流式**：不需要，一次性返回（复用 axios 即可，但需独立实例避开 request.js 的 10s 超时和全局弹窗）。
4. **布局**：表格全屏 + 悬浮可拖拽收起的 AI 对话面板。

预期结果：打开 agentReport 页面后，表格占满主区域，右下角悬浮一个 AI 助手面板，输入中文指令即可操作表格，并在面板内显示执行结果（成功/失败/读取到的数据）。

## 与 univer 官方 AI 能力的对比与借鉴（已调研）

调研 univer 官方 AI 能力后确认：官方**没有**"第三方 Vue2 web 页面内嵌对话框操控当前 univer 实例"的现成方案。官方能力盘点：

| 官方能力 | 适用性 |
|---|---|
| Univer MCP（`dream-num/univer-mcp`，云端 `https://mcp.univer.ai/mcp/` + API key） | ❌ 架构是"IDE agent → 云端 MCP → 独立 univer 进程（Playground/start-kit）"，操控的不是当前页面表格；要复用得自建 MCP server + 浏览器 WebSocket 桥 + 把 mcp-bridge 硬塞进 Webpack4，成本远超自研 |
| Univer Go（桌面应用 AI 模板） | ❌ 桌面端闭源应用，不可 web 嵌入 |
| Headless + Facade API for LLM（官方推荐思路） | ✅ **思路可用**——印证了"LLM → 工具调用 → Facade API"是官方推荐方向 |
| univer-mcp-start-kit（React/Vite） | ❌ 独立 React 项目，要替换现有 univer 初始化；其 agent-ui 分支已 revert |

**本方案的定位**：本质上就是官方 mcp-bridge 思路的**单页面轻量版**——把"LLM 工具调用 → Facade API 操控表格"这条链路在浏览器内闭环实现，不依赖云端 MCP server / API key / 独立进程。借鉴点：
1. 官方 mcp-bridge 把表格操作封装成工具供 LLM 调用 → 本方案第 3 节的 JSON 指令 schema + 第 4 节执行器就是这个角色（每个 action type = 一个工具）。
2. 官方推荐 LLM 直接调 Facade API → 本方案执行器全部基于已确认的 FRange/FWorksheet/FWorkbook facade API（setValue/merge/insertRows/undo 等）。
3. 若未来要升级到标准 MCP 协议对接外部 agent，本方案的 schema 与执行器可平滑外暴露为 MCP tools（actions 结构与 MCP tool call 同构）。

---

## 1. 文件结构

### 新增
| 路径 | 职责 |
|---|---|
| `src/api/ai/chat.js` | AI 接口封装 + 独立 axios 实例 + mock 回退 |
| `src/views/application/agentReport/ai-action-schema.js` | 指令白名单、A1 解析、值转换、校验、越界截断、公式安全校验 |
| `src/views/application/agentReport/univer-context.js` | 表格上下文采集（紧凑，最多 50×26） |
| `src/views/application/agentReport/univer-action-executor.js` | 指令执行器纯函数（actions+univerAPI→结果日志） |
| `src/views/application/agentReport/ai-mock-parser.js` | 本地中文指令→actions 关键词/正则解析器 |
| `src/views/application/agentReport/components/AiChatPanel.vue` | 悬浮可拖拽收起的对话面板（消息气泡内联，不另拆 ChatMessage） |

### 修改
| 路径 | 修改点 |
|---|---|
| `src/views/application/agentReport/index.vue` | 引入 AiChatPanel；`.univer-wrapper` 高度改占满；data 加 `aiPanelVisible`/`aiUseMock`；新增 `onAiExecuted` 更新 statusMessage |

**不改**：request.js、vue.config.js、main.js、permission.js、store/*。新功能自包含在 agentReport 子树 + src/api/ai/。

---

## 2. 后端接口契约（可直接发后端）

- **路径**：`POST /ai/chat`（前端实际 `/dev-api/ai/chat`，网关剥前缀），需登录态，`Authorization: Bearer <token>`，Content-Type: json，**非流式**，整体响应 ≤30s。
- **请求体**：
  ```jsonc
  {
    "message": "把A1改成100",          // 必填
    "conversationId": "uuid",          // 可选，多轮会话id
    "history": [                       // 可选，最近6轮 [{role:'user'|'assistant', content, actions?}]
      {"role":"user","content":"清空B2:D10"},
      {"role":"assistant","content":"已清空","actions":[...]}
    ],
    "context": {                       // 必填，表格状态摘要
      "workbookName":"销售月报.xlsx",
      "sheetName":"Sheet1","sheetId":"sheet-0",
      "rowCount":100,"columnCount":26,
      "usedRange":{"startRow":0,"endRow":12,"startColumn":0,"endColumn":5},
      "values":[["姓名","语文",...],["张三",90,...]],   // usedRange内二维裸值，null=空
      "formulas":[[null,...],["=SUM(B2:D2)",...]],      // 无公式时该字段省略
      "merges":[{"startRow":0,"endRow":0,"startColumn":0,"endColumn":5}],
      "selection":{"a1":"B2:D5","startRow":1,"endRow":4,"startColumn":1,"endColumn":3},
      "activeCell":{"row":1,"column":1,"a1":"B2"}
    },
    "clientMeta":{"locale":"zh-CN","univerVersion":"0.25.1","truncated":false}
  }
  ```
- **响应体**（沿用 ruoyi `{code,msg,data}` 外壳）：
  ```jsonc
  {"code":200,"msg":"ok","data":{
    "reply":"已把A1设为100，并合并B1:C1。",   // 给用户的自然语言
    "conversationId":"uuid",
    "actions":[                              // 结构化指令，schema见第3节
      {"type":"setCell","range":"A1","value":100},
      {"type":"merge","range":"B1:C1","mode":"default"}
    ],
    "needFeedback":false
  }}
  ```
- **错误码**：200 成功；400 入参非法；401 未登录；429 限流；500 模型超时/上游错误；501 模型识别但不支持；503 服务不可用（前端自动回退 mock）。
- **约束**：`actions[].type` 必须命中前端白名单（第3节），不在白名单的指令前端跳过并在 reply 提示；后端模型 prompt 由后端维护，只产出符合 schema 的 JSON。

---

## 3. JSON 指令 Schema

每个 action 是 plain object `{type, range?, ...}`，type 必须命中白名单。

| type | 必填参数 | 可选参数 | 映射 facade API |
|---|---|---|---|
| `setCell` | range, value | isFormula | `range.setValue(value)`（以=开头识别为公式） |
| `setCellForCell` | range(单格), value | — | `range.setValueForCell(value)` |
| `setValues` | range, values(二维) | — | `range.setValues(values)` 维度须匹配 |
| `clearRange` | range | clearWhat("content"\|"format"\|"all",默认content) | content:setValue(null)；format:重置样式+breakApart |
| `setFormula` | range, formula | — | `range.setValue("=" + formula)` |
| `merge` | range | mode("default"\|"across"\|"vertically"), force | `range.merge/mergeAcross/mergeVertically({defaultMerge:true,isForceMerge:force})` |
| `breakApart` | range | — | `range.breakApart()` |
| `insertRows` | rowIndex, count=1 | position("at"\|"after"\|"before") | `insertRows/insertRowsAfter/insertRowsBefore` |
| `deleteRows` | rowPosition, count=1 | — | `deleteRows(rowPosition, count)` |
| `insertColumns` | columnIndex, count=1 | position("at"\|"after") | `insertColumns/insertColumnAfter` |
| `deleteColumns` | columnPosition, count=1 | — | `deleteColumns(columnPosition, count)` |
| `setRowHeight` | rowPosition, height | — | `setRowHeight(r,h)` |
| `setColumnWidth` | columnPosition, width | — | `setColumnWidth(c,w)` |
| `setStyle` | range | background,fontColor,fontSize,fontFamily,fontWeight,hAlign,vAlign,wrap,textRotation | 链式调用对应 setter |
| `undo` | — | — | `workbook.undo()` |
| `redo` | — | — | `workbook.redo()` |
| `getRangeValue` | range | — | **读取类**，`range.getValues()` 返回二维裸值 |
| `getRangeFormulas` | range | — | `range.getFormulas()` |
| `getMerges` | — | — | `sheet.getMergeData()` |
| `getSelection` | — | — | `sheet.getActiveRange()`+`getA1Notation()` |
| `sumToCell` | sourceRange, targetRange(单格) | — | 语义糖，转 `target.setValueForCell("=SUM(src)")` |

**白名单外（一律跳过）**：删 sheet、保护/解除保护、宏、`HYPERLINK(http...)`/`WEBSERVICE`/`IMPORTDATA`/`IMPORTXML`/`IMAGE` 等含外部链接/网络的公式。

**读取类回传**：执行结果对象含 `result` 字段，渲染为表格卡片挂在 AI 气泡下方；摘要拼入下一轮 history 让 AI 基于结果继续对话。

---

## 4. 执行器（`univer-action-executor.js`）

纯函数 + 副作用隔离。导出：
```js
executeActions(univerAPI, actions, options) → { results:[ActionResult], summary:{ok,failed,skipped}, notes:[string] }
executeAction(univerAPI, action, options) → ActionResult   // {type,range,ok,error,result,note,durationMs}
resolveSheet(univerAPI, sheetName?) → FWorksheet
resolveRange(sheet, rangeStr) → FRange
```
options: `{ sheetName?, onProgress(i,total,action), continueOnError=true }`。单次 actions 上限 30。

**派发表**：`handlers = { setCell: handleSetCell, merge: handleMerge, ... }`。executeAction 流程：查 handler → `validateAction`（白名单+参数+公式安全）→ 解析 sheet → 调 handler → try-catch 包裹 → 返回 ActionResult。

**代表性 handler**（伪代码）：
```js
function handleSetCell(univerAPI, sheet, action) {
  const { range: clamped, changed } = clampRange(parseA1(action.range), 0, sheet.getRowCount()-1, 0, sheet.getColumnCount()-1)
  const range = sheet.getRange(clamped.startRow, clamped.startColumn, clamped.endRow-clamped.startRow+1, clamped.endColumn-clamped.startColumn+1)
  let value = action.value
  if (action.isFormula || (typeof value==='string' && value.startsWith('='))) {
    value = String(value).replace(/^=?/, '=')
    assertSafeFormula(value)   // 拒绝 HYPERLINK/WEBSERVICE 等
  }
  range.setValue(value)
  return changed ? { note: 'range 已截断到合法边界' } : null
}

function handleMerge(univerAPI, sheet, action) {
  const range = resolveRange(sheet, action.range)
  const opts = { defaultMerge:true, isForceMerge:!!action.force }
  if (action.mode==='across') range.mergeAcross(opts)
  else if (action.mode==='vertically') range.mergeVertically(opts)
  else range.merge(opts)
}

function handleClearRange(univerAPI, sheet, action) {
  const range = resolveRange(sheet, action.range)
  const what = action.clearWhat || 'content'
  if (what==='content'||what==='all') range.setValue(null)
  if (what==='format'||what==='all') {
    range.setBackgroundColor('#ffffff').setFontColor('#000000').setFontSize(11)
         .setFontWeight('normal').setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(false)
    if (range.isPartOfMerge()) range.breakApart()
  }
}

function handleGetRangeValue(univerAPI, sheet, action) {
  const range = resolveRange(sheet, action.range)
  const values = range.getValues().map(row => row.map(v => (v===null||v===undefined)?null : (typeof v==='object'?(v.v??null):v)))
  return { values, a1: action.range }
}

function handleUndo(univerAPI, sheet, action) { univerAPI.getActiveWorkbook().undo() }
```

---

## 5. AI API 封装 + Mock（`src/api/ai/chat.js` + `ai-mock-parser.js`）

### chat.js — 独立 axios 实例（关键：避开 request.js）
```js
import axios from 'axios'
import { getToken } from '@/utils/auth'
import { mockParse } from '@/views/application/agentReport/ai-mock-parser'

const aiHttp = axios.create({
  baseURL: process.env.VUE_APP_BASE_API,  // /dev-api
  timeout: 60000,                          // 60s，覆盖 request.js 的 10s
  headers: { 'Content-Type':'application/json;charset=utf-8' }
})
aiHttp.interceptors.request.use(config => {
  const token = getToken()
  if (token) config.headers['Authorization'] = 'Bearer ' + token
  return config
})
aiHttp.interceptors.response.use(res => res.data, err => Promise.reject(err))  // 不挂全局弹窗

export async function sendChat(payload, opts={}) {
  if (opts.useMock) return mockToResponse(payload, mockParse(payload.message, payload.context))
  try {
    const data = await aiHttp.post('/ai/chat', payload)
    if (data && data.code===200) return data
    if (data && data.code===503 && opts.allowMockFallback)
      return mockToResponse(payload, mockParse(payload.message, payload.context), 'mock-fallback')
    const err = new Error((data&&data.msg)||'AI接口返回异常'); err.code = data&&data.code; throw err
  } catch (e) {
    if (opts.allowMockFallback && isNetworkError(e))
      return mockToResponse(payload, mockParse(payload.message, payload.context), 'mock-fallback')
    throw e
  }
}
```
- `opts.useMock`：用户在面板勾"本地模式"强制走 mock。
- `opts.allowMockFallback` 默认 true：后端 404/网络异常/503 自动降级 mock。

### ai-mock-parser.js — 中文规则解析
`mockParse(message, context) → { reply, actions }`，纯函数，规则按优先级匹配（命中多条返回多 action）。核心规则：
- `撤销|undo` → `{type:'undo'}`；`重做|恢复|redo` → `{type:'redo'}`
- `清空\s*([A-Z]+\d+:[A-Z]+\d+)` → `{type:'clearRange',range}`；`清空第(\d+)行` → 清行；`清空第(\d+)列` → 清列
- `合并\s*([A-Z]+\d+:[A-Z]+\d+)` → `{type:'merge',range,mode:'default'}`；`横向合并...`→across；`拆分|取消合并` → breakApart
- `把\s*([A-Z]+\d+)\s*(?:改成?|改为|设为)\s*(.+)` → `{type:'setCell',range,value:coerceValue(v)}`（coerceValue：纯数字转 number，否则字符串）
- `([A-Z]+\d+(?::[A-Z]+\d+)?)\s*求和.{0,6}填到?\s*([A-Z]+\d+)` → `{type:'sumToCell',sourceRange,targetRange}`
- `...平均...填到...` → `{type:'setFormula',range:target,formula:'AVERAGE(src)'}`
- `([A-Z]+\d+:[A-Z]+\d+)\s*背景.{0,4}(红色|绿色|...)` → `{type:'setStyle',range,background:zhColorToEn(color)}`
- `...加粗` → fontWeight:'bold'；`...居中` → hAlign:'center'；`...字号\s*(\d+)` → fontSize
- `在第(\d+)行(前|后)插入(\d+)?行` → insertRows；`删除第(\d+)行(到第(\d+)行)?` → deleteRows（count上限5）
- `(?:读取|看一下|查询)\s*([A-Z]+\d+:[A-Z]+\d+)` → getRangeValue；`当前选(区|中)` → getSelection
- 未命中 → reply 提示可用指令示例，actions 为空。

**安全**：mock 规则不生成 deleteColumns；deleteRows count 截断到 5；只产出 SUM/AVERAGE 公式。

---

## 6. AiChatPanel.vue — 悬浮对话面板

### Props/Events
```js
props: {
  univerAPI: { type:Object, required:true },
  visible: { type:Boolean, default:true },
  useMock: { type:Boolean, default:false },
  allowMockFallback: { type:Boolean, default:true }
}
emits: { 'update:visible', 'executed'(result), 'context-needed' }
```
data：`panelPos{x:80,y:80}, size{w:380,h:520}, collapsed, messages[], history[], conversationId, input, loading, execProgress, dragging`

### UI 结构
```
.ai-panel (fixed, z-index:2000, transform:translate3d(x,y,0))
  ├ header (cursor:move, 拖拽区，标题+本地模式switch+收起+清空按钮)
  ├ body (overflow-y:auto, 消息列表)
  │   └ 用户气泡(右) / AI气泡(左) + 操作日志卡片(summary + 每条result + 读取结果table)
  └ footer (el-input textarea + 发送按钮 + 快捷指令按钮组)
.ai-panel--collapsed (收起态，56×56 圆形 FAB，点击展开)
```

### 拖拽（原生事件）
- header `mousedown` 记录起点 → `document.mousemove` 算 dx/dy → 边界约束（至少80px留在视口）→ `mouseup` 解绑 + `localStorage` 持久化位置。
- 用 `transform: translate3d(${x}px,${y}px,0)` 移动（性能优于改 left/top）。
- 收起前记忆位置 `fabPos`，展开时恢复；collapsed 状态也持久化。

### 消息渲染（轻量，不引入 markdown）
- `{{ msg.content }}` 插值（Vue 自动转义，安全）+ CSS `white-space:pre-wrap` 保留换行。
- AI 气泡下方挂操作日志卡片：summary（成功N/失败M）+ 每条 result（type/range/ok图标/error/note）+ 读取类结果渲染为小 HTML table（最多8行）。

### 发送流程
```js
async send() {
  if (!this.input.trim() || this.loading || !this.univerAPI) return
  const userText = this.input; this.input = ''
  this.messages.push({role:'user',content:userText}); this.loading = true
  try {
    const context = collectContext(this.univerAPI)
    const resp = await sendChat({message:userText, conversationId:this.conversationId, history:this.history, context},
                                {useMock:this.useMock, allowMockFallback:this.allowMockFallback})
    this.conversationId = resp.data.conversationId
    const exec = executeActions(this.univerAPI, resp.data.actions||[],
                                {onProgress:(i,t)=>this.execProgress=`${i+1}/${t}`})
    this.messages.push({role:'assistant', content:resp.data.reply, execResults:exec.results, summary:exec.summary, notes:exec.notes})
    this.history = truncateHistory(this.history, userText, resp.data.reply, resp.data.actions, 6)
    this.$emit('executed', exec)
  } catch(e) {
    this.messages.push({role:'assistant', content:'执行失败：'+e.message, error:true})
  } finally { this.loading=false; this.execProgress='' }
}
```
快捷指令按钮：`[把A1改成100][清空B2:D10][合并A1:C1][B2:B5求和填到D1][撤销]`，点击直接填入输入框或直接发送。

---

## 7. 父页面改造（`index.vue`）

- template：在 `.univer-container` 后追加 `<ai-chat-panel :univer-api="univerAPI" :visible.sync="aiPanelVisible" :use-mock.sync="aiUseMock" :allow-mock-fallback="true" @executed="onAiExecuted" />`
- 容器样式：`.univer-wrapper` 高度从 `calc(100vh - 320px)` 改为 `calc(100vh - 200px)` 让表格更高（腾出空间给悬浮面板，但面板 fixed 不挤占）。
- script：`components:{ AiChatPanel }`，data 加 `aiPanelVisible:true, aiUseMock:false`，methods 加：
  ```js
  onAiExecuted(exec) {
    const s = exec.summary
    this.statusMessage = `AI 执行完成：成功 ${s.ok} 条，失败 ${s.failed} 条` + (exec.notes.length?('；'+exec.notes.join('；')):'')
    if (s.failed > s.ok && s.failed > 0) this.$modal.msgWarning('部分 AI 指令执行失败，详情见对话框')
  }
  ```
- 时序：univerAPI 在 `mounted→$nextTick→initUniver` 后赋值，是响应式 prop，子组件自动接收；子组件在 univerAPI 为空时禁用输入并提示"表格引擎尚未就绪"。

---

## 8. 上下文采集（`univer-context.js`）

`collectContext(univerAPI) → context对象|null`。流程：
1. 取 `getActiveWorkbook().getActiveSheet()`（或 `getSheets()[0]`），取 rowCount/columnCount。
2. 探测 usedRange：`sheet.getRange(0,0,Math.min(rowCount,200),Math.min(colCount,50)).getValues()` 一次性取，找最后一个非空行列。
3. 截断到 `MAX_ROWS=50, MAX_COLS=26`，逐格 `safeGetCell`（getValue 取裸值，对象取 .v；getFormulas 取公式字符串；无公式则不发 formulas 字段）。
4. `sheet.getMergeData()` 映射为 `{startRow,endRow,startColumn,endColumn,a1}` 数组，最多 50 个。
5. `sheet.getActiveRange()`/`getActiveCell()` 取选区和活动单元格，`getA1Notation()` 转 A1。
6. 超出截断时 `truncated:true`，让后端模型知道还有更多数据。
7. 整体 context ≤30KB。

---

## 9. 安全与边界（在 `ai-action-schema.js`）

- **白名单**：`ACTION_WHITELIST` 数组；type 不命中 → 跳过并记 error。
- **A1 解析** `parseA1(str) → {startRow,endRow,startColumn,endColumn}|null`：支持 `A1`、`A1:B2`、`Sheet1!A1`（剥 sheet 前缀）；失败抛错。
- **越界截断** `clampRange(range,minR,maxR,minC,maxC)`：自动截到边界，记 note（非报错卡死，避免"Z999"卡住对话）；保证 start≤end。
- **数量上限**：insert/delete 的 count 上限 100；单次 actions 上限 30。
- **setValues 维度匹配**：执行器内解析 range 拿行列数后校验二维数组维度，不匹配跳过该 action（不阻断后续）。
- **公式安全** `assertSafeFormula(formula)`：黑名单 `HYPERLINK(http`/`WEBSERVICE`/`IMPORTDATA`/`IMPORTXML`/`IMAGE` 等；白名单函数 `SUM/AVERAGE/MAX/MIN/COUNT/COUNTIF/SUMIF/IF/IFS/VLOOKUP/HLOOKUP/INDEX/MATCH/CONCATENATE/TEXT/ROUND/LEFT/RIGHT/MID/LEN/TRIM/UPPER/LOWER/ABS/SQRT/MOD/INT/DATE/TODAY/NOW/YEAR/MONTH/DAY/IFERROR/ISBLANK/ISNUMBER/ROW/COLUMN/OFFSET` 等；强制以 `=` 开头。
- **颜色校验**：`^#([0-9a-fA-F]{6})$|^(red|green|blue|yellow|black|white|gray|orange|pink|purple)$`。
- **工具函数**：`coerceValue(str)`（纯数字→number，"true"/"false"→boolean，否则字符串）、`zhColorToEn(zh)`（红色→red、#rrggbb 原样）。

---

## 10. 验证

### 无后端 — mock 模式端到端
1. `npm run dev`（8080 没起也行，会自动回退 mock；或面板勾"本地模式"强制 mock）。
2. 打开 agentReport 页面，导入含示例数据的 Excel 或点"新建表格"。
3. 测试用例：
   | 输入 | 预期 |
   |---|---|
   | 把A1改成100 | A1=100 |
   | 把B2改成张三 | B2="张三" |
   | B2:B5求和填到D1 | D1==SUM(B2:B5) |
   | 清空B2:D10 | B2:D10 全空 |
   | 合并A1:C1 | A1:C1 合并 |
   | 把B2:B5背景设为红色 | B2:B5 背景#ff0000 |
   | A1:A3加粗 | A1:A3 bold |
   | 在第2行后插入3行 | 出现3空行 |
   | 删除第2行到第4行 | 一次性删3行 |
   | 撤销 | workbook.undo() 回退 |
   | 读取一下A1:C3 | AI气泡下出现结果表格 |
   | 当前选区是什么 | 返回选区A1 |
4. 边界：构造 AI 返回写 Z999 → 自动截断+note；返回 setFormula =WEBSERVICE(...) → 跳过+error；一次返回3个 setCell → 全执行 summary=3 ok。

### 有后端 — 全链路
1. 后端实现 `POST /ai/chat` 接大模型。
2. 关闭本地模式，用例同上但指令由模型生成。
3. 重点：多轮对话（"把A1改成100"→"再加50"应理解 A1=A1+50）；选区指代（选中 B2:B5 后说"把这些求和填到D1"）；大表截断（导入200行验证 context 只发50行+truncated:true）；错误码 429/500 的 UI 表现。

### 调试
- AiChatPanel 内 console.log payload 和 data.actions。
- 执行器每条 action 执行前后 console.log。
- 浏览器控制台 `univerAPI.getActiveWorkbook()._commandService` 观察 undo 栈。
- Vue Devtools 选中 AiChatPanel 改 input 数据驱动回放。

---

## 11. 实现顺序

1. `ai-action-schema.js`（白名单+parseA1+coerceValue+zhColorToEn+validateAction+assertSafeFormula+clampRange）
2. `univer-context.js`（采集函数）
3. `univer-action-executor.js`（依赖 schema，先跑通 setCell/merge/clearRange/undo 四条）
4. `ai-mock-parser.js`（依赖 schema）
5. `src/api/ai/chat.js`（依赖 mock-parser；独立 axios 实例；mock/远程分支）
6. `components/AiChatPanel.vue`（拼装上面四块+拖拽UI+消息渲染）
7. `index.vue` 改造（引入 panel、调容器高度、监听 executed）
8. 联调：先纯 mock 跑通用例，再接后端。

---

## 关键文件路径
- 修改：`g:\code\agentReport\ruoyi-cloude-vue2\src\views\application\agentReport\index.vue`
- 新增：`g:\code\agentReport\ruoyi-cloude-vue2\src\api\ai\chat.js`
- 新增：`g:\code\agentReport\ruoyi-cloude-vue2\src\views\application\agentReport\ai-action-schema.js`
- 新增：`g:\code\agentReport\ruoyi-cloude-vue2\src\views\application\agentReport\univer-context.js`
- 新增：`g:\code\agentReport\ruoyi-cloude-vue2\src\views\application\agentReport\univer-action-executor.js`
- 新增：`g:\code\agentReport\ruoyi-cloude-vue2\src\views\application\agentReport\ai-mock-parser.js`
- 新增：`g:\code\agentReport\ruoyi-cloude-vue2\src\views\application\agentReport\components\AiChatPanel.vue`

## 参考文件（不修改，实现时对照）
- `g:\code\agentReport\ruoyi-cloude-vue2\src\utils\request.js`（timeout/拦截器行为，刻意不复用）
- `g:\code\agentReport\ruoyi-cloude-vue2\src\utils\auth.js`（getToken）
- `g:\code\agentReport\ruoyi-cloude-vue2\src\plugins\modal.js`（$modal API）
- `g:\code\agentReport\ruoyi-cloude-vue2\node_modules\@univerjs\sheets\lib\es\facade.js`（FRange/FWorksheet/FWorkbook 方法签名逐一核对）
