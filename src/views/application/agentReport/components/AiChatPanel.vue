<template>
  <div v-if="visible" class="ai-chat-wrapper">
    <!-- 收起态：悬浮圆按钮 -->
    <div
      v-if="collapsed"
      class="ai-fab"
      :style="fabStyle"
      @click="expand"
      title="展开 AI 助手"
    >
      <i class="el-icon-chat-dot-round"></i>
    </div>

    <!-- 展开态：完整面板 -->
    <div v-else class="ai-panel" :style="panelStyle">
      <!-- 标题栏（可拖拽） -->
      <div class="ai-panel__header" @mousedown="onHeaderMouseDown">
        <span class="ai-panel__title">
          <i class="el-icon-magic-stick"></i>
          AI 助手
        </span>
        <span class="ai-panel__source" v-if="lastSource">{{ lastSourceLabel }}</span>
        <div class="ai-panel__actions" @mousedown.stop>
          <el-tooltip content="本地模式：纯前端解析，无需后端" placement="bottom">
            <el-switch
              :value="useMock"
              @change="onMockChange"
              active-text="本地"
              inactive-text=""
              :width="46"
            />
          </el-tooltip>
          <el-tooltip content="清空对话" placement="bottom">
            <el-button
              type="text"
              icon="el-icon-delete"
              class="ai-panel__btn"
              @click="clearMessages"
            />
          </el-tooltip>
          <el-tooltip content="收起" placement="bottom">
            <el-button
              type="text"
              icon="el-icon-minus"
              class="ai-panel__btn"
              @click="collapse"
            />
          </el-tooltip>
          <el-tooltip content="关闭" placement="bottom">
            <el-button
              type="text"
              icon="el-icon-close"
              class="ai-panel__btn"
              @click="close"
            />
          </el-tooltip>
        </div>
      </div>

      <!-- 消息列表 -->
      <div ref="body" class="ai-panel__body">
        <div v-if="messages.length === 0" class="ai-empty">
          <i class="el-icon-chat-line-square"></i>
          <p>输入中文指令操纵表格，例如：</p>
          <el-button
            v-for="q in quickCommands"
            :key="q"
            type="text"
            class="ai-empty__quick"
            @click="sendQuick(q)"
          >{{ q }}</el-button>
        </div>

        <div
          v-for="(msg, i) in messages"
          :key="i"
          class="ai-msg"
          :class="msg.role === 'user' ? 'ai-msg--user' : 'ai-msg--ai'"
        >
          <div class="ai-msg__bubble" :class="{ 'ai-msg__bubble--error': msg.error }">
            <pre class="ai-msg__content">{{ msg.content }}</pre>
          </div>
          <!-- AI 消息的操作日志卡片 -->
          <div
            v-if="msg.role === 'assistant' && msg.execResults && msg.execResults.length"
            class="exec-card"
          >
            <div class="exec-card__summary">
              共 {{ msg.execResults.length }} 条 ·
              <span class="exec-ok">成功 {{ msg.summary.ok }}</span> ·
              <span class="exec-fail" v-if="msg.summary.failed">失败 {{ msg.summary.failed }}</span>
              <span v-if="msg.notes && msg.notes.length" class="exec-notes">（{{ msg.notes.join('；') }}）</span>
            </div>
            <div
              v-for="(r, ri) in msg.execResults"
              :key="ri"
              class="exec-item"
            >
              <i
                :class="r.ok ? 'el-icon-success' : 'el-icon-error'"
                :style="{ color: r.ok ? '#67c23a' : '#f56c6c' }"
              ></i>
              <span class="exec-type">{{ r.type }}</span>
              <span class="exec-range" v-if="r.range">{{ r.range }}</span>
              <span class="exec-error" v-if="!r.ok">：{{ r.error }}</span>
              <span class="exec-note" v-if="r.note">（{{ r.note }}）</span>
              <!-- 图表结果 -->
              <div
                v-if="r.result && r.result.chartOption"
                class="exec-chart"
              >
                <chart-card :option="r.result.chartOption" height="280px" />
              </div>
              <!-- 读取类结果表格 -->
              <table
                v-else-if="r.result && r.result.values && Array.isArray(r.result.values)"
                class="exec-table"
              >
                <tr v-for="(row, rri) in r.result.values.slice(0, 8)" :key="rri">
                  <td v-for="(cell, cci) in row" :key="cci">{{ cell === null || cell === undefined ? '' : cell }}</td>
                </tr>
              </table>
              <div
                v-else-if="r.result && r.result.selection"
                class="exec-result-text"
              >
                当前选区：{{ r.result.selection.a1 || '(无)' }}
              </div>
              <div
                v-else-if="r.result && r.result.merges"
                class="exec-result-text"
              >
                合并区域 {{ r.result.merges.length }} 个：{{ r.result.merges.map(m=>m.a1||'?').join(', ') }}
              </div>
            </div>
          </div>
        </div>

        <div v-if="loading" class="ai-loading">
          <i class="el-icon-loading"></i>
          <span>AI 思考中{{ execProgress ? ' (' + execProgress + ')' : '' }}...</span>
        </div>
      </div>

      <!-- 输入区 -->
      <div class="ai-panel__footer">
        <el-input
          v-model="input"
          type="textarea"
          :rows="2"
          resize="none"
          :placeholder="workbook ? '输入指令，Ctrl+Enter 发送' : '表格引擎尚未就绪'"
          :disabled="!workbook || loading"
          @keydown.native.ctrl.enter="send"
          @keydown.native.meta.enter="send"
        />
        <div class="ai-panel__send-row">
          <div class="ai-panel__quicks">
            <el-button
              v-for="q in quickCommands"
              :key="q"
              type="text"
              size="mini"
              class="ai-quick-btn"
              @click="sendQuick(q)"
            >{{ q }}</el-button>
          </div>
          <el-button
            type="primary"
            size="mini"
            icon="el-icon-s-promotion"
            :loading="loading"
            :disabled="!workbook || !input.trim()"
            @click="send"
          >发送</el-button>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
import { sendChat, getHistory } from '@/api/ai/chat'
import { collectContext } from '../univer-context'
import { executeActions } from '../univer-action-executor'
import ChartCard from './ChartCard.vue'

const QUICK_COMMANDS = [
  '把A1改成100',
  '清空B2:D10',
  '合并A1:C1',
  'B2:B5求和填到D1',
  '撤销'
]

const STORAGE_POS = 'aiPanel.pos'
const STORAGE_COLLAPSED = 'aiPanel.collapsed'
const STORAGE_CONVERSATION = 'aiPanel.conversationId'

export default {
  name: 'AiChatPanel',
  components: { ChartCard },
  props: {
    univerAPI: { type: Object, default: null },
    workbook: { type: Object, default: null },
    visible: { type: Boolean, default: true },
    useMock: { type: Boolean, default: false },
    allowMockFallback: { type: Boolean, default: true }
  },
  data() {
    return {
      collapsed: false,
      panelPos: { x: 80, y: 120 },
      messages: [],
      history: [],
      conversationId: null,
      input: '',
      loading: false,
      execProgress: '',
      dragging: false,
      dragStart: null,
      lastSource: 'mock-fallback',
      quickCommands: QUICK_COMMANDS
    }
  },
  computed: {
    panelStyle() {
      return {
        transform: 'translate3d(' + this.panelPos.x + 'px, ' + this.panelPos.y + 'px, 0)'
      }
    },
    fabStyle() {
      return {
        transform: 'translate3d(' + this.panelPos.x + 'px, ' + this.panelPos.y + 'px, 0)'
      }
    },
    lastSourceLabel() {
      if (!this.lastSource) return ''
      if (this.lastSource === 'mock') return '本地模式'
      if (this.lastSource === 'mock-fallback') return '兜底模式'
      return ''
    }
  },
  watch: {
    messages: {
      handler() {
        this.$nextTick(this.scrollToBottom)
      },
      deep: true
    },
    loading() {
      this.$nextTick(this.scrollToBottom)
    }
  },
  mounted() {
    this.restoreState()
    this.loadHistory()
  },
  beforeDestroy() {
    this.unbindDrag()
  },
  created() {},
  methods: {
    // ============ 状态控制 ============
    collapse() {
      this.collapsed = true
      this.persistState()
    },
    expand() {
      this.collapsed = false
      this.persistState()
    },
    close() {
      this.$emit('update:visible', false)
    },
    onMockChange(v) {
      this.$emit('update:useMock', v)
    },
    clearMessages() {
      this.messages = []
      this.history = []
      this.conversationId = null
      this.lastSource = null
      try { localStorage.removeItem(STORAGE_CONVERSATION) } catch (e) { /* ignore */ }
    },

    // ============ 拖拽 ============
    onHeaderMouseDown(e) {
      if (this.collapsed) return
      // 点击按钮区域不拖拽（按钮已 stop，这里再兜底）
      if (e.target.closest('.ai-panel__actions')) return
      this.dragging = true
      this.dragStart = {
        mx: e.clientX,
        my: e.clientY,
        x: this.panelPos.x,
        y: this.panelPos.y
      }
      document.addEventListener('mousemove', this.onDragMove)
      document.addEventListener('mouseup', this.onDragEnd)
      document.body.style.userSelect = 'none'
    },
    onDragMove(e) {
      if (!this.dragging) return
      const dx = e.clientX - this.dragStart.mx
      const dy = e.clientY - this.dragStart.my
      let nx = this.dragStart.x + dx
      let ny = this.dragStart.y + dy
      // 边界约束：至少 80px 宽度、整个高度保留在视口内
      const maxX = window.innerWidth - 80
      const maxY = window.innerHeight - 60
      nx = Math.max(-300, Math.min(maxX, nx))
      ny = Math.max(0, Math.min(maxY, ny))
      this.panelPos = { x: nx, y: ny }
    },
    onDragEnd() {
      if (!this.dragging) return
      this.dragging = false
      this.unbindDrag()
      document.body.style.userSelect = ''
      this.persistState()
    },
    unbindDrag() {
      document.removeEventListener('mousemove', this.onDragMove)
      document.removeEventListener('mouseup', this.onDragEnd)
    },

    // ============ 持久化 ============
    persistState() {
      try {
        localStorage.setItem(STORAGE_POS, JSON.stringify(this.panelPos))
        localStorage.setItem(STORAGE_COLLAPSED, this.collapsed ? '1' : '0')
      } catch (e) {
        // ignore
      }
    },
    restoreState() {
      try {
        const pos = localStorage.getItem(STORAGE_POS)
        if (pos) {
          const p = JSON.parse(pos)
          if (p && typeof p.x === 'number') this.panelPos = p
        }
        this.collapsed = localStorage.getItem(STORAGE_COLLAPSED) === '1'
        const cid = localStorage.getItem(STORAGE_CONVERSATION)
        if (cid) this.conversationId = cid
      } catch (e) {
        // ignore
      }
    },

    // ============ 发送 ============
    sendQuick(text) {
      this.input = text
      this.send()
    },
    async send() {
      const userText = (this.input || '').trim()
      if (!userText || this.loading || !this.workbook) return
      this.input = ''
      this.messages.push({ role: 'user', content: userText })
      this.loading = true
      this.execProgress = ''
      try {
        const context = collectContext(this.workbook)
        const resp = await sendChat(
          {
            message: userText,
            conversationId: this.conversationId,
            history: this.history,
            context: context,
            clientMeta: { locale: 'zh-CN', univerVersion: '0.25.1' }
          },
          { useMock: this.useMock, allowMockFallback: this.allowMockFallback }
        )
        if (resp && resp.source) this.lastSource = resp.source
        if (resp && resp.data && resp.data.conversationId) {
          this.conversationId = resp.data.conversationId
          try { localStorage.setItem(STORAGE_CONVERSATION, this.conversationId) } catch (e) { /* ignore */ }
        }
        const actions = (resp && resp.data && resp.data.actions) || []
        const exec = executeActions(this.univerAPI, this.workbook, actions, {
          onProgress: (i, total) => {
            this.execProgress = i + 1 + '/' + total
          }
        })
        this.messages.push({
          role: 'assistant',
          content: (resp && resp.data && resp.data.reply) || '(无回复)',
          execResults: exec.results,
          summary: exec.summary,
          notes: exec.notes
        })
        // 维护多轮历史
        this.history = truncateHistory(
          this.history,
          userText,
          (resp && resp.data && resp.data.reply) || '',
          actions,
          6
        )
        this.$emit('executed', exec)
      } catch (e) {
        this.messages.push({
          role: 'assistant',
          content: '执行失败：' + (e && e.message ? e.message : String(e)),
          error: true
        })
      } finally {
        this.loading = false
        this.execProgress = ''
      }
    },

    // ============ 历史记录 ============
    async loadHistory() {
      if (!this.conversationId) return
      try {
        const data = await getHistory(this.conversationId)
        if (data && data.code === 200 && Array.isArray(data.data) && data.data.length > 0) {
          // 只回显 role + content，不重放 actions（历史操作已执行过）
          this.messages = data.data.map((m) => ({
            role: m.role,
            content: m.content || ''
          }))
          this.$nextTick(this.scrollToBottom)
        }
      } catch (e) {
        // 后端未就绪或网络异常时静默忽略，不打扰用户
      }
    },

    // ============ 工具 ============
    scrollToBottom() {
      const body = this.$refs.body
      if (body) body.scrollTop = body.scrollHeight
    }
  }
}

/**
 * 截断历史到最近 maxRounds 轮（一轮 = user + assistant）
 */
function truncateHistory(history, userText, reply, actions, maxRounds) {
  const next = history.slice(-maxRounds * 2)
  next.push({ role: 'user', content: userText })
  next.push({ role: 'assistant', content: reply, actions: actions })
  return next.slice(-maxRounds * 2)
}
</script>

<style lang="scss" scoped>
.ai-chat-wrapper {
  position: fixed;
  top: 0;
  left: 0;
  z-index: 2000;
  pointer-events: none;
}

.ai-fab,
.ai-panel {
  pointer-events: auto;
  position: fixed;
  top: 0;
  left: 0;
}

.ai-fab {
  width: 52px;
  height: 52px;
  border-radius: 26px;
  background: linear-gradient(135deg, #409eff, #66b1ff);
  color: #fff;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 24px;
  cursor: pointer;
  box-shadow: 0 4px 16px rgba(64, 158, 255, 0.45);
  transition: transform 0.15s;
  &:hover {
    transform: translate3d(var(--x, 0), var(--y, 0), 0) scale(1.06);
  }
}

.ai-panel {
  width: 380px;
  height: 520px;
  max-height: calc(100vh - 40px);
  background: #fff;
  border-radius: 10px;
  box-shadow: 0 8px 28px rgba(0, 0, 0, 0.18);
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.ai-panel__header {
  cursor: move;
  background: linear-gradient(135deg, #409eff, #66b1ff);
  color: #fff;
  padding: 8px 12px;
  display: flex;
  align-items: center;
  gap: 8px;
  user-select: none;
  flex-shrink: 0;
}
.ai-panel__title {
  font-weight: 600;
  font-size: 14px;
  flex: 1;
  i {
    margin-right: 4px;
  }
}
.ai-panel__source {
  font-size: 11px;
  background: rgba(255, 255, 255, 0.25);
  padding: 1px 6px;
  border-radius: 8px;
}
.ai-panel__actions {
  display: flex;
  align-items: center;
  gap: 2px;
  cursor: default;
}
.ai-panel__btn {
  color: #fff !important;
  padding: 4px !important;
  &:hover {
    color: #fff !important;
    background: rgba(255, 255, 255, 0.18) !important;
  }
}
::v-deep .ai-panel__actions .el-switch__label {
  color: #fff;
}
::v-deep .ai-panel__actions .el-switch__label.is-active {
  color: #fff;
}

.ai-panel__body {
  flex: 1;
  overflow-y: auto;
  padding: 12px;
  background: #f7f8fa;
}

.ai-empty {
  text-align: center;
  color: #909399;
  padding: 24px 8px;
  i {
    font-size: 36px;
    color: #c0c4cc;
  }
  p {
    margin: 8px 0 12px;
    font-size: 13px;
  }
  &__quick {
    display: block;
    margin: 4px auto;
    padding: 4px 10px !important;
    font-size: 12px;
  }
}

.ai-msg {
  margin-bottom: 12px;
  display: flex;
  flex-direction: column;
}
.ai-msg--user {
  align-items: flex-end;
}
.ai-msg--ai {
  align-items: flex-start;
}
.ai-msg__bubble {
  max-width: 88%;
  padding: 8px 12px;
  border-radius: 10px;
  font-size: 13px;
  line-height: 1.5;
  word-break: break-word;
}
.ai-msg--user .ai-msg__bubble {
  background: #409eff;
  color: #fff;
  border-bottom-right-radius: 2px;
}
.ai-msg--ai .ai-msg__bubble {
  background: #fff;
  color: #303133;
  border: 1px solid #e4e7ed;
  border-bottom-left-radius: 2px;
}
.ai-msg__bubble--error {
  background: #fef0f0 !important;
  color: #f56c6c !important;
  border-color: #fbc4c4 !important;
}
.ai-msg__content {
  margin: 0;
  white-space: pre-wrap;
  word-break: break-word;
  font-family: inherit;
  font-size: inherit;
}

.exec-card {
  margin-top: 6px;
  max-width: 100%;
  background: #f0f9ff;
  border: 1px solid #d9ecff;
  border-radius: 6px;
  padding: 6px 8px;
  font-size: 12px;
}
.exec-card__summary {
  color: #606266;
  margin-bottom: 4px;
}
.exec-ok {
  color: #67c23a;
}
.exec-fail {
  color: #f56c6c;
}
.exec-notes {
  color: #e6a23c;
}
.exec-item {
  padding: 2px 0;
  color: #606266;
  i {
    margin-right: 4px;
  }
}
.exec-type {
  font-weight: 600;
  color: #303133;
  margin-right: 4px;
}
.exec-range {
  color: #409eff;
  margin-right: 4px;
  font-family: Consolas, Monaco, monospace;
}
.exec-error {
  color: #f56c6c;
}
.exec-note {
  color: #e6a23c;
}
.exec-table {
  margin-top: 4px;
  border-collapse: collapse;
  max-width: 100%;
  td {
    border: 1px solid #d9ecff;
    padding: 2px 6px;
    font-size: 11px;
    color: #303133;
    max-width: 80px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
}
.exec-result-text {
  margin-top: 4px;
  color: #409eff;
}
.exec-chart {
  margin-top: 6px;
  padding: 6px;
  background: #fff;
  border: 1px solid #ebeef5;
  border-radius: 6px;
}

.ai-loading {
  text-align: center;
  color: #909399;
  font-size: 12px;
  padding: 8px;
  i {
    margin-right: 4px;
  }
}

.ai-panel__footer {
  border-top: 1px solid #ebeef5;
  padding: 8px;
  background: #fff;
  flex-shrink: 0;
}
.ai-panel__send-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-top: 6px;
  gap: 8px;
}
.ai-panel__quicks {
  flex: 1;
  display: flex;
  flex-wrap: wrap;
  gap: 2px;
  overflow: hidden;
}
.ai-quick-btn {
  font-size: 11px !important;
  padding: 2px 6px !important;
  color: #409eff !important;
}
</style>
