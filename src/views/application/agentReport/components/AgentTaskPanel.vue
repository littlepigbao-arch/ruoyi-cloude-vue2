<template>
  <section v-if="enabled || recoveryAvailable || syncRequired" class="agent-task">
    <div class="agent-controls">
      <el-switch v-model="useAgent" active-text="计划执行" :disabled="busy || !snapshotWritesEnabled" @change="$emit('mode', $event)" />
      <el-checkbox v-if="useAgent && !busy" v-model="allowOverwrite">允许覆盖目标区域已有内容</el-checkbox>
    </div>
    <template v-if="run && run.runId">
      <strong>{{ run.goal }}</strong>
      <p>状态：{{ statusLabel }} · 计划 {{ run.planVersion || 0 }} · 模型调用 {{ run.modelCalls || 0 }} · {{ run.tokens || 0 }} tokens</p>
      <ol v-if="run.plan">
        <li v-for="step in run.plan.steps" :key="step.stepId">{{ label(step.status) }} {{ step.objective }}</li>
      </ol>
      <p v-if="run.plan && run.plan.changeReason">计划说明：{{ run.plan.changeReason }}</p>
      <p v-if="run.message">{{ run.message }}</p>
      <p>草稿：{{ run.persistenceStatus === 'SAVED' ? '已保存' : '尚未确认保存' }}；正式版本请在版本管理中主动提交。</p>
      <p v-if="run.completedSteps && run.completedSteps.length">已有 {{ run.completedSteps.length }} 个步骤生效或验证通过，停止任务不会撤销已保存结果。</p>
      <el-alert v-if="error" :title="error" type="warning" :closable="false" />
      <div v-if="run.question || run.status === 'WAITING_USER'">
        <p>{{ run.question }}</p>
        <el-input v-model="answer" type="textarea" placeholder="补充字段映射或业务口径" />
        <el-button size="mini" :disabled="!answer.trim() || operating" @click="sendAnswer">提交补充条件</el-button>
      </div>
      <el-button v-if="busy" size="mini" :disabled="operating" @click="stop">停止任务</el-button>
      <el-button v-if="syncRequired || ['PAUSED', 'CANCELLING'].includes(run.status)" size="mini" :disabled="operating" @click="recover">载入服务器草稿并核实停止</el-button>
      <el-button size="mini" type="text" @click="details = !details">{{ details ? '收起' : '查看' }}执行记录</el-button>
      <ul v-if="details"><li v-for="event in events" :key="event.sequence">{{ event.sequence }}. {{ event.message }}</li></ul>
    </template>
    <el-alert v-else-if="error" :title="error" type="warning" :closable="false" />
    <el-button v-if="syncRequired && !run" size="mini" :disabled="operating" @click="attach">重新核实工作区</el-button>
  </section>
</template>

<script>
import { agentCapabilities, agentRequest, activeAgent, getAgent } from '@/api/ai/agent'
import { createOperationId } from '../change-capture'
const TERMINAL = ['COMPLETED', 'CANCELLED', 'FAILED']
export default {
  name: 'AgentTaskPanel',
  props: { workbookId: { type: [Number, String], default: null }, workspace: { type: Object, required: true } },
  data: () => ({ enabled: false, recoveryAvailable: false, snapshotWritesEnabled: false, syncRequired: false, generation: 0,
    useAgent: false, allowOverwrite: false, run: null, events: [], error: '', answer: '', details: false,
    timer: null, heartbeatTimer: null, polling: false, operating: false, clientId: createOperationId(), pendingRequest: null }),
  computed: { activeRun() { return !!(this.run && this.run.runId && !TERMINAL.includes(this.run.status)) },
    busy() { return this.activeRun || this.syncRequired }, statusLabel() { return this.label(this.run.status) } },
  watch: {
    workbookId() { this.attach() },
    busy(value) { this.$emit('busy', value); if (value) this.workspace.lock(true) }
  },
  async mounted() {
    try { const c = await agentCapabilities(); this.enabled = c.enabled; this.recoveryAvailable = c.recoveryAvailable;
      this.snapshotWritesEnabled = c.snapshotWritesEnabled === true;
      if (this.enabled || this.recoveryAvailable) await this.attach() } catch (e) { this.error = '智能体服务暂不可用：' + e.message }
  },
  beforeDestroy() { this.generation++; clearTimeout(this.timer); clearInterval(this.heartbeatTimer) },
  methods: {
    label(s) { return ({ QUEUED: '排队', PLANNING: '制定计划', EXECUTING: '执行中', REPLANNING: '修订计划', VERIFYING: '验证', WAITING_USER: '等待补充', PAUSED: '暂停', CANCELLING: '核实停止', CANCELLED: '已停止', COMPLETED: '完成', FAILED: '失败', PENDING: '待执行', RUNNING: '执行中', SUCCEEDED: '已验证', SUPERSEDED: '被新计划替代' })[s] || s },
    async attach() {
      clearTimeout(this.timer); clearInterval(this.heartbeatTimer)
      const generation = ++this.generation
      this.run = null; this.events = []; this.error = ''
      if ((!this.enabled && !this.recoveryAvailable) || !this.workbookId) return
      const id = this.workbookId
      this.syncRequired = true; this.workspace.lock(true)
      try {
        const run = await activeAgent(id)
        if (generation !== this.generation || String(id) !== String(this.workbookId)) return
        this.run = run.runId ? run : null
        if (this.activeRun) { this.useAgent = true; this.$emit('mode', true); this.schedule() }
        else if (this.run) await this.synchronize(generation)
        else { this.syncRequired = false; this.workspace.lock(false); this.$emit('busy', false) }
      } catch (e) { if (generation === this.generation) this.error = e.message }
    },
    async synchronize(generation = this.generation) {
      this.syncRequired = true; this.workspace.lock(true)
      await this.workspace.reload()
      if (generation !== this.generation) throw new Error('页面已切换，保持保护状态')
      this.syncRequired = false
      if (!this.activeRun) { this.workspace.lock(false); this.$emit('busy', false) }
    },
    async start(goal, conversationId) {
      if (!this.enabled || !this.snapshotWritesEnabled || this.busy || this.operating) throw new Error('A1 已阻断完整 Snapshot 智能体写入，可核实或取消已有任务')
      this.operating = true; this.error = ''; this.workspace.lock(true)
      try {
        const state = await this.workspace.prepare()
        // Retain this exact request when creation response is lost. User resubmits only after reconciliation.
        this.pendingRequest = { requestId: createOperationId(), goal, conversationId, workbookId: this.workbookId,
          branchId: state.branchId, expectedRevision: state.revision, allowOverwrite: this.allowOverwrite }
        try { this.run = await agentRequest('/runs', this.pendingRequest) }
        catch (e) { this.run = await agentRequest('/runs', this.pendingRequest) }
        this.events = []; this.schedule()
      } catch (e) {
        this.error = e.message
        // Never unlock after an ambiguous create response until the server has been queried.
        try { const existing = await activeAgent(this.workbookId); this.run = existing.runId ? existing : null; if (this.activeRun) this.schedule(); else await this.synchronize() } catch (_) { this.syncRequired = true; this.error += '；任务创建结果未知，请重新打开文档核实'; }
        throw e
      } finally { this.operating = false }
    },
    schedule() {
      clearTimeout(this.timer); clearInterval(this.heartbeatTimer)
      this.heartbeatTimer = setInterval(() => {
        if (this.activeRun && !['PAUSED', 'CANCELLING'].includes(this.run.status)) agentRequest('/runs/' + this.run.runId + '/lease-heartbeat', { fencingToken: this.run.fencingToken }).catch(e => { this.error = e.message })
      }, 20000)
      this.timer = setTimeout(this.poll, 500)
    },
    async poll() {
      if (this.polling || !this.run || !this.run.runId) return
      this.polling = true
      const id = this.run.runId
      const generation = this.generation
      try {
        const latest = await getAgent(id)
        if (generation !== this.generation || !this.run || this.run.runId !== id) return
        this.syncRequired = true
        this.run = latest
        const after = this.events.length ? this.events[this.events.length - 1].sequence : 0
        const events = await agentRequest('/runs/' + id + '/events?after=' + after, null, 'get')
        if (generation !== this.generation || !this.run || this.run.runId !== id) return
        this.events = this.events.concat(events).slice(-200)
        if (this.snapshotWritesEnabled && this.run.status === 'EXECUTING' && this.run.call && this.run.call.status === 'PENDING') {
          const call = await agentRequest('/runs/' + id + '/tool-claims', { clientId: this.clientId, fencingToken: this.run.fencingToken })
          if (call.callId) {
            let receipt = { callId: call.callId, claimToken: call.claimToken, argsHash: call.argsHash, fencingToken: call.fencingToken, ok: false }
            try { receipt.snapshot = await this.workspace.apply(call.snapshot); receipt.ok = true; receipt.calculationVerified = true } catch (e) { this.error = e.message }
            // Only retry the identical receipt, never the browser operation.
            try { this.run = await agentRequest('/runs/' + id + '/tool-results', receipt) }
            catch (e) { this.run = await agentRequest('/runs/' + id + '/tool-results', receipt) }
            if (receipt.ok) await this.workspace.saved()
          }
        }
        if (!this.activeRun) { clearInterval(this.heartbeatTimer); await this.synchronize(generation) }
      } catch (e) { if (generation === this.generation) this.error = e.message }
      finally { this.polling = false; if (generation === this.generation && this.activeRun) this.timer = setTimeout(this.poll, 2000) }
    },
    async stop() { this.operating = true; this.syncRequired = true; this.workspace.lock(true); const generation = this.generation;
      try { const run = await agentRequest('/runs/' + this.run.runId + '/cancel', {});
        if (generation !== this.generation) return; this.run = run;
        if (!this.activeRun) await this.synchronize(generation); else this.schedule()
      } catch (e) { if (generation === this.generation) this.error = e.message } finally { this.operating = false } },
    async recover() {
      this.operating = true
      const generation = this.generation
      try {
        if (!this.activeRun) { await this.synchronize(generation); return }
        const id = this.run.runId
        const cancelled = await agentRequest('/runs/' + id + '/cancel', {})
        if (generation !== this.generation) return
        this.run = cancelled
        if (!this.activeRun) { await this.synchronize(generation); return }
        const state = await this.workspace.reload()
        if (generation !== this.generation) return
        const resolved = await agentRequest('/runs/' + id + '/resume', { reloadedServerDraft: true, expectedRevision: state.revision })
        if (generation !== this.generation) return
        this.run = resolved
        this.error = ''; if (this.activeRun) this.schedule(); else await this.synchronize(generation)
      } catch (e) { this.error = e.message } finally { this.operating = false }
    },
    async sendAnswer() {
      this.operating = true
      try { this.run = await agentRequest('/runs/' + this.run.runId + '/answers', { answer: this.answer }); this.answer = ''; this.error = ''; this.schedule() }
      catch (e) { this.error = e.message } finally { this.operating = false }
    }
  }
}
</script>

<style scoped>
.agent-task { padding: 12px; margin: 8px 0; border: 1px solid #c7d9ef; border-radius: 6px; font-size: 12px; line-height: 1.7; background: #f6f9fe; }
.agent-controls { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 10px; }
.agent-task ol, .agent-task ul { padding-left: 20px; }
</style>
