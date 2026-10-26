<template>
  <section v-if="enabled || recoveryAvailable || syncRequired" class="agent-task">
    <div class="agent-controls">
      <el-switch v-model="useAgent" active-text="计划执行" :disabled="busy || !(controlledWritesEnabled || snapshotWritesEnabled)" @change="$emit('mode', $event)" />
      <span v-if="controlledWritesEnabled">AI 写入仅限发送时的选区；读取展示无需写入授权。可用操作以服务端能力校验为准。</span>
      <el-checkbox v-if="controlledWritesEnabled && !busy" v-model="allowOverwrite">允许覆盖授权区域已有内容</el-checkbox>
    </div>
    <template v-if="run && run.runId">
      <strong>{{ run.goal }}</strong>
      <p v-if="run.kind === 'CONTROLLED_WRITE'">普通 AI 受控写入：{{ statusLabel }}</p>
      <p v-else>状态：{{ statusLabel }} · 计划 {{ run.planVersion || 0 }} · 模型调用 {{ run.modelCalls || 0 }} · {{ run.tokens || 0 }} tokens</p>
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
import { classifyAiActions, submitControlledCall, createOrResolveControlledRun, resolveControlledCreation } from '../controlled-write-executor'
const TERMINAL = ['COMPLETED', 'CANCELLED', 'FAILED']
export default {
  name: 'AgentTaskPanel',
  props: { workbookId: { type: [Number, String], default: null }, workspace: { type: Object, required: true } },
  data: () => ({ enabled: false, recoveryAvailable: false, snapshotWritesEnabled: false, controlledWritesEnabled: false, ordinaryWritesEnabled: false, controlledActions: [], syncRequired: false, generation: 0,
    useAgent: false, allowOverwrite: false, run: null, events: [], error: '', answer: '', details: false,
    timer: null, heartbeatTimer: null, polling: false, operating: false, responseEpoch: 0, responseSequence: 0, acceptedStateSequence: 0,
    clientId: createOperationId(), pendingRequest: null, pendingKind: null, creationUncertain: false }),
  computed: { activeRun() { return !!(this.run && this.run.runId && !TERMINAL.includes(this.run.status)) },
    busy() { return this.activeRun || this.syncRequired }, statusLabel() { return this.label(this.run.status) } },
  watch: {
    workbookId() { this.attach() },
    busy(value) { this.$emit('busy', value); if (value) this.workspace.lock(true) }
  },
  async mounted() {
    try { const c = await agentCapabilities(); this.enabled = c.enabled; this.recoveryAvailable = c.recoveryAvailable;
      this.snapshotWritesEnabled = c.snapshotWritesEnabled === true;
      this.controlledWritesEnabled = c.controlledWritesEnabled === true;
      this.ordinaryWritesEnabled = c.ordinaryAiControlledEnabled === true;
      this.controlledActions = Array.isArray(c.controlledActions) ? c.controlledActions.filter(type => typeof type === 'string') : [];
      if (this.workbookId) await this.attach() } catch (e) { this.error = '智能体服务暂不可用：' + e.message; if (this.workbookId) await this.attach() }
  },
  beforeDestroy() { this.generation++; this.beginStateChange() },
  methods: {
    beginStateChange() {
      clearTimeout(this.timer); clearInterval(this.heartbeatTimer)
      this.responseEpoch++
    },
    stateTicket(runId = null) {
      return { generation: this.generation, epoch: this.responseEpoch, workbookId: this.workbookId, runId, sequence: ++this.responseSequence }
    },
    currentTicket(ticket) {
      return ticket.generation === this.generation && ticket.epoch === this.responseEpoch && String(ticket.workbookId) === String(this.workbookId) &&
        (!ticket.runId || (this.run && this.run.runId === ticket.runId))
    },
    acceptRun(run, ticket, allowNew = false) {
      if (!this.currentTicket(ticket) || !run || !run.runId || (!allowNew && run.runId !== ticket.runId)) return false
      if (run.workbookId != null && String(run.workbookId) !== String(this.workbookId)) throw new Error('RESULT_UNKNOWN: 任务工作簿身份不符，保持保护')
      if (this.run && this.run.runId === run.runId) {
        const current = this.run.version, incoming = run.version
        if (Number.isSafeInteger(current) && !Number.isSafeInteger(incoming)) throw new Error('RESULT_UNKNOWN: 任务状态版本缺失，保持保护')
        if (Number.isSafeInteger(current) && Number.isSafeInteger(incoming)) {
          if (incoming < current) return false
          if (incoming === current && ticket.sequence < this.acceptedStateSequence) return false
          if (incoming === current && !this.activeRun && !this.syncRequired && !allowNew) return false
        } else if (ticket.sequence < this.acceptedStateSequence) return false
      }
      // Server version wins over request order. Sequence only breaks equal/unversioned ties.
      this.run = run; this.acceptedStateSequence = Math.max(this.acceptedStateSequence, ticket.sequence)
      return true
    },
    async refreshRun(ticket) {
      if (!this.currentTicket(ticket)) return false
      const read = this.stateTicket(ticket.runId)
      const fresh = await getAgent(ticket.runId)
      if (!this.currentTicket(read)) return false
      if (!fresh || fresh.runId !== ticket.runId || (Number.isSafeInteger(this.run.version) && (!Number.isSafeInteger(fresh.version) || fresh.version < this.run.version))) throw new Error('RESULT_UNKNOWN: 任务核实查询为空、身份不符或已过期，保持保护')
      this.acceptRun(fresh, read)
      return true
    },
    async reconcileStateError(ticket, error) {
      if (!this.currentTicket(ticket)) return
      this.error = error.message; this.syncRequired = true; this.workspace.lock(true)
      try {
        // A lost mutation response may have committed. Query facts; never repeat the mutation.
        if (!await this.refreshRun(ticket)) return
        if (this.activeRun) this.schedule(); else await this.synchronize(ticket.generation)
      } catch (queryError) { if (this.currentTicket(ticket)) this.error += '；状态核实未完成：' + queryError.message }
    },
    rememberCreation(kind) {
      const record = { kind, request: this.pendingRequest }
      sessionStorage.setItem('agentReport:create:' + this.pendingRequest.workbookId, JSON.stringify(record))
      this.pendingKind = kind; this.creationUncertain = true; this.syncRequired = true
    },
    clearCreation() {
      if (this.pendingRequest && typeof sessionStorage.removeItem === 'function') sessionStorage.removeItem('agentReport:create:' + this.pendingRequest.workbookId)
      this.creationUncertain = false; this.pendingKind = null
    },
    restoreCreation() {
      if (this.pendingRequest && String(this.pendingRequest.workbookId) !== String(this.workbookId)) { this.pendingRequest = null; this.pendingKind = null; this.creationUncertain = false }
      if (typeof sessionStorage.getItem !== 'function') return
      const stored = sessionStorage.getItem('agentReport:create:' + this.workbookId)
      if (stored) { const record = JSON.parse(stored); if (!record.request || String(record.request.workbookId) !== String(this.workbookId) || !['PLAN_AGENT', 'CONTROLLED_WRITE'].includes(record.kind)) throw new Error('创建核实记录无效，保持保护'); this.pendingRequest = record.request; this.pendingKind = record.kind; this.creationUncertain = true }
    },
    label(s) { return ({ QUEUED: '排队', PLANNING: '制定计划', EXECUTING: '执行中', REPLANNING: '修订计划', VERIFYING: '验证', WAITING_USER: '等待补充', PAUSED: '暂停', CANCELLING: '核实停止', CANCELLED: '已停止', COMPLETED: '完成', FAILED: '失败', PENDING: '待执行', RUNNING: '执行中', SUCCEEDED: '已验证', SUPERSEDED: '被新计划替代' })[s] || s },
    async attach() {
      this.beginStateChange()
      const generation = ++this.generation
      const ticket = this.stateTicket()
      this.run = null; this.events = []; this.error = ''
      if (!this.workbookId) return
      const id = this.workbookId
      this.syncRequired = true; this.workspace.lock(true)
      try {
        this.restoreCreation()
        if (!this.enabled && !this.recoveryAvailable && !this.creationUncertain) {
          this.syncRequired = false; this.workspace.lock(false); this.$emit('busy', false); return
        }
        const run = this.creationUncertain && this.pendingRequest
          ? await resolveControlledCreation(this.pendingKind, this.pendingRequest, agentRequest)
          : await activeAgent(id)
        if (!this.currentTicket(ticket) || String(id) !== String(this.workbookId)) return
        if (this.creationUncertain && this.pendingRequest && String(this.pendingRequest.workbookId) === String(id) && run.requestId !== this.pendingRequest.requestId) throw new Error('RESULT_UNKNOWN: 创建请求尚未核实，保持保护，请勿重新提交')
        if (this.creationUncertain) this.clearCreation()
        if (run.runId) this.acceptRun(run, ticket, true)
        else this.run = null
        if (this.activeRun) { this.useAgent = this.run.kind !== 'CONTROLLED_WRITE'; this.$emit('mode', this.useAgent); this.schedule() }
        else if (this.run) await this.synchronize(generation)
        else { this.syncRequired = false; this.workspace.lock(false); this.$emit('busy', false) }
      } catch (e) { if (this.currentTicket(ticket)) this.error = e.message }
    },
    async synchronize(generation = this.generation) {
      const ticket = this.stateTicket(this.run && this.run.runId)
      if (generation !== this.generation) throw new Error('页面已切换，保持保护状态')
      this.syncRequired = true; this.workspace.lock(true)
      if (this.creationUncertain) throw new Error('RESULT_UNKNOWN: 原创建请求尚未核实，不能解锁')
      const current = () => this.currentTicket(ticket)
      if (!(this.workspace.syncControlled && this.run && await this.workspace.syncControlled(this.run, current))) {
        if (!current()) throw new Error('任务核实操作已切换，保持保护状态')
        await this.workspace.reload(current)
      }
      if (!current()) throw new Error('任务或页面已切换，保持保护状态')
      this.syncRequired = false
      if (!this.activeRun) { this.workspace.lock(false); this.$emit('busy', false) }
    },
    async start(goal, conversationId) {
      if (this.creationUncertain) throw new Error('RESULT_UNKNOWN: 请先核实原创建请求')
      if (!this.enabled || !this.controlledWritesEnabled || this.busy || this.operating) throw new Error('智能体写入入口未开启，可核实或取消已有任务')
      this.beginStateChange()
      const ticket = this.stateTicket()
      const generation = this.generation
      this.operating = true; this.error = ''; this.workspace.lock(true)
      let attempted = false
      try {
        const state = await this.workspace.prepare()
        if (!this.currentTicket(ticket)) throw new Error('工作区已切换')
        this.pendingRequest = { requestId: createOperationId(), goal, conversationId, workbookId: this.workbookId,
          branchId: state.branchId, expectedRevision: state.revision, allowOverwrite: this.allowOverwrite,
          writeScopes: state.writeScopes, activeSheetId: state.activeSheetId, loadedRevision: state.loadedRevision,
          expectedDraftVersion: state.draftVersion, expectedBranchVersion: state.branchVersion }
        this.rememberCreation('PLAN_AGENT'); attempted = true
        const created = await createOrResolveControlledRun('PLAN_AGENT', this.pendingRequest, agentRequest)
        if (!this.acceptRun(created, ticket, true)) throw new Error('工作区已切换，原创建结果需核实')
        this.clearCreation(); this.events = []; this.schedule()
      } catch (e) {
        if (this.currentTicket(ticket)) {
          this.error = e.message
          if (attempted && (e.creationUncertain || !e.definitive)) { this.creationUncertain = true; this.syncRequired = true; this.run = null }
          else if (attempted) { this.clearCreation(); this.operating = false; await this.attach() }
          else this.workspace.lock(false)
        }
        throw e
      } finally { if (this.currentTicket(ticket)) this.operating = false }
    },
    captureOrdinaryAuthorization() {
      return { workspace: this.workspace.captureAuthorization(), allowOverwrite: this.allowOverwrite }
    },
    async startActions(actions, conversationId, authorization, requestId, goal) {
      if (this.creationUncertain) throw new Error('RESULT_UNKNOWN: 请先核实原创建请求')
      if (classifyAiActions(actions, this.controlledActions) !== 'controlled-values' || !this.controlledWritesEnabled || !this.ordinaryWritesEnabled || this.busy || this.operating) throw new Error('CONTROLLED_DISABLED: 当前不能开始普通 AI 受控写入')
      this.beginStateChange()
      const ticket = this.stateTicket()
      const generation = this.generation
      this.operating = true; this.error = ''; this.workspace.lock(true)
      let attempted = false
      try {
        if (!authorization || !authorization.workspace) throw new Error('缺少发送时的写入授权')
        const state = await this.workspace.prepare(authorization.workspace)
        if (!this.currentTicket(ticket)) throw new Error('工作区已切换')
        this.pendingRequest = { entryPoint: 'ordinary-ai', requestId, goal, workbookId: this.workbookId, branchId: state.branchId,
          loadedRevision: state.loadedRevision, expectedDraftVersion: state.draftVersion, expectedBranchVersion: state.branchVersion,
          activeSheetId: state.activeSheetId, writeScopes: state.writeScopes, overwrite: authorization.allowOverwrite ? 'confirmed' : 'empty-only', actions }
        if (conversationId) this.pendingRequest.conversationId = conversationId
        this.rememberCreation('CONTROLLED_WRITE'); attempted = true
        const created = await createOrResolveControlledRun('CONTROLLED_WRITE', this.pendingRequest, agentRequest)
        if (!this.acceptRun(created, ticket, true)) throw new Error('工作区已切换，创建结果需核实')
        this.clearCreation()
        this.syncRequired = true; this.events = []; this.schedule()
        return this.run
      } catch (e) {
        if (this.currentTicket(ticket)) {
          this.error = e.message
          if (attempted && (e.creationUncertain || !e.definitive)) { this.creationUncertain = true; this.syncRequired = true; this.run = null; this.error += '；创建结果需核实，保持保护，请勿重新提交'; }
          else if (attempted) { this.clearCreation(); this.operating = false; this.syncRequired = true; await this.attach() }
          else this.workspace.lock(false)
        }
        throw e
      } finally { if (this.currentTicket(ticket)) this.operating = false }
    },
    publishOrdinary(error) {
      if (this.run && this.run.entryPoint === 'ordinary-ai') this.$emit('ordinary-result', { run: this.run, error: error && error.message })
    },
    schedule() {
      clearTimeout(this.timer); clearInterval(this.heartbeatTimer)
      this.heartbeatTimer = setInterval(() => {
        if (this.activeRun && !['PAUSED', 'CANCELLING'].includes(this.run.status)) {
          const ticket = this.stateTicket(this.run.runId)
          agentRequest('/runs/' + this.run.runId + '/lease-heartbeat', { fencingToken: this.run.fencingToken }).catch(e => { if (this.currentTicket(ticket)) this.error = e.message })
        }
      }, 20000)
      this.timer = setTimeout(this.poll, 500)
    },
    async poll() {
      if (this.operating || !this.run || !this.run.runId) return
      if (this.polling) {
        // An old context may still own the in-flight HTTP request. Keep the current task observable.
        if (this.activeRun) { clearTimeout(this.timer); this.timer = setTimeout(this.poll, 500) }
        return
      }
      this.polling = true
      const id = this.run.runId
      const generation = this.generation
      const ticket = this.stateTicket(id)
      try {
        const latest = await getAgent(id)
        if (!this.acceptRun(latest, ticket)) return
        this.syncRequired = true
        const after = this.events.length ? this.events[this.events.length - 1].sequence : 0
        const events = await agentRequest('/runs/' + id + '/events?after=' + after, null, 'get')
        if (!this.currentTicket(ticket)) return
        this.events = this.events.concat(events).slice(-200)
        if (this.run.status === 'EXECUTING' && this.run.call && this.run.call.transportVersion === 2 && this.run.call.status === 'PENDING') {
          const call = await agentRequest('/runs/' + id + '/controlled-claims', { clientId: this.clientId, fencingToken: this.run.fencingToken })
          if (!this.currentTicket(ticket)) throw new Error('页面或核实操作已切换，不执行旧调用')
          if (call.claimToken && call.actions) {
            const acknowledged = await submitControlledCall(this.workspace, call, id, agentRequest,
              () => this.currentTicket(ticket))
            if (!this.acceptRun(acknowledged, ticket)) return
            if (this.run.lastControlledResult && this.run.lastControlledResult.callId === call.callId && this.run.lastControlledResult.persistenceStatus === 'SAVED') await this.workspace.acceptControlled(this.run.lastControlledResult)
          }
        }
        if (this.snapshotWritesEnabled && this.run.status === 'EXECUTING' && this.run.call && this.run.call.status === 'PENDING') {
          const call = await agentRequest('/runs/' + id + '/tool-claims', { clientId: this.clientId, fencingToken: this.run.fencingToken })
          if (call.callId) {
            let receipt = { callId: call.callId, claimToken: call.claimToken, argsHash: call.argsHash, fencingToken: call.fencingToken, ok: false }
            try { receipt.snapshot = await this.workspace.apply(call.snapshot); receipt.ok = true; receipt.calculationVerified = true } catch (e) { this.error = e.message }
            // Only retry the identical receipt, never the browser operation.
            let result
            try { result = await agentRequest('/runs/' + id + '/tool-results', receipt) }
            catch (e) { if (!this.currentTicket(ticket)) return; result = await agentRequest('/runs/' + id + '/tool-results', receipt) }
            if (!this.acceptRun(result, ticket)) return
            if (receipt.ok) await this.workspace.saved()
          }
        }
        if (!this.activeRun) { clearInterval(this.heartbeatTimer); await this.synchronize(generation) }
        this.publishOrdinary()
      } catch (e) { if (this.currentTicket(ticket)) { if (e.run) this.acceptRun(e.run, ticket); this.error = e.message; this.publishOrdinary(e) } }
      finally { this.polling = false; if (this.currentTicket(ticket) && this.activeRun && !this.operating) this.timer = setTimeout(this.poll, 2000) }
    },
    async stop() { if (this.operating) return; this.beginStateChange(); const ticket = this.stateTicket(this.run.runId); this.operating = true; this.syncRequired = true; this.workspace.lock(true); const generation = this.generation;
      try { const run = await agentRequest('/runs/' + this.run.runId + '/cancel', {});
        if (!this.currentTicket(ticket)) return; this.acceptRun(run, ticket);
        if (!await this.refreshRun(ticket)) return;
        if (!this.activeRun) await this.synchronize(generation); else this.schedule()
      } catch (e) { await this.reconcileStateError(ticket, e) } finally { if (this.currentTicket(ticket)) this.operating = false } },
    async recover() {
      if (this.creationUncertain) { await this.attach(); return }
      if (this.operating) return
      this.beginStateChange()
      const ticket = this.stateTicket(this.run.runId)
      this.operating = true
      const generation = this.generation
      try {
        if (!this.activeRun) { await this.synchronize(generation); return }
        const id = this.run.runId
        const cancelled = await agentRequest('/runs/' + id + '/cancel', {})
        if (!this.currentTicket(ticket)) return
        this.acceptRun(cancelled, ticket)
        if (!await this.refreshRun(ticket)) return
        if (!this.activeRun) { await this.synchronize(generation); return }
        const state = await this.workspace.reload(() => this.currentTicket(ticket))
        if (!this.currentTicket(ticket)) return
        const resolved = await agentRequest('/runs/' + id + '/resume', { reloadedServerDraft: true, expectedRevision: state.revision })
        if (!this.currentTicket(ticket)) return
        this.acceptRun(resolved, ticket)
        if (!await this.refreshRun(ticket)) return
        this.error = ''; if (this.activeRun) this.schedule(); else await this.synchronize(generation)
      } catch (e) { await this.reconcileStateError(ticket, e) } finally { if (this.currentTicket(ticket)) this.operating = false }
    },
    async sendAnswer() {
      if (this.operating) return
      this.beginStateChange()
      const ticket = this.stateTicket(this.run.runId)
      this.operating = true
      try { const run = await agentRequest('/runs/' + ticket.runId + '/answers', { answer: this.answer }); if (!this.currentTicket(ticket)) return; this.acceptRun(run, ticket); if (!await this.refreshRun(ticket)) return; this.answer = ''; this.error = ''; if (this.activeRun) this.schedule(); else await this.synchronize(ticket.generation) }
      catch (e) { await this.reconcileStateError(ticket, e) } finally { if (this.currentTicket(ticket)) this.operating = false }
    }
  }
}
</script>

<style scoped>
.agent-task { padding: 12px; margin: 8px 0; border: 1px solid #c7d9ef; border-radius: 6px; font-size: 12px; line-height: 1.7; background: #f6f9fe; }
.agent-controls { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 10px; }
.agent-task ol, .agent-task ul { padding-left: 20px; }
</style>
