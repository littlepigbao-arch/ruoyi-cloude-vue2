<template>
  <div class="app-container">
    <el-alert v-if="conflictDrafts.length" title="已保留同步前的本地编辑草稿，可下载后进行人工合并" type="warning" :closable="false" />
    <el-button v-if="conflictDrafts.length" size="mini" @click="downloadConflictDrafts">下载保留的编辑草稿</el-button>
    <el-button v-if="branchCheckoutUncertain && !versionDialogVisible" :disabled="persistBusy" size="mini" type="warning" @click="reconcileBranchCheckout">核实分支切换</el-button>
    <el-row :gutter="10" class="mb8">
      <el-col :span="1.5">
        <el-upload
          class="upload-demo"
          action="#"
          :auto-upload="false"
          :show-file-list="false"
          :before-upload="beforeUpload"
          :on-change="handleFileChange"
          accept=".xlsx,.xls"
        >
          <el-button
            type="primary"
            plain
            icon="el-icon-upload2"
            size="mini"
          >导入Excel</el-button>
        </el-upload>
      </el-col>
      <el-col :span="1.5">
        <el-button
          type="primary"
          plain
          icon="el-icon-time"
          size="mini"
          @click="openVersionDialog"
          :disabled="!currentDocumentId"
        >版本管理</el-button>
      </el-col>
      <el-col :span="1.5">
        <el-button
          type="success"
          plain
          icon="el-icon-download"
          size="mini"
          @click="handleExport"
          :disabled="!hasData"
        >导出Excel数据</el-button>
        <el-button plain size="mini" :disabled="!hasData" @click="handleSnapshotExport">完整备份JSON</el-button>
      </el-col>
      <el-col :span="1.5">
        <el-button
          type="warning"
          plain
          icon="el-icon-document"
          size="mini"
          @click="createNewSheet"
        >新建表格</el-button>
      </el-col>
      <el-col :span="1.5">
        <el-button
          type="primary"
          plain
          icon="el-icon-check"
          size="mini"
          @click="saveCurrentWorkbook(true)"
          :disabled="!hasData"
        >保存</el-button>
      </el-col>
      <el-col :span="1.5">
        <el-button
          type="info"
          plain
          icon="el-icon-delete"
          size="mini"
          @click="clearSheet"
          :disabled="!hasData"
        >清空数据</el-button>
      </el-col>
      <el-col v-if="largeDataMode" :span="1.5">
        <el-tag type="warning" size="small" effect="dark">
          大数据模式 · {{ formatLargeDataStats() }}
        </el-tag>
      </el-col>
      <right-toolbar :showSearch.sync="showSearch" @queryTable="getList"></right-toolbar>
    </el-row>

    <el-alert
      title="提示信息"
      :description="statusMessage"
      type="info"
      :closable="false"
      show-icon
      class="mb8"
    />

    <transition name="persist-fade">
      <div v-if="persistBusy" class="persist-mask" @click.stop @mousedown.stop @keydown.stop>
        <div class="persist-card">
          <i class="el-icon-loading persist-card__icon"></i>
          <div class="persist-card__title">{{ persistTitle }}</div>
          <el-progress
            :percentage="persistProgress"
            :stroke-width="12"
            :show-text="true"
            status="success"
          />
          <div class="persist-card__hint">正在处理大数据，请勿关闭页面或刷新浏览器</div>
        </div>
      </div>
    </transition>

    <div class="content-row">
      <document-list
        :documents="documents"
        :active-id="currentDocumentId"
        @select="loadDocument"
        @delete="deleteDocument"
        @rename="openRenameDialog"
        @refresh="fetchDocuments"
      />
      <div class="univer-container">
        <div ref="univerContainer" id="univer-container" class="univer-wrapper"></div>
      </div>
    </div>

    <!-- AI 助手悬浮面板 -->
    <ai-chat-panel
      :univerAPI="univerAPI"
      :workbook="activeWorkbook"
      :workbook-id="currentDocumentId"
      :agent-workspace="agentWorkspaceAdapter"
      :agent-locked="agentLocked"
      :read-only="historyPreview || persistBusy"
      :visible.sync="aiPanelVisible"
      :use-mock.sync="aiUseMock"
      :allow-mock-fallback="true"
      :large-data-mode="largeDataMode"
      :workbook-stats="workbookStats"
      @executed="onAiExecuted"
      @rollback="rollbackAiOperation"
      @execution-state="aiExecutionInProgress = $event"
    />

    <!-- 重命名文档对话框 -->
    <el-dialog
      title="重命名文档"
      :visible.sync="renameDialogVisible"
      width="420px"
      append-to-body
    >
      <el-form label-width="80px">
        <el-form-item label="文档名称">
          <el-input
            v-model="renameName"
            placeholder="请输入新名称"
            maxlength="100"
            show-word-limit
            @keyup.enter.native="confirmRename"
          />
        </el-form-item>
      </el-form>
      <div slot="footer">
        <el-button size="mini" @click="renameDialogVisible = false">取消</el-button>
        <el-button type="primary" size="mini" :loading="renameLoading" @click="confirmRename">确定</el-button>
      </div>
    </el-dialog>

    <el-dialog title="版本管理" :visible.sync="versionDialogVisible" width="760px" append-to-body>
      <el-alert
        :title="'当前分支：' + (currentBranchName || 'main') + ' · HEAD：' + (headCommitId || '无')"
        :description="'未提交变化：' + pendingChanges.length + ' 项'"
        type="info" :closable="false" class="mb8"
      />
      <el-form inline size="mini">
        <el-form-item v-if="branchCheckoutUncertain">
          <el-button type="warning" :disabled="persistBusy" @click="reconcileBranchCheckout">核实分支切换</el-button>
        </el-form-item>
        <el-form-item label="切换分支">
          <el-select v-model="selectedBranchId" style="width:190px" :disabled="agentLocked || agentSyncRequired || persistBusy || historyPreview" @change="checkoutSelectedBranch">
            <el-option v-for="item in branches" :key="item.branchId" :label="item.branchName" :value="item.branchId" />
          </el-select>
        </el-form-item>
        <el-form-item label="新建草稿">
          <el-input v-model="newBranchName" style="width:180px" placeholder="例如：预算调整" />
        </el-form-item>
        <el-form-item>
          <el-button @click="createDraftBranch">创建</el-button>
        </el-form-item>
        <el-form-item v-if="currentBranchName !== 'main'">
          <el-button type="success" plain @click="fastForwardCurrentBranch">安全纳入 main</el-button>
        </el-form-item>
      </el-form>
      <el-form inline size="mini">
        <el-form-item label="提交说明">
          <el-input v-model="commitMessage" style="width:360px" maxlength="500" placeholder="说明本次修改的业务目的" />
        </el-form-item>
        <el-form-item>
          <el-button type="primary" :disabled="!pendingChanges.length" :loading="commitLoading" @click="commitPendingChanges">提交</el-button>
          <el-button type="danger" plain :disabled="agentLocked || persistBusy || historyPreview || draftVersion < 0" @click="discardCurrentDraft">丢弃已保存草稿</el-button>
        </el-form-item>
      </el-form>
      <el-table :data="commitHistory" size="mini" max-height="420">
        <el-table-column prop="commitNo" label="#" width="60" />
        <el-table-column prop="commitMessage" label="说明" min-width="220" />
        <el-table-column prop="source" label="来源" width="90" />
        <el-table-column prop="authorName" label="作者" width="100" />
        <el-table-column prop="createdTime" label="时间" width="170" />
        <el-table-column label="操作" width="150">
          <template slot-scope="scope">
            <el-button type="text" size="mini" @click="previewVersion(scope.row)">预览</el-button>
            <el-button type="text" size="mini" @click="restoreToVersion(scope.row)">恢复</el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-dialog>
  </div>
</template>

<script>
import {
  LocaleType,
  mergeLocales,
  createUniver,
  UniverInstanceType,
  defaultTheme
} from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting'
import sheetsConditionalFormattingZhCN from '@univerjs/preset-sheets-conditional-formatting/locales/zh-CN'
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter'
import sheetsFilterZhCN from '@univerjs/preset-sheets-filter/locales/zh-CN'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import sheetsDrawingZhCN from '@univerjs/preset-sheets-drawing/locales/zh-CN'

import '@univerjs/presets/lib/styles/preset-sheets-core.css'
import '@univerjs/presets/lib/styles/preset-sheets-conditional-formatting.css'
import '@univerjs/presets/lib/styles/preset-sheets-filter.css'
import '@univerjs/presets/lib/styles/preset-sheets-drawing.css'

// 注册 sheets facade mixin（FUniverSheetsMixin），让 univerAPI.getActiveWorkbook() 可用
// 关键：必须用 lib/facade.js，与 @univerjs/preset-sheets-core 内部 import 的
// "@univerjs/sheets/lib/facade" 解析到同一个文件。若误用 lib/es/facade.js，
// 会多加载一个 FRange/FWorksheet 类，导致 FRange.extend(FRangeSheetsFilterMixin)
// 扩展错类，运行时 fRange.getFilter()/createFilter() 恒为 undefined。
import '@univerjs/sheets/lib/facade.js'
// 注册 filter facade mixin（FRangeSheetsFilterMixin），让 FRange.getFilter()/createFilter() 可用
import '@univerjs/sheets-filter/lib/facade.js'

import LuckyExcel from 'luckyexcel'
import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'
import AiChatPanel from './components/AiChatPanel.vue'
import DocumentList from './components/DocumentList.vue'
import { agentWorkspace, activeAgent } from '@/api/ai/agent'
import { executeControlledWrite, sameControlledSnapshot, sameControlledFormatSnapshot, sameControlledFormatTransition, sameControlledFilterSnapshot, sameControlledFilterTransition, sameControlledFormulaSnapshot, sameControlledFormulaTransition, sameControlledMergeSnapshot, sameControlledMergeTransition, sameControlledStructureSnapshot, sameControlledStructureTransition } from './controlled-write-executor'
import {
  listWorkbooks, getWorkbook, saveWorkbook, deleteWorkbook, renameWorkbook,
  getVersionStatus, saveWorkingState, discardWorkingState, createCommit, listCommits, getVersionSnapshot, restoreVersion,
  listBranches, createBranch, checkoutBranch, fastForwardBranchToMain
} from '@/api/ai/workbook'
import { createOperationId, buildManualChanges, buildLightweightManualChanges } from './change-capture'
import { workingStatePayload, sameWorkspace, versionMutationBinding } from './working-state-sync'
import { exportWorkbook, exportSheet } from './excel-export'

const LARGE_DATA_CELL_THRESHOLD = 200000
const LARGE_DATA_ROW_THRESHOLD = 20000
// 大数据模式：最后一次编辑后 10 分钟自动保存
const LARGE_DATA_AUTOSAVE_DELAY = 1000 * 60 * 1000
// 纯导航命令不会改变工作簿数据，不能进入变更捕获和自动保存队列。
const NON_MUTATING_SHEET_COMMANDS = new Set([
  'sheet.command.move-selection',
  'sheet.command.move-selection-enter-tab',
  'sheet.command.expand-selection',
  'sheet.command.set-scroll-relative',
  'sheet.command.scroll-view',
  'sheet.command.scroll-to-cell',
  'sheet.command.scroll-view-reset',
  'sheet.command.set-worksheet-activate'
])

export default {
  name: 'ApplicationExcel',
  components: { AiChatPanel, DocumentList },
  data() {
    return {
      showSearch: false,
      statusMessage: '请导入Excel文件或点击"新建表格"开始使用',
      univer: null,
      univerAPI: null,
      hasData: false,
      currentWorkbookId: null,
      // createUnit() 返回的 workbook 引用（FWorkbook），传给 AI 执行器操控
      activeWorkbook: null,
      // AI 助手面板状态
      aiPanelVisible: true,
      aiUseMock: false,
      // 文档列表状态
      documents: [],
      currentDocumentId: null,
      currentDocumentName: '',
      currentDocumentType: '',
      currentDocumentVersion: null,
      currentBranchId: null,
      currentBranchName: 'main',
      selectedBranchId: null,
      branches: [],
      newBranchName: '',
      branchVersion: null,
      draftVersion: -1,
      loadedRevision: null,
      loadedSnapshotFingerprint: null,
      workspaceGeneration: 0,
      branchCheckoutUncertain: false,
      excelImportGeneration: 0,
      pendingExcelImport: null,
      agentSyncRequired: false,
      conflictDrafts: [],
      agentLocked: false,
      agentApplying: false,
      controlledExecution: null,
      headCommitId: null,
      pendingChanges: [],
      pendingActions: [],
      versionDialogVisible: false,
      commitHistory: [],
      commitMessage: '',
      commitLoading: false,
      aiExecutionInProgress: false,
      changeCaptureDisposables: [],
      manualCommands: [],
      manualBeforeSnapshot: null,
      manualCaptureTimer: null,
      autosaveTimer: null,
      persistBusy: false,
      persistProgress: 0,
      persistTitle: '正在保存表格…',
      persistProgressTimer: null,
      suspendChangeCapture: false,
      historyPreview: false,
      largeDataMode: false,
      workbookStats: {
        sheetCount: 0,
        populatedCells: 0,
        maxUsedRows: 0,
        maxUsedColumns: 0,
        sheets: {}
      },
      // 重命名对话框状态
      renameDialogVisible: false,
      renameLoading: false,
      renameTarget: null,
      renameName: ''
    }
  },
  computed: {
    agentWorkspaceAdapter() {
      return { lock: this.lockAgentWorkspace, prepare: this.prepareAgentWorkspace, apply: this.applyAgentSnapshot,
        saved: this.refreshAgentDraft, reload: this.reloadAgentDraft, applyControlled: this.applyControlledWrite,
        acceptControlled: this.acceptControlledResult, syncControlled: this.syncControlledResult, captureAuthorization: this.captureAgentAuthorization, captureReadSelection: this.captureReadSelection }
    }
  },
  mounted() {
    try {
      for (let i = 0; i < sessionStorage.length; i++) {
        const key = sessionStorage.key(i)
        if (key && key.startsWith('agentReport:conflict:')) this.conflictDrafts.push(JSON.parse(sessionStorage.getItem(key)))
      }
    } catch (e) { this.statusMessage = '保留草稿读取失败，请勿关闭已有编辑页面' }
    this.$nextTick(() => {
      this.initUniver()
    })
    this.fetchDocuments()
  },
  beforeDestroy() {
    this.cancelExcelImport()
    clearTimeout(this.autosaveTimer)
    this.autosaveTimer = null
    clearInterval(this.persistProgressTimer)
    this.persistProgressTimer = null
    this.disposeUniver()
  },
  methods: {
    lockAgentWorkspace(value) {
      if (!value && this.agentSyncRequired) return
      if (value) this.cancelExcelImport()
      this.agentLocked = value
      if (value) { clearTimeout(this.autosaveTimer); this.autosaveTimer = null }
    },
    captureReadSelection() {
      const sheet = this.univerAPI.getActiveWorkbook().getActiveSheet()
      const range = sheet.getActiveRange(), cell = sheet.getActiveCell()
      if (!range) return null
      const raw = cell && (typeof cell.getRange === 'function' ? cell.getRange() : cell._range)
      return { sheetId: sheet.getSheetId(), range: range.getA1Notation(),
        activeCell: raw ? { row: raw.startRow, column: raw.startColumn } : null }
    },
    captureAgentAuthorization() {
      const sheet = this.univerAPI.getActiveWorkbook().getActiveSheet(), selection = sheet.getActiveRange()
      return { identity: this.workspaceIdentity(), loadedRevision: this.loadedRevision, draftVersion: this.draftVersion,
        branchVersion: this.branchVersion, fingerprint: JSON.stringify(this.getSnapshot()), activeSheetId: sheet.getSheetId(),
        writeScopes: selection && typeof selection.getA1Notation === 'function' ? [{ sheetId: sheet.getSheetId(), range: selection.getA1Notation() }] : [] }
    },
    async prepareAgentWorkspace(authorization) {
      if (!this.currentDocumentId || !this.currentBranchId || this.historyPreview || this.persistBusy) throw new Error('请先保存并打开可编辑的工作簿')
      if (authorization && (!sameWorkspace(this, authorization.identity) || authorization.loadedRevision !== this.loadedRevision || authorization.draftVersion !== this.draftVersion || authorization.branchVersion !== this.branchVersion || authorization.fingerprint !== JSON.stringify(this.getSnapshot()))) throw new Error('WRITE_VERSION_CONFLICT: 等待 AI 回复期间工作区已变化，请保留编辑并重新发起请求')
      this.flushManualChanges(true)
      const snapshot = this.getSnapshot()
      if (!snapshot) throw new Error('无法获取当前工作簿')
      const identity = this.workspaceIdentity()
      const response = await saveWorkingState(this.currentDocumentId, workingStatePayload(this, snapshot))
      if (!response || response.code !== 200) throw new Error((response && response.msg) || '运行前草稿保存失败')
      if (!sameWorkspace(this, identity)) throw new Error('工作簿已切换，忽略旧保存响应')
      this.adoptSavedDraft(response.data, snapshot)
      const state = await agentWorkspace(this.currentDocumentId)
      if (!sameWorkspace(this, identity)) throw new Error('工作簿已切换，忽略旧工作区响应')
      const captured = authorization || this.captureAgentAuthorization()
      if (!captured.writeScopes.length) throw new Error('请先选中允许写入的目标区域')
      state.activeSheetId = captured.activeSheetId
      state.writeScopes = captured.writeScopes
      return state
    },
    async applyControlledWrite(manifest) {
      const identity = this.workspaceIdentity()
      if (!this.agentLocked || String(manifest.workbookId) !== String(identity.workbookId) || String(manifest.branchId) !== String(identity.branchId) || manifest.loadedRevision !== this.loadedRevision) throw new Error('受控目标或版本已变化，保持保护')
      if (!this.loadedSnapshotFingerprint || JSON.stringify(this.getSnapshot()) !== this.loadedSnapshotFingerprint) throw new Error('WRITE_VERSION_CONFLICT: 受控执行前本地工作簿已变化，请核实服务器草稿')
      if (this.controlledExecution && this.controlledExecution.callId === manifest.callId) throw new Error('同一调用不得重复执行，请核实结果')
      this.agentSyncRequired = true
      this.controlledExecution = { callId: manifest.callId, identity, fingerprint: null }
      this.agentApplying = true; this.aiExecutionInProgress = true
      try {
        const result = await executeControlledWrite(this.univerAPI, this.univerAPI.getActiveWorkbook(), manifest)
        if (!sameWorkspace(this, identity)) throw new Error('页面已切换，不能提交旧执行结果')
        this.controlledExecution.fingerprint = JSON.stringify(this.getSnapshot())
        return result
      } finally { this.agentApplying = false; this.aiExecutionInProgress = false }
    },
    async acceptControlledResult(result) {
      const executed = this.controlledExecution
      if (!result || result.validation !== 'PASSED' || result.persistenceStatus !== 'SAVED' || !executed || executed.callId !== result.callId || !sameWorkspace(this, executed.identity) || !executed.fingerprint || executed.fingerprint !== JSON.stringify(this.getSnapshot())) throw new Error('执行结果或页面已变化，必须核实服务器草稿')
      this.draftVersion = result.draftVersion; this.loadedRevision = result.loadedRevision
      this.loadedSnapshotFingerprint = executed.fingerprint
      if (result.executedChange && !this.pendingChanges.some(c => c.operationId === result.callId)) {
        this.pendingChanges.push(result.executedChange)
        this.pendingActions.push(...(result.actions || []))
      }
    },
    async syncControlledResult(run, isCurrent = () => true) {
      const identity = this.workspaceIdentity(), result = run.lastControlledResult
      if (!result) return false
      const state = await agentWorkspace(identity.workbookId)
      if (!isCurrent() || !sameWorkspace(this, identity) || String(state.branchId) !== String(identity.branchId)) throw new Error('页面或任务核实操作已切换，保持保护')
      if (state.loadedRevision !== result.loadedRevision || state.draftVersion !== result.draftVersion) return false
      const manifest = run.lastControlledManifest
      if (manifest && manifest.structurePatch) {
        if (manifest.callId !== result.callId || result.validation !== 'PASSED' || result.persistenceStatus !== 'SAVED') throw new Error('RESULT_UNKNOWN: 结构调用缺少可信保存结果')
        const patch = manifest.structurePatch
        if (this.loadedRevision === result.loadedRevision && this.draftVersion === result.draftVersion) {
          if (JSON.stringify(this.getSnapshot()) !== this.loadedSnapshotFingerprint || !sameControlledStructureSnapshot(this.getSnapshot(), state.snapshot, patch)) {
            this.agentSyncRequired = true; this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 结构保存后页面已变化，保持保护')
          }
          if (patch.type === 'deleteSheet' && this.getSnapshot().activeSheet === patch.sheetId) {
            const restored = await this.reloadAgentDraft(isCurrent, { skipPreserve: true })
            if (!isCurrent() || restored.loadedRevision !== result.loadedRevision || this.getSnapshot().activeSheet !== patch.fallbackSheetId || !sameControlledStructureSnapshot(this.getSnapshot(), restored.snapshot, patch)) throw new Error('RESULT_UNKNOWN: 删除后活动 Sheet 恢复失败，保持保护')
          }
          this.agentSyncRequired = false
          return true
        }
        const executed = this.controlledExecution && this.controlledExecution.callId === result.callId ? this.controlledExecution : null
        if (executed) {
          if (!sameWorkspace(this, executed.identity) || !executed.fingerprint || executed.fingerprint !== JSON.stringify(this.getSnapshot()) || !sameControlledStructureSnapshot(this.getSnapshot(), state.snapshot, patch)) {
            this.agentSyncRequired = true; this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 结构执行与服务器草稿不一致，保持保护')
          }
        } else {
          if (manifest.loadedRevision !== this.loadedRevision || JSON.stringify(this.getSnapshot()) !== this.loadedSnapshotFingerprint || !sameControlledStructureTransition(this.getSnapshot(), state.snapshot, patch)) {
            this.agentSyncRequired = true; this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 结构恢复含目标外变化，保持保护')
          }
        }
        const reloaded = await this.reloadAgentDraft(isCurrent, { skipPreserve: !!executed })
        if (!isCurrent() || reloaded.loadedRevision !== result.loadedRevision || reloaded.draftVersion !== result.draftVersion || !sameControlledStructureSnapshot(this.getSnapshot(), reloaded.snapshot, patch)) {
          this.agentSyncRequired = true; this.lockAgentWorkspace(true)
          throw new Error('RESULT_UNKNOWN: 结构重载与可信草稿不一致，保持保护')
        }
        return true
      }
      if (manifest && manifest.mergePatch) {
        if (manifest.callId !== result.callId || result.validation !== 'PASSED' || result.persistenceStatus !== 'SAVED') throw new Error('RESULT_UNKNOWN: 合并调用缺少可信保存结果')
        const patch = manifest.mergePatch
        if (this.loadedRevision === result.loadedRevision && this.draftVersion === result.draftVersion) {
          if (JSON.stringify(this.getSnapshot()) !== this.loadedSnapshotFingerprint || !sameControlledMergeSnapshot(this.getSnapshot(), state.snapshot, patch)) {
            this.agentSyncRequired = true; this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 合并保存后页面已变化，保持保护')
          }
          this.agentSyncRequired = false
          return true
        }
        const executed = this.controlledExecution && this.controlledExecution.callId === result.callId ? this.controlledExecution : null
        if (executed) {
          if (!sameWorkspace(this, executed.identity) || !executed.fingerprint || executed.fingerprint !== JSON.stringify(this.getSnapshot()) || !sameControlledMergeSnapshot(this.getSnapshot(), state.snapshot, patch)) {
            this.agentSyncRequired = true; this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 合并执行与服务器草稿不一致，保持保护')
          }
        } else {
          if (manifest.loadedRevision !== this.loadedRevision || JSON.stringify(this.getSnapshot()) !== this.loadedSnapshotFingerprint || !sameControlledMergeTransition(this.getSnapshot(), state.snapshot, patch)) {
            this.agentSyncRequired = true; this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 合并恢复含目标外变化，保持保护')
          }
        }
        const reloaded = await this.reloadAgentDraft(isCurrent, { skipPreserve: !!executed })
        if (!isCurrent() || reloaded.loadedRevision !== result.loadedRevision || reloaded.draftVersion !== result.draftVersion || !sameControlledMergeSnapshot(this.getSnapshot(), reloaded.snapshot, patch)) {
          this.agentSyncRequired = true; this.lockAgentWorkspace(true)
          throw new Error('RESULT_UNKNOWN: 合并重载与可信草稿不一致，保持保护')
        }
        return true
      }
      if (manifest && manifest.formulaPatch) {
        if (manifest.callId !== result.callId || result.validation !== 'PASSED' || result.persistenceStatus !== 'SAVED') throw new Error('RESULT_UNKNOWN: 公式调用缺少可信保存结果')
        const patch = manifest.formulaPatch
        if (this.loadedRevision === result.loadedRevision && this.draftVersion === result.draftVersion) {
          if (JSON.stringify(this.getSnapshot()) !== this.loadedSnapshotFingerprint || !sameControlledFormulaSnapshot(this.getSnapshot(), state.snapshot, patch)) {
            this.agentSyncRequired = true; this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 公式保存后页面已变化，保持保护')
          }
          this.agentSyncRequired = false
          return true
        }
        const executed = this.controlledExecution && this.controlledExecution.callId === result.callId ? this.controlledExecution : null
        if (executed) {
          if (!sameWorkspace(this, executed.identity) || !executed.fingerprint || executed.fingerprint !== JSON.stringify(this.getSnapshot()) || !sameControlledFormulaSnapshot(this.getSnapshot(), state.snapshot, patch)) {
            this.agentSyncRequired = true; this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 公式执行与服务器草稿不一致，保持保护')
          }
        } else {
          if (manifest.loadedRevision !== this.loadedRevision || JSON.stringify(this.getSnapshot()) !== this.loadedSnapshotFingerprint || !sameControlledFormulaTransition(this.getSnapshot(), state.snapshot, patch)) {
            this.agentSyncRequired = true; this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 公式恢复含目标外变化，保持保护')
          }
        }
        const reloaded = await this.reloadAgentDraft(isCurrent, { skipPreserve: !!executed })
        if (!isCurrent() || reloaded.loadedRevision !== result.loadedRevision || reloaded.draftVersion !== result.draftVersion || !sameControlledFormulaSnapshot(this.getSnapshot(), reloaded.snapshot, patch)) {
          this.agentSyncRequired = true; this.lockAgentWorkspace(true)
          throw new Error('RESULT_UNKNOWN: 公式重载与可信草稿不一致，保持保护')
        }
        return true
      }
      if (manifest && manifest.filterPatch) {
        if (manifest.callId !== result.callId || result.validation !== 'PASSED' || result.persistenceStatus !== 'SAVED') throw new Error('RESULT_UNKNOWN: 筛选调用缺少可信保存结果')
        const patch = manifest.filterPatch
        if (this.loadedRevision === result.loadedRevision && this.draftVersion === result.draftVersion) {
          if (JSON.stringify(this.getSnapshot()) !== this.loadedSnapshotFingerprint || !sameControlledFilterSnapshot(this.getSnapshot(), state.snapshot, patch)) {
            this.agentSyncRequired = true; this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 筛选保存后页面已变化，保持保护')
          }
          this.agentSyncRequired = false
          return true
        }
        const executed = this.controlledExecution && this.controlledExecution.callId === result.callId ? this.controlledExecution : null
        if (executed) {
          if (!sameWorkspace(this, executed.identity) || !executed.fingerprint || executed.fingerprint !== JSON.stringify(this.getSnapshot()) || !sameControlledFilterSnapshot(this.getSnapshot(), state.snapshot, patch)) {
            this.agentSyncRequired = true; this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 筛选执行与服务器草稿不一致，保持保护')
          }
        } else {
          if (manifest.loadedRevision !== this.loadedRevision || JSON.stringify(this.getSnapshot()) !== this.loadedSnapshotFingerprint || !sameControlledFilterTransition(this.getSnapshot(), state.snapshot, patch)) {
            this.agentSyncRequired = true; this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 筛选恢复含目标外变化，保持保护')
          }
        }
        const reloaded = await this.reloadAgentDraft(isCurrent, { skipPreserve: !!executed })
        if (!isCurrent() || reloaded.loadedRevision !== result.loadedRevision || reloaded.draftVersion !== result.draftVersion || !sameControlledFilterSnapshot(this.getSnapshot(), reloaded.snapshot, patch)) {
          this.agentSyncRequired = true; this.lockAgentWorkspace(true)
          throw new Error('RESULT_UNKNOWN: 筛选重载与可信草稿不一致，保持保护')
        }
        return true
      }
      if (manifest && manifest.patch) {
        if (manifest.callId !== result.callId || result.validation !== 'PASSED' || result.persistenceStatus !== 'SAVED') throw new Error('RESULT_UNKNOWN: 格式调用缺少可信保存结果')
        // A later poll can see the same committed call after the first authoritative reload.
        // Verify the unchanged local baseline against that same server draft before acknowledging it again.
        if (this.loadedRevision === result.loadedRevision && this.draftVersion === result.draftVersion) {
          if (JSON.stringify(this.getSnapshot()) !== this.loadedSnapshotFingerprint ||
            !sameControlledFormatSnapshot(this.getSnapshot(), state.snapshot, manifest.patch)) {
            this.agentSyncRequired = true
            this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 格式保存后页面已变化，保持保护')
          }
          this.agentSyncRequired = false
          return true
        }
        const executed = this.controlledExecution && this.controlledExecution.callId === result.callId ? this.controlledExecution : null
        if (executed) {
          if (!sameWorkspace(this, executed.identity) || !executed.fingerprint || executed.fingerprint !== JSON.stringify(this.getSnapshot())) throw new Error('RESULT_UNKNOWN: 格式执行后页面已变化，保持保护')
          if (!sameControlledFormatSnapshot(this.getSnapshot(), state.snapshot, manifest.patch)) {
            this.agentSyncRequired = true
            this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 格式执行回读与服务器草稿不一致，保留页面并核实')
          }
        } else {
          if (manifest.loadedRevision !== this.loadedRevision || JSON.stringify(this.getSnapshot()) !== this.loadedSnapshotFingerprint) throw new Error('RESULT_UNKNOWN: 格式恢复基线已变化，保持保护')
          if (!sameControlledFormatTransition(this.getSnapshot(), state.snapshot, manifest.patch)) {
            this.agentSyncRequired = true
            this.lockAgentWorkspace(true)
            throw new Error('RESULT_UNKNOWN: 格式草稿含目标外变化，保留页面并核实')
          }
        }
        const reloaded = await this.reloadAgentDraft(isCurrent, { skipPreserve: !!executed })
        if (!isCurrent() || reloaded.loadedRevision !== result.loadedRevision || reloaded.draftVersion !== result.draftVersion || !sameControlledFormatSnapshot(this.getSnapshot(), reloaded.snapshot, manifest.patch)) {
          this.agentSyncRequired = true
          this.lockAgentWorkspace(true)
          throw new Error('RESULT_UNKNOWN: 格式保存后服务器草稿与页面回读不一致，保持保护')
        }
        return true
      }
      if (!this.controlledExecution || this.controlledExecution.callId !== result.callId) {
        // Other pages may apply only a proved single committed transition to an unchanged baseline.
        if (!manifest || manifest.loadedRevision !== this.loadedRevision || JSON.stringify(this.getSnapshot()) !== this.loadedSnapshotFingerprint) return false
        await this.applyControlledWrite(manifest)
      }
      if (!isCurrent()) throw new Error('任务核实操作已切换，保持保护')
      await this.acceptControlledResult(result)
      if (!isCurrent()) throw new Error('任务核实操作已切换，保持保护')
      if (!sameControlledSnapshot(this.getSnapshot(), state.snapshot)) return false
      this.applyVersionStatus({ branchId: state.branchId, branchName: state.branchName, version: state.branchVersion,
        headCommitId: state.headCommitId, workbookVersion: state.workbookVersion, workingState: state.workingState }, true)
      this.agentSyncRequired = false
      return true
    },
    async applyAgentSnapshot(snapshot) {
      if (!this.agentLocked) throw new Error('工作区未锁定，拒绝执行')
      if (!this.univer || !this.univerAPI) throw new Error('表格引擎尚未就绪，保持保护状态')
      const checkpoint = this.getSnapshot()
      this.agentApplying = true
      this.aiExecutionInProgress = true
      try {
        // A bounded proposal is produced by the server's registered tools; never execute model code.
        this.loadUniverWorkbook(JSON.parse(JSON.stringify(snapshot)))
        await this.$nextTick()
        const hasFormulas = Object.values(snapshot.sheets || {}).some(sheet => Object.values(sheet.cellData || {}).some(row => Object.values(row).some(cell => cell && cell.f)))
        if (hasFormulas) {
          const engine = this.univerAPI.getFormula && this.univerAPI.getFormula()
          if (!engine || typeof engine.onCalculationResultApplied !== 'function') throw new Error('当前表格引擎不能确认公式计算完成，任务已暂停')
          await engine.onCalculationResultApplied(15000)
        }
        const actual = this.getSnapshot()
        if (!actual) throw new Error('执行后无法回读工作簿')
        return actual
      } catch (e) {
        if (checkpoint) this.loadUniverWorkbook(checkpoint)
        throw e
      } finally { this.agentApplying = false; this.aiExecutionInProgress = false }
    },
    async refreshAgentDraft() {
      return this.reloadAgentDraft()
    },
    async reloadAgentDraft(isCurrent = () => true, options = {}) {
      const identity = this.workspaceIdentity()
      this.agentSyncRequired = true
      this.lockAgentWorkspace(true)
      const state = await agentWorkspace(this.currentDocumentId)
      if (!isCurrent() || !sameWorkspace(this, identity) || String(state.workbookId) !== String(identity.workbookId) || String(state.branchId) !== String(identity.branchId)) throw new Error('工作簿、分支或任务核实操作已变化，忽略旧同步响应')
      if (!state.loadedRevision || !Number.isInteger(state.draftVersion) || !state.snapshot) throw new Error('服务器同步版本信息不完整')
      if (!options.skipPreserve) this.preserveLocalDraft()
      await this.applyAgentSnapshot(state.snapshot)
      if (!isCurrent() || !sameWorkspace(this, identity)) throw new Error('工作簿或任务核实操作已变化，保持保护状态')
      this.applyVersionStatus({ branchId: state.branchId, branchName: state.branchName, version: state.branchVersion,
        headCommitId: state.headCommitId, workbookVersion: state.workbookVersion, workingState: state.workingState }, true)
      this.draftVersion = state.draftVersion
      this.loadedRevision = state.loadedRevision
      this.loadedSnapshotFingerprint = JSON.stringify(this.getSnapshot())
      this.agentSyncRequired = false
      return state
    },
    workspaceIdentity() {
      return { workbookId: this.currentDocumentId, branchId: this.currentBranchId, generation: this.workspaceGeneration }
    },
    adoptSavedDraft(draft, snapshot) {
      this.draftVersion = draft.version
      this.loadedRevision = draft.loadedRevision || draft.contentHash
      this.loadedSnapshotFingerprint = JSON.stringify(snapshot)
    },
    preserveLocalDraft(force = false) {
      this.flushManualChanges(true)
      const snapshot = this.getSnapshot()
      if (!snapshot || (force !== true && JSON.stringify(snapshot) === this.loadedSnapshotFingerprint)) return
      const draft = { workbookId: this.currentDocumentId, branchId: this.currentBranchId, loadedRevision: this.loadedRevision,
        expectedDraftVersion: this.draftVersion, snapshot, plannedActions: this.pendingActions, executedChanges: this.pendingChanges, savedAt: Date.now() }
      this.conflictDrafts.push(draft)
      try { sessionStorage.setItem('agentReport:conflict:' + draft.workbookId + ':' + draft.savedAt, JSON.stringify(draft)) }
      catch (e) { throw new Error('本地编辑保留空间不足，已保持原内容和保护状态；请下载保留草稿') }
    },
    downloadConflictDrafts() {
      saveAs(new Blob([JSON.stringify(this.conflictDrafts)], { type: 'application/json;charset=utf-8' }), '保留的编辑草稿.json')
    },
    /** 初始化 Univer 表格引擎 */
    initUniver() {
      try {
        const container = this.$refs.univerContainer
        if (!container) {
          this.statusMessage = '错误：无法获取容器元素'
          return
        }

        const zhCNLocales = mergeLocales(
          {},
          UniverPresetSheetsCoreZhCN.default || UniverPresetSheetsCoreZhCN,
          sheetsConditionalFormattingZhCN.default || sheetsConditionalFormattingZhCN,
          sheetsFilterZhCN.default || sheetsFilterZhCN,
          sheetsDrawingZhCN.default || sheetsDrawingZhCN
        )

        const { univer, univerAPI } = createUniver({
          locale: LocaleType.ZH_CN,
          locales: {
            [LocaleType.ZH_CN]: zhCNLocales
          },
          theme: defaultTheme,
          presets: [
            UniverSheetsCorePreset({ container: container }),
            UniverSheetsConditionalFormattingPreset(),
            UniverSheetsFilterPreset(),
            UniverSheetsDrawingPreset()
          ]
        })

        this.univer = univer
        this.univerAPI = univerAPI
        this.registerManualChangeCapture()
        this.statusMessage = 'Univer 表格引擎初始化成功，请导入Excel文件或新建表格'
      } catch (error) {
        console.error('Univer 初始化失败:', error)
        this.statusMessage = 'Univer 初始化失败: ' + (error.message || String(error))
      }
    },

    /** 销毁 Univer 实例 */
    disposeUniver() {
      this.disposeManualChangeCapture()
      if (this.univer) {
        try {
          this.univer.dispose()
        } catch (e) {
          console.warn('销毁 Univer 实例时出错:', e)
        }
        this.univer = null
        this.univerAPI = null
        this.activeWorkbook = null
      }
    },

    /** 正确销毁指定工作簿 unit（Univer 顶层无 disposeUnit，须经实例服务） */
    disposeUnitById(unitId) {
      if (!unitId || !this.univer) return false
      try {
        const instanceService = this.univer._univerInstanceService
        if (instanceService && typeof instanceService.disposeUnit === 'function') {
          return instanceService.disposeUnit(unitId)
        }
      } catch (e) {
        console.warn('销毁工作簿时出错:', e)
      }
      return false
    },

    /** 重新创建 Univer 实例 */
    recreateUniver() {
      this.disposeUniver()
      this.$nextTick(() => {
        this.initUniver()
      })
    },

    /** 监听人工顶层 Command，并在 2 秒连续输入窗口内合并为一个变更组。 */
    registerManualChangeCapture() {
      this.disposeManualChangeCapture()
      if (!this.univerAPI || !this.univerAPI.Event) return
      const before = this.univerAPI.addEvent(this.univerAPI.Event.BeforeCommandExecute, event => {
        if ((this.historyPreview || (this.agentLocked && !this.agentApplying)) && event && event.type === 0 && String(event.id || '').indexOf('sheet.') === 0) {
          event.cancel = true
          return
        }
        if (!this.shouldCaptureManualCommand(event)) return
        if (!this.largeDataMode && !this.manualCommands.length) this.manualBeforeSnapshot = this.getSnapshot()
      })
      const after = this.univerAPI.addEvent(this.univerAPI.Event.CommandExecuted, event => {
        if (!this.shouldCaptureManualCommand(event)) return
        this.manualCommands.push({ id: event.id, type: event.type, params: event.params || {} })
        clearTimeout(this.manualCaptureTimer)
        this.manualCaptureTimer = setTimeout(() => this.flushManualChanges(), 2000)
      })
      this.changeCaptureDisposables = [before, after]
    },

    shouldCaptureManualCommand(event) {
      // CommandType.COMMAND = 0；选区 OPERATION 和内部 MUTATION 不重复记录。
      const commandId = String((event && event.id) || '')
      return !!event && event.type === 0 && !NON_MUTATING_SHEET_COMMANDS.has(commandId) &&
        !this.suspendChangeCapture && !this.aiExecutionInProgress && !this.historyPreview && !this.agentLocked &&
        this.hasData && !!this.currentDocumentId && commandId.indexOf('sheet.') === 0
    },

    disposeManualChangeCapture() {
      clearTimeout(this.manualCaptureTimer)
      this.manualCaptureTimer = null
      ;(this.changeCaptureDisposables || []).forEach(item => {
        try { if (item && typeof item.dispose === 'function') item.dispose() } catch (e) { /* ignore */ }
      })
      this.changeCaptureDisposables = []
      this.manualCommands = []
      this.manualBeforeSnapshot = null
    },

    flushManualChanges(skipAutosave) {
      if (!this.manualCommands.length || this.suspendChangeCapture || this.aiExecutionInProgress) return
      const commands = this.manualCommands.slice()
      const before = this.manualBeforeSnapshot
      const after = this.largeDataMode ? null : this.getSnapshot()
      this.manualCommands = []
      this.manualBeforeSnapshot = null
      this.manualCaptureTimer = null
      const capture = this.largeDataMode
        ? buildLightweightManualChanges(commands)
        : buildManualChanges(commands, before, after)
      if (!capture.changes.length) return
      this.pendingChanges = this.pendingChanges.concat(capture.changes)
      if (!skipAutosave) this.autosaveWorkingState()
      this.statusMessage = this.largeDataMode
        ? '大数据模式：已记录修改，将在停止操作 10 分钟后合并保存'
        : '已记录 ' + capture.changes.length + ' 项人工修改（尚未提交）'
    },

    // ============ 文档持久化 ============
    /** 拉取文档列表 */
    fetchDocuments() {
      listWorkbooks().then((res) => {
        if (res && res.code === 200) {
          this.documents = res.data || []
        }
      }).catch((e) => {
        console.error('获取文档列表失败:', e)
      })
    },

    /** 获取当前表格快照 */
    getSnapshot() {
      try {
        const workbook = this.univerAPI.getActiveWorkbook()
        if (!workbook) return null
        let snapshot = null
        try {
          snapshot = workbook.save()
        } catch (e) {
          console.warn('使用 save() 获取快照失败，回退手动快照:', e)
        }
        if (!snapshot) snapshot = this.manualSnapshot()
        return snapshot
      } catch (e) {
        console.error('获取表格快照失败:', e)
        return null
      }
    },

    /** 保存当前表格到后端（showMsg 控制是否弹成功/失败提示） */
    saveCurrentWorkbook(showMsg) {
      if (this.agentLocked) return Promise.reject(new Error('请先停止或完成智能体任务'))
      if (!this.hasData || !this.univerAPI) {
        if (showMsg) this.$modal.msgWarning('没有可保存的数据')
        return Promise.resolve()
      }
      if (this.persistBusy) return Promise.resolve()
      // 已纳入版本管理的工作簿，“保存”只保存分支草稿，不生成 Commit，也不覆盖 HEAD。
      if (this.currentDocumentId) {
        this.flushManualChanges(true)
        return this.autosaveWorkingState(true).then(() => {
          this.statusMessage = '当前分支草稿已保存（尚未提交）'
          if (showMsg) this.$modal.msgSuccess('草稿保存成功')
        }).catch(e => {
          if (showMsg) this.$modal.msgError('草稿保存失败：' + (e.message || String(e)))
        })
      }
      const creationIdentity = this.workspaceIdentity(), creationFingerprint = JSON.stringify(this.getSnapshot())
      let createdDocumentId = null
      return this.runPersistTask('正在保存表格…', updateProgress => {
        if (!sameWorkspace(this, creationIdentity) || creationFingerprint !== JSON.stringify(this.getSnapshot()) || this.agentLocked || this.agentSyncRequired || this.historyPreview) throw new Error('首次保存前工作区已变化，未发送创建请求')
        updateProgress(12, '正在生成表格快照…')
        const snapshot = this.getSnapshot()
        if (!snapshot) throw new Error('获取表格快照失败')
        updateProgress(35, '正在整理大数据…')
        const payload = {
          workbookId: this.currentDocumentId || null,
          name: this.currentDocumentName || snapshot.name || '未命名文档',
          type: this.currentDocumentType || 'created',
          version: this.currentDocumentVersion,
          content: JSON.stringify(snapshot)
        }
        updateProgress(48, '正在上传表格数据…')
        return saveWorkbook(payload, percent => {
          updateProgress(48 + Math.round(percent * 0.42), percent >= 100 ? '正在等待服务器完成保存…' : '正在上传表格数据…')
        }).then(res => {
          if (!res || res.code !== 200 || !res.data) throw new Error((res && res.msg) || '保存失败')
          createdDocumentId = res.data.workbookId
          return res
        })
      }).then((res) => {
        if (!sameWorkspace(this, creationIdentity) || creationFingerprint !== JSON.stringify(this.getSnapshot()) || this.agentLocked || this.agentSyncRequired || this.historyPreview) {
          this.fetchDocuments()
          throw new Error('新文档已创建（ID ' + createdDocumentId + '），页面已变化，未应用旧保存结果；请从文件列表核实')
        }
        if (res && res.code === 200 && res.data) {
          this.currentDocumentId = res.data.workbookId
          this.currentDocumentName = res.data.name
          this.currentDocumentType = res.data.type
          this.currentDocumentVersion = res.data.version
          this.statusMessage = '已保存：' + res.data.name
          this.loadVersionStatus(false)
          if (showMsg) this.$modal.msgSuccess('保存成功')
          this.fetchDocuments()
        } else {
          throw new Error((res && res.msg) || '保存失败')
        }
      }).catch((e) => {
        console.error('保存文档失败:', e)
        if (showMsg || createdDocumentId || sameWorkspace(this, creationIdentity)) this.$modal.msgError('保存失败：' + (e && e.message ? e.message : String(e)))
      })
    },

    /** 在浏览器完成一次绘制后再执行同步快照，确保进度层先显示。 */
    nextUiFrame() {
      return new Promise(resolve => {
        this.$nextTick(() => setTimeout(resolve, 50))
      })
    },

    updatePersistProgress(percentage, title) {
      this.persistProgress = Math.max(this.persistProgress, Math.min(99, Number(percentage) || 0))
      if (title) this.persistTitle = title
    },

    runPersistTask(title, executor) {
      if (this.persistBusy) return Promise.reject(new Error('已有保存或提交任务正在进行'))
      this.cancelExcelImport()
      clearTimeout(this.autosaveTimer)
      this.autosaveTimer = null
      this.persistBusy = true
      this.persistProgress = 3
      this.persistTitle = title
      clearInterval(this.persistProgressTimer)
      this.persistProgressTimer = setInterval(() => {
        if (this.persistProgress < 88) this.persistProgress += this.persistProgress < 45 ? 2 : 1
      }, 900)
      const update = (percentage, nextTitle) => this.updatePersistProgress(percentage, nextTitle)
      return this.nextUiFrame()
        .then(() => executor(update))
        .then(result => {
          this.persistProgress = 100
          this.persistTitle = title.indexOf('提交') >= 0
            ? '提交完成'
            : (title.indexOf('打开') >= 0 || title.indexOf('切换') >= 0 ? '文件打开完成' : '保存完成')
          return result
        })
        .catch(error => {
          this.persistTitle = title.indexOf('提交') >= 0
            ? '提交失败'
            : (title.indexOf('打开') >= 0 || title.indexOf('切换') >= 0 ? '文件打开失败' : '保存失败')
          throw error
        })
        .finally(() => {
          clearInterval(this.persistProgressTimer)
          this.persistProgressTimer = null
          setTimeout(() => {
            this.persistBusy = false
            this.persistProgress = 0
          }, 350)
        })
    },

    /** 打开文档 */
    loadDocument(doc) {
      if (this.agentLocked) return this.$modal.msgWarning('任务运行或待核实期间不能切换工作簿')
      if (!doc || !doc.workbookId || this.persistBusy) return
      if (String(doc.workbookId) === String(this.currentDocumentId) && !this.historyPreview) return
      const targetWorkbookId = doc.workbookId
      const loadGeneration = ++this.workspaceGeneration
      this.runPersistTask('正在切换文件…', updateProgress => {
        // 切换前先收口当前编辑，避免大数据模式尚未触发的延迟保存丢失。
        this.flushManualChanges(true)
        let saveCurrent = Promise.resolve()
        if (this.currentDocumentId && this.currentBranchId && this.pendingChanges.length) {
          const previousWorkbookId = this.currentDocumentId
          updateProgress(8, '正在保存当前文件的修改…')
          const snapshot = this.getSnapshot()
          if (!snapshot) throw new Error('无法保存当前文件的修改')
          const identity = this.workspaceIdentity()
          const payload = workingStatePayload(this, snapshot)
          saveCurrent = saveWorkingState(previousWorkbookId, payload, percent => {
            updateProgress(8 + Math.round(percent * 0.2), '正在保存当前文件的修改…')
          }).then(res => {
            if (!res || res.code !== 200) throw new Error((res && res.msg) || '当前文件保存失败')
            if (!sameWorkspace(this, identity)) throw new Error('工作簿已变化，忽略旧保存响应')
            this.adoptSavedDraft(res.data, snapshot)
          })
        }
        return saveCurrent.then(() => {
          updateProgress(32, '正在下载目标文件…')
          // 文档内容和版本状态并行获取，避免串行等待。
          return Promise.all([
            getWorkbook(targetWorkbookId),
            getVersionStatus(targetWorkbookId).catch(() => null)
          ])
        }).then(results => {
          if (loadGeneration !== this.workspaceGeneration) throw new Error('文件加载已过期，忽略旧响应')
          const workbookRes = results[0]
          const statusRes = results[1]
          if (!workbookRes || workbookRes.code !== 200 || !workbookRes.data) {
            throw new Error((workbookRes && workbookRes.msg) || '加载文档失败')
          }
          const wb = workbookRes.data
          const status = statusRes && statusRes.code === 200 ? statusRes.data : null
          if (!status || !status.loadedRevision || !status.content) throw new Error('无法获取一致的文件内容与版本，请稍后重试')
          const draft = status && status.workingState
          const content = status.content
          if (!content) throw new Error('文档内容为空')
          updateProgress(55, draft && draft.content ? '正在解析未提交草稿…' : '正在解析文件数据…')
          const snapshot = JSON.parse(content)
          updateProgress(68, '正在构建大数据表格…')
          this.preserveLocalDraft()
          this.currentDocumentId = wb.workbookId
          this.currentDocumentName = wb.name
          this.currentDocumentType = wb.type
          this.currentDocumentVersion = wb.version
          this.loadUniverWorkbook(snapshot)
          this.applyVersionStatus(status, true)
          this.loadedSnapshotFingerprint = JSON.stringify(this.getSnapshot())
          this.historyPreview = false
          updateProgress(94, '正在完成文件切换…')
          this.statusMessage = draft && draft.content
            ? '已打开并恢复未提交草稿：' + wb.name
            : '已打开：' + wb.name
          return wb
        })
      }).catch((e) => {
        console.error('加载文档失败:', e)
        this.$modal.msgError('加载文档失败：' + (e.message || String(e)))
      })
    },

    /** 删除文档 */
    deleteDocument(doc) {
      if (!doc || !doc.workbookId) return
      this.$modal.confirm('确认删除文档「' + doc.name + '」？').then(() => {
        deleteWorkbook(doc.workbookId).then((res) => {
          if (res && res.code === 200) {
            this.$modal.msgSuccess('删除成功')
            if (this.currentDocumentId === doc.workbookId) {
              this.currentDocumentId = null
              this.currentDocumentName = ''
              this.currentDocumentType = ''
          this.currentDocumentVersion = null
          this.resetVersionState()
            }
            this.fetchDocuments()
          } else {
            this.$modal.msgError((res && res.msg) || '删除失败')
          }
        }).catch((e) => {
          console.error('删除文档失败:', e)
          this.$modal.msgError('删除失败')
        })
      }).catch(() => {})
    },

    /** 打开重命名对话框 */
    openRenameDialog(doc) {
      if (!doc || !doc.workbookId) return
      this.renameTarget = doc
      this.renameName = doc.name || ''
      this.renameDialogVisible = true
    },

    /** 确认重命名 */
    confirmRename() {
      const name = (this.renameName || '').trim()
      if (!name) {
        this.$modal.msgWarning('文档名称不能为空')
        return
      }
      if (!this.renameTarget || !this.renameTarget.workbookId) {
        this.renameDialogVisible = false
        return
      }
      if (name === this.renameTarget.name) {
        this.renameDialogVisible = false
        return
      }
      this.renameLoading = true
      renameWorkbook({ workbookId: this.renameTarget.workbookId, name }).then((res) => {
        if (res && res.code === 200 && res.data) {
          this.$modal.msgSuccess('重命名成功')
          if (this.currentDocumentId === res.data.workbookId) {
            this.currentDocumentName = res.data.name
            this.statusMessage = '已重命名：' + res.data.name
          }
          this.renameDialogVisible = false
          this.fetchDocuments()
        } else {
          this.$modal.msgError((res && res.msg) || '重命名失败')
        }
      }).catch((e) => {
        console.error('重命名失败:', e)
        this.$modal.msgError('重命名失败')
      }).finally(() => {
        this.renameLoading = false
      })
    },

    /** 生成不重复的文件名（与现有列表同名则追加时间戳） */
    uniqueName(base) {
      let name = (base || '未命名文档').trim()
      if (this.documents.some((d) => d.name === name)) {
        name = name + '_' + this.formatDate(new Date())
      }
      return name
    },

    getList() {},

    /** 上传前校验 */
    beforeUpload(file) {
      if (!file || !file.name) return false
      if (this.agentSyncRequired || this.persistBusy || this.historyPreview) { this.$modal.msgWarning('请先结束保存、预览或待核实状态再导入'); return false }
      if (this.agentLocked) { this.$modal.msgWarning('请先停止或完成智能体任务'); return false }
      const isExcel =
        file.name.endsWith('.xlsx') ||
        file.name.endsWith('.xls') ||
        file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        file.type === 'application/vnd.ms-excel'
      if (!isExcel) {
        this.$modal.msgError('只能上传 Excel 文件（.xlsx 或 .xls）')
        return false
      }
      const isLt10M = file.size / 1024 / 1024 < 10
      if (!isLt10M) {
        this.$modal.msgError('文件大小不能超过 10MB')
        return false
      }
      return true
    },

    cancelExcelImport() {
      const pending = this.pendingExcelImport
      this.pendingExcelImport = null
      this.excelImportGeneration++
      if (pending && pending.reader && pending.reader.readyState === 1) pending.reader.abort()
    },

    beginExcelImport(file) {
      if (!this.beforeUpload(file)) return null
      this.cancelExcelImport()
      const context = { generation: this.excelImportGeneration, identity: this.workspaceIdentity(),
        baseline: this.excelImportBaseline(), fingerprint: JSON.stringify(this.getSnapshot()), reader: null }
      this.pendingExcelImport = context
      this.statusMessage = '正在解析 Excel 文件...'
      return context
    },

    excelImportBaseline() {
      return JSON.stringify([this.headCommitId, this.branchVersion, this.draftVersion, this.loadedRevision])
    },

    isExcelImportCurrent(context) {
      return context && this.pendingExcelImport === context && context.generation === this.excelImportGeneration &&
        sameWorkspace(this, context.identity) && !this.agentLocked && !this.agentSyncRequired && !this.persistBusy &&
        !this.historyPreview && context.baseline === this.excelImportBaseline() &&
        context.fingerprint === JSON.stringify(this.getSnapshot())
    },

    expireExcelImport(context) {
      if (this.pendingExcelImport !== context) return
      this.cancelExcelImport()
      if (sameWorkspace(this, context.identity) && !this.agentLocked && !this.agentSyncRequired && !this.persistBusy && !this.historyPreview) {
        this.statusMessage = '导入期间工作区已变化，当前内容已保留；请重新选择文件'
        this.$modal.msgWarning(this.statusMessage)
      }
    },

    failExcelImport(context, error) {
      if (!this.isExcelImportCurrent(context)) { this.expireExcelImport(context); return }
      this.cancelExcelImport()
      this.statusMessage = '导入失败：' + (error.message || String(error))
      this.$modal.msgError(this.statusMessage)
    },

    /** 每次选择只启动本地解析，不向 upload 的占位地址提交文件。 */
    handleFileChange(file) {
      if (file && file.raw) return this.importExcelWithXLSX(file.raw)
    },

    importExcelWithXLSX(file) {
      const context = this.beginExcelImport(file)
      if (!context) return
      try {
        const reader = new FileReader()
        context.reader = reader
        reader.onload = e => {
          if (!this.isExcelImportCurrent(context)) { this.expireExcelImport(context); return }
          let snapshot
          try {
            const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' })
            snapshot = this.convertXlsxWorkbookToUniverData(workbook, file.name)
          } catch (error) {
            if (!this.isExcelImportCurrent(context)) { this.expireExcelImport(context); return }
            this.statusMessage = 'XLSX 解析失败，尝试备用方案...'
            this.importExcelWithLuckyExcel(file, context)
            return
          }
          this.completeExcelImport(snapshot, file, context, false)
        }
        reader.onerror = () => this.failExcelImport(context, new Error('文件读取失败'))
        reader.onabort = () => this.failExcelImport(context, new Error('文件读取已取消'))
        reader.readAsArrayBuffer(file)
      } catch (error) { this.failExcelImport(context, error) }
      return context
    },

    importExcelWithLuckyExcel(file, existingContext) {
      const context = existingContext || this.beginExcelImport(file)
      if (!context) return
      if (!this.isExcelImportCurrent(context)) { this.expireExcelImport(context); return }
      context.reader = null
      // A parser may call success/error more than once; only its first callback is eligible.
      let received = false
      try {
        LuckyExcel.transformExcelToLucky(file, result => {
          if (received) return
          received = true
          if (!this.isExcelImportCurrent(context)) { this.expireExcelImport(context); return }
          if (!result || !result.sheets || !result.sheets.length) return this.failExcelImport(context, new Error('文件中没有有效数据'))
          let snapshot
          try { snapshot = this.convertLuckyJsonToUniver(result, file.name) }
          catch (error) { this.failExcelImport(context, error); return }
          this.completeExcelImport(snapshot, file, context, true)
        }, error => {
          if (received) return
          received = true
          this.failExcelImport(context, error || new Error('备用解析失败'))
        })
      } catch (error) { if (!received) this.failExcelImport(context, error) }
      return context
    },

    completeExcelImport(snapshot, file, context, fallback) {
      if (!this.isExcelImportCurrent(context)) { this.expireExcelImport(context); return }
      try {
        // Do not enter loadUniverWorkbook's delayed engine-recreation path with stale import data.
        if (!this.univer || !this.univerAPI) throw new Error('表格引擎尚未就绪，请稍后重新导入')
        if (!snapshot || !snapshot.sheetOrder || !snapshot.sheetOrder.length) throw new Error('文件中没有有效工作表')
        this.preserveLocalDraft(true)
      } catch (error) { this.failExcelImport(context, error); return }
      this.cancelExcelImport()
      try { this.loadUniverWorkbook(snapshot) }
      catch (error) {
        this.agentSyncRequired = true
        this.lockAgentWorkspace(true)
        this.statusMessage = '导入载入失败，原页面副本已保留，请下载草稿并重新打开文档'
        this.$modal.msgError(this.statusMessage)
        return
      }
      this.currentDocumentId = null
      this.currentDocumentName = this.uniqueName(snapshot.name || file.name)
      this.currentDocumentType = 'imported'
      this.currentDocumentVersion = null
      this.resetVersionState()
      this.statusMessage = fallback ? 'Excel 文件导入成功（备用方案）' : 'Excel 文件导入成功'
      this.$modal.msgSuccess('导入成功')
      if (!this.largeDataMode) this.saveCurrentWorkbook(false)
      else this.statusMessage = 'Excel 导入成功，已自动启用大数据模式；请点击“保存”完成首次保存'
    },

    /** 转换 LuckyExcel 数据为 Univer 数据格式 */
    convertLuckyJsonToUniver(luckyJson, fileName) {
      const sheets = {}
      const sheetOrder = []

      luckyJson.sheets.forEach((sheet, index) => {
        const sheetId = 'sheet-' + index
        sheetOrder.push(sheetId)

        const cellData = {}
        const cellSource = sheet.celldata || sheet.data || []

        if (cellSource.length > 0) {
          if (
            Array.isArray(cellSource) &&
            typeof cellSource[0] === 'object' &&
            cellSource[0].r !== undefined
          ) {
            cellSource.forEach(cell => {
              const r = cell.r
              const c = cell.c
              if (!cellData[r]) cellData[r] = {}
              const v = cell.v || {}
              const rawValue = v.v !== undefined ? v.v : ''
              cellData[r][c] = {
                v: rawValue,
                t: typeof rawValue === 'number' ? 2 : typeof rawValue === 'boolean' ? 3 : 1,
                f: v.f ? (v.f.startsWith('=') ? v.f : '=' + v.f) : undefined
              }
            })
          } else if (Array.isArray(cellSource)) {
            cellSource.forEach((row, r) => {
              if (!Array.isArray(row)) return
              row.forEach((cell, c) => {
                if (!cell) return
                if (!cellData[r]) cellData[r] = {}
                const rawValue = cell.v !== undefined ? cell.v : (cell.m || '')
                cellData[r][c] = {
                  v: rawValue,
                  t: typeof rawValue === 'number' ? 2 : typeof rawValue === 'boolean' ? 3 : 1,
                  f: cell.f ? (cell.f.startsWith('=') ? cell.f : '=' + cell.f) : undefined
                }
              })
            })
          }
        }

        const mergeData = []
        if (sheet.config && sheet.config.merge) {
          Object.keys(sheet.config.merge).forEach(key => {
            const m = sheet.config.merge[key]
            mergeData.push({
              startRow: m.r,
              endRow: m.r + m.rs - 1,
              startColumn: m.c,
              endColumn: m.c + m.cs - 1
            })
          })
        }

        sheets[sheetId] = {
          id: sheetId,
          name: sheet.name || ('Sheet' + (index + 1)),
          rowCount: sheet.rowCount || (sheet.row && sheet.row.len) || 100,
          columnCount: sheet.columnCount || (sheet.column && sheet.column.len) || 26,
          cellData: cellData,
          mergeData: mergeData
        }
      })

      return {
        id: 'workbook-' + Date.now(),
        name: fileName || '工作簿',
        sheetOrder: sheetOrder,
        sheets: sheets,
        activeSheet: sheetOrder[0]
      }
    },

    /** 转换 XLSX workbook 到 Univer 数据 */
    convertXlsxWorkbookToUniverData(workbook, fileName) {
      const sheets = {}
      const sheetOrder = []

      workbook.SheetNames.forEach((sheetName, index) => {
        const sheetId = 'sheet-' + index
        sheetOrder.push(sheetId)
        const ws = workbook.Sheets[sheetName]
        const cellData = {}
        let maxRow = 0
        let maxCol = 0

        // 只遍历真实存在的单元格，避免对大范围稀疏表逐格探测空白区域。
        Object.keys(ws).forEach(cellAddr => {
          if (cellAddr.charAt(0) === '!' || !/^[A-Z]+[1-9]\d*$/i.test(cellAddr)) return
          const cell = ws[cellAddr]
          if (!cell) return
          const pos = XLSX.utils.decode_cell(cellAddr)
          const r = pos.r
          const c = pos.c
          if (!cellData[r]) cellData[r] = {}
          const t = cell.t === 'b' ? 3 : cell.t === 'n' || cell.t === 'd' ? 2 : 1
          const errors = { 0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?', 36: '#NUM!', 42: '#N/A', 43: '#GETTING_DATA' }
          cellData[r][c] = {
            v: cell.t === 'e' ? (errors[cell.v] || cell.w || '#VALUE!') : cell.v !== undefined ? cell.v : (cell.w || ''),
            t: t,
            f: cell.f ? (cell.f.startsWith('=') ? cell.f : '=' + cell.f) : undefined
          }
          if (r > maxRow) maxRow = r
          if (c > maxCol) maxCol = c
        })

        const mergeData = []
        if (ws['!merges'] && ws['!merges'].length > 0) {
          ws['!merges'].forEach(m => {
            mergeData.push({
              startRow: m.s.r,
              endRow: m.e.r,
              startColumn: m.s.c,
              endColumn: m.e.c
            })
          })
        }

        sheets[sheetId] = {
          id: sheetId,
          name: sheetName,
          rowCount: Math.max(maxRow + 50, 100),
          columnCount: Math.max(maxCol + 10, 26),
          cellData: cellData,
          mergeData: mergeData
        }
      })

      return {
        id: 'workbook-' + Date.now(),
        name: fileName || '工作簿',
        sheetOrder: sheetOrder,
        sheets: sheets,
        activeSheet: sheetOrder[0]
      }
    },

    /** 加载 Univer 工作簿数据 */
    loadUniverWorkbook(workbookData) {
      if (!this.univer || !this.univerAPI) {
        this.recreateUniver()
        const self = this
        setTimeout(() => {
          self.loadUniverWorkbook(workbookData)
        }, 500)
        return
      }

      try {
        this.suspendChangeCapture = true
        this.applyDataMode(workbookData)
        this.disposeUnitById(this.currentWorkbookId)

        const workbook = this.univer.createUnit(UniverInstanceType.UNIVER_SHEET, workbookData)
        // 优先用 facade FWorkbook（含 setValue/merge 等方法），facade 不可用时退回核心 workbook
        this.activeWorkbook = (this.univerAPI.getActiveWorkbook && this.univerAPI.getActiveWorkbook()) || workbook
        if (workbook && typeof workbook.getUnitId === 'function') {
          this.currentWorkbookId = workbook.getUnitId()
        } else if (workbookData.id) {
          this.currentWorkbookId = workbookData.id
        }
        this.hasData = true
      } catch (error) {
        console.error('创建 Univer 工作簿失败:', error)
        throw error
      } finally {
        this.$nextTick(() => { this.suspendChangeCapture = false })
      }
    },

    /** 根据实际有值的单元格自动切换大数据模式。 */
    applyDataMode(workbookData) {
      const stats = {
        sheetCount: 0,
        populatedCells: 0,
        maxUsedRows: 0,
        maxUsedColumns: 0,
        sheets: {}
      }
      const sheets = (workbookData && workbookData.sheets) || {}
      const ids = (workbookData && workbookData.sheetOrder) || Object.keys(sheets)
      ids.forEach(sheetId => {
        const sheet = sheets[sheetId] || {}
        const cellData = sheet.cellData || {}
        let populatedCells = 0
        let maxUsedRows = 0
        let maxUsedColumns = 0
        Object.keys(cellData).forEach(rowKey => {
          const row = cellData[rowKey] || {}
          const columns = Object.keys(row)
          if (!columns.length) return
          populatedCells += columns.length
          maxUsedRows = Math.max(maxUsedRows, Number(rowKey) + 1)
          columns.forEach(columnKey => {
            maxUsedColumns = Math.max(maxUsedColumns, Number(columnKey) + 1)
          })
        })
        stats.sheetCount += 1
        stats.populatedCells += populatedCells
        stats.maxUsedRows = Math.max(stats.maxUsedRows, maxUsedRows)
        stats.maxUsedColumns = Math.max(stats.maxUsedColumns, maxUsedColumns)
        stats.sheets[sheetId] = { populatedCells, maxUsedRows, maxUsedColumns }
      })
      this.workbookStats = stats
      this.largeDataMode = stats.populatedCells >= LARGE_DATA_CELL_THRESHOLD ||
        stats.maxUsedRows >= LARGE_DATA_ROW_THRESHOLD
    },

    formatLargeDataStats() {
      const stats = this.workbookStats || {}
      return Number(stats.maxUsedRows || 0).toLocaleString() + ' 行 / ' +
        Number(stats.populatedCells || 0).toLocaleString() + ' 单元格'
    },

    captureExportSnapshot() {
      if (this.persistBusy || this.agentApplying || this.aiExecutionInProgress || this.suspendChangeCapture || this.pendingExcelImport) throw new Error('请等待当前载入或写入完成后再导出')
      if (!this.hasData || !this.univerAPI) throw new Error('没有可导出的数据')
      const workbook = this.univerAPI.getActiveWorkbook()
      if (!workbook || typeof workbook.save !== 'function') throw new Error('无法获取完整工作簿快照，已停止导出')
      const snapshot = workbook.save()
      if (!snapshot || !snapshot.sheets) throw new Error('完整快照不可用，已停止导出')
      return JSON.parse(JSON.stringify(snapshot, (key, value) => {
        if (typeof value === 'number' && !Number.isFinite(value)) throw new Error('快照含非有限数字，已停止导出')
        return value
      }))
    },

    exportFileName(snapshot, suffix) {
      const name = String(this.currentDocumentName || snapshot.name || '表格导出').replace(/[\x00-\x1f<>:"/\\|?*]/g, '-').slice(0, 100)
      return name + '_' + this.formatDate(new Date()) + suffix
    },

    handleExport() {
      try {
        const snapshot = this.captureExportSnapshot(), workbook = this.convertUniverSnapshotToXlsx(snapshot)
        const data = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
        const name = this.exportFileName(snapshot, '.xlsx')
        saveAs(new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name)
        this.statusMessage = '已导出 Excel 数据与公式：' + name + '；完整样式及配置可另存 JSON 备份'
        this.$modal.msgSuccess('Excel 数据导出成功')
      } catch (error) {
        this.statusMessage = '导出失败：' + (error.message || String(error))
        this.$modal.msgError(this.statusMessage)
      }
    },

    handleSnapshotExport() {
      try {
        const snapshot = this.captureExportSnapshot(), name = this.exportFileName(snapshot, '.snapshot.json')
        saveAs(new Blob([JSON.stringify(snapshot, null, 2)], { type: 'application/json;charset=utf-8' }), name)
        this.statusMessage = '已下载完整快照备份：' + name
        this.$modal.msgSuccess('完整快照备份已下载')
      } catch (error) {
        this.statusMessage = '备份失败：' + (error.message || String(error))
        this.$modal.msgError(this.statusMessage)
      }
    },

    /** 手动构建快照（当 save() 失败时备用） */
    manualSnapshot() {
      try {
        const workbook = this.univerAPI.getActiveWorkbook()
        if (!workbook) return null

        const sheets = {}
        const sheetOrder = []
        const workbookSheets = workbook.getSheets()
        let activeSheet = null

        workbookSheets.forEach((sheet, index) => {
          const sheetId = 'sheet-' + index
          sheetOrder.push(sheetId)

          const cellData = {}
          // 兼容 facade（getMaxRows/getMaxColumns）与核心（getRowCount/getColumnCount）
          const rowCount = typeof sheet.getMaxRows === 'function' ? sheet.getMaxRows() : sheet.getRowCount()
          const colCount = typeof sheet.getMaxColumns === 'function' ? sheet.getMaxColumns() : sheet.getColumnCount()

          for (let r = 0; r < Math.min(rowCount, 5000); r++) {
            for (let c = 0; c < Math.min(colCount, 500); c++) {
              try {
                const range = sheet.getRange(r, c, 1, 1)
                if (!range) continue
                const value = range.getValue()
                if (value === null || value === undefined || value === '') continue
                if (!cellData[r]) cellData[r] = {}
                cellData[r][c] = {
                  v: value,
                  t: typeof value === 'number' ? 1 : 2
                }
              } catch (e) {
                // skip error cell
              }
            }
          }

          sheets[sheetId] = {
            id: sheetId,
            name: (typeof sheet.getSheetName === 'function' ? sheet.getSheetName() : sheet.getName()) || ('Sheet' + (index + 1)),
            rowCount: rowCount,
            columnCount: colCount,
            cellData: cellData
          }

          if (sheet.isActive()) {
            activeSheet = sheetId
          }
        })

        return {
          id: 'workbook-' + Date.now(),
          name: workbook.getName() || '工作簿',
          sheetOrder: sheetOrder,
          sheets: sheets,
          activeSheet: activeSheet || sheetOrder[0]
        }
      } catch (e) {
        console.error('手动构建快照失败:', e)
        return null
      }
    },

    convertUniverSnapshotToXlsx(snapshot) { return exportWorkbook(snapshot) },
    convertUniverSheetToXlsx(sheet) { return exportSheet(sheet) },

    /** 创建新的空白表格 */
    createNewSheet() {
      if (this.agentLocked) return this.$modal.msgWarning('请先停止或完成智能体任务')
      const workbookData = {
        id: 'workbook-' + Date.now(),
        name: '新建工作簿_' + this.formatDate(new Date()),
        sheetOrder: ['sheet-0'],
        sheets: {
          'sheet-0': {
            id: 'sheet-0',
            name: 'Sheet1',
            rowCount: 200,
            columnCount: 52,
            cellData: {
              0: {
                0: { v: '欢迎使用 Univer 表格', t: 2 },
                1: { v: '请在此处编辑或导入 Excel 文件', t: 2 }
              },
              1: {
                0: { v: '提示：', t: 2 },
                1: { v: '1. 点击左上角"导入Excel"按钮导入文件', t: 2 },
                2: { v: '2. 点击"导出Excel"按钮导出当前内容', t: 2 }
              },
              2: {
                1: { v: '3. 支持基础单元格编辑', t: 2 },
                2: { v: 10, t: 1 },
                3: { v: 20, t: 1 }
              }
            }
          }
        },
        activeSheet: 'sheet-0'
      }

      this.loadUniverWorkbook(workbookData)
      this.currentDocumentId = null
      this.currentDocumentName = workbookData.name
      this.currentDocumentType = 'created'
      this.currentDocumentVersion = null
      this.resetVersionState()
      this.statusMessage = '已创建新的空白工作簿'
      this.$modal.msgSuccess('创建成功')
      // 新建后自动保存
      this.saveCurrentWorkbook(false)
    },

    /** 清空当前数据 */
    clearSheet() {
      if (this.agentLocked) return this.$modal.msgWarning('请先停止或完成智能体任务')
      this.$modal.confirm('确认清空当前表格中的所有数据？').then(() => {
        this.disposeUnitById(this.currentWorkbookId)
        this.currentWorkbookId = null
        this.activeWorkbook = null
        this.hasData = false
        this.currentDocumentId = null
        this.currentDocumentName = ''
        this.currentDocumentType = ''
        this.currentDocumentVersion = null
        this.resetVersionState()
        this.largeDataMode = false
        this.workbookStats = { sheetCount: 0, populatedCells: 0, maxUsedRows: 0, maxUsedColumns: 0, sheets: {} }
        this.statusMessage = '数据已清空，请导入Excel文件或新建表格'
        this.$modal.msgSuccess('清空成功')
      }).catch(() => {})
    },

    /** 格式化日期 */
    formatDate(date) {
      const y = date.getFullYear()
      const m = String(date.getMonth() + 1).padStart(2, '0')
      const d = String(date.getDate()).padStart(2, '0')
      const h = String(date.getHours()).padStart(2, '0')
      const min = String(date.getMinutes()).padStart(2, '0')
      const s = String(date.getSeconds()).padStart(2, '0')
      return y + m + d + '_' + h + min + s
    },

    /** AI 执行结果回调：把执行概要刷到顶部状态条 */
    onAiExecuted(exec) {
      if (!exec || !exec.summary) return
      const s = exec.summary
      const ok = s.ok || 0
      const failed = s.failed || 0
      let msg = 'AI 执行完成：成功 ' + ok + ' 条'
      if (failed > 0) msg += '，失败 ' + failed + ' 条'
      if (exec.notes && exec.notes.length) msg += '（' + exec.notes.join('；') + '）'
      this.statusMessage = msg
      if (failed === 0 && exec.executedChanges && exec.executedChanges.length) {
        this.pendingChanges = this.pendingChanges.concat(exec.executedChanges)
        this.pendingActions = this.pendingActions.concat(exec.plannedActions || [])
        this.autosaveWorkingState()
      }
    },

    /** AI 动作组失败时恢复整本工作簿 checkpoint。 */
    rollbackAiOperation(snapshot) {
      if (!snapshot) return
      try {
        this.loadUniverWorkbook(snapshot)
      } catch (e) {
        console.error('恢复 AI 操作 checkpoint 失败:', e)
        this.$modal.msgError('动作执行失败，且自动恢复工作簿失败，请重新打开文档')
      }
    },

    resetVersionState() {
      this.branchCheckoutUncertain = false;
      this.cancelExcelImport()
      this.workspaceGeneration++
      this.loadedRevision = null
      this.loadedSnapshotFingerprint = null
      this.draftVersion = -1
      this.currentBranchId = null
      this.currentBranchName = 'main'
      this.selectedBranchId = null
      this.branches = []
      this.branchVersion = null
      this.headCommitId = null
      this.pendingChanges = []
      this.pendingActions = []
      this.commitHistory = []
      this.historyPreview = false
    },

    /** 只应用版本元数据；工作簿内容由调用方决定何时加载。 */
    applyVersionStatus(data, contentLoaded = false) {
      if (!data) {
        this.resetVersionState()
        return null
      }
      if (!contentLoaded && this.loadedRevision) return data.workingState
      this.currentBranchId = data.branchId
      this.currentBranchName = data.branchName || 'main'
      this.selectedBranchId = data.branchId
      this.branchVersion = data.version
      this.headCommitId = data.headCommitId
      if (data.workbookVersion !== undefined) this.currentDocumentVersion = data.workbookVersion
      const draft = data.workingState
      // Metadata dialogs must not bless stale local content with a newer token.
      this.draftVersion = draft ? draft.version : -1
      this.loadedRevision = data.loadedRevision || (draft && draft.contentHash) || data.contentHash
      this.pendingChanges = []
      this.pendingActions = []
      if (draft && draft.executedChangesJson) {
        try { this.pendingChanges = JSON.parse(draft.executedChangesJson) || [] } catch (e) { /* ignore */ }
      }
      if (draft && draft.draftActionsJson) {
        try { this.pendingActions = JSON.parse(draft.draftActionsJson) || [] } catch (e) { /* ignore */ }
      }
      return draft
    },

    loadVersionStatus(restoreDraft) {
      if (!this.currentDocumentId) return Promise.resolve()
      const identity = this.workspaceIdentity()
      return getVersionStatus(this.currentDocumentId).then(res => {
        if (!sameWorkspace(this, identity)) return
        if (res && res.code === 200 && res.data) {
          // Only initial save or an actual reload can adopt fresh draft tokens.
          if (restoreDraft || !this.loadedRevision) {
            if (!res.data.content || !res.data.loadedRevision) throw new Error('服务器内容与版本不完整，不能更新页面版本')
            this.preserveLocalDraft()
            this.loadUniverWorkbook(JSON.parse(res.data.content))
            this.applyVersionStatus(res.data, true)
            this.loadedSnapshotFingerprint = JSON.stringify(this.getSnapshot())
          }
        }
      }).catch(e => console.warn('加载版本状态失败:', e))
    },

    autosaveWorkingState(strict) {
      if (!this.currentDocumentId || !this.currentBranchId || !this.pendingChanges.length) return Promise.resolve()
      if (this.largeDataMode && !strict) {
        clearTimeout(this.autosaveTimer)
        this.autosaveTimer = setTimeout(() => {
          this.autosaveTimer = null
          this.performAutosaveWorkingState(false)
        }, LARGE_DATA_AUTOSAVE_DELAY)
        return Promise.resolve()
      }
      clearTimeout(this.autosaveTimer)
      this.autosaveTimer = null
      return this.performAutosaveWorkingState(strict)
    },

    performAutosaveWorkingState(strict) {
      if (this.agentLocked || this.agentSyncRequired) return strict ? Promise.reject(new Error('工作区待核实，编辑已保留')) : Promise.resolve()
      if (!this.currentDocumentId || !this.currentBranchId || !this.pendingChanges.length) return Promise.resolve()
      if (this.persistBusy) return strict ? Promise.reject(new Error('已有保存或提交任务正在进行')) : Promise.resolve()
      return this.runPersistTask(strict ? '正在保存草稿…' : '正在自动保存大数据表格…', updateProgress => {
        updateProgress(12, '正在生成表格快照…')
        const snapshot = this.getSnapshot()
        if (!snapshot) throw new Error('无法获取当前工作簿快照')
        updateProgress(35, '正在整理草稿数据…')
        const identity = this.workspaceIdentity()
        const payload = workingStatePayload(this, snapshot)
        updateProgress(48, '正在上传草稿数据…')
        return saveWorkingState(this.currentDocumentId, payload, percent => {
          updateProgress(48 + Math.round(percent * 0.42), percent >= 100 ? '正在等待服务器保存草稿…' : '正在上传草稿数据…')
        }).then(res => {
          if (!res || res.code !== 200) throw new Error((res && res.msg) || '保存分支草稿失败')
          if (!sameWorkspace(this, identity)) throw new Error('工作簿已变化，忽略旧保存响应')
          this.adoptSavedDraft(res.data, snapshot)
          return res
        })
      }).then(res => {
        if (!res || res.code !== 200) throw new Error((res && res.msg) || '保存分支草稿失败')
        return res
      }).catch(e => {
        if (strict) throw e
        console.warn('自动保存版本草稿失败:', e)
      })
    },

    openVersionDialog() {
      if (this.agentLocked) return this.$modal.msgWarning('请先停止或完成智能体任务')
      if (!this.currentDocumentId) return
      this.versionDialogVisible = true
      const identity = this.workspaceIdentity()
      return Promise.all([this.loadVersionStatus(), listCommits(identity.workbookId).then(res => {
        if (!sameWorkspace(this, identity)) return
        if (res && res.code === 200) this.commitHistory = res.data || []
      }), this.loadBranches()]).catch(e => { if (sameWorkspace(this, identity)) this.$modal.msgError('加载版本信息失败：' + (e.message || String(e))) })
    },

    loadBranches() {
      if (!this.currentDocumentId) return Promise.resolve()
      const identity = this.workspaceIdentity()
      return listBranches(identity.workbookId).then(res => {
        if (!sameWorkspace(this, identity)) return
        if (!res || res.code !== 200) throw new Error((res && res.msg) || '加载分支失败')
        this.branches = res.data || []
      })
    },

    createDraftBranch() {
      if (this.agentLocked || this.agentSyncRequired || this.persistBusy || this.historyPreview) return this.$modal.msgWarning('请先结束任务、保存或预览再创建分支')
      const identity = this.workspaceIdentity()
      const name = (this.newBranchName || '').trim()
      if (!name) return this.$modal.msgWarning('请输入草稿分支名称')
      return createBranch(identity.workbookId, {
        branchName: name,
        fromCommitId: this.headCommitId,
        operationId: createOperationId()
      }).then(res => {
        if (!sameWorkspace(this, identity)) return
        if (!res || res.code !== 200) throw new Error((res && res.msg) || '创建分支失败')
        if ((this.newBranchName || '').trim() === name) this.newBranchName = ''
        this.$modal.msgSuccess('草稿分支已创建')
        return this.loadBranches()
      }).catch(e => { if (sameWorkspace(this, identity)) this.$modal.msgError(e.message || String(e)) })
    },

    checkoutSelectedBranch(branchId) {
      const identity = this.workspaceIdentity()
      if (this.agentLocked || this.agentSyncRequired || this.persistBusy || this.historyPreview) {
        this.selectedBranchId = this.currentBranchId
        return this.$modal.msgWarning('请先结束任务、保存、预览或核实状态再切换分支')
      }
      if (!branchId || String(branchId) === String(this.currentBranchId)) return
      const target = this.branches.find(item => String(item.branchId) === String(branchId))
      if (!target || !Number.isInteger(target.version) || !Number.isInteger(target.draftVersion) || !target.loadedRevision || !Number.isInteger(this.currentDocumentVersion)) {
        this.selectedBranchId = this.currentBranchId
        return this.$modal.msgWarning('请刷新并核实源、目标分支信息后再切换')
      }
      this.flushManualChanges(true)
      const fingerprint = JSON.stringify(this.getSnapshot()), targetBinding = { targetVersion: target.version,
        targetHeadCommitId: target.headCommitId, expectedTargetDraftVersion: target.draftVersion, targetLoadedRevision: target.loadedRevision }
      let source
      try { source = versionMutationBinding(this) } catch (e) { this.selectedBranchId = this.currentBranchId; return this.$modal.msgError(e.message) }
      const workbookVersion = this.currentDocumentVersion
      let sent = false
      const verify = () => {
        if (!sameWorkspace(this, identity) || this.agentLocked || this.agentSyncRequired || this.historyPreview || fingerprint !== JSON.stringify(this.getSnapshot())) throw new Error('切换期间工作区或编辑已变化，未覆盖页面')
      }
      return this.runPersistTask('正在切换分支…', async () => {
        verify()
        if (JSON.stringify(source) !== JSON.stringify(versionMutationBinding(this)) || workbookVersion !== this.currentDocumentVersion) throw new Error('源版本已变化，未发送切换请求')
        this.preserveLocalDraft(true)
        if (this.pendingChanges.length) {
          const snapshot = this.getSnapshot(), saved = await saveWorkingState(identity.workbookId, workingStatePayload(this, snapshot))
          if (!saved || saved.code !== 200 || !saved.data) throw new Error((saved && saved.msg) || '源分支草稿保存失败')
          if (!sameWorkspace(this, identity)) throw new Error('源草稿保存响应属于旧工作区')
          this.adoptSavedDraft(saved.data, snapshot)
        }
        verify()
        sent = true
        const loaded = versionMutationBinding(this)
        const res = await checkoutBranch(identity.workbookId, branchId, { ...targetBinding, operationId: createOperationId(),
          currentBranchId: loaded.branchId, baseCommitId: loaded.baseCommitId, baseVersion: loaded.baseVersion,
          expectedDraftVersion: loaded.expectedDraftVersion, loadedRevision: loaded.loadedRevision, expectedWorkbookVersion: workbookVersion })
        if (res && res.code !== 200) throw new Error(res.msg || '分支切换被拒绝')
        if (!res || !res.data || !res.data.content || !res.data.loadedRevision || String(res.data.branchId) !== String(branchId) || !Number.isInteger(res.data.version) || !Number.isInteger(res.data.workbookVersion) || (res.data.workingState && !Number.isInteger(res.data.workingState.version))) throw new Error('分支切换响应不完整，请核实服务器')
        verify()
        if (!this.univer || !this.univerAPI) throw new Error('表格引擎尚未就绪，请核实服务器分支')
        this.loadUniverWorkbook(JSON.parse(res.data.content))
        this.applyVersionStatus(res.data, true)
        this.loadedSnapshotFingerprint = JSON.stringify(this.getSnapshot())
        this.workspaceGeneration++
        this.statusMessage = '已切换到分支：' + res.data.branchName
        const switchedIdentity = this.workspaceIdentity()
        try { await this.loadBranches() } catch (e) { if (sameWorkspace(this, switchedIdentity)) this.$modal.msgWarning('切换成功，分支列表刷新失败，请刷新版本管理后再切换') }
        return res
      }).catch(e => {
        if (!sameWorkspace(this, identity)) return
        this.selectedBranchId = this.currentBranchId
        if (sent) {
          this.branchCheckoutUncertain = true
          this.agentSyncRequired = true
          this.lockAgentWorkspace(true)
          try { this.preserveLocalDraft(true) } catch (backupError) { this.$modal.msgError(backupError.message) }
        }
        this.$modal.msgError('切换失败：' + (e.message || String(e)))
      })
    },

    reconcileBranchCheckout() {
      if (!this.branchCheckoutUncertain || this.persistBusy) return
      const identity = this.workspaceIdentity(), fingerprint = JSON.stringify(this.getSnapshot())
      const verify = () => {
        if (!sameWorkspace(this, identity) || fingerprint !== JSON.stringify(this.getSnapshot())) throw new Error('核实期间工作区或编辑已变化，已保持原页面')
      }
      const checkRun = async () => {
        const run = await activeAgent(identity.workbookId)
        if (!run || typeof run !== 'object' || Array.isArray(run)) throw new Error('无法核实任务状态，保持保护')
        if (run.runId && !['COMPLETED', 'FAILED', 'CANCELLED'].includes(run.status)) throw new Error('请先停止或完成当前任务，再核实分支')
      }
      return this.runPersistTask('正在核实服务器分支…', async () => {
        this.preserveLocalDraft(true)
        await checkRun()
        const response = await getVersionStatus(identity.workbookId)
        verify()
        const state = response && response.code === 200 && response.data
        if (!state || String(state.workbookId) !== String(identity.workbookId) || !state.branchId || !state.content || !state.loadedRevision || !Number.isInteger(state.version) || !Number.isInteger(state.workbookVersion) || (state.workingState && !Number.isInteger(state.workingState.version))) throw new Error('服务器分支信息不完整，保持保护')
        await checkRun()
        verify()
        if (!this.univer || !this.univerAPI) throw new Error('表格引擎尚未就绪，保持保护')
        this.loadUniverWorkbook(JSON.parse(state.content))
        this.applyVersionStatus(state, true)
        this.loadedSnapshotFingerprint = JSON.stringify(this.getSnapshot())
        this.workspaceGeneration++
        this.branchCheckoutUncertain = false
        this.agentSyncRequired = false
        this.lockAgentWorkspace(false)
        this.statusMessage = '已核实服务器当前分支：' + state.branchName
        const reconciledIdentity = this.workspaceIdentity()
        try { await this.loadBranches() } catch (e) { if (sameWorkspace(this, reconciledIdentity)) this.$modal.msgWarning('核实成功，分支列表刷新失败，请刷新版本管理') }
        return state
      }).catch(e => { if (sameWorkspace(this, identity)) this.$modal.msgError('核实失败：' + (e.message || String(e))) })
    },

    discardCurrentDraft() {
      if (this.agentLocked || this.persistBusy || this.historyPreview) return this.$modal.msgWarning('请先结束当前任务或预览，再核实草稿')
      let binding
      try { binding = versionMutationBinding(this) } catch (e) { return this.$modal.msgError(e.message) }
      if (binding.expectedDraftVersion < 0) return this.$modal.msgWarning('当前页面没有已加载的服务器草稿')
      const identity = this.workspaceIdentity(), fingerprint = JSON.stringify(this.getSnapshot())
      return this.$confirm('将丢弃当前分支中本页面已加载的服务器草稿，正式历史保留。页面副本会先保留到可下载草稿，再载入服务器当前内容。是否继续？', '丢弃已保存草稿', {
        confirmButtonText: '确认丢弃', cancelButtonText: '取消', type: 'warning'
      }).then(() => {
        if (!sameWorkspace(this, identity) || this.agentLocked || fingerprint !== JSON.stringify(this.getSnapshot()) || JSON.stringify(binding) !== JSON.stringify(versionMutationBinding(this))) throw new Error('确认期间工作区已变化，未发送丢弃请求，请重新核实')
        return this.runPersistTask('正在丢弃已核实草稿…', async () => {
          this.preserveLocalDraft(true)
          const res = await discardWorkingState(identity.workbookId, { ...binding, confirmed: true })
          if (!res || res.code !== 200) throw new Error((res && res.msg) || '丢弃失败，本地副本已保留')
          if (!sameWorkspace(this, identity)) throw new Error('丢弃结果属于旧工作区，请核实服务器状态')
          const state = await getVersionStatus(identity.workbookId)
          if (!sameWorkspace(this, identity)) throw new Error('工作区已切换，忽略旧核实响应')
          if (!state || state.code !== 200 || !state.data || String(state.data.branchId) !== String(identity.branchId)) throw new Error('服务器分支已变化，请重新载入并核实')
          if (fingerprint !== JSON.stringify(this.getSnapshot())) { this.preserveLocalDraft(true); throw new Error('等待期间有新编辑，已保留，本次不覆盖页面') }
          this.loadUniverWorkbook(JSON.parse(state.data.content))
          this.applyVersionStatus(state.data, true)
          this.loadedSnapshotFingerprint = JSON.stringify(this.getSnapshot())
          this.workspaceGeneration++
          return state
        })
      }).then(() => { this.$modal.msgSuccess('已丢弃指定草稿，页面副本已保留'); return this.openVersionDialog() }).catch(e => {
        if (e !== 'cancel' && e !== 'close') this.$modal.msgError(e.message || String(e))
      })
    },

    fastForwardCurrentBranch() {
      if (this.agentLocked || this.persistBusy || this.historyPreview) return this.$modal.msgWarning('请先结束任务或预览，再核实分支')
      this.flushManualChanges(true)
      if (this.pendingChanges.length) return this.$modal.msgWarning('源分支还有未提交变化，请先提交')
      const main = this.branches.find(item => item.branchName === 'main')
      if (!main || !Number.isInteger(main.draftVersion) || !main.loadedRevision) return this.$modal.msgWarning('请刷新并核实 main 分支信息后再合入')
      let source
      try { source = versionMutationBinding(this) } catch (e) { return this.$modal.msgError(e.message) }
      const identity = this.workspaceIdentity(), fingerprint = JSON.stringify(this.getSnapshot())
      const request = { operationId: createOperationId(), sourceHeadCommitId: source.baseCommitId, sourceVersion: source.baseVersion,
        mainHeadCommitId: main.headCommitId, mainVersion: main.version, expectedMainDraftVersion: main.draftVersion, mainLoadedRevision: main.loadedRevision }
      return this.$confirm('仅当 main 未分叉、源分支无未提交草稿且目标 main 无未提交变化时才会合入。任何版本变化都会停止。是否继续？', '安全纳入 main', {
        confirmButtonText: '继续', cancelButtonText: '取消', type: 'warning'
      }).then(() => {
        if (!sameWorkspace(this, identity) || this.agentLocked || fingerprint !== JSON.stringify(this.getSnapshot()) || JSON.stringify(source) !== JSON.stringify(versionMutationBinding(this))) throw new Error('确认期间工作区已变化，未发送合入请求')
        return this.runPersistTask('正在安全纳入 main…', async () => {
          const res = await fastForwardBranchToMain(identity.workbookId, identity.branchId, request)
          if (!res || res.code !== 200 || !res.data) throw new Error((res && res.msg) || '合入失败')
          if (!sameWorkspace(this, identity)) throw new Error('合入响应属于旧工作区，未覆盖当前页面')
          if (fingerprint !== JSON.stringify(this.getSnapshot())) { this.preserveLocalDraft(true); throw new Error('等待期间有新编辑，已保留，请重新核实服务器') }
          this.historyPreview = false
          this.loadUniverWorkbook(JSON.parse(res.data.content))
          this.applyVersionStatus({ ...res.data, loadedRevision: res.data.loadedRevision || res.data.contentHash, workingState: null }, true)
          this.loadedSnapshotFingerprint = JSON.stringify(this.getSnapshot())
          this.workspaceGeneration++
          return res
        })
      }).then(() => { this.$modal.msgSuccess('源提交已安全纳入 main'); return this.openVersionDialog() }).catch(e => {
        if (e !== 'cancel' && e !== 'close') this.$modal.msgError(e.message || String(e))
      })
    },

    commitPendingChanges() {
      if (this.agentLocked) return this.$modal.msgWarning('请先停止或完成智能体任务')
      if (this.persistBusy) return
      this.flushManualChanges(true)
      const message = (this.commitMessage || '').trim()
      if (!message) {
        this.$modal.msgWarning('请输入提交说明')
        return
      }
      this.commitLoading = true
      this.runPersistTask('正在提交版本…', updateProgress => {
        updateProgress(12, '正在生成提交快照…')
        const snapshot = this.getSnapshot()
        if (!snapshot) throw new Error('无法获取当前工作簿快照')
        updateProgress(35, '正在整理版本数据…')
        const identity = this.workspaceIdentity()
        const submittedChanges = this.pendingChanges.length
        const submittedActions = this.pendingActions.length
        const payload = {
          ...versionMutationBinding(this),
          commitMessage: message,
          operationId: createOperationId(),
          content: JSON.stringify(snapshot),
          source: this.pendingChanges.some(item => item.source === 'ai') && this.pendingChanges.some(item => item.source === 'manual')
            ? 'mixed' : (this.pendingChanges.some(item => item.source === 'ai') ? 'ai' : 'manual'),
          actionSchemaVersion: '2.0',
          plannedActions: this.pendingActions,
          executedChanges: this.pendingChanges,
          diffSummary: { changeCount: this.pendingChanges.length }
        }
        updateProgress(48, '正在上传并创建版本…')
        return createCommit(this.currentDocumentId, payload, percent => {
          updateProgress(48 + Math.round(percent * 0.42), percent >= 100 ? '正在等待服务器创建版本…' : '正在上传并创建版本…')
        }).then(res => {
          if (!res || res.code !== 200) throw new Error((res && res.msg) || '提交失败')
          if (!sameWorkspace(this, identity)) throw new Error('提交响应已过期')
          this.headCommitId = res.data.commitId
          this.branchVersion = res.data.branchVersion === undefined ? payload.baseVersion + 1 : res.data.branchVersion
          this.currentDocumentVersion = res.data.workbookVersion === undefined ? this.currentDocumentVersion + 1 : res.data.workbookVersion
          this.adoptSavedDraft({ version: -1, loadedRevision: res.data.afterHash }, snapshot)
          this.pendingChanges = this.pendingChanges.slice(submittedChanges)
          this.pendingActions = this.pendingActions.slice(submittedActions)
          return res
        })
      }).then(res => {
        if (!res || res.code !== 200) throw new Error((res && res.msg) || '提交失败')
        this.$modal.msgSuccess('提交成功')
        this.commitMessage = ''
        return this.openVersionDialog()
      }).catch(e => this.$modal.msgError(e.message || String(e))).finally(() => {
        this.commitLoading = false
      })
    },

    previewVersion(row) {
      if (!row || !row.commitId) return
      const identity = this.workspaceIdentity()
      getVersionSnapshot(this.currentDocumentId, row.commitId).then(res => {
        if (!sameWorkspace(this, identity)) throw new Error('历史预览响应已过期')
        if (!res || res.code !== 200 || !res.data || !res.data.content) throw new Error((res && res.msg) || '快照不存在')
        const snapshot = JSON.parse(res.data.content)
        this.$confirm('历史版本将以只读预览方式加载。当前未提交草稿已经自动保存，是否继续？', '版本预览', {
          confirmButtonText: '继续', cancelButtonText: '取消', type: 'warning'
        }).then(() => this.autosaveWorkingState(true)).then(() => {
          if (!sameWorkspace(this, identity)) throw new Error('工作簿已切换，取消历史预览')
          this.preserveLocalDraft()
          this.historyPreview = true
          this.loadUniverWorkbook(snapshot)
          this.statusMessage = '正在预览版本 #' + row.commitNo + '（请重新打开当前文档返回编辑状态）'
          this.versionDialogVisible = false
        }).catch(() => {})
      }).catch(e => this.$modal.msgError('预览失败：' + (e.message || String(e))))
    },

    restoreToVersion(row) {
      if (this.agentLocked) return this.$modal.msgWarning('工作区待核实，请先结束当前任务')
      if (!row || !row.commitId) return
      const identity = this.workspaceIdentity()
      this.flushManualChanges()
      if (this.pendingChanges.length) {
        this.$modal.msgWarning('当前分支还有未提交变化，请先提交后再恢复历史版本')
        return
      }
      this.$confirm('恢复会把整本工作簿替换为版本 #' + row.commitNo + '，并新增一条恢复提交；原历史不会删除。是否继续？', '恢复历史版本', {
        confirmButtonText: '确认恢复', cancelButtonText: '取消', type: 'warning'
      }).then(() => {
        if (!sameWorkspace(this, identity)) throw new Error('工作簿已切换，取消恢复')
        return restoreVersion(identity.workbookId, row.commitId, {
          ...versionMutationBinding(this),
          operationId: createOperationId(),
          commitMessage: '恢复到版本 #' + row.commitNo,
          actionSchemaVersion: '2.0'
        })
      }).then(res => {
        if (!sameWorkspace(this, identity)) throw new Error('历史恢复响应已过期')
        if (!res || res.code !== 200) throw new Error((res && res.msg) || '恢复失败')
        this.$modal.msgSuccess('恢复成功，已生成新的提交')
        this.versionDialogVisible = false
        return getVersionStatus(this.currentDocumentId)
      }).then(res => {
        if (res && res.code === 200 && res.data) {
          if (!sameWorkspace(this, identity)) throw new Error('工作簿已切换，忽略恢复响应')
          this.preserveLocalDraft()
          this.historyPreview = false
          this.loadUniverWorkbook(JSON.parse(res.data.content))
          this.applyVersionStatus(res.data, true)
          this.loadedSnapshotFingerprint = JSON.stringify(this.getSnapshot())
        }
      }).catch(e => {
        if (e !== 'cancel' && e !== 'close') this.$modal.msgError('恢复失败：' + (e.message || String(e)))
      })
    }
  }
}
</script>

<style lang="scss" scoped>
.content-row {
  display: flex;
  gap: 10px;
  align-items: stretch;
}

.univer-container {
  flex: 1;
  min-width: 0;
  border: 1px solid #dcdfe6;
  border-radius: 4px;
  overflow: hidden;
  background: #fff;
}

.univer-wrapper {
  width: 100%;
  height: calc(100vh - 320px);
  min-height: 500px;
}

.persist-mask {
  position: fixed;
  inset: 0;
  z-index: 4000;
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgba(16, 24, 40, 0.46);
  backdrop-filter: blur(2px);
  cursor: wait;
}

.persist-card {
  width: 420px;
  max-width: calc(100vw - 48px);
  padding: 30px 34px 26px;
  border-radius: 12px;
  background: #fff;
  box-shadow: 0 16px 48px rgba(0, 0, 0, 0.22);
  text-align: center;
}

.persist-card__icon {
  margin-bottom: 12px;
  color: #409eff;
  font-size: 34px;
}

.persist-card__title {
  margin-bottom: 20px;
  color: #303133;
  font-size: 17px;
  font-weight: 600;
}

.persist-card__hint {
  margin-top: 14px;
  color: #909399;
  font-size: 12px;
}

.persist-fade-enter-active,
.persist-fade-leave-active {
  transition: opacity 0.18s ease;
}

.persist-fade-enter,
.persist-fade-leave-to {
  opacity: 0;
}

.mb8 {
  margin-bottom: 8px;
}
</style>
