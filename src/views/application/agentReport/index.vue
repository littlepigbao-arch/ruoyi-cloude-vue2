<template>
  <div class="app-container">
    <el-row :gutter="10" class="mb8">
      <el-col :span="1.5">
        <el-upload
          class="upload-demo"
          action="#"
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
        >导出Excel</el-button>
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
        <el-form-item label="切换分支">
          <el-select v-model="selectedBranchId" style="width:190px" @change="checkoutSelectedBranch">
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
import {
  listWorkbooks, getWorkbook, saveWorkbook, deleteWorkbook, renameWorkbook,
  getVersionStatus, saveWorkingState, createCommit, listCommits, getVersionSnapshot, restoreVersion,
  listBranches, createBranch, checkoutBranch, fastForwardBranchToMain
} from '@/api/ai/workbook'
import { createOperationId, buildManualChanges, buildLightweightManualChanges } from './change-capture'

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
  mounted() {
    this.$nextTick(() => {
      this.initUniver()
    })
    this.fetchDocuments()
  },
  beforeDestroy() {
    clearTimeout(this.autosaveTimer)
    this.autosaveTimer = null
    clearInterval(this.persistProgressTimer)
    this.persistProgressTimer = null
    this.disposeUniver()
  },
  methods: {
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
        if (this.historyPreview && event && event.type === 0 && String(event.id || '').indexOf('sheet.') === 0) {
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
        !this.suspendChangeCapture && !this.aiExecutionInProgress && !this.historyPreview &&
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
      return this.runPersistTask('正在保存表格…', updateProgress => {
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
          return res
        })
      }).then((res) => {
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
        if (showMsg) this.$modal.msgError('保存失败：' + (e && e.message ? e.message : String(e)))
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
      if (!doc || !doc.workbookId || this.persistBusy) return
      if (String(doc.workbookId) === String(this.currentDocumentId) && !this.historyPreview) return
      const targetWorkbookId = doc.workbookId
      this.runPersistTask('正在切换文件…', updateProgress => {
        // 切换前先收口当前编辑，避免大数据模式尚未触发的延迟保存丢失。
        this.flushManualChanges(true)
        let saveCurrent = Promise.resolve()
        if (this.currentDocumentId && this.currentBranchId && this.pendingChanges.length) {
          const previousWorkbookId = this.currentDocumentId
          updateProgress(8, '正在保存当前文件的修改…')
          const snapshot = this.getSnapshot()
          if (!snapshot) throw new Error('无法保存当前文件的修改')
          const payload = {
            branchId: this.currentBranchId,
            baseCommitId: this.headCommitId,
            baseVersion: this.branchVersion,
            content: JSON.stringify(snapshot),
            plannedActions: this.pendingActions,
            executedChanges: this.pendingChanges
          }
          saveCurrent = saveWorkingState(previousWorkbookId, payload, percent => {
            updateProgress(8 + Math.round(percent * 0.2), '正在保存当前文件的修改…')
          }).then(res => {
            if (!res || res.code !== 200) throw new Error((res && res.msg) || '当前文件保存失败')
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
          const workbookRes = results[0]
          const statusRes = results[1]
          if (!workbookRes || workbookRes.code !== 200 || !workbookRes.data) {
            throw new Error((workbookRes && workbookRes.msg) || '加载文档失败')
          }
          const wb = workbookRes.data
          const status = statusRes && statusRes.code === 200 ? statusRes.data : null
          const draft = status && status.workingState
          const content = draft && draft.content ? draft.content : wb.content
          if (!content) throw new Error('文档内容为空')
          updateProgress(55, draft && draft.content ? '正在解析未提交草稿…' : '正在解析文件数据…')
          const snapshot = JSON.parse(content)
          updateProgress(68, '正在构建大数据表格…')
          this.currentDocumentId = wb.workbookId
          this.currentDocumentName = wb.name
          this.currentDocumentType = wb.type
          this.currentDocumentVersion = wb.version
          this.applyVersionStatus(status)
          this.loadUniverWorkbook(snapshot)
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

    /** 处理文件上传 */
    handleFileChange(file) {
      if (!file || !file.raw) return
      if (!this.beforeUpload(file.raw)) return

      this.statusMessage = '正在解析 Excel 文件...'
      this.importExcelWithXLSX(file.raw)
    },

    /** 使用 XLSX 库导入 Excel */
    importExcelWithXLSX(file) {
      const self = this
      const reader = new FileReader()
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result)
          const workbook = XLSX.read(data, { type: 'array' })
          const univerData = self.convertXlsxWorkbookToUniverData(workbook, file.name)
          self.loadUniverWorkbook(univerData)
          self.currentDocumentId = null
          self.currentDocumentName = self.uniqueName(univerData.name || file.name)
          self.currentDocumentType = 'imported'
          self.currentDocumentVersion = null
          self.resetVersionState()
          self.statusMessage = 'Excel 文件导入成功'
          self.$modal.msgSuccess('导入成功')
          // 大表先保证界面可操作，由用户确认后再做首次完整保存。
          if (!self.largeDataMode) self.saveCurrentWorkbook(false)
          else self.statusMessage = 'Excel 导入成功，已自动启用大数据模式；请点击“保存”完成首次保存'
        } catch (error) {
          console.error('XLSX 解析失败，尝试 LuckyExcel:', error)
          self.statusMessage = 'XLSX 解析失败，尝试备用方案...'
          self.importExcelWithLuckyExcel(file)
        }
      }
      reader.onerror = function() {
        self.statusMessage = '文件读取失败'
        self.$modal.msgError('文件读取失败')
      }
      reader.readAsArrayBuffer(file)
    },

    /** 使用 LuckyExcel 备用导入方案 */
    importExcelWithLuckyExcel(file) {
      const self = this
      try {
        LuckyExcel.transformExcelToLucky(
          file,
          function(exportJson) {
            if (!exportJson || !exportJson.sheets || exportJson.sheets.length === 0) {
              self.$modal.msgError('未能读取到 Excel 中的有效数据')
              self.statusMessage = '导入失败：文件中没有有效数据'
              return
            }
            try {
              const univerData = self.convertLuckyJsonToUniver(exportJson, file.name)
              self.loadUniverWorkbook(univerData)
              self.currentDocumentId = null
              self.currentDocumentName = self.uniqueName(univerData.name || file.name)
              self.currentDocumentType = 'imported'
              self.currentDocumentVersion = null
              self.resetVersionState()
              self.statusMessage = 'Excel 文件导入成功（备用方案）'
              self.$modal.msgSuccess('导入成功')
              if (!self.largeDataMode) self.saveCurrentWorkbook(false)
              else self.statusMessage = 'Excel 导入成功，已自动启用大数据模式；请点击“保存”完成首次保存'
            } catch (err) {
              console.error('LuckyExcel 数据转换失败:', err)
              self.statusMessage = '导入失败: ' + (err.message || String(err))
              self.$modal.msgError('导入失败: ' + (err.message || String(err)))
            }
          },
          function(error) {
            console.error('LuckyExcel 解析失败:', error)
            self.statusMessage = '导入失败: ' + (error.message || String(error))
            self.$modal.msgError('导入失败: ' + (error.message || String(error)))
          }
        )
      } catch (error) {
        console.error('LuckyExcel 导入异常:', error)
        this.statusMessage = '导入失败: ' + (error.message || String(error))
        this.$modal.msgError('导入失败: ' + (error.message || String(error)))
      }
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
                t: typeof rawValue === 'number' ? 1 : 2,
                f: v.f || undefined
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
                  t: typeof rawValue === 'number' ? 1 : 2,
                  f: cell.f || undefined
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
          let t = 2
          if (cell.t === 'n' || cell.t === 'b' || cell.t === 'd') t = 1
          cellData[r][c] = {
            v: cell.v !== undefined ? cell.v : (cell.w || ''),
            t: t,
            f: cell.f || undefined
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

    /** 导出 Excel */
    handleExport() {
      if (!this.hasData || !this.univerAPI) {
        this.$modal.msgWarning('没有可导出的数据')
        return
      }

      try {
        this.statusMessage = '正在导出 Excel 文件...'
        const workbook = this.univerAPI.getActiveWorkbook()
        if (!workbook) {
          this.$modal.msgError('获取工作簿失败')
          return
        }

        const snapshot = this.getSnapshot()

        if (!snapshot) {
          this.$modal.msgError('获取工作簿数据失败')
          return
        }

        const xlsxWb = this.convertUniverSnapshotToXlsx(snapshot)
        const wbout = XLSX.write(xlsxWb, { bookType: 'xlsx', type: 'array' })
        const blob = new Blob([wbout], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })
        const fileName = (snapshot.name || '表格导出') + '_' + this.formatDate(new Date()) + '.xlsx'
        saveAs(blob, fileName)
        this.statusMessage = '导出成功：' + fileName
        this.$modal.msgSuccess('导出成功')
      } catch (error) {
        console.error('导出失败:', error)
        this.statusMessage = '导出失败: ' + (error.message || String(error))
        this.$modal.msgError('导出失败: ' + (error.message || String(error)))
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

    /** 将 Univer snapshot 转换为 XLSX workbook */
    convertUniverSnapshotToXlsx(snapshot) {
      const wb = XLSX.utils.book_new()
      const sheetOrder = snapshot.sheetOrder || Object.keys(snapshot.sheets || {})

      sheetOrder.forEach(sheetId => {
        const sheet = snapshot.sheets[sheetId]
        if (!sheet) return
        const ws = this.convertUniverSheetToXlsx(sheet)
        XLSX.utils.book_append_sheet(wb, ws, (sheet.name || 'Sheet').substring(0, 31))
      })

      return wb
    },

    /** 将 Univer 工作表转换为 XLSX 工作表 */
    convertUniverSheetToXlsx(sheet) {
      const aoa = []
      const cellData = sheet.cellData || {}
      const mergeData = sheet.mergeData || []
      let maxRow = sheet.rowCount || 0
      let maxCol = sheet.columnCount || 0

      const rowKeys = Object.keys(cellData)
        .map(k => Number(k))
        .filter(n => !isNaN(n))

      if (rowKeys.length > 0) {
        const actualMaxRow = Math.max(...rowKeys)
        rowKeys.forEach(r => {
          const colKeys = Object.keys(cellData[r] || {})
            .map(k => Number(k))
            .filter(n => !isNaN(n))
          if (colKeys.length > 0) {
            maxCol = Math.max(maxCol, Math.max(...colKeys) + 1)
          }
        })
        maxRow = Math.max(maxRow, actualMaxRow + 1)
      }

      const limitRow = Math.min(maxRow, 100000)
      const limitCol = Math.min(maxCol, 100)

      for (let r = 0; r < limitRow; r++) {
        const row = []
        for (let c = 0; c < limitCol; c++) {
          if (cellData[r] && cellData[r][c] !== undefined) {
            const cell = cellData[r][c]
            const cellValue = cell.v
            const xlsxCell = {}
            if (cell.f) {
              xlsxCell.f = cell.f
              if (cellValue !== undefined && cellValue !== null && cellValue !== '') {
                xlsxCell.v = cellValue
              }
            } else {
              xlsxCell.v = cellValue
            }
            if (typeof cellValue === 'number') {
              xlsxCell.t = 'n'
            } else if (typeof cellValue === 'boolean') {
              xlsxCell.t = 'b'
            } else {
              xlsxCell.t = 's'
            }
            row.push(xlsxCell)
          } else {
            row.push(null)
          }
        }
        aoa.push(row)
      }

      const ws = XLSX.utils.aoa_to_sheet(aoa)

      if (mergeData && mergeData.length > 0) {
        ws['!merges'] = mergeData
          .filter(m =>
            m.startRow !== undefined &&
            m.endRow !== undefined &&
            m.startColumn !== undefined &&
            m.endColumn !== undefined
          )
          .map(m => ({
            s: { r: m.startRow, c: m.startColumn },
            e: { r: m.endRow, c: m.endColumn }
          }))
      }

      return ws
    },

    /** 创建新的空白表格 */
    createNewSheet() {
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
    applyVersionStatus(data) {
      if (!data) {
        this.resetVersionState()
        return null
      }
      this.currentBranchId = data.branchId
      this.currentBranchName = data.branchName || 'main'
      this.selectedBranchId = data.branchId
      this.branchVersion = data.version
      this.headCommitId = data.headCommitId
      if (data.workbookVersion !== undefined) this.currentDocumentVersion = data.workbookVersion
      const draft = data.workingState
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
      return getVersionStatus(this.currentDocumentId).then(res => {
        if (res && res.code === 200 && res.data) {
          const draft = this.applyVersionStatus(res.data)
          if (restoreDraft && draft && draft.content) {
            this.loadUniverWorkbook(JSON.parse(draft.content))
            this.statusMessage = '已恢复分支未提交草稿：' + (res.data.branchName || 'main')
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
      if (!this.currentDocumentId || !this.currentBranchId || !this.pendingChanges.length) return Promise.resolve()
      if (this.persistBusy) return strict ? Promise.reject(new Error('已有保存或提交任务正在进行')) : Promise.resolve()
      return this.runPersistTask(strict ? '正在保存草稿…' : '正在自动保存大数据表格…', updateProgress => {
        updateProgress(12, '正在生成表格快照…')
        const snapshot = this.getSnapshot()
        if (!snapshot) throw new Error('无法获取当前工作簿快照')
        updateProgress(35, '正在整理草稿数据…')
        const payload = {
          branchId: this.currentBranchId,
          baseCommitId: this.headCommitId,
          baseVersion: this.branchVersion,
          content: JSON.stringify(snapshot),
          plannedActions: this.pendingActions,
          executedChanges: this.pendingChanges
        }
        updateProgress(48, '正在上传草稿数据…')
        return saveWorkingState(this.currentDocumentId, payload, percent => {
          updateProgress(48 + Math.round(percent * 0.42), percent >= 100 ? '正在等待服务器保存草稿…' : '正在上传草稿数据…')
        }).then(res => {
          if (!res || res.code !== 200) throw new Error((res && res.msg) || '保存分支草稿失败')
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
      if (!this.currentDocumentId) return
      this.versionDialogVisible = true
      Promise.all([this.loadVersionStatus(), listCommits(this.currentDocumentId).then(res => {
        if (res && res.code === 200) this.commitHistory = res.data || []
      }), this.loadBranches()]).catch(e => this.$modal.msgError('加载版本信息失败：' + (e.message || String(e))))
    },

    loadBranches() {
      if (!this.currentDocumentId) return Promise.resolve()
      return listBranches(this.currentDocumentId).then(res => {
        if (!res || res.code !== 200) throw new Error((res && res.msg) || '加载分支失败')
        this.branches = res.data || []
      })
    },

    createDraftBranch() {
      const name = (this.newBranchName || '').trim()
      if (!name) return this.$modal.msgWarning('请输入草稿分支名称')
      createBranch(this.currentDocumentId, {
        branchName: name,
        fromCommitId: this.headCommitId,
        operationId: createOperationId()
      }).then(res => {
        if (!res || res.code !== 200) throw new Error((res && res.msg) || '创建分支失败')
        this.newBranchName = ''
        this.$modal.msgSuccess('草稿分支已创建')
        return this.loadBranches()
      }).catch(e => this.$modal.msgError(e.message || String(e)))
    },

    checkoutSelectedBranch(branchId) {
      if (!branchId || branchId === this.currentBranchId) return
      const previousBranchId = this.currentBranchId
      const previousSelection = previousBranchId
      this.flushManualChanges(true)
      this.autosaveWorkingState(true).then(() => checkoutBranch(this.currentDocumentId, branchId, {
        operationId: createOperationId(),
        currentBranchId: previousBranchId,
        baseCommitId: this.headCommitId
      })).then(res => {
        if (!res || res.code !== 200 || !res.data || !res.data.content) throw new Error((res && res.msg) || '切换分支失败')
        this.loadUniverWorkbook(JSON.parse(res.data.content))
        this.currentBranchId = res.data.branchId
        this.currentBranchName = res.data.branchName
        this.selectedBranchId = res.data.branchId
        this.branchVersion = res.data.version
        this.headCommitId = res.data.headCommitId
        this.currentDocumentVersion = res.data.workbookVersion
        const draft = res.data.workingState
        this.pendingChanges = draft && draft.executedChangesJson ? JSON.parse(draft.executedChangesJson) : []
        this.pendingActions = draft && draft.draftActionsJson ? JSON.parse(draft.draftActionsJson) : []
        this.statusMessage = '已切换到分支：' + res.data.branchName
      }).catch(e => {
        this.selectedBranchId = previousSelection
        this.$modal.msgError('切换失败：' + (e.message || String(e)))
      })
    },

    fastForwardCurrentBranch() {
      this.flushManualChanges()
      if (this.pendingChanges.length) return this.$modal.msgWarning('草稿分支还有未提交变化，请先提交')
      const main = this.branches.find(item => item.branchName === 'main')
      if (!main) return this.$modal.msgError('未找到 main 分支')
      this.$confirm('仅当 main 自草稿分支创建后没有分叉时才会执行；否则系统会停止并要求人工处理冲突。是否继续？', '安全纳入 main', {
        confirmButtonText: '继续', cancelButtonText: '取消', type: 'warning'
      }).then(() => fastForwardBranchToMain(this.currentDocumentId, this.currentBranchId, {
        operationId: createOperationId(),
        mainHeadCommitId: main.headCommitId,
        mainVersion: main.version
      })).then(res => {
        if (!res || res.code !== 200 || !res.data) throw new Error((res && res.msg) || '纳入 main 失败')
        this.historyPreview = false
        this.loadUniverWorkbook(JSON.parse(res.data.content))
        this.currentBranchId = res.data.branchId
        this.currentBranchName = 'main'
        this.selectedBranchId = res.data.branchId
        this.branchVersion = res.data.version
        this.headCommitId = res.data.headCommitId
        this.currentDocumentVersion = res.data.workbookVersion
        this.$modal.msgSuccess('草稿提交已安全纳入 main')
        return this.openVersionDialog()
      }).catch(e => {
        if (e !== 'cancel' && e !== 'close') this.$modal.msgError(e.message || String(e))
      })
    },

    commitPendingChanges() {
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
        const payload = {
          branchId: this.currentBranchId,
          commitMessage: message,
          operationId: createOperationId(),
          baseCommitId: this.headCommitId,
          baseVersion: this.branchVersion,
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
          return res
        })
      }).then(res => {
        if (!res || res.code !== 200) throw new Error((res && res.msg) || '提交失败')
        this.$modal.msgSuccess('提交成功')
        this.pendingChanges = []
        this.pendingActions = []
        this.commitMessage = ''
        return this.openVersionDialog()
      }).catch(e => this.$modal.msgError(e.message || String(e))).finally(() => {
        this.commitLoading = false
      })
    },

    previewVersion(row) {
      if (!row || !row.commitId) return
      getVersionSnapshot(this.currentDocumentId, row.commitId).then(res => {
        if (!res || res.code !== 200 || !res.data || !res.data.content) throw new Error((res && res.msg) || '快照不存在')
        const snapshot = JSON.parse(res.data.content)
        this.$confirm('历史版本将以只读预览方式加载。当前未提交草稿已经自动保存，是否继续？', '版本预览', {
          confirmButtonText: '继续', cancelButtonText: '取消', type: 'warning'
        }).then(() => this.autosaveWorkingState(true)).then(() => {
          this.historyPreview = true
          this.loadUniverWorkbook(snapshot)
          this.statusMessage = '正在预览版本 #' + row.commitNo + '（请重新打开当前文档返回编辑状态）'
          this.versionDialogVisible = false
        }).catch(() => {})
      }).catch(e => this.$modal.msgError('预览失败：' + (e.message || String(e))))
    },

    restoreToVersion(row) {
      if (!row || !row.commitId) return
      this.flushManualChanges()
      if (this.pendingChanges.length) {
        this.$modal.msgWarning('当前分支还有未提交变化，请先提交后再恢复历史版本')
        return
      }
      this.$confirm('恢复会把整本工作簿替换为版本 #' + row.commitNo + '，并新增一条恢复提交；原历史不会删除。是否继续？', '恢复历史版本', {
        confirmButtonText: '确认恢复', cancelButtonText: '取消', type: 'warning'
      }).then(() => restoreVersion(this.currentDocumentId, row.commitId, {
        branchId: this.currentBranchId,
        operationId: createOperationId(),
        baseCommitId: this.headCommitId,
        baseVersion: this.branchVersion,
        commitMessage: '恢复到版本 #' + row.commitNo,
        actionSchemaVersion: '2.0'
      })).then(res => {
        if (!res || res.code !== 200) throw new Error((res && res.msg) || '恢复失败')
        this.$modal.msgSuccess('恢复成功，已生成新的提交')
        this.versionDialogVisible = false
        return getWorkbook(this.currentDocumentId)
      }).then(res => {
        if (res && res.code === 200 && res.data) {
          this.historyPreview = false
          this.loadUniverWorkbook(JSON.parse(res.data.content))
          this.currentDocumentVersion = res.data.version
          this.pendingChanges = []
          this.pendingActions = []
          this.loadVersionStatus()
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
