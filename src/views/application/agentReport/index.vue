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

    <div class="content-row">
      <document-list
        :documents="documents"
        :active-id="currentDocumentId"
        @select="loadDocument"
        @delete="deleteDocument"
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
      :visible.sync="aiPanelVisible"
      :use-mock.sync="aiUseMock"
      :allow-mock-fallback="true"
      @executed="onAiExecuted"
    />
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
import { listWorkbooks, getWorkbook, saveWorkbook, deleteWorkbook } from '@/api/ai/workbook'

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
      currentDocumentType: ''
    }
  },
  mounted() {
    this.$nextTick(() => {
      this.initUniver()
    })
    this.fetchDocuments()
  },
  beforeDestroy() {
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
        this.statusMessage = 'Univer 表格引擎初始化成功，请导入Excel文件或新建表格'
      } catch (error) {
        console.error('Univer 初始化失败:', error)
        this.statusMessage = 'Univer 初始化失败: ' + (error.message || String(error))
      }
    },

    /** 销毁 Univer 实例 */
    disposeUniver() {
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
        return
      }
      const snapshot = this.getSnapshot()
      if (!snapshot) {
        if (showMsg) this.$modal.msgError('获取表格快照失败')
        return
      }
      const payload = {
        workbookId: this.currentDocumentId || null,
        name: this.currentDocumentName || snapshot.name || '未命名文档',
        type: this.currentDocumentType || 'created',
        content: JSON.stringify(snapshot)
      }
      saveWorkbook(payload).then((res) => {
        if (res && res.code === 200 && res.data) {
          this.currentDocumentId = res.data.workbookId
          this.currentDocumentName = res.data.name
          this.currentDocumentType = res.data.type
          this.statusMessage = '已保存：' + res.data.name
          if (showMsg) this.$modal.msgSuccess('保存成功')
          this.fetchDocuments()
        } else {
          if (showMsg) this.$modal.msgError((res && res.msg) || '保存失败')
        }
      }).catch((e) => {
        console.error('保存文档失败:', e)
        if (showMsg) this.$modal.msgError('保存失败：' + (e && e.message ? e.message : String(e)))
      })
    },

    /** 打开文档 */
    loadDocument(doc) {
      if (!doc || !doc.workbookId) return
      getWorkbook(doc.workbookId).then((res) => {
        if (res && res.code === 200 && res.data) {
          const wb = res.data
          if (!wb.content) {
            this.$modal.msgWarning('文档内容为空')
            return
          }
          let snapshot = null
          try {
            snapshot = JSON.parse(wb.content)
          } catch (e) {
            console.error('文档内容解析失败:', e)
            this.$modal.msgError('文档内容解析失败')
            return
          }
          this.loadUniverWorkbook(snapshot)
          this.currentDocumentId = wb.workbookId
          this.currentDocumentName = wb.name
          this.currentDocumentType = wb.type
          this.statusMessage = '已打开：' + wb.name
        } else {
          this.$modal.msgError((res && res.msg) || '加载文档失败')
        }
      }).catch((e) => {
        console.error('加载文档失败:', e)
        this.$modal.msgError('加载文档失败')
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
          self.currentDocumentName = univerData.name || file.name
          self.currentDocumentType = 'imported'
          self.statusMessage = 'Excel 文件导入成功'
          self.$modal.msgSuccess('导入成功')
          // 导入后自动保存
          self.saveCurrentWorkbook(false)
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
              self.currentDocumentName = univerData.name || file.name
              self.currentDocumentType = 'imported'
              self.statusMessage = 'Excel 文件导入成功（备用方案）'
              self.$modal.msgSuccess('导入成功')
              // 导入后自动保存
              self.saveCurrentWorkbook(false)
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

        if (ws['!ref']) {
          const range = XLSX.utils.decode_range(ws['!ref'])
          for (let r = range.s.r; r <= range.e.r; r++) {
            for (let c = range.s.c; c <= range.e.c; c++) {
              const cellAddr = XLSX.utils.encode_cell({ r: r, c: c })
              const cell = ws[cellAddr]
              if (cell) {
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
              }
            }
          }
        }

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
      }
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

      const limitRow = Math.min(maxRow, 10000)
      const limitCol = Math.min(maxCol, 1000)

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
        name: '新建工作簿',
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

.mb8 {
  margin-bottom: 8px;
}
</style>
