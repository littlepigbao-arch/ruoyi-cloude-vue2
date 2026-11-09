<template>
  <div class="app-container c3-accept">
    <el-alert
      type="info"
      :closable="false"
      title="C3 导入链路验收专页（独立 harness，不影响主页面）"
      description="本页面向 /ai/agent/import/* 全链路：导入（幂等）→ 预览（保精度字面量）→ Univer 重放（显示/计算证据）→ 保存重载（页内 SDK + 真实后端保存/重载，逐格比对）。自动断言一览常显于顶部（连通自检 401/403 判鉴权失败）；所有网络/业务错误原样展示不吞；证据留证推荐「复制日志（存 G 盘）」——浏览器下载目录默认在 C 盘，项目与证据文件一律落 G 盘。"
    />

    <!-- 0. 断言汇总面板 -->
    <el-card class="panel" shadow="never">
      <div slot="header" class="panel-header">
        <span>断言汇总（自动判定）</span>
        <el-button size="mini" type="primary" plain icon="el-icon-download" @click="exportAcceptanceLog">导出验收日志</el-button>
        <el-button size="mini" type="primary" plain icon="el-icon-document-copy" @click="copyAcceptanceLog">复制日志（存 G 盘）</el-button>
      </div>
      <el-alert
        v-if="allAssertionsGreen"
        type="success"
        :closable="false"
        :title="greenBannerText"
        class="mb8"
      />
      <el-table :data="assertionRows" size="mini" border>
        <el-table-column prop="name" label="用例" min-width="260" />
        <el-table-column prop="result" label="结果" width="110" align="center">
          <template slot-scope="scope">
            <el-tag size="mini" :type="resultTagType(scope.row.result)">{{ scope.row.result }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column prop="evidence" label="证据摘要" min-width="420" show-overflow-tooltip />
      </el-table>
    </el-card>

    <!-- 1. 连接面板 -->
    <el-card class="panel" shadow="never">
      <div slot="header"><span>连接面板</span></div>
      <el-form inline size="mini" label-width="130px">
        <el-form-item label="VUE_APP_BASE_API">
          <span class="mono">{{ baseApi }}</span>
        </el-form-item>
        <el-form-item label="getToken() 有值">
          <el-tag size="mini" :type="tokenPresent ? 'success' : 'info'">{{ tokenPresent ? '是（已有登录 Cookie）' : '否' }}</el-tag>
        </el-form-item>
        <el-form-item label="Token（JWT）">
          <el-input
            v-model="tokenInput"
            style="width: 420px"
            placeholder="粘贴后端签发的 JWT；应用后写入登录 Cookie，供 getToken() 读取"
            clearable
          />
        </el-form-item>
        <el-form-item>
          <el-button size="mini" type="primary" :disabled="!tokenInput.trim()" @click="applyToken">应用 Token</el-button>
          <el-button size="mini" type="warning" plain @click="clearToken">清除（还原 Cookie）</el-button>
          <el-button size="mini" :loading="probeRunning" @click="runProbe">连通自检</el-button>
        </el-form-item>
      </el-form>
      <div class="hint">
        说明：「应用 Token」会把 Token 写入登录 Cookie（Admin-Token，<b>与主应用共享</b>），并留存于 sessionStorage（c3-accept 前缀键）供本页刷新后自动重放；
        import.js 的 Authorization 取自 getToken()，因此粘贴的 Token 必须进 Cookie 才会生效。「清除」还原进入本页前的原 Cookie 值（无原值则删除该 Cookie）。验收完成后请务必点击清除或重新登录。
      </div>
      <el-alert
        v-if="probeResult"
        :type="probeResult.kind === 'business' || probeResult.kind === 'unexpected-ok' ? 'success' : 'error'"
        :closable="false"
        class="mb8"
        :title="probeTitle"
      >
        <div class="mono small">{{ probeDetail }}</div>
      </el-alert>
    </el-card>

    <!-- 2. 导入面板 -->
    <el-card class="panel" shadow="never">
      <div slot="header"><span>导入面板（createImportJob / getImportJob / getImportJobByRequest，requestId 幂等）</span></div>
      <el-form inline size="mini" label-width="130px">
        <el-form-item label="源文件(.xlsx)">
          <input ref="fileInput" type="file" accept=".xlsx" style="display: none" @change="onFileChange">
          <el-button size="mini" @click="$refs.fileInput.click()">选择文件</el-button>
          <span class="mono file-name">{{ fileName || '未选择' }}</span>
        </el-form-item>
        <el-form-item label="requestId">
          <el-input v-model="requestId" style="width: 340px" class="mono" />
          <el-button size="mini" icon="el-icon-refresh" @click="requestId = genRequestId()">重生成</el-button>
        </el-form-item>
        <el-form-item label="预览行数 rows">
          <el-input-number v-model="rowsRequested" :min="1" :max="5000" size="mini" controls-position="right" style="width: 120px" />
        </el-form-item>
      </el-form>
      <el-form inline size="mini">
        <el-form-item>
          <el-button size="mini" type="primary" :loading="importBusy" :disabled="!file || !requestId" @click="createImport">创建导入任务</el-button>
          <el-button size="mini" :loading="importBusy" :disabled="!jobId" @click="refreshStatus">刷新状态</el-button>
          <el-button size="mini" :loading="importBusy" :disabled="!requestId" @click="queryByRequest">按 requestId 查询</el-button>
          <el-button size="mini" type="warning" plain :loading="importBusy" :disabled="!file || !jobId" @click="resendSameRequest">同 requestId 重发（幂等断言）</el-button>
          <el-button size="mini" type="danger" plain :loading="budgetBusy" @click="runBudgetRefusalCheck">预算拒绝自检（免选文件）</el-button>
        </el-form-item>
      </el-form>
      <el-alert v-if="importError" type="error" :closable="false" class="mb8" :title="'导入面板最近一次错误（原样）'">
        <div class="mono small">{{ pretty(importError) }}</div>
      </el-alert>
      <div v-if="budgetResult" class="hint mono small">
        预算拒绝自检：status={{ budgetResult.status }} · kind={{ budgetResult.kind }} · limit={{ budgetResult.limit }} · observed={{ budgetResult.observed }}
        · jobId={{ budgetResult.jobId || '-' }} · requestId={{ budgetResult.requestId }}
      </div>
      <el-collapse v-model="openImportJson">
        <el-collapse-item title="创建任务返回（原始 JSON）" name="create">
          <pre class="raw-json">{{ pretty(createResult) }}</pre>
        </el-collapse-item>
        <el-collapse-item title="刷新状态返回（原始 JSON）" name="status">
          <pre class="raw-json">{{ pretty(jobStatusResult) }}</pre>
        </el-collapse-item>
        <el-collapse-item title="按 requestId 查询返回（原始 JSON）" name="byRequest">
          <pre class="raw-json">{{ pretty(byRequestResult) }}</pre>
        </el-collapse-item>
        <el-collapse-item title="同 requestId 重发返回（原始 JSON）" name="resend">
          <pre class="raw-json">{{ pretty(resendResult) }}</pre>
        </el-collapse-item>
      </el-collapse>
    </el-card>

    <!-- 3. 预览面板 -->
    <el-card class="panel" shadow="never">
      <div slot="header"><span>预览面板（previewImport：v 原样全精度显示）</span></div>
      <el-form inline size="mini">
        <el-form-item>
          <el-button size="mini" type="primary" :loading="previewBusy" :disabled="!jobId || !sourceHash" @click="queryPreview">查询预览</el-button>
          <span class="mono small">jobId={{ jobId || '-' }} · sourceHash={{ sourceHash ? shortHash(sourceHash) : '-' }} · rows={{ rowsRequested }}</span>
        </el-form-item>
      </el-form>
      <el-alert v-if="previewError" type="error" :closable="false" class="mb8" :title="'预览面板最近一次错误（原样）'">
        <div class="mono small">{{ pretty(previewError) }}</div>
      </el-alert>
      <el-table v-if="previewMetaRows.length" :data="previewMetaRows" size="mini" border class="mb8">
        <el-table-column prop="key" label="元信息" width="220" />
        <el-table-column prop="value" label="值">
          <template slot-scope="scope">
            <span class="mono">{{ scope.row.value }}</span>
          </template>
        </el-table-column>
      </el-table>
      <el-table v-if="previewCellRows.length" :data="previewCellRows" size="mini" border max-height="360">
        <el-table-column prop="sheetId" label="sheetId" width="140" show-overflow-tooltip />
        <el-table-column prop="row" label="rowIndex" width="90" align="center" />
        <el-table-column prop="col" label="列 key" width="80" align="center" />
        <el-table-column prop="t" label="t" width="60" align="center" />
        <el-table-column label="v（原样，字符串加引号前缀；数字 String(v) 全精度）" min-width="320">
          <template slot-scope="scope">
            <span
              :class="['mono', scope.row.precisionLiteral ? 'precision-literal' : '']"
              :title="showScalar(scope.row.cell.v)"
            >{{ showScalar(scope.row.cell.v) }}</span>
            <el-tag v-if="scope.row.hasF" size="mini" type="info" class="ml8">f={{ scope.row.cell.f }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column label="标记" width="140">
          <template slot-scope="scope">
            <el-tag v-if="scope.row.precisionLiteral" size="mini" type="warning">保精度字面量</el-tag>
            <el-tag v-else-if="scope.row.hasF" size="mini" type="success">公式缓存值</el-tag>
          </template>
        </el-table-column>
      </el-table>
      <div v-if="preview" class="hint">
        展示口径：typeof v==='string' 的单元格加引号前缀显示；数字用 String(v) 全精度显示（禁止 toFixed/截断）；橙色边框 = t=2 且 typeof v==='string' 的保精度字面量（import.js parsePreservingNumbers 停放的原始字面量）。<br>
        字段契约：创建/查询（createImportJob/getImportJob/getImportJobByRequest）响应字段名为 sourceContentHash，预览响应字段名为 sourceHash——两者是同一个源内容哈希；「查询预览」以查询参数 sourceHash（=sourceContentHash）发起。
      </div>
    </el-card>

    <!-- 4. Univer 重放面板 -->
    <el-card class="panel" shadow="never">
      <div slot="header"><span>Univer 重放面板（显示/计算证据，口径同 c3-import-types-regression.cjs）</span></div>
      <el-form inline size="mini">
        <el-form-item>
          <el-button size="mini" type="primary" :loading="replayBusy" :disabled="!hasPreviewRows" @click="replayToWorkbook">重放到工作簿</el-button>
          <el-button size="mini" plain @click="disposeUniver">销毁 Univer 实例</el-button>
        </el-form-item>
      </el-form>
      <el-alert v-if="replayError" type="error" :closable="false" class="mb8" :title="'重放面板最近一次错误（原样）'">
        <div class="mono small">{{ pretty(replayError) }}</div>
      </el-alert>
      <div ref="univerHost" class="univer-host" />
      <template v-if="readbackRows.length">
        <div class="section-title">显示证据：三点对照（预览 v / 存储快照 save().cellData / 显示通道 getValues）。判定口径（2026-11-08 实证）：数值格以存储层与预览值全等为硬判据；显示通道 15 位截断属 SDK 显示行为，单独标注不判 FAIL（保精度字面量按 F2 契约走文本通道 t=1 重放，读回必须仍是逐字符相等的字符串）</div>
        <el-table :data="readbackRows" size="mini" border max-height="320">
          <el-table-column prop="sheetId" label="sheetId" width="130" show-overflow-tooltip />
          <el-table-column prop="row" label="行" width="60" align="center" />
          <el-table-column prop="col" label="列" width="60" align="center" />
          <el-table-column prop="t" label="t" width="55" align="center" />
          <el-table-column label="预览 v" min-width="170">
            <template slot-scope="scope">
              <span :class="['mono', scope.row.precisionLiteral ? 'precision-literal' : '']" :title="showScalar(scope.row.v)">{{ showScalar(scope.row.v) }}</span>
            </template>
          </el-table-column>
          <el-table-column label="存储快照（save）" min-width="170">
            <template slot-scope="scope">
              <span class="mono" :title="showScalar(scope.row.saveSnapshot && scope.row.saveSnapshot.v)">{{ scope.row.saveSnapshot ? showScalar(scope.row.saveSnapshot.v) : '(无)' }}</span>
            </template>
          </el-table-column>
          <el-table-column label="显示通道（getValues）" min-width="170">
            <template slot-scope="scope">
              <span class="mono" :title="showScalar(scope.row.readback)">{{ showScalar(scope.row.readback) }}</span>
            </template>
          </el-table-column>
          <el-table-column label="标记" width="140" align="center">
            <template slot-scope="scope">
              <el-tag v-if="scope.row.textLane" size="mini" type="warning">文本通道</el-tag>
              <el-tooltip v-if="scope.row.displayTruncated" content="显示管线格式化截断（15 位有效数字）——SDK 显示行为，存储层无损，不判 FAIL">
                <el-tag size="mini" type="info">显示截断</el-tag>
              </el-tooltip>
            </template>
          </el-table-column>
          <el-table-column prop="verdict" label="判定" width="100" align="center">
            <template slot-scope="scope">
              <el-tag size="mini" :type="resultTagType(scope.row.verdict)">{{ scope.row.verdict }}</el-tag>
            </template>
          </el-table-column>
        </el-table>
      </template>
      <template v-if="formulaRows.length">
        <div class="section-title">计算证据：显式重算 lane（executeCalculation 触发 → 有界轮询直到显示值离开预览缓存 → 取计算结果与存储快照）。计算结果与期望值（evalSimpleFormula）双精度全等判定；saveSnapshot.v === 计算结果为"引擎真实重算写模型"的证据；期望不可判定记人工核对</div>
        <el-table :data="formulaRows" size="mini" border max-height="240">
          <el-table-column prop="sheetId" label="sheetId" width="130" show-overflow-tooltip />
          <el-table-column prop="row" label="行" width="60" align="center" />
          <el-table-column prop="col" label="列" width="60" align="center" />
          <el-table-column label="f（原样）" min-width="140">
            <template slot-scope="scope"><span class="mono">{{ scope.row.f }}</span></template>
          </el-table-column>
          <el-table-column label="getFormula() 读回" min-width="130">
            <template slot-scope="scope"><span class="mono">{{ scope.row.formulaReadback }}</span></template>
          </el-table-column>
          <el-table-column label="预览缓存 v" width="110" align="center">
            <template slot-scope="scope"><span class="mono">{{ showScalar(scope.row.cachedV) }}</span></template>
          </el-table-column>
          <el-table-column label="计算结果" min-width="110">
            <template slot-scope="scope"><span class="mono" :title="showScalar(scope.row.computed)">{{ showScalar(scope.row.computed) }}</span></template>
          </el-table-column>
          <el-table-column label="存储快照 v（save）" min-width="130">
            <template slot-scope="scope"><span class="mono" :title="showScalar(scope.row.saveSnapshot && scope.row.saveSnapshot.v)">{{ scope.row.saveSnapshot ? showScalar(scope.row.saveSnapshot.v) : '(无)' }}</span></template>
          </el-table-column>
          <el-table-column label="重算已应用" width="100" align="center">
            <template slot-scope="scope">
              <el-tag size="mini" :type="scope.row.recalcApplied ? 'success' : 'info'">{{ scope.row.recalcApplied ? '是' : '否' }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column label="期望值（evalSimpleFormula）" min-width="140">
            <template slot-scope="scope"><span class="mono">{{ scope.row.decidable ? showScalar(scope.row.expected) : '不可判定' }}</span></template>
          </el-table-column>
          <el-table-column prop="verdict" label="判定" width="100" align="center">
            <template slot-scope="scope">
              <el-tag size="mini" :type="resultTagType(scope.row.verdict)">{{ scope.row.verdict }}</el-tag>
            </template>
          </el-table-column>
        </el-table>
      </template>
    </el-card>

    <!-- 5. 保存重载面板 -->
    <el-card class="panel" shadow="never">
      <div slot="header"><span>保存重载面板（save → dispose → 快照重建 → 逐格比对 t/v）</span></div>
      <el-form inline size="mini">
        <el-form-item>
          <el-button size="mini" type="primary" :loading="reloadBusy" :disabled="!replayDone" @click="saveReloadCompare">保存→重载→比对</el-button>
        </el-form-item>
      </el-form>
      <el-alert v-if="reloadError" type="error" :closable="false" class="mb8" :title="'保存重载面板最近一次错误（原样）'">
        <div class="mono small">{{ pretty(reloadError) }}</div>
      </el-alert>
      <el-alert
        v-if="compareSummary"
        :type="compareSummary.diffCount === 0 ? 'success' : 'error'"
        :closable="false"
        class="mb8"
        :title="compareTitle"
      />
      <el-table v-if="compareRows.length" :data="compareRows" size="mini" border max-height="320">
        <el-table-column prop="sheetId" label="sheetId" width="130" show-overflow-tooltip />
        <el-table-column prop="row" label="行" width="60" align="center" />
        <el-table-column prop="col" label="列" width="60" align="center" />
        <el-table-column label="首次 save t/v" min-width="200">
          <template slot-scope="scope">
            <span class="mono" :title="showScalar(scope.row.v1)">t={{ scope.row.t1 === undefined ? '-' : scope.row.t1 }} · v={{ showScalar(scope.row.v1) }}</span>
          </template>
        </el-table-column>
        <el-table-column label="重载后 t/v" min-width="200">
          <template slot-scope="scope">
            <span class="mono" :title="showScalar(scope.row.v2)">t={{ scope.row.t2 === undefined ? '-' : scope.row.t2 }} · v={{ showScalar(scope.row.v2) }}</span>
          </template>
        </el-table-column>
        <el-table-column prop="same" label="比对" width="90" align="center">
          <template slot-scope="scope">
            <el-tag size="mini" :type="scope.row.same ? 'success' : 'danger'">{{ scope.row.same ? 'PASS' : 'FAIL' }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column prop="anchor" label="锚点" width="120" align="center">
          <template slot-scope="scope">
            <el-tag v-if="scope.row.anchor" size="mini" type="warning">{{ scope.row.anchor }}</el-tag>
          </template>
        </el-table-column>
      </el-table>

      <!-- 真实后端保存/重载（模拟刷新）：走产品文档保存链路，不做页内自比 -->
      <div class="section-title">真实后端保存/重载（模拟刷新；走产品文档保存链路 POST /ai/workbook/save + GET /ai/workbook/{id}）</div>
      <div class="hint">
        分支/草稿 CAS 属主页面 C4 范围，不在本 harness 内。基线 = 「真实后端保存」时定格的页内 save() 快照（与上传 content 同源）；比对口径：t 与 v 全等，保精度字面量经 SDK save→后端→重载 后必须仍是逐字符相等的文本，公式 f 必须原样（f 不一致即 FAIL，F4b 整改）。「再查一次」连续两次从后端重载比对，证明非缓存巧合。
      </div>
      <el-form inline size="mini">
        <el-form-item>
          <el-button size="mini" type="primary" :loading="backendSaveBusy" :disabled="!replayDone" @click="backendSave">真实后端保存</el-button>
          <el-button size="mini" :loading="backendReloadBusy" :disabled="!backendWorkbookId || !backendBaseline" @click="backendReloadCompare(false)">从后端重载比对（模拟刷新）</el-button>
          <el-button size="mini" plain :loading="backendReloadBusy" :disabled="!backendWorkbookId || !backendBaseline" @click="backendReloadCompare(true)">再查一次</el-button>
        </el-form-item>
      </el-form>
      <el-alert v-if="backendSaveError" type="error" :closable="false" class="mb8" title="真实后端保存最近一次错误（原样）">
        <div class="mono small">{{ pretty(backendSaveError) }}</div>
      </el-alert>
      <el-alert v-if="backendReloadError" type="error" :closable="false" class="mb8" title="从后端重载比对最近一次错误（原样）">
        <div class="mono small">{{ pretty(backendReloadError) }}</div>
      </el-alert>
      <div v-if="backendWorkbookId" class="hint">已保存 workbookId={{ backendWorkbookId }}（name=C3验收-{{ requestId }}），已完成重载比对 {{ backendReloadAttempts }} 次</div>
      <el-alert
        v-if="backendReloadSummary"
        :type="backendReloadSummary.verdict === 'PASS' ? 'success' : 'error'"
        :closable="false"
        class="mb8"
        :title="backendReloadTitle"
      />
      <el-table v-if="backendReloadRows.length" :data="backendReloadRows" size="mini" border max-height="320">
        <el-table-column prop="sheetId" label="sheetId" width="130" show-overflow-tooltip />
        <el-table-column prop="row" label="行" width="60" align="center" />
        <el-table-column prop="col" label="列" width="60" align="center" />
        <el-table-column label="基线 t/v" min-width="200">
          <template slot-scope="scope">
            <span class="mono" :title="showScalar(scope.row.v1)">t={{ scope.row.t1 === undefined ? '-' : scope.row.t1 }} · v={{ showScalar(scope.row.v1) }}</span>
          </template>
        </el-table-column>
        <el-table-column label="后端重载 t/v" min-width="200">
          <template slot-scope="scope">
            <span class="mono" :title="showScalar(scope.row.v2)">t={{ scope.row.t2 === undefined ? '-' : scope.row.t2 }} · v={{ showScalar(scope.row.v2) }}</span>
          </template>
        </el-table-column>
        <el-table-column prop="same" label="比对" width="90" align="center">
          <template slot-scope="scope">
            <el-tag size="mini" :type="scope.row.same ? 'success' : 'danger'">{{ scope.row.same ? 'PASS' : 'FAIL' }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column prop="anchor" label="锚点" width="120" align="center">
          <template slot-scope="scope">
            <el-tag v-if="scope.row.anchor" size="mini" type="warning">{{ scope.row.anchor }}</el-tag>
          </template>
        </el-table-column>
      </el-table>
    </el-card>

    <!-- 请求日志 -->
    <el-card class="panel" shadow="never">
      <div slot="header"><span>请求/响应摘要日志（供导出留证）</span></div>
      <el-table :data="requestLog" size="mini" border max-height="260">
        <el-table-column prop="time" label="时间" width="180" />
        <el-table-column prop="label" label="动作" width="200" show-overflow-tooltip />
        <el-table-column label="摘要" min-width="480">
          <template slot-scope="scope">
            <span class="mono small">{{ compact(scope.row.payload) }}</span>
          </template>
        </el-table-column>
      </el-table>
    </el-card>
  </div>
</template>

<script>
// C3 集成验收专页（独立 harness）：不得侵入 index.vue 主页面；只依赖既有依赖，禁止新增 package.json 依赖。
// 重放口径对齐 tests/agentReport/c3-import-types-regression.cjs：遍历 preview rows，
// sheet.getRange(rowIndex, col).setValue({ v, t, ...(f ? { f } : {}) })（t 1=STRING / 2=NUMBER / 3=BOOLEAN 直传；
// 唯一分道例外：t=2 且 typeof v==='string' 的保精度字面量按 F2 契约走文本通道 t=1 重放——若按 t=2 直传，
// Univer 会在数字通道把字符串物化成 Number，29 位字面量在重放即损；读回判定为逐字符全等）。
//
// 字段契约（创建→预览打通）：创建/查询（createImportJob/getImportJob/getImportJobByRequest）响应字段名为
// sourceContentHash（ExcelImportService.rowView .put("sourceContentHash", ...)），预览查询参数名为 sourceHash
// 且必须与该任务绑定的 source_content_hash 全等（ExcelImportService.preview require(recorded.equals(sourceHash))）。
// 页面统一以 sourceContentHash ?? sourceHash 取值驱动预览按钮，三个查询入口的响应都会刷新 this.sourceHash。
//
// 公式证据：计算结果必须与 ./c3-acceptance-formula 的 evalSimpleFormula 期望值双精度全等（可判定时）；
// 期望不可判定（函数名/字符串/引用缺值/非有限结果）记人工核对，绝不放行（不再"读出任意数值即 PASS"）。
//
// 公式计算引擎注册依据（实测本仓 node_modules，@univerjs 0.25.1 presets 体系）：
// 1) @univerjs/preset-sheets-core@0.25.1 的 package.json dependencies 已含
//    @univerjs/sheets-formula@0.25.1（开源公式计算引擎）与 @univerjs/engine-formula@0.25.1，
//    且其 lib/cjs/index.js 内部直接 import 了 '@univerjs/sheets-formula/lib/facade'、
//    '@univerjs/engine-formula/lib/facade' —— 即 UniverSheetsCorePreset 已经注册公式引擎插件，
//    并把 univerAPI.getFormula()（FFormula）挂进 FUniver facade。
// 2) 因此本页【不引入】@univerjs/preset-sheets-advanced：advanced 捆绑的是 @univerjs-pro/engine-formula
//    （Pro 引擎，需 license），对 "=A1*B1" 级计算证据没有必要。
// 3) 计算触发与等待姿势（2026-11-08 headless 真实 SDK 实证，证据：G:/code/agentExcel/.tmp-c3-replay/evidence-formula-sequences.json）：
//    a. 触发：univerAPI.getFormula().executeCalculation() → SetTriggerFormulaCalculationStartMutation
//       {forceCalculation:true}（engine-formula facade），经 dirty 管理器 100ms 合流后真正起算。
//    b. 等待陷阱：onCalculationResultApplied(timeout) 内部是 FormulaCalculationSessionService
//       .waitForLatestApplied——它对"安静引擎"走 500ms start-watchdog 直接 resolve，对已存在的
//       会话取"首个 applied 终态"，均不保证【本次触发】的结果已应用（sheets-formula cjs 414/505 行实证）。
//       浏览器联调证据（2026-11-08-C3联调证据-c3-replay-blocked.json）即此竞态：viaFacade=true 而
//       computed 读出应用前缓存 13。旧轮询又被"缓存值本身是数字"立即满足而空转。
//    c. 实证结论（headless 0.25.1 引擎级插件环境，页面同款序列）：
//       - 触发后立即读 = 13（应用前缓存，红）；
//       - executeCalculation + 上述 facade 等待 + 数值轮询 = -5627.25（本环境可绿，但等待不可靠）；
//       - 整改姿势（本页采用）：executeCalculation → facade 等待作沉降门（尽力而为）→
//         【有界轮询直到公式格显示值离开预览缓存值】(≤5s) → 读计算结果，并同步取 save() 存储快照
//         （saveSnapshot.v 必须等于计算结果——引擎真实重算写模型的直接证据，非手填）。
//    d. 存储 vs 显示（A3 实证 evidence-a3-three-point.json）：0.25.1 显示通道 getValues() 对
//       17 位有效数字 double 仅保 15 位（45838.99998842592 → 45838.9999884259），而
//       save().cellData 无损。故重放读回断言口径：数值格以【存储层 save 快照 vs 预览值全等】为
//       硬判据，getValues() 仅作显示通道证据并标注 displayTruncated（SDK 显示行为，不判 FAIL）。
// 4) 公式注册路径（2026-11-09 真实浏览器 UI 插件栈实证，证据：.tmp-c3-replay/evidence-ui-stack-strategies.json）：
//    headless 引擎级环境与浏览器 UI 栈行为【不同】——setValue({v,t,f}) 在 UI 栈只把 f 写进 cellData
//    模型，公式引擎不注册：executeCalculation 空转，读回恒为缓存 13，即使 calculationEnd 回调已触发
//    (state=2) 或重复触发（S1/S2/S6/S7/S8）。range.setFormula(f) 走公式命令路径完成注册与标脏，
//    引擎真实重算并写回存储：S4/S5/S9/S10 均得 -5627.25（S10 = 本页口径，553ms，存储同步）。
//    另：setFormula 会先清空缓存再补算，故"离开缓存"轮询必须要求【有限数值且 ≠ 缓存值】，
//    否则瞬时空值会被误判为应用落地。重放阶段仍只写 setValue({v,t,f}) 保留缓存 13——
//    "导入不重算"守卫（重放后 save() 仍为 13）由 c3-replay-calc-regression.cjs 的 CACHE GUARD 固化。
import { LocaleType, mergeLocales, createUniver, UniverInstanceType, defaultTheme } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'

import '@univerjs/presets/lib/styles/preset-sheets-core.css'
// 注册 sheets facade mixin（FUniverSheetsMixin）：getActiveWorkbook / FRange.setValue / getValues / getFormula。
// 关键：必须用 lib/facade.js，与 @univerjs/preset-sheets-core 内部 import 的 '@univerjs/sheets/lib/facade'
// 解析到同一文件（同 index.vue 242 行注释结论，避免加载出第二套 FRange 类）。
import '@univerjs/sheets/lib/facade.js'

import Cookies from 'js-cookie'
import { getToken } from '@/utils/auth'
import { createImportJob, getImportJob, getImportJobByRequest, previewImport } from '@/api/ai/import'
import { saveWorkbook, getWorkbook } from '@/api/ai/workbook'
import { evalSimpleFormula } from './c3-acceptance-formula'
// F1 整改：预览保精度断言的独立期望表——期望值按真实序列化规则数学推导得出
// （Codex 隔离库实测：1e400 负 scale 折叠 401 位整数 FIXTURE_D1_JSON_LITERAL_LENGTH=401；
// 1E-400/4E-324/1.23456789012345E-310 正 scale 保留 BigDecimal toString E 记法原文），
// 独立于任何 API 响应；assertPreviewPrecision 以该表逐格核对（见方法内注释）。
import { FIXTURE_EXPECTATIONS } from './c3-acceptance-fixture-expectations'

// Token 留存键（c3-accept 前缀）：sessionStorage 存粘贴值便于刷新重放；Admin-Token 才是 getToken() 的真实来源。
const C3_TOKEN_KEY = 'c3-accept-token'
const C3_TOKEN_BACKUP_KEY = 'c3-accept-token-backup'
const ADMIN_TOKEN_COOKIE = 'Admin-Token' // 与 src/utils/auth.js 的 TokenKey 一致（getToken 读 Cookie）

// 鉴权失败判定（整改 6 + 二轮 F2 整改）：业务 code 或 HTTP status 命中 401/403 即判鉴权失败，
// 不得当成"连通+鉴权 OK"。调用方须先 Number() 归一（Axios 字符串传输码 Number('ERR_*')=NaN，
// 与 401/403 全等比较恒为 false，天然不误判）。
const isAuthDenied = code => code === 401 || code === 403
// 终审 P2-b：探针（getImportJob('c3-acceptance-probe')）的"预期拒绝"白名单。该 id 必然不存在，
// 正常后端只会回 404/409（IMPORT_FORBIDDEN/NOT_FOUND）或 422；只有这类拒绝能证明链路与鉴权正常。
// 其他业务 code（典型：500 缺表/SQL 异常）属环境故障，绝不能判"连通+鉴权 OK"。
const EXPECTED_PROBE_CODES = new Set([404, 409, 422])

// 断言初始态（data 与 resetDownstreamState 共用，保证跨文件清除后复位口径一致）。
function makeInitialAssertions() {
  return {
    connectivity: { name: '连通与鉴权：业务封套可达；业务 code 或 HTTP status（含 Axios 传输错误的 response.status，ERR_* 传输码不作业务 code）命中 401/403 判鉴权失败', result: '待执行', evidence: '尚未执行连通自检' },
    idempotentReplay: { name: '导入幂等：同 requestId 重发返回同一 jobId', result: '待执行', evidence: '尚未触发重发' },
    previewPrecision: { name: '预览保精度：c3-acceptance.xlsx 夹具期望表逐格核对（独立期望值）；未命中夹具签名只记人工核对', result: '待执行', evidence: '尚未查询预览' },
    replayTypes: { name: '重放读回类型：数字全等 / 布尔 1-0 / 文本逐字符（保精度字面量走文本通道）', result: '待执行', evidence: '尚未重放' },
    formulaCalc: { name: '公式计算：f 原样且可判定项与期望值（evalSimpleFormula）全等；存在不可判定项整体记人工核对不放行', result: '待执行', evidence: '尚未重放' },
    saveReload: { name: '保存→重载：cellData 逐格 t/v 全等', result: '待执行', evidence: '尚未执行保存重载' },
    backendSaveReload: { name: '真实后端保存→重载（模拟刷新）：cellData 逐格 t/v 全等且公式 f 原样', result: '待执行', evidence: '尚未执行真实后端保存' },
    budgetRefusal: { name: '预算拒绝：c3-budget-refusal.xlsx 必须被拒（FAILED），结构化 kind=COMPRESSION_RATIO / limit=100 / observed≥100（流式触发取整口径）', result: '待执行', evidence: '尚未执行预算拒绝自检' }
  }
}

// requestId 生成：crypto.randomUUID() 优先，非安全上下文（http 非本机）或旧浏览器降级为随机串。
function genRequestId() {
  try {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
      return 'c3-accept-' + crypto.randomUUID()
    }
  } catch (e) { /* 降级 */ }
  return 'c3-accept-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 10) + '-' + Math.random().toString(36).slice(2, 10)
}

export default {
  name: 'AgentReportImportAcceptance',
  data() {
    return {
      // 连接面板
      baseApi: process.env.VUE_APP_BASE_API || '(未配置)',
      tokenInput: '',
      tokenPresent: false,
      probeRunning: false,
      probeResult: null,
      // 导入面板
      file: null,
      fileName: '',
      requestId: genRequestId(),
      rowsRequested: 50,
      importBusy: false,
      budgetBusy: false,
      budgetResult: null,
      createResult: null,
      jobStatusResult: null,
      byRequestResult: null,
      resendResult: null,
      jobId: '',
      sourceHash: '',
      importError: null,
      openImportJson: ['create'],
      // 预览面板
      previewBusy: false,
      preview: null,
      previewError: null,
      // Univer 重放面板
      univer: null,
      univerAPI: null,
      replayBusy: false,
      replayDone: false,
      replayError: null,
      readbackRows: [],
      formulaRows: [],
      // 保存重载面板（页内 SDK save→重建自比）
      reloadBusy: false,
      reloadError: null,
      compareRows: [],
      compareSummary: null,
      // 真实后端保存/重载面板（产品文档保存链路 POST /ai/workbook/save + GET /ai/workbook/{id}）
      backendSaveBusy: false,
      backendSaveError: null,
      backendWorkbookId: null,
      backendBaseline: null, // 「真实后端保存」时定格的页内 save() 基线 cellData（对比口径的左侧）
      backendReloadBusy: false,
      backendReloadError: null,
      backendReloadRows: [],
      backendReloadSummary: null,
      backendReloadAttempts: 0,
      // 保精度字面量位置集（'sid|row|col'），由重放阶段登记，供页内/后端两个比对口径打锚点标记。
      precisionPositions: new Set(),
      // 断言汇总（result: 待执行 / PASS / FAIL / NA / 人工核对）
      assertions: makeInitialAssertions(),
      requestLog: []
    }
  },
  computed: {
    assertionRows() {
      return Object.keys(this.assertions).map(key => ({ key, ...this.assertions[key] }))
    },
    hasPreviewRows() {
      return Boolean(this.preview && Array.isArray(this.preview.rows) && this.preview.rows.length)
    },
    allAssertionsGreen() {
      const results = this.assertionRows.map(r => r.result)
      return results.every(r => r === 'PASS') && results.some(r => r === 'PASS')
    },
    greenBannerText() {
      return 'C3 自动断言全绿，人工核对项见验收用例文档'
    },
    previewMetaRows() {
      if (!this.preview) return []
      const p = this.preview
      return [
        { key: 'jobId', value: String(p.jobId !== undefined ? p.jobId : '-') },
        { key: 'sourceHash', value: String(p.sourceHash !== undefined ? p.sourceHash : '-') },
        // 预览响应的 sourceHash 即该任务绑定的源内容哈希（创建/查询响应中字段名为 sourceContentHash），同步展示对齐两者。
        { key: 'sourceContentHash', value: String(p.sourceHash !== undefined ? p.sourceHash : '-') },
        { key: 'sheetCount', value: String(p.sheetCount !== undefined ? p.sheetCount : '-') },
        { key: 'totalRowCount', value: String(p.totalRowCount !== undefined ? p.totalRowCount : '-') },
        { key: 'requestedRows', value: String(p.requestedRows !== undefined ? p.requestedRows : '-') },
        { key: 'returnedRows', value: String(p.returnedRows !== undefined ? p.returnedRows : (p.rows ? p.rows.length : '-')) },
        { key: 'limited', value: String(p.limited !== undefined ? p.limited : '-') },
        { key: 'note', value: String(p.note !== undefined ? p.note : '-') }
      ]
    },
    previewCellRows() {
      if (!this.hasPreviewRows) return []
      const rows = []
      for (const record of this.preview.rows) {
        const sid = String(record.sheetId !== undefined ? record.sheetId : 'sheet-' + record.sheetIndex)
        const row = Number(record.rowIndex) || 0
        for (const colKey of Object.keys(record.cells || {})) {
          const cell = record.cells[colKey] || {}
          rows.push({
            sheetId: sid,
            row,
            col: colKey,
            t: cell.t,
            cell,
            hasF: Boolean(cell.f),
            precisionLiteral: cell.t === 2 && typeof cell.v === 'string'
          })
        }
      }
      return rows
    },
    probeTitle() {
      if (!this.probeResult) return ''
      if (this.probeResult.kind === 'auth-failed') return '鉴权失败（Token 无效/过期）——业务 code 或 HTTP status 命中 401/403，连通断言置 FAIL'
      if (this.probeResult.kind === 'business') return '连通 + 鉴权通过（收到业务错误封套，符合预期）'
      if (this.probeResult.kind === 'unexpected-ok') return '连通（意外返回 code=200，探测 id 竟存在）'
      if (this.probeResult.kind === 'http') return 'HTTP 层异常（非业务封套）'
      return '网络异常：后端可能未启动或地址不可达'
    },
    probeDetail() {
      if (!this.probeResult) return ''
      const r = this.probeResult
      return JSON.stringify({ code: r.code, msg: r.msg, definitive: r.definitive, httpStatus: r.httpStatus, detail: r.detail, raw: r.raw })
    },
    compareTitle() {
      if (!this.compareSummary) return ''
      const s = this.compareSummary
      return `逐格比对完成：共 ${s.total} 格，t/v 不一致 ${s.diffCount} 处，f 不一致 ${s.fDiffCount} 处（布尔/前导零/日期序列/保精度字面量锚点 ${s.anchorCount} 格）`
    },
    backendReloadTitle() {
      if (!this.backendReloadSummary) return ''
      const s = this.backendReloadSummary
      return `第 ${s.attempts} 次从后端重载比对完成（模拟刷新）：共 ${s.total} 格，t/v 不一致 ${s.diffCount} 处，f 不一致 ${s.fDiffCount} 处（保精度字面量/布尔/前导零/日期序列锚点 ${s.anchorCount} 格）`
    }
  },
  mounted() {
    // 刷新后自动恢复粘贴过的 Token：sessionStorage 有值则重写 Cookie，保证 getToken() 能取到。
    const saved = sessionStorage.getItem(C3_TOKEN_KEY)
    if (saved) Cookies.set(ADMIN_TOKEN_COOKIE, saved)
    this.refreshTokenState()
  },
  beforeDestroy() {
    this.disposeUniver()
  },
  methods: {
    // ---------- 通用 ----------
    pretty(value) {
      if (value === undefined || value === null) return '(无)'
      try { return JSON.stringify(value, null, 2) } catch (e) { return String(value) }
    },
    compact(value) {
      try { const s = JSON.stringify(value); return s && s.length > 300 ? s.slice(0, 300) + '…' : s } catch (e) { return String(value) }
    },
    safeClone(value) {
      try { return JSON.parse(JSON.stringify(value)) } catch (e) { return String(value) }
    },
    showScalar(v) {
      if (v === undefined) return '(undefined)'
      if (v === null) return '(null)'
      if (typeof v === 'string') return "'" + v + "'"
      if (typeof v === 'number' || typeof v === 'boolean') return String(v) // 全精度，禁止 toFixed/截断
      try { return JSON.stringify(v) } catch (e) { return String(v) }
    },
    shortHash(hash) {
      const s = String(hash || '')
      return s.length > 16 ? s.slice(0, 8) + '…' + s.slice(-6) : s
    },
    genRequestId,
    resultTagType(result) {
      if (result === 'PASS') return 'success'
      if (result === 'FAIL') return 'danger'
      if (result === '人工核对') return 'warning'
      return 'info'
    },
    setAssertion(key, result, evidence) {
      this.assertions[key] = { name: this.assertions[key].name, result, evidence: String(evidence || '') }
    },
    logRequest(label, payload) {
      this.requestLog.push({ time: new Date().toISOString(), label, payload: this.safeClone(payload) })
      if (this.requestLog.length > 400) this.requestLog.splice(0, this.requestLog.length - 400)
    },
    normalizeError(err) {
      if (!err) return null
      const out = {
        name: err.name || 'Error',
        message: err.message || String(err),
        code: err.code === undefined ? null : err.code,
        definitive: err.definitive === true,
        networkLike: !err.response && err.code === undefined
      }
      if (err.response) out.httpStatus = err.response.status
      return out
    },
    // ---------- 连接面板 ----------
    refreshTokenState() {
      this.tokenPresent = Boolean(getToken())
    },
    applyToken() {
      const token = String(this.tokenInput || '').trim()
      if (!token) { this.$message.warning('请先粘贴 Token'); return }
      // 首次应用前备份原 Cookie，供「清除」还原（与主应用共享 Cookie，必须可回滚）。
      if (sessionStorage.getItem(C3_TOKEN_BACKUP_KEY) === null) {
        const existing = Cookies.get(ADMIN_TOKEN_COOKIE)
        if (existing !== undefined) sessionStorage.setItem(C3_TOKEN_BACKUP_KEY, existing)
      }
      Cookies.set(ADMIN_TOKEN_COOKIE, token)
      sessionStorage.setItem(C3_TOKEN_KEY, token)
      this.refreshTokenState()
      // 证据脱敏：日志只记录长度，不记录任何 token 片段（JWT 头前缀也属不必要暴露）。
      this.logRequest('应用 Token（写入 Admin-Token Cookie + sessionStorage）', { tokenLength: token.length, tokenHead: '(redacted)' })
      this.$message.success('Token 已写入登录 Cookie，getToken() 现在应返回它')
    },
    clearToken() {
      const backup = sessionStorage.getItem(C3_TOKEN_BACKUP_KEY)
      if (backup !== null) {
        Cookies.set(ADMIN_TOKEN_COOKIE, backup)
        sessionStorage.removeItem(C3_TOKEN_BACKUP_KEY)
      } else {
        Cookies.remove(ADMIN_TOKEN_COOKIE)
      }
      sessionStorage.removeItem(C3_TOKEN_KEY)
      this.refreshTokenState()
      this.logRequest('清除 Token（还原 Cookie）', { restoredBackup: backup !== null })
      this.$message.success('已清除本页写入的 Token 并还原 Cookie')
    },
    // 连通自检：getImportJob('c3-acceptance-probe') —— 该 id 必不存在，预期业务错误封套。
    // 有 error.code（含 definitive）=> 网络与鉴权链路通；无响应（Network Error）=> 后端未启动。
    // 鉴权失败口径（整改 6 + F2 整改）：两类错误形态必须区分——
    //   a) 业务封套错误：import.js importRequest 在 code!==200 时抛 Error 并挂【数值】code
    //      （error.code = response.code，definitive 仅 409/422），【没有 response 属性】；
    //   b) Axios 传输错误：HTTP 层有响应时 err.code 是【字符串传输码】（ERR_BAD_REQUEST /
    //      ERR_NETWORK...）且携带 err.response（内含真实 HTTP status）；网络层无响应时同样
    //      是字符串码且无 response。ERR_* 是传输码不是业务 code——旧实现
    //      if (err.code !== undefined) 先行分支把 ERR_BAD_REQUEST 当业务 code 吃掉，导致
    //      HTTP 403 被判"连通+鉴权 OK"假绿。
    // 合并判定：isAuthDenied(Number(err.code)) || (err.response && isAuthDenied(err.response.status))，
    // 任一命中即 auth-failed FAIL；业务封套分支仅在 typeof err.code === 'number' &&
    // Number.isFinite(err.code) 时进入；带 err.response 的 Axios 错误一律按 HTTP status 分类
    // （401/403→auth-failed，其余→http FAIL），绝不把 ERR_* 当业务 code。
    async runProbe() {
      this.probeRunning = true
      this.probeResult = null
      try {
        const data = await getImportJob('c3-acceptance-probe')
        this.probeResult = { kind: 'unexpected-ok', raw: this.safeClone(data) }
        this.setAssertion('connectivity', 'PASS', '连通+鉴权 OK（意外返回 code=200，探测 id 竟存在，链路可达且 Token 有效）')
        this.logRequest('连通自检（意外 code=200）', { data })
      } catch (err) {
        const normalized = this.normalizeError(err)
        const authFailed = (err && err.code !== undefined && err.code !== null && isAuthDenied(Number(err.code))) ||
          Boolean(err && err.response && isAuthDenied(err.response.status))
        const businessEnvelope = Boolean(err) && typeof err.code === 'number' && Number.isFinite(err.code)
        if (businessEnvelope) {
          // 终审 P2-b：只有"探针的预期拒绝"才证明连通+鉴权 OK。预期 = 明确的任务不存在/无权/
          // 参数类拒绝（404/409/422）且错误码前缀是 IMPORT_*。未知服务端错误（如 500 缺表 /
          // SQL 异常 / AI_CAPABILITY_BLOCKED 以外的内部错误）一律 FAIL——真实故障曾表现为
          // code=500 却被判"连通+鉴权 OK"，那会把环境故障当成验收通过。
          const expectedProbeRejection = EXPECTED_PROBE_CODES.has(err.code) && /^\s*IMPORT_[A-Z_]+:/.test(String(err.message || ''))
          this.probeResult = {
            kind: authFailed ? 'auth-failed' : (expectedProbeRejection ? 'business' : 'unexpected-business'),
            code: err.code,
            msg: err.message,
            definitive: err.definitive === true
          }
          if (authFailed) {
            this.setAssertion('connectivity', 'FAIL', `鉴权失败（Token 无效/过期）：业务 code=${err.code}，msg=${err.message}（401/403 不得判为鉴权通过）`)
          } else if (expectedProbeRejection) {
            this.setAssertion('connectivity', 'PASS', `连通+鉴权 OK（预期探针拒绝：code=${err.code}，msg=${err.message}，definitive=${err.definitive === true}）`)
          } else {
            this.setAssertion('connectivity', 'FAIL',
              `未预期服务端错误：code=${err.code}，msg=${err.message}——未知错误不能证明"连通+鉴权 OK"` +
              `（只有预期探针拒绝 ${[...EXPECTED_PROBE_CODES].join('/')} 且消息为 IMPORT_* 才判通过）`)
          }
        } else if (err && err.response) {
          // Axios 传输错误（字符串传输码 + response）：按真实 HTTP status 分类，err.code 仅留证。
          const statusAuth = isAuthDenied(err.response.status)
          this.probeResult = {
            kind: statusAuth ? 'auth-failed' : 'http',
            code: err.code,
            httpStatus: err.response.status,
            detail: normalized.message
          }
          this.setAssertion('connectivity', 'FAIL', statusAuth
            ? `鉴权失败（Token 无效/过期）：HTTP status=${err.response.status}（Axios 传输码 ${err.code} 按 response.status 判定，不是业务 code；401/403 不得判为鉴权通过）`
            : `未连通：HTTP 层异常 status=${err.response.status}（${normalized.message}）`)
        } else {
          this.probeResult = {
            kind: authFailed ? 'auth-failed' : 'network',
            code: err && err.code !== undefined ? err.code : null,
            detail: normalized.message
          }
          this.setAssertion('connectivity', 'FAIL', authFailed
            ? `鉴权失败（Token 无效/过期）：code=${err.code}（401/403 不得判为鉴权通过）`
            : `未连通（网络异常）：${normalized.message}`)
        }
        this.logRequest('连通自检（错误原样）', { normalized })
      } finally {
        this.probeRunning = false
        this.refreshTokenState()
      }
    },
    // ---------- 导入面板 ----------
    // 跨文件遗留状态清除（整改 3c）：选中新文件或创建新任务成功时调用。
    // 旧 jobId/sourceHash/preview/重放/比对/真实后端保存重载全部残留会让"第二次验收"跑的是
    // 文件 A 的预览 + 文件 B 的任务，故这里一次性清空下游全部状态并复位断言（连通断言与文件无关，保留）。
    resetDownstreamState(reason) {
      const had = Boolean(
        this.jobId || this.sourceHash || this.preview || this.previewError ||
        this.readbackRows.length || this.formulaRows.length || this.replayDone || this.univer ||
        this.compareRows.length || this.compareSummary ||
        this.backendWorkbookId || this.backendBaseline || this.backendReloadRows.length || this.backendReloadSummary
      )
      const keepConnectivity = this.assertions.connectivity
      this.preview = null
      this.previewError = null
      this.readbackRows = []
      this.formulaRows = []
      this.replayDone = false
      this.replayError = null
      this.disposeUniver()
      this.compareRows = []
      this.compareSummary = null
      this.reloadError = null
      // 真实后端保存/重载面板状态一并清
      this.backendWorkbookId = null
      this.backendBaseline = null
      this.backendReloadRows = []
      this.backendReloadSummary = null
      this.backendReloadError = null
      this.backendSaveError = null
      this.backendReloadAttempts = 0
      this.precisionPositions = new Set()
      // 旧任务归属的 jobId/sourceHash 一并清除：预览按钮随之为禁用，防止拿旧 hash 预览新任务
      this.jobId = ''
      this.sourceHash = ''
      this.assertions = makeInitialAssertions()
      if (keepConnectivity && keepConnectivity.result !== '待执行') this.assertions.connectivity = keepConnectivity
      this.logRequest('跨文件遗留状态已清除', { reason, had })
    },
    onFileChange(e) {
      const file = e.target.files && e.target.files[0]
      if (!file) return
      if (!/\.xlsx$/i.test(file.name)) {
        this.$message.warning('请选择 .xlsx 文件')
        e.target.value = ''
        return
      }
      this.file = file
      this.fileName = file.name
      this.resetDownstreamState('成功选中新源文件: ' + file.name)
      this.logRequest('选择源文件', { name: file.name, size: file.size, type: file.type })
      e.target.value = '' // 允许重复选择同一文件
    },
    extractJobInfo(data) {
      const info = {}
      if (data && typeof data === 'object') {
        info.jobId = data.jobId !== undefined ? data.jobId : (data.id !== undefined ? data.id : undefined)
        // 字段契约（整改 1）：创建/查询响应字段名为 sourceContentHash（ExcelImportService.rowView），
        // 预览查询参数名为 sourceHash（后端要求与 source_content_hash 全等）。统一取
        // sourceContentHash ?? sourceHash 兼容字段名演进——此前只读 data.sourceHash 导致取到
        // undefined，预览按钮永禁用。
        info.sourceHash = data.sourceContentHash ?? data.sourceHash
        info.status = data.status !== undefined ? data.status : undefined
        info.requestId = data.requestId !== undefined ? data.requestId : undefined
      }
      return info
    },
    async createImport() {
      this.importBusy = true
      this.importError = null
      try {
        const data = await createImportJob(this.file, this.requestId)
        this.createResult = data
        // 新任务创建成功：上一任务的下游状态全部清除（再由本次响应回填 jobId/sourceHash）。
        this.resetDownstreamState('createImport 成功，清除上一任务下游状态')
        const info = this.extractJobInfo(data)
        if (info.jobId !== undefined) this.jobId = String(info.jobId)
        if (info.sourceHash !== undefined && info.sourceHash !== null) this.sourceHash = String(info.sourceHash)
        this.logRequest('createImportJob', { requestId: this.requestId, file: this.fileName, data })
      } catch (err) {
        this.importError = this.normalizeError(err)
        this.logRequest('createImportJob 失败（原样）', { normalized: this.importError })
      } finally {
        this.importBusy = false
      }
    },
    async refreshStatus() {
      this.importBusy = true
      this.importError = null
      try {
        const data = await getImportJob(this.jobId)
        this.jobStatusResult = data
        // view 响应字段名为 sourceContentHash：预览可能晚于创建，hash 可从 view 补齐/刷新。
        const info = this.extractJobInfo(data)
        if (info.jobId !== undefined && info.jobId !== null) this.jobId = String(info.jobId)
        if (info.sourceHash !== undefined && info.sourceHash !== null) this.sourceHash = String(info.sourceHash)
        this.logRequest('getImportJob', { jobId: this.jobId, sourceHash: this.sourceHash, data })
      } catch (err) {
        this.importError = this.normalizeError(err)
        this.logRequest('getImportJob 失败（原样）', { normalized: this.importError })
      } finally {
        this.importBusy = false
      }
    },
    async queryByRequest() {
      this.importBusy = true
      this.importError = null
      try {
        const data = await getImportJobByRequest(this.requestId)
        this.byRequestResult = data
        // byRequest 响应同 view 口径：sourceContentHash 刷新进 sourceHash 状态（补齐/对齐）。
        const info = this.extractJobInfo(data)
        if (info.jobId !== undefined && info.jobId !== null) this.jobId = String(info.jobId)
        if (info.sourceHash !== undefined && info.sourceHash !== null) this.sourceHash = String(info.sourceHash)
        this.logRequest('getImportJobByRequest', { requestId: this.requestId, jobId: this.jobId, sourceHash: this.sourceHash, data })
      } catch (err) {
        this.importError = this.normalizeError(err)
        this.logRequest('getImportJobByRequest 失败（原样）', { normalized: this.importError })
      } finally {
        this.importBusy = false
      }
    },
    // 幂等断言：同 requestId 再次 createImportJob，服务端必须返回与首次一致的 jobId。
    async resendSameRequest() {
      this.importBusy = true
      this.importError = null
      try {
        const data = await createImportJob(this.file, this.requestId)
        this.resendResult = data
        const info = this.extractJobInfo(data)
        const resendJobId = info.jobId !== undefined ? String(info.jobId) : ''
        const pass = Boolean(this.jobId) && resendJobId === this.jobId
        this.setAssertion('idempotentReplay', pass ? 'PASS' : 'FAIL',
          `首次 jobId=${this.jobId || '-'}，重发 jobId=${resendJobId || '-'}，requestId=${this.requestId}`)
        this.logRequest('createImportJob（同 requestId 重发）', { requestId: this.requestId, data, pass })
      } catch (err) {
        this.importError = this.normalizeError(err)
        this.setAssertion('idempotentReplay', 'FAIL', '重发请求失败：' + JSON.stringify(this.importError))
        this.logRequest('createImportJob 重发失败（原样）', { normalized: this.importError })
      } finally {
        this.importBusy = false
      }
    },
    // 预算拒绝自检（UC-09 独立证据）：从站点静态路径取 c3-budget-refusal.xlsx 字节，
    // 仍走【生产 createImportJob】（同一 multipart + requestId + 鉴权路径），只是文件来源换成
    // fetch 而非文件选择器——浏览器自动化（IAB）不支持 filechooser 上传，该入口让 UC-09 可一键复现。
    // 判据：任务必须被拒（FAILED 或业务错误）且结构化 kind=COMPRESSION_RATIO / limit=100 /
    // observed>=100（流式触发 floor 取整口径，实测 100；最终压缩比 505 不等于观测值）。
    async runBudgetRefusalCheck() {
      if (this.budgetBusy) return
      this.budgetBusy = true
      this.budgetResult = null
      try {
        const fixtureUrl = (process.env.BASE_URL || '/') + 'c3-fixtures/c3-budget-refusal.xlsx'
        const resp = await fetch(fixtureUrl)
        if (!resp.ok) throw new Error('夹具获取失败：' + fixtureUrl + ' HTTP ' + resp.status)
        const blob = await resp.blob()
        const file = new File([blob], 'c3-budget-refusal.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
        const requestId = genRequestId()
        let job = null
        let errMessage = ''
        let errCode = null
        let errDefinitive = null
        try {
          job = await createImportJob(file, requestId)
        } catch (err) {
          errMessage = String((err && err.message) || err)
          errCode = err && err.code !== undefined ? err.code : null
          errDefinitive = err && err.definitive !== undefined ? err.definitive : null
          // 后端契约（ExcelImportService.create）：先落库 PARSING → 解析被预算拒绝时 fail() 写
          // FAILED + fail_code/fail_phase/fail_message/fail_detail → 再抛异常由控制器转 409。
          // 即：任务行确实存在且为 FAILED，只是 409 响应体不带 jobId。故按 requestId 回查任务行，
          // 取回服务端账本里的 FAILED 记录（与用户预期的 FAILED/kind/limit/observed 形状对齐）。
          if (errCode === 409 || errCode === 422) {
            try {
              const row = await getImportJobByRequest(requestId)
              if (row && typeof row === 'object') job = row
            } catch (e) { /* 回查失败不影响拒绝判定；evidence 会标注 jobId=- */ }
          }
        }
        const detail = (job && job.failDetail) || {}
        const text = [errMessage, job && job.failMessage, job && job.failDetail ? JSON.stringify(job.failDetail) : ''].filter(Boolean).join(' ')
        // 终审 P2-a：消息解析只作诊断证据，判定必须来自服务端账本。
        const messageKind = (() => {
          const m = /([A-Z][A-Z_]+)\s*上限\s*(\d+)[，,\s]*观察值\s*(\d+)/.exec(text)
          return m ? { kind: m[1], limit: m[2], observed: m[3] } : null
        })()
        const jobStatus = String((job && job.status) || '').toUpperCase()
        const ledgerFailed = jobStatus === 'FAILED'
        // 账本结构化三键必须齐全才可能 PASS：kind/limit/observed 均取自 failDetail（服务端落库事实）。
        // 只拿到拒绝消息、查不到 FAILED 行（或行里没有 failDetail）一律 FAIL——否则"消息像拒绝"
        // 就能冒充"服务端确实拒绝了"。
        const ledgerStructured = Boolean(
          ledgerFailed && detail && typeof detail === 'object' &&
          detail.kind !== undefined && detail.limit !== undefined && detail.observed !== undefined
        )
        const ledgerMatches = ledgerStructured &&
          detail.kind === 'COMPRESSION_RATIO' && Number(detail.limit) === 100 && Number(detail.observed) >= 100
        const kind = ledgerStructured ? detail.kind : (messageKind ? messageKind.kind : null)
        const limit = ledgerStructured ? detail.limit : (messageKind ? messageKind.limit : null)
        const observed = ledgerStructured ? detail.observed : (messageKind ? messageKind.observed : null)
        const pass = ledgerMatches
        const jobRefused = ledgerFailed || Boolean(errMessage)
        this.budgetResult = {
          requestId,
          jobId: job && job.jobId ? job.jobId : null,
          status: jobStatus || (errMessage ? 'ERROR' : '(未知)'),
          ledgerFailed,
          ledgerStructured,
          jobRefused,
          errorCode: errCode,
          errorDefinitive: errDefinitive,
          kind: kind === undefined ? null : kind,
          limit: limit === undefined ? null : limit,
          observed: observed === undefined ? null : observed,
          messageKind: messageKind ? messageKind.kind : null,
          messageStructured: messageKind || null,
          errorMessage: errMessage || null,
          failMessage: (job && job.failMessage) || null,
          failDetail: (job && job.failDetail) || null
        }
        const verdictReason = ledgerMatches
          ? '服务端账本 FAILED 行 + failDetail 结构化三键对齐'
          : (!ledgerFailed
            ? '未取得 FAILED 账本行（仅凭拒绝消息不得判通过）'
            : (!ledgerStructured
              ? 'FAILED 账本行缺少结构化 failDetail（kind/limit/observed）'
              : `账本结构化事实不符（需 COMPRESSION_RATIO / 100 / observed≥100）`))
        this.setAssertion('budgetRefusal', pass ? 'PASS' : 'FAIL',
          `夹具=${fixtureUrl}，requestId=${requestId}，jobId=${this.budgetResult.jobId || '-'}，status=${this.budgetResult.status}` +
          `，业务 code=${this.budgetResult.errorCode === null ? '-' : this.budgetResult.errorCode} definitive=${this.budgetResult.errorDefinitive === null ? '-' : this.budgetResult.errorDefinitive}` +
          `；账本 kind=${this.budgetResult.kind} limit=${this.budgetResult.limit} observed=${this.budgetResult.observed}` +
          `；判定依据：${verdictReason}` +
          `${messageKind ? `；消息诊断（不参与判定）kind=${messageKind.kind} limit=${messageKind.limit} observed=${messageKind.observed}` : ''}` +
          `${errMessage ? '；业务错误原文：' + errMessage : ''}`)
        this.logRequest('预算拒绝自检（UC-09）', this.safeClone(this.budgetResult))
      } catch (err) {
        this.setAssertion('budgetRefusal', 'FAIL', '预算拒绝自检未完成：' + String((err && err.message) || err))
        this.logRequest('预算拒绝自检失败（原样）', { normalized: this.normalizeError(err) })
      } finally {
        this.budgetBusy = false
      }
    },
    // ---------- 预览面板 ----------
    async queryPreview() {
      this.previewBusy = true
      this.previewError = null
      try {
        // 跨文件遗留防线（整改 3c）：残留 preview 与当前 jobId 不匹配时强制重查（正常由
        // resetDownstreamState 拦截，这里兜底——例如开发者工具改状态或历史遗留快照）。
        if (this.preview && this.preview.jobId !== undefined && this.jobId && String(this.preview.jobId) !== this.jobId) {
          this.logRequest('预览与当前任务不匹配，强制重查', { stalePreviewJobId: String(this.preview.jobId), currentJobId: this.jobId })
          this.preview = null
        }
        const data = await previewImport(this.jobId, this.sourceHash, this.rowsRequested)
        this.preview = data
        this.assertPreviewPrecision()
        this.logRequest('previewImport', { jobId: this.jobId, sourceHash: this.sourceHash, rows: this.rowsRequested, returnedRows: data && data.rows ? data.rows.length : 0 })
      } catch (err) {
        this.previewError = this.normalizeError(err)
        this.logRequest('previewImport 失败（原样）', { normalized: this.previewError })
      } finally {
        this.previewBusy = false
      }
    },
    // 保精度字面量断言（F1 整改）：旧实现只做形态检查（length===31 && /^0\.\d{29}$/），
    // 把正确字面量换成 0.99999999999999999999999999999 仍 PASS——形态合法但值错误。
    // 整改后改为 c3-acceptance-fixture-expectations.js 独立期望表逐格核对：期望值按
    // 真实序列化规则数学推导得出（Codex 隔离库实测：1e400 负 scale 折叠 401 位整数；
    // 1E-400/4E-324/1.23456789012345E-310 正 scale 保留 E 记法原文），
    // 独立于任何 API 响应。流程：
    //   1) 先探测 preview 是否含期望表全部坐标（夹具签名）——命中则逐格比对 t 类型全等 +
    //      v 逐字符全等（字符串）或数值/布尔全等，任何一格不符即 FAIL 并列出差异明细；
    //   2) 未命中夹具签名（用户上传了其他文件）时，保留旧 beyond-double 检查作证据留档，
    //      但无独立期望值不得宣称通过——断言只能记「人工核对」。
    assertPreviewPrecision() {
      const expectationCount = Object.keys(FIXTURE_EXPECTATIONS).length
      const signature = this.matchFixtureSignature()
      if (!signature) {
        const literalCell = this.findPrecisionLiteralCell()
        if (!literalCell) {
          this.setAssertion('previewPrecision', '人工核对',
            `未命中 c3-acceptance.xlsx 夹具签名（期望表 ${expectationCount} 格坐标未全部出现）且预览中无 t=2 且 typeof v==='string' 的保精度字面量单元格——无独立期望值，不自动判定，请人工核对`)
          return
        }
        const v = literalCell.cell.v
        const beyondDouble = typeof v === 'string' && Number.isFinite(Number(v)) && String(Number(v)) !== v
        const formOk = typeof v === 'string' && v.length === 31 && /^0\.\d{29}$/.test(v)
        this.setAssertion('previewPrecision', '人工核对',
          `未命中夹具签名（用户上传了其他文件），无独立期望值不得自动 PASS——仅留证据：字面量='${v}'` +
          `（typeof=${typeof v}，长度=${v.length}，形态 ${formOk ? '匹配' : '不匹配'} 0.x29×31；` +
          `beyond-double ${beyondDouble ? '成立：String(Number(v))!==v，字面量未被收敛' : '不成立'}）` +
          ` @ row=${literalCell.row} col=${literalCell.col}；请人工核对`)
        return
      }
      const diffs = []
      for (const key of Object.keys(FIXTURE_EXPECTATIONS)) {
        diffs.push(...this.expectationDiffsOf(key, FIXTURE_EXPECTATIONS[key], signature.cells.get(key) || {}))
      }
      if (diffs.length) {
        this.setAssertion('previewPrecision', 'FAIL',
          `夹具签名命中（期望表 ${expectationCount} 格），但 ${diffs.length} 处与独立期望不符：${diffs.join('；')}`)
        return
      }
      this.setAssertion('previewPrecision', 'PASS',
        `夹具签名命中：c3-acceptance.xlsx 已知 ${expectationCount} 格逐格核对全等` +
        `（t 类型全等；字符串 v 逐字符全等、数值/布尔 v 全等；期望值依据真实序列化规则数学推导——负 scale 折叠整数、正 scale BigDecimal toString E 记法，独立于本次响应）`)
    },
    // 夹具签名探测：preview 是否包含期望表全部坐标（'{sheetIndex}:{row}:{col}'，0 基）。
    // 全部出现才算命中（任一缺失即视为其他文件），返回坐标→单元格映射供逐格核对。
    matchFixtureSignature() {
      const cellsByCoordinate = new Map()
      for (const record of (this.preview && this.preview.rows) || []) {
        const sheetIndex = Number(record.sheetIndex) || 0
        const row = Number(record.rowIndex) || 0
        for (const colKey of Object.keys(record.cells || {})) {
          cellsByCoordinate.set(sheetIndex + ':' + row + ':' + (Number(colKey) || 0), record.cells[colKey] || {})
        }
      }
      const missing = Object.keys(FIXTURE_EXPECTATIONS).filter(key => !cellsByCoordinate.has(key))
      return missing.length ? null : { cells: cellsByCoordinate, missing }
    },
    // 单格比对（F1 整改）：t 类型全等；期望字符串 v 逐字符全等（定位首个差异字符），
    // 数值/布尔 v 全等。返回差异明细（坐标/期望/实际），空数组即全等。
    expectationDiffsOf(key, expected, actual) {
      // 长字面量（如 D1 的 401 位）在证据里截断展示：头 32 + 尾 16 + 总长，保证可定位又不淹没明细。
      const clip = s => {
        const t = String(s)
        return t.length > 64 ? t.slice(0, 32) + '…' + t.slice(-16) + `（共 ${t.length} 字符）` : t
      }
      const diffs = []
      if (actual.t !== expected.t) {
        diffs.push(`${key} t 不符：期望 ${expected.t}，实际 ${actual.t === undefined ? '(undefined)' : actual.t}`)
      }
      const ev = expected.v
      const av = actual.v
      if (typeof ev === 'string') {
        if (typeof av !== 'string') {
          diffs.push(`${key} v 不符：期望字符串 '${clip(ev)}'（长度 ${ev.length}），实际 ${typeof av} ${clip(this.showScalar(av))}`)
        } else if (av !== ev) {
          let at = -1
          for (let i = 0; i < Math.max(av.length, ev.length); i++) {
            if (av[i] !== ev[i]) { at = i; break }
          }
          diffs.push(`${key} v 不符：期望 '${clip(ev)}'（长度 ${ev.length}），实际 '${clip(av)}'（长度 ${av.length}）${at >= 0 ? `，第 ${at + 1} 个字符起不一致` : ''}`)
        }
      } else if (typeof ev === 'number') {
        if (typeof av !== 'number' || av !== ev) {
          diffs.push(`${key} v 不符：期望数值 ${this.showScalar(ev)}，实际 ${this.showScalar(av)}`)
        }
      } else if (typeof ev === 'boolean') {
        if (typeof av !== 'boolean' || av !== ev) {
          diffs.push(`${key} v 不符：期望布尔 ${this.showScalar(ev)}，实际 ${this.showScalar(av)}`)
        }
      }
      return diffs
    },
    findPrecisionLiteralCell() {
      for (const record of (this.preview && this.preview.rows) || []) {
        for (const colKey of Object.keys(record.cells || {})) {
          const cell = record.cells[colKey] || {}
          if (cell.t === 2 && typeof cell.v === 'string') {
            return { row: Number(record.rowIndex) || 0, col: Number(colKey), cell }
          }
        }
      }
      return null
    },
    // ---------- Univer 重放 ----------
    buildUniver() {
      // 与 index.vue 相同的 core preset 组装（含公式引擎，见顶部注释依据），容器为本页专属 div。
      this.disposeUniver()
      const zhCNLocales = mergeLocales({}, UniverPresetSheetsCoreZhCN.default || UniverPresetSheetsCoreZhCN)
      const { univer, univerAPI } = createUniver({
        locale: LocaleType.ZH_CN,
        locales: { [LocaleType.ZH_CN]: zhCNLocales },
        theme: defaultTheme,
        presets: [
          UniverSheetsCorePreset({ container: this.$refs.univerHost })
        ]
      })
      this.univer = univer
      this.univerAPI = univerAPI
    },
    disposeUniver() {
      // 无论 univer 是否存在都清空两个引用：只清 half-state 会留下 univerAPI 为旧实例的悬空引用
      // （联调出现 "Cannot read properties of null (reading 'getActiveWorkbook')" 的成因之一）。
      if (this.univer) {
        try {
          this.univer.dispose()
        } catch (e) {
          console.warn('销毁 Univer 实例时出错:', e)
        }
      }
      this.univer = null
      this.univerAPI = null
    },
    // 生命周期等待：UI 插件栈的 active workbook 登记可能异步完成。有界轮询（默认 3s），
    // 不把时序问题当失败；超时返回 null 由调用方给出结构化错误。
    async waitForActiveWorkbook(budgetMs) {
      const deadline = Date.now() + (budgetMs || 3000)
      while (Date.now() < deadline) {
        if (this.univerAPI) {
          try {
            const book = this.univerAPI.getActiveWorkbook()
            if (book) return book
          } catch (e) { /* 实例已销毁按未就绪处理 */ }
        }
        await new Promise(resolve => setTimeout(resolve, 20))
      }
      return null
    },
    // 由 preview rows 构建重放工作簿：每个 sheetId 一张 sheet，行列数按预览数据外扩 2 行/列。
    buildReplayWorkbookData() {
      const meta = new Map()
      for (const record of (this.preview && this.preview.rows) || []) {
        const sid = String(record.sheetId !== undefined ? record.sheetId : 'sheet-' + record.sheetIndex)
        const m = meta.get(sid) || { id: sid, name: sid, index: Number(record.sheetIndex) || 0, maxRow: 0, maxCol: 0 }
        m.maxRow = Math.max(m.maxRow, Number(record.rowIndex) || 0)
        for (const colKey of Object.keys(record.cells || {})) m.maxCol = Math.max(m.maxCol, Number(colKey) || 0)
        meta.set(sid, m)
      }
      const ordered = [...meta.values()].sort((a, b) => a.index - b.index)
      const sheets = {}
      for (const m of ordered) {
        sheets[m.id] = { id: m.id, name: m.name, rowCount: m.maxRow + 2, columnCount: m.maxCol + 2 }
      }
      return {
        id: 'c3-acceptance',
        name: 'C3 导入验收重放',
        sheetOrder: ordered.map(m => m.id),
        sheets,
        activeSheet: ordered.length ? ordered[0].id : undefined
      }
    },
    resolveSheet(book, sid) {
      if (typeof book.getSheetBySheetId === 'function') {
        const sheet = book.getSheetBySheetId(sid)
        if (sheet) return sheet
      }
      const all = typeof book.getSheets === 'function' ? book.getSheets() : []
      const found = all.find(s => s && typeof s.getSheetId === 'function' && s.getSheetId() === sid)
      return found || book.getActiveSheet()
    },
    collectReplayTargets() {
      // 汇总全部待重放单元格（含所属 sheetId），重放与读回共用。
      const targets = []
      for (const record of (this.preview && this.preview.rows) || []) {
        const sid = String(record.sheetId !== undefined ? record.sheetId : 'sheet-' + record.sheetIndex)
        const row = Number(record.rowIndex) || 0
        for (const colKey of Object.keys(record.cells || {})) {
          const cell = record.cells[colKey] || {}
          targets.push({ sid, row, col: Number(colKey), cell })
        }
      }
      return targets
    },
    // 布尔口径（与 c3-import-types-regression.cjs 一致）：线缆 JSON true/false 在 SDK 边界归一为 1/0。
    expectedReadback(cell) {
      if (cell.t === 3) return typeof cell.v === 'boolean' ? (cell.v ? 1 : 0) : cell.v
      if (cell.t === 2 && typeof cell.v === 'number') return cell.v
      if (cell.t === 1) return cell.v
      // 保精度字面量（t=2 且 typeof v==='string'）按 F2 契约走文本通道 t=1 重放，读回必须仍是
      // 逐字符相等的字符串（整改 2：不再"人工核对"放行——重放损坏即 FAIL）。
      if (cell.t === 2 && typeof cell.v === 'string') return cell.v
      return null
    },
    async replayToWorkbook() {
      if (!this.hasPreviewRows) return
      // 并发守卫（联调证据 2026-11-09：getActiveWorkbook 空引用后重试成功）：重放会 dispose 并重建
      // Univer 实例，两次并发重放会互相销毁对方的实例，产生 null 引用与"看似失败其实被覆盖"的假象。
      if (this.replayBusy) return
      this.replayBusy = true
      this.replayError = null
      this.readbackRows = []
      this.formulaRows = []
      this.replayDone = false
      try {
        await this.$nextTick()
        this.buildUniver()
        this.univer.createUnit(UniverInstanceType.UNIVER_SHEET, this.buildReplayWorkbookData())
        // 生命周期等待：UI 插件栈下 active workbook 的登记可能晚于 createUnit（联调曾出现
        // getActiveWorkbook 为空）。有界等待（≤3s）而不是立即抛错，避免把时序问题报成失败。
        const book = await this.waitForActiveWorkbook(3000)
        if (!book) throw new Error('Univer 工作簿创建失败（getActiveWorkbook 3s 内为空）')
        const targets = this.collectReplayTargets()
        // 登记保精度字面量位置，供页内与后端两个比对口径打"保精度字面量"锚点（整改 2 / 整改 4）。
        this.precisionPositions = new Set(
          targets.filter(t => t.cell.t === 2 && typeof t.cell.v === 'string').map(t => t.sid + '|' + t.row + '|' + t.col)
        )
        const sheetCache = new Map()
        const sheetOf = sid => {
          if (!sheetCache.has(sid)) sheetCache.set(sid, this.resolveSheet(book, sid))
          return sheetCache.get(sid)
        }
        // 重放（口径同 c3-import-types-regression.cjs replay()）：setValue({ v, t, ...(f?{f}:{}) })，t 1/2/3 直传。
        // 分道整改（整改 2，F2 契约"C4 走 Univer 文本通道"）：t=2 且 typeof v==='string' 的保精度
        // 字面量改走文本通道 t=1——t=2 数字通道会把字符串字面量物化成 Number，29 位字面量在重放即损。
        for (const target of targets) {
          const textLane = target.cell.t === 2 && typeof target.cell.v === 'string'
          const payload = { v: target.cell.v, t: textLane ? 1 : target.cell.t }
          if (target.cell.f) payload.f = target.cell.f
          sheetOf(target.sid).getRange(target.row, target.col).setValue(payload)
        }
        this.collectReadbackEvidence(targets, sheetOf)
        await this.collectFormulaEvidence(targets, sheetOf)
        this.assertReplayTypes()
        this.assertFormulaCalc()
        this.replayDone = true
        this.logRequest('Univer 重放', { cells: targets.length, readbackRows: this.readbackRows.length, formulaCells: this.formulaRows.length, textLaneCells: this.precisionPositions.size })
      } catch (err) {
        this.replayError = this.normalizeError(err)
        this.logRequest('Univer 重放失败（原样）', { normalized: this.replayError })
      } finally {
        this.replayBusy = false
      }
    },
    collectReadbackEvidence(targets, sheetOf) {
      // 存储通道快照：每个 sheetId 只 save() 一次，按工作表真实 sheetId 索引 cellData。
      // 实证依据（.tmp-c3-replay/evidence-a3-three-point.json）：0.25.1 getValues() 显示通道
      // 对 17 位 double 仅保 15 位，save().cellData 无损——数值格硬判据落在存储层。
      // 实例可能已被并发操作销毁：防御式取用，缺失时快照为 null → 后续按 FAIL 判定，不抛 null 引用。
      const book = this.univerAPI ? this.univerAPI.getActiveWorkbook() : null
      const snapshotBySid = new Map()
      const cellDataOf = sid => {
        if (!snapshotBySid.has(sid)) {
          let cellData = {}
          try {
            const ws = sheetOf(sid)
            const sheetId = typeof ws.getSheetId === 'function' ? ws.getSheetId() : sid
            const snap = book.save()
            cellData = (snap && snap.sheets && snap.sheets[sheetId] && snap.sheets[sheetId].cellData) || {}
          } catch (e) { /* 快照缺失按 undefined 处理，后续判 FAIL */ }
          snapshotBySid.set(sid, cellData)
        }
        return snapshotBySid.get(sid)
      }
      const rows = []
      for (const target of targets) {
        const readback = sheetOf(target.sid).getRange(target.row, target.col).getValues()[0][0]
        const expected = this.expectedReadback(target.cell)
        const precisionLiteral = target.cell.t === 2 && typeof target.cell.v === 'string'
        const textLane = precisionLiteral
        const numericCell = target.cell.t === 2 && typeof target.cell.v === 'number'
        const rowData = cellDataOf(target.sid)[target.row]
        const saveSnapshot = rowData ? rowData[target.col] : undefined
        let verdict
        if (expected === null) {
          verdict = '人工核对'
        } else if (numericCell) {
          // 硬判据：存储层（save 快照）与预览值全等（t 与 v 双等，double 直接 ===）。
          // 显示层差异（getValues 15 位截断）只记录 displayTruncated，不判 FAIL——不静默放宽：
          // 存储不符仍是 FAIL，显示差异单独计数并在三点对照中展示。
          verdict = (saveSnapshot && saveSnapshot.t === 2 && saveSnapshot.v === target.cell.v) ? 'PASS' : 'FAIL'
        } else {
          verdict = readback === expected ? 'PASS' : 'FAIL'
        }
        rows.push({
          sheetId: target.sid,
          row: target.row,
          col: target.col,
          t: target.cell.t,
          v: target.cell.v,
          f: target.cell.f || '',
          precisionLiteral,
          textLane,
          readback,
          expected,
          saveSnapshot: saveSnapshot === undefined ? null : saveSnapshot,
          // 显示通道截断标注：数值格显示值与预览值（或存储值）不一致即 SDK 显示行为证据。
          displayTruncated: numericCell && readback !== target.cell.v,
          verdict
        })
      }
      this.readbackRows = rows
    },
    // 引用解析（整改 3a）：把 evalSimpleFormula 回传的规范化引用（大写、无 $，如 'A1'）解析为
    // 0 基行列，再从重放后的 Univer sheet 读该格当前值；非有限数值（空格/文本/错误值）返回 null
    // ——整条公式随之"不可判定"，记人工核对。
    parseCellRef(ref) {
      const m = /^([A-Za-z]{1,3})([0-9]{1,7})$/.exec(String(ref || ''))
      if (!m) return null
      const row = Number(m[2]) - 1
      if (row < 0) return null
      let col = 0
      for (let i = 0; i < m[1].length; i++) col = col * 26 + (m[1].charCodeAt(i) - 64)
      return { row, col: col - 1 }
    },
    // 计算证据（2026-11-08 实证整改，依据见顶部注释 3）：
    //   触发 = facade executeCalculation()（forceCalculation 全量重算）；
    //   等待 = facade onCalculationResultApplied 仅作沉降门（其实现可对"首个 applied 会话/安静引擎"
    //          提前 resolve，不保证本次触发的结果已应用——浏览器联调 FAIL 即此竞态），
    //          真正有界等待 = 轮询公式格显示值【离开预览缓存值】（≤5s）；
    //   取证 = 计算结果（getValues 显示通道）+ saveSnapshot（save() 存储通道快照）：
    //          引擎真实重算会把结果写回 cellData，saveSnapshot.v === computed 即"真实计算非手填"的
    //          模型层证据；recalcApplied 记录显示值是否离开过缓存（应用发生过的直接观察）。
    // 缓存语义守卫：重放阶段（replayToWorkbook）对公式格只 setValue({v,t,f}) 写入缓存 13，
    //   显式重算发生在本 lane 内——重放后、executeCalculation 前 save() 中 v 仍=13 由
    //   c3-replay-calc-regression.cjs 的缓存语义守卫用例固化。
    // 期望值口径（整改 3a）：每个公式单元格的期望 = evalSimpleFormula(f, 引用→重放后数值)，
    // 计算结果与期望双精度全等才算 PASS；期望不可判定（函数名/字符串/引用缺值等）记人工核对。
    async collectFormulaEvidence(targets, sheetOf) {
      const formulaTargets = targets.filter(t => t.cell && t.cell.f)
      if (!formulaTargets.length) {
        this.formulaRows = []
        return
      }
      const tracked = formulaTargets.map(t => ({
        ...t,
        range: sheetOf(t.sid).getRange(t.row, t.col)
      }))
      // 重算前存储快照（缓存语义证据）：此刻 save() 中公式格仍是预览缓存值（13）——证明重放阶段
      // 没有偷偷重算，真实计算只发生在本 lane 之后。
      const preCalcSnapshotOf = fc => {
        try {
          const b = this.univerAPI ? this.univerAPI.getActiveWorkbook() : null
          if (!b || typeof b.save !== 'function') return null
          const ws = sheetOf(fc.sid)
          const sheetId = typeof ws.getSheetId === 'function' ? ws.getSheetId() : fc.sid
          const snap = b.save()
          const rowData = snap && snap.sheets && snap.sheets[sheetId] && snap.sheets[sheetId].cellData ? snap.sheets[sheetId].cellData[fc.row] : undefined
          return rowData ? rowData[fc.col] || null : null
        } catch (e) { return null }
      }
      for (const fc of tracked) fc.preCalcSnapshot = preCalcSnapshotOf(fc)
      // 公式注册（2026-11-09 真实浏览器实证，G:/code/agentExcel/.tmp-c3-replay/probe 结果）：
      //   setValue({v,t,f}) 只把 f 写进 cellData 模型，UI 插件栈的公式引擎【不会注册】该公式——
      //   executeCalculation 空转，读回永远是预览缓存 13，即使 calculationEnd 回调已触发（state=2）
      //   或重复触发也一样（S1/S2/S3/S6/S7 均 13）。range.setFormula(f) 走公式命令路径完成注册与
      //   标脏，引擎真实重算并把结果写回存储（S4/S5：-5627.25，存储同步，~300ms）。
      //   因此注册必须发生在本 lane；重放阶段仍只写 setValue({v,t,f}) 保留缓存，"导入不重算"守卫
      //   （重放后 save() 仍为 13）由 c3-replay-calc-regression.cjs 固化。
      let registeredVia = 'none'
      for (const fc of tracked) {
        try {
          if (typeof fc.range.setFormula === 'function') {
            fc.range.setFormula(fc.cell.f)
            registeredVia = 'setFormula'
          }
        } catch (e) { /* 注册失败保留证据：后续判定会按"未算出期望值"如实 FAIL */ }
      }
      const deadlineMs = 5000
      let viaFacade = false
      const fe = (this.univerAPI && typeof this.univerAPI.getFormula === 'function') ? this.univerAPI.getFormula() : null
      if (fe && typeof fe.executeCalculation === 'function') {
        viaFacade = true
        try { fe.executeCalculation() } catch (e) { viaFacade = false }
        if (viaFacade && typeof fe.onCalculationResultApplied === 'function') {
          try { await fe.onCalculationResultApplied(deadlineMs) } catch (e) { /* 沉降门提前/超时均不视为失败，转取值轮询 */ }
        }
      }
      // 有界轮询：直到每个公式格显示值离开其预览缓存值（应用落地的可观察信号）。
      // 只认"有限数值且不等于缓存值"——setFormula 会先把缓存清成空值再补算，空值不算应用落地。
      // 若重算结果恰等于缓存值（理论上可能），轮询跑满 5s 后按当时显示值判定（该值即期望值时
      // 证据中 recalcApplied=false 如实标注，不额外放行）。
      const deadline = Date.now() + deadlineMs
      const leftCache = new Set()
      while (Date.now() < deadline) {
        for (const fc of tracked) {
          try {
            const val = fc.range.getValues()[0][0]
            if (typeof val === 'number' && Number.isFinite(val) && val !== fc.cell.v) leftCache.add(fc.sid + '|' + fc.row + '|' + fc.col)
          } catch (e) { /* 读取失败按未离开缓存处理 */ }
        }
        if (leftCache.size === tracked.length) break
        await new Promise(resolve => setTimeout(resolve, 100))
      }
      // 解析器在计算完成后取值：引用格经 Univer 重放/重算后的当前值才是"重放后的数值"。
      const resolveReplayNumber = (sid, ref) => {
        const parsed = this.parseCellRef(ref)
        if (!parsed) return null
        try {
          const val = sheetOf(sid).getRange(parsed.row, parsed.col).getValues()[0][0]
          return typeof val === 'number' && Number.isFinite(val) ? val : null
        } catch (e) { return null }
      }
      const book = this.univerAPI ? this.univerAPI.getActiveWorkbook() : null
      this.formulaRows = tracked.map(fc => {
        const formulaReadback = typeof fc.range.getFormula === 'function' ? fc.range.getFormula() : null
        const computed = fc.range.getValues()[0][0]
        const book = this.univerAPI ? this.univerAPI.getActiveWorkbook() : null
        const saveSnapshot = (() => {
        try {
          const ws = sheetOf(fc.sid)
          const sheetId = typeof ws.getSheetId === 'function' ? ws.getSheetId() : fc.sid
          const snap = book.save()
          const rowData = (snap && snap.sheets && snap.sheets[sheetId] && snap.sheets[sheetId].cellData) ? snap.sheets[sheetId].cellData[fc.row] : undefined
          const cell = rowData ? rowData[fc.col] : undefined
          return cell === undefined ? null : cell
        } catch (e) { return null }
        })()
        const expected = evalSimpleFormula(fc.cell.f, ref => resolveReplayNumber(fc.sid, ref))
        const decidable = expected !== null
        const recalcApplied = leftCache.has(fc.sid + '|' + fc.row + '|' + fc.col)
        return {
          sheetId: fc.sid,
          row: fc.row,
          col: fc.col,
          f: fc.cell.f,
          cachedV: fc.cell.v,
          formulaReadback: formulaReadback === null || formulaReadback === undefined ? '(不可读)' : String(formulaReadback),
          computed,
          saveSnapshot,
          // 重算前存储快照：应为预览缓存（v=13, f 原样）——"重放未重算"的证据，与 computed 对照。
          preCalcSnapshot: fc.preCalcSnapshot || null,
          recalcApplied,
          registeredVia,
          expected,
          decidable,
          match: decidable && computed === expected, // 双精度直接 ===（Univer 与 JS 对同样乘法产生同一 double）
          verdict: decidable ? (computed === expected ? 'PASS' : 'FAIL') : '人工核对',
          viaFacade
        }
      })
    },
    assertReplayTypes() {
      const assertable = this.readbackRows.filter(r => r.verdict !== '人工核对')
      const failed = assertable.filter(r => r.verdict === 'FAIL')
      const countBy = predicate => this.readbackRows.filter(predicate).length
      const numbers = countBy(r => r.t === 2 && typeof r.v === 'number')
      const booleans = countBy(r => r.t === 3)
      const texts = countBy(r => r.t === 1)
      const precision = countBy(r => r.precisionLiteral)
      // 显示层差异单独计数（实证：0.25.1 getValues 显示通道对 17 位 double 仅保 15 位，
      // save().cellData 无损——A3 三点对照 evidence-a3-three-point.json）。显示截断是 SDK 显示
      // 行为证据，不是存储损失，不判 FAIL；数值格硬判据已落在存储层（saveSnapshot 与预览值全等）。
      const displayTruncated = this.readbackRows.filter(r => r.displayTruncated)
      if (!assertable.length) {
        this.setAssertion('replayTypes', 'NA', `无可自动判定的单元格（数字 ${numbers}/布尔 ${booleans}/文本 ${texts}，保精度字面量 ${precision}）`)
        return
      }
      const sampleFail = failed.slice(0, 3).map(r => `r${r.row}c${r.col}: 预览 ${this.showScalar(r.v)} → 存储 ${this.showScalar(r.saveSnapshot && r.saveSnapshot.v)}`).join('；')
      const sampleDisplay = displayTruncated.slice(0, 3).map(r => `r${r.row}c${r.col}: 预览 ${this.showScalar(r.v)} → 存储 ${this.showScalar(r.saveSnapshot && r.saveSnapshot.v)} → 显示 ${this.showScalar(r.readback)}`).join('；')
      // 汇总口径（整改 2 + 2026-11-08 实证）：数值格 = 存储层（save 快照）与预览值全等判 FAIL；
      // 保精度字面量走文本通道 t=1 逐字符判定；布尔按 SDK 边界 1/0；显示层截断单独计数展示。
      this.setAssertion('replayTypes', failed.length === 0 ? 'PASS' : 'FAIL',
        `数字 ${numbers}（存储层全等，显示层截断 ${displayTruncated.length} 格另计不判 FAIL）/布尔 ${booleans}（SDK 边界 1/0 口径）/文本 ${texts}（逐字符）/保精度字面量 ${precision}（文本通道 t=1 重放，逐字符判定）可判定项 ${assertable.length}，FAIL ${failed.length}` +
        `${failed.length ? '：' + sampleFail : ''}${displayTruncated.length ? '；显示通道三点对照：' + sampleDisplay : ''}`)
    },
    assertFormulaCalc() {
      if (!this.formulaRows.length) {
        this.setAssertion('formulaCalc', 'NA', '预览数据中无含 f 的公式单元格')
        return
      }
      const badVerbatim = this.formulaRows.filter(r => r.formulaReadback !== r.f)
      const undecidable = this.formulaRows.filter(r => !r.decidable)
      const mismatch = this.formulaRows.filter(r => r.decidable && !r.match)
      // 终审 P2-c：重算必须"真的发生并写回模型"，不能只看 computed===expected。
      // 反例：computed=expected 但 recalcApplied=false 且 save().v 仍是预览缓存 13——旧实现判 PASS。
      // 硬判据（逐公式）：f 原样 + recalcApplied + 存储快照 f 原样 + 存储 v===computed===expected。
      const notApplied = this.formulaRows.filter(r => r.decidable && r.recalcApplied !== true)
      const storageMissing = this.formulaRows.filter(r => r.decidable && (!r.saveSnapshot || typeof r.saveSnapshot !== 'object'))
      const storageMismatch = this.formulaRows.filter(r => r.decidable && r.saveSnapshot && typeof r.saveSnapshot === 'object' && r.saveSnapshot.v !== r.computed)
      const storageFormulaLost = this.formulaRows.filter(r => r.decidable && r.saveSnapshot && typeof r.saveSnapshot === 'object' && r.saveSnapshot.f !== r.f)
      const sample = this.formulaRows.slice(0, 3).map(r => `${r.f} → 计算 ${this.showScalar(r.computed)} / 期望 ${r.decidable ? this.showScalar(r.expected) : '不可判定'} / 存储 ${r.saveSnapshot ? this.showScalar(r.saveSnapshot.v) : '(无)'} / 已应用 ${r.recalcApplied === true}`).join('；')
      // 口径同步（整改 3a + F4a 整改）：可判定项 = 计算结果与 evalSimpleFormula 期望值双精度全等。
      // 旧实现 ok=true 只看"无 mismatch 且无 badVerbatim"，混合场景（1 个可判定通过 + 1 个
      // 不可判定）也会 PASS，属假绿。整改：存在任一不可判定项（函数名/字符串/引用缺值）→
      // 整体记「人工核对（待验）」，不得 PASS；PASS 仅当全部公式可判定、f 原样且与
      // evalSimpleFormula 期望全等；mismatch/badVerbatim>0 仍 FAIL。
      const failed = badVerbatim.length > 0 || mismatch.length > 0 ||
        notApplied.length > 0 || storageMissing.length > 0 || storageMismatch.length > 0 || storageFormulaLost.length > 0
      this.setAssertion('formulaCalc', failed ? 'FAIL' : (undecidable.length > 0 ? '人工核对' : 'PASS'),
        `公式 ${this.formulaRows.length} 个：f 原样读回不符 ${badVerbatim.length} 处；可判定 ${this.formulaRows.length - undecidable.length} 个（与 evalSimpleFormula 期望值双精度全等，不符 ${mismatch.length} 处）；` +
        `重算未应用 ${notApplied.length} 处 / 存储快照缺失 ${storageMissing.length} 处 / 存储值≠计算结果 ${storageMismatch.length} 处 / 存储公式原文丢失 ${storageFormulaLost.length} 处（PASS 要求 f 原样 + 重算已应用 + 存储 v===computed===expected），` +
        `不可判定 ${undecidable.length} 个${undecidable.length ? '（存在不可判定项，整体记人工核对待验，不得 PASS）' : ''}；示例：${sample}`)
    },
    // ---------- 保存重载 ----------
    flattenCellData(sheetsObject) {
      const out = {}
      for (const [sid, sheet] of Object.entries(sheetsObject || {})) {
        out[sid] = (sheet && sheet.cellData) || {}
      }
      return out
    },
    anchorTag(cell) {
      if (!cell) return ''
      if (cell.t === 3) return '布尔1/0'
      if (cell.t === 1 && cell.v === '001') return '前导零文本'
      if (cell.t === 2 && typeof cell.v === 'number' && !Number.isInteger(cell.v)) return '日期序列/小数'
      return ''
    },
    async saveReloadCompare() {
      const book = this.univerAPI && this.univerAPI.getActiveWorkbook()
      if (!book) {
        this.reloadError = { message: '无活动工作簿：请先执行重放' }
        return
      }
      this.reloadBusy = true
      this.reloadError = null
      try {
        // 1) 重放后首次 save：深拷贝定格基线（防 createUnit 就地规范化污染基线）。
        const firstSave = book.save()
        const firstCellData = this.safeClone(this.flattenCellData((firstSave && firstSave.sheets) || {}))
        // 2) dispose → 用保存快照重建工作簿（同配置 createUniver + createUnit，unitData 用快照）。
        this.disposeUniver()
        await this.$nextTick()
        this.buildUniver()
        const sheets = this.safeClone((firstSave && firstSave.sheets) || {})
        const sheetOrder = (firstSave && Array.isArray(firstSave.sheetOrder) && firstSave.sheetOrder.length)
          ? firstSave.sheetOrder
          : Object.keys(sheets)
        this.univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
          id: 'c3-acceptance-reload',
          name: 'C3 导入验收-保存重载',
          sheetOrder,
          sheets,
          activeSheet: sheetOrder[0]
        })
        const book2 = this.univerAPI.getActiveWorkbook()
        if (!book2) throw new Error('保存快照重建工作簿失败（getActiveWorkbook 为空）')
        // 3) 重载后再次 save，读 cellData 与首次 save 逐格比对（t 与 v 全等）。
        const secondSave = book2.save()
        const secondCellData = this.flattenCellData((secondSave && secondSave.sheets) || {})
        this.compareCellData(firstCellData, secondCellData)
        this.logRequest('保存→重载→比对', { summary: this.compareSummary })
      } catch (err) {
        this.reloadError = this.normalizeError(err)
        this.logRequest('保存重载失败（原样）', { normalized: this.reloadError })
      } finally {
        this.reloadBusy = false
      }
    },
    // 逐格比对公共实现（页内 SDK 与真实后端两个口径共用）：t 与 v 全等；保精度字面量格以
    // precisionPositions 登记的位置打"保精度字面量"锚点——此类格子经 save→（后端）→重载 后
    // 必须仍是逐字符相等的文本（t=1 + 字符串逐字符相等由 cellEqual 保证）。
    // f 差异口径分两层：页内 SDK save→重载（compareCellData）仍作信息项；真实后端重载
    // （compareBackendCellData）自 F4b 整改起为硬判据——fDiff>0 即 FAIL（公式往返必须 f 原样）。
    diffCellData(firstCellData, secondCellData) {
      const rows = []
      let total = 0
      let diffCount = 0
      let fDiffCount = 0
      let anchorCount = 0
      let precisionCount = 0
      const sheetIds = [...new Set([...Object.keys(firstCellData), ...Object.keys(secondCellData)])]
      for (const sid of sheetIds) {
        const a = firstCellData[sid] || {}
        const b = secondCellData[sid] || {}
        const rowKeys = [...new Set([...Object.keys(a), ...Object.keys(b)])]
        for (const rowKey of rowKeys) {
          const rowA = a[rowKey] || {}
          const rowB = b[rowKey] || {}
          const colKeys = [...new Set([...Object.keys(rowA), ...Object.keys(rowB)])]
          for (const colKey of colKeys) {
            const ca = rowA[colKey]
            const cb = rowB[colKey]
            total++
            const same = this.cellEqual(ca, cb)
            const fDiff = String((ca && ca.f) || '') !== String((cb && cb.f) || '')
            if (!same) diffCount++
            if (fDiff) fDiffCount++
            const isPrecision = (this.precisionPositions || new Set()).has(sid + '|' + rowKey + '|' + colKey)
            if (isPrecision) precisionCount++
            const anchor = isPrecision ? '保精度字面量' : this.anchorTag(cb || ca || {})
            if (anchor) anchorCount++
            rows.push({
              sheetId: sid,
              row: rowKey,
              col: colKey,
              t1: ca === undefined ? undefined : ca.t,
              t2: cb === undefined ? undefined : cb.t,
              v1: ca === undefined ? undefined : ca.v,
              v2: cb === undefined ? undefined : cb.v,
              f1: (ca && ca.f) || '',
              f2: (cb && cb.f) || '',
              fDiff,
              same,
              anchor
            })
          }
        }
      }
      return { rows, total, diffCount, fDiffCount, anchorCount, precisionCount }
    },
    failSamplesOf(rows) {
      return rows.filter(r => !r.same).slice(0, 3)
        .map(r => `r${r.row}c${r.col}: t/v ${this.showScalar(r.v1)}(t${r.t1 === undefined ? '-' : r.t1}) → ${this.showScalar(r.v2)}(t${r.t2 === undefined ? '-' : r.t2})`).join('；')
    },
    compareCellData(firstCellData, secondCellData) {
      const d = this.diffCellData(firstCellData, secondCellData)
      this.compareRows = d.rows
      this.compareSummary = { total: d.total, diffCount: d.diffCount, fDiffCount: d.fDiffCount, anchorCount: d.anchorCount }
      if (!d.total) {
        this.setAssertion('saveReload', 'NA', '两次 save 的 cellData 均为空，无可比对单元格')
      } else {
        this.setAssertion('saveReload', d.diffCount === 0 ? 'PASS' : 'FAIL',
          `共比对 ${d.total} 格（布尔/前导零/日期序列/保精度字面量锚点 ${d.anchorCount} 格，其中保精度字面量经 save→重载 逐字符相等 ${d.precisionCount} 格），t/v 不一致 ${d.diffCount} 处${d.diffCount ? '：' + this.failSamplesOf(d.rows) : ''}；f 不一致 ${d.fDiffCount} 处（信息项）`)
      }
    },
    // ---------- 真实后端保存/重载（整改 4，产品文档保存链路） ----------
    countCells(cellDataBySheet) {
      let total = 0
      for (const rows of Object.values(cellDataBySheet || {})) {
        for (const row of Object.values(rows || {})) total += Object.keys(row || {}).length
      }
      return total
    },
    // 真实后端保存：payload 契约与主页面 index.vue 首存一致（workbookId:null/type:'created'，
    // content = JSON.stringify(book.save())）；保存前先把页内 save() 快照定格为基线。
    async backendSave() {
      const book = this.univerAPI && this.univerAPI.getActiveWorkbook()
      if (!book || !this.replayDone) {
        this.backendSaveError = { message: '无活动工作簿或尚未重放：请先执行"重放到工作簿"' }
        return
      }
      this.backendSaveBusy = true
      this.backendSaveError = null
      try {
        const snapshot = this.safeClone(book.save())
        this.backendBaseline = this.flattenCellData((snapshot && snapshot.sheets) || {})
        const payload = {
          workbookId: null,
          // 唯一后缀：同一 requestId 反复验收会撞"文件名已存在"，每次保存必须得到独立工作簿
          // （指纹留在 name 里，日志可追溯是第几次保存）。
          name: 'C3验收-' + this.requestId + '-' + Date.now().toString(36),
          type: 'created',
          version: undefined,
          content: JSON.stringify(snapshot)
        }
        const res = await saveWorkbook(payload)
        if (!res || res.code !== 200 || !res.data || res.data.workbookId === undefined || res.data.workbookId === null) {
          throw new Error((res && res.msg) || '保存响应缺少 workbookId')
        }
        this.backendWorkbookId = res.data.workbookId
        this.backendReloadRows = []
        this.backendReloadSummary = null
        this.backendReloadError = null
        this.backendReloadAttempts = 0
        this.logRequest('真实后端保存（POST /ai/workbook/save）', { workbookId: this.backendWorkbookId, name: payload.name, baselineCells: this.countCells(this.backendBaseline) })
      } catch (err) {
        this.backendSaveError = this.normalizeError(err)
        this.logRequest('真实后端保存失败（原样）', { normalized: this.backendSaveError })
      } finally {
        this.backendSaveBusy = false
      }
    },
    // 从后端重载比对（模拟刷新）：GET /ai/workbook/{id} → 解析 data.content → cellData 与基线
    // 逐格比对（t 与 v 全等；保精度字面量必须仍是逐字符相等的文本；f 差异仅信息项）。
    // 「再查一次」重复本流程（attempts 递增），证明结果非缓存巧合。
    async backendReloadCompare(again) {
      if (!this.backendWorkbookId || !this.backendBaseline) return
      this.backendReloadBusy = true
      this.backendReloadError = null
      try {
        const res = await getWorkbook(this.backendWorkbookId)
        if (!res || res.code !== 200 || !res.data || typeof res.data.content !== 'string') {
          throw new Error((res && res.msg) || '重载响应缺少 content')
        }
        const snapshot = JSON.parse(res.data.content)
        const reloadedCellData = this.flattenCellData((snapshot && snapshot.sheets) || {})
        this.backendReloadAttempts++
        this.compareBackendCellData(reloadedCellData)
        this.logRequest(again ? '从后端重载比对（再查一次）' : '从后端重载比对（模拟刷新）', { workbookId: this.backendWorkbookId, summary: this.backendReloadSummary })
      } catch (err) {
        this.backendReloadError = this.normalizeError(err)
        this.logRequest('从后端重载比对失败（原样）', { normalized: this.backendReloadError })
      } finally {
        this.backendReloadBusy = false
      }
    },
    compareBackendCellData(reloadedCellData) {
      const d = this.diffCellData(this.backendBaseline, reloadedCellData)
      this.backendReloadRows = d.rows
      // F4b 整改：verdict 进入 summary，供面板 alert 用真实判定（而非仅 diffCount）着色。
      const verdict = (d.diffCount === 0 && d.fDiffCount === 0) ? 'PASS' : 'FAIL'
      this.backendReloadSummary = {
        total: d.total, diffCount: d.diffCount, fDiffCount: d.fDiffCount,
        anchorCount: d.anchorCount, precisionCount: d.precisionCount, attempts: this.backendReloadAttempts, verdict
      }
      if (!d.total) {
        this.setAssertion('backendSaveReload', 'NA', '后端快照 cellData 为空，无可比对单元格')
      } else {
        // F4b 整改：公式往返必须校验 f 原样——旧实现 fDiffCount 只作信息项，后端重载后公式 f
        // 丢失、缓存值不变时 PASS 属假绿。fDiffCount>0 即 FAIL，并列出 f 丢失/变更坐标
        // （diffCellData 的 f 差异明细行保留，供表格展示）。
        const fDiffSamples = d.rows.filter(r => r.fDiff).slice(0, 5)
          .map(r => `r${r.row}c${r.col}: f ${r.f1 ? "'" + r.f1 + "'" : '(丢失)'} → ${r.f2 ? "'" + r.f2 + "'" : '(丢失)'}`).join('；')
        this.setAssertion('backendSaveReload', verdict,
          `第 ${this.backendReloadAttempts} 次从后端重载（workbookId=${this.backendWorkbookId}，GET /ai/workbook/{id}）：共比对 ${d.total} 格` +
          `（布尔/前导零/日期序列/保精度字面量锚点 ${d.anchorCount} 格，保精度字面量经 SDK save→后端→重载 后逐字符相等 ${d.precisionCount} 格），` +
          `t/v 不一致 ${d.diffCount} 处${d.diffCount ? '：' + this.failSamplesOf(d.rows) : ''}；` +
          `f 不一致 ${d.fDiffCount} 处（公式往返必须 f 原样，fDiff>0 即 FAIL）${d.fDiffCount ? '：' + fDiffSamples : ''}`)
      }
    },
    cellEqual(a, b) {
      if (a === undefined && b === undefined) return true
      if (a === undefined || b === undefined) return false
      return a.t === b.t && a.v === b.v // 口径：t 与 v 全等（f 差异作为信息项另行展示）
    },
    // ---------- 导出验收日志 ----------
    exportAcceptanceLog() {
      const results = this.assertionRows.map(r => r.result)
      const log = {
        generatedAt: new Date().toISOString(),
        scope: 'C3 导入链路验收 harness（显示/计算/保存重载[页内 SDK+真实后端] 人工+自动双通道）',
        baseUrl: this.baseApi,
        tokenPresent: this.tokenPresent,
        requestId: this.requestId,
        jobId: this.jobId,
        // 字段契约（整改 1）：创建/查询响应 sourceContentHash 已对齐进本状态；预览查询参数名为
        // sourceHash（两者同值：源内容 sha256）。日志同步展示两个名字，避免字段名演进歧义。
        sourceHash: this.sourceHash,
        sourceContentHash: this.sourceHash,
        rowsRequested: this.rowsRequested,
        probeResult: this.probeResult,
        errors: {
          importError: this.importError,
          previewError: this.previewError,
          replayError: this.replayError,
          reloadError: this.reloadError,
          backendSaveError: this.backendSaveError,
          backendReloadError: this.backendReloadError
        },
        importResults: {
          create: this.safeClone(this.createResult),
          jobStatus: this.safeClone(this.jobStatusResult),
          byRequest: this.safeClone(this.byRequestResult),
          resend: this.safeClone(this.resendResult)
        },
        previewMeta: this.preview ? {
          jobId: this.preview.jobId, sourceHash: this.preview.sourceHash, sourceContentHash: this.preview.sourceHash, sheetCount: this.preview.sheetCount,
          totalRowCount: this.preview.totalRowCount, requestedRows: this.preview.requestedRows,
          returnedRows: this.preview.returnedRows, limited: this.preview.limited, note: this.preview.note
        } : null,
        replayEvidence: {
          readback: this.safeClone(this.readbackRows),
          formula: this.safeClone(this.formulaRows)
        },
        saveReload: {
          summary: this.compareSummary,
          mismatches: this.compareRows.filter(r => !r.same).slice(0, 200),
          fMismatches: this.compareRows.filter(r => r.fDiff).slice(0, 200)
        },
        // 真实后端保存/重载证据（整改 4）：产品文档保存链路，模拟刷新重载比对 + 连续两次重载留证。
        backendSaveReload: {
          workbookId: this.backendWorkbookId,
          attempts: this.backendReloadAttempts,
          summary: this.backendReloadSummary,
          mismatches: this.backendReloadRows.filter(r => !r.same).slice(0, 200),
          fMismatches: this.backendReloadRows.filter(r => r.fDiff).slice(0, 200)
        },
        assertions: this.safeClone(this.assertions),
        budgetRefusal: this.safeClone(this.budgetResult),
        allGreen: results.every(r => r === 'PASS'),
        requestLog: this.safeClone(this.requestLog)
      }
      const blob = new Blob([JSON.stringify(log, null, 2)], { type: 'application/json' })
      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url
      link.download = 'c3-acceptance-log.json'
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
      URL.revokeObjectURL(url)
      this.logRequest('导出验收日志', { file: 'c3-acceptance-log.json', allGreen: log.allGreen })
      this.$message.info('下载位置由浏览器决定：请另存到 G 盘；或改用「复制日志（存 G 盘）」')
    },
    // 证据文件落盘口径（用户要求：项目文件一律 G 盘，不写 C 盘）。
    // 浏览器下载目录不可由页面控制，故提供复制通道：复制 JSON 后粘到
    // G:/code/agentExcel/docs/acceptance/evidence-browser-logs/ 下的文件里。
    buildAcceptanceLog() {
      // 复用导出逻辑的同一份 JSON（避免两条证据口径分叉）
      const results = this.assertionRows.map(r => r.result)
      return {
        generatedAt: new Date().toISOString(),
        scope: 'C3 导入链路验收 harness（显示/计算/保存重载[页内 SDK+真实后端] 人工+自动双通道）',
        baseUrl: this.baseApi,
        tokenPresent: this.tokenPresent,
        requestId: this.requestId,
        jobId: this.jobId,
        sourceHash: this.sourceHash,
        sourceContentHash: this.sourceHash,
        rowsRequested: this.rowsRequested,
        allGreen: results.every(r => r === 'PASS'),
        assertions: this.safeClone(this.assertions),
        budgetRefusal: this.safeClone(this.budgetResult),
        replayEvidence: { readback: this.safeClone(this.readbackRows), formula: this.safeClone(this.formulaRows) },
        backendSaveReload: {
          workbookId: this.backendWorkbookId,
          attempts: this.backendReloadAttempts,
          summary: this.backendReloadSummary,
          mismatches: this.backendReloadRows.filter(r => !r.same).slice(0, 200),
          fMismatches: this.backendReloadRows.filter(r => r.fDiff).slice(0, 200)
        },
        requestLog: this.safeClone(this.requestLog)
      }
    },
    copyAcceptanceLog() {
      const text = JSON.stringify(this.buildAcceptanceLog(), null, 2)
      const done = () => {
        this.logRequest('复制验收日志（存 G 盘）', { chars: text.length })
        this.$message.success('已复制 JSON；请粘贴保存到 G:/code/agentExcel/docs/acceptance/evidence-browser-logs/ 下（勿存 C 盘）')
      }
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(done).catch(() => this.copyAcceptanceLogFallback(text, done))
        return
      }
      this.copyAcceptanceLogFallback(text, done)
    },
    copyAcceptanceLogFallback(text, done) {
      const ta = document.createElement('textarea')
      ta.value = text
      ta.style.position = 'fixed'
      ta.style.opacity = '0'
      document.body.appendChild(ta)
      ta.select()
      let ok = false
      try { ok = document.execCommand('copy') } catch (e) { ok = false }
      document.body.removeChild(ta)
      if (ok) done()
      else this.$message.error('复制失败：请点「导出验收日志」并另存到 G 盘（勿用默认下载目录）')
    }
  }
}
</script>

<style scoped>
.c3-accept .panel {
  margin-bottom: 12px;
}
.panel-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
}
.mb8 {
  margin-bottom: 8px;
}
.ml8 {
  margin-left: 8px;
}
.mono {
  font-family: Consolas, Monaco, 'Courier New', monospace;
}
.small {
  font-size: 12px;
  word-break: break-all;
}
.file-name {
  margin-left: 8px;
  color: #606266;
}
.raw-json {
  max-height: 240px;
  overflow: auto;
  margin: 0;
  padding: 8px;
  background: #f5f7fa;
  border: 1px solid #ebeef5;
  border-radius: 4px;
  font-family: Consolas, Monaco, 'Courier New', monospace;
  font-size: 12px;
  white-space: pre-wrap;
  word-break: break-all;
}
.precision-literal {
  border: 1px solid #e6a23c;
  background: #fdf6ec;
  padding: 1px 4px;
  border-radius: 2px;
  color: #b88230;
}
.univer-host {
  height: 460px;
  border: 1px solid #dcdfe6;
  border-radius: 4px;
  margin-bottom: 10px;
  overflow: hidden;
}
.section-title {
  font-weight: 600;
  font-size: 13px;
  color: #303133;
  margin: 8px 0 6px;
}
.hint {
  font-size: 12px;
  color: #909399;
  line-height: 1.6;
  margin-top: 6px;
}
</style>
