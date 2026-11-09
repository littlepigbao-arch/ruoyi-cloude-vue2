<template>
  <div class="probe">
    <h3>Univer 公式重算策略探针（真实 UI 插件栈）</h3>
    <p class="hint">
      本页在真实浏览器 UI 栈（createUniver + UniverSheetsCorePreset）中逐个策略实证"=A1*B1 能否真的算出 -5627.25"。
      每次策略都新建独立 Univer 实例，结束后销毁；结果以 JSON 打印，可复制留证。
    </p>
    <el-form inline size="mini">
      <el-form-item>
        <el-button type="primary" :loading="running" @click="runAll">运行全部策略</el-button>
      </el-form-item>
      <el-form-item>
        <el-button v-for="s in strategyList" :key="s.name" :disabled="running" @click="runOne(s.name)">{{ s.name }}</el-button>
      </el-form-item>
      <el-form-item>
        <el-button :disabled="running" @click="clearResults">清空结果</el-button>
      </el-form-item>
    </el-form>
    <div ref="host" class="univer-host" />
    <el-alert v-if="error" type="error" :closable="false" :title="'探针错误（原样）'" class="mb8">
      <div class="mono small">{{ error }}</div>
    </el-alert>
    <pre class="mono small result">{{ JSON.stringify(results, null, 2) }}</pre>
  </div>
</template>

<script>
import { createUniver, LocaleType, mergeLocales, defaultTheme, UniverInstanceType } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import '@univerjs/presets/lib/styles/preset-sheets-core.css'
import '@univerjs/sheets/lib/facade.js'

const WAIT_APPLY_MS = 6000

export default {
  name: 'FormulaProbe',
  data() {
    return {
      running: false,
      error: null,
      results: [],
      strategyList: [
        { name: 'S1_current_immediate' },
        { name: 'S2_settle_then_facade' },
        { name: 'S3_settle_1000_then_facade' },
        { name: 'S4_setFormula_then_facade' },
        { name: 'S5_formulaFirst_then_facade' },
        { name: 'S6_calculationEnd_wait' },
        { name: 'S7_facade_twice' },
        { name: 'S8_no_trigger_selfcalc' },
        { name: 'S9_setFormula_no_trigger_autorecalc' },
        { name: 'S10_pageStyle_register_in_calc_lane' }
      ]
    }
  },
  created() {
    // 结果持久化：dev 热重载/误点重跑都不丢证据。
    try {
      const saved = sessionStorage.getItem('c3-formula-probe')
      if (saved) this.results = JSON.parse(saved)
    } catch (e) { /* ignore */ }
  },
  beforeDestroy() { this.dispose() },
  methods: {
    dispose() {
      if (this.univer) {
        try { this.univer.dispose() } catch (e) { /* ignore */ }
      }
      this.univer = null
      this.univerAPI = null
    },
    build() {
      this.dispose()
      const zhCNLocales = mergeLocales({}, UniverPresetSheetsCoreZhCN.default || UniverPresetSheetsCoreZhCN)
      const { univer, univerAPI } = createUniver({
        locale: LocaleType.ZH_CN,
        locales: { [LocaleType.ZH_CN]: zhCNLocales },
        theme: defaultTheme,
        presets: [UniverSheetsCorePreset({ container: this.$refs.host })]
      })
      this.univer = univer
      this.univerAPI = univerAPI
    },
    // 创建 unit 并等待"当前工作簿可读"——记录等待耗时（生命周期证据）。
    async createBookAndWait(id) {
      const t0 = Date.now()
      this.univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
        id,
        name: id,
        sheetOrder: ['s'],
        sheets: { s: { id: 's', name: 'T', rowCount: 20, columnCount: 10 } }
      })
      let book = null
      let polls = 0
      while (Date.now() - t0 < 3000) {
        book = this.univerAPI ? this.univerAPI.getActiveWorkbook() : null
        if (book) break
        polls++
        await this.sleep(20)
      }
      return { book, waitMs: Date.now() - t0, polls }
    },
    sleep(ms) { return new Promise(r => setTimeout(r, ms)) },
    async nextFrame() {
      await new Promise(r => (typeof requestAnimationFrame === 'function' ? requestAnimationFrame(() => r()) : setTimeout(r, 16)))
    },
    // 写引用 + 公式缓存值（页面现状口径）。
    writeSeed(sheet) {
      sheet.getRange(0, 0).setValue({ v: 123, t: 2 })
      sheet.getRange(0, 1).setValue({ v: -45.75, t: 2 })
      sheet.getRange(4, 0).setValue({ v: 13, t: 2, f: '=A1*B1' })
    },
    // 有界等待：显示值离开缓存值（应用落地信号）；返回是否离开与耗时。samples 可选，按 100ms 采样留证。
    async waitLeaveCache(range, cached, budgetMs, samples) {
      const t0 = Date.now()
      while (Date.now() - t0 < budgetMs) {
        const v = range.getValues()[0][0]
        if (samples) samples.push({ ms: Date.now() - t0, v })
        if (v !== cached) return { left: true, ms: Date.now() - t0, v }
        await this.sleep(100)
      }
      return { left: false, ms: Date.now() - t0, v: range.getValues()[0][0] }
    },
    // 单个策略执行体：seed 方式 + 触发方式 + 等待方式。
    async runStrategy(name, opts) {
      const rec = { strategy: name, notes: opts.note || '' }
      this.build()
      const created = await this.createBookAndWait('probe-' + name)
      rec.activeWorkbookWaitMs = created.waitMs
      rec.activeWorkbookPolls = created.polls
      rec.activeWorkbookNullOnFirstRead = created.book === null
      if (!created.book) { rec.error = 'getActiveWorkbook 始终为空'; this.record(rec); return rec }
      const sheet = created.book.getActiveSheet()
      const formulaRange = sheet.getRange(4, 0)
      const fe = this.univerAPI && typeof this.univerAPI.getFormula === 'function' ? this.univerAPI.getFormula() : null
      rec.formulaFacadePresent = !!fe
      rec.hasExecuteCalculation = !!(fe && typeof fe.executeCalculation === 'function')
      rec.hasCalculationEnd = !!(fe && typeof fe.calculationEnd === 'function')
      rec.hasSetFormulaOnRange = typeof formulaRange.setFormula === 'function'
      try {
        if (opts.seed === 'setValueWithF') {
          this.writeSeed(sheet)
        } else if (opts.seed === 'setFormula') {
          sheet.getRange(0, 0).setValue({ v: 123, t: 2 })
          sheet.getRange(0, 1).setValue({ v: -45.75, t: 2 })
          formulaRange.setValue(13)
          formulaRange.setFormula('=A1*B1')
        } else if (opts.seed === 'pageStyleRegister') {
          // 修复后页面口径：重放阶段只写 setValue({v,t,f}) 保留缓存；计算 lane 内 setFormula 注册。
          this.writeSeed(sheet)
          rec.preCalcStorage = (() => {
            try {
              const snap = created.book.save()
              const cd = snap && snap.sheets && snap.sheets.s && snap.sheets.s.cellData
              return cd && cd[4] ? cd[4][0] : null
            } catch (e) { return null }
          })()
          formulaRange.setFormula('=A1*B1')
          rec.registeredVia = 'setFormula'
        } else if (opts.seed === 'formulaFirst') {
          formulaRange.setValue(13)
          formulaRange.setFormula('=A1*B1')
          sheet.getRange(0, 0).setValue({ v: 123, t: 2 })
          sheet.getRange(0, 1).setValue({ v: -45.75, t: 2 })
        }
        rec.cachedBeforeTrigger = formulaRange.getValues()[0][0]
        if (opts.settleMs) await this.sleep(opts.settleMs)
        if (opts.frames) { for (let i = 0; i < opts.frames; i++) await this.nextFrame() }

        if (opts.trigger === 'facade') {
          fe.executeCalculation()
        } else if (opts.trigger === 'facadeTwice') {
          fe.executeCalculation()
          await this.sleep(opts.gapMs || 150)
          fe.executeCalculation()
        } else if (opts.trigger === 'calculationEnd') {
          const endPromise = new Promise(resolve => {
            let done = false
            const finish = v => { if (!done) { done = true; resolve(v) } }
            try { fe.calculationEnd(state => finish(state)) } catch (e) { finish('error:' + e.message) }
            setTimeout(() => finish('timeout'), opts.endBudgetMs || 5000)
          })
          fe.executeCalculation()
          rec.calculationEndState = await endPromise
        }
        // 无触发策略：只等待自算
        const samples = opts.samples ? [] : undefined
        const waited = await this.waitLeaveCache(formulaRange, opts.leaveFrom !== undefined ? opts.leaveFrom : rec.cachedBeforeTrigger, opts.waitBudgetMs || WAIT_APPLY_MS, samples)
        if (samples) rec.samples = samples
        rec.leftCache = waited.left
        rec.waitMs = waited.ms
        rec.display = formulaRange.getValues()[0][0]
        const snap = created.book.save()
        rec.storage = snap && snap.sheets && snap.sheets.s && snap.sheets.s.cellData ? snap.sheets.s.cellData[4][0] : null
        rec.formulaReadback = typeof formulaRange.getFormula === 'function' ? formulaRange.getFormula() : null
        rec.expectMatch = rec.display === -5627.25
      } catch (e) {
        rec.error = String((e && e.message) || e)
      }
      this.record(rec)
      return rec
    },
    strategyOptions() {
      return {
        S1_current_immediate: { seed: 'setValueWithF', trigger: 'facade', note: '页面现状：setValue({v,t,f}) 后立刻 executeCalculation' },
        S2_settle_then_facade: { seed: 'setValueWithF', frames: 2, settleMs: 250, trigger: 'facade', note: '写完后等 2 帧 + 250ms 再触发' },
        S3_settle_1000_then_facade: { seed: 'setValueWithF', settleMs: 1000, trigger: 'facade', note: '写完后等 1000ms 再触发' },
        S4_setFormula_then_facade: { seed: 'setFormula', frames: 2, settleMs: 100, trigger: 'facade', note: 'setValue(13) + setFormula(=A1*B1) 后触发' },
        S5_formulaFirst_then_facade: { seed: 'formulaFirst', settleMs: 100, trigger: 'facade', note: '先写公式再写引用值' },
        S6_calculationEnd_wait: { seed: 'setValueWithF', settleMs: 100, trigger: 'calculationEnd', note: '用 calculationEnd 回调判定完成' },
        S7_facade_twice: { seed: 'setValueWithF', settleMs: 100, trigger: 'facadeTwice', gapMs: 200, note: '触发两次，间隔 200ms' },
        S8_no_trigger_selfcalc: { seed: 'setValueWithF', trigger: 'none', note: '不调用任何触发，只观察自算' },
        S9_setFormula_no_trigger_autorecalc: { seed: 'setFormula', trigger: 'none', waitBudgetMs: 3000, samples: true, note: 'setFormula 写入后不做任何显式触发，采样观察是否自动重算（导入不重算守卫的关键事实）' },
        S10_pageStyle_register_in_calc_lane: { seed: 'pageStyleRegister', trigger: 'facade', note: '修复后页面口径：重放 setValue({v,t,f}) 保留缓存 → 计算 lane setFormula 注册 → 触发并等待真实结果' }
      }
    },
    // 记录 + 持久化（热重载/误重跑不丢证据）；同策略重跑覆盖旧记录。
    record(rec) {
      const kept = this.results.filter(r => r.strategy !== rec.strategy)
      kept.push(rec)
      kept.sort((a, b) => String(a.strategy).localeCompare(String(b.strategy)))
      this.results = kept
      try { sessionStorage.setItem('c3-formula-probe', JSON.stringify(this.results)) } catch (e) { /* ignore */ }
    },
    clearResults() {
      this.results = []
      try { sessionStorage.removeItem('c3-formula-probe') } catch (e) { /* ignore */ }
    },
    // 单策略执行（重入保护：并发跑会互相 dispose 实例，结果不可信）。
    async runOne(name) {
      if (this.running) return
      this.running = true
      this.error = null
      try {
        await this.runStrategy(name, this.strategyOptions()[name] || {})
      } catch (e) {
        this.error = String((e && e.stack) || e)
      } finally {
        this.dispose()
        this.running = false
      }
    },
    async runAll() {
      if (this.running) return // 重入保护
      this.running = true
      this.error = null
      try {
        const opts = this.strategyOptions()
        for (const s of this.strategyList) {
          await this.runStrategy(s.name, opts[s.name] || {})
        }
      } catch (e) {
        this.error = String((e && e.stack) || e)
      } finally {
        this.dispose()
        this.running = false
      }
    }
  }
}
</script>

<style scoped>
.probe { padding: 12px; }
.hint { color: #666; font-size: 12px; }
.univer-host { height: 320px; border: 1px dashed #ccc; margin: 8px 0; }
.result { background: #f5f5f5; padding: 8px; max-height: 380px; overflow: auto; }
.mono { font-family: Consolas, Monaco, monospace; }
.small { font-size: 12px; }
.mb8 { margin-bottom: 8px; }
</style>
