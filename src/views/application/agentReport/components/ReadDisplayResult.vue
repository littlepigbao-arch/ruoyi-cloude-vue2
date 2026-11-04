<template>
  <div class="read-display-result">
    <chart-card v-if="result.chartOption" :option="result.chartOption" height="280px" />
    <template v-else-if="grid">
      <p>{{ result.formulas ? '公式' : '值' }} {{ result.a1 || '' }} · {{ grid.length }} 行，{{ columnCount }} 列<span v-if="grid.length > 8 || columnCount > 16">（仅预览前 8 行、16 列）</span></p>
      <div class="read-grid"><table><tr v-for="(row, i) in grid.slice(0, 8)" :key="i"><td v-for="(cell, j) in row.slice(0, 16)" :key="j">{{ cellText(cell) }}</td></tr></table></div>
    </template>
    <template v-else-if="Object.prototype.hasOwnProperty.call(result, 'selection')">
      <p>{{ result.capturedAtStart ? '启动时选区' : '当前选区' }}：{{ result.selection && result.selection.a1 || '(无)' }}</p>
      <p v-if="result.activeCell">活动单元格：{{ result.activeCell.a1 }}</p>
    </template>
    <template v-else-if="Array.isArray(result.merges)">
      <p>合并区域 {{ result.merges.length }} 个：{{ result.merges.slice(0, 20).map(m => m.a1 || '(无坐标)').join('、') || '(无)' }}<span v-if="result.merges.length > 20">（仅显示前 20 个）</span></p>
    </template>
    <template v-else-if="typeof result.hasFilter === 'boolean'">
      <p>{{ result.hasFilter ? '筛选范围：' + (result.range || '(未知)') : '未设置筛选' }}</p>
      <template v-if="result.hasFilter">
        <p>筛选条件 {{ (result.columns || []).length }} 列</p>
        <pre v-for="(column, i) in (result.columns || []).slice(0, 20)" :key="i">第 {{ column.column + 1 }} 列：{{ JSON.stringify(column.criteria) }}</pre>
        <p v-if="(result.columns || []).length > 20">仅显示前 20 列条件。</p>
        <p>被筛选隐藏的行（共 {{ (result.filteredOutRows || []).length }} 行）：{{ (result.filteredOutRows || []).slice(0, 50).map(row => row + 1).join('、') || '(无)' }}<span v-if="(result.filteredOutRows || []).length > 50">（仅显示前 50 行）</span></p>
      </template>
    </template>
  </div>
</template>
<script>
import ChartCard from './ChartCard.vue'
export default {
  name: 'ReadDisplayResult',
  components: { ChartCard },
  props: { result: { type: Object, required: true } },
  computed: {
    grid() { return Array.isArray(this.result.values) ? this.result.values : Array.isArray(this.result.formulas) ? this.result.formulas : null },
    columnCount() { return this.grid ? this.grid.reduce((max, row) => Math.max(max, row.length), 0) : 0 }
  },
  methods: { cellText(value) { return value == null ? '' : String(value) } }
}
</script>
<style scoped>
.read-display-result { min-width: 0; width: 100%; }
.read-display-result p { margin: 4px 0; }
.read-grid { max-width: 100%; overflow: auto; }
.read-grid table { border-collapse: collapse; }
.read-grid td { border: 1px solid #dce4ee; padding: 3px 6px; white-space: pre-wrap; min-width: 32px; }
.read-display-result pre { white-space: pre-wrap; overflow-wrap: anywhere; }
</style>
