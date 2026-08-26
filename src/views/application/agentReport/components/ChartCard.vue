<template>
  <div ref="chart" class="chart-card" :style="{ height }" />
</template>

<script>
import * as echarts from 'echarts'

/**
 * 轻量 ECharts 卡片：接收 option 渲染，支持窗口 resize，销毁时释放实例
 */
export default {
  name: 'ChartCard',
  props: {
    option: { type: Object, default: null },
    height: { type: String, default: '260px' }
  },
  data() {
    return {
      chart: null
    }
  },
  watch: {
    option: {
      handler(val) {
        if (val) this.$nextTick(() => this.renderChart(val))
      },
      deep: true
    }
  },
  mounted() {
    this.$nextTick(() => {
      this.chart = echarts.init(this.$el)
      if (this.option) this.renderChart(this.option)
    })
    window.addEventListener('resize', this.onResize)
  },
  beforeDestroy() {
    window.removeEventListener('resize', this.onResize)
    if (this.chart) {
      this.chart.dispose()
      this.chart = null
    }
  },
  methods: {
    renderChart(option) {
      if (!this.chart) this.chart = echarts.init(this.$el)
      this.chart.setOption(option, true)
    },
    onResize() {
      if (this.chart) this.chart.resize()
    }
  }
}
</script>

<style lang="scss" scoped>
.chart-card {
  width: 100%;
  min-height: 200px;
}
</style>
