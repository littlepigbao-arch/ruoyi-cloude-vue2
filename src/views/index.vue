<template>
  <div class="app-container home">
    <el-row :gutter="20">
      <el-col :span="24">
        <el-card class="welcome-card" shadow="never">
          <h2>系统工程软件管理平台</h2>
          <p>欢迎使用本平台，请从左侧菜单或下方功能入口开始使用。</p>
        </el-card>
      </el-col>
    </el-row>

    <el-row :gutter="20" class="feature-row">
      <el-col :xs="24" :sm="12" :md="8">
        <el-card class="feature-card" shadow="hover" @click.native="goTarget('/agent-report')">
          <div class="feature-icon"><i class="el-icon-s-grid"></i></div>
          <div class="feature-title">智能报表</div>
          <div class="feature-desc">在线表格编辑、Excel 导入导出与 AI 智能助手</div>
        </el-card>
      </el-col>
    </el-row>
  </div>
</template>

<script>
export default {
  name: "Index",
  methods: {
    goTarget(path) {
      this.$router.push(path)
    }
  }
}
</script>

<style scoped lang="scss">
.home {
  font-family: "open sans", "Helvetica Neue", Helvetica, Arial, sans-serif;
  color: #676a6c;
}

.welcome-card {
  h2 {
    margin-top: 0;
    margin-bottom: 12px;
    color: #303133;
  }

  p {
    margin: 0;
    font-size: 14px;
    color: #909399;
  }
}

.feature-row {
  margin-top: 20px;
}

.feature-card {
  cursor: pointer;

  .feature-icon {
    font-size: 36px;
    color: #409eff;
    margin-bottom: 12px;
  }

  .feature-title {
    font-size: 16px;
    font-weight: 600;
    color: #303133;
    margin-bottom: 8px;
  }

  .feature-desc {
    font-size: 13px;
    color: #909399;
    line-height: 1.5;
  }
}
</style>
