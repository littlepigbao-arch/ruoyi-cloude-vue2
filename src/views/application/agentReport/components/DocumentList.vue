<template>
  <div class="doc-sidebar">
    <div class="doc-sidebar__header">
      <span class="doc-sidebar__title">
        <i class="el-icon-folder-opened"></i> 文档列表
      </span>
      <el-tooltip content="刷新列表" placement="bottom">
        <el-button
          type="text"
          icon="el-icon-refresh"
          size="mini"
          class="doc-sidebar__refresh"
          @click="$emit('refresh')"
        />
      </el-tooltip>
    </div>

    <div class="doc-sidebar__body">
      <div v-if="documents.length === 0" class="doc-sidebar__empty">
        暂无文档
      </div>

      <div
        v-for="doc in documents"
        :key="doc.workbookId"
        class="doc-item"
        :class="{ 'doc-item--active': doc.workbookId === activeId }"
        @click="$emit('select', doc)"
      >
        <div class="doc-item__main">
          <div class="doc-item__name" :title="doc.name">{{ doc.name }}</div>
          <div class="doc-item__meta">
            <el-tag size="mini" :type="doc.type === 'imported' ? 'warning' : 'success'">
              {{ doc.type === 'imported' ? '导入' : '新建' }}
            </el-tag>
            <span class="doc-item__time">{{ formatTime(doc.updateTime || doc.createTime) }}</span>
          </div>
        </div>
        <el-tooltip content="重命名" placement="top">
          <el-button
            type="text"
            icon="el-icon-edit"
            size="mini"
            class="doc-item__rename"
            @click.stop="$emit('rename', doc)"
          />
        </el-tooltip>
        <el-tooltip content="删除" placement="top">
          <el-button
            type="text"
            icon="el-icon-delete"
            size="mini"
            class="doc-item__del"
            @click.stop="$emit('delete', doc)"
          />
        </el-tooltip>
      </div>
    </div>
  </div>
</template>

<script>
export default {
  name: 'DocumentList',
  props: {
    documents: { type: Array, default: () => [] },
    activeId: { type: [Number, String], default: null }
  },
  methods: {
    formatTime(value) {
      if (!value) return ''
      let d
      if (typeof value === 'string' || typeof value === 'number') {
        d = new Date(value)
      } else {
        d = value
      }
      if (!d || isNaN(d.getTime())) return ''
      const y = d.getFullYear()
      const m = String(d.getMonth() + 1).padStart(2, '0')
      const day = String(d.getDate()).padStart(2, '0')
      const h = String(d.getHours()).padStart(2, '0')
      const min = String(d.getMinutes()).padStart(2, '0')
      return y + '-' + m + '-' + day + ' ' + h + ':' + min
    }
  }
}
</script>

<style lang="scss" scoped>
.doc-sidebar {
  width: 230px;
  flex-shrink: 0;
  display: flex;
  flex-direction: column;
  border: 1px solid #dcdfe6;
  border-radius: 4px;
  background: #fff;
  overflow: hidden;

  &__header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 8px 12px;
    border-bottom: 1px solid #ebeef5;
    background: #fafafa;
  }

  &__title {
    font-size: 14px;
    font-weight: 600;
    color: #303133;

    i {
      color: #409eff;
    }
  }

  &__refresh {
    padding: 3px;
    color: #909399;
  }

  &__body {
    flex: 1;
    overflow-y: auto;
    padding: 6px;
  }

  &__empty {
    padding: 24px 0;
    text-align: center;
    color: #c0c4cc;
    font-size: 13px;
  }
}

.doc-item {
  display: flex;
  align-items: center;
  padding: 8px 10px;
  margin-bottom: 4px;
  border-radius: 4px;
  cursor: pointer;
  transition: background 0.15s;

  &:hover {
    background: #f5f7fa;

    .doc-item__del,
    .doc-item__rename {
      opacity: 1;
    }
  }

  &--active {
    background: #ecf5ff;
    border-left: 3px solid #409eff;
  }

  &__main {
    flex: 1;
    min-width: 0;
  }

  &__name {
    font-size: 13px;
    color: #303133;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    margin-bottom: 4px;
  }

  &__meta {
    display: flex;
    align-items: center;
    gap: 6px;
  }

  &__time {
    font-size: 12px;
    color: #c0c4cc;
  }

  &__del {
    padding: 3px;
    color: #f56c6c;
    opacity: 0;
    transition: opacity 0.15s;
  }

  &__rename {
    padding: 3px;
    color: #409eff;
    opacity: 0;
    transition: opacity 0.15s;
  }
}
</style>
