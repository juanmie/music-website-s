<template>
  <div class="sidebar">
    <el-menu
      class="sidebar-el-menu"
      background-color="#ffffff"
      active-text-color="#30a4fc"
      default-active="/Home/Demand"
      router
      :collapse="isCollapse"
    >
      <el-menu-item index="/Home/Demand">
        <el-icon><List /></el-icon>
        <span>需求管理</span>
      </el-menu-item>
    </el-menu>
  </div>
</template>

<script lang="ts" setup>
import { computed } from 'vue'
import { useStore } from 'vuex'
import { List } from '@element-plus/icons-vue'

const store = useStore()
const isCollapse = computed(() => store.getters.isCollapse)
</script>

<style scoped>
.sidebar {
  display: block;
  position: fixed;
  left: 0;
  top: 60px;
  bottom: 0;
  overflow-y: auto;
}

.sidebar::-webkit-scrollbar {
  width: 0;
}

.sidebar-el-menu:not(.el-menu--collapse) {
  width: 200px;
  height: 100%;
}
</style>
