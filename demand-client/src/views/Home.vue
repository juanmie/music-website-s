<template>
  <div>
    <yin-header></yin-header>
    <yin-aside></yin-aside>
    <div class="content-box" :class="{ 'content-collapse': isCollapse }">
      <router-view></router-view>
    </div>
  </div>
</template>

<script lang="ts">
import { computed, defineComponent } from 'vue'
import { useStore } from 'vuex'
import YinHeader from '@/components/layouts/YinHeader.vue'
import YinAside from '@/components/layouts/YinAside.vue'

export default defineComponent({
  name: 'YinHome',
  components: { YinHeader, YinAside },
  setup() {
    const store = useStore()
    const isCollapse = computed(() => store.getters.isCollapse)
    return { isCollapse }
  }
})
</script>

<style scoped>
.content-box {
  position: fixed;
  left: 200px;
  right: 0;
  top: 60px;
  bottom: 0;
  overflow-y: auto;
  transition: left 0.3s ease-in-out;
  padding: 20px;
  background-color: #f5f5f5;
}

.content-collapse {
  left: 64px;
}
</style>
