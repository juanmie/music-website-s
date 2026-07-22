<template>
  <el-dialog title="确认删除" v-model="visible" width="400px">
    <span>确定要删除吗？</span>
    <template #footer>
      <el-button @click="visible = false">取消</el-button>
      <el-button type="danger" @click="$emit('confirm')">确定</el-button>
    </template>
  </el-dialog>
</template>

<script lang="ts">
import { defineComponent, ref, watch } from 'vue'

export default defineComponent({
  props: {
    delVisible: { type: Boolean, default: false }
  },
  emits: ['confirm', 'cancelRow'],
  setup(props, { emit }) {
    const visible = ref(false)
    watch(() => props.delVisible, (val) => { visible.value = val })
    watch(visible, (val) => { if (!val) emit('cancelRow', false) })
    return { visible }
  }
})
</script>
