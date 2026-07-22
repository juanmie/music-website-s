import { createStore } from 'vuex'

export default createStore({
  state: {
    isCollapse: false
  },
  mutations: {
    setIsCollapse(state, val) {
      state.isCollapse = val
    }
  },
  getters: {
    isCollapse(state) {
      return state.isCollapse
    }
  }
})
