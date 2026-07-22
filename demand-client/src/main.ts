import { createApp } from 'vue'
import ElementPlus from 'element-plus'
import App from './App.vue'
import router from './router'
import store from './store'
import 'element-plus/dist/index.css'
import './assets/css/main.css'

import { Store } from 'vuex'
declare module '@vue/runtime-core' {
  interface State {
    count: number
  }
  interface ComponentCustomProperties {
    $store: Store<State>
  }
}

createApp(App).use(store).use(router as any).use(ElementPlus as any).mount('#app')
