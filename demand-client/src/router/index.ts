import { createRouter, createWebHistory, RouteRecordRaw } from 'vue-router'

const routes: Array<RouteRecordRaw> = [
  {
    path: '/Home',
    component: () => import('@/views/Home.vue'),
    children: [
      {
        path: 'Demand',
        component: () => import('@/views/DemandPage.vue'),
        meta: { title: '需求管理' }
      }
    ]
  },
  {
    path: '/',
    redirect: '/Home/Demand'
  }
]

const router = createRouter({
  history: createWebHistory(process.env.BASE_URL),
  routes
})

export default router
