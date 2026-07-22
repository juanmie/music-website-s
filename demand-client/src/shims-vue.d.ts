declare interface ResponseBody {
    code: number
    message: string
    type: string
    success: boolean
    data: any
}

declare module '*.vue' {
    import type { DefineComponent } from 'vue'
    const component: DefineComponent<{}, {}, any>
    export default component
}

declare var process: {
    env: {
        NODE_ENV: string
        BASE_URL: string
        NODE_HOST: string
    }
}
declare interface ResponseBody {
    code: number
    message: string
    type: string
    success: boolean
    data: any
}
