import axios from 'axios'

const BASE_URL = process.env.NODE_HOST

axios.defaults.timeout = 10000
axios.defaults.withCredentials = true
axios.defaults.baseURL = BASE_URL
axios.defaults.headers.post['Content-Type'] = 'application/json;charset=UTF-8'

axios.interceptors.response.use(
  (response) => {
    if (response.status === 200) {
      return Promise.resolve(response)
    }
    return Promise.reject(response)
  },
  (error) => {
    return Promise.reject(error)
  }
)

export function getBaseURL() {
  return BASE_URL
}

export function get(url: string, params?: object) {
  return new Promise((resolve, reject) => {
    axios.get(url, params).then(
      (response) => resolve(response.data),
      (error) => reject(error)
    )
  })
}

/**
 * 长时间请求（用于 AI 分析等耗时操作）
 * @param timeout 超时时间（毫秒），默认 5 分钟
 */
export function getLong(url: string, params?: object, timeout = 300000) {
  return new Promise((resolve, reject) => {
    axios.get(url, { ...params, timeout }).then(
      (response) => resolve(response.data),
      (error) => reject(error)
    )
  })
}

export function post(url: string, data = {}, timeout?: number) {
  return new Promise((resolve, reject) => {
    const config: any = {}
    if (timeout) config.timeout = timeout
    axios.post(url, data, config).then(
      (response) => resolve(response.data),
      (error) => reject(error)
    )
  })
}

export function deletes(url: string, data = {}) {
  return new Promise((resolve, reject) => {
    axios.delete(url, data).then(
      (response) => resolve(response.data),
      (error) => reject(error)
    )
  })
}

/**
 * 下载文件（返回 Blob）
 * @param timeout 超时时间（毫秒），默认 3 分钟
 */
export function download(url: string, params?: object, timeout = 180000) {
  return new Promise((resolve, reject) => {
    axios.get(url, { ...params, responseType: 'blob', timeout }).then(
      (response) => resolve(response),
      (error) => reject(error)
    )
  })
}

/**
 * POST 下载文件（返回 Blob，用于传递请求体）
 * @param timeout 超时时间（毫秒），默认 3 分钟
 */
export function postBlob(url: string, data: any = {}, timeout = 180000) {
  return new Promise((resolve, reject) => {
    axios.post(url, data, { responseType: 'blob', timeout }).then(
      (response) => {
        // 检查是否是错误响应（服务器返回的 blob 可能是 JSON 错误）
        const contentType = String(response.headers['content-type'] || '')
        if (contentType.includes('application/json')) {
          // 读取 blob 内容作为错误信息
          const reader = new FileReader()
          reader.onload = () => {
            reject(new Error(reader.result as string))
          }
          reader.readAsText(response.data)
        } else {
          resolve(response)
        }
      },
      (error) => {
        if (error.response && error.response.data instanceof Blob) {
          const reader = new FileReader()
          reader.onload = () => {
            reject(new Error(reader.result as string))
          }
          reader.readAsText(error.response.data)
        } else {
          reject(error)
        }
      }
    )
  })
}
