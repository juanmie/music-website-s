import { get, getLong, post, download, postBlob } from './request'

const HttpManager = {
    // =======================> 需求评估 API
    // 获取所有需求
    getAllDemand: () => get(`demandEvaluation`),
    // 获取需求详情
    getDemandOfId: (id) => get(`demandEvaluation/detail?id=${id}`),
    // 添加需求
    addDemand: (data) => post(`demandEvaluation/add`, data),
    // 更新需求
    updateDemand: (data) => post(`demandEvaluation/update`, data),
    // 删除需求
    deleteDemand: (id) => get(`demandEvaluation/delete?id=${id}`),
    // 导出需求清单 + API匹配结果（POST 传入需求列表）
    exportWithApiMatch: (data) => postBlob(`demandEvaluation/exportWithApiMatch`, data),

    // =======================> AI 分析 API（使用长超时 5 分钟）
    // AI 自动分类
    aiClassify: (data) => post(`ai/classify`, data),
    // AI 工时估算
    aiEstimate: (data) => post(`ai/estimate`, data),
    // AI 智能汇总
    aiSummary: () => getLong(`ai/summary`),
    // AI 优化建议
    aiOptimize: (data) => post(`ai/optimize`, data),
    // AI 批量分析（POST 传入需求列表）
    aiBatchAnalysis: (data) => post(`ai/batchAnalysis`, data, 300000),
    // AI 增强导出（POST 传入需求列表）
    aiExportEnhanced: (data) => post(`ai/exportEnhanced`, data, 300000),
}

export { HttpManager }
