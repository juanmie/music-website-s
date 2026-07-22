package com.example.yin.controller;

import com.example.yin.common.R;
import com.example.yin.model.domain.DemandEvaluation;
import com.example.yin.service.AiAnalysisService;
import com.example.yin.service.ApiMatchService;
import com.example.yin.service.DemandEvaluationService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.util.List;
import java.util.Map;

/**
 * AI 分析控制器
 * 提供基于 Kimi API 的智能匹配功能
 */
@RestController
@RequestMapping("/ai")
public class AiAnalysisController {

    @Autowired
    private AiAnalysisService aiAnalysisService;

    @Autowired
    private DemandEvaluationService demandEvaluationService;

    @Autowired
    private ApiMatchService apiMatchService;

    /**
     * 获取所有需求列表（内部方法）
     */
    @SuppressWarnings("unchecked")
    private List<DemandEvaluation> getDemandList() {
        R result = demandEvaluationService.getAllDemand();
        return (List<DemandEvaluation>) result.getData();
    }

    /**
     * AI 智能匹配 - 基于需求清单与接口清单进行智能语义匹配
     * @param demands 前端传入的需求列表（若为空则取全部）
     */
    @PostMapping("/batchAnalysis")
    public R batchAnalysis(@RequestBody(required = false) List<DemandEvaluation> demands) {
        try {
            if (demands == null || demands.isEmpty()) {
                demands = getDemandList();
            }
            List<Map<String, String>> interfaceList = apiMatchService.getInterfaceList();
            Map<String, Object> result = aiAnalysisService.batchAnalysis(demands, interfaceList);
            return R.success("AI 智能匹配成功", result);
        } catch (Exception e) {
            return R.fatal("AI 智能匹配失败: " + e.getMessage());
        }
    }

    /**
     * AI 智能匹配导出
     * @param demands 前端传入的需求列表（若为空则取全部）
     */
    @PostMapping("/exportEnhanced")
    public R exportEnhanced(@RequestBody(required = false) List<DemandEvaluation> demands) {
        try {
            if (demands == null || demands.isEmpty()) {
                demands = getDemandList();
            }
            List<Map<String, String>> interfaceList = apiMatchService.getInterfaceList();
            Map<String, Object> result = aiAnalysisService.batchAnalysis(demands, interfaceList);
            
            // 添加导出时间戳
            result.put("exportTime", System.currentTimeMillis());
            result.put("totalDemands", demands.size());
            
            return R.success("AI 智能匹配导出成功", result);
        } catch (Exception e) {
            return R.fatal("AI 智能匹配导出失败: " + e.getMessage());
        }
    }
}
