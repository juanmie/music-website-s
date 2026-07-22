package com.example.yin.service;

import com.example.yin.model.domain.DemandEvaluation;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.*;
import org.springframework.http.client.SimpleClientHttpRequestFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;

import java.util.*;

/**
 * AI 分析服务 - 使用 Kimi API 进行需求分析
 */
@Service
public class AiAnalysisService {

    private static final Logger log = LoggerFactory.getLogger(AiAnalysisService.class);

    // 主 API: Kimi 通用平台
    @Value("${kimi.api.key}")
    private String apiKey;

    @Value("${kimi.api.url}")
    private String apiUrl;

    @Value("${kimi.model}")
    private String model;

    // 备用 API: Kimi Code
    @Value("${kimi.fallback.key}")
    private String fallbackApiKey;

    @Value("${kimi.fallback.url}")
    private String fallbackUrl;

    @Value("${kimi.fallback.model}")
    private String fallbackModel;

    // 配置较长的超时时间：连接 30 秒，读取 2 分钟
    private final RestTemplate restTemplate = createRestTemplate();
    private final ObjectMapper objectMapper = new ObjectMapper();

    private static RestTemplate createRestTemplate() {
        SimpleClientHttpRequestFactory factory = new SimpleClientHttpRequestFactory();
        factory.setConnectTimeout(30000);  // 连接超时 30 秒
        factory.setReadTimeout(120000);    // 读取超时 2 分钟
        return new RestTemplate(factory);
    }

    /**
     * 自动分类/评级 - 根据规格说明建议客制类型和难度等级
     */
    public Map<String, Object> autoClassify(DemandEvaluation demand) {
        String prompt = String.format(
            "你是一个软件需求分析专家。请根据以下需求信息，建议合适的客制类型和难度等级。\n\n" +
            "程序名称: %s\n" +
            "程序代号: %s\n" +
            "规格说明: %s\n\n" +
            "客制类型选项: 1.新增 / 2.修改标准 / 3.小修客制 / 4.版更\n" +
            "难度等级选项: 1:一级(入门) / 2:二级(初级) / 3:三级(中级) / 4:四级(高级) / 5:五级(资级)\n\n" +
            "请以JSON格式返回，包含 customType 和 difficultyLevel 两个字段，只返回JSON不要其他内容。",
            demand.getProgramName(),
            demand.getProgramCode(),
            demand.getSpecDesc()
        );

        String response = callKimiApi(prompt);
        return parseJsonResponse(response);
    }

    /**
     * 工时估算 - 基于需求描述估算开发时数
     */
    public Map<String, Object> estimateHours(DemandEvaluation demand) {
        String prompt = String.format(
            "你是一个软件项目管理专家。请根据以下需求信息，估算开发所需的时数。\n\n" +
            "程序名称: %s\n" +
            "程序代号: %s\n" +
            "规格说明: %s\n" +
            "客制类型: %s\n" +
            "难度等级: %s\n\n" +
            "请估算：\n" +
            "1. billingHoursCustomer: 客户计费时数\n" +
            "2. dispatchHours: 派工时数\n" +
            "3. 估算依据说明\n\n" +
            "请以JSON格式返回，包含 billingHoursCustomer(数字)、dispatchHours(数字)、reason(字符串) 三个字段，只返回JSON不要其他内容。",
            demand.getProgramName(),
            demand.getProgramCode(),
            demand.getSpecDesc(),
            demand.getCustomType(),
            demand.getDifficultyLevel()
        );

        String response = callKimiApi(prompt);
        Map<String, Object> result = parseJsonResponse(response);
        
        // 添加估算说明
        if (result.containsKey("reason")) {
            result.put("estimateReason", result.get("reason"));
        }
        
        return result;
    }

    /**
     * 智能汇总 - 对需求列表进行汇总分析
     */
    public Map<String, Object> summarizeRequirements(List<DemandEvaluation> demands) {
        StringBuilder demandList = new StringBuilder();
        for (int i = 0; i < demands.size(); i++) {
            DemandEvaluation d = demands.get(i);
            demandList.append(String.format(
                "%d. [%s] %s - %s (难度: %s, 类型: %s)\n",
                i + 1, d.getSystemCode(), d.getProgramCode(), d.getProgramName(),
                d.getDifficultyLevel(), d.getCustomType()
            ));
        }

        String prompt = String.format(
            "你是一个软件需求管理专家。请对以下需求清单进行汇总分析：\n\n%s\n" +
            "请提供：\n" +
            "1. total: 需求总数\n" +
            "2. byCustomType: 按客制类型分类统计\n" +
            "3. byDifficulty: 按难度等级分类统计\n" +
            "4. bySystem: 按系统分类统计\n" +
            "5. summary: 总体分析摘要（100字以内）\n" +
            "6. suggestions: 改进建议（可选）\n\n" +
            "请以JSON格式返回，只返回JSON不要其他内容。",
            demandList.toString()
        );

        String response = callKimiApi(prompt);
        return parseJsonResponse(response);
    }

    /**
     * 需求优化建议 - 检查需求描述的完整性
     */
    public Map<String, Object> optimizeSuggestions(DemandEvaluation demand) {
        String prompt = String.format(
            "你是一个软件需求分析专家。请检查以下需求描述的完整性，并提出优化建议。\n\n" +
            "程序名称: %s\n" +
            "程序代号: %s\n" +
            "系统代号: %s\n" +
            "规格说明: %s\n" +
            "客制类型: %s\n" +
            "难度等级: %s\n\n" +
            "请评估：\n" +
            "1. completenessScore: 完整性评分(0-100)\n" +
            "2. missingFields: 缺失的关键信息列表\n" +
            "3. suggestions: 优化建议列表\n" +
            "4. improvedDesc: 改进后的规格说明（可选）\n\n" +
            "请以JSON格式返回，只返回JSON不要其他内容。",
            demand.getProgramName(),
            demand.getProgramCode(),
            demand.getSystemCode(),
            demand.getSpecDesc(),
            demand.getCustomType(),
            demand.getDifficultyLevel()
        );

        String response = callKimiApi(prompt);
        return parseJsonResponse(response);
    }

    /**
     * AI 智能匹配 - 基于需求清单与接口清单进行智能语义匹配
     * 参考 api-document-retrieval skill 的 match_ai_assist.js
     * 
     * @param demands 需求清单
     * @param interfaceList 接口清单（从【2026】接口清单---汇总.xlsx读取）
     * @return 匹配结果
     */
    public Map<String, Object> batchAnalysis(List<DemandEvaluation> demands, List<Map<String, String>> interfaceList) {
        if (demands == null || demands.isEmpty()) {
            Map<String, Object> emptyResult = new HashMap<>();
            emptyResult.put("matches", Collections.emptyList());
            emptyResult.put("summary", Collections.singletonMap("summary", "无需求数据"));
            return emptyResult;
        }

        // 构建需求列表 - 包含完整信息用于语义分析
        StringBuilder demandList = new StringBuilder();
        for (int i = 0; i < demands.size(); i++) {
            DemandEvaluation d = demands.get(i);
            demandList.append(String.format(
                "%d. [ID:%d]\n" +
                "   程序代号: %s\n" +
                "   程序名称: %s\n" +
                "   系统代号: %s\n" +
                "   开发分类: %s\n" +
                "   产品集成: %s\n" +
                "   客制类型: %s\n" +
                "   难度等级: %s\n" +
                "   规格说明: %s\n" +
                "   备注: %s\n\n",
                i + 1, d.getId(), 
                d.getProgramCode() != null ? d.getProgramCode() : "",
                d.getProgramName() != null ? d.getProgramName() : "",
                d.getSystemCode() != null ? d.getSystemCode() : "",
                d.getDevCategory() != null ? d.getDevCategory() : "",
                d.getIntegrateProductCode() != null ? d.getIntegrateProductCode() : "",
                d.getCustomType() != null ? d.getCustomType() : "",
                d.getDifficultyLevel() != null ? d.getDifficultyLevel() : "",
                d.getSpecDesc() != null ? d.getSpecDesc().substring(0, Math.min(300, d.getSpecDesc().length())) : "",
                d.getRemark() != null ? d.getRemark() : ""
            ));
        }

        // 构建候选接口列表（取前 100 个，避免超出 token 限制）
        StringBuilder candidateList = new StringBuilder();
        int limit = Math.min(interfaceList.size(), 100);
        for (int i = 0; i < limit; i++) {
            Map<String, String> api = interfaceList.get(i);
            candidateList.append(String.format(
                "%d. 流程名称:%s | 接口方向:%s | 适配类型:%s | 模组:%s\n",
                i + 1, 
                api.getOrDefault("流程名称", ""), 
                api.getOrDefault("接口方向", ""),
                api.getOrDefault("适配类型", ""),
                api.getOrDefault("模组", "")
            ));
        }

        // 构建 AI 提示词 - 增强版
        String prompt = String.format(
            "你是一个 T100/OA 系统集成专家。请根据需求清单与接口清单进行智能匹配分析。\n\n" +
            "=== 需求清单 ===\n%s\n" +
            "=== 接口清单 ===\n%s\n" +
            "=== 匹配分析规则 ===\n" +
            "1. **程序代号匹配**：需求的程序代号必须在接口流程名称中出现（如 axmt540、aimm200）\n" +
            "2. **程序名称匹配**：根据程序名称的语义判断匹配哪个具体流程\n" +
            "   - 例如：\"创建出货单\" 应匹配 \"axmt540创建出货单-客制新增\" 而非 \"axmt540过账更新出货单-客制新增\"\n" +
            "3. **开发分类判断流向**：\n" +
            "   - \"新增\"、\"创建\"、\"生单\" → 客制新增类型\n" +
            "   - \"更新\"、\"修改\"、\"过账\" → 客制更新类型\n" +
            "4. **产品集成**：考虑产品集成代码与接口的关联\n" +
            "5. **备注信息**：备注中可能包含关键的匹配线索\n\n" +
            "=== 分析输出要求 ===\n" +
            "对每条需求，请分析并返回：\n" +
            "1. matchedFlowName: 匹配到的流程名称（必须从接口清单中选择）\n" +
            "2. interfaceDirection: 接口方向\n" +
            "3. isCustomNew: 是否为客制新增（true/false）\n" +
            "4. customType: 客制类型（客制新增/客制更新/客制修改）\n" +
            "5. confidence: 匹配确信度 (0-100)\n" +
            "6. reason: 匹配理由（简要说明为什么匹配到这个流程）\n\n" +
            "=== 返回格式 ===\n" +
            "{\n" +
            "  \"summary\": {\n" +
            "    \"total\": 需求总数,\n" +
            "    \"matched\": 匹配成功数,\n" +
            "    \"unmatched\": 未匹配数,\n" +
            "    \"customNewCount\": 客制新增数量,\n" +
            "    \"customUpdateCount\": 客制更新数量,\n" +
            "    \"summary\": \"匹配结果摘要(100字以内)\"\n" +
            "  },\n" +
            "  \"matches\": [\n" +
            "    {\n" +
            "      \"id\": 需求ID,\n" +
            "      \"matchedFlowName\": \"匹配的流程名称\",\n" +
            "      \"interfaceDirection\": \"接口方向\",\n" +
            "      \"isCustomNew\": true或false,\n" +
            "      \"customType\": \"客制新增\"或\"客制更新\",\n" +
            "      \"confidence\": 匹配度(0-100),\n" +
            "      \"reason\": \"匹配理由\"\n" +
            "    }\n" +
            "  ]\n" +
            "}\n\n" +
            "注意：只返回JSON，不要其他内容。",
            demandList.toString(),
            candidateList.toString()
        );

        String response = callKimiApi(prompt);
        Map<String, Object> result = parseJsonResponse(response);

        // 确保返回结构完整
        if (!result.containsKey("summary")) {
            Map<String, Object> fallbackSummary = new HashMap<>();
            fallbackSummary.put("total", demands.size());
            fallbackSummary.put("matched", 0);
            fallbackSummary.put("unmatched", demands.size());
            fallbackSummary.put("summary", "AI 匹配完成");
            result.put("summary", fallbackSummary);
        }
        if (!result.containsKey("matches")) {
            result.put("matches", new ArrayList<>());
        }

        return result;
    }

    /**
     * AI 辅助匹配 - 当本地规则匹配失败时，使用 AI 语义匹配作为兜底
     * 参考 api-document-retrieval skill 的 match_ai_assist.js
     * 
     * @param failedDemands 本地匹配失败的需求
     * @param interfaceList 接口清单（程序代号、流程名称、接口方向）
     * @return AI 匹配结果
     */
    public List<Map<String, Object>> aiAssistMatch(List<DemandEvaluation> failedDemands, 
                                                     List<Map<String, String>> interfaceList) {
        if (failedDemands == null || failedDemands.isEmpty()) {
            return Collections.emptyList();
        }

        // 构建需求列表 - 包含完整信息
        StringBuilder demandList = new StringBuilder();
        for (int i = 0; i < failedDemands.size(); i++) {
            DemandEvaluation d = failedDemands.get(i);
            demandList.append(String.format(
                "%d. [ID:%d]\n" +
                "   程序代号: %s\n" +
                "   程序名称: %s\n" +
                "   开发分类: %s\n" +
                "   产品集成: %s\n" +
                "   规格说明: %s\n" +
                "   备注: %s\n\n",
                i + 1, d.getId(), 
                d.getProgramCode() != null ? d.getProgramCode() : "",
                d.getProgramName() != null ? d.getProgramName() : "",
                d.getDevCategory() != null ? d.getDevCategory() : "",
                d.getIntegrateProductCode() != null ? d.getIntegrateProductCode() : "",
                d.getSpecDesc() != null ? d.getSpecDesc().substring(0, Math.min(200, d.getSpecDesc().length())) : "",
                d.getRemark() != null ? d.getRemark() : ""
            ));
        }

        // 构建候选接口列表（取前 60 个）
        StringBuilder candidateList = new StringBuilder();
        int limit = Math.min(interfaceList.size(), 60);
        for (int i = 0; i < limit; i++) {
            Map<String, String> api = interfaceList.get(i);
            candidateList.append(String.format(
                "%d. 流程名称:%s | 接口方向:%s | 适配类型:%s\n",
                i + 1, 
                api.getOrDefault("流程名称", ""), 
                api.getOrDefault("接口方向", ""),
                api.getOrDefault("适配类型", "")
            ));
        }

        String prompt = String.format(
            "你是一个 T100/OA 系统集成专家。以下需求无法通过本地规则匹配到接口，请使用语义分析进行匹配。\n\n" +
            "=== 待匹配需求 ===\n%s\n" +
            "=== 候选接口清单 ===\n%s\n" +
            "=== 匹配任务 ===\n" +
            "请为每个需求从候选接口中找到最匹配的接口，分析要点：\n" +
            "1. 程序代号必须匹配（如 axmt540 必须出现在流程名称中）\n" +
            "2. 程序名称的语义要与流程名称匹配（如 \"创建出货单\" 匹配 \"axmt540创建出货单-客制新增\"）\n" +
            "3. 根据开发分类判断是客制新增还是客制更新\n" +
            "4. 备注信息可能包含关键线索\n\n" +
            "=== 返回格式 ===\n" +
            "{\n" +
            "  \"matches\": [\n" +
            "    { \n" +
            "      \"id\": 需求ID, \n" +
            "      \"matchedFlowName\": \"匹配的流程名称\", \n" +
            "      \"interfaceDirection\": \"接口方向\",\n" +
            "      \"isCustomNew\": true或false,\n" +
            "      \"confidence\": 80, \n" +
            "      \"reason\": \"匹配理由\" \n" +
            "    }\n" +
            "  ]\n" +
            "}\n\n" +
            "注意：\n" +
            "1. confidence 为 0-100 的匹配确信度\n" +
            "2. 如果找不到合适的接口，confidence 设为 0\n" +
            "3. 只返回JSON，不要其他内容。",
            demandList.toString(),
            candidateList.toString()
        );

        String response = callKimiApi(prompt);
        Map<String, Object> result = parseJsonResponse(response);

        // 提取 matches 数组
        Object matchesObj = result.get("matches");
        if (matchesObj instanceof List) {
            @SuppressWarnings("unchecked")
            List<Map<String, Object>> matches = (List<Map<String, Object>>) matchesObj;
            return matches;
        }

        return Collections.emptyList();
    }

    /**
     * 调用 Kimi API（带自动回退）
     * 优先使用 Kimi 通用平台，流量耗尽时回退到 Kimi Code
     */
    private String callKimiApi(String prompt) {
        // 1. 先尝试主 API（Kimi 通用平台）
        try {
            return doCallKimiApi(apiKey, apiUrl, model, prompt, false);
        } catch (Exception e) {
            String errorMsg = e.getMessage();
            // 判断是否需要回退（流量耗尽、配额超限、速率限制等）
            if (shouldFallback(errorMsg)) {
                log.warn("主 API 不可用 ({}), 回退到 Kimi Code", errorMsg);
                // 2. 回退到备用 API（Kimi Code）
                try {
                    return doCallKimiApi(fallbackApiKey, fallbackUrl, fallbackModel, prompt, true);
                } catch (Exception fallbackEx) {
                    throw new RuntimeException("Kimi API 调用失败（主 API 和备用 API 均不可用）: " + fallbackEx.getMessage(), fallbackEx);
                }
            }
            throw new RuntimeException("Kimi API 调用失败: " + errorMsg, e);
        }
    }

    /**
     * 判断是否应该回退到备用 API
     */
    private boolean shouldFallback(String errorMsg) {
        if (errorMsg == null) return false;
        String lower = errorMsg.toLowerCase();
        return lower.contains("quota") || 
               lower.contains("rate limit") || 
               lower.contains("insufficient") || 
               lower.contains("exceeded") ||
               lower.contains("balance") ||
               lower.contains("429") ||
               lower.contains("402");
    }

    /**
     * 实际调用 Kimi API
     * @param isFallback 是否为回退调用（用于日志）
     */
    private String doCallKimiApi(String key, String url, String modelName, String prompt, boolean isFallback) {
        try {
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_JSON);
            headers.setBearerAuth(key);

            ObjectNode requestBody = objectMapper.createObjectNode();
            requestBody.put("model", modelName);
            
            ArrayNode messages = requestBody.putArray("messages");
            ObjectNode message = messages.addObject();
            message.put("role", "user");
            message.put("content", prompt);
            
            // Kimi Code 要求 temperature=1，通用平台可以用其他值
            requestBody.put("temperature", 1);

            HttpEntity<String> entity = new HttpEntity<>(objectMapper.writeValueAsString(requestBody), headers);
            ResponseEntity<String> response = restTemplate.exchange(url, HttpMethod.POST, entity, String.class);

            JsonNode responseJson = objectMapper.readTree(response.getBody());
            return responseJson.path("choices").path(0).path("message").path("content").asText();
        } catch (Exception e) {
            throw new RuntimeException((isFallback ? "[备用API] " : "") + "API 调用失败: " + e.getMessage(), e);
        }
    }

    /**
     * 解析 JSON 响应
     */
    private Map<String, Object> parseJsonResponse(String response) {
        try {
            // 尝试提取 JSON 部分
            String jsonStr = response.trim();
            if (jsonStr.startsWith("```json")) {
                jsonStr = jsonStr.substring(7);
            }
            if (jsonStr.startsWith("```")) {
                jsonStr = jsonStr.substring(3);
            }
            if (jsonStr.endsWith("```")) {
                jsonStr = jsonStr.substring(0, jsonStr.length() - 3);
            }
            jsonStr = jsonStr.trim();
            
            @SuppressWarnings("unchecked")
            Map<String, Object> result = objectMapper.readValue(jsonStr, Map.class);
            return result;
        } catch (Exception e) {
            Map<String, Object> errorResult = new HashMap<>();
            errorResult.put("rawResponse", response);
            errorResult.put("error", "JSON 解析失败: " + e.getMessage());
            return errorResult;
        }
    }
}
