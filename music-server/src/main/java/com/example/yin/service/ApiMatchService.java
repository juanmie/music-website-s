package com.example.yin.service;

import com.example.yin.model.domain.DemandEvaluation;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import javax.annotation.PostConstruct;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * API 文档匹配服务
 * 根据需求清单自动匹配 T100/OA 集成标准 API 对照清单
 */
@Service
public class ApiMatchService {

    private static final Logger log = LoggerFactory.getLogger(ApiMatchService.class);

    // 项目根目录（不含中文）
    private static final String PROJECT_ROOT = "/Users/dsh/Documents/music-website-s";

    @Value("${api.match.data-dir:}")
    private String dataDir;

    @Value("${api.match.interface-list-name:}")
    private String interfaceListName;

    @Value("${api.match.field-mapping-name:}")
    private String fieldMappingName;

    @Value("${api.match.output-template-dir:}")
    private String outputTemplateDir;

    @Value("${api.match.output-template-name:}")
    private String outputTemplateName;

    // 获取实际使用的数据目录（带兜底）
    private String getEffectiveDataDir() {
        // 先尝试配置路径
        if (dataDir != null && !dataDir.isEmpty()) {
            File dir = new File(dataDir);
            if (dir.exists() && dir.isDirectory()) {
                return dataDir;
            }
        }
        // 兜底：扫描项目根目录查找包含接口清单xlsx的目录
        File root = new File(PROJECT_ROOT);
        if (root.exists()) {
            File[] subdirs = root.listFiles(File::isDirectory);
            if (subdirs != null) {
                for (File subdir : subdirs) {
                    File[] xlsxFiles = subdir.listFiles((d, name) -> name.endsWith(".xlsx") && name.contains("2026"));
                    if (xlsxFiles != null && xlsxFiles.length >= 2) {
                        log.info("通过扫描找到数据目录: {}", subdir.getAbsolutePath());
                        return subdir.getAbsolutePath();
                    }
                }
            }
        }
        log.warn("无法找到数据目录");
        return PROJECT_ROOT;
    }

    // 获取实际使用的接口清单文件名（带兜底）
    private String getEffectiveInterfaceListName() {
        if (interfaceListName != null && !interfaceListName.isEmpty()) {
            return interfaceListName;
        }
        // 兜底：在数据目录中查找包含"接口清单"的xlsx文件
        return findFileInDir(getEffectiveDataDir(), "xlsx", null);
    }

    // 获取实际使用的字段对照文件名（带兜底）
    private String getEffectiveFieldMappingName() {
        if (fieldMappingName != null && !fieldMappingName.isEmpty()) {
            return fieldMappingName;
        }
        return findFileInDir(getEffectiveDataDir(), "xlsx", null);
    }

    // 获取实际使用的输出模板目录（带兜底）
    private String getEffectiveOutputTemplateDir() {
        if (outputTemplateDir != null && !outputTemplateDir.isEmpty()) {
            File dir = new File(outputTemplateDir);
            if (dir.exists() && dir.isDirectory()) {
                return outputTemplateDir;
            }
        }
        // 兜底：扫描项目根目录，查找包含 "AI" 或 "摸板" 或 "模板" xlsx文件的目录
        File root = new File(PROJECT_ROOT);
        if (root.exists()) {
            File[] subdirs = root.listFiles(File::isDirectory);
            if (subdirs != null) {
                for (File subdir : subdirs) {
                    File[] xlsxFiles = subdir.listFiles((d, name) -> name.endsWith(".xlsx"));
                    if (xlsxFiles != null) {
                        for (File f : xlsxFiles) {
                            String name = f.getName();
                            // 查找包含 "AI" 或 "摸板" 或 "模板" 或 "输出物" 的文件
                            if (name.contains("AI") || name.contains("摸板") || name.contains("模板") || name.contains("输出物")) {
                                log.info("通过扫描找到输出模板目录: {}", subdir.getAbsolutePath());
                                return subdir.getAbsolutePath();
                            }
                        }
                    }
                }
            }
        }
        log.warn("无法找到输出模板目录");
        return PROJECT_ROOT;
    }

    // 获取实际使用的输出模板文件名（带兜底）
    private String getEffectiveOutputTemplateName() {
        if (outputTemplateName != null && !outputTemplateName.isEmpty()) {
            return outputTemplateName;
        }
        return findFileInDir(getEffectiveOutputTemplateDir(), "xlsx", null);
    }

    /**
     * 在目录中查找匹配的文件名（通过扫描实际文件名避免编码问题）
     */
    private String findFileInDir(String dirPath, String extension, String keyword) {
        File dir = new File(dirPath);
        if (!dir.exists() || !dir.isDirectory()) return null;
        File[] files = dir.listFiles((d, name) -> name.endsWith("." + extension));
        if (files == null || files.length == 0) return null;
        
        // 如果有关键词，优先匹配
        if (keyword != null) {
            for (File f : files) {
                if (f.getName().contains(keyword)) return f.getName();
            }
        }
        // 否则返回第一个
        return files[0].getName();
    }
    
    /**
     * 在目录中查找包含关键词的文件（通过扫描实际文件名避免编码问题）
     */
    private File findFileByKeyword(String dirPath, String keyword) {
        File dir = new File(dirPath);
        if (!dir.exists() || !dir.isDirectory()) return null;
        File[] files = dir.listFiles((d, name) -> name.endsWith(".xlsx"));
        if (files == null) return null;
        for (File f : files) {
            // 通过 File.getName() 获取实际文件名，避免编码问题
            String name = f.getName();
            log.debug("扫描文件: {}", name);
            if (name.contains(keyword)) {
                log.info("找到匹配文件: {}", f.getAbsolutePath());
                return f;
            }
        }
        return null;
    }

    // 需要从输出模板复制的标准工作表
    public static final List<String> TEMPLATE_SHEETS_TO_COPY = Arrays.asList(
        "集成配置", "开发计划", "问题跟进管制-TB导入联动"
    );

    // 开发分类 → 流向映射规则
    private static final Map<String, String> FLOW_DIRECTION_RULES = new HashMap<>();
    static {
        FLOW_DIRECTION_RULES.put("702", "XX→T100");
        FLOW_DIRECTION_RULES.put("703", "XX→T100");
        FLOW_DIRECTION_RULES.put("704", "XX→T100");
        FLOW_DIRECTION_RULES.put("705", "XX→T100");
        FLOW_DIRECTION_RULES.put("710", "T100→XX");
        FLOW_DIRECTION_RULES.put("711", "T100→XX");
        FLOW_DIRECTION_RULES.put("202", "T100→XX");
    }

    // 业务关键词列表
    private static final List<String> BUSINESS_KEYWORDS = Arrays.asList(
        "员工", "部门", "料号", "客户", "供应商", "采购", "销售", "订单",
        "核价", "请购", "入库", "出库", "库存", "杂收", "杂发", "付款",
        "收款", "报销", "借款", "预算", "资产", "理财", "发票", "核销",
        "申请", "维护", "查询", "新增", "修改", "删除", "审核", "送签",
        "推送", "同步", "抛转", "回写", "更新", "创建", "生成"
    );

    // 程序代号正则
    private static final Pattern PROG_CODE_PATTERN = Pattern.compile("([a-z]{2,4}\\d{3})");
    private static final Pattern DEV_CLASS_PATTERN = Pattern.compile("(\\d{3})");

    // 缓存的接口清单数据
    private List<Map<String, String>> cachedInterfaceList;
    private long lastLoadTime = 0;
    private static final long CACHE_TTL = 5 * 60 * 1000; // 5分钟缓存
    
    // 缓存字段对照工作簿和输出模板工作簿
    private Workbook cachedFieldMappingWb;
    private long fieldMappingLoadTime = 0;
    private Workbook cachedOutputTemplateWb;
    private long outputTemplateLoadTime = 0;
    private static final long WB_CACHE_TTL = 30 * 60 * 1000; // 30分钟缓存

    @PostConstruct
    public void init() {
        log.info("API匹配服务初始化");
        log.info("  配置数据目录: {} -> 实际使用: {}", dataDir, getEffectiveDataDir());
        log.info("  配置接口清单: {} -> 实际使用: {}", interfaceListName, getEffectiveInterfaceListName());
        log.info("  配置字段对照: {} -> 实际使用: {}", fieldMappingName, getEffectiveFieldMappingName());
        log.info("  配置输出模板目录: {} -> 实际使用: {}", outputTemplateDir, getEffectiveOutputTemplateDir());
        log.info("  配置输出模板文件: {} -> 实际使用: {}", outputTemplateName, getEffectiveOutputTemplateName());
        
        // 预加载 Excel 文件，避免首次导出时从磁盘加载的延迟
        long preloadStart = System.currentTimeMillis();
        try {
            loadInterfaceList();
            log.info("  预加载接口清单完成, 共{}条", cachedInterfaceList != null ? cachedInterfaceList.size() : 0);
        } catch (Exception e) {
            log.warn("  预加载接口清单失败: {}", e.getMessage());
        }
        try {
            getFieldMappingWorkbook();
            log.info("  预加载字段对照工作簿完成, 工作表数: {}", cachedFieldMappingWb != null ? cachedFieldMappingWb.getNumberOfSheets() : 0);
        } catch (Exception e) {
            log.warn("  预加载字段对照工作簿失败: {}", e.getMessage());
        }
        try {
            getOutputTemplateWorkbook();
            log.info("  预加载输出模板工作簿完成");
        } catch (Exception e) {
            log.warn("  预加载输出模板工作簿失败: {}", e.getMessage());
        }
        log.info("API匹配服务初始化完成, 预加载总耗时: {}ms", System.currentTimeMillis() - preloadStart);
    }

    /**
     * 解析数据文件路径（支持绝对路径和相对路径）
     */
    private File resolveDataFile(String fileName) {
        String userDir = System.getProperty("user.dir");
        log.debug("解析文件路径, user.dir={}, dataDir={}, fileName={}", userDir, dataDir, fileName);
        
        // 1. 如果 dataDir 是绝对路径，直接使用
        File absDir = new File(dataDir);
        if (absDir.isAbsolute() && absDir.exists()) {
            File absFile = new File(absDir, fileName);
            if (absFile.exists()) {
                log.debug("找到文件(绝对路径): {}", absFile.getAbsolutePath());
                return absFile;
            }
        }
        
        // 2. 尝试 user.dir + dataDir
        File file1 = new File(userDir, dataDir + File.separator + fileName);
        if (file1.exists()) {
            log.debug("找到文件(user.dir): {}", file1.getAbsolutePath());
            return file1;
        }
        
        // 3. 尝试 user.dir/../dataDir (当工作目录是 music-server 时)
        File file2 = new File(userDir + File.separator + ".." + File.separator + dataDir, fileName);
        if (file2.exists()) {
            log.debug("找到文件(上级目录): {}", file2.getAbsolutePath());
            return file2;
        }
        
        // 4. 尝试固定路径 (项目根目录)
        String projectRoot = "/Users/dsh/Documents/music-website-s";
        File file3 = new File(projectRoot, dataDir + File.separator + fileName);
        if (file3.exists()) {
            log.debug("找到文件(项目根目录): {}", file3.getAbsolutePath());
            return file3;
        }
        
        // 5. 尝试 user.dir/music-server/../dataDir
        File file4 = new File(userDir + File.separator + "music-server" + File.separator + ".." + File.separator + dataDir, fileName);
        if (file4.exists()) {
            log.debug("找到文件(music-server上级): {}", file4.getAbsolutePath());
            return file4;
        }
        
        // 返回默认路径，记录警告
        File defaultFile = new File(userDir, dataDir + File.separator + fileName);
        log.warn("文件不存在，尝试的路径: {}", defaultFile.getAbsolutePath());
        return defaultFile;
    }

    private File getInterfaceListFile() {
        String effectiveDir = getEffectiveDataDir();
        // 通过扫描文件名查找接口清单文件（包含 "接口清单" 或 "汇总" 的 xlsx）
        File found = findFileByKeyword(effectiveDir, "汇总");
        if (found == null) {
            found = findFileByKeyword(effectiveDir, "接口");
        }
        if (found != null) {
            log.info("接口清单文件: {}", found.getAbsolutePath());
            return found;
        }
        // 兜底：使用配置的文件名
        String effectiveName = getEffectiveInterfaceListName();
        File file = new File(effectiveDir, effectiveName);
        log.info("接口清单文件路径(配置): {}", file.getAbsolutePath());
        return file;
    }

    private File getFieldMappingFile() {
        String effectiveDir = getEffectiveDataDir();
        // 通过扫描文件名查找字段对照文件（包含 "字段对照" 或 "颗粒度" 的 xlsx）
        File found = findFileByKeyword(effectiveDir, "颗粒度");
        if (found == null) {
            found = findFileByKeyword(effectiveDir, "字段");
        }
        if (found != null) {
            log.info("字段对照文件: {}", found.getAbsolutePath());
            return found;
        }
        // 兜底：使用配置的文件名
        String effectiveName = getEffectiveFieldMappingName();
        File file = new File(effectiveDir, effectiveName);
        log.info("字段对照文件路径(配置): {}", file.getAbsolutePath());
        return file;
    }

    /**
     * 解析输出模板文件路径
     */
    private File getOutputTemplateFile() {
        String effectiveDir = getEffectiveOutputTemplateDir();
        // 通过扫描文件名查找输出模板文件（包含 "摸板" 或 "模板" 或 "AI" 的 xlsx）
        File found = findFileByKeyword(effectiveDir, "摸板");
        if (found == null) {
            found = findFileByKeyword(effectiveDir, "模板");
        }
        if (found == null) {
            found = findFileByKeyword(effectiveDir, "AI");
        }
        if (found != null) {
            log.info("输出模板文件: {}", found.getAbsolutePath());
            return found;
        }
        // 兜底：使用配置的文件名
        String effectiveName = getEffectiveOutputTemplateName();
        File file = new File(effectiveDir, effectiveName);
        log.info("输出模板文件路径(配置): {}", file.getAbsolutePath());
        return file;
    }

    /**
     * 获取 AI 输出物模板工作簿（用于复制集成配置、开发计划、问题跟进管制等工作表）
     */
    public Workbook getOutputTemplateWorkbook() {
        // 检查缓存是否有效
        if (cachedOutputTemplateWb != null && System.currentTimeMillis() - outputTemplateLoadTime < WB_CACHE_TTL) {
            log.debug("使用缓存的输出模板工作簿");
            return cachedOutputTemplateWb;
        }
        
        // 关闭旧的工作簿
        if (cachedOutputTemplateWb != null) {
            try { cachedOutputTemplateWb.close(); } catch (IOException ignored) {}
        }
        
        File file = getOutputTemplateFile();
        if (file == null || !file.exists()) {
            log.warn("找不到AI输出物模板: {}", file != null ? file.getAbsolutePath() : "未配置");
            return null;
        }

        try {
            long start = System.currentTimeMillis();
            FileInputStream fis = new FileInputStream(file);
            cachedOutputTemplateWb = new XSSFWorkbook(fis);
            outputTemplateLoadTime = System.currentTimeMillis();
            log.info("加载输出模板工作簿耗时: {}ms", outputTemplateLoadTime - start);
            return cachedOutputTemplateWb;
        } catch (IOException e) {
            log.error("读取AI输出物模板失败: {}", e.getMessage(), e);
            return null;
        }
    }

    /**
     * 对需求列表进行 API 匹配
     */
    public List<MatchResult> matchAll(List<DemandEvaluation> demands) {
        loadInterfaceListIfNeeded();
        
        List<MatchResult> results = new ArrayList<>();
        for (DemandEvaluation demand : demands) {
            MatchResult result = matchSingle(demand);
            results.add(result);
        }
        
        long matchedCount = results.stream().filter(r -> r.getScore() >= 40).count();
        log.info("匹配完成: {}/{} 条需求找到匹配API", matchedCount, demands.size());
        
        return results;
    }

    /**
     * 获取接口清单（供 AI 辅助匹配使用）
     */
    public List<Map<String, String>> getInterfaceList() {
        loadInterfaceListIfNeeded();
        return cachedInterfaceList != null ? cachedInterfaceList : Collections.emptyList();
    }

    /**
     * 获取匹配失败的需求（匹配度 < 40）
     */
    public List<DemandEvaluation> getFailedDemands(List<DemandEvaluation> demands, List<MatchResult> results) {
        List<DemandEvaluation> failed = new ArrayList<>();
        for (int i = 0; i < demands.size(); i++) {
            if (i < results.size() && results.get(i).getScore() < 40) {
                failed.add(demands.get(i));
            }
        }
        return failed;
    }

    /**
     * 对单条需求进行匹配
     * 匹配逻辑：程序代号 + 程序名称 共同确定唯一的流程名称
     */
    public MatchResult matchSingle(DemandEvaluation demand) {
        String progCode = safeStr(demand.getProgramCode()).toLowerCase().trim();
        String devCategory = safeStr(demand.getDevCategory());
        String specDesc = safeStr(demand.getSpecDesc());
        String programName = safeStr(demand.getProgramName()).trim();
        String remark = safeStr(demand.getRemark()).trim();

        log.debug("开始匹配需求: progCode={}, programName={}, devCategory={}, remark={}", progCode, programName, devCategory, remark);
        
        // 根据备注判断适配类型：客制新增(复用函数) vs OPENAPI(标准API)
        String expectedAdaptType = "";
        if (remark.isEmpty() || remark.contains("客制新增") || remark.contains("复用")) {
            expectedAdaptType = "客制新增";
        } else if (remark.contains("OPENAPI") || remark.contains("openapi") || remark.contains("标准")) {
            expectedAdaptType = "OPENAPI";
        }

        MatchResult result = new MatchResult();
        result.setDemandId(demand.getId());
        result.setProgramCode(progCode);
        result.setProgramName(programName);

        if (progCode.isEmpty()) {
            result.setMatched(false);
            result.setScore(0);
            result.setReason("程序代号为空");
            return result;
        }

        // 推断流向
        String inferredFlow = inferFlowDirection(devCategory, specDesc);
        result.setInferredFlow(inferredFlow);
        log.debug("推断流向: {}", inferredFlow);

        // 在接口清单中匹配
        Map<String, Object> bestMatch = null;
        int bestScore = 0;
        String bestReason = "";
        int matchCount = 0;

        for (Map<String, String> apiRow : cachedInterfaceList) {
            String flowName = apiRow.getOrDefault("流程名称", "");
            String apiDirection = apiRow.getOrDefault("接口方向", "");

            if (flowName.isEmpty()) continue;

            int score = 0;
            List<String> reasons = new ArrayList<>();

            // 1. 程序代号匹配（必须）
            String apiProgCode = extractProgCode(flowName);
            boolean progCodeMatch = false;
            if (apiProgCode.equals(progCode)) {
                score += 50;
                reasons.add("程序代号精确匹配: " + progCode);
                progCodeMatch = true;
                log.debug("找到精确匹配! flowName={}, apiProgCode={}", flowName, apiProgCode);
            } else if (flowName.toLowerCase().contains(progCode)) {
                score += 30;
                reasons.add("程序代号包含匹配: " + progCode);
                progCodeMatch = true;
                log.debug("找到包含匹配! flowName={}", flowName);
            }
            
            // 如果程序代号不匹配，跳过此条
            if (!progCodeMatch) {
                continue;
            }

            // 2. 程序名称匹配（关键）
            if (!programName.isEmpty()) {
                String flowNameLower = flowName.toLowerCase();
                String programNameLower = programName.toLowerCase();
                
                // 检查流程名称是否包含程序名称
                if (flowNameLower.contains(programNameLower)) {
                    score += 40;
                    reasons.add("程序名称完全匹配: " + programName);
                    log.debug("程序名称完全匹配! flowName={}, programName={}", flowName, programName);
                } else {
                    // 提取程序名称中的关键词进行匹配
                    List<String> nameKeywords = extractNameKeywords(programName);
                    int matchedKeywords = 0;
                    for (String keyword : nameKeywords) {
                        if (flowNameLower.contains(keyword)) {
                            matchedKeywords++;
                        }
                    }
                    
                    if (!nameKeywords.isEmpty() && matchedKeywords > 0) {
                        int keywordScore = (int) ((matchedKeywords * 30.0) / nameKeywords.size());
                        score += keywordScore;
                        reasons.add("程序名称关键词匹配: " + matchedKeywords + "/" + nameKeywords.size());
                        log.debug("程序名称关键词匹配: {}/{}", matchedKeywords, nameKeywords.size());
                    } else {
                        // 程序名称不匹配，大幅扣分
                        score -= 20;
                        reasons.add("程序名称不匹配: " + programName);
                    }
                }
            }

            // 3. 流向匹配（加分项）
            if (!inferredFlow.isEmpty() && score > 0) {
                String normalizedApiDir = normalizeFlowDirection(apiDirection);
                if (inferredFlow.equals(normalizedApiDir)) {
                    score += 10;
                    reasons.add("流向匹配: " + inferredFlow + " ↔ " + apiDirection);
                } else if (!normalizedApiDir.isEmpty()) {
                    score -= 5;
                    reasons.add("流向不匹配: " + inferredFlow + " vs " + apiDirection);
                }
            }

            // 4. 规格说明提及程序代号（加分项）
            if (!progCode.isEmpty() && specDesc.toLowerCase().contains(progCode) && score > 0) {
                score += 5;
                reasons.add("规格说明中提及程序代号");
            }
            
            // 5. 适配类型匹配（根据备注判断是复用函数还是标准API）
            if (!expectedAdaptType.isEmpty() && score > 0) {
                String apiAdaptType = apiRow.getOrDefault("适配类型", "");
                if (apiAdaptType.contains(expectedAdaptType)) {
                    score += 15;
                    reasons.add("适配类型匹配: " + expectedAdaptType);
                } else if (!apiAdaptType.isEmpty()) {
                    score -= 10;
                    reasons.add("适配类型不匹配: 期望" + expectedAdaptType + ", 实际" + apiAdaptType);
                }
            }

            log.debug("匹配结果: flowName={}, score={}, reasons={}", flowName, score, reasons);

            if (score >= 50 && score > bestScore) {
                bestScore = score;
                bestMatch = new HashMap<>(apiRow);
                bestReason = String.join(" | ", reasons);
                matchCount++;
            }
        }

        log.debug("匹配完成: progCode={}, bestScore={}, matchCount={}", progCode, bestScore, matchCount);

        if (bestMatch != null) {
            result.setMatched(true);
            result.setScore(Math.min(bestScore, 100));
            result.setMatchedFlowName(String.valueOf(bestMatch.getOrDefault("流程名称", "")));
            result.setInterfaceDirection(String.valueOf(bestMatch.getOrDefault("接口方向", "")));
            result.setReason(bestReason);
            result.setReferenceSheet(String.valueOf(bestMatch.getOrDefault("_sheet", "")));
            result.setReferenceRow(String.valueOf(bestMatch.getOrDefault("_row_num", "")));
        } else {
            result.setMatched(false);
            result.setScore(0);
            result.setReason(bestReason.isEmpty() ? "无匹配接口" : bestReason);
        }

        return result;
    }
    
    /**
     * 从程序名称中提取关键词
     * 例如："创建出货单" -> ["创建", "出货", "出货单"]
     */
    private List<String> extractNameKeywords(String programName) {
        if (programName == null || programName.isEmpty()) {
            return Collections.emptyList();
        }
        
        List<String> keywords = new ArrayList<>();
        // 添加完整的程序名称
        keywords.add(programName.toLowerCase());
        
        // 添加程序名称中的关键词
        String[] parts = programName.split("[\\s,，、]+");
        for (String part : parts) {
            if (part.length() >= 2) {
                keywords.add(part.toLowerCase());
            }
        }
        
        // 添加业务关键词
        List<String> businessKeywords = extractBusinessKeywords(programName);
        keywords.addAll(businessKeywords);
        
        // 去重
        return keywords.stream().distinct().collect(Collectors.toList());
    }

    /**
     * 根据开发分类推断流向
     */
    private String inferFlowDirection(String devClass, String specDesc) {
        if (devClass == null || devClass.isEmpty()) return "";

        Matcher matcher = DEV_CLASS_PATTERN.matcher(devClass);
        if (matcher.find()) {
            String classCode = matcher.group(1);
            String flow = FLOW_DIRECTION_RULES.getOrDefault(classCode, "");

            // 202 特殊处理：根据规格说明推断
            if ("202".equals(classCode) && specDesc != null) {
                String desc = specDesc.toLowerCase();
                if (desc.contains("查询") || desc.contains("封装") || 
                    desc.contains("ipaas") || desc.contains("api")) {
                    flow = "T100→XX";
                }
            }
            return flow;
        }
        return "";
    }

    /**
     * 标准化流向格式
     */
    private String normalizeFlowDirection(String flowDir) {
        if (flowDir == null || flowDir.isEmpty()) return "";

        String flow = flowDir.toUpperCase()
            .replaceAll("\\s", "")
            .replace("--", "→")
            .replace("->", "→");

        if (flow.startsWith("ERP→") || flow.startsWith("T100→")) {
            return "T100→XX";
        }
        if (flow.startsWith("OA→") || flow.startsWith("XX→")) {
            return "XX→T100";
        }
        if (flow.contains("→")) {
            String[] parts = flow.split("→");
            if (parts[0].contains("ERP") || parts[0].contains("T100")) {
                return "T100→XX";
            } else {
                return "XX→T100";
            }
        }
        return flow;
    }

    /**
     * 从文本中提取程序代号
     */
    private String extractProgCode(String text) {
        if (text == null || text.isEmpty()) return "";
        Matcher matcher = PROG_CODE_PATTERN.matcher(text.toLowerCase());
        return matcher.find() ? matcher.group(1) : "";
    }

    /**
     * 提取业务关键词
     */
    private List<String> extractBusinessKeywords(String text) {
        if (text == null || text.isEmpty()) return Collections.emptyList();
        String textLower = text.toLowerCase();
        return BUSINESS_KEYWORDS.stream()
            .filter(textLower::contains)
            .collect(Collectors.toList());
    }

    /**
     * 加载接口清单（带缓存）
     */
    private void loadInterfaceListIfNeeded() {
        if (cachedInterfaceList != null && System.currentTimeMillis() - lastLoadTime < CACHE_TTL) {
            return;
        }
        cachedInterfaceList = loadInterfaceList();
        lastLoadTime = System.currentTimeMillis();
    }

    /**
     * 从 Excel 文件加载接口清单
     */
    private List<Map<String, String>> loadInterfaceList() {
        List<Map<String, String>> allInterfaces = new ArrayList<>();

        File file = getInterfaceListFile();
        if (file == null || !file.exists()) {
            log.warn("找不到接口清单文件: {}", file != null ? file.getAbsolutePath() : "未配置");
            return allInterfaces;
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = new XSSFWorkbook(fis)) {

            for (int sheetIdx = 0; sheetIdx < wb.getNumberOfSheets(); sheetIdx++) {
                Sheet sheet = wb.getSheetAt(sheetIdx);
                String sheetName = sheet.getSheetName();

                // 查找表头行
                int headerRowIdx = 0;
                Row headerRow = null;
                for (int i = 0; i < Math.min(10, sheet.getPhysicalNumberOfRows()); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        for (Cell cell : row) {
                            String val = getCellStringValue(cell);
                            if ("序号".equals(val) || "流程名称".equals(val)) {
                                headerRowIdx = i;
                                headerRow = row;
                                break;
                            }
                        }
                        if (headerRow != null) break;
                    }
                }

                if (headerRow == null) {
                    log.warn("工作表 {} 未找到表头（查找 '序号' 或 '流程名称'）", sheetName);
                    continue;
                }
                
                log.info("工作表 {} 表头行: {}", sheetName, headerRowIdx);

                // 解析表头，支持列名映射
                Map<Integer, String> headers = new HashMap<>();
                for (int col = 0; col < headerRow.getLastCellNum(); col++) {
                    Cell cell = headerRow.getCell(col);
                    String name = getCellStringValue(cell).trim();
                    if (!name.isEmpty()) {
                        // 列名映射：将"字段对照表及报文"映射为"流程名称"
                        if ("字段对照表及报文".equals(name)) {
                            name = "流程名称";
                        } else if ("函数名称".equals(name)) {
                            name = "函数名称";
                        }
                        headers.put(col, name);
                    }
                }
                log.info("工作表 {} 表头列: {}", sheetName, headers.values());

                // 读取数据行
                for (int rowIdx = headerRowIdx + 1; rowIdx <= sheet.getLastRowNum(); rowIdx++) {
                    Row row = sheet.getRow(rowIdx);
                    if (row == null) continue;

                    Map<String, String> rowData = new LinkedHashMap<>();
                    rowData.put("_sheet", sheetName);
                    rowData.put("_row_num", String.valueOf(rowIdx + 1));

                    for (Map.Entry<Integer, String> entry : headers.entrySet()) {
                        Cell cell = row.getCell(entry.getKey());
                        rowData.put(entry.getValue(), getCellStringValue(cell).trim());
                    }

                    String flowName = rowData.getOrDefault("流程名称", "").trim();
                    if (!flowName.isEmpty()) {
                        allInterfaces.add(rowData);
                    }
                }
            }

            log.info("已加载 {} 条接口记录（来自 {} 个工作表）", allInterfaces.size(), wb.getNumberOfSheets());
            
            // 打印前 3 条记录以便调试（仅首次加载时）
            if (!allInterfaces.isEmpty()) {
                log.debug("接口清单示例（前3条）:");
                for (int i = 0; i < Math.min(3, allInterfaces.size()); i++) {
                    Map<String, String> row = allInterfaces.get(i);
                    log.debug("  [{}] 流程名称={}, 接口方向={}, sheet={}", 
                        i + 1, row.get("流程名称"), row.get("接口方向"), row.get("_sheet"));
                }
            }

        } catch (IOException e) {
            log.error("读取接口清单文件失败: {}", e.getMessage(), e);
        }

        return allInterfaces;
    }

    /**
     * 获取字段对照清单工作簿（用于复制匹配的工作表）- 带缓存
     */
    public Workbook getFieldMappingWorkbook() {
        // 检查缓存是否有效
        if (cachedFieldMappingWb != null && System.currentTimeMillis() - fieldMappingLoadTime < WB_CACHE_TTL) {
            log.debug("使用缓存的字段对照工作簿");
            return cachedFieldMappingWb;
        }
        
        // 关闭旧的工作簿
        if (cachedFieldMappingWb != null) {
            try { cachedFieldMappingWb.close(); } catch (IOException ignored) {}
        }
        
        File file = getFieldMappingFile();
        if (file == null || !file.exists()) {
            log.warn("找不到字段对照清单文件: {}", file != null ? file.getAbsolutePath() : "未配置");
            return null;
        }

        try {
            long start = System.currentTimeMillis();
            FileInputStream fis = new FileInputStream(file);
            cachedFieldMappingWb = new XSSFWorkbook(fis);
            fieldMappingLoadTime = System.currentTimeMillis();
            log.info("加载字段对照工作簿耗时: {}ms, 工作表数: {}", fieldMappingLoadTime - start, cachedFieldMappingWb.getNumberOfSheets());
            return cachedFieldMappingWb;
        } catch (IOException e) {
            log.error("读取字段对照清单文件失败: {}", e.getMessage(), e);
            return null;
        }
    }

    /**
     * 根据流程名称查找字段对照工作表
     * @param wb 工作簿
     * @param flowName 完整的流程名称（如 "axmt540创建出货单-客制新增"）
     * @param progCode 程序代号（如 "axmt540"）
     */
    public String findFieldMappingSheet(Workbook wb, String flowName, String progCode) {
        if (wb == null || progCode == null || progCode.isEmpty()) return null;

        String progCodeLower = progCode.toLowerCase();
        String flowNameLower = flowName != null ? flowName.toLowerCase() : "";
        log.info("查找字段对照工作表, flowName={}, progCode={}", flowName, progCodeLower);
        
        // 1. 优先精确匹配：工作表名 == 流程名称
        if (!flowNameLower.isEmpty()) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                String sheetName = wb.getSheetName(i);
                if (sheetName.equalsIgnoreCase(flowName)) {
                    log.info("精确匹配到工作表(流程名称): {}", sheetName);
                    return sheetName;
                }
            }
        }
        
        // 2. 精确匹配：工作表名 == 程序代号
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String sheetName = wb.getSheetName(i);
            String sheetNameLower = sheetName.toLowerCase();
            if (sheetNameLower.equals(progCodeLower)) {
                log.info("精确匹配到工作表(程序代号): {}", sheetName);
                return sheetName;
            }
        }
        
        // 3. 流程名称包含匹配：工作表名包含流程名称的关键词
        if (!flowNameLower.isEmpty()) {
            // 提取流程名称中程序代号之后的部分作为匹配关键词
            String nameAfterCode = flowNameLower;
            int codeIndex = flowNameLower.indexOf(progCodeLower);
            if (codeIndex >= 0) {
                nameAfterCode = flowNameLower.substring(codeIndex + progCodeLower.length());
            }
            
            // 查找包含这些关键词的工作表
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                String sheetName = wb.getSheetName(i);
                String sheetNameLower = sheetName.toLowerCase();
                if (sheetNameLower.contains(progCodeLower) && sheetNameLower.contains(nameAfterCode)) {
                    log.info("流程名称关键词匹配到工作表: {}", sheetName);
                    return sheetName;
                }
            }
        }
        
        // 4. 包含匹配：工作表名包含程序代号，优先"创建/生单"类型
        List<String> candidateSheets = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String sheetName = wb.getSheetName(i);
            String sheetNameLower = sheetName.toLowerCase();
            if (sheetNameLower.contains(progCodeLower)) {
                candidateSheets.add(sheetName);
            }
        }
        
        if (!candidateSheets.isEmpty()) {
            // 优先返回"创建"类型且不是"过账/更新"类型的工作表
            for (String sheetName : candidateSheets) {
                String sheetNameLower = sheetName.toLowerCase();
                boolean isCreate = sheetNameLower.contains("创建") || sheetNameLower.contains("生单");
                boolean isUpdate = sheetNameLower.contains("过账") || sheetNameLower.contains("更新");
                if (isCreate && !isUpdate) {
                    log.info("包含匹配到工作表(创建类型): {}", sheetName);
                    return sheetName;
                }
            }
            // 其次返回不是"过账/更新"类型的工作表
            for (String sheetName : candidateSheets) {
                String sheetNameLower = sheetName.toLowerCase();
                boolean isUpdate = sheetNameLower.contains("过账") || sheetNameLower.contains("更新");
                if (!isUpdate) {
                    log.info("包含匹配到工作表(非更新类型): {}", sheetName);
                    return sheetName;
                }
            }
            // 否则返回第一个
            log.info("包含匹配到工作表(默认): {}", candidateSheets.get(0));
            return candidateSheets.get(0);
        }
        
        log.warn("未找到匹配的工作表: {}", progCode);
        return null;
    }
    
    /**
     * 根据程序代号查找字段对照工作表（向后兼容）
     */
    public String findFieldMappingSheet(Workbook wb, String progCode) {
        return findFieldMappingSheet(wb, null, progCode);
    }

    /**
     * 获取单元格字符串值
     */
    private String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                double numVal = cell.getNumericCellValue();
                if (numVal == Math.floor(numVal)) {
                    return String.valueOf((long) numVal);
                }
                return String.valueOf(numVal);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (Exception e2) {
                        return "";
                    }
                }
            default:
                return "";
        }
    }

    private String safeStr(String s) {
        return s == null ? "" : s;
    }

    /**
     * 匹配结果
     */
    public static class MatchResult {
        private Long demandId;
        private String programCode;
        private String programName;
        private boolean matched;
        private int score;
        private String matchedFlowName;
        private String interfaceDirection;
        private String inferredFlow;
        private String reason;
        private String referenceSheet;
        private String referenceRow;

        // Getters and Setters
        public Long getDemandId() { return demandId; }
        public void setDemandId(Long demandId) { this.demandId = demandId; }
        public String getProgramCode() { return programCode; }
        public void setProgramCode(String programCode) { this.programCode = programCode; }
        public String getProgramName() { return programName; }
        public void setProgramName(String programName) { this.programName = programName; }
        public boolean isMatched() { return matched; }
        public void setMatched(boolean matched) { this.matched = matched; }
        public int getScore() { return score; }
        public void setScore(int score) { this.score = score; }
        public String getMatchedFlowName() { return matchedFlowName; }
        public void setMatchedFlowName(String matchedFlowName) { this.matchedFlowName = matchedFlowName; }
        public String getInterfaceDirection() { return interfaceDirection; }
        public void setInterfaceDirection(String interfaceDirection) { this.interfaceDirection = interfaceDirection; }
        public String getInferredFlow() { return inferredFlow; }
        public void setInferredFlow(String inferredFlow) { this.inferredFlow = inferredFlow; }
        public String getReason() { return reason; }
        public void setReason(String reason) { this.reason = reason; }
        public String getReferenceSheet() { return referenceSheet; }
        public void setReferenceSheet(String referenceSheet) { this.referenceSheet = referenceSheet; }
        public String getReferenceRow() { return referenceRow; }
        public void setReferenceRow(String referenceRow) { this.referenceRow = referenceRow; }
    }
}
