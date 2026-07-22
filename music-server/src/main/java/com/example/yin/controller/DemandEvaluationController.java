package com.example.yin.controller;

import com.example.yin.common.R;
import com.example.yin.model.domain.DemandEvaluation;
import com.example.yin.service.ApiMatchService;
import com.example.yin.service.DemandEvaluationService;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.type.CollectionType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.time.LocalDate;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 需求评估控制器
 */
@RestController
@RequestMapping("/demandEvaluation")
public class DemandEvaluationController {

    private static final Logger log = LoggerFactory.getLogger(DemandEvaluationController.class);
    private static final Pattern PROG_CODE_PATTERN = Pattern.compile("([a-z]{2,4}\\d{3})");

    @Autowired
    private DemandEvaluationService demandEvaluationService;

    @Autowired
    private ApiMatchService apiMatchService;

    // 返回所有需求
    @GetMapping
    public R getAllDemand() {
        return demandEvaluationService.getAllDemand();
    }

    // 根据ID获取需求详情
    @GetMapping("/detail")
    public R getDemandOfId(@RequestParam Long id) {
        return demandEvaluationService.getDemandOfId(id);
    }

    // 添加需求
    @PostMapping("/add")
    public R addDemand(@RequestBody DemandEvaluation demand) {
        return demandEvaluationService.addDemand(demand);
    }

    // 更新需求信息
    @PostMapping("/update")
    public R updateDemand(@RequestBody DemandEvaluation demand) {
        return demandEvaluationService.updateDemand(demand);
    }

    // 删除需求
    @GetMapping("/delete")
    public R deleteDemand(@RequestParam Long id) {
        return demandEvaluationService.deleteDemand(id);
    }

    /**
     * 导出需求清单 + API匹配结果
     * 输出格式：适配AI输出模板（输入物+集成配置+开发计划+问题跟进+字段对照）
     * 支持可选的 AI 分析数据（aiData），有则追加 AI 分析 sheets
     * @param requestBody 包含 demands（需求列表）和可选 aiData（AI分析结果）
     */
    @PostMapping("/exportWithApiMatch")
    public void exportWithApiMatch(@RequestBody(required = false) Object body, HttpServletResponse response) {
        long totalStart = System.currentTimeMillis();
        try {
            // 解析请求参数（兼容旧格式 List<DemandEvaluation> 和新格式 {demands, aiData}）
            List<DemandEvaluation> demands = null;
            Map<String, Object> aiData = null;
            ObjectMapper objectMapper = new ObjectMapper();
            
            if (body instanceof Map) {
                @SuppressWarnings("unchecked")
                Map<String, Object> requestBody = (Map<String, Object>) body;
                // 尝试提取 aiData
                Object aiDataObj = requestBody.get("aiData");
                if (aiDataObj instanceof Map) {
                    @SuppressWarnings("unchecked")
                    Map<String, Object> temp = (Map<String, Object>) aiDataObj;
                    aiData = temp;
                }
                // 提取 demands
                Object demandsObj = requestBody.get("demands");
                if (demandsObj instanceof List) {
                    CollectionType listType = objectMapper.getTypeFactory()
                        .constructCollectionType(List.class, DemandEvaluation.class);
                    demands = objectMapper.convertValue(demandsObj, listType);
                }
            } else if (body instanceof List) {
                CollectionType listType = objectMapper.getTypeFactory()
                    .constructCollectionType(List.class, DemandEvaluation.class);
                demands = objectMapper.convertValue(body, listType);
            }
            
            // 1. 获取需求（若前端未传入则取全部）
            long stepStart = System.currentTimeMillis();
            if (demands == null || demands.isEmpty()) {
                R result = demandEvaluationService.getAllDemand();
                @SuppressWarnings("unchecked")
                List<DemandEvaluation> allDemands = (List<DemandEvaluation>) result.getData();
                demands = allDemands;
            }
            log.info("[导出耗时] 获取需求: {}ms", System.currentTimeMillis() - stepStart);
            
            if (demands == null || demands.isEmpty()) {
                response.setContentType("application/json;charset=UTF-8");
                response.getWriter().write("{\"error\":\"没有需求数据可导出\"}");
                return;
            }

            // 2. 执行 API 匹配
            stepStart = System.currentTimeMillis();
            List<ApiMatchService.MatchResult> matchResults = apiMatchService.matchAll(demands);
            log.info("[导出耗时] API匹配: {}ms", System.currentTimeMillis() - stepStart);

            // 3. 生成 Excel
            stepStart = System.currentTimeMillis();
            Workbook wb = new XSSFWorkbook();
            
            // Sheet 1: 输入物-TB导入EXCEL摸板（需求匹配结果）
            Sheet resultSheet = wb.createSheet("输入物-TB导入EXCEL摸板");
            writeMatchResultSheet(resultSheet, demands, matchResults);
            log.info("[导出耗时] 写入匹配结果Sheet: {}ms", System.currentTimeMillis() - stepStart);
            
            // Sheet 2-4: 从 AI输出物摸板 复制标准工作表
            stepStart = System.currentTimeMillis();
            Workbook outputTemplateWb = apiMatchService.getOutputTemplateWorkbook();
            if (outputTemplateWb != null) {
                for (String sheetName : ApiMatchService.TEMPLATE_SHEETS_TO_COPY) {
                    if (outputTemplateWb.getSheet(sheetName) != null) {
                        copySheet(outputTemplateWb, sheetName, wb, sheetName);
                        log.debug("已复制模板工作表: {}", sheetName);
                    } else {
                        log.warn("模板中未找到工作表: {}", sheetName);
                    }
                }
                // 不关闭，保持缓存
            }
            log.info("[导出耗时] 复制模板工作表: {}ms", System.currentTimeMillis() - stepStart);
            
            // Sheet 5~N: 匹配到的字段对照详情
            stepStart = System.currentTimeMillis();
            Workbook fieldMappingWb = apiMatchService.getFieldMappingWorkbook();
            Set<String> copiedSheets = new HashSet<>();
            
            for (ApiMatchService.MatchResult mr : matchResults) {
                if (!mr.isMatched() || mr.getScore() < 40) continue;
                
                String flowName = mr.getMatchedFlowName();
                if (flowName == null || copiedSheets.contains(flowName)) continue;
                
                String progCode = extractProgCode(flowName);
                if (progCode.isEmpty()) continue;
                
                String sourceSheetName = apiMatchService.findFieldMappingSheet(fieldMappingWb, flowName, progCode);
                if (sourceSheetName == null) continue;
                
                String targetName = sourceSheetName.length() > 31 
                    ? sourceSheetName.substring(0, 28) + "..." 
                    : sourceSheetName;
                
                String finalName = targetName;
                int suffix = 1;
                while (wb.getSheet(finalName) != null) {
                    finalName = targetName.substring(0, Math.min(29, targetName.length())) + "_" + suffix;
                    suffix++;
                }
                
                copySheet(fieldMappingWb, sourceSheetName, wb, finalName);
                copiedSheets.add(flowName);
            }
            log.info("[导出耗时] 复制字段对照工作表: {}ms", System.currentTimeMillis() - stepStart);
            
            // 4. 如果有 AI 分析数据，追加 AI 分析 sheets
            if (aiData != null && !aiData.isEmpty()) {
                stepStart = System.currentTimeMillis();
                writeAiSheets(wb, aiData);
                log.info("[导出耗时] 写入AI分析sheets: {}ms", System.currentTimeMillis() - stepStart);
            }
            
            // 5. 输出到响应
            stepStart = System.currentTimeMillis();
            String fileName = "需求清单_API匹配结果_" + LocalDate.now() + ".xlsx";
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition", "attachment;filename=" + 
                URLEncoder.encode(fileName, "UTF-8"));
            
            wb.write(response.getOutputStream());
            wb.close();
            log.info("[导出耗时] 写入响应流: {}ms", System.currentTimeMillis() - stepStart);
            
            log.info("[导出总耗时] {}ms, 共{}条需求, 匹配{}条, 复制字段对照{}个", 
                System.currentTimeMillis() - totalStart, demands.size(), 
                matchResults.stream().filter(ApiMatchService.MatchResult::isMatched).count(), copiedSheets.size());
            
        } catch (Exception e) {
            log.error("导出失败: {}", e.getMessage(), e);
            try {
                response.setContentType("application/json;charset=UTF-8");
                response.getWriter().write("{\"error\":\"导出失败: " + e.getMessage() + "\"}");
            } catch (IOException ignored) {}
        }
    }

    /**
     * 写入匹配结果工作表
     */
    private void writeMatchResultSheet(Sheet sheet, List<DemandEvaluation> demands, 
                                        List<ApiMatchService.MatchResult> matchResults) {
        // 表头样式
        CellStyle headerStyle = sheet.getWorkbook().createCellStyle();
        Font headerFont = sheet.getWorkbook().createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);

        // 表头
        String[] headers = {
            "序号", "程序代号", "程序名称", "系统", "规格说明", "客制类型", "难度",
            "开发分类", "集成产品", "急单", "提出人", "状态", "TB单号",
            "匹配状态", "匹配流程名称", "接口方向", "推断流向", "匹配度", "匹配依据", "参考位置"
        };
        
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        // 数据行样式
        CellStyle dataStyle = sheet.getWorkbook().createCellStyle();
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setWrapText(true);

        // 匹配成功样式（绿色）
        CellStyle matchOkStyle = sheet.getWorkbook().createCellStyle();
        matchOkStyle.cloneStyleFrom(dataStyle);
        Font greenFont = sheet.getWorkbook().createFont();
        greenFont.setColor(IndexedColors.GREEN.getIndex());
        matchOkStyle.setFont(greenFont);

        // 匹配失败样式（红色）
        CellStyle matchFailStyle = sheet.getWorkbook().createCellStyle();
        matchFailStyle.cloneStyleFrom(dataStyle);
        Font redFont = sheet.getWorkbook().createFont();
        redFont.setColor(IndexedColors.RED.getIndex());
        matchFailStyle.setFont(redFont);

        // 填充数据
        for (int i = 0; i < demands.size(); i++) {
            DemandEvaluation d = demands.get(i);
            ApiMatchService.MatchResult mr = matchResults.get(i);
            
            Row row = sheet.createRow(i + 1);
            
            // 原始需求数据
            createCell(row, 0, d.getSeqNo() != null ? d.getSeqNo().toString() : "", dataStyle);
            createCell(row, 1, d.getProgramCode(), dataStyle);
            createCell(row, 2, d.getProgramName(), dataStyle);
            createCell(row, 3, d.getSystemCode(), dataStyle);
            createCell(row, 4, d.getSpecDesc(), dataStyle);
            createCell(row, 5, d.getCustomType(), dataStyle);
            createCell(row, 6, d.getDifficultyLevel(), dataStyle);
            createCell(row, 7, d.getDevCategory(), dataStyle);
            createCell(row, 8, d.getIntegrateProductCode(), dataStyle);
            createCell(row, 9, d.getUrgentFlag(), dataStyle);
            createCell(row, 10, d.getDemandProposer(), dataStyle);
            createCell(row, 11, d.getDemandStatus(), dataStyle);
            createCell(row, 12, d.getTbBillno(), dataStyle);
            
            // 匹配结果
            CellStyle matchStyle = mr.isMatched() ? matchOkStyle : matchFailStyle;
            createCell(row, 13, mr.isMatched() ? "✅ 已匹配" : "❌ 未匹配", matchStyle);
            createCell(row, 14, mr.getMatchedFlowName(), dataStyle);
            createCell(row, 15, mr.getInterfaceDirection(), dataStyle);
            createCell(row, 16, mr.getInferredFlow(), dataStyle);
            createCell(row, 17, mr.getScore() + "%", dataStyle);
            createCell(row, 18, mr.getReason(), dataStyle);
            createCell(row, 19, mr.getReferenceSheet() + " 行" + mr.getReferenceRow(), dataStyle);
        }

        // 设置列宽
        int[] widths = {6, 12, 20, 8, 40, 12, 15, 18, 18, 6, 10, 10, 15, 12, 25, 15, 12, 8, 40, 20};
        for (int i = 0; i < widths.length; i++) {
            sheet.setColumnWidth(i, widths[i] * 256);
        }
    }

    private void createCell(Row row, int col, String value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }

    /**
     * 写入 AI 分析 sheets（智能汇总 + 需求分析详情）
     */
    @SuppressWarnings("unchecked")
    private void writeAiSheets(Workbook wb, Map<String, Object> aiData) {
        // Sheet: AI智能汇总
        Object summaryObj = aiData.get("summary");
        if (summaryObj instanceof Map) {
            Map<String, Object> summary = (Map<String, Object>) summaryObj;
            Sheet summarySheet = wb.createSheet("AI智能汇总");
            
            CellStyle headerStyle = wb.createCellStyle();
            Font headerFont = wb.createFont();
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            headerStyle.setBorderTop(BorderStyle.THIN);
            headerStyle.setBorderLeft(BorderStyle.THIN);
            headerStyle.setBorderRight(BorderStyle.THIN);
            
            CellStyle dataStyle = wb.createCellStyle();
            dataStyle.setBorderBottom(BorderStyle.THIN);
            dataStyle.setBorderTop(BorderStyle.THIN);
            dataStyle.setBorderLeft(BorderStyle.THIN);
            dataStyle.setBorderRight(BorderStyle.THIN);
            dataStyle.setWrapText(true);
            
            Row headerRow = summarySheet.createRow(0);
            createCell(headerRow, 0, "项目", headerStyle);
            createCell(headerRow, 1, "值", headerStyle);
            
            int rowIdx = 1;
            rowIdx = addSummaryRow(summarySheet, rowIdx, "需求总数", String.valueOf(summary.getOrDefault("total", "")), dataStyle);
            rowIdx = addSummaryRow(summarySheet, rowIdx, "匹配成功", String.valueOf(summary.getOrDefault("matched", "")), dataStyle);
            rowIdx = addSummaryRow(summarySheet, rowIdx, "未匹配", String.valueOf(summary.getOrDefault("unmatched", "")), dataStyle);
            rowIdx = addSummaryRow(summarySheet, rowIdx, "客制新增", String.valueOf(summary.getOrDefault("customNewCount", "")), dataStyle);
            rowIdx = addSummaryRow(summarySheet, rowIdx, "客制更新", String.valueOf(summary.getOrDefault("customUpdateCount", "")), dataStyle);
            rowIdx = addSummaryRow(summarySheet, rowIdx, "分析摘要", String.valueOf(summary.getOrDefault("summary", "")), dataStyle);
            
            // 添加分类统计
            Object byCustomType = summary.get("byCustomType");
            if (byCustomType instanceof Map) {
                for (Map.Entry<String, Object> entry : ((Map<String, Object>) byCustomType).entrySet()) {
                    rowIdx = addSummaryRow(summarySheet, rowIdx, "客制类型-" + entry.getKey(), String.valueOf(entry.getValue()), dataStyle);
                }
            }
            Object byDifficulty = summary.get("byDifficulty");
            if (byDifficulty instanceof Map) {
                for (Map.Entry<String, Object> entry : ((Map<String, Object>) byDifficulty).entrySet()) {
                    rowIdx = addSummaryRow(summarySheet, rowIdx, "难度等级-" + entry.getKey(), String.valueOf(entry.getValue()), dataStyle);
                }
            }
            
            summarySheet.setColumnWidth(0, 20 * 256);
            summarySheet.setColumnWidth(1, 60 * 256);
        }
        
        // Sheet: AI需求分析
        Object analysesObj = aiData.get("demandAnalyses");
        if (analysesObj instanceof List) {
            List<Map<String, Object>> analyses = (List<Map<String, Object>>) analysesObj;
            if (!analyses.isEmpty()) {
                Sheet analysisSheet = wb.createSheet("AI需求分析");
                
                CellStyle headerStyle = wb.createCellStyle();
                Font headerFont = wb.createFont();
                headerFont.setBold(true);
                headerStyle.setFont(headerFont);
                headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                headerStyle.setBorderBottom(BorderStyle.THIN);
                headerStyle.setBorderTop(BorderStyle.THIN);
                headerStyle.setBorderLeft(BorderStyle.THIN);
                headerStyle.setBorderRight(BorderStyle.THIN);
                
                CellStyle dataStyle = wb.createCellStyle();
                dataStyle.setBorderBottom(BorderStyle.THIN);
                dataStyle.setBorderTop(BorderStyle.THIN);
                dataStyle.setBorderLeft(BorderStyle.THIN);
                dataStyle.setBorderRight(BorderStyle.THIN);
                dataStyle.setWrapText(true);
                
                String[] headers = {"ID", "程序代号", "程序名称", "AI建议-客制类型", "AI建议-难度等级",
                    "AI估算-计费时数", "AI估算-派工时数", "AI估算-依据", "AI完整性评分", "AI优化建议"};
                Row headerRow = analysisSheet.createRow(0);
                for (int i = 0; i < headers.length; i++) {
                    createCell(headerRow, i, headers[i], headerStyle);
                }
                
                for (int i = 0; i < analyses.size(); i++) {
                    Map<String, Object> item = analyses.get(i);
                    Row row = analysisSheet.createRow(i + 1);
                    
                    Map<String, Object> classification = item.get("classification") instanceof Map ? (Map<String, Object>) item.get("classification") : Collections.emptyMap();
                    Map<String, Object> hourEstimate = item.get("hourEstimate") instanceof Map ? (Map<String, Object>) item.get("hourEstimate") : Collections.emptyMap();
                    Map<String, Object> optimization = item.get("optimization") instanceof Map ? (Map<String, Object>) item.get("optimization") : Collections.emptyMap();
                    
                    String suggestions = "";
                    Object sugObj = optimization.get("suggestions");
                    if (sugObj instanceof List) {
                        suggestions = String.join("; ", (List<String>) sugObj);
                    } else if (sugObj instanceof String) {
                        suggestions = (String) sugObj;
                    }
                    
                    createCell(row, 0, String.valueOf(item.getOrDefault("id", "")), dataStyle);
                    createCell(row, 1, String.valueOf(item.getOrDefault("programCode", "")), dataStyle);
                    createCell(row, 2, String.valueOf(item.getOrDefault("programName", "")), dataStyle);
                    createCell(row, 3, String.valueOf(classification.getOrDefault("customType", "")), dataStyle);
                    createCell(row, 4, String.valueOf(classification.getOrDefault("difficultyLevel", "")), dataStyle);
                    createCell(row, 5, String.valueOf(hourEstimate.getOrDefault("billingHoursCustomer", "")), dataStyle);
                    createCell(row, 6, String.valueOf(hourEstimate.getOrDefault("dispatchHours", "")), dataStyle);
                    String reason = String.valueOf(hourEstimate.getOrDefault("reason", hourEstimate.getOrDefault("estimateReason", "")));
                    createCell(row, 7, reason, dataStyle);
                    createCell(row, 8, String.valueOf(optimization.getOrDefault("completenessScore", "")), dataStyle);
                    createCell(row, 9, suggestions, dataStyle);
                }
                
                int[] widths = {8, 15, 20, 15, 15, 12, 12, 40, 12, 50};
                for (int i = 0; i < widths.length; i++) {
                    analysisSheet.setColumnWidth(i, widths[i] * 256);
                }
            }
        }
    }
    
    private int addSummaryRow(Sheet sheet, int rowIdx, String label, String value, CellStyle style) {
        Row row = sheet.createRow(rowIdx);
        createCell(row, 0, label, style);
        createCell(row, 1, value, style);
        return rowIdx + 1;
    }

    /**
     * 复制工作表（保留格式，使用样式缓存避免重复创建）
     */
    private void copySheet(Workbook sourceWb, String sourceName, Workbook targetWb, String targetName) {
        Sheet sourceSheet = sourceWb.getSheet(sourceName);
        if (sourceSheet == null) return;
        
        Sheet targetSheet = targetWb.createSheet(targetName);
        
        // 样式缓存：源样式索引 -> 目标样式，避免每个单元格都创建新样式
        Map<Integer, CellStyle> styleCache = new HashMap<>();
        
        // 复制列宽
        if (sourceSheet.getRow(0) != null) {
            Row firstRow = sourceSheet.getRow(0);
            for (int i = 0; i < firstRow.getLastCellNum(); i++) {
                targetSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
            }
        }
        
        // 复制行和单元格
        for (int rowIdx = 0; rowIdx <= sourceSheet.getLastRowNum(); rowIdx++) {
            Row sourceRow = sourceSheet.getRow(rowIdx);
            if (sourceRow == null) continue;
            
            Row targetRow = targetSheet.createRow(rowIdx);
            if (sourceRow.getHeight() > 0) {
                targetRow.setHeight(sourceRow.getHeight());
            }
            
            for (int colIdx = 0; colIdx < sourceRow.getLastCellNum(); colIdx++) {
                Cell sourceCell = sourceRow.getCell(colIdx);
                if (sourceCell == null) continue;
                
                Cell targetCell = targetRow.createCell(colIdx);
                copyCellValue(sourceCell, targetCell);
                
                // 复制样式（使用缓存，相同源样式只创建一次）
                CellStyle srcStyle = sourceCell.getCellStyle();
                if (srcStyle != null) {
                    int srcStyleIdx = srcStyle.getIndex();
                    CellStyle targetStyle = styleCache.computeIfAbsent(srcStyleIdx, k -> {
                        CellStyle ns = targetWb.createCellStyle();
                        ns.cloneStyleFrom(srcStyle);
                        return ns;
                    });
                    targetCell.setCellStyle(targetStyle);
                }
            }
        }
        
        // 复制合并单元格
        for (int i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            targetSheet.addMergedRegion(sourceSheet.getMergedRegion(i));
        }
        
        log.debug("已复制工作表: {} → {}, 样式缓存大小: {}", sourceName, targetName, styleCache.size());
    }

    /**
     * 复制单元格值
     */
    private void copyCellValue(Cell source, Cell target) {
        switch (source.getCellType()) {
            case STRING:
                target.setCellValue(source.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(source)) {
                    target.setCellValue(source.getDateCellValue());
                } else {
                    target.setCellValue(source.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                target.setCellValue(source.getBooleanCellValue());
                break;
            case FORMULA:
                target.setCellFormula(source.getCellFormula());
                break;
            case BLANK:
                target.setBlank();
                break;
            default:
                break;
        }
    }

    private String extractProgCode(String text) {
        if (text == null || text.isEmpty()) return "";
        Matcher matcher = PROG_CODE_PATTERN.matcher(text.toLowerCase());
        return matcher.find() ? matcher.group(1) : "";
    }
}
