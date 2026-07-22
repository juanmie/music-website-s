#!/usr/bin/env node
// -*- coding: utf-8 -*-
/**
 * API Document Retrieval V5 - 能率需求清单智能匹配（Node.js + ExcelJS版本）
 * 优化匹配逻辑：
 * 1. 开发分类 → 流向自动推断
 *    - 702/703/704/705: 服务端接口 → XX→T100 (OA/BPM等外部系统推送数据到T100)
 *    - 710/711: 客户端接口 → T100→XX (T100推送数据到OA/BPM等外部系统)
 *    - 202: 基本资料作业 → 根据需求描述推断
 * 2. 程序代号 + 推断流向 → 匹配接口清单"流程名称"
 * 3. 支持 .xlsx 格式（使用 ExcelJS 库，完整保留格式：边框、颜色、字体、合并单元格、列宽、行高）
 * 
 * 用法: node match_nenglv_v5.js <需求清单xlsx文件>
 */

const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// 配置
const WORKSPACE_DIR = process.cwd();
const API_MAPPING_FILE = path.join(WORKSPACE_DIR, '需求对照工具表', '【2026】字段对照---最小颗粒度.xlsx');
const INTERFACE_LIST_FILE = path.join(WORKSPACE_DIR, '需求对照工具表', '【2026】接口清单---汇总.xlsx');
const ONLINE_DOC_URL = 'https://docs.qq.com/sheet/DWmF4RXZmdW5RWnVk?tab=nxp5uc';
const FIELD_MAPPING_DOC_URL = 'https://docs.qq.com/sheet/DWm9PT0xtYWdYRW5R?scene=a815f6149367bdcef5a2dc6d3cXfw1&tab=mcy037';

/**
 * 格式化日期时间戳：yyyy-mm-dd_HH-MM-SS
 */
function formatTimestamp(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hour = String(date.getHours()).padStart(2, '0');
    const minute = String(date.getMinutes()).padStart(2, '0');
    const second = String(date.getSeconds()).padStart(2, '0');
    return `${year}-${month}-${day}_${hour}-${minute}-${second}`;
}

// 开发分类 → 流向映射规则
const FLOW_DIRECTION_RULES = {
    '702': 'XX→T100',
    '703': 'XX→T100',
    '704': 'XX→T100',
    '705': 'XX→T100',
    '710': 'T100→XX',
    '711': 'T100→XX',
    '202': 'T100→XX',
};

/**
 * 从接口清单汇总文件中加载接口列表（用于匹配流程名称和方向）
 */
function loadInterfaceList() {
    console.log(`正在加载接口清单汇总: ${path.basename(INTERFACE_LIST_FILE)}`);
    
    if (!fs.existsSync(INTERFACE_LIST_FILE)) {
        console.error(`错误: 找不到接口清单汇总文件: ${INTERFACE_LIST_FILE}`);
        console.error(`\n建议使用在线文档: ${ONLINE_DOC_URL}`);
        process.exit(1);
    }
    
    const wb = XLSX.readFile(INTERFACE_LIST_FILE);
    const allInterfaces = [];
    
    // 遍历所有工作表（OA接口清单、WMS接口清单、MES接口清单）
    wb.SheetNames.forEach(sheetName => {
        const sheet = wb.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        let headerRowIdx = 0;
        for (let i = 0; i < Math.min(10, data.length); i++) {
            const row = data[i];
            if (row && row.some(cell => cell === '序号' || cell === '流程名称')) {
                headerRowIdx = i;
                break;
            }
        }
        
        const headers = data[headerRowIdx].map((h, i) => ({
            col: i,
            name: h ? String(h).trim() : `Column${i + 1}`
        }));
        
        for (let rowIdx = headerRowIdx + 1; rowIdx < data.length; rowIdx++) {
            const row = data[rowIdx];
            if (!row || row.length === 0) continue;
            
            const rowData = { _row_num: rowIdx + 1, _sheet: sheetName };
            
            headers.forEach(header => {
                const value = row[header.col];
                rowData[header.name] = value !== undefined ? String(value).trim() : '';
            });
            
            if (rowData['流程名称'] && rowData['流程名称'].trim()) {
                allInterfaces.push(rowData);
            }
        }
    });
    
    console.log(`已加载 ${allInterfaces.length} 条接口记录（来自 ${wb.SheetNames.length} 个工作表）\n`);
    return allInterfaces;
}

/**
 * 使用 xlsx 加载API字段对照清单（用于数据匹配）
 * 现在从【2026】字段对照---最小颗粒度.xlsx 读取
 */
function loadApiMapping() {
    console.log(`正在加载API字段对照清单: ${path.basename(API_MAPPING_FILE)}`);
    
    if (!fs.existsSync(API_MAPPING_FILE)) {
        console.error(`错误: 找不到本地API字段对照清单文件: ${API_MAPPING_FILE}`);
        console.error(`\n建议使用在线文档: ${ONLINE_DOC_URL}`);
        process.exit(1);
    }
    
    const wb = XLSX.readFile(API_MAPPING_FILE);
    
    // 新文件结构：每个工作表对应一个流程，工作表名包含程序代号
    // 接口清单数据现在从 INTERFACE_LIST_FILE 加载
    // 这里只返回工作表名称列表，用于后续复制详情页签
    const sheetNames = wb.SheetNames.filter(name => wb.Sheets[name] && 
        XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 }).length > 0);
    
    console.log(`找到 ${sheetNames.length} 个有效字段对照工作表`);
    console.log(`工作表列表: ${sheetNames.join(', ')}\n`);
    
    return { sheetNames, wb };
}

/**
 * 根据开发分类推断流向
 */
function inferFlowDirection(devClass, specDesc) {
    if (!devClass) return '';
    
    const match = String(devClass).match(/(\d{3})/);
    if (match) {
        const classCode = match[1];
        let flow = FLOW_DIRECTION_RULES[classCode] || '';
        
        if (classCode === '202' && specDesc) {
            const desc = String(specDesc).toLowerCase();
            if (desc.includes('查询') || desc.includes('封装') || desc.includes('ipaas') || desc.includes('api')) {
                flow = 'T100→XX';
            }
        }
        
        return flow;
    }
    return '';
}

/**
 * 标准化流向格式
 */
function normalizeFlowDirection(flowDir) {
    if (!flowDir) return '';
    
    let flow = String(flowDir).toUpperCase()
        .replace(/\s/g, '')
        .replace(/--/g, '→')
        .replace(/->/g, '→');
    
    if (flow.startsWith('ERP→') || flow.startsWith('T100→')) {
        return 'T100→XX';
    }
    if (flow.startsWith('OA→') || flow.startsWith('XX→')) {
        return 'XX→T100';
    }
    
    if (flow.includes('→')) {
        const parts = flow.split('→');
        if (parts[0].includes('ERP') || parts[0].includes('T100')) {
            return 'T100→XX';
        } else {
            return 'XX→T100';
        }
    }
    
    return flow;
}

/**
 * 从文本中提取程序代号
 */
function extractProgCode(text) {
    if (!text) return '';
    const match = String(text).toLowerCase().match(/([a-z]{2,4}\d{3})/);
    return match ? match[1] : '';
}

/**
 * 提取业务关键词
 */
function extractBusinessKeywords(text) {
    if (!text) return [];
    
    const keywords = [
        '员工', '部门', '料号', '客户', '供应商', '采购', '销售', '订单',
        '核价', '请购', '入库', '出库', '库存', '杂收', '杂发', '付款',
        '收款', '报销', '借款', '预算', '资产', '理财', '发票', '核销',
        '申请', '维护', '查询', '新增', '修改', '删除', '审核', '送签',
        '推送', '同步', '抛转', '回写', '更新', '创建', '生成'
    ];
    
    const textLower = String(text).toLowerCase();
    return keywords.filter(kw => textLower.includes(kw));
}

/**
 * 匹配接口
 */
function matchInterface(requirement, apiData) {
    const matches = [];
    
    const devClass = requirement['开发分类'] || requirement['开发分类(*)'] || '';
    const progCode = String(requirement['程序代号'] || requirement['程序代号(*)'] || '').toLowerCase().trim();
    const reqName = requirement['程序名称'] || requirement['程序名称(*)'] || '';
    const specDesc = requirement['规格说明'] || requirement['规格说明(*)'] || '';
    
    if (!progCode) {
        return matches;
    }
    
    const inferredFlow = inferFlowDirection(devClass, specDesc);
    requirement._inferredFlow = inferredFlow;
    
    console.log(`  程序代号: ${progCode}, 开发分类: ${devClass}, 推断流向: ${inferredFlow}`);
    
    for (const apiRow of apiData) {
        const flowName = apiRow['流程名称'] || '';
        const apiDirection = apiRow['接口方向'] || '';
        
        if (!flowName) continue;
        
        let score = 0;
        const reasons = [];
        
        const apiProgCode = extractProgCode(flowName);
        if (apiProgCode === progCode) {
            score += 60;
            reasons.push(`程序代号精确匹配: ${progCode}`);
        } else if (flowName.toLowerCase().includes(progCode)) {
            score += 40;
            reasons.push(`程序代号包含匹配: ${progCode}`);
        }
        
        if (inferredFlow && score > 0) {
            const normalizedApiDir = normalizeFlowDirection(apiDirection);
            
            if (inferredFlow === normalizedApiDir) {
                score += 30;
                reasons.push(`流向匹配: ${inferredFlow} ↔ ${apiDirection}`);
            } else if (normalizedApiDir) {
                score -= 20;
                reasons.push(`流向相反: ${inferredFlow} vs ${apiDirection}`);
            }
        }
        
        if (score > 0 && reqName) {
            const reqKeywords = extractBusinessKeywords(reqName);
            const flowKeywords = extractBusinessKeywords(flowName);
            
            const commonKeywords = reqKeywords.filter(kw => flowKeywords.includes(kw));
            if (commonKeywords.length > 0) {
                const keywordScore = Math.min(commonKeywords.length * 5, 10);
                score += keywordScore;
                reasons.push(`业务关键词匹配: ${commonKeywords.join(', ')}`);
            }
        }
        
        if (progCode && String(specDesc).toLowerCase().includes(progCode) && score > 0) {
            score += 5;
            reasons.push('规格说明中提及程序代号');
        }
        
        if (score >= 40) {
            matches.push({
                api: apiRow,
                score: Math.min(score, 100),
                reason: reasons.join(' | '),
                matchType: '流程名称匹配'
            });
        }
    }
    
    matches.sort((a, b) => b.score - a.score);
    
    return matches;
}

/**
 * 读取需求清单文件
 */
function readRequirementFile(filePath) {
    console.log(`读取需求清单: ${path.basename(filePath)}`);
    
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    
    const headers = data[0].map((h, i) => ({
        col: i,
        name: h ? String(h).trim() : `Column${i + 1}`
    }));
    
    const requirements = [];
    for (let rowIdx = 1; rowIdx < data.length; rowIdx++) {
        const row = data[rowIdx];
        const reqData = { _rowNum: rowIdx + 1 };
        
        headers.forEach(header => {
            const value = row[header.col];
            reqData[header.name] = value !== undefined ? String(value).trim() : '';
        });
        
        const progCode = reqData['程序代号'] || reqData['程序代号(*)'];
        if (progCode && progCode.trim()) {
            requirements.push(reqData);
        }
    }
    
    console.log(`工作表: ${wb.SheetNames[0]}`);
    console.log(`提取到 ${requirements.length} 条需求记录\n`);
    return { requirements, headers, wb };
}

/**
 * 使用 ExcelJS 复制工作表，完整保留所有格式
 */
async function copyWorksheetWithFormat(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName) {
    const sourceWs = sourceWorkbook.getWorksheet(sourceSheetName);
    if (!sourceWs) {
        console.log(`  警告: 未找到源工作表: ${sourceSheetName}`);
        return null;
    }
    
    // 创建新工作表
    const targetWs = targetWorkbook.addWorksheet(targetSheetName);
    
    // 复制列属性（列宽、隐藏状态等）
    // 注意：sourceWs.columns 可能包含 null 元素，需要安全处理
    const maxCol = sourceWs.columnCount;
    for (let colIndex = 1; colIndex <= maxCol; colIndex++) {
        const sourceCol = sourceWs.getColumn(colIndex);
        const targetCol = targetWs.getColumn(colIndex);
        
        if (sourceCol) {
            if (sourceCol.width !== undefined) {
                targetCol.width = sourceCol.width;
            }
            if (sourceCol.hidden !== undefined) {
                targetCol.hidden = sourceCol.hidden;
            }
            if (sourceCol.outlineLevel !== undefined) {
                targetCol.outlineLevel = sourceCol.outlineLevel;
            }
        }
    }
    
    // 复制行属性（行高、隐藏状态等）和单元格
    sourceWs.eachRow({ includeEmpty: true }, (sourceRow, rowNumber) => {
        const targetRow = targetWs.getRow(rowNumber);
        
        // 复制行高
        if (sourceRow.height !== undefined && sourceRow.height !== null) {
            targetRow.height = sourceRow.height;
        }
        
        // 复制行隐藏状态
        if (sourceRow.hidden !== undefined) {
            targetRow.hidden = sourceRow.hidden;
        }
        
        // 复制每个单元格
        sourceRow.eachCell({ includeEmpty: true }, (sourceCell, colNumber) => {
            const targetCell = targetRow.getCell(colNumber);
            
            // 复制值（先处理特殊类型）
            if (sourceCell.formula) {
                targetCell.value = { formula: sourceCell.formula, result: sourceCell.result };
            } else if (sourceCell.hyperlink) {
                targetCell.value = { text: sourceCell.text, hyperlink: sourceCell.hyperlink };
            } else {
                targetCell.value = sourceCell.value;
            }
            
            // 复制样式（深拷贝）
            if (sourceCell.style && Object.keys(sourceCell.style).length > 0) {
                targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
            }
            
            // 复制数据验证
            if (sourceCell.dataValidation) {
                targetCell.dataValidation = JSON.parse(JSON.stringify(sourceCell.dataValidation));
            }
        });
    });
    
    // 复制合并单元格
    if (sourceWs._merges) {
        Object.keys(sourceWs._merges).forEach(mergeKey => {
            const mergeRange = sourceWs._merges[mergeKey];
            if (mergeRange && mergeRange.left && mergeRange.top && mergeRange.right && mergeRange.bottom) {
                targetWs.mergeCells(
                    mergeRange.top, 
                    mergeRange.left, 
                    mergeRange.bottom, 
                    mergeRange.right
                );
            }
        });
    }
    
    // 复制工作表属性
    if (sourceWs.properties) {
        try {
            targetWs.properties = JSON.parse(JSON.stringify(sourceWs.properties));
        } catch (e) {
            // 忽略属性复制错误
        }
    }
    
    // 复制页面设置
    if (sourceWs.pageSetup) {
        try {
            targetWs.pageSetup = JSON.parse(JSON.stringify(sourceWs.pageSetup));
        } catch (e) {
            // 忽略页面设置复制错误
        }
    }
    
    // 复制打印区域
    if (sourceWs.printArea) {
        targetWs.printArea = sourceWs.printArea;
    }
    
    console.log(`  已复制工作表: ${sourceSheetName} → ${targetSheetName} (保留格式)`);
    return targetWs;
}

/**
 * 生成输出文件（使用 ExcelJS 保留格式）
 */
async function generateOutput(requirements, results, originalFile, apiMappingData) {
    console.log('\n生成匹配结果文件...');
    
    // 使用 ExcelJS 读取原始文件以保留格式
    const originalWorkbook = new ExcelJS.Workbook();
    await originalWorkbook.xlsx.readFile(originalFile);
    const originalWs = originalWorkbook.getWorksheet(1); // 第一个工作表
    
    // 获取原始表头（从第一行提取）
    const originalHeaders = [];
    const firstRow = originalWs.getRow(1);
    firstRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        originalHeaders[colNumber - 1] = cell.value ? String(cell.value).trim() : '';
    });
    
    // 添加匹配结果列
    const matchHeaders = [
        '匹配状态',
        '匹配流程名称',
        '接口方向',
        '推断流向',
        '匹配度',
        '匹配依据',
        '参考位置',
        '在线文档链接'
    ];
    
    const baseCol = originalHeaders.length + 1;
    
    // 添加匹配结果表头到原始工作表（保留原有格式，新列添加表头）
    matchHeaders.forEach((h, idx) => {
        const cell = originalWs.getRow(1).getCell(baseCol + idx);
        cell.value = h;
        // 给新表头添加基本样式（与原始表头一致）
        const firstHeaderCell = originalWs.getRow(1).getCell(1);
        if (firstHeaderCell.style) {
            cell.style = JSON.parse(JSON.stringify(firstHeaderCell.style));
        }
    });
    
    // 填充数据行
    for (let i = 0; i < results.length; i++) {
        const result = results[i];
        const rowIdx = result.requirement._rowNum;
        const row = originalWs.getRow(rowIdx);
        
        if (result.api && result.score >= 40) {
            row.getCell(baseCol).value = '✅ 已匹配';
            row.getCell(baseCol + 1).value = result.api['流程名称'] || '';
            row.getCell(baseCol + 2).value = result.api['接口方向'] || '';
            row.getCell(baseCol + 3).value = result.inferredFlow || '';
            row.getCell(baseCol + 4).value = `${result.score}%`;
            row.getCell(baseCol + 5).value = result.reason;
            row.getCell(baseCol + 6).value = `${result.api._sheet} 行${result.api._row_num}`;
            row.getCell(baseCol + 7).value = ONLINE_DOC_URL;
        } else {
            row.getCell(baseCol).value = '❌ 未匹配';
            row.getCell(baseCol + 1).value = '';
            row.getCell(baseCol + 2).value = '';
            row.getCell(baseCol + 3).value = result.inferredFlow || '';
            row.getCell(baseCol + 4).value = '0%';
            row.getCell(baseCol + 5).value = result.reason || '无匹配接口';
            row.getCell(baseCol + 6).value = '';
            row.getCell(baseCol + 7).value = ONLINE_DOC_URL;
        }
    }
    
    // 设置新列的列宽
    matchHeaders.forEach((h, idx) => {
        originalWs.getColumn(baseCol + idx).width = 25;
    });
    
    // 创建新的输出工作簿，将原始工作表作为第一个sheet
    const outputWorkbook = new ExcelJS.Workbook();
    
    // 复制原始工作表到输出工作簿（保留所有格式）
    await copyWorksheetWithFormat(originalWorkbook, originalWorkbook.worksheets[0].name, outputWorkbook, '需求匹配结果');
    
    // 使用 ExcelJS 读取API字段对照清单（用于复制匹配的工作表，保留样式）
    const apiWorkbook = new ExcelJS.Workbook();
    await apiWorkbook.xlsx.readFile(API_MAPPING_FILE);
    
    // 复制匹配的工作表到输出文件
    console.log('\n复制匹配的工作表...');
    let copiedCount = 0;
    const copiedSheets = new Set();
    
    for (const result of results) {
        if (result.api && result.score >= 40) {
            const flowName = result.api['流程名称'];
            if (!flowName || copiedSheets.has(flowName)) continue;
            
            // 从流程名称提取程序代号，匹配字段对照表的工作表
            const flowProgCode = extractProgCode(flowName);
            if (!flowProgCode) {
                console.log(`  警告: 无法从流程名称提取程序代号: ${flowName}`);
                continue;
            }
            
            // 查找匹配的字段对照工作表（工作表名包含程序代号）
            const sourceSheetName = apiWorkbook.worksheets.find(ws => {
                const wsNameLower = ws.name.toLowerCase();
                return wsNameLower.includes(flowProgCode);
            })?.name;
            
            if (!sourceSheetName) {
                console.log(`  警告: 未找到对应字段对照工作表: ${flowName} (程序代号: ${flowProgCode})`);
                continue;
            }
            
            // 工作表名称限制31字符
            let targetSheetName = sourceSheetName;
            if (targetSheetName.length > 31) {
                targetSheetName = targetSheetName.substring(0, 28) + '...';
            }
            
            // 确保名称唯一
            let finalName = targetSheetName;
            let suffix = 1;
            while (outputWorkbook.getWorksheet(finalName)) {
                finalName = targetSheetName.substring(0, 29) + '_' + suffix;
                suffix++;
            }
            
            // 复制工作表（保留所有格式）
            await copyWorksheetWithFormat(apiWorkbook, sourceSheetName, outputWorkbook, finalName);
            
            copiedSheets.add(flowName);
            copiedCount++;
        }
    }
    
    console.log(`共复制 ${copiedCount} 个工作表\n`);
    
    // 确定输出目录
    const OUTPUT_DIR = path.join(WORKSPACE_DIR, '集成需求字段清单产出');
    if (!fs.existsSync(OUTPUT_DIR)) {
        fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    }
    
    const timestamp = formatTimestamp(new Date());
    
    const outputFile = path.join(
        OUTPUT_DIR,
        `${path.basename(originalFile, path.extname(originalFile))}_API匹配结果_${timestamp}.xlsx`
    );
    
    await outputWorkbook.xlsx.writeFile(outputFile);
    
    const stats = fs.statSync(outputFile);
    console.log(`匹配结果已保存: ${path.basename(outputFile)}`);
    console.log(`文件大小: ${(stats.size / 1024).toFixed(2)} KB\n`);
    
    return outputFile;
}

/**
 * 主函数
 */
async function main() {
    console.log('='.repeat(80));
    console.log('     API Document Retrieval V5 - 能率需求清单智能匹配');
    console.log('     优化: 开发分类自动推断流向 + 程序代号+流向匹配流程名称');
    console.log('     格式保留: 使用 ExcelJS 完整保留边框/颜色/字体/合并单元格/列宽/行高');
    console.log('='.repeat(80));
    console.log();
    
    const args = process.argv.slice(2);
    if (args.length < 1) {
        console.log('用法: node match_nenglv_v5.js <需求清单xlsx文件>');
        console.log('\n支持格式: .xlsx');
        console.log('\n示例:');
        console.log('  node match_nenglv_v5.js 需求评估整理/能率需求清单-bpm集成.xlsx');
        console.log('  node match_nenglv_v5.js 需求评估整理/能率需求清单-国家平台集成.xlsx');
        process.exit(1);
    }
    
    const requirementFile = args[0];
    const filePath = path.resolve(requirementFile);
    
    if (!fs.existsSync(filePath)) {
        console.error(`错误: 找不到文件 ${requirementFile}`);
        process.exit(1);
    }
    
    if (!filePath.endsWith('.xlsx')) {
        console.error('错误: 仅支持 Excel 文件 (.xlsx)');
        process.exit(1);
    }
    
    // 执行智能匹配
    const apiMappingData = loadApiMapping();
    const apiData = loadInterfaceList();
    
    console.log(`\n📖 处理需求清单: ${path.basename(filePath)}`);
    console.log('='.repeat(80));
    
    const { requirements, headers } = readRequirementFile(filePath);
    
    if (requirements.length === 0) {
        console.log('未找到有效的需求记录');
        process.exit(1);
    }
    
    // 匹配接口
    console.log('开始匹配API...');
    console.log('='.repeat(80));
    
    const results = [];
    let matchedCount = 0;
    
    for (let i = 0; i < requirements.length; i++) {
        const req = requirements[i];
        const progCode = req['程序代号(*)'] || req['程序代号'] || '';
        const reqName = req['程序名称(*)'] || req['程序名称'] || '';
        const devClass = req['开发分类(*)'] || req['开发分类'] || '';
        
        console.log(`[${i + 1}/${requirements.length}] ${progCode} - ${reqName.substring(0, 40)}...`);
        console.log(`  开发分类: ${devClass}`);
        
        const matches = matchInterface(req, apiData);
        
        if (matches.length > 0) {
            const bestMatch = matches[0];
            matchedCount++;
            console.log(`  [✓] 匹配成功 (${bestMatch.score}%) - ${bestMatch.api['流程名称']}`);
            console.log(`  依据: ${bestMatch.reason}\n`);
            
            results.push({
                requirement: req,
                api: bestMatch.api,
                score: bestMatch.score,
                reason: bestMatch.reason,
                inferredFlow: req._inferredFlow
            });
        } else {
            console.log(`  [✗] 未找到匹配\n`);
            results.push({
                requirement: req,
                api: null,
                score: 0,
                reason: '',
                inferredFlow: req._inferredFlow
            });
        }
    }
    
    console.log('='.repeat(80));
    console.log(`匹配完成: ${matchedCount}/${requirements.length} 条需求找到匹配API\n`);
    
    // 检查匹配率
    const matchRate = (matchedCount / requirements.length) * 100;
    if (matchRate < 50) {
        console.log('⚠️ 警告: 匹配率低于 50%，建议更新本地 API 对照清单');
        console.log(`\n当前本地文件: ${API_MAPPING_FILE}`);
        console.log(`在线文档地址: ${ONLINE_DOC_URL}`);
        console.log('\n📥 请手动下载最新版本:');
        console.log('  1. 访问上述在线文档地址');
        console.log('  2. 点击右上角「文件」→「导出为」→「Excel(.xlsx)」');
        console.log('  3. 保存到: 需求对照工具表/【2026】字段对照---最小颗粒度.xlsx');
        console.log('  4. 保存到: 需求对照工具表/【2026】接口清单---汇总.xlsx');
        console.log('\n尝试自动打开浏览器...\n');
        
        try {
            const { exec } = require('child_process');
            const platform = process.platform;
            if (platform === 'win32') {
                exec(`start "" "${ONLINE_DOC_URL}"`);
            } else if (platform === 'darwin') {
                exec(`open "${ONLINE_DOC_URL}"`);
            } else {
                exec(`xdg-open "${ONLINE_DOC_URL}"`);
            }
            console.log('✓ 已自动打开浏览器访问在线文档\n');
        } catch (e) {
            console.log('✗ 自动打开浏览器失败，请手动访问上述地址\n');
        }
    }
    
    // 生成结果文件
    const outputFile = await generateOutput(requirements, results, filePath, apiMappingData);
    
    console.log('='.repeat(80));
    console.log('智能匹配完成！');
    console.log('='.repeat(80));
    console.log(`\n输出文件: ${outputFile}`);
    console.log(`在线文档: ${ONLINE_DOC_URL}`);
    console.log(`\n请打开输出文件查看匹配结果！`);
}

// 运行主函数
main().catch(err => {
    console.error('发生错误:', err);
    process.exit(1);
});
