#!/usr/bin/env node
// -*- coding: utf-8 -*-
/**
 * API Document Retrieval V4 - 能率需求清单智能匹配（Node.js版本）
 * 优化匹配逻辑：
 * 1. 开发分类 → 流向自动推断
 *    - 702/703/704/705: 服务端接口 → XX→T100 (OA/BPM等外部系统推送数据到T100)
 *    - 710/711: 客户端接口 → T100→XX (T100推送数据到OA/BPM等外部系统)
 *    - 202: 基本资料作业 → 根据需求描述推断
 * 2. 程序代号 + 推断流向 → 匹配接口清单"流程名称"
 * 3. 支持 .xlsx 格式（使用 xlsx 库）
 * 
 * 用法: node match_nenglv_v4.js <需求清单xlsx文件>
 */

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// 配置 xlsx 库以保留单元格样式
XLSX.set_fs ? XLSX.set_fs(fs) : null;

// 配置
const WORKSPACE_DIR = process.cwd();
const API_MAPPING_FILE = path.join(WORKSPACE_DIR, '需求对照工具表', '【2026】T100&OA类系统集成清单字段对照表.xlsx');
const ONLINE_DOC_URL = 'https://docs.qq.com/sheet/DSmdhek1TT3BPenRa?tab=BB08J2';

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
    // 服务端接口：外部系统 → T100
    '702': 'XX→T100',  // 服务端_生单接口 (BPM/OA推送数据到T100生成单据)
    '703': 'XX→T100',  // 服务端_更新接口
    '704': 'XX→T100',  // 服务端_更新接口(审核) (BPM回写审批状态到T100)
    '705': 'XX→T100',  // 服务端_其他接口
    // 客户端接口：T100 → 外部系统
    '710': 'T100→XX',  // 客户端_推送接口 (T100推送数据到BPM/OA)
    '711': 'T100→XX',  // 客户端_其他推送
    // 基本资料作业：根据需求描述推断，默认T100提供数据
    '202': 'T100→XX',  // 基本资料作业 (通常是T100提供数据查询接口给外部系统)
};

/**
 * 加载API对照清单
 */
function loadApiMapping() {
    console.log(`正在加载API对照清单: ${path.basename(API_MAPPING_FILE)}`);
    
    if (!fs.existsSync(API_MAPPING_FILE)) {
        console.error(`错误: 找不到本地API对照清单文件: ${API_MAPPING_FILE}`);
        console.error(`\n建议使用在线文档: ${ONLINE_DOC_URL}`);
        process.exit(1);
    }
    
    const wb = XLSX.readFile(API_MAPPING_FILE, { cellStyles: true, cellNF: true, cellDates: true });
    
    // 查找"接口清单"页签
    let sheetName = '接口清单';
    if (!wb.SheetNames.includes(sheetName)) {
        console.log(`未找到'接口清单'页签，使用第一个页签: ${wb.SheetNames[0]}`);
        sheetName = wb.SheetNames[0];
    } else {
        console.log(`找到接口清单页签: ${sheetName}`);
    }
    
    const sheet = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    // 找到真正的表头行（包含"序号"或"流程名称"的行）
    let headerRowIdx = 0;
    for (let i = 0; i < Math.min(10, data.length); i++) {
        const row = data[i];
        if (row && row.some(cell => cell === '序号' || cell === '流程名称')) {
            headerRowIdx = i;
            break;
        }
    }
    
    console.log(`表头在第 ${headerRowIdx + 1} 行`);
    
    // 读取表头
    const headers = data[headerRowIdx].map((h, i) => ({
        col: i,
        name: h ? String(h).trim() : `Column${i + 1}`
    }));
    
    console.log(`表头字段: ${headers.slice(0, 8).map(h => h.name).join(', ')}`);
    
    // 读取所有行数据
    const apiData = [];
    for (let rowIdx = headerRowIdx + 1; rowIdx < data.length; rowIdx++) {
        const row = data[rowIdx];
        if (!row || row.length === 0) continue;
        
        const rowData = { _row_num: rowIdx + 1, _sheet: sheetName };
        
        headers.forEach(header => {
            const value = row[header.col];
            rowData[header.name] = value !== undefined ? String(value).trim() : '';
        });
        
        // 只保留有流程名称的行
        if (rowData['流程名称'] && rowData['流程名称'].trim()) {
            apiData.push(rowData);
        }
    }
    
    console.log(`已加载 ${apiData.length} 条接口记录\n`);
    return apiData;
}

/**
 * 根据开发分类推断流向
 */
function inferFlowDirection(devClass, specDesc) {
    if (!devClass) return '';
        
    // 提取开发分类代码（数字部分）
    const match = String(devClass).match(/(\d{3})/);
    if (match) {
        const classCode = match[1];
        let flow = FLOW_DIRECTION_RULES[classCode] || '';
            
        // 对于202基本资料作业，根据规格说明进一步确认流向
        if (classCode === '202' && specDesc) {
            const desc = String(specDesc).toLowerCase();
            // 如果描述中包含"查询"、"封装成api"等，通常是 T100提供数据给外部系统
            if (desc.includes('查询') || desc.includes('封装') || desc.includes('ipaas') || desc.includes('api')) {
                flow = 'T100→XX';  // T100提供数据查询接口
            }
        }
            
        return flow;
    }
    return '';
}

/**
 * 标准化流向格式 - 只判断方向，忽略目标系统（OA/国家平台/其他）
 */
function normalizeFlowDirection(flowDir) {
    if (!flowDir) return '';
    
    let flow = String(flowDir).toUpperCase()
        .replace(/\s/g, '')
        .replace(/--/g, '→')
        .replace(/->/g, '→');
    
    // 只判断方向，不区分目标系统
    // T100/ERP 开头 → 向外推送
    if (flow.startsWith('ERP→') || flow.startsWith('T100→')) {
        return 'T100→XX';
    }
    // OA/外部系统 开头 → 向T100推送
    if (flow.startsWith('OA→') || flow.startsWith('XX→')) {
        return 'XX→T100';
    }
    
    // 兜底判断：包含→的位置
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
    
    // 提取关键信息
    const devClass = requirement['开发分类'] || requirement['开发分类(*)'] || '';
    const progCode = String(requirement['程序代号'] || requirement['程序代号(*)'] || '').toLowerCase().trim();
    const reqName = requirement['程序名称'] || requirement['程序名称(*)'] || '';
    const specDesc = requirement['规格说明'] || requirement['规格说明(*)'] || '';
    
    if (!progCode) {
        return matches;
    }
    
    // 推断流向
    const inferredFlow = inferFlowDirection(devClass, specDesc);
    requirement._inferredFlow = inferredFlow;
    
    console.log(`  程序代号: ${progCode}, 开发分类: ${devClass}, 推断流向: ${inferredFlow}`);
    
    for (const apiRow of apiData) {
        const flowName = apiRow['流程名称'] || '';
        const apiDirection = apiRow['接口方向'] || '';
        
        if (!flowName) continue;
        
        let score = 0;
        const reasons = [];
        
        // 规则1: 程序代号精确匹配（权重 60%）
        const apiProgCode = extractProgCode(flowName);
        if (apiProgCode === progCode) {
            score += 60;
            reasons.push(`程序代号精确匹配: ${progCode}`);
        } else if (flowName.toLowerCase().includes(progCode)) {
            score += 40;
            reasons.push(`程序代号包含匹配: ${progCode}`);
        }
        
        // 规则2: 流向匹配（权重 30%）- 只判断方向，忽略目标系统
        if (inferredFlow && score > 0) {
            const normalizedApiDir = normalizeFlowDirection(apiDirection);
            
            if (inferredFlow === normalizedApiDir) {
                // 流向一致（都是 T100→XX 或都是 XX→T100）
                score += 30;
                reasons.push(`流向匹配: ${inferredFlow} ↔ ${apiDirection}`);
            } else if (normalizedApiDir) {
                // 流向相反
                score -= 20;
                reasons.push(`流向相反: ${inferredFlow} vs ${apiDirection}`);
            }
        }
        
        // 规则3: 需求名称辅助匹配（权重 10%）
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
        
        // 规则4: 规格说明中的程序代号提及（权重 5%）
        if (progCode && String(specDesc).toLowerCase().includes(progCode) && score > 0) {
            score += 5;
            reasons.push('规格说明中提及程序代号');
        }
        
        // 记录匹配结果
        if (score >= 40) {  // 阈值
            matches.push({
                api: apiRow,
                score: Math.min(score, 100),
                reason: reasons.join(' | '),
                matchType: '流程名称匹配'
            });
        }
    }
    
    // 按匹配度排序
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
    
    // 读取表头
    const headers = data[0].map((h, i) => ({
        col: i,
        name: h ? String(h).trim() : `Column${i + 1}`
    }));
    
    // 读取数据行
    const requirements = [];
    for (let rowIdx = 1; rowIdx < data.length; rowIdx++) {
        const row = data[rowIdx];
        const reqData = { _rowNum: rowIdx + 1 };
        
        headers.forEach(header => {
            const value = row[header.col];
            reqData[header.name] = value !== undefined ? String(value).trim() : '';
        });
        
        // 只保留有程序代号的行
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
 * 生成输出文件
 */
/**
 * 从API对照清单中复制指定工作表的内容（保留格式）
 */
function copyMatchedSheet(flowName, apiWorkbook) {
    if (!flowName || !apiWorkbook) return null;
    
    // 查找匹配的工作表（流程名称对应的工作表）
    const sheetName = apiWorkbook.SheetNames.find(name => {
        // 提取工作表中的程序代号
        const sheetProgCode = extractProgCode(name);
        const flowProgCode = extractProgCode(flowName);
        return sheetProgCode && sheetProgCode === flowProgCode;
    });
    
    if (!sheetName) {
        console.log(`  警告: 未找到对应工作表: ${flowName}`);
        return null;
    }
    
    // 获取原始工作表（保留所有格式信息）
    const sheet = apiWorkbook.Sheets[sheetName];
    
    // 复制工作表对象（深拷贝，保留格式）
    const clonedSheet = JSON.parse(JSON.stringify(sheet));
    
    console.log(`  已复制工作表: ${sheetName} (保留格式)`);
    return { sheetName, sheet: clonedSheet };
}

function generateOutput(requirements, results, originalFile) {
    console.log('\n生成匹配结果文件...');
    
    // 读取原始表头
    const wbIn = XLSX.readFile(originalFile);
    const wsIn = wbIn.Sheets[wbIn.SheetNames[0]];
    const originalData = XLSX.utils.sheet_to_json(wsIn, { header: 1 });
    const originalHeaders = originalData[0];
    
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
    
    // 构建输出数据
    const outputData = [...originalData];
    
    // 添加表头行（如果不存在匹配结果列）
    const firstRow = outputData[0];
    let baseCol = firstRow.length;
    matchHeaders.forEach(h => {
        if (!firstRow.includes(h)) {
            firstRow.push(h);
        }
    });
    
    // 填充数据行
    for (let i = 0; i < results.length; i++) {
        const result = results[i];
        const rowIdx = result.requirement._rowNum;
        
        // 确保行存在
        while (outputData.length <= rowIdx) {
            outputData.push([]);
        }
        
        const row = outputData[rowIdx];
        
        // 确保行有足够列
        while (row.length < firstRow.length) {
            row.push('');
        }
        
        if (result.api && result.score >= 40) {
            row[baseCol] = '✅ 已匹配';
            row[baseCol + 1] = result.api['流程名称'] || '';
            row[baseCol + 2] = result.api['接口方向'] || '';
            row[baseCol + 3] = result.inferredFlow || '';
            row[baseCol + 4] = `${result.score}%`;
            row[baseCol + 5] = result.reason;
            row[baseCol + 6] = `接口清单 行${result.api._row_num}`;
            row[baseCol + 7] = ONLINE_DOC_URL;
        } else {
            row[baseCol] = '❌ 未匹配';
            row[baseCol + 1] = '';
            row[baseCol + 2] = '';
            row[baseCol + 3] = result.inferredFlow || '';
            row[baseCol + 4] = '0%';
            row[baseCol + 5] = result.reason || '无匹配接口';
            row[baseCol + 6] = '';
            row[baseCol + 7] = ONLINE_DOC_URL;
        }
    }
    
    // 创建新的工作簿
    const wbOut = XLSX.utils.book_new();
    const wsOut = XLSX.utils.aoa_to_sheet(outputData);
    
    // 设置列宽
    const colWidths = {};
    for (let i = 0; i < firstRow.length; i++) {
        colWidths[XLSX.utils.encode_col(i)] = { wch: 25 };
    }
    wsOut['!cols'] = Object.values(colWidths);
    
    XLSX.utils.book_append_sheet(wbOut, wsOut, '需求匹配结果');
    
    // 读取API对照清单工作簿（用于复制匹配的工作表，保留样式）
    const apiWb = XLSX.readFile(API_MAPPING_FILE, { cellStyles: true, cellNF: true, cellDates: true });
    
    // 复制匹配的工作表到输出文件
    console.log('\n复制匹配的工作表...');
    let copiedCount = 0;
    const copiedSheets = new Set(); // 避免重复复制
    
    for (const result of results) {
        if (result.api && result.score >= 40) {
            const flowName = result.api['流程名称'];
            if (!flowName || copiedSheets.has(flowName)) continue;
            
            const sheetData = copyMatchedSheet(flowName, apiWb);
            if (sheetData) {
                // 直接使用克隆的工作表（保留所有格式）
                const ws = sheetData.sheet;
                
                // 工作表名称限制31字符
                let sheetName = sheetData.sheetName;
                if (sheetName.length > 31) {
                    sheetName = sheetName.substring(0, 28) + '...';
                }
                
                // 确保名称唯一
                let finalName = sheetName;
                let suffix = 1;
                while (wbOut.SheetNames.includes(finalName)) {
                    finalName = sheetName.substring(0, 29) + '_' + suffix;
                    suffix++;
                }
                
                XLSX.utils.book_append_sheet(wbOut, ws, finalName);
                copiedSheets.add(flowName);
                copiedCount++;
            }
        }
    }
    
    console.log(`共复制 ${copiedCount} 个工作表\n`);
    
    // 确定输出目录：集成需求字段清单产出
    const OUTPUT_DIR = path.join(WORKSPACE_DIR, '集成需求字段清单产出');
    if (!fs.existsSync(OUTPUT_DIR)) {
        fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    }
    
    // 生成时间戳（格式：yyyy-mm-dd_HH-MM-SS）
    const timestamp = formatTimestamp(new Date());
    
    // 保存文件（统一命名格式：原文件名_API匹配结果_时间戳.xlsx）
    const outputFile = path.join(
        OUTPUT_DIR,
        `${path.basename(originalFile, path.extname(originalFile))}_API匹配结果_${timestamp}.xlsx`
    );
    
    XLSX.writeFile(wbOut, outputFile);
    
    const stats = fs.statSync(outputFile);
    console.log(`匹配结果已保存: ${path.basename(outputFile)}`);
    console.log(`文件大小: ${(stats.size / 1024).toFixed(2)} KB\n`);
    
    return outputFile;
}

/**
 * 主函数
 */
function main() {
    console.log('='.repeat(80));
    console.log('     API Document Retrieval V4 - 能率需求清单智能匹配');
    console.log('     优化: 开发分类自动推断流向 + 程序代号+流向匹配流程名称');
    console.log('='.repeat(80));
    console.log();
    
    const args = process.argv.slice(2);
    if (args.length < 1) {
        console.log('用法: node match_nenglv_v4.js <需求清单xlsx文件>');
        console.log('\n支持格式: .xlsx');
        console.log('\n示例:');
        console.log('  node match_nenglv_v4.js 需求评估整理/能率需求清单-bpm集成.xlsx');
        console.log('  node match_nenglv_v4.js 需求评估整理/能率需求清单-国家平台集成.xlsx');
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
    const apiData = loadApiMapping();
    
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
    
    // 检查匹配率，如果低于阈值提示更新在线文档
    const matchRate = (matchedCount / requirements.length) * 100;
    if (matchRate < 50) {
        console.log('⚠️ 警告: 匹配率低于 50%，建议更新本地 API 对照清单');
        console.log(`\n当前本地文件: ${API_MAPPING_FILE}`);
        console.log(`在线文档地址: ${ONLINE_DOC_URL}`);
        console.log('\n📥 请手动下载最新版本:');
        console.log('  1. 访问上述在线文档地址');
        console.log('  2. 点击右上角「文件」→「导出为」→「Excel(.xlsx)」');
        console.log('  3. 保存到: 需求对照工具表/【2026】T100&OA类系统集成清单字段对照表.xlsx');
        console.log('\n尝试自动打开浏览器...\n');
        
        // 尝试自动打开浏览器
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
    const outputFile = generateOutput(requirements, results, filePath);
    
    console.log('='.repeat(80));
    console.log('智能匹配完成！');
    console.log('='.repeat(80));
    console.log(`\n输出文件: ${outputFile}`);
    console.log(`在线文档: ${ONLINE_DOC_URL}`);
    console.log(`\n请打开输出文件查看匹配结果！`);
}

// 运行主函数
main();
