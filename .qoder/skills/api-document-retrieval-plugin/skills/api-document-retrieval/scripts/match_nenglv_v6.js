#!/usr/bin/env node
// -*- coding: utf-8 -*-
/**
 * API Document Retrieval V6 - 统一模板智能匹配（Node.js + ExcelJS版本）
 * 基于AI输入模板和输出模板的标准格式进行匹配
 * 
 * 输入模板结构：
 * - 工作表1: 导入模板（需求清单主表）
 * - 工作表2-5: 下拉选项配置（clientCustomType/difficultyLevel/developmentClass/integrateProduct）
 * 
 * 输出模板结构：
 * - 工作表1: 输入物-TB导入EXCEL摸板（保留原需求数据+匹配结果列）
 * - 工作表2: 集成配置（从AI输出物摸板复制）
 * - 工作表3: 开发计划（从AI输出物摸板复制）
 * - 工作表4: 问题跟进管制-TB导入联动（从AI输出物摸板复制）
 * - 工作表5-N: 匹配到的字段对照详情页签（从【2026】字段对照---最小颗粒度.xlsx复制）
 * 
 * 用法: node match_nenglv_v6.js <需求清单xlsx文件>
 */

const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// 配置
const WORKSPACE_DIR = process.cwd();
const API_MAPPING_FILE = path.join(WORKSPACE_DIR, '需求对照工具表', '【2026】字段对照---最小颗粒度.xlsx');
const INTERFACE_LIST_FILE = path.join(WORKSPACE_DIR, '需求对照工具表', '【2026】接口清单---汇总.xlsx');
const OUTPUT_TEMPLATE_FILE = path.join(WORKSPACE_DIR, '集成需求字段清单产出', 'AI输出物摸板.xlsx');
const ONLINE_DOC_URL = 'https://docs.qq.com/sheet/DWmF4RXZmdW5RWnVk?tab=nxp5uc';
const FIELD_MAPPING_DOC_URL = 'https://docs.qq.com/sheet/DWm9PT0xtYWdYRW5R?scene=a815f6149367bdcef5a2dc6d3cXfw1&tab=hhd00s';

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
// 规则：
//   - 服务端接口（702/703/704/705）：默认 XX→ERP
//   - 客户端接口（710/711）：默认 ERP→XX
//   - 集成产品编号进一步细化流向：CRM→ERP, OA→ERP, WMS→ERP 等
const FLOW_DIRECTION_RULES = {
    '702': 'XX→ERP',
    '703': 'XX→ERP',
    '704': 'XX→ERP',
    '705': 'XX→ERP',
    '710': 'ERP→XX',
    '711': 'ERP→XX',
    '202': 'ERP→XX',
};

// 集成产品编号 → 外部系统映射
// 根据集成产品编号推断具体的外部系统名称
// 支持格式：F28:CRM_纷享, I03:HR_钉钉, C57:WMS_科大
const INTEGRATE_PRODUCT_RULES = {
    'CRM': 'CRM',
    'OA': 'OA',
    'WMS': 'WMS',
    'MES': 'MES',
    'SRM': 'SRM',
    'PLM': 'PLM',
    'HR': 'OA',      // HR系统映射到OA
    'PIM': 'PIM',    // 产品信息管理系统
    'BI': 'BI',      // 商业智能系统
    'QMS': 'QMS',    // 质量管理系统
    'TMS': 'TMS',    // 运输管理系统
    'APS': 'APS',    // 高级计划排程系统
    'SCM': 'SCM',    // 供应链管理系统
    'EAM': 'EAM',    // 企业资产管理系统
    'FMS': 'FMS',    // 财务管理系统
    'BPM': 'BPM',    // 业务流程管理系统
    'ESB': 'ESB',    // 企业服务总线
    'MDM': 'MDM',    // 主数据管理系统
    'DMS': 'DMS',    // 文档管理系统
    'LIMS': 'LIMS',  // 实验室信息管理系统
    'OMS': 'OMS',    // 订单管理系统
    'CMS': 'CMS',    // 内容管理系统
    'POS': 'POS',    // 销售点系统
    'EC': 'EC',      // 电子商务系统
    'APP': 'APP',    // 移动应用
    'WECHAT': '微信', // 微信集成
    'DINGTALK': '钉钉', // 钉钉集成
    'FEISHU': '飞书',   // 飞书集成
};

/**
 * 从集成产品编号中提取外部系统
 * 支持格式：F28:CRM_纷享, I03:HR_钉钉, C57:WMS_科大
 */
function extractExternalSystem(integrateProduct) {
    if (!integrateProduct) return 'XX';
    const productUpper = String(integrateProduct).toUpperCase().trim();
    
    // 如果直接匹配已知系统，返回
    if (INTEGRATE_PRODUCT_RULES[productUpper]) {
        return INTEGRATE_PRODUCT_RULES[productUpper];
    }
    
    // 解析格式：F28:CRM_纷享, I03:HR_钉钉, C57:WMS_科大
    // 提取冒号后的系统代码（CRM, OA, WMS, MES, SRM, HR等）
    const match = productUpper.match(/:\s*([A-Z]+)/);
    if (match) {
        const systemCode = match[1];
        if (INTEGRATE_PRODUCT_RULES[systemCode]) {
            return INTEGRATE_PRODUCT_RULES[systemCode];
        }
        return systemCode;
    }
    
    // 尝试在字符串中查找任何已知的系统代码
    for (const [key, value] of Object.entries(INTEGRATE_PRODUCT_RULES)) {
        if (productUpper.includes(key)) {
            return value;
        }
    }
    
    return 'XX';
}

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
        
        if (!data || data.length === 0) {
            console.log(`  跳过空工作表: ${sheetName}`);
            return;
        }
        
        let headerRowIdx = -1;
        for (let i = 0; i < Math.min(10, data.length); i++) {
            const row = data[i];
            if (row && row.some(cell => cell === '序号' || cell === '流程名称')) {
                headerRowIdx = i;
                break;
            }
        }
        
        if (headerRowIdx === -1 || !data[headerRowIdx]) {
            console.log(`  跳过无表头工作表: ${sheetName}`);
            return;
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
        console.error(`\n建议使用在线文档: ${FIELD_MAPPING_DOC_URL}`);
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
 * 根据开发分类和集成产品编号推断流向
 * 
 * 规则：
 * 1. 开发分类包含"服务端" → 外部系统→ERP（XX→ERP）
 * 2. 开发分类包含"客户端" → ERP→外部系统（ERP→XX）
 * 3. 结合集成产品编号确定具体外部系统：
 *    - CRM → CRM→ERP 或 ERP→CRM
 *    - OA → OA→ERP 或 ERP→OA
 *    - WMS → WMS→ERP 或 ERP→WMS
 *    - MES → MES→ERP 或 ERP→MES
 * 
 * @param {string} devClass - 开发分类（如"702:服务端_生单接口"）
 * @param {string} integrateProduct - 集成产品编号（如"CRM"）
 * @param {string} specDesc - 规格说明（备用）
 * @returns {string} 推断的流向
 */
function inferFlowDirection(devClass, integrateProduct, specDesc) {
    if (!devClass) return '';
    
    const devClassStr = String(devClass);
    const classCodeMatch = devClassStr.match(/(\d{3})/);
    const classCode = classCodeMatch ? classCodeMatch[1] : '';
    
    // 判断是否为服务端接口
    const isServerSide = devClassStr.includes('服务端') || 
                         ['702', '703', '704', '705'].includes(classCode);
    
    // 判断是否为客户端接口
    const isClientSide = devClassStr.includes('客户端') || 
                         ['710', '711', '202'].includes(classCode);
    
    // 从集成产品编号推断外部系统
    let externalSystem = extractExternalSystem(integrateProduct);
    
    // 根据服务端/客户端确定流向
    if (isServerSide) {
        // 服务端接口：外部系统 → ERP（XX→ERP）
        return `${externalSystem}→ERP`;
    } else if (isClientSide) {
        // 客户端接口：ERP → 外部系统（ERP→XX）
        return `ERP→${externalSystem}`;
    }
    
    // 备用：根据旧规则映射
    if (classCode) {
        let flow = FLOW_DIRECTION_RULES[classCode] || '';
        
        if (classCode === '202' && specDesc) {
            const desc = String(specDesc).toLowerCase();
            if (desc.includes('查询') || desc.includes('封装') || desc.includes('ipaas') || desc.includes('api')) {
                flow = 'ERP→XX';
            }
        }
        
        return flow;
    }
    
    return '';
}

/**
 * 标准化流向格式
 * 统一将 ERP 和 T100 视为等价（都是指ERP系统）
 */
function normalizeFlowDirection(flowDir) {
    if (!flowDir) return '';
    
    let flow = String(flowDir).toUpperCase()
        .replace(/\s/g, '')
        .replace(/--/g, '→')
        .replace(/->/g, '→')
        .replace(/T100/g, 'ERP');  // 统一将T100替换为ERP
    
    if (flow.startsWith('ERP→')) {
        return 'ERP→XX';
    }
    if (flow.startsWith('OA→') || flow.startsWith('CRM→') || flow.startsWith('WMS→') || flow.startsWith('MES→') || flow.startsWith('SRM→') || flow.startsWith('XX→')) {
        return 'XX→ERP';
    }
    
    if (flow.includes('→')) {
        const parts = flow.split('→');
        if (parts[0].includes('ERP')) {
            return 'ERP→XX';
        } else {
            return 'XX→ERP';
        }
    }
    
    return flow;
}

/**
 * 从文本中提取程序代号（支持客制前缀如 cwssp9001 -> axmt500）
 * 提取规则：
 * 1. 标准程序代号: [a-z]{2,4}\d{3}（如 axmt500, aooi130）
 * 2. 客制程序代号: c[a-z]{2,4}\d{3,4}（如 cwssp9001, caxmt500）
 * 3. 从客制代号中提取标准代号: 去掉前缀 c 或 cwssp 等
 */
function extractProgCode(text) {
    if (!text) return '';
    const textLower = String(text).toLowerCase();
    
    // 先尝试匹配标准程序代号 (如 axmt500, aooi130)
    const standardMatch = textLower.match(/([a-z]{2,4}\d{3})/);
    if (standardMatch) {
        return standardMatch[1];
    }
    
    return '';
}

/**
 * 从客制程序代号中提取标准程序代号
 * 规则:
 *   - 新增客制接口: cwssp* (如 cwssp9001) -> 从程序名称提取标准代号
 *   - 修改标准接口: wssp* (如 wssp339) -> 从程序名称提取标准代号，标记为OPENAPI
 *   - 标准接口: axmt500, aooi130 等 -> 直接使用
 *   - 带后缀: axmt500_send -> 去掉后缀
 * 
 * 返回对象: { standardCode, isOpenApi }
 *   - standardCode: 标准程序代号
 *   - isOpenApi: 是否为OPENAPI接口 (wssp* 前缀)
 */
function extractStandardProgCode(customProgCode, progName) {
    if (!customProgCode) return { standardCode: '', isOpenApi: false };
    const codeLower = String(customProgCode).toLowerCase().trim();
    
    // 判断是否为 wssp* 前缀（修改标准接口，使用OPENAPI）
    const isOpenApi = codeLower.startsWith('wssp');
    
    // 判断是否为 cwssp* 前缀（新增客制接口）
    const isCustom = codeLower.startsWith('cwssp');
    
    // 去掉下划线后缀: axmt500_send -> axmt500
    let cleanedCode = codeLower.replace(/_.*$/, '');
    
    // 如果清洗后符合标准格式，直接返回
    const standardMatch = cleanedCode.match(/([a-z]{2,4}\d{3})/);
    if (standardMatch && standardMatch[0] === cleanedCode) {
        return { standardCode: cleanedCode, isOpenApi: false };
    }
    
    // 如果是 cwssp* 或 wssp* 前缀，从程序名称提取标准代号
    if ((isCustom || isOpenApi) && progName) {
        const nameMatch = String(progName).toLowerCase().match(/([a-z]{2,4}\d{3})/);
        if (nameMatch) {
            return { standardCode: nameMatch[1], isOpenApi: isOpenApi };
        }
    }
    
    return { standardCode: cleanedCode, isOpenApi: isOpenApi };
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
 * 提取操作类型关键词（查询、推送、创建、修改等）
 * 用于判断需求名称和接口流程名称的操作类型是否一致
 */
function extractOperationType(text) {
    if (!text) return '';
    const textLower = String(text).toLowerCase();
    
    // 操作类型优先级：查询 > 推送 > 创建 > 修改 > 删除 > 审核 > 过账
    const operationTypes = [
        { type: '查询', keywords: ['查询', '获取', '读取', 'read', 'query', 'get'] },
        { type: '推送', keywords: ['推送', '发送', '抛转', 'publish', 'push', 'send'] },
        { type: '创建', keywords: ['创建', '新增', '生成', 'create', 'add', 'insert'] },
        { type: '修改', keywords: ['修改', '更新', '变更', 'update', 'modify', 'change'] },
        { type: '删除', keywords: ['删除', '作废', '移除', 'delete', 'remove', 'cancel'] },
        { type: '审核', keywords: ['审核', '审批', '核准', 'approve', 'review', 'check'] },
        { type: '过账', keywords: ['过账', 'post', '过帐'] }
    ];
    
    for (const op of operationTypes) {
        for (const kw of op.keywords) {
            if (textLower.includes(kw)) {
                return op.type;
            }
        }
    }
    
    return '';
}

/**
 * 从字段对照页签中提取方向信息
 * 页签第一行通常包含格式如：流程名称：aimm200料件主档维护作业申请新增流程(OA-->ERP)
 * @param {ExcelJS.Worksheet} sheet - ExcelJS工作表对象
 * @returns {string} 提取的方向，如 'OA→ERP'，未找到则返回空字符串
 */
function extractSheetDirection(sheet) {
    if (!sheet) return '';
    
    // 读取第一行第一列的内容
    const firstRow = sheet.getRow(1);
    const firstCell = firstRow.getCell(1);
    const cellValue = firstCell.value ? String(firstCell.value) : '';
    
    // 匹配括号中的方向信息，如 (OA-->ERP), (SRM→ERP), (ERP→WMS) 等
    const directionMatch = cellValue.match(/\(([A-Za-z]+)\s*[-–—→>]+\s*([A-Za-z]+)\)/);
    if (directionMatch) {
        const from = directionMatch[1].toUpperCase();
        const to = directionMatch[2].toUpperCase();
        // 统一将T100替换为ERP
        const normalizedFrom = from === 'T100' ? 'ERP' : from;
        const normalizedTo = to === 'T100' ? 'ERP' : to;
        return `${normalizedFrom}→${normalizedTo}`;
    }
    
    // 也尝试匹配不包含括号的格式
    const looseMatch = cellValue.match(/([A-Za-z]+)\s*[-–—→>]+\s*([A-Za-z]+)/);
    if (looseMatch) {
        const from = looseMatch[1].toUpperCase();
        const to = looseMatch[2].toUpperCase();
        const normalizedFrom = from === 'T100' ? 'ERP' : from;
        const normalizedTo = to === 'T100' ? 'ERP' : to;
        return `${normalizedFrom}→${normalizedTo}`;
    }
    
    return '';
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
    const integrateProduct = requirement['集成产品编号'] || requirement['集成产品编号(*)'] || '';
    
    if (!progCode) {
        return matches;
    }
    
    // 201 开发分类（建档/视图）不是接口开发，跳过接口匹配
    const devClassStr = String(devClass);
    const classCodeMatch = devClassStr.match(/(\d{3})/);
    const classCode = classCodeMatch ? classCodeMatch[1] : '';
    if (classCode === '201' || devClassStr.includes('建档')) {
        console.log(`  [跳过] 开发分类为201(建档/视图)，非接口开发，跳过匹配`);
        requirement._inferredFlow = '';
        requirement._isOpenApi = false;
        return matches;
    }
    
    // 从客制程序代号中提取标准程序代号
    const { standardCode: standardProgCode, isOpenApi } = extractStandardProgCode(progCode, reqName);
    requirement._isOpenApi = isOpenApi;
    
    const inferredFlow = inferFlowDirection(devClass, integrateProduct, specDesc);
    requirement._inferredFlow = inferredFlow;
    
    console.log(`  程序代号: ${progCode}, 标准代号: ${standardProgCode}, OPENAPI: ${isOpenApi}, 开发分类: ${devClass}, 推断流向: ${inferredFlow}`);
    
    for (const apiRow of apiData) {
        const flowName = apiRow['流程名称'] || '';
        const apiDirection = apiRow['接口方向'] || '';
        
        if (!flowName) continue;
        
        let score = 0;
        const reasons = [];
        
        const apiProgCode = extractProgCode(flowName);
        
        // 匹配逻辑：优先使用标准程序代号进行匹配
        if (apiProgCode === standardProgCode) {
            score += 60;
            reasons.push(`程序代号精确匹配: ${standardProgCode}`);
        } else if (flowName.toLowerCase().includes(standardProgCode)) {
            score += 40;
            reasons.push(`程序代号包含匹配: ${standardProgCode}`);
        } else if (apiProgCode === progCode) {
            // 原始程序代号精确匹配（备用）
            score += 60;
            reasons.push(`程序代号精确匹配: ${progCode}`);
        } else if (flowName.toLowerCase().includes(progCode)) {
            // 原始程序代号包含匹配（备用）
            score += 40;
            reasons.push(`程序代号包含匹配: ${progCode}`);
        }
        
        // 流向匹配逻辑：根据推断流向和接口方向判断是否匹配
        if (inferredFlow && score > 0) {
            // 标准化接口方向（用于比较流向是否一致）
            const normalizedApiDir = normalizeFlowDirection(apiDirection);
            
            // 从推断流向提取外部系统
            const inferredExternal = inferredFlow.split('→')[0];
            
            // 从原始接口方向提取外部系统（不经过归一化，保留原始系统信息）
            // 处理格式：ERP-->OA, WMS-->ERP, OA→ERP 等
            let apiDirExternal = '';
            const apiDirUpper = String(apiDirection).toUpperCase().replace(/\s/g, '');
            
            // 提取箭头后的系统代码（对于 ERP→XX 类型，外部系统是箭头后的系统）
            // 提取箭头前的系统代码（对于 XX→ERP 类型，外部系统是箭头前的系统）
            const arrowMatch = apiDirUpper.match(/^([A-Z]+)\s*[-–—→>]+\s*([A-Z]+)/);
            if (arrowMatch) {
                const fromSystem = arrowMatch[1];
                const toSystem = arrowMatch[2];
                
                // 判断哪个是外部系统：
                // 如果 ERP 在箭头前（ERP→XX），外部系统是箭头后的系统
                // 如果 ERP 在箭头后（XX→ERP），外部系统是箭头前的系统
                if (fromSystem === 'ERP' || fromSystem === 'T100') {
                    apiDirExternal = toSystem;
                } else {
                    apiDirExternal = fromSystem;
                }
            }
            
            // 检查外部系统是否一致（PLM→ERP 不能匹配 OA→ERP）
            const externalSystemMatch = inferredExternal === apiDirExternal || 
                                       apiDirExternal === '' ||
                                       apiDirExternal === 'XX';
            
            // 调试输出
            if (flowName.toLowerCase().includes('axmt540')) {
                console.log(`    [DEBUG] inferredFlow=${inferredFlow}, normalizedApiDir=${normalizedApiDir}`);
                console.log(`    [DEBUG] inferredExternal=${inferredExternal}, apiDirExternal=${apiDirExternal}`);
                console.log(`    [DEBUG] externalSystemMatch=${externalSystemMatch}`);
            }
            
            // 将推断流向也归一化，用于比较
            const normalizedInferredFlow = normalizeFlowDirection(inferredFlow);
            
            if (normalizedInferredFlow === normalizedApiDir) {
                // 归一化后的流向匹配，但还需要检查外部系统是否一致
                if (!externalSystemMatch) {
                    // 外部系统不一致（如 ERP→WMS vs ERP→OA），直接跳过
                    console.log(`    跳过: 外部系统不匹配 需求[${inferredFlow}] vs 接口[${apiDirection}]`);
                    continue;
                }
                // 流向完全匹配，加分
                score += 30;
                reasons.push(`流向匹配: ${inferredFlow} ↔ ${apiDirection}`);
            } else if (normalizedApiDir) {
                // 流向不匹配，检查外部系统是否一致
                if (!externalSystemMatch) {
                    // 外部系统不一致（如 PLM→ERP vs OA→ERP），直接跳过
                    console.log(`    跳过: 外部系统不匹配 需求[${inferredFlow}] vs 接口[${apiDirection}]`);
                    continue;
                }
                // 外部系统一致但流向相反（如 SRM→ERP vs ERP→SRM），大幅减分
                score -= 40;
                reasons.push(`流向相反: ${inferredFlow} vs ${apiDirection}`);
            }
        }
        
        // 操作类型匹配：检查需求名称和接口流程名称的操作类型是否一致
        // 如果操作类型不一致（如查询 vs 推送），直接过滤掉，不加入匹配结果
        if (score > 0 && reqName) {
            const reqOpType = extractOperationType(reqName);
            const flowOpType = extractOperationType(flowName);
            
            if (reqOpType && flowOpType) {
                if (reqOpType === flowOpType) {
                    // 操作类型一致，加分
                    score += 15;
                    reasons.push(`操作类型匹配: ${reqOpType}`);
                } else {
                    // 操作类型不一致，直接跳过此接口（不加入匹配结果）
                    console.log(`    跳过: 操作类型不一致 需求[${reqOpType}] vs 接口[${flowOpType}]`);
                    continue;
                }
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
 * 读取需求清单文件（适配新模板：第一个工作表为"导入模板"）
 */
function readRequirementFile(filePath) {
    console.log(`读取需求清单: ${path.basename(filePath)}`);
    
    const wb = XLSX.readFile(filePath);
    
    // 新模板：第一个工作表是"导入模板"，包含需求数据
    // 也可能直接是第一个工作表（兼容旧格式）
    let targetSheetName = wb.SheetNames[0];
    
    // 查找包含"导入"或"模板"的工作表
    const importSheet = wb.SheetNames.find(name => 
        name.includes('导入') || name.includes('模板') || name.includes('输入')
    );
    if (importSheet) {
        targetSheetName = importSheet;
    }
    
    const ws = wb.Sheets[targetSheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    
    if (!data || data.length === 0) {
        console.error('错误: 需求清单文件为空');
        process.exit(1);
    }
    
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
    
    console.log(`工作表: ${targetSheetName}`);
    console.log(`提取到 ${requirements.length} 条需求记录\n`);
    return { requirements, headers, wb };
}

/**
 * 使用 ExcelJS 复制工作表，完整保留所有格式
 * 如果源工作表为空（无有效数据行），则返回 null 不复制
 */
async function copyWorksheetWithFormat(sourceWorkbook, sourceSheetName, targetWorkbook, targetSheetName) {
    const sourceWs = sourceWorkbook.getWorksheet(sourceSheetName);
    if (!sourceWs) {
        console.log(`  警告: 未找到源工作表: ${sourceSheetName}`);
        return null;
    }
    
    // 检查工作表是否为空（无有效数据行）
    let hasData = false;
    let dataRowCount = 0;
    sourceWs.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        // 跳过表头行（通常第1行），检查是否有数据行
        if (rowNumber > 1) {
            dataRowCount++;
            hasData = true;
        }
    });
    
    if (!hasData || dataRowCount === 0) {
        console.log(`  提示: 工作表为空，跳过复制: ${sourceSheetName} (无有效数据行)`);
        return null;
    }
    
    // 创建新工作表
    const targetWs = targetWorkbook.addWorksheet(targetSheetName);
    
    // 复制列属性（列宽、隐藏状态等）
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
    
    // 复制行属性（行高、隐藏状态等）和单元格 - 优化版本
    // 使用批量复制减少函数调用开销
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
        
        // 复制每个单元格 - 优化：减少重复属性访问
        // 注意：_cells 是 0-based 数组，colNumber 是 1-based
        const cells = sourceRow._cells;
        if (!cells) return;
        
        for (let colIndex = 0; colIndex < cells.length; colIndex++) {
            const sourceCell = cells[colIndex];
            if (!sourceCell) continue;
            
            const colNumber = colIndex + 1; // 转换回 1-based 列号
            const targetCell = targetRow.getCell(colNumber);
            
            // 复制值（先处理特殊类型）
            if (sourceCell.formula) {
                targetCell.value = { formula: sourceCell.formula, result: sourceCell.result };
            } else if (sourceCell.hyperlink) {
                targetCell.value = { text: sourceCell.text, hyperlink: sourceCell.hyperlink };
            } else {
                targetCell.value = sourceCell.value;
            }
            
            // 复制样式（深拷贝，避免样式对象引用共享导致修改一个单元格影响其他单元格）
            if (sourceCell.style) {
                try {
                    targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
                } catch (e) {
                    // 如果深拷贝失败，回退到引用赋值（可能仍有共享问题）
                    targetCell.style = sourceCell.style;
                }
            }
            
            // 复制数据验证
            if (sourceCell.dataValidation) {
                targetCell.dataValidation = sourceCell.dataValidation;
            }
        }
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
 * 生成输出文件（适配新输出模板格式）
 */
async function generateOutput(requirements, results, originalFile, apiMappingData) {
    console.log('\n生成匹配结果文件...');
    
    // 使用 ExcelJS 读取原始文件以保留格式
    const originalWorkbook = new ExcelJS.Workbook();
    await originalWorkbook.xlsx.readFile(originalFile);
    
    // 找到导入模板工作表（第一个包含"导入"或"模板"的工作表）
    let importTemplateWs = originalWorkbook.worksheets[0];
    for (const ws of originalWorkbook.worksheets) {
        if (ws.name.includes('导入') || ws.name.includes('模板') || ws.name.includes('输入')) {
            importTemplateWs = ws;
            break;
        }
    }
    
    // 获取原始表头（从第一行提取）
    const originalHeaders = [];
    const firstRow = importTemplateWs.getRow(1);
    firstRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        originalHeaders[colNumber - 1] = cell.value ? String(cell.value).trim() : '';
    });
    
    // 创建新的输出工作簿
    const outputWorkbook = new ExcelJS.Workbook();
    
    // 复制导入模板工作表到输出工作簿（保留所有格式）
    await copyWorksheetWithFormat(originalWorkbook, importTemplateWs.name, outputWorkbook, '输入物-TB导入EXCEL摸板');
    
    // 使用 ExcelJS 读取AI输出物摸板（用于复制集成配置/开发计划/问题跟进管制工作表）
    const outputTemplateWorkbook = new ExcelJS.Workbook();
    if (fs.existsSync(OUTPUT_TEMPLATE_FILE)) {
        await outputTemplateWorkbook.xlsx.readFile(OUTPUT_TEMPLATE_FILE);
    } else {
        console.log(`  警告: 未找到AI输出物摸板: ${OUTPUT_TEMPLATE_FILE}`);
    }
    
    // 复制AI输出物摸板中的工作表（集成配置、开发计划、问题跟进管制-TB导入联动）
    console.log('\n复制AI输出物摸板工作表...');
    const templateSheetsToCopy = ['集成配置', '开发计划', '问题跟进管制-TB导入联动'];
    for (const sheetName of templateSheetsToCopy) {
        const sourceWs = outputTemplateWorkbook.getWorksheet(sheetName);
        if (sourceWs) {
            await copyWorksheetWithFormat(outputTemplateWorkbook, sheetName, outputWorkbook, sheetName);
        } else {
            console.log(`  警告: AI输出物摸板中未找到工作表: ${sheetName}`);
        }
    }
    
    // 使用 ExcelJS 读取API字段对照清单（用于复制匹配的工作表，保留样式）
    const apiWorkbook = new ExcelJS.Workbook();
    await apiWorkbook.xlsx.readFile(API_MAPPING_FILE);
    
    // 复制匹配的字段对照工作表到输出文件
    console.log('\n复制匹配的字段对照工作表...');
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
            
            // 判断是否为OPENAPI接口 (wssp* 前缀或需求标记为OPENAPI)
            const isOpenApi = result.requirement._isOpenApi || false;
            
            // 查找匹配的字段对照工作表
            // 如果是OPENAPI接口，优先匹配包含"OPENAPI"的工作表
            let sourceSheetName = null;
            
            if (isOpenApi) {
                // OPENAPI接口：匹配包含程序代号且包含"OPENAPI"的工作表
                sourceSheetName = apiWorkbook.worksheets.find(ws => {
                    const wsNameLower = ws.name.toLowerCase();
                    return wsNameLower.includes(flowProgCode) && wsNameLower.includes('openapi');
                })?.name;
                
                if (sourceSheetName) {
                    console.log(`  OPENAPI匹配: ${flowName} -> ${sourceSheetName}`);
                }
            }
            
            // 如果不是OPENAPI或没找到OPENAPI匹配，则普通匹配
            // 根据操作类型严格匹配对应的工作表
            if (!sourceSheetName) {
                const reqOpType = extractOperationType(result.requirement['程序名称'] || result.requirement['程序名称(*)'] || '');
                const inferredFlow = result.requirement._inferredFlow || '';
                
                // 先获取所有包含程序代号的页签
                let matchingSheets = apiWorkbook.worksheets.filter(ws => {
                    const wsNameLower = ws.name.toLowerCase();
                    return wsNameLower.includes(flowProgCode);
                });
                
                // 方向校验：如果需求推断出方向，过滤掉方向不一致的页签
                if (inferredFlow && matchingSheets.length > 0) {
                    const directionMatchedSheets = [];
                    
                    for (const ws of matchingSheets) {
                        const sheetDirection = extractSheetDirection(ws);
                        
                        // 如果页签有方向信息，检查是否与需求方向一致
                        if (sheetDirection) {
                            // 提取需求和页签的外部系统
                            const inferredExternal = inferredFlow.split('→')[0];
                            const sheetExternal = sheetDirection.split('→')[0];
                            
                            // 检查流向是否一致（外部系统必须相同，或页签是通用格式XX）
                            const inferredBase = inferredFlow.replace(/^[A-Z]+→/, 'XX→');
                            const sheetBase = sheetDirection.replace(/^[A-Z]+→/, 'XX→');
                            
                            // 外部系统必须匹配（PLM→ERP 不能匹配 OA→ERP）
                            const externalSystemMatch = inferredExternal === sheetExternal || sheetExternal === 'XX';
                            const flowDirectionMatch = inferredBase === sheetBase;
                            
                            if (externalSystemMatch && flowDirectionMatch) {
                                directionMatchedSheets.push(ws);
                                console.log(`  方向匹配: ${ws.name} 方向=${sheetDirection} (需求方向=${inferredFlow})`);
                            } else {
                                console.log(`  方向不匹配，跳过: ${ws.name} 方向=${sheetDirection} (需求方向=${inferredFlow})`);
                            }
                        } else {
                            // 页签没有方向信息，保留（可能是旧格式）
                            directionMatchedSheets.push(ws);
                        }
                    }
                    
                    // 如果方向过滤后还有页签，使用过滤后的结果
                    if (directionMatchedSheets.length > 0) {
                        matchingSheets = directionMatchedSheets;
                    } else if (matchingSheets.length > 0) {
                        // 方向过滤后没有页签，说明所有页签方向都不匹配，标记为未匹配
                        console.log(`  未匹配: 所有字段对照页签方向与需求不一致 (需求方向: ${inferredFlow})`);
                        result.api = null;
                        result.score = 0;
                        result.reason = `字段对照页签方向不匹配: 需求方向${inferredFlow}`;
                        continue;
                    }
                }
                
                if (matchingSheets.length === 0) {
                    // 最小颗粒度表中没有对应程序代号的任何工作表，标记为未匹配
                    console.log(`  未匹配: 最小颗粒度表中无对应页签: ${flowName} (程序代号: ${flowProgCode})`);
                    result.api = null;
                    result.score = 0;
                    result.reason = `最小颗粒度表中无对应页签: ${flowProgCode}`;
                    continue;
                } else if (matchingSheets.length === 1) {
                    sourceSheetName = matchingSheets[0].name;
                } else if (matchingSheets.length > 1 && reqOpType) {
                    // 多个匹配时，根据操作类型严格选择最匹配的工作表
                    // 操作类型到关键词的映射（用于工作表名称匹配）
                    const opKeywords = {
                        '查询': ['查询', '读取'],
                        '推送': ['推送', '发送', '抛转'],
                        '创建': ['创建', '新增', '生成'],
                        '修改': ['修改', '更新', '变更'],
                        '删除': ['删除', '作废', '移除'],
                        '审核': ['审核', '审批'],
                        '过账': ['过账', 'post', '过账更新']
                    };
                    
                    const keywords = opKeywords[reqOpType] || [reqOpType];
                    
                    // 严格匹配：只选择包含操作类型关键词的工作表
                    let bestMatch = null;
                    let bestScore = -1;
                    
                    for (const ws of matchingSheets) {
                        const wsNameLower = ws.name.toLowerCase();
                        let score = 0;
                        for (const kw of keywords) {
                            if (wsNameLower.includes(kw.toLowerCase())) {
                                score += 10;
                                // 完全匹配关键词加分
                                if (kw === '过账更新' && wsNameLower.includes('过账更新')) {
                                    score += 5;
                                }
                            }
                        }
                        if (score > bestScore) {
                            bestScore = score;
                            bestMatch = ws;
                        }
                    }
                    
                    // 如果找到匹配的工作表，使用它；否则标记为未匹配
                    if (bestMatch && bestScore > 0) {
                        sourceSheetName = bestMatch.name;
                        console.log(`  多工作表匹配: 根据操作类型[${reqOpType}]选择 -> ${sourceSheetName}`);
                    } else {
                        console.log(`  未匹配: 找不到操作类型[${reqOpType}]对应的字段对照工作表`);
                        result.api = null;
                        result.score = 0;
                        result.reason = `找不到操作类型[${reqOpType}]对应的字段对照工作表`;
                        continue;
                    }
                } else if (matchingSheets.length > 0) {
                    sourceSheetName = matchingSheets[0].name;
                }
            }
            
            if (!sourceSheetName) {
                console.log(`  警告: 未找到对应字段对照工作表: ${flowName} (程序代号: ${flowProgCode}, OPENAPI: ${isOpenApi})`);
                continue;
            }
            
            // 检查源字段对照工作表是否为空（无有效数据行）
            const sourceWs = apiWorkbook.getWorksheet(sourceSheetName);
            if (sourceWs) {
                let dataRowCount = 0;
                sourceWs.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                    if (rowNumber > 1) dataRowCount++;
                });
                
                if (dataRowCount === 0) {
                    // 字段对照页签为空，标记为未匹配
                    console.log(`  未匹配: ${flowName} 对应的字段对照工作表为空 (${sourceSheetName})`);
                    result.api = null;
                    result.score = 0;
                    result.reason = '字段对照工作表为空';
                    continue;
                }
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
            const copiedWs = await copyWorksheetWithFormat(apiWorkbook, sourceSheetName, outputWorkbook, finalName);
            
            if (!copiedWs) {
                // 工作表为空，标记为未匹配
                console.log(`  未匹配: ${flowName} 对应的字段对照工作表为空`);
                result.api = null;
                result.score = 0;
                result.reason = '字段对照工作表为空';
                continue;
            }
            
            // 将接口清单中的功能描述写入到产出文件对应页签的A2单元格
            if (result.api && result.api['功能描述']) {
                const funcDesc = result.api['功能描述'];
                const a2Cell = copiedWs.getRow(2).getCell(1);
                a2Cell.value = funcDesc;
                console.log(`  已写入功能描述到A2: ${funcDesc.substring(0, 50)}${funcDesc.length > 50 ? '...' : ''}`);
            }
            
            copiedSheets.add(flowName);
            copiedCount++;
        }
    }
    
    console.log(`共复制 ${copiedCount} 个字段对照工作表\n`);
    
    // 设置背景色：在字段对照页签复制完成后，根据最终结果设置背景色
    console.log('设置匹配结果背景色...');
    
    // 获取输出工作簿中的输入物工作表
    const outputImportWs = outputWorkbook.getWorksheet('输入物-TB导入EXCEL摸板');
    if (!outputImportWs) {
        console.log('  警告: 输出工作簿中未找到输入物工作表');
        console.log('  可用工作表: ' + outputWorkbook.worksheets.map(ws => ws.name).join(', '));
    } else {
        console.log('  outputImportWs 找到: ' + outputImportWs.name);
        console.log('  importTemplateWs === outputImportWs: ' + (importTemplateWs === outputImportWs));
    }
    
    for (let i = 0; i < results.length; i++) {
        const result = results[i];
        const rowIdx = result.requirement._rowNum;
        
        // 在输出工作簿的工作表上设置背景色
        const outputRow = outputImportWs ? outputImportWs.getRow(rowIdx) : null;
        const originalRow = importTemplateWs.getRow(rowIdx);
        
        // 匹配度阈值：50%（低于50%视为未匹配）
        const isMatched = result.api && result.score >= 50;
        
        if (isMatched) {
            // 匹配成功的行用浅黄色填充背景
            const lightYellowFill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFFCC' }  // 浅黄色
            };
            
            // 给匹配成功的数据行的所有列添加浅黄色背景
            if (rowIdx > 1 && outputRow) {
                for (let col = 1; col <= originalHeaders.length; col++) {
                    const cell = outputRow.getCell(col);
                    if (!cell.style) {
                        cell.style = {};
                    }
                    cell.style.fill = lightYellowFill;
                }
            }
        } else {
            // 未匹配的行清除所有列的填充颜色
            if (rowIdx > 1 && outputRow) {
                for (let col = 1; col <= originalHeaders.length; col++) {
                    const cell = outputRow.getCell(col);
                    if (cell.style && cell.style.fill) {
                        delete cell.style.fill;
                    }
                }
            }
        }
    }
    
    // 自适应列宽：根据内容自动调整每列的宽度
    console.log('自动调整列宽...');
    if (outputImportWs) {
        // 遍历所有列，根据内容计算最佳宽度
        for (let col = 1; col <= originalHeaders.length; col++) {
            const column = outputImportWs.getColumn(col);
            let maxLength = 0;
            
            // 遍历该列的所有行，计算最大内容长度
            outputImportWs.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                const cell = row.getCell(col);
                if (cell.value) {
                    const cellText = String(cell.value);
                    // 考虑中文字符（占2个宽度单位）和换行符
                    let cellWidth = 0;
                    for (const char of cellText) {
                        // 中文字符和全角字符占2个单位
                        if (/[\u4e00-\u9fa5\u3000-\u303f\uff00-\uffef]/.test(char)) {
                            cellWidth += 2;
                        } else {
                            cellWidth += 1;
                        }
                    }
                    // 考虑换行符，取最长行
                    const lines = cellText.split('\n');
                    for (const line of lines) {
                        let lineWidth = 0;
                        for (const char of line) {
                            if (/[\u4e00-\u9fa5\u3000-\u303f\uff00-\uffef]/.test(char)) {
                                lineWidth += 2;
                            } else {
                                lineWidth += 1;
                            }
                        }
                        if (lineWidth > maxLength) {
                            maxLength = lineWidth;
                        }
                    }
                }
            });
            
            // 设置列宽：最大内容长度 + 边距（2个字符），最小宽度10，最大宽度80
            if (maxLength > 0) {
                const newWidth = Math.min(Math.max(maxLength + 2, 10), 80);
                column.width = newWidth;
            }
        }
        console.log('  列宽调整完成');
    }
    
    // 确定输出目录
    const OUTPUT_DIR = path.join(WORKSPACE_DIR, '集成需求字段清单产出');
    if (!fs.existsSync(OUTPUT_DIR)) {
        fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    }
    
    // 生成输出文件名（固定名称，只保留最新版本）
    const outputFile = path.join(
        OUTPUT_DIR,
        `${path.basename(originalFile, path.extname(originalFile))}_API匹配结果.xlsx`
    );
    
    // 如果文件已存在，先删除旧文件
    if (fs.existsSync(outputFile)) {
        try {
            fs.unlinkSync(outputFile);
            console.log(`  已删除旧版本: ${path.basename(outputFile)}`);
        } catch (err) {
            if (err.code === 'EBUSY') {
                console.log(`  警告: 文件被占用，无法删除旧版本: ${path.basename(outputFile)}`);
                console.log(`  请关闭Excel中的该文件后重试，或使用不同的输出文件名`);
                // 生成带时间戳的备用文件名
                const timestamp = formatTimestamp(new Date());
                const altOutputFile = path.join(
                    OUTPUT_DIR,
                    `${path.basename(originalFile, path.extname(originalFile))}_API匹配结果_${timestamp}.xlsx`
                );
                console.log(`  使用备用文件名: ${path.basename(altOutputFile)}`);
                return await outputWorkbook.xlsx.writeFile(altOutputFile);
            } else {
                throw err;
            }
        }
    }
    
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
    console.log('     API Document Retrieval V6 - 统一模板智能匹配');
    console.log('     优化: 开发分类自动推断流向 + 程序代号+流向匹配流程名称');
    console.log('     格式保留: 使用 ExcelJS 完整保留边框/颜色/字体/合并单元格/列宽/行高');
    console.log('     输出格式: 适配AI输出模板（输入物+接口清单+集成配置+开发计划+问题跟进+字段对照）');
    console.log('='.repeat(80));
    console.log();
    
    const args = process.argv.slice(2);
    if (args.length < 1) {
        console.log('用法: node match_nenglv_v6.js <需求清单xlsx文件>');
        console.log('\n支持格式: .xlsx');
        console.log('\n示例:');
        console.log('  node match_nenglv_v6.js 需求评估整理/AI输入摸板-联动TB导入EXCEL摸板.xlsx');
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
        console.log('  3. 保存字段对照清单到: 需求对照工具表/【2026】字段对照---最小颗粒度.xlsx');
        console.log('  4. 保存接口清单汇总到: 需求对照工具表/【2026】接口清单---汇总.xlsx');
        console.log('');
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
