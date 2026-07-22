#!/usr/bin/env node
// -*- coding: utf-8 -*-
/**
 * API Document Retrieval - AI Assist Matching
 * 使用 Kimi API 进行语义检索匹配，作为本地规则匹配失败时的兜底方案
 *
 * 用法: node match_ai_assist.js <需求清单xlsx文件> [选项]
 *
 * 选项:
 *   --api-key <key>     指定 Kimi API Key
 *   --threshold <n>     本地匹配度阈值，低于此值才调用 AI（默认 40）
 *   --output <dir>      输出目录（默认 集成需求字段清单产出/）
 *   --all               对所有行都调用 AI 分析（不筛选失败行）
 *   --model <model>     指定 Kimi 模型（默认 moonshot-v1-32k）
 *   --max-retries <n>   API 调用失败最大重试次数（默认 3）
 *   --help              显示帮助信息
 *
 * 示例:
 *   node match_ai_assist.js 需求评估整理/AI输入摸板.xlsx
 *   node match_ai_assist.js 需求评估整理/AI输入摸板.xlsx --api-key sk-xxx --threshold 50
 *   node match_ai_assist.js 需求评估整理/AI输入摸板.xlsx --all
 */

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const https = require('https');

// 配置
const WORKSPACE_DIR = process.cwd();
const INTERFACE_LIST_FILE = path.join(WORKSPACE_DIR, '需求对照工具表', '【2026】接口清单---汇总.xlsx');
const DEFAULT_OUTPUT_DIR = path.join(WORKSPACE_DIR, '集成需求字段清单产出');
const KIMI_API_URL = 'https://api.moonshot.cn/v1/chat/completions';
const DEFAULT_MODEL = 'moonshot-v1-32k';
const DEFAULT_THRESHOLD = 40;
const DEFAULT_MAX_RETRIES = 3;

// 开发分类 → 流向映射规则（与 V6 保持一致）
const FLOW_DIRECTION_RULES = {
    '702': 'XX→ERP',
    '703': 'XX→ERP',
    '704': 'XX→ERP',
    '705': 'XX→ERP',
    '710': 'ERP→XX',
    '711': 'ERP→XX',
    '202': 'ERP→XX',
};

// 集成产品编号 → 外部系统映射（与 V6 保持一致）
const INTEGRATE_PRODUCT_RULES = {
    'CRM': 'CRM', 'OA': 'OA', 'WMS': 'WMS', 'MES': 'MES', 'SRM': 'SRM',
    'PLM': 'PLM', 'HR': 'OA', 'PIM': 'PIM', 'BI': 'BI', 'QMS': 'QMS',
    'TMS': 'TMS', 'APS': 'APS', 'SCM': 'SCM', 'EAM': 'EAM', 'FMS': 'FMS',
    'BPM': 'BPM', 'ESB': 'ESB', 'MDM': 'MDM', 'DMS': 'DMS', 'LIMS': 'LIMS',
    'OMS': 'OMS', 'CMS': 'CMS', 'POS': 'POS', 'EC': 'EC', 'APP': 'APP',
    'WECHAT': '微信', 'DINGTALK': '钉钉', 'FEISHU': '飞书',
};

/**
 * 解析命令行参数
 */
function parseArgs() {
    const args = process.argv.slice(2);
    const options = {
        apiKey: process.env.KIMI_API_KEY || '',
        threshold: DEFAULT_THRESHOLD,
        outputDir: DEFAULT_OUTPUT_DIR,
        all: false,
        model: DEFAULT_MODEL,
        maxRetries: DEFAULT_MAX_RETRIES,
        help: false,
    };

    let inputFile = null;

    for (let i = 0; i < args.length; i++) {
        const arg = args[i];
        switch (arg) {
            case '--api-key':
                options.apiKey = args[++i] || '';
                break;
            case '--threshold':
                options.threshold = parseInt(args[++i], 10) || DEFAULT_THRESHOLD;
                break;
            case '--output':
                options.outputDir = args[++i] || DEFAULT_OUTPUT_DIR;
                break;
            case '--all':
                options.all = true;
                break;
            case '--model':
                options.model = args[++i] || DEFAULT_MODEL;
                break;
            case '--max-retries':
                options.maxRetries = parseInt(args[++i], 10) || DEFAULT_MAX_RETRIES;
                break;
            case '--help':
                options.help = true;
                break;
            default:
                if (!arg.startsWith('--') && !inputFile) {
                    inputFile = arg;
                }
                break;
        }
    }

    return { inputFile, options };
}

/**
 * 显示帮助信息
 */
function showHelp() {
    console.log(`
API Document Retrieval - AI Assist Matching
使用 Kimi API 进行语义检索匹配，作为本地规则匹配失败时的兜底方案

用法:
  node match_ai_assist.js <需求清单xlsx文件> [选项]

选项:
  --api-key <key>     指定 Kimi API Key（也可通过 KIMI_API_KEY 环境变量设置）
  --threshold <n>     本地匹配度阈值，低于此值才调用 AI（默认 40）
  --output <dir>      输出目录（默认 集成需求字段清单产出/）
  --all               对所有行都调用 AI 分析（不筛选失败行）
  --model <model>     指定 Kimi 模型（默认 moonshot-v1-32k）
  --max-retries <n>   API 调用失败最大重试次数（默认 3）
  --help              显示帮助信息

示例:
  node match_ai_assist.js 需求评估整理/AI输入摸板.xlsx
  node match_ai_assist.js 需求评估整理/AI输入摸板.xlsx --api-key sk-xxx
  node match_ai_assist.js 需求评估整理/AI输入摸板.xlsx --all --threshold 50

环境变量:
  KIMI_API_KEY        Kimi API Key
`);
}

/**
 * 从集成产品编号中提取外部系统
 */
function extractExternalSystem(integrateProduct) {
    if (!integrateProduct) return 'XX';
    const productUpper = String(integrateProduct).toUpperCase().trim();

    if (INTEGRATE_PRODUCT_RULES[productUpper]) {
        return INTEGRATE_PRODUCT_RULES[productUpper];
    }

    const match = productUpper.match(/:\s*([A-Z]+)/);
    if (match) {
        const systemCode = match[1];
        if (INTEGRATE_PRODUCT_RULES[systemCode]) {
            return INTEGRATE_PRODUCT_RULES[systemCode];
        }
        return systemCode;
    }

    for (const [key, value] of Object.entries(INTEGRATE_PRODUCT_RULES)) {
        if (productUpper.includes(key)) {
            return value;
        }
    }

    return 'XX';
}

/**
 * 根据开发分类和集成产品编号推断流向
 */
function inferFlowDirection(devClass, integrateProduct) {
    if (!devClass) return '';

    const devClassStr = String(devClass);
    const classCodeMatch = devClassStr.match(/(\d{3})/);
    const classCode = classCodeMatch ? classCodeMatch[1] : '';

    const isServerSide = devClassStr.includes('服务端') || ['702', '703', '704', '705'].includes(classCode);
    const isClientSide = devClassStr.includes('客户端') || ['710', '711', '202'].includes(classCode);

    let externalSystem = extractExternalSystem(integrateProduct);

    if (isServerSide) {
        return `${externalSystem}→ERP`;
    } else if (isClientSide) {
        return `ERP→${externalSystem}`;
    }

    if (classCode) {
        return FLOW_DIRECTION_RULES[classCode] || '';
    }

    return '';
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
 * 从接口清单汇总文件中加载接口列表
 */
function loadInterfaceList() {
    console.log(`正在加载接口清单汇总: ${path.basename(INTERFACE_LIST_FILE)}`);

    if (!fs.existsSync(INTERFACE_LIST_FILE)) {
        console.error(`错误: 找不到接口清单汇总文件: ${INTERFACE_LIST_FILE}`);
        console.error('\n请确保文件存在: 需求对照工具表/【2026】接口清单---汇总.xlsx');
        process.exit(1);
    }

    const wb = XLSX.readFile(INTERFACE_LIST_FILE);
    const allInterfaces = [];

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

    console.log(`已加载 ${allInterfaces.length} 条接口记录\n`);
    return allInterfaces;
}

/**
 * 读取需求清单文件
 */
function readRequirementFile(filePath) {
    console.log(`读取需求清单: ${path.basename(filePath)}`);

    const wb = XLSX.readFile(filePath);

    let targetSheetName = wb.SheetNames[0];
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
    return { requirements, headers };
}

/**
 * 筛选候选接口（基于程序代号前缀和外部系统做初步过滤）
 */
function filterCandidateInterfaces(requirement, allInterfaces) {
    const progCode = String(requirement['程序代号'] || requirement['程序代号(*)'] || '').toLowerCase().trim();
    const reqName = requirement['程序名称'] || requirement['程序名称(*)'] || '';
    const devClass = requirement['开发分类'] || requirement['开发分类(*)'] || '';
    const integrateProduct = requirement['集成产品编号'] || requirement['集成产品编号(*)'] || '';

    const inferredFlow = inferFlowDirection(devClass, integrateProduct);
    const externalSystem = extractExternalSystem(integrateProduct);

    // 提取程序代号前缀（前2-4个字母）
    const prefixMatch = progCode.match(/^([a-z]{2,4})/);
    const prefix = prefixMatch ? prefixMatch[1] : '';

    // 评分并排序候选接口
    const scoredCandidates = allInterfaces.map(api => {
        let score = 0;
        const flowName = api['流程名称'] || '';
        const apiDirection = api['接口方向'] || '';
        const flowNameLower = flowName.toLowerCase();

        // 程序代号前缀匹配
        if (prefix && flowNameLower.includes(prefix)) {
            score += 30;
        }

        // 完整程序代号匹配
        if (progCode && flowNameLower.includes(progCode)) {
            score += 50;
        }

        // 外部系统匹配
        if (externalSystem !== 'XX' && apiDirection.includes(externalSystem)) {
            score += 20;
        }

        // 流向匹配
        if (inferredFlow && apiDirection.includes(inferredFlow.split('→')[0])) {
            score += 10;
        }

        // 业务关键词匹配（从需求名称提取）
        const reqKeywords = extractBusinessKeywords(reqName);
        const flowKeywords = extractBusinessKeywords(flowName);
        const commonKeywords = reqKeywords.filter(kw => flowKeywords.includes(kw));
        score += commonKeywords.length * 5;

        return { api, score };
    });

    // 按分数排序，取前 30 个作为候选
    scoredCandidates.sort((a, b) => b.score - a.score);
    return scoredCandidates.slice(0, 30).map(item => item.api);
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
        '推送', '同步', '抛转', '回写', '更新', '创建', '生成', '过账'
    ];

    const textLower = String(text).toLowerCase();
    return keywords.filter(kw => textLower.includes(kw));
}

/**
 * 构造 Kimi API Prompt
 */
function buildPrompt(requirement, candidates) {
    const progCode = requirement['程序代号'] || requirement['程序代号(*)'] || '';
    const reqName = requirement['程序名称'] || requirement['程序名称(*)'] || '';
    const devClass = requirement['开发分类'] || requirement['开发分类(*)'] || '';
    const specDesc = requirement['规格说明'] || requirement['规格说明(*)'] || '';
    const integrateProduct = requirement['集成产品编号'] || requirement['集成产品编号(*)'] || '';
    const inferredFlow = inferFlowDirection(devClass, integrateProduct);

    const interfaceList = candidates.map((api, idx) => {
        return `${idx + 1}. 流程名称: ${api['流程名称'] || 'N/A'}
   接口方向: ${api['接口方向'] || 'N/A'}
   功能描述: ${api['功能描述'] || 'N/A'}
   适配类型: ${api['适配类型'] || 'N/A'}`;
    }).join('\n\n');

    return `你是一位T100 ERP系统集成专家。请根据以下需求信息，从提供的接口清单中找出最匹配的接口。

【需求信息】
- 程序代号: ${progCode}
- 程序名称: ${reqName}
- 开发分类: ${devClass}
- 规格说明: ${specDesc}
- 集成产品编号: ${integrateProduct}
- 推断流向: ${inferredFlow}

【候选接口清单】（共 ${candidates.length} 条）
${interfaceList}

请分析以上需求信息和候选接口，返回最匹配的 1-3 个接口。返回格式必须是严格的 JSON：

{
  "matched": true,
  "recommendations": [
    {
      "flowName": "流程名称",
      "confidence": 85,
      "reason": "匹配理由说明"
    }
  ],
  "analysis": "整体分析说明"
}

如果没有匹配的接口，返回：
{
  "matched": false,
  "recommendations": [],
  "analysis": "未匹配原因说明"
}

注意：
1. confidence 是 0-100 的整数，表示匹配确信度
2. reason 需要具体说明为什么匹配，包括程序代号、业务场景、操作类型等
3. analysis 是对整体匹配情况的分析总结`;
}

/**
 * 调用 Kimi API
 */
function callKimiAPI(prompt, apiKey, model, maxRetries) {
    return new Promise((resolve, reject) => {
        const requestData = JSON.stringify({
            model: model,
            messages: [
                { role: 'system', content: '你是一位T100 ERP系统集成专家，擅长根据需求信息匹配标准API接口。请严格按照要求的JSON格式返回结果。' },
                { role: 'user', content: prompt }
            ],
            temperature: 0.3,
            response_format: { type: 'json_object' }
        });

        const options = {
            hostname: 'api.moonshot.cn',
            port: 443,
            path: '/v1/chat/completions',
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`,
                'Content-Length': Buffer.byteLength(requestData)
            }
        };

        let retries = 0;

        function doRequest() {
            const req = https.request(options, (res) => {
                let data = '';

                res.on('data', (chunk) => {
                    data += chunk;
                });

                res.on('end', () => {
                    try {
                        const response = JSON.parse(data);

                        if (response.error) {
                            if (retries < maxRetries) {
                                retries++;
                                const delay = Math.pow(2, retries) * 1000;
                                console.log(`  API 错误，${delay}ms 后重试 (${retries}/${maxRetries})...`);
                                setTimeout(doRequest, delay);
                                return;
                            }
                            reject(new Error(`Kimi API 错误: ${response.error.message}`));
                            return;
                        }

                        if (!response.choices || response.choices.length === 0) {
                            reject(new Error('Kimi API 返回空结果'));
                            return;
                        }

                        const content = response.choices[0].message.content;
                        resolve(content);
                    } catch (e) {
                        if (retries < maxRetries) {
                            retries++;
                            const delay = Math.pow(2, retries) * 1000;
                            console.log(`  解析错误，${delay}ms 后重试 (${retries}/${maxRetries})...`);
                            setTimeout(doRequest, delay);
                            return;
                        }
                        reject(new Error(`解析 Kimi API 响应失败: ${e.message}`));
                    }
                });
            });

            req.on('error', (error) => {
                if (retries < maxRetries) {
                    retries++;
                    const delay = Math.pow(2, retries) * 1000;
                    console.log(`  网络错误，${delay}ms 后重试 (${retries}/${maxRetries})...`);
                    setTimeout(doRequest, delay);
                    return;
                }
                reject(new Error(`调用 Kimi API 失败: ${error.message}`));
            });

            req.on('timeout', () => {
                req.destroy();
                if (retries < maxRetries) {
                    retries++;
                    const delay = Math.pow(2, retries) * 1000;
                    console.log(`  请求超时，${delay}ms 后重试 (${retries}/${maxRetries})...`);
                    setTimeout(doRequest, delay);
                    return;
                }
                reject(new Error('调用 Kimi API 超时'));
            });

            req.setTimeout(60000);
            req.write(requestData);
            req.end();
        }

        doRequest();
    });
}

/**
 * 解析 Kimi API 返回结果
 */
function parseAIResponse(content) {
    try {
        // 尝试直接解析 JSON
        const result = JSON.parse(content);
        return {
            matched: result.matched || false,
            recommendations: result.recommendations || [],
            analysis: result.analysis || ''
        };
    } catch (e) {
        // 如果直接解析失败，尝试提取 JSON 部分
        const jsonMatch = content.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
            try {
                const result = JSON.parse(jsonMatch[0]);
                return {
                    matched: result.matched || false,
                    recommendations: result.recommendations || [],
                    analysis: result.analysis || ''
                };
            } catch (e2) {
                // 忽略
            }
        }

        // 如果都无法解析，返回空结果
        return {
            matched: false,
            recommendations: [],
            analysis: `无法解析 AI 返回结果: ${content.substring(0, 200)}`
        };
    }
}

/**
 * 生成 Markdown 报告
 */
function generateMarkdownReport(requirements, aiResults, localResults, options) {
    const timestamp = new Date().toLocaleString('zh-CN');
    const report = [];

    report.push('# AI 辅助匹配报告');
    report.push('');
    report.push(`生成时间: ${timestamp}`);
    report.push(`使用模型: ${options.model}`);
    report.push(`匹配阈值: ${options.threshold}%`);
    report.push('');

    // 统计
    let localMatched = 0;
    let aiMatched = 0;
    let stillUnmatched = 0;

    requirements.forEach((req, idx) => {
        const localResult = localResults[idx];
        const aiResult = aiResults[idx];

        if (localResult && localResult.score >= options.threshold) {
            localMatched++;
        } else if (aiResult && aiResult.matched && aiResult.recommendations.length > 0) {
            aiMatched++;
        } else {
            stillUnmatched++;
        }
    });

    report.push('## 匹配统计');
    report.push(`- 本地规则匹配成功: ${localMatched} 条`);
    report.push(`- AI 辅助匹配成功: ${aiMatched} 条`);
    report.push(`- 仍未匹配: ${stillUnmatched} 条`);
    report.push(`- 总计: ${requirements.length} 条`);
    report.push('');

    // AI 匹配结果详情
    report.push('## AI 匹配结果详情');
    report.push('');

    let aiResultCount = 0;
    requirements.forEach((req, idx) => {
        const localResult = localResults[idx];
        const aiResult = aiResults[idx];

        // 只显示本地匹配失败但 AI 匹配成功的，或 --all 模式下所有行
        const isLocalFailed = !localResult || localResult.score < options.threshold;

        if (!isLocalFailed && !options.all) {
            return;
        }

        if (!aiResult || !aiResult.matched || aiResult.recommendations.length === 0) {
            return;
        }

        aiResultCount++;
        const progCode = req['程序代号'] || req['程序代号(*)'] || '';
        const reqName = req['程序名称'] || req['程序名称(*)'] || '';
        const devClass = req['开发分类'] || req['开发分类(*)'] || '';

        report.push(`### ${aiResultCount}. 程序代号: ${progCode} | 程序名称: ${reqName}`);
        report.push('');
        report.push(`- **行号**: ${req._rowNum}`);
        report.push(`- **开发分类**: ${devClass}`);
        report.push(`- **本地匹配结果**: ${localResult ? `匹配度 ${localResult.score}%` : '未匹配'}`);
        report.push('');

        aiResult.recommendations.forEach((rec, recIdx) => {
            report.push(`**AI 推荐 #${recIdx + 1}**:`);
            report.push(`- 流程名称: ${rec.flowName || 'N/A'}`);
            report.push(`- 匹配确信度: ${rec.confidence || 0}%`);
            report.push(`- 匹配理由: ${rec.reason || 'N/A'}`);
            report.push('');
        });

        report.push(`**AI 分析**: ${aiResult.analysis || 'N/A'}`);
        report.push('');
        report.push('---');
        report.push('');
    });

    // 仍未匹配
    report.push('## 仍未匹配的需求');
    report.push('');

    let unmatchedCount = 0;
    requirements.forEach((req, idx) => {
        const localResult = localResults[idx];
        const aiResult = aiResults[idx];

        const isLocalFailed = !localResult || localResult.score < options.threshold;
        const isAIFailed = !aiResult || !aiResult.matched || aiResult.recommendations.length === 0;

        if (isLocalFailed && isAIFailed) {
            unmatchedCount++;
            const progCode = req['程序代号'] || req['程序代号(*)'] || '';
            const reqName = req['程序名称'] || req['程序名称(*)'] || '';
            report.push(`- 行 ${req._rowNum}: ${progCode} - ${reqName}`);
        }
    });

    if (unmatchedCount === 0) {
        report.push('无');
    }

    report.push('');
    report.push('---');
    report.push('*本报告由 API Document Retrieval AI Assist 自动生成*');

    return report.join('\n');
}

/**
 * 生成 JSON 数据文件
 */
function generateJSONData(requirements, aiResults, localResults) {
    const data = [];

    requirements.forEach((req, idx) => {
        const localResult = localResults[idx];
        const aiResult = aiResults[idx];

        data.push({
            rowNum: req._rowNum,
            progCode: req['程序代号'] || req['程序代号(*)'] || '',
            reqName: req['程序名称'] || req['程序名称(*)'] || '',
            devClass: req['开发分类'] || req['开发分类(*)'] || '',
            localMatch: localResult ? {
                flowName: localResult.api ? localResult.api['流程名称'] : null,
                score: localResult.score,
                reason: localResult.reason
            } : null,
            aiMatch: aiResult && aiResult.matched ? {
                recommendations: aiResult.recommendations,
                analysis: aiResult.analysis
            } : null
        });
    });

    return JSON.stringify(data, null, 2);
}

/**
 * 主函数
 */
async function main() {
    const { inputFile, options } = parseArgs();

    if (options.help || !inputFile) {
        showHelp();
        if (!inputFile && !options.help) {
            console.error('\n错误: 请提供需求清单文件路径');
            process.exit(1);
        }
        process.exit(0);
    }

    // 检查 API Key
    if (!options.apiKey) {
        console.error('错误: 未设置 Kimi API Key');
        console.error('\n请通过以下方式之一设置 API Key:');
        console.error('  1. 环境变量: KIMI_API_KEY=sk-xxx');
        console.error('  2. 命令行参数: --api-key sk-xxx');
        process.exit(1);
    }

    console.log('='.repeat(80));
    console.log('     API Document Retrieval - AI Assist Matching');
    console.log('     使用 Kimi API 进行语义检索匹配');
    console.log('     模型: ' + options.model);
    console.log('     阈值: ' + options.threshold + '%');
    console.log('     模式: ' + (options.all ? '全量分析' : '失败兜底'));
    console.log('='.repeat(80));
    console.log();

    const filePath = path.resolve(inputFile);

    if (!fs.existsSync(filePath)) {
        console.error(`错误: 找不到文件 ${inputFile}`);
        process.exit(1);
    }

    if (!filePath.endsWith('.xlsx')) {
        console.error('错误: 仅支持 Excel 文件 (.xlsx)');
        process.exit(1);
    }

    // 加载数据
    const apiData = loadInterfaceList();
    const { requirements } = readRequirementFile(filePath);

    if (requirements.length === 0) {
        console.log('未找到有效的需求记录');
        process.exit(1);
    }

    // 模拟本地规则匹配结果（实际使用时，这里应该调用 V6 的 matchInterface 函数）
    // 为了独立运行，我们做一个简化的本地匹配
    console.log('执行本地规则匹配（简化版）...');
    console.log('='.repeat(80));

    const localResults = [];
    requirements.forEach(req => {
        const progCode = String(req['程序代号'] || req['程序代号(*)'] || '').toLowerCase().trim();
        const devClass = req['开发分类'] || req['开发分类(*)'] || '';
        const integrateProduct = req['集成产品编号'] || req['集成产品编号(*)'] || '';

        // 201 开发分类跳过
        const devClassStr = String(devClass);
        const classCodeMatch = devClassStr.match(/(\d{3})/);
        const classCode = classCodeMatch ? classCodeMatch[1] : '';
        if (classCode === '201' || devClassStr.includes('建档')) {
            localResults.push({ score: 0, api: null, reason: '201开发分类跳过' });
            return;
        }

        const inferredFlow = inferFlowDirection(devClass, integrateProduct);
        let bestScore = 0;
        let bestMatch = null;
        let bestReason = '';

        for (const apiRow of apiData) {
            const flowName = apiRow['流程名称'] || '';
            const apiDirection = apiRow['接口方向'] || '';
            if (!flowName) continue;

            let score = 0;
            const reasons = [];
            const flowNameLower = flowName.toLowerCase();

            // 程序代号匹配
            const apiProgCode = extractProgCode(flowName);
            if (apiProgCode === progCode) {
                score += 60;
                reasons.push(`程序代号精确匹配: ${progCode}`);
            } else if (flowNameLower.includes(progCode)) {
                score += 40;
                reasons.push(`程序代号包含匹配: ${progCode}`);
            }

            // 流向匹配
            if (inferredFlow && score > 0) {
                const normalizedApiDir = apiDirection.replace(/T100/g, 'ERP').replace(/->/g, '→');
                if (normalizedApiDir.includes(inferredFlow.split('→')[0])) {
                    score += 30;
                    reasons.push(`流向匹配: ${inferredFlow}`);
                }
            }

            if (score > bestScore) {
                bestScore = score;
                bestMatch = apiRow;
                bestReason = reasons.join(' | ');
            }
        }

        localResults.push({
            score: bestScore,
            api: bestMatch,
            reason: bestReason
        });
    });

    console.log('\n本地规则匹配完成');
    console.log('='.repeat(80));

    // 确定需要 AI 辅助分析的行
    const needAIAnalysis = [];
    requirements.forEach((req, idx) => {
        const localResult = localResults[idx];
        const isLocalFailed = !localResult || localResult.score < options.threshold;

        if (options.all || isLocalFailed) {
            needAIAnalysis.push(idx);
        }
    });

    console.log(`\n需要 AI 辅助分析: ${needAIAnalysis.length} 条`);
    console.log('='.repeat(80));

    // 调用 Kimi API 进行 AI 辅助匹配
    const aiResults = new Array(requirements.length).fill(null);

    for (let i = 0; i < needAIAnalysis.length; i++) {
        const idx = needAIAnalysis[i];
        const req = requirements[idx];
        const progCode = req['程序代号'] || req['程序代号(*)'] || '';
        const reqName = req['程序名称'] || req['程序名称(*)'] || '';

        console.log(`\n[${i + 1}/${needAIAnalysis.length}] AI 分析: ${progCode} - ${reqName.substring(0, 40)}...`);

        // 筛选候选接口
        const candidates = filterCandidateInterfaces(req, apiData);
        console.log(`  候选接口: ${candidates.length} 条`);

        if (candidates.length === 0) {
            console.log('  无候选接口，跳过 AI 分析');
            aiResults[idx] = {
                matched: false,
                recommendations: [],
                analysis: '无候选接口可供匹配'
            };
            continue;
        }

        // 构造 Prompt
        const prompt = buildPrompt(req, candidates);

        // 调用 Kimi API
        try {
            console.log('  调用 Kimi API...');
            const responseContent = await callKimiAPI(
                prompt,
                options.apiKey,
                options.model,
                options.maxRetries
            );

            const parsedResult = parseAIResponse(responseContent);
            aiResults[idx] = parsedResult;

            if (parsedResult.matched && parsedResult.recommendations.length > 0) {
                const best = parsedResult.recommendations[0];
                console.log(`  [✓] AI 匹配成功 (${best.confidence || 0}%) - ${best.flowName || 'N/A'}`);
            } else {
                console.log('  [✗] AI 未找到匹配');
            }
        } catch (error) {
            console.error(`  [✗] AI 调用失败: ${error.message}`);
            aiResults[idx] = {
                matched: false,
                recommendations: [],
                analysis: `AI 调用失败: ${error.message}`
            };
        }

        // 避免 Rate Limit，每 5 个请求暂停 1 秒
        if ((i + 1) % 5 === 0 && i < needAIAnalysis.length - 1) {
            console.log('  暂停 1 秒...');
            await new Promise(resolve => setTimeout(resolve, 1000));
        }
    }

    console.log('\n' + '='.repeat(80));
    console.log('AI 辅助匹配完成！');
    console.log('='.repeat(80));

    // 生成输出文件
    if (!fs.existsSync(options.outputDir)) {
        fs.mkdirSync(options.outputDir, { recursive: true });
    }

    const baseName = path.basename(inputFile, path.extname(inputFile));

    // 生成 Markdown 报告
    const markdownReport = generateMarkdownReport(requirements, aiResults, localResults, options);
    const markdownFile = path.join(options.outputDir, `${baseName}_AI辅助匹配报告.md`);
    fs.writeFileSync(markdownFile, markdownReport, 'utf-8');
    console.log(`\nMarkdown 报告: ${markdownFile}`);

    // 生成 JSON 数据文件
    const jsonData = generateJSONData(requirements, aiResults, localResults);
    const jsonFile = path.join(options.outputDir, `${baseName}_ai_matches.json`);
    fs.writeFileSync(jsonFile, jsonData, 'utf-8');
    console.log(`JSON 数据文件: ${jsonFile}`);

    console.log('\n' + '='.repeat(80));
    console.log('全部完成！');
    console.log('='.repeat(80));
}

// 运行主函数
main().catch(err => {
    console.error('发生错误:', err);
    process.exit(1);
});
