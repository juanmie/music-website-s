#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
能率需求清单-bpm集成 - 智能API匹配工具 V2
基于需求说明和程序代号进行精确匹配
"""

import sys
import os
from pathlib import Path
import xlrd
import xlwt
from datetime import datetime
import shutil
import re

# 配置
WORKSPACE_DIR = Path(__file__).parent.parent.parent.parent.parent
# 优先使用 .xls 版本（解决 Python 3.14 兼容性问题）
API_MAPPING_FILE_XLS = WORKSPACE_DIR / "需求对照工具表" / "【2026】T100&OA类系统集成清单字段对照表.xls"
API_MAPPING_FILE_XLSX = WORKSPACE_DIR / "需求对照工具表" / "【2026】T100&OA类系统集成清单字段对照表.xlsx"

# 选择可用的文件
if API_MAPPING_FILE_XLS.exists():
    API_MAPPING_FILE = API_MAPPING_FILE_XLS
elif API_MAPPING_FILE_XLSX.exists():
    API_MAPPING_FILE = API_MAPPING_FILE_XLSX
else:
    API_MAPPING_FILE = API_MAPPING_FILE_XLSX

ONLINE_DOC_URL = "https://docs.qq.com/sheet/DSmdhek1TT3BPenRa?tab=BB08J2"


def load_api_mapping():
    """加载API对照清单"""
    if not API_MAPPING_FILE.exists():
        print(f"警告: 找不到本地API对照清单文件")
        print(f"  .xls: {API_MAPPING_FILE_XLS}")
        print(f"  .xlsx: {API_MAPPING_FILE_XLSX}")
        print(f"\n建议使用在线文档: {ONLINE_DOC_URL}")
        sys.exit(1)
    
    print(f"正在加载API对照清单: {API_MAPPING_FILE.name}")
    
    try:
        wb = xlrd.open_workbook(API_MAPPING_FILE)
        print(f"可用页签: {wb.sheet_names()}")
        
        # 查找"接口清单"页签
        interface_sheet = None
        for sheet_name in wb.sheet_names():
            if '接口清单' in sheet_name or 'interface' in sheet_name.lower():
                interface_sheet = wb.sheet_by_name(sheet_name)
                print(f"找到接口清单页签: {sheet_name}")
                break
        
        if not interface_sheet:
            print(f"未找到'接口清单'页签，使用第一个页签: {wb.sheet_names()[0]}")
            interface_sheet = wb.sheet_by_index(0)
        
        # 读取表头
        headers = []
        for col_idx in range(interface_sheet.ncols):
            cell_value = interface_sheet.cell(0, col_idx).value
            if cell_value:
                headers.append({
                    'col': col_idx,
                    'name': str(cell_value).strip()
                })
        
        # 读取所有行数据
        api_data = {
            'sheet_name': interface_sheet.name,
            'headers': headers,
            'rows': []
        }
        
        for row_idx in range(1, interface_sheet.nrows):
            row_data = {'_row_num': row_idx + 1, '_sheet': interface_sheet.name}
            for header in headers:
                cell_value = interface_sheet.cell(row_idx, header['col']).value
                row_data[header['name']] = cell_value if cell_value else ''
            
            api_data['rows'].append(row_data)
        
        print(f"已加载 {len(api_data['rows'])} 条接口记录\n")
        return api_data
        
    except Exception as e:
        print(f"读取失败: {e}")
        print(f"\n建议使用在线文档: {ONLINE_DOC_URL}")
        sys.exit(1)


def read_requirement_xls(file_path):
    """读取能率需求清单"""
    print(f"读取需求清单: {file_path.name}")
    
    wb = xlrd.open_workbook(file_path)
    ws = wb.sheet_by_index(0)
    print(f"工作表: {ws.name}")
    print(f"行数: {ws.nrows}")
    print(f"列数: {ws.ncols}\n")
    
    # 提取表头
    headers = []
    for col_idx in range(ws.ncols):
        cell_value = ws.cell(0, col_idx).value
        if cell_value:
            headers.append({
                'col': col_idx,
                'name': str(cell_value).strip()
            })
    
    print(f"表头字段: {[h['name'] for h in headers]}\n")
    
    # 提取需求数据
    requirements = []
    for row_idx in range(1, ws.nrows):
        req_data = {'_row_num': row_idx + 1}
        
        for header in headers:
            cell_value = ws.cell(row_idx, header['col']).value
            req_data[header['name']] = cell_value if cell_value else ''
        
        if req_data.get('序号(*)', '') or req_data.get('程序代号(*)', ''):
            requirements.append(req_data)
    
    print(f"提取到 {len(requirements)} 条需求记录\n")
    return requirements, headers


def extract_keywords_from_description(desc):
    """从需求说明中提取关键信息"""
    keywords = {
        'program_codes': [],  # 程序代号
        'actions': [],  # 动作：推送、签核、回传、作废等
        'doc_types': [],  # 单据类型
        'direction': '',  # 流向
        'is_bpm': False,  # 是否BPM集成
        'has_process_id': False  # 是否需要流程实例ID
    }
    
    if not desc:
        return keywords
    
    desc_lower = desc.lower()
    
    # 提取程序代号（如 apmt200, wssp085 等）
    prog_pattern = r'[a-z]{2,4}\d{3}'
    codes = re.findall(prog_pattern, desc_lower)
    keywords['program_codes'] = list(set(codes))
    
    # 提取动作关键词
    action_words = ['推送', '签核', '回传', '回写', '作废', '同步', '创建', '新增', '更新', '查询']
    for action in action_words:
        if action in desc:
            keywords['actions'].append(action)
    
    # 提取单据类型
    doc_types = ['供应商', '核价单', '销售订单', '采购订单', '杂收单', '杂发单', '仓退单', '价格表']
    for doc_type in doc_types:
        if doc_type in desc:
            keywords['doc_types'].append(doc_type)
    
    # 判断流向
    if 'BPM' in desc or 'bpm' in desc_lower:
        keywords['is_bpm'] = True
        if '推送BPM' in desc or '发起BPM' in desc:
            keywords['direction'] = 'T100→BPM'
        elif 'BPM回传' in desc or 'BPM签核回传' in desc or '回传更新' in desc:
            keywords['direction'] = 'BPM→T100'
        elif 'BPM发起' in desc or 'BPM新增' in desc or 'BPM作废' in desc:
            keywords['direction'] = 'BPM→T100'
    
    # 检查是否需要流程实例ID
    if '流程实例id' in desc or '流程实例ID' in desc or '审批示例ID' in desc:
        keywords['has_process_id'] = True
    
    return keywords


def calculate_match_score(req, api_row):
    """计算需求与API的匹配度"""
    score = 0
    match_reasons = []
    
    # 获取需求信息
    prog_code = str(req.get('程序代号(*)', '')).lower()
    req_name = str(req.get('程序名称(*)', ''))
    dev_class = str(req.get('开发分类(*)', ''))
    req_desc = str(req.get('规格说明(*)', ''))
    custom_type = str(req.get('客制类型(*)', ''))
    
    # 提取需求说明中的关键词
    keywords = extract_keywords_from_description(req_desc)
    
    # 获取API信息
    flow_name = str(api_row.get('流程名称', ''))
    api_name = str(api_row.get('接口名称', ''))
    api_direction = str(api_row.get('接口方向', ''))
    
    # ========== 匹配规则 1: 程序代号精确匹配（最高优先级 60%） ==========
    if prog_code and prog_code in flow_name.lower():
        score += 60
        match_reasons.append(f"程序代号精确匹配: {prog_code}")
    elif prog_code and prog_code in api_name.lower():
        score += 50
        match_reasons.append(f"程序代号在接口名称中: {prog_code}")
    
    # ========== 匹配规则 2: 规格说明中的程序代号匹配（20%） ==========
    if keywords['program_codes']:
        for code in keywords['program_codes']:
            if code in flow_name.lower() or code in api_name.lower():
                score += 20
                match_reasons.append(f"规格说明中程序代号匹配: {code}")
                break
    
    # ========== 匹配规则 3: 开发分类和流向匹配（10%） ==========
    if '710' in dev_class and '推送接口' in dev_class:
        if keywords['direction'] == 'T100→BPM':
            if 'T100' in flow_name or 'ERP' in flow_name:
                score += 10
                match_reasons.append("开发分类为推送接口且流向T100→BPM")
    elif '702' in dev_class and '生单接口' in dev_class:
        if keywords['direction'] == 'BPM→T100':
            score += 10
            match_reasons.append("开发分类为生单接口且流向BPM→T100")
    elif '704' in dev_class and '更新接口' in dev_class:
        if keywords['direction'] == 'BPM→T100' and ('更新' in keywords['actions'] or '回传' in keywords['actions']):
            score += 10
            match_reasons.append("开发分类为更新接口且BPM回传更新")
    
    # ========== 匹配规则 4: 客制类型（5%） ==========
    if '新增' in custom_type:
        score += 5
        match_reasons.append("客制类型为新增")
    
    # ========== 匹配规则 5: 单据类型关键词匹配（5%） ==========
    if keywords['doc_types']:
        for doc_type in keywords['doc_types']:
            if doc_type in flow_name or doc_type in api_name:
                score += 5
                match_reasons.append(f"单据类型匹配: {doc_type}")
                break
    
    return min(score, 100), ' | '.join(match_reasons)


def match_requirements(requirements, api_data):
    """为每个需求匹配API"""
    print("开始匹配API...")
    print("=" * 80)
    
    results = []
    matched_count = 0
    
    for i, req in enumerate(requirements, 1):
        prog_code = req.get('程序代号(*)', '')
        req_name = req.get('程序名称(*)', '')
        dev_class = req.get('开发分类(*)', '')
        
        print(f"[{i}/{len(requirements)}] 处理: {prog_code} - {req_name[:30]}...")
        
        # 计算所有API的匹配度
        matches = []
        for api_row in api_data['rows']:
            score, reason = calculate_match_score(req, api_row)
            if score >= 50:  # 匹配度阈值
                matches.append({
                    'score': score,
                    'reason': reason,
                    'api': api_row
                })
        
        # 按匹配度排序
        matches.sort(key=lambda x: x['score'], reverse=True)
        
        if matches:
            matched_count += 1
            best_match = matches[0]
            print(f"  [OK] 匹配成功 (匹配度: {best_match['score']}%)")
            print(f"  匹配接口: {best_match['api'].get('接口名称', 'N/A')}")
            print(f"  匹配依据: {best_match['reason']}\n")
        else:
            print(f"  [WARN] 未找到匹配\n")
        
        results.append({
            'requirement': req,
            'matches': matches
        })
    
    print("=" * 80)
    print(f"匹配完成: {matched_count}/{len(requirements)} 条需求找到匹配API\n")
    return results


def generate_matched_excel(requirements, results, original_file):
    """生成包含匹配结果的新Excel文件"""
    print("生成匹配结果文件...")
    
    # 备份原文件
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_file = original_file.parent / f"{original_file.stem}_备份_{timestamp}{original_file.suffix}"
    shutil.copy2(original_file, backup_file)
    print(f"原文件已备份: {backup_file.name}")
    
    # 创建新的工作簿
    wb_out = xlwt.Workbook()
    ws_out = wb_out.add_sheet("需求匹配结果")
    
    # 读取原始xls获取表头
    wb_in = xlrd.open_workbook(original_file)
    ws_in = wb_in.sheet_by_index(0)
    
    # 写入原始表头
    original_headers = []
    for col_idx in range(ws_in.ncols):
        cell_value = ws_in.cell(0, col_idx).value
        original_headers.append(str(cell_value) if cell_value else '')
    
    # 添加匹配结果列
    match_headers = [
        '匹配状态',
        '最佳匹配接口',
        '流程名称',
        '接口方向',
        '匹配度',
        '匹配依据',
        '参考位置(工作表-行号)',
        '在线文档链接'
    ]
    
    all_headers = original_headers + match_headers
    
    # 创建表头样式
    header_style = xlwt.XFStyle()
    header_style.font = xlwt.Font()
    header_style.font.bold = True
    header_style.font.colour_index = xlwt.Style.colour_map['white']
    header_style.pattern = xlwt.Pattern()
    header_style.pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    header_style.pattern.pattern_fore_colour = xlwt.Style.colour_map['blue']
    header_style.alignment = xlwt.Alignment()
    header_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
    header_style.alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT
    
    # 写入表头
    for col_idx, header in enumerate(all_headers):
        ws_out.write(0, col_idx, header, header_style)
    
    # 写入数据
    for i, result in enumerate(results, 1):
        req = result['requirement']
        matches = result['matches']
        
        # 写入原始数据
        for col_idx in range(len(original_headers)):
            header = original_headers[col_idx]
            value = req.get(header, '')
            ws_out.write(i, col_idx, value)
        
        # 写入匹配结果
        base_col = len(original_headers)
        
        if matches:
            best = matches[0]
            api = best['api']
            
            ws_out.write(i, base_col, "已匹配")
            ws_out.write(i, base_col + 1, api.get('接口名称', ''))
            ws_out.write(i, base_col + 2, api.get('流程名称', ''))
            ws_out.write(i, base_col + 3, api.get('接口方向', ''))
            ws_out.write(i, base_col + 4, f"{best['score']}%")
            ws_out.write(i, base_col + 5, best['reason'])
            ws_out.write(i, base_col + 6, f"{api.get('_sheet', '')} 行{api.get('_row_num', '')}")
            ws_out.write(i, base_col + 7, ONLINE_DOC_URL)
        else:
            ws_out.write(i, base_col, "未匹配")
    
    # 设置列宽
    for col_idx in range(len(all_headers)):
        ws_out.col(col_idx).width = 256 * 20
    
    # 保存文件
    output_file = original_file.parent / f"{original_file.stem}_已匹配{original_file.suffix}"
    wb_out.save(output_file)
    print(f"匹配结果已保存: {output_file.name}")
    print(f"文件大小: {output_file.stat().st_size / 1024:.2f} KB\n")
    
    return output_file


def main():
    """主函数"""
    print("=" * 80)
    print("              能率需求清单-bpm集成 - 智能API匹配工具 V2")
    print("=" * 80)
    print()
    
    if len(sys.argv) < 2:
        print("用法: python match_nenglv_bpm_v2.py <需求清单文件>")
        print("\n示例:")
        print("  python match_nenglv_bpm_v2.py 需求评估整理/能率需求清单-bpm集成.xls")
        sys.exit(1)
    
    requirement_file = sys.argv[1]
    
    if not Path(requirement_file).exists():
        print(f"错误: 找不到文件 {requirement_file}")
        sys.exit(1)
    
    if not requirement_file.endswith(('.xls', '.xlsx')):
        print("错误: 仅支持 Excel 文件 (.xls, .xlsx)")
        sys.exit(1)
    
    # 步骤1: 加载API对照清单
    api_data = load_api_mapping()
    
    # 步骤2: 读取需求清单
    requirements, headers = read_requirement_xls(Path(requirement_file))
    
    # 步骤3: 执行匹配
    results = match_requirements(requirements, api_data)
    
    # 步骤4: 生成结果文件
    output_file = generate_matched_excel(requirements, results, Path(requirement_file))
    
    print("=" * 80)
    print("匹配完成！")
    print("=" * 80)
    print(f"\n输出文件: {output_file}")
    print(f"在线文档: {ONLINE_DOC_URL}")
    print(f"\n请打开输出文件查看匹配结果！")


if __name__ == '__main__':
    main()
