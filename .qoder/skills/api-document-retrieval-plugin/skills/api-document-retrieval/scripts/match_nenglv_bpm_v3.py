#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
能率需求清单-bpm集成 - 智能API匹配工具 V3
基于页签名称进行匹配（每个接口一个页签）
"""

import sys
from pathlib import Path
import xlrd
import xlwt
from datetime import datetime
import shutil
import re

# 配置
WORKSPACE_DIR = Path(__file__).parent.parent.parent.parent.parent
API_MAPPING_FILE_XLS = WORKSPACE_DIR / "需求对照工具表" / "【2026】T100&OA类系统集成清单字段对照表.xls"
API_MAPPING_FILE_XLSX = WORKSPACE_DIR / "需求对照工具表" / "【2026】T100&OA类系统集成清单字段对照表.xlsx"

if API_MAPPING_FILE_XLS.exists():
    API_MAPPING_FILE = API_MAPPING_FILE_XLS
elif API_MAPPING_FILE_XLSX.exists():
    API_MAPPING_FILE = API_MAPPING_FILE_XLSX
else:
    API_MAPPING_FILE = API_MAPPING_FILE_XLSX

ONLINE_DOC_URL = "https://docs.qq.com/sheet/DSmdhek1TT3BPenRa?tab=BB08J2"


def load_api_mapping():
    """加载API对照清单 - 基于页签名称"""
    if not API_MAPPING_FILE.exists():
        print(f"警告: 找不到本地API对照清单文件")
        sys.exit(1)
    
    print(f"正在加载API对照清单: {API_MAPPING_FILE.name}")
    
    wb = xlrd.open_workbook(API_MAPPING_FILE)
    sheet_names = wb.sheet_names()
    
    print(f"总页签数: {len(sheet_names)}")
    print(f"部分页签: {sheet_names[:10]}\n")
    
    # 构建API索引（基于页签名称）
    api_list = []
    for sheet_name in sheet_names:
        # 跳过非接口页签
        if sheet_name in ['接口清单', 'T100&OA集成清单', '接口规划', '接口实现统计']:
            continue
        
        # 从页签名称提取信息
        # 格式如: apmt200供应商申请送签OA(ERP-->OA)
        api_info = {
            'sheet_name': sheet_name,
            'program_code': '',
            'business_name': '',
            'direction': ''
        }
        
        # 提取程序代号（开头的字母+数字）
        match = re.match(r'^([a-z]{2,4}\d{3})', sheet_name.lower())
        if match:
            api_info['program_code'] = match.group(1)
        
        # 提取流向
        if 'ERP-->OA' in sheet_name or 'T100-->OA' in sheet_name:
            api_info['direction'] = 'T100→OA'
        elif 'OA-->ERP' in sheet_name or 'OA-->T100' in sheet_name:
            api_info['direction'] = 'OA→T100'
        
        # 提取业务名称（程序代号之后，流向之前的部分）
        if api_info['program_code']:
            remaining = sheet_name[len(api_info['program_code']):]
            # 去除流向部分
            for pattern in ['(ERP-->OA)', '(OA-->ERP)', '(T100-->OA)', '(OA-->T100)']:
                remaining = remaining.replace(pattern, '')
            api_info['business_name'] = remaining.strip()
        
        api_list.append(api_info)
    
    print(f"提取到 {len(api_list)} 个接口页签\n")
    return api_list


def read_requirement_xls(file_path):
    """读取能率需求清单"""
    print(f"读取需求清单: {file_path.name}")
    
    wb = xlrd.open_workbook(file_path)
    ws = wb.sheet_by_index(0)
    
    print(f"工作表: {ws.name}")
    print(f"行数: {ws.nrows}, 列数: {ws.ncols}\n")
    
    # 提取表头
    headers = []
    for col_idx in range(ws.ncols):
        cell_value = ws.cell(0, col_idx).value
        if cell_value:
            headers.append({
                'col': col_idx,
                'name': str(cell_value).strip()
            })
    
    # 提取需求数据
    requirements = []
    for row_idx in range(1, ws.nrows):
        req_data = {'_row_num': row_idx + 1}
        for header in headers:
            cell_value = ws.cell(row_idx, header['col']).value
            req_data[header['name']] = cell_value if cell_value else ''
        
        if req_data.get('程序代号(*)', ''):
            requirements.append(req_data)
    
    print(f"提取到 {len(requirements)} 条需求记录\n")
    return requirements, headers


def match_requirement_to_api(req, api_list):
    """将单个需求匹配到API"""
    prog_code = str(req.get('程序代号(*)', '')).lower().strip()
    req_name = str(req.get('程序名称(*)', ''))
    dev_class = str(req.get('开发分类(*)', ''))
    req_desc = str(req.get('规格说明(*)', ''))
    
    if not prog_code:
        return None, 0, ''
    
    # 查找匹配的API
    matches = []
    
    for api in api_list:
        score = 0
        reasons = []
        
        # 规则1: 程序代号精确匹配（80%）
        if api['program_code'] == prog_code:
            score += 80
            reasons.append(f"程序代号精确匹配: {prog_code}")
        
        # 规则2: 程序代号包含匹配（60%）
        elif prog_code in api['sheet_name'].lower() or api['program_code'] in prog_code:
            score += 60
            reasons.append(f"程序代号包含匹配: {prog_code}")
        
        # 规则3: 开发分类辅助判断（10%）
        if score > 0:
            if '710' in dev_class and '推送接口' in dev_class:
                if api['direction'] == 'T100→OA':
                    score += 10
                    reasons.append("推送接口方向匹配")
            elif '702' in dev_class and '生单接口' in dev_class:
                if api['direction'] == 'OA→T100':
                    score += 10
                    reasons.append("生单接口方向匹配")
            elif '704' in dev_class and '更新接口' in dev_class:
                if api['direction'] == 'OA→T100':
                    score += 10
                    reasons.append("更新接口方向匹配")
        
        # 规则4: 需求说明中提及的程序代号（5%）
        if prog_code in req_desc.lower() and score > 0:
            score += 5
            reasons.append("规格说明中提及")
        
        if score >= 60:  # 阈值
            matches.append({
                'api': api,
                'score': score,
                'reason': ' | '.join(reasons)
            })
    
    # 返回最佳匹配
    if matches:
        matches.sort(key=lambda x: x['score'], reverse=True)
        best = matches[0]
        return best['api'], best['score'], best['reason']
    
    return None, 0, ''


def generate_matched_excel(requirements, results, original_file, api_list):
    """生成包含匹配结果的新Excel文件"""
    print("\n生成匹配结果文件...")
    
    # 备份原文件
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_file = original_file.parent / f"{original_file.stem}_备份_{timestamp}.xls"
    shutil.copy2(original_file, backup_file)
    print(f"原文件已备份: {backup_file.name}")
    
    # 创建新的工作簿
    wb_out = xlwt.Workbook()
    ws_out = wb_out.add_sheet("需求匹配结果")
    
    # 读取原始表头
    wb_in = xlrd.open_workbook(original_file)
    ws_in = wb_in.sheet_by_index(0)
    
    original_headers = []
    for col_idx in range(ws_in.ncols):
        cell_value = ws_in.cell(0, col_idx).value
        original_headers.append(str(cell_value) if cell_value else '')
    
    # 添加匹配结果列
    match_headers = [
        '匹配状态',
        '匹配接口(页签名称)',
        '接口方向',
        '匹配度',
        '匹配依据',
        '在线文档链接'
    ]
    
    all_headers = original_headers + match_headers
    
    # 表头样式
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
        api = result['api']
        score = result['score']
        reason = result['reason']
        
        # 写入原始数据
        for col_idx in range(len(original_headers)):
            header = original_headers[col_idx]
            value = req.get(header, '')
            ws_out.write(i, col_idx, value)
        
        # 写入匹配结果
        base_col = len(original_headers)
        
        if api and score >= 60:
            ws_out.write(i, base_col, "已匹配")
            ws_out.write(i, base_col + 1, api['sheet_name'])
            ws_out.write(i, base_col + 2, api['direction'])
            ws_out.write(i, base_col + 3, f"{score}%")
            ws_out.write(i, base_col + 4, reason)
            ws_out.write(i, base_col + 5, ONLINE_DOC_URL)
        else:
            ws_out.write(i, base_col, "未匹配")
            ws_out.write(i, base_col + 1, "")
            ws_out.write(i, base_col + 2, "")
            ws_out.write(i, base_col + 3, "0%")
            ws_out.write(i, base_col + 4, "无匹配接口")
            ws_out.write(i, base_col + 5, ONLINE_DOC_URL)
    
    # 设置列宽
    for col_idx in range(len(all_headers)):
        ws_out.col(col_idx).width = 256 * 25
    
    # 保存文件（使用不同的文件名避免冲突）
    output_file = original_file.parent / f"{original_file.stem}_API匹配结果.xls"
    
    try:
        wb_out.save(output_file)
        print(f"匹配结果已保存: {output_file.name}")
        print(f"文件大小: {output_file.stat().st_size / 1024:.2f} KB\n")
    except PermissionError:
        # 如果文件被占用，使用时间戳
        output_file = original_file.parent / f"{original_file.stem}_API匹配结果_{timestamp}.xls"
        wb_out.save(output_file)
        print(f"匹配结果已保存: {output_file.name}")
        print(f"文件大小: {output_file.stat().st_size / 1024:.2f} KB\n")
    
    return output_file


def main():
    """主函数"""
    print("=" * 80)
    print("         能率需求清单-bpm集成 - 智能API匹配工具 V3")
    print("=" * 80)
    print()
    
    if len(sys.argv) < 2:
        print("用法: python match_nenglv_bpm_v3.py <需求清单文件>")
        sys.exit(1)
    
    requirement_file = sys.argv[1]
    
    if not Path(requirement_file).exists():
        print(f"错误: 找不到文件 {requirement_file}")
        sys.exit(1)
    
    # 步骤1: 加载API对照清单
    api_list = load_api_mapping()
    
    # 步骤2: 读取需求清单
    requirements, headers = read_requirement_xls(Path(requirement_file))
    
    # 步骤3: 执行匹配
    print("开始匹配API...")
    print("=" * 80)
    
    results = []
    matched_count = 0
    
    for i, req in enumerate(requirements, 1):
        prog_code = req.get('程序代号(*)', '')
        req_name = req.get('程序名称(*)', '')
        
        print(f"[{i}/{len(requirements)}] {prog_code} - {req_name[:30]}...")
        
        api, score, reason = match_requirement_to_api(req, api_list)
        
        if api and score >= 60:
            matched_count += 1
            print(f"  [OK] 匹配成功 ({score}%) - {api['sheet_name']}")
            print(f"  依据: {reason}\n")
        else:
            print(f"  [WARN] 未找到匹配\n")
        
        results.append({
            'requirement': req,
            'api': api,
            'score': score,
            'reason': reason
        })
    
    print("=" * 80)
    print(f"匹配完成: {matched_count}/{len(requirements)} 条需求找到匹配API\n")
    
    # 步骤4: 生成结果文件
    output_file = generate_matched_excel(requirements, results, Path(requirement_file), api_list)
    
    print("=" * 80)
    print("全部完成！")
    print("=" * 80)
    print(f"\n输出文件: {output_file}")
    print(f"在线文档: {ONLINE_DOC_URL}")
    print(f"\n请打开输出文件查看匹配结果！")


if __name__ == '__main__':
    main()
