#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
从在线文档自动下载 API 对照清单
支持腾讯文档导出功能（需要登录cookie或API token）
"""

import sys
import os
from pathlib import Path
import requests
from datetime import datetime

# 配置
WORKSPACE_DIR = Path(__file__).parent.parent.parent.parent
OUTPUT_DIR = WORKSPACE_DIR / "需求对照工具表"
ONLINE_DOC_URL = "https://docs.qq.com/sheet/DSmdhek1TT3BPenRa?tab=BB08J2"
OUTPUT_FILE = OUTPUT_DIR / "【2026】T100&OA类系统集成清单字段对照表.xlsx"

class TencentDocDownloader:
    """腾讯文档下载器"""
    
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
    
    def download_with_cookie(self, cookie_str):
        """使用 Cookie 下载（需要用户手动获取）
        
        Args:
            cookie_str: 浏览器中的 Cookie 字符串
        """
        print("📥 使用 Cookie 方式下载...")
        
        # 解析 Cookie
        cookies = {}
        for item in cookie_str.split(';'):
            if '=' in item:
                key, value = item.strip().split('=', 1)
                cookies[key] = value
        
        self.session.cookies.update(cookies)
        
        # 尝试导出 Excel
        # 注意：腾讯文档的导出URL可能变化，需要抓包获取
        export_url = "https://docs.qq.com/api/docs/export/DSmdhek1TT3BPenRa"
        
        try:
            response = self.session.get(export_url, stream=True)
            response.raise_for_status()
            
            # 保存文件
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            with open(OUTPUT_FILE, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            print(f"✓ 下载成功: {OUTPUT_FILE}")
            print(f"✓ 文件大小: {OUTPUT_FILE.stat().st_size / 1024:.2f} KB")
            return True
            
        except Exception as e:
            print(f"✗ 下载失败: {str(e)}")
            return False
    
    def download_with_playwright(self):
        """使用 Playwright 浏览器自动化下载
        
        需要安装: pip install playwright
        然后运行: playwright install chromium
        """
        print("📥 使用浏览器自动化方式下载...")
        print("⚠️ 此方式需要额外配置")
        print("\n安装步骤:")
        print("  1. pip install playwright")
        print("  2. playwright install chromium")
        print("\n或者使用手动下载方式更简单")
        return False
    
    def guide_manual_download(self):
        """引导用户手动下载"""
        print("\n" + "=" * 60)
        print("📋 手动下载指引")
        print("=" * 60)
        print(f"\n1. 打开在线文档:")
        print(f"   {ONLINE_DOC_URL}")
        print(f"\n2. 点击右上角「下载」或「导出」按钮")
        print(f"\n3. 选择 Excel (.xlsx) 格式")
        print(f"\n4. 保存到以下路径:")
        print(f"   {OUTPUT_FILE}")
        print(f"\n5. 确保文件名:")
        print(f"   {OUTPUT_FILE.name}")
        print("\n" + "=" * 60)


def main():
    """主函数"""
    print("=" * 60)
    print("    API 对照清单在线下载工具")
    print("=" * 60)
    
    downloader = TencentDocDownloader()
    
    # 检查是否提供了 Cookie
    if len(sys.argv) > 1 and sys.argv[1] == '--cookie':
        if len(sys.argv) < 3:
            print("用法: python download_online_doc.py --cookie '<cookie字符串>'")
            print("\n如何获取 Cookie:")
            print("  1. 在浏览器中打开腾讯文档并登录")
            print("  2. 按 F12 打开开发者工具")
            print("  3. 切换到 Application/Storage 标签")
            print("  4. 复制 Cookies 值")
            sys.exit(1)
        
        cookie = sys.argv[2]
        success = downloader.download_with_cookie(cookie)
        
        if not success:
            print("\n💡 建议改用手动下载方式:")
            downloader.guide_manual_download()
    
    else:
        # 默认显示手动下载指引
        print("\n📥 自动下载腾讯文档需要登录凭证")
        print("\n选择下载方式:\n")
        print("  1. 手动下载（推荐）- 最简单可靠")
        print("  2. Cookie 下载 - 需要提供浏览器Cookie")
        print("  3. 浏览器自动化 - 需要安装Playwright")
        print()
        
        downloader.guide_manual_download()
        
        print("\n💡 提示:")
        print("  - 手动下载只需 10 秒钟")
        print("  - 下载后无需频繁操作，定期更新即可")
        print("  - 自动化方案维护成本高，不推荐")


if __name__ == '__main__':
    main()
