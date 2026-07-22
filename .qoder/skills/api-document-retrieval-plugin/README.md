# API Document Retrieval Plugin

## 插件概述

**API Document Retrieval** 是一个 Qoder 插件，能够根据需求清单中的关键字，自动从 **T100与OA类系统集成标准API对照清单** 中匹配对应的API接口信息和字段映射。

## 功能特性

- **智能关键字提取**：自动识别业务对象、字段名称、系统标识、操作类型
- **多维度API匹配**：高/中/低优先级匹配策略
- **多种匹配版本**：
  - V1 - 关键字匹配（文本/docx/txt）
  - V2 - 智能匹配增强版（Excel）
  - V5 - 能率需求清单智能匹配（Node.js + ExcelJS，完整保留格式）
  - V6 - 统一模板智能匹配（最新推荐版本）
- **格式完整保留**：使用 ExcelJS 完整保留边框、颜色、字体、合并单元格、列宽、行高
- **自动推断流向**：根据开发分类自动推断接口流向
- **字段对照页签复制**：匹配成功后自动复制对应的字段对照工作表到输出文件

## 安装方法

### 方式 1：直接复制到项目（推荐）

1. 将本插件文件夹复制到目标项目的 `.qoder/skills/` 目录下：
   ```
   <你的项目>/.qoder/skills/api-document-retrieval/
   ```

2. 确保依赖已安装：
   ```bash
   # Python 依赖（V1/V2）
   pip install openpyxl python-docx

   # Node.js 依赖（V5/V6）
   npm install exceljs xlsx
   ```

3. 准备数据源文件（放到项目根目录的 `需求对照工具表/` 目录下）：
   - `【2026】字段对照---最小颗粒度.xlsx`（字段对照清单）
   - `【2026】接口清单---汇总.xlsx`（接口清单汇总）

### 方式 2：通过 Qoder 插件管理器安装

（如果 Qoder 支持插件管理器，可将本插件打包为 `.zip` 后安装）

## 文件结构

```
api-document-retrieval-plugin/
  .qoder-plugin/
    plugin.json          # 插件清单
  README.md              # 本文件
  skills/
    api-document-retrieval/
      SKILL.md           # 技能主文件（Qoder 使用说明）
      README.md          # 技能详细文档
      examples.md        # 使用示例
      V2智能匹配使用说明.md   # V2版本详细说明
      scripts/
        match_api.py           # V1 - 关键字匹配
        match_api_v2.py        # V2 - 智能匹配增强版
        match_nenglv_v5.js     # V5 - 能率需求清单智能匹配（ExcelJS格式保留）
        match_nenglv_v6.js     # V6 - 统一模板智能匹配（最新推荐）
        download_online_doc.py # 在线文档下载工具
        match_nenglv_bpm.py    # BPM专用匹配（旧版）
        match_nenglv_bpm_v2.py # BPM专用匹配V2（旧版）
        match_nenglv_bpm_v3.py # BPM专用匹配V3（旧版）
        match_nenglv_v4.js     # V4版本（旧版，不推荐使用）
        match_nenglv_v4.py     # V4 Python版本（旧版）
```

## 使用方式

### V1 - 关键字匹配
```bash
python .qoder/skills/api-document-retrieval/scripts/match_api.py <需求清单文件路径>
```

### V2 - 智能匹配增强版
```bash
python .qoder/skills/api-document-retrieval/scripts/match_api_v2.py <需求清单Excel文件>
```

### V5 - 能率需求清单智能匹配（推荐）
```bash
node .qoder/skills/api-document-retrieval/scripts/match_nenglv_v5.js <需求清单xlsx文件>
```

### V6 - 统一模板智能匹配（最新推荐）
```bash
node .qoder/skills/api-document-retrieval/scripts/match_nenglv_v6.js <需求清单xlsx文件>
```

## 数据源要求

本插件需要以下本地数据源文件（放在项目根目录的 `需求对照工具表/` 目录下）：

1. **字段对照清单**：`【2026】字段对照---最小颗粒度.xlsx`
   - 每个工作表对应一个流程的详细字段对照
   - 包含接收方/发起方字段、参数类型、赋值逻辑等

2. **接口清单汇总**：`【2026】接口清单---汇总.xlsx`
   - 所有接口的汇总列表
   - 包含流程名称、接口方向、功能描述、适配类型等

如果本地文件不存在，脚本会提示访问在线文档手动下载：
- 接口清单汇总：https://docs.qq.com/sheet/DWmF4RXZmdW5RWnVk?tab=nxp5uc
- 字段对照清单：https://docs.qq.com/sheet/DWm9PT0xtYWdYRW5R?scene=a815f6149367bdcef5a2dc6d3cXfw1&tab=mcy037

## 迁移注意事项

1. **路径依赖**：脚本中的路径是基于工作区根目录的相对路径（`需求对照工具表/`），迁移后需要确保目录结构一致
2. **依赖安装**：目标机器需要安装 Python（V1/V2）或 Node.js（V5/V6）以及对应的库
3. **数据源文件**：需要手动复制或重新下载 `需求对照工具表/` 目录下的 Excel 文件
4. **在线文档**：如果本地文件缺失，脚本会提示访问腾讯文档在线版本

## 版本历史

- **v2.0.0** (2026-07-10) - 打包为 Qoder 插件，支持迁移
- **v1.0.0** (2026-06-24) - 初始版本发布

## 技术支持

如有问题，请查看：
- `skills/api-document-retrieval/SKILL.md` - 技能详细说明
- `skills/api-document-retrieval/examples.md` - 使用示例
- `skills/api-document-retrieval/V2智能匹配使用说明.md` - V2版本详细说明

