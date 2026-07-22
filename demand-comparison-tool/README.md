# demand-comparison-tool（需求对照工具表）

本目录用于存放 T100/OA 集成标准 API 对照清单文件。

## 所需文件

请将以下两个 Excel 文件放置在此目录：

### 1. 接口清单汇总
- **文件名**: `【2026】接口清单---汇总.xlsx`
- **内容**: 所有接口的汇总列表，包含流程名称、接口方向、功能描述、适配类型等
- **结构**: 按系统分类（OA接口清单、WMS接口清单、MES接口清单）

### 2. 字段对照清单
- **文件名**: `【2026】字段对照---最小颗粒度.xlsx`
- **内容**: 每个工作表对应一个流程的详细字段对照（接收方/发起方字段、参数类型、赋值逻辑等）
- **结构**: 工作表名包含程序代号（如 axmt540、aimm200）

## 文件获取方式

### 方式 1: 从在线文档下载（推荐）
1. 访问接口清单汇总在线文档：https://docs.qq.com/sheet/DWmF4RXZmdW5RWnVk?tab=nxp5uc
2. 访问字段对照清单在线文档：https://docs.qq.com/sheet/DWm9PT0xtYWdYRW5R?scene=a815f6149367bdcef5a2dc6d3cXfw1&tab=mcy037
3. 点击右上角「文件」→「导出为」→「Excel(.xlsx)」
4. 保存文件到本目录，确保文件名与上述一致

### 方式 2: 使用下载脚本
```bash
# 需要先安装依赖
pip install playwright

# 运行下载脚本
python .qoder/skills/api-document-retrieval-plugin/skills/api-document-retrieval/scripts/download_online_doc.py
```

## 配置说明

后端配置文件 (`music-server/src/main/resources/application-dev.properties`) 中的相关配置：

```properties
# 需求对照工具表路径
api.match.data-dir=demand-comparison-tool
api.match.interface-list-name=【2026】接口清单---汇总.xlsx
api.match.field-mapping-name=【2026】字段对照---最小颗粒度.xlsx

# AI 输出物模板路径（位于上级目录 integration-field-output/）
api.match.output-template-dir=integration-field-output
api.match.output-template-name=AI输出物摸板.xlsx
```

## 输出模板

AI 输出物模板文件位于 `integration-field-output/AI输出物摸板.xlsx`，包含以下标准工作表：
- 集成配置
- 开发计划
- 问题跟进管制-TB导入联动

导出时会自动复制这些工作表到输出文件中。

## 使用场景

- **AI 智能分析导出**: 结合 Kimi AI 分析需求，并匹配 API 接口
- **API 对照匹配导出**: 基于本地规则匹配 API 对照清单，生成匹配报告

## 注意事项

1. 文件名必须与配置一致，否则无法读取
2. 文件更新后，后端服务会自动重新加载（5分钟缓存）
3. 建议定期从在线文档同步最新版本
