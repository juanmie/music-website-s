# API Document Retrieval 使用示例

## 示例 1: 从 Word 需求文档匹配 API

### 输入文件
`各地区集成需求/上海中柔_OA抛转调拨单接口需求确认书.docx`

### 执行命令
```bash
python .qoder/skills/api-document-retrieval/scripts/match_api.py "各地区集成需求/上海中柔_OA抛转调拨单接口需求确认书.docx"
```

### 提取的关键字
- 业务对象: 调拨单、审批流程
- 关键字段: 单号、状态、日期、申请人
- 涉及系统: OA、T100
- 操作类型: 抛转、创建

### 输出报告
生成 `API匹配报告_20260624.md`，包含：
- 高匹配度：调拨单创建 API (OA → T100)
- 字段映射表：OA字段 ↔ T100字段
- 转换规则：状态码映射、日期格式

---

## 示例 2: 从 Excel 需求清单批量匹配

### 输入文件
`集成需求清单产出/上海中柔_OA抛转调拨单接口需求确认书_标准化.xlsx`

### 执行命令
```bash
python .qoder/skills/api-document-retrieval/scripts/match_api.py "集成需求清单产出/上海中柔_OA抛转调拨单接口需求确认书_标准化.xlsx"
```

### 匹配结果
自动读取 Excel 中的：
- 需求概览表：提取需求名称、客户信息
- 集成接口表：提取接口方向、系统类型
- 字段映射表：提取具体字段列表

### 输出
生成详细的 Markdown 报告，包含多个匹配结果

---

## 示例 3: 从文本文件快速匹配

### 输入文件
创建 `test_requirement.txt`:
```
客户需求：在OA系统中发起采购申请，审批通过后自动在T100中创建采购订单。
需要字段：
- 申请单号
- 申请人
- 申请日期
- 物料编码
- 采购数量
- 单价
- 总金额
- 审批状态
```

### 执行命令
```bash
python .qoder/skills/api-document-retrieval/scripts/match_api.py test_requirement.txt
```

### 提取关键字
- 业务对象: 采购订单、审批流程
- 关键字段: 单号、申请人、日期、金额、状态
- 涉及系统: OA、T100
- 操作类型: 创建、审批、自动

### 预期匹配
- 高匹配度: CreatePurchaseOrder API
- 字段映射: 完整的采购订单字段对照
- 转换规则: 审批状态映射、金额精度

---

## 示例 4: 自定义关键字匹配

### 创建自定义输入
```bash
echo "销售订单同步接口，从T100同步到OA系统，包含订单号、客户、金额、状态" > custom_req.txt
```

### 执行匹配
```bash
python .qoder/skills/api-document-retrieval/scripts/match_api.py custom_req.txt
```

---

## 解读匹配报告

### 报告结构
```markdown
# API 匹配报告

## 📋 需求信息          ← 原始需求概览
## 📊 匹配统计          ← 匹配结果统计
## 🔴 高匹配度 API      ← 最相关的接口（重点查看）
## 🟡 中匹配度 API      ← 可能相关的接口
## 🟢 低匹配度 API      ← 参考性接口
## 💡 建议              ← 改进建议
```

### 高匹配度结果解读
```markdown
### 1. API 接口 (匹配度: 95%)

**参考位置**: 工作表 `采购接口` | 行号 `45-67`
**匹配依据**: 业务对象: 采购订单, 字段: 单号, 系统: OA, 动作: 创建

#### 字段映射
| 源字段 | 目标字段 | 类型 | 转换规则 |
|--------|---------|------|---------|
| docNo | docno | String→varchar(30) | 直接映射 |
| status | status | String→char(1) | Approved→Y |
```

**关键信息**：
- 匹配度 95%：非常相关，优先采用
- 参考位置：可在原对照清单中快速定位
- 匹配依据：说明为什么匹配
- 字段映射：具体的字段对照关系

---

## 提高匹配准确度的技巧

### 1. 丰富需求描述
❌ **差**: "采购接口"  
✅ **好**: "OA审批流程对接T100创建采购订单，包含单号、申请人、物料、金额等字段"

### 2. 明确接口方向
❌ **模糊**: "T100和OA对接"  
✅ **明确**: "OA推送数据到T100" 或 "T100同步数据到OA"

### 3. 列出关键字段
❌ **少**: "需要字段映射"  
✅ **多**: "字段：docno, status, requester, amount, date, remarks"

### 4. 使用标准术语
- 采购订单 (而非 "买东西的单子")
- 审批流程 (而非 "领导签字")
- 抛转 (而非 "传过去")

---

## 批量处理多个需求

### 创建批处理脚本
```bash
#!/bin/bash
# batch_match.sh

for file in 各地区集成需求/*.docx; do
    echo "处理: $file"
    python .qoder/skills/api-document-retrieval/scripts/match_api.py "$file"
    echo "---"
done
```

### 执行批处理
```bash
chmod +x batch_match.sh
./batch_match.sh
```

---

## 常见问题

### Q1: 匹配结果为空怎么办？
**A**: 
- 检查需求文件是否包含有效文本
- 尝试使用更详细的业务描述
- 确认 API 对照清单文件存在

### Q2: 匹配度都很低怎么办？
**A**:
- 补充业务场景描述
- 添加更多字段信息
- 使用标准业务术语

### Q3: 如何查看完整的字段映射？
**A**:
- 报告中会列出主要字段映射
- 根据"参考位置"打开原对照清单查看完整信息
- 工作表和行号帮助快速定位

### Q4: 脚本执行失败？
**A**:
```bash
# 检查依赖
pip install openpyxl python-docx

# 检查文件路径
ls -la 需求对照工具表/

# 查看错误日志
python .qoder/skills/api-document-retrieval/scripts/match_api.py <file> 2>&1 | head -20
```

---

## 进阶用法

### 集成到需求标准化流程
```bash
# 1. 标准化需求文档
node convert_requirements.js

# 2. 匹配 API 对照清单
python .qoder/skills/api-document-retrieval/scripts/match_api.py "集成需求清单产出/xxx_标准化.xlsx"

# 3. 查看报告
cat API匹配报告_*.md
```

### 自动化工作流
```python
# workflow.py
import subprocess
from pathlib import Path

# 处理所有需求文件
for req_file in Path("各地区集成需求").glob("*.docx"):
    subprocess.run([
        'python', 
        '.qoder/skills/api-document-retrieval/scripts/match_api.py',
        str(req_file)
    ])
```

---

## 输出示例

查看完整的输出示例，请参考工作区中生成的 `API匹配报告_*.md` 文件。

每个报告包含：
- ✅ 需求信息概览
- ✅ 匹配度统计
- ✅ 高/中/低匹配结果
- ✅ 字段映射详情
- ✅ 参考位置指引
- ✅ 改进建议
