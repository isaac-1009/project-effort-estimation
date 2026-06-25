# 项目估算脚本使用说明

## 📁 文件结构

```
project-effort-estimation/
├── data/
│   ├── current/                    # 当前估算项目（工作区）
│   │   ├── metadata.json          # 项目元数据
│   │   ├── estimate.xlsx          # Excel估算文件
│   │   ├── requirements.md        # 原始需求
│   │   ├── pending_questions.md   # 待确认问题
│   │   └── analysis.md           # 项目分析总结
│   ├── estimations/                # 历史估算项目（已归档）
│   │   ├── index.json            # 项目索引
│   │   ├── CSR活动管理系统_20250320/
│   │   └── ...
│   └── history.json              # 功能历史数据库
├── scripts/
│   ├── generate_excel.py          # Excel生成脚本
│   └── archive_estimation.py      # 归档脚本
└── templates/
    └── analysis_template.md       # 分析总结模板
```

## 🔄 工作流程

### 1. 创建新估算

当接到新需求时，在 `data/current/` 文件夹中创建以下文件：

```bash
# 1. 创建 metadata.json（必需）
# 参考：data/current/metadata.json

# 2. 生成 Excel 文件
python scripts/generate_excel.py data/current/metadata.json

# 3. 创建其他文件（可选）
# - requirements.md：原始需求
# - pending_questions.md：待确认问题
# - analysis.md：项目分析总结
```

### 2. 审查和修改

在 `data/current/` 中审查生成的文件：
- 检查 Excel 文件格式
- 修改工时估算
- 补充待确认问题
- 完善项目分析

### 3. 归档到历史

当估算确认无误后：

```bash
# 归档当前估算到历史文件夹
python scripts/archive_estimation.py

# 此操作会：
# 1. 将 data/current/ 移动到 data/estimations/{项目名}_{日期}/
# 2. 更新 data/estimations/index.json
# 3. 清空 data/current/ 准备下次使用
```

## 📝 metadata.json 格式

```json
{
  "project_id": "项目ID_日期",
  "name": "项目名称",
  "created_date": "2025-03-26",
  "status": "estimated",
  "total_effort": 67,
  "breakdown": {
    "development": 47,
    "testing": 7,
    "other": 13
  },
  "systems": ["系统1", "系统2"],
  "tech_stack": {
    "frontend": "前端技术",
    "backend": "后端技术"
  },
  "features": [
    {
      "system": "系统名称",
      "module": "模块名称",
      "name": "功能名称",
      "effort_days": 5,
      "description": "功能描述",
      "complexity": "中"
    }
  ],
  "tags": ["标签1", "标签2"],
  "notes": "备注"
}
```

## 🎯 Excel 生成特性

- ✅ 只有一个 Sheet，名称为"工时估算"
- ✅ 从 B 列（第2列）开始记录数据
- ✅ 相同系统/模块自动合并单元格
- ✅ 自动添加边框和样式
- ✅ 冻结首行
- ✅ 自动计算测试工时（15%）和总工时

## 📚 查找历史项目

### 按项目名称查找

查看 `data/estimations/index.json` 中的 `projects` 数组。

### 按功能查找

查看 `data/estimations/index.json` 中的 `feature_index` 对象。

### 示例

```bash
# 查找包含"认证"功能的项目
grep "认证" data/estimations/index.json
```

## 🔧 常见问题

### Q: 如何修改Excel样式？
A: 编辑 `scripts/generate_excel.py` 中的样式定义部分。

### Q: 如何添加新的功能模式？
A: 编辑 `data/history.json` 中的 `feature_patterns` 对象。

### Q: 归档后如何恢复？
A: 归档后的文件保存在 `data/estimations/` 文件夹中，可以手动查看。

### Q: 如何重新估算已有项目？
A: 在 `data/current/` 中创建新的 `metadata.json`，基于历史项目参考。
