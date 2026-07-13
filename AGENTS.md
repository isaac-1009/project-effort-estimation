# AGENTS.md

本文件为 AI 编码助手（Codex / ZCode 等）在此代码库中工作时提供指导。

## 项目概述

这是一个基于 AI 编码助手 **rules 和 skills 系统**的项目工时估算工具。它通过分析需求和参考历史数据来估算软件项目工时（单位：人天）。

**核心概念**：这不是一个传统的软件项目，而是一个使用结构化 rules 和 skills 提供智能工时估算的 AI 助手。

## 架构

### 核心组件

**Skill**（在用户全局配置中定义，不在此仓库中）：

- `~/.Codex/skills/project-estimation/SKILL.md` - 估算工作流、执行流程
- `~/.Codex/skills/project-estimation/references/effort-standards.md` - 标准工时表、功能模式、Excel 格式规范

> 注：实际路径取决于本机配置（如 `%USERPROFILE%\.Codex\skills\...`）。仓库内不包含 skill 本体，只含数据与辅助脚本。

**数据结构**：

- `data/current/` - 当前估算工作区（每个项目完成后清空）
- `data/estimations/` - 已归档的历史项目
- `data/estimations/index.json` - 项目索引和功能关键词索引
- `data/history.json` - 功能模式数据库及标准工时范围

**脚本**（Python 辅助工具）：

- `scripts/metadata_schema.py` - **共享**：metadata.json schema 校验（单一真相源，被多个脚本 import）
- `scripts/history.py` - **共享**：history.json 自学习写入逻辑（归档时回填功能模式库）
- `scripts/logging_setup.py` - **共享**：统一日志配置（UTF-8 + [LEVEL] 前缀）
- `scripts/generate_excel.py` - 从 `metadata.json` 生成 Excel 文件（**8 列布局**，支持 sub_systems 多端结构）
- `scripts/read_excel.py` - 读取工时估算 Excel 文件，显示详细工时信息
- `scripts/archive_estimation.py` - 将当前估算归档到历史（同时更新 index.json + history.json）
- `scripts/generate_dev_plan.py` - 生成开发计划甘特图 HTML

> 已废弃：旧版 6 列 Excel 布局（#/系统/模块/功能/工时/描述）已被 8 列布局取代；
> `scripts/migrate_history.py` 为一次性迁移脚本，已完成使命并删除。

## 工作流程

详细的估算工作流、工时标准、质量检查清单等均定义在 `project-estimation` skill 中。
当用户提供需求时，Codex 会自动触发 skill 执行估算。

简要流程：

1. 查找历史参考 → 2. 需求分析 → 3. 工时估算 → 4. 生成文件并验证 → 5. 用户确认与归档

**第4步详细说明**（生成文件并验证）：
   - 运行 `generate_excel.py` 生成 `{项目名}估算.xlsx`（8 列布局）
   - 运行 `read_excel.py` 读取并核对工时合计是否与 `metadata.json` 一致
   - 如不一致，检查并修正 `metadata.json` 后重新生成

### 文件管理

**当前估算**（`data/current/`）：

- `metadata.json` - 项目元数据（必需）
- `{项目名}估算.xlsx` - 生成的 Excel 文件（由 `generate_excel.py` 产出）
- `analysis.md` - 项目分析总结
- `requirements.md` - 原始需求
- `pending_questions.md` - 待确认问题（如有）

**确认后**：

- 运行 `python scripts/archive_estimation.py`（支持 `--force` 跳过交互确认）
- 脚本将 `data/current/` 移动到 `data/estimations/{项目名}_{YYYYMMDD}/`
- `project_id` 统一用归档日期生成（格式 `{name}_YYYYMMDD`），与 folder 同源，避免年份漂移
- 更新 `data/estimations/index.json`（项目索引 + 功能关键词索引，关键词数据驱动）
- 更新 `data/history.json`（自学习：把新 feature 并入 projects，命中的功能模式回填 examples）
- 清空 `data/current/` 准备下一个项目
- 若目标 folder 已存在，旧目录会备份为 `{folder}.bak.{时间戳}` 而非直接删除

**切勿**直接修改 `data/estimations/` 中的历史估算文件。

## Excel 格式

当前使用 **8 列布局**（由 `generate_excel.py` 自动生成），格式规范详见 skill 的 `references/effort-standards.md`。

### 列布局

| 列 | 表头 | 合并规则 | 对齐 |
|----|------|---------|------|
| A | # | 按 (module, name) 分组合并 | 居中+上对齐 |
| B | 模块 | 按连续相同值合并 | 左+上对齐 |
| C | 功能 | 按 name 分组合并 | 左+上对齐 |
| D | 需求类型 | 按连续相同值合并 | 左+上对齐 |
| E | 描述 | 不合并（每行独立） | 左+上对齐 |
| F | 系统 | 不合并（企微/PC后台/-） | 左+上对齐 |
| G | 工时(人天) | 不合并 | **右+上对齐** |
| H | 备注 | 不合并 | 左+上对齐 |

### metadata.json 结构（新）

支持 `sub_systems` 数组，一个功能可跨企微+PC后台多行：

```json
{
  "module": "功能开发",
  "name": "评论支持图片",
  "requirement_type": "功能优化",
  "description": "推文评论支持上传图片",
  "sub_systems": [
    {"system": "企微", "effort_days": 4, "remark": "上传9张图片"},
    {"system": "PC后台", "effort_days": 1, "remark": "评论管理显示缩略图"}
  ]
}
```

可选字段：`requirement_type`（需求类型）、`remark`（备注）。
旧的单 system 字段结构仍向后兼容。

### 公式

- 功能测试行：`=ROUNDUP(SUM(开发区间)*0.15,0)`（动态计算）
- 合计行：`=ROUNDUP(SUM(所有行),0)`，F 列标"合计"

## 开发约定

修改脚本时请遵守以下约定（详见 `README.md` 的项目结构）：

**模块复用**：
- `metadata.json` 的 schema 校验统一走 `scripts/metadata_schema.py`，**不要**在其它脚本里复制校验逻辑
- `history.json` 的自学习写入统一走 `scripts/history.py` 的 `update_history()`
- 日志统一用 `scripts/logging_setup.py` 的 `setup_logging()`，**不要**用 `print("[INFO]...")` 手拼前缀

**编码规范**：
- 脚本入口用 `argparse`，库函数内**不要**调 `sys.exit()`（抛异常由 `__main__` 决定退出码）
- magic number（阈值、颜色、列宽等）提取为模块顶部命名常量
- 修改 `generate_excel.py` 后须保证 Excel 产出不变（可用 `read_excel.py` 核对工时合计）

**测试**：
- 纯函数（校验、四舍五入、模式匹配等）须有 `tests/` 下的单元测试覆盖
- 改动后运行 `python -m pytest tests/` 确保全绿

**数据安全**：
- 切勿直接修改 `data/estimations/` 中的历史估算文件
- 归档时若目标目录已存在，会自动备份为 `.bak.{时间戳}`，不会静默覆盖

## 常用命令

```bash
# 从 metadata 生成 Excel
python scripts/generate_excel.py data/current/metadata.json

# 读取工时估算Excel
python scripts/read_excel.py <excel文件路径>

# 归档当前估算到历史
python scripts/archive_estimation.py

# 生成开发计划甘特图（需要 config.json 配置文件）
python scripts/generate_dev_plan.py data/current/dev-plan-config.json

# 查看项目索引
cat data/estimations/index.json

# 查看功能模式
cat data/history.json

# 安装依赖（首次）
pip install -r requirements.txt

# 运行单元测试
python -m pytest tests/
```

### 开发计划甘特图

使用 `generate_dev_plan.py` 生成可视化的开发排期甘特图：

```bash
python scripts/generate_dev_plan.py <config.json> [output.html]
```

**配置结构**：
```json
{
    "title": "项目名称",
    "total_weeks": 11,
    "background": "transparent",
    "month_labels": ["第1月", "第2月", "第3月"],
    "legend": [
        {"label": "需求阶段", "color": "primary"},
        {"label": "UI 设计", "color": "violet"},
        {"label": "系统设计", "color": "sky"},
        {"label": "功能开发", "color": "secondary"},
        {"label": "系统测试", "color": "test"},
        {"label": "客户验收", "color": "accent"},
        {"label": "系统发布", "color": "rose"}
    ],
    "rows": [
        {"label": "需求沟通与评审", "legend": "需求阶段", "start_week": 1, "end_week": 3},
        {"label": "UI设计与评审", "legend": "UI 设计", "start_week": 1, "end_week": 4},
        {"label": "系统设计", "legend": "系统设计", "start_week": 1, "end_week": 4},
        {"label": "功能开发", "legend": "功能开发", "start_week": 2, "end_week": 8},
        {"label": "系统测试", "legend": "系统测试", "start_week": 7, "end_week": 9},
        {"label": "客户验收", "legend": "客户验收", "start_week": 9, "end_week": 11},
        {"label": "系统发布", "legend": "系统发布", "start_week": 11, "end_week": 11, "is_release": true}
    ]
}
```

**字段说明**：

| 字段 | 必需 | 说明 |
|------|------|------|
| `title` | ✓ | 项目标题 |
| `total_weeks` | ✓ | 总周数 |
| `background` | - | 背景：`"transparent"` 或颜色值（默认 `#f5f5f5`） |
| `month_labels` | - | 自定义月份标签（默认自动生成"第X月"） |
| `legend` | - | 图例定义（省略时从 rows 自动生成） |
| `rows` | ✓ | 甘特图行数据 |

**legend 条目**：
- `label`: 显示名称
- `color`: 颜色键名（见下方颜色选项）

**row 条目**：
- `label`: 行标签文本
- `legend`: 引用 legend 中的 label（获取颜色）
- `start_week` / `end_week`: 起止周（1-based，闭区间）
- `is_release`: `true` 时显示 🚀 图标
- `step_num`: 自定义步骤编号（如 `"⑥"`，省略时自动编号 ①②③...）
- `color`: 直接指定颜色键（向后兼容，不推荐）

**颜色选项**：`primary`, `secondary`, `accent`, `test`, `violet`, `rose`, `sky`

**输出**：HTML 文件包含甘特图、图例、PNG 导出按钮

## 估算约束

### 固定工时规则

**环境搭建工时**：
- **新项目（独立系统）**：需要环境搭建工时（2人天）
- **现有系统新增功能**：**不需要环境搭建工时**
  - 判断标准：项目复用已有系统的技术栈、部署环境、基础设施
  - 示例：在"美外之家"企微应用内新增Helpdesk功能 → 无环境搭建工时

**需求沟通工时**：
- **1人天**：需求文档完整（功能列表+交互稿+接口定义齐全）
- **2人天**：需求文档部分缺失（有功能描述但缺少交互稿或接口定义之一）
- **3人天**：仅有口头描述或需求模糊

**UI/UX设计工时**：
- **0人天**：复用现有系统样式，无需新界面设计
- **2人天**：新界面但沿用设计规范（有组件库）
- **3-5人天**：全新视觉风格/品牌设计
  - 页面数量 ≤5：3人天
  - 页面数量 6-15：4人天
  - 页面数量 >15：5人天

**验收发布工时**：
- 标准工时：1人天（固定）

### 工时计算公式

```
总工时 = 开发工时 + 测试工时 + UI/UX设计工时 + 需求沟通 + 验收发布

新项目固定工时 = 需求沟通(1-3) + 环境搭建(2) + 验收发布(1) = 4-8人天
现有系统固定工时 = 需求沟通(1-3) + 验收发布(1) = 2-5人天
```

## 备注

- 此工具通过归档积累历史数据以提高估算准确度（自学习机制：归档时
  `archive_estimation.py` 把新 feature 并入 `history.json`，命中的功能模式
  回填 `examples`）
- 当前库状态：11 个项目、223 个 feature、24 种功能模式均有历史实例
- 准确度随项目数提升：1-5 个 ±30%、5-10 个 ±20%、10-20 个 ±15%
- 始终优先参考历史数据而非标准工时表
- 有疑问时，主动询问用户以明确需求
