# 项目工时估算工具

基于 Claude Code 的项目工时估算系统，通过 `project-estimation` skill 实现智能工时估算。

## 整体架构

```
用户需求 → project-estimation skill → data/current/ → 确认 → data/estimations/
           (5步工作流)                  (工作区)        (归档)
```

系统由 3 个层次组成：

| 层级 | 位置 | 职责 |
|------|------|------|
| **Skill** | `~/.claude/skills/project-estimation/` | 工作流定义、工时标准、执行逻辑 |
| **数据** | `project-effort-estimation/data/` | 历史项目、功能模式库、当前工作区 |
| **脚本** | `project-effort-estimation/scripts/` | Excel 生成、项目归档 |

## 操作流程

### 用户视角

```
用户："/project-estimation 请估算以下需求的工时：[需求文本]"
         │
         ▼
┌─────────────────────────────┐
│ Step 1: 查找历史参考          │  ← Read index.json + history.json
│ 匹配相似项目，输出参考工时     │
├─────────────────────────────┤
│ Step 2: 需求分析              │  ← 识别系统、模块、功能
│ 分类 Must-have / 可选         │     如不明确则 AskUserQuestion
├─────────────────────────────┤
│ Step 3: 处理待确认问题        │  ← A/B/C 三级分类
│ 不确定的标记"需求未定"，工时=0  │
├─────────────────────────────┤
│ Step 4: 工时估算              │  ← 匹配24种功能模式
│ 创建 metadata.json            │     运行 generate_excel.py
│ 生成 Excel + analysis.md      │
├─────────────────────────────┤
│ Step 5: 质量检查 → 用户确认    │
│ 确认后运行 archive_estimation.py → 归档
└─────────────────────────────┘
```

### 使用方法

1. **提供需求** — 向 Claude 提供项目需求文本
2. **自动处理** — Claude 执行 5 步工作流，生成文件到 `data/current/`
3. **审查确认** — 查看 `data/current/` 中的 Excel 和分析文档
4. **归档** — 确认无误后运行 `python scripts/archive_estimation.py`

## 工时估算原理

### 1. 功能模式匹配

每个功能匹配到 24 种标准模式之一（定义在 `references/effort-standards.md`）：

| 模式 | 基准工时 | 复杂度范围 |
|------|---------|-----------|
| CRUD操作 | 3人天 | 2-5 |
| 认证授权 | 4人天 | 3-8 |
| 文件上传 | 4人天 | 3-6 |
| 富文本编辑 | 5人天 | 4-8 |
| 审核流程 | 6人天 | 4-8 |
| 数据统计 | 5人天 | 4-7 |
| 排行榜 | 6人天 | 5-8 |
| 打卡签到 | 2人天 | 2-4 |
| 勋章积分 | 7人天 | 6-10 |
| 问答系统 | 6人天 | 5-8 |
| 二维码 | 1人天 | 1-2 |
| 系统管理 | 6人天 | 5-8 |
| 内容管理 | 10人天 | 8-12 |
| 社交互动 | 8人天 | 6-10 |
| 工单系统 | 15人天 | 10-20 |
| ... | ... | ... |

完整 24 种模式见 `C:\Users\admin\.claude\skills\project-estimation\references\effort-standards.md`

### 2. 复杂度系数

| 复杂度 | 系数 | 适用场景 |
|--------|------|---------|
| 低 | 0.8x | 标准CRUD，无特殊逻辑，无第三方集成 |
| 中 | 1.0x | 有业务逻辑，文件处理，简单审核流程 |
| 高 | 1.3x | 复杂规则，多系统集成，高并发 |

### 3. 历史数据优先

估算优先级：**历史参考 > 功能模式 > 标准工时表**

- 每个归档项目自动更新 `index.json`（项目索引）和 `history.json`（功能模式库）
- 归档时 `archive_estimation.py` 会把新项目的每个 feature 与 24 种标准模式
  做关键词匹配，命中的 feature id 回填到 `feature_patterns[].examples`，
  实现"越用越准"的自学习
- 下次估算通过标签、技术栈、功能关键词匹配相似项目
- 匹配到相似功能时直接参考历史工时

### 4. 工时汇总公式

```
开发工时 = Σ(各功能估算工时)
测试工时 = 开发工时 × 15%（默认）
固定工时 = 需求沟通(3) + 环境搭建(2) + 验收发布(1) = 6人天
────────────────────────────────────────
总工时   = 开发工时 + 测试工时 + 固定工时
```

所有工时四舍五入到 0.5 的整数倍（最小估算单位）。

## 数据流转

```
data/
├── history.json                    ← 功能模式库(24种) + 历史项目
├── estimations/
│   ├── index.json                  ← 项目索引 + 功能关键词索引
│   └── {name}_{date}/              ← 已归档项目
│       ├── metadata.json
│       ├── 工时估算.xlsx
│       └── analysis.md
└── current/                        ← 当前工作区（估算完成后清空）
    ├── metadata.json               ← 由 skill 生成
    ├── 工时估算.xlsx               ← 由 generate_excel.py 生成
    └── analysis.md                 ← 由 skill 生成
```

## 项目结构

```
project-effort-estimation/
├── CLAUDE.md                        # Claude Code 项目指南
├── README.md                        # 本文件
├── requirements.txt                 # Python 依赖（xlsxwriter, openpyxl）
├── pyproject.toml                   # pytest/ruff 配置 + 依赖声明
├── data/
│   ├── current/                     # 当前估算工作区
│   ├── estimations/                 # 历史估算（已归档）
│   │   └── index.json              # 项目索引 + 功能关键词索引
│   └── history.json                # 功能模式库(24种) + 历史项目
├── scripts/
│   ├── metadata_schema.py          # 共享：metadata.json schema 校验（单一真相源）
│   ├── history.py                  # 共享：history.json 自学习写入逻辑
│   ├── logging_setup.py            # 共享：统一日志配置
│   ├── generate_excel.py           # 从 metadata.json 生成 Excel 估算
│   ├── read_excel.py               # 读取/校验 Excel 估算文件
│   ├── archive_estimation.py       # 归档当前估算到历史库
│   └── generate_dev_plan.py        # 生成甘特图 HTML
└── tests/                          # pytest 单元测试
```

## Skill 配置

估算工作流定义在以下文件中（全局配置，不在本仓库中）：

| 文件 | 内容 |
|------|------|
| `~/.claude/skills/project-estimation/SKILL.md` | 5步工作流、执行逻辑、质量检查 |
| `~/.claude/skills/project-estimation/references/effort-standards.md` | 24种功能模式、复杂度系数、Excel格式规范 |

## 准确度提升机制

| 项目数量 | 准确度 | 原因 |
|---------|--------|------|
| 1-5 | ±30% | 基准估算，依赖标准工时表 |
| 5-10 | ±20% | 有少量历史参考 |
| 10-20 | ±15% | 功能模式库较完善 |

> 当前库已积累 10 个项目、162 个 feature，22 种功能模式均有历史实例支撑。

## 安装与开发

### 环境要求

- Python 3.9+

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行测试

```bash
pip install pytest        # 首次需要
python -m pytest tests/   # 运行全部单元测试
```

### 版本控制

项目使用 git 管理。`.claude/settings.local.json` 已在 `.gitignore` 中排除
（含本机敏感配置）。Excel/HTML 等二进制产物通过 `.gitattributes` 标记为
binary，文本文件统一 LF 行尾。
