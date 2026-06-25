# 历史估算存储系统

## 📁 文件夹结构

```
data/
├── estimations/                    # 历史估算文件夹
│   ├── index.json                 # 索引文件（所有项目汇总）
│   ├── CSR活动管理系统_20250320/   # 项目文件夹（项目名_日期）
│   │   ├── estimate.xlsx          # 估算Excel文件
│   │   ├── metadata.json          # 项目元数据
│   │   ├── requirements.md        # 需求文档
│   │   └── pending_questions.md   # 待确认问题（如果有）
│   └── 新项目名称_20250326/       # 新项目...
│       └── ...
├── history.json                    # 功能历史数据库
└── README.md                       # 本文件
```

## 🔄 累积流程

### 当完成一个新项目估算时：

1. **创建项目文件夹**
   ```
   格式：{项目名称}_{YYYYMMDD}
   示例：在线商城_20250326
   ```

2. **保存文件**
   - `estimate.xlsx` - Excel 估算文件
   - `metadata.json` - 项目元数据
   - `requirements.md` - 原始需求
   - `pending_questions.md` - 待确认问题（如果有）

3. **更新索引** (`index.json`)
   - 添加项目信息
   - 更新功能索引
   - 更新统计信息

4. **更新功能数据库** (`history.json`)
   - 添加新功能模式
   - 更新工时数据

## 🔍 如何查找历史项目

### 按项目名称查找

```python
def find_project_by_name(name):
    with open('data/estimations/index.json', 'r', encoding='utf-8') as f:
        index = json.load(f)

    for project in index['projects']:
        if name in project['name']:
            return project
    return None
```

### 按功能查找

```python
def find_projects_by_feature(feature_name):
    with open('data/estimations/index.json', 'r', encoding='utf-8') as f:
        index = json.load(f)

    # 在功能索引中查找
    project_ids = index['feature_index'].get(feature_name, [])

    # 返回项目详情
    return [p for p in index['projects'] if p['id'] in project_ids]
```

### 按标签查找

```python
def find_projects_by_tag(tag):
    with open('data/estimations/index.json', 'r', encoding='utf-8') as f:
        index = json.load(f)

    return [p for p in index['projects'] if tag in p.get('tags', [])]
```

## 📊 索引文件结构

### index.json

```json
{
  "last_updated": "2025-03-26",
  "total_projects": 2,
  "projects": [
    {
      "id": "CSR_20250320",
      "name": "CSR活动管理系统",
      "date": "2025-03-20",
      "folder": "CSR活动管理系统_20250320",
      "total_effort": 99,
      "systems": ["员工端", "管理后台"],
      "tech_stack": {
        "frontend": "React",
        "backend": "Spring Boot",
        "database": "MySQL"
      },
      "features_count": 17,
      "status": "completed",
      "tags": ["活动管理", "排行榜", "打卡", "勋章", "认证"]
    }
  ],
  "feature_index": {
    "认证": ["CSR_20250320"],
    "排行榜": ["CSR_20250320"],
    "打卡": ["CSR_20250320"],
    ...
  }
}
```

### metadata.json

```json
{
  "project_id": "CSR_20250320",
  "name": "CSR活动管理系统",
  "created_date": "2025-03-20",
  "total_effort": 99,
  "systems": ["员工端", "管理后台"],
  "features": [
    {
      "system": "员工端",
      "module": "认证",
      "name": "登录/注册",
      "effort_days": 3,
      "description": "邮箱",
      "complexity": "中"
    }
  ],
  "tags": ["活动管理", "排行榜", "打卡", "勋章", "认证"]
}
```

## 🎯 使用场景

### 场景1：估算新项目时查找参考

```python
# 新需求：需要一个排行榜功能
keyword = "排行榜"

# 查找历史
related_projects = find_projects_by_feature(keyword)

# 结果：找到 CSR活动管理系统
# → 查看该项目的排行榜功能工时：6人天
# → 参考："排行榜（个人/区域/工厂），6人天"
```

### 场景2：统计分析

```python
# 统计所有包含"认证"功能的项目
projects = find_projects_by_feature("认证")

# 分析工时范围
efforts = [extract_effort(p, "认证") for p in projects]
min_effort = min(efforts)  # 最小值
max_effort = max(efforts)  # 最大值
avg_effort = sum(efforts) / len(efforts)  # 平均值

# 结果：认证功能平均 3-4 人天
```

### 场景3：相似项目推荐

```python
# 新项目：员工激励系统
# 分析技术栈和功能类型

# 查找相似项目
similar = find_similar_projects(
    tech_stack={"frontend": "React", "backend": "Spring Boot"},
    features=["排行榜", "勋章", "打卡"]
)

# 推荐：CSR活动管理系统相似度 85%
# → 可以重点参考该项目的工时分布
```

## 📈 数据增长

随着项目累积：

```
第1个项目：1个参考点
第2个项目：2个参考点
第3个项目：3个参考点
...
第10个项目：10个参考点
```

**累积效应：**
- 功能相似度匹配越来越准确
- 工时估算越来越精准
- 可以分析不同类型项目的工时分布
- 可以识别工时估算的偏差模式

## ⚙️ 自动化脚本

可以使用以下脚本自动维护历史系统：

```python
# scripts/add_estimation_to_history.py
def add_estimation_to_history(project_name, estimate_path):
    # 1. 创建项目文件夹
    folder = create_project_folder(project_name)

    # 2. 复制文件
    copy_files(estimate_path, folder)

    # 3. 生成 metadata.json
    metadata = extract_metadata(estimate_path)
    save_metadata(folder, metadata)

    # 4. 更新 index.json
    update_index(metadata)

    # 5. 更新功能数据库
    update_feature_database(metadata)

    print(f"✅ 项目已添加到历史：{project_name}")
```

## 🔒 数据完整性

每次添加新项目时检查：

- [ ] 项目ID唯一
- [ ] 文件完整（Excel, metadata, requirements）
- [ ] 索引正确更新
- [ ] 功能索引正确
- [ ] 数据格式正确

## 📊 统计分析

可以基于历史数据生成：

1. **功能工时分布图**
2. **技术栈偏好统计**
3. **项目规模趋势**
4. **工时估算准确度分析**

---

**系统维护**：每次完成新项目估算后，自动添加到此历史系统
