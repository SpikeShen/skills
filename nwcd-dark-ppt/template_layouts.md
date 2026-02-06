# NWCD Dark 模板版式详解

## 版式索引

模板共提供 16 种版式（索引 0-15），每种版式都有特定的用途和占位符配置。

## 详细版式说明

### [0] Title Slide - 标题幻灯片
**用途**: 演示文稿封面
**占位符**: 4 个
- 演讲者/团队名称 (位置: 1.4", 6.2")
- 日期 (位置: 1.4", 6.7")
- 主标题 (位置: 1.4", 2.7")
- 副标题 (位置: 1.4", 4.1")

**使用场景**: 
- 演示文稿第一页
- 章节封面

**代码示例**:
```python
slide = prs.slides.add_slide(prs.slide_layouts[0])
title = slide.placeholders[12]  # 主标题
subtitle = slide.placeholders[13]  # 副标题
presenter = slide.placeholders[10]  # 演讲者
date = slide.placeholders[11]  # 日期

title.text = "Kiro MCP 架构自动化"
subtitle.text = "基于 Model Context Protocol 的智能架构生成"
presenter.text = "技术团队"
date.text = "2026年2月"
```

---

### [1] Content List - 内容列表
**用途**: 目录页或内容概览
**占位符**: 0 个（使用预定义形状）
**总形状数**: 3

**使用场景**:
- 演示文稿目录
- 章节概览

**注意**: 此版式使用固定形状，不建议修改

---

### [2] Summary - 总结页
**用途**: 章节总结或要点概括
**占位符**: 2 个
- 标题 (位置: 2.0", 1.8")
- 正文内容 (位置: 2.0", 3.0")

**使用场景**:
- 章节总结
- 关键要点
- 简短说明

**代码示例**:
```python
slide = prs.slides.add_slide(prs.slide_layouts[2])
title = slide.placeholders[0]
body = slide.placeholders[10]

title.text = "Summary"
body.text = "Amazon ECS 和 Amazon EKS 提供了灵活的容器编排方案..."
```

---

### [3] Chapter Topics - 章节主题
**用途**: 章节分隔页
**占位符**: 2 个
- 章节标题 (位置: 5.1", 2.9")
- 副标题 (位置: 5.1", 4.2")

**使用场景**:
- 新章节开始
- 主题切换

**代码示例**:
```python
slide = prs.slides.add_slide(prs.slide_layouts[3])
chapter = slide.placeholders[10]
subheading = slide.placeholders[11]

chapter.text = "技术架构"
subheading.text = "MCP 协议与 Draw.io 集成"
```

---

### [4] Blank Page - 空白页
**用途**: 自定义内容
**占位符**: 3 个
- 日期 (位置: 0.9", 0.3")
- 页脚 (位置: 2.7", 6.8")
- 页码 (位置: 11.4", 0.3")

**使用场景**:
- 完全自定义布局
- 插入大图
- 特殊设计

---

### [5] Only Title - 仅标题
**用途**: 只有标题的页面
**占位符**: 1 个
- 标题 (位置: 1.0", 0.3")

**使用场景**:
- 需要大量自定义内容
- 图表展示
- 架构图

**代码示例**:
```python
slide = prs.slides.add_slide(prs.slide_layouts[5])
title = slide.placeholders[0]
title.text = "系统架构图"

# 然后添加自定义形状或图片
```

---

### [6] Content - 标题+内容
**用途**: 标准内容页
**占位符**: 2 个
- 标题 (位置: 1.0", 0.3")
- 内容对象 (位置: 1.0", 1.4")

**使用场景**:
- 单列内容展示
- 文本说明
- 列表展示

**代码示例**:
```python
slide = prs.slides.add_slide(prs.slide_layouts[6])
title = slide.placeholders[0]
content = slide.placeholders[10]

title.text = "工作流程"
tf = content.text_frame
tf.text = "1. 用户输入需求"
p = tf.add_paragraph()
p.text = "2. Kiro 分析架构"
p.level = 0
```

---

### [7] Two Content - 双列内容
**用途**: 并列对比或双栏展示
**占位符**: 3 个
- 标题 (位置: 1.0", 0.3")
- 左侧内容 (位置: 1.0", 1.4")
- 右侧内容 (位置: 6.8", 1.4")

**使用场景**:
- 对比分析
- 优缺点对比
- 并列展示

**代码示例**:
```python
slide = prs.slides.add_slide(prs.slide_layouts[7])
title = slide.placeholders[0]
left = slide.placeholders[10]
right = slide.placeholders[11]

title.text = "使用场景"
left.text = "开发环境\n• 快速原型\n• 本地测试"
right.text = "生产环境\n• 高可用\n• 自动扩展"
```

---

### [8] Three Content - 三列内容 ⭐ 推荐
**用途**: 三栏并列展示
**占位符**: 4 个
- 标题 (位置: 1.0", 0.3")
- 左侧内容 (位置: 1.0", 1.4")
- 中间内容 (位置: 4.9", 1.4")
- 右侧内容 (位置: 8.7", 1.4")

**使用场景**:
- 技术栈展示
- 三要素说明
- 多维度对比

**代码示例**:
```python
slide = prs.slides.add_slide(prs.slide_layouts[8])
title = slide.placeholders[0]
left = slide.placeholders[10]
middle = slide.placeholders[11]
right = slide.placeholders[12]

title.text = "技术栈"
left.text = "前端\n• React\n• TypeScript"
middle.text = "后端\n• Python\n• FastAPI"
right.text = "基础设施\n• AWS\n• Docker"
```

---

### [9] Master Title - 主标题页
**用途**: 大标题展示
**占位符**: 2 个
- 主标题 (位置: 0.9", 2.6")
- 副标题 (位置: 0.9", 3.9")

**使用场景**:
- 重要声明
- 核心观点
- 章节开始

---

### [10] Content Title - 内容标题页
**用途**: 带详细说明的标题页
**占位符**: 3 个
- 内容标题 (位置: 1.0", 1.7")
- 副标题 (位置: 1.0", 2.3")
- 详细说明 (位置: 1.0", 3.0")

**使用场景**:
- 详细介绍
- 多层次说明

---

### [11] Goals - 目标/要点（四宫格）
**用途**: 四个要点展示
**占位符**: 5 个
- 标题 (位置: 1.0", 0.3")
- 要点1 (位置: 1.1", 1.4")
- 要点2 (位置: 1.1", 4.2")
- 要点3 (位置: 7.2", 1.4")
- 要点4 (位置: 7.2", 4.2")

**使用场景**:
- 四大优势
- 核心功能
- 关键目标

**代码示例**:
```python
slide = prs.slides.add_slide(prs.slide_layouts[11])
title = slide.placeholders[0]
goal1 = slide.placeholders[18]
goal2 = slide.placeholders[19]
goal3 = slide.placeholders[20]
goal4 = slide.placeholders[21]

title.text = "核心优势"
goal1.text = "1. 自动化"
goal2.text = "2. 智能化"
goal3.text = "3. 可扩展"
goal4.text = "4. 易集成"
```

---

### [12] Process - 流程页
**用途**: 流程展示
**占位符**: 1 个
- 标题 (位置: 1.0", 0.3")

**使用场景**:
- 工作流程
- 步骤说明

---

### [13] Analysis - 分析页（带表格）
**用途**: 数据分析或对比
**占位符**: 2 个
- 标题 (位置: 1.0", 0.3")
- 表格 (位置: 1.0", 1.2")

**使用场景**:
- 竞品分析
- 数据对比
- 特性比较

---

### [14] Positioning Map - 定位图
**用途**: 定位或矩阵展示
**占位符**: 1 个
- 标题 (位置: 1.0", 0.3")

**使用场景**:
- 市场定位
- 技术选型矩阵

---

### [15] Thanks - 结束页 ⭐ 推荐
**用途**: 演示文稿结束
**占位符**: 2 个
- 感谢标题 (位置: 0.7", 2.3")
- 联系信息 (位置: 0.7", 3.8")

**使用场景**:
- 演示文稿最后一页
- Q&A 页面

**代码示例**:
```python
slide = prs.slides.add_slide(prs.slide_layouts[15])
title = slide.placeholders[0]
contact = slide.placeholders[10]

title.text = "Thanks！"
contact.text = "联系方式：tech@example.com"
```

---

## 推荐使用组合

### 标准技术演示（6页）
1. [0] Title Slide - 封面
2. [8] Three Content - 技术栈
3. [6] Content - 工作流程
4. [11] Goals - 核心优势
5. [7] Two Content - 使用场景
6. [15] Thanks - 结束

### 产品介绍（8页）
1. [0] Title Slide - 封面
2. [1] Content List - 目录
3. [3] Chapter Topics - 产品概述
4. [6] Content - 产品特性
5. [11] Goals - 核心价值
6. [13] Analysis - 竞品对比
7. [7] Two Content - 应用场景
8. [15] Thanks - 结束

### 架构说明（5页）
1. [0] Title Slide - 封面
2. [5] Only Title - 架构图
3. [8] Three Content - 组件说明
4. [6] Content - 技术细节
5. [15] Thanks - 结束
