---
name: NWCD Dark PPT Generation
version: 3.0.0
description: 生成符合 NWCD Dark 模板风格的 PowerPoint 演示文稿。融合了前端设计美学、PPTX 技术能力和演示适配性优化，提供从模板分析、视觉设计到内容填充的完整工作流。
---

# NWCD Dark PPT Generation Skill (v3.0)

融合了四个能力维度：
1. **NWCD 模板知识** — 版式结构、占位符映射、品牌色彩
2. **前端设计美学** — 视觉层次、色彩理论、排版节奏、装饰元素
3. **PPTX 技术能力** — python-pptx / PptxGenJS / XML 编辑工作流
4. **演示适配性** — 投影仪/投屏场景优化、高对比度、内容溢出防护

## 前置环境检查（生成前必须执行）

在生成 PPT 之前，agent 必须先执行以下检查：

1. **检查 python-pptx 是否已安装**：运行 `python3 -c "import pptx; print(pptx.__version__)"`
2. **如果未安装**：安装依赖 `pip3 install python-pptx`（或使用 references 目录中的 `requirements.txt`）
3. **确认模板文件存在**：模板文件 `Deck_Template_NWCD_dark_202103.pptx` 位于 skill 的 `references/` 目录中，使用前需确认其路径

## 模板信息

- **模板**: `Deck_Template_NWCD_dark_202103.pptx`（存放于 `references/` 目录）
- **版式数**: 16 种（索引 0-15）
- **风格**: 深色主题，专业商务
- **主色调**: 橙色 `#FFBD50`、深灰 `#313A40`、辅助灰 `#A0A8B0`

## 设计思维（源自前端设计美学）

生成 PPT 前，先确定设计方向：

- **色彩主导**: 一种颜色占60-70%视觉权重，1-2种辅助色，一种锐利强调色。不要平均分配。
- **视觉 motif**: 选定一个贯穿全部幻灯片的装饰元素（左侧色条、圆形图标背景、卡片边框等）。
- **排版层次**: 标题36-44pt，节标题20-24pt，正文14-16pt，注释10-12pt。大小对比要明显。
- **每页都要有视觉元素**: 色块、图标、形状、图表。纯文本页面是失败的。
- **布局多样性**: 不要每页都用相同版式。交替使用单列、双列、三列、四宫格卡片、统计数字等。
- **字体配对**: 标题用有个性的字体（Cambria、Georgia），正文用清晰的字体（Calibri）。避免Arial。

### NWCD 色彩体系

| 角色 | 颜色 | HEX | 用途 |
|------|------|-----|------|
| 品牌主色 | NWCD Orange | `#FFBD50` | 标题、强调、装饰 |
| 背景深色 | Dark Gray | `#313A40` | 幻灯片背景 |
| 卡片背景 | Card Gray | `#3A4349` | 卡片填充色 |
| 辅助灰 | Light Gray | `#A0A8B0` | 次要文本、注释（⚠️ 不要用 `#797979`，对比度不足） |
| 正文白 | White | `#FFFFFF` | 正文内容（默认文字颜色） |
| 浅灰白 | Light | `#CCCCCC` | 辅助说明文字 |
| 科技感 | Cyan | `#00D9FF` | 技术内容、链接、小标题 |
| 成功/增长 | Teal | `#00A896` | 正面指标、✅ 标记 |
| 警示/重点 | Coral | `#FF6B6B` | 关键警告、❌ 标记 |
| 创新/AI | Purple | `#8B5CF6` | 高级功能 |
| 提示/亮点 | Gold | `#FFC832` | 新功能标记 |

### 设计规则

1. **视觉焦点元素不换行**: 大字号（36pt+）的统计数字、日期、百分比等必须确保文本框宽度足够，设置 `word_wrap = False`，根据内容长度动态调整字号。
2. **不要在标题下加装饰线**: 这是 AI 生成幻灯片的典型特征，用留白或背景色代替。
3. **左对齐正文**: 段落和列表左对齐，只有标题可以居中。
4. **间距一致**: 选定 0.3" 或 0.5" 间距并全局统一使用。
5. **最小边距 0.5"**: 内容不要贴边。
6. **深色背景用浅色文字**: 确保对比度足够。

## ⭐ 演示适配性规范（v3.0 新增）

PPT 经常需要通过投影仪或投屏播放，必须考虑实际演示环境的限制。

### 高对比度原则

投影仪亮度不足是常见问题，文字与背景的对比度必须足够高：

| 场景 | 推荐颜色 | 禁止颜色 | 说明 |
|------|----------|----------|------|
| 深色背景上的正文 | `#FFFFFF` 白色 | `#797979` 中灰 | 中灰在投影仪上几乎不可见 |
| 深色背景上的次要文字 | `#A0A8B0` 浅灰 | `#666666` 以下 | 最低可接受亮度 |
| 深色背景上的注释 | `#A0A8B0` 浅灰 | `#888888` 以下 | 注释也必须可读 |
| 卡片背景 `#3A4349` 上的文字 | `#FFFFFF` 白色 | `#999999` 以下 | 卡片内文字同样需要高对比 |

**核心规则：在深色背景（`#313A40`）上，所有可读文字的颜色亮度值不得低于 `#A0A8B0`。**

**⚠️ 绝对禁止：使用 `#797979` 或更暗的灰色作为深色背景上的文字颜色。** 这在投影仪环境下几乎不可见，是非常不专业的表现。

### 内容溢出防护

文字与模板底部页脚（Logo、公司名称等）重叠是严重的排版错误，必须避免：

**安全区域定义：**
- 模板底部页脚区域约占 `Y > 6.8"` 的空间（包含 Logo、公司名称等）
- 所有内容必须在 `Y ≤ 6.6"` 的安全区域内
- 标题区域：`Y: 0.3" ~ 1.0"`
- 内容区域：`Y: 1.2" ~ 6.6"`（最大可用高度约 5.4"）

**内容量评估（生成前必须执行）：**

| 每列条目数 | 推荐方案 | 说明 |
|-----------|---------|------|
| ≤ 6 条 | 占位符版式（[6][7][8]） | 直接使用模板占位符 |
| 7-10 条 | 占位符 + 缩小字号（11-12pt） | 减少行间距，压缩空行 |
| 11-15 条 | **卡片布局**（[5] Only Title + 自定义） | 拆分为 2×2 四宫格卡片 |
| > 15 条 | **拆分为多页** | 单页信息过多影响阅读 |

**⚠️ 当单列内容超过 8 条时，强烈建议改用卡片布局或拆分多页，而不是缩小字号硬塞。**

### 卡片布局（内容密集页的推荐方案）

当内容较多时，使用 `[5] Only Title` 版式 + 自定义卡片布局，将内容分组到多个卡片中：

**四宫格卡片标准参数：**
```python
# 四宫格卡片布局（2×2）
# 每个卡片：宽 5.3"，高 2.7"
# 上排 Y=1.2"，下排 Y=4.2"
# 左列 X=0.6"，右列 X=6.2"
cards = [
    (0.6, 1.2, "左上标题", TEAL,   [...]),  # 左上
    (6.2, 1.2, "右上标题", TEAL,   [...]),  # 右上
    (0.6, 4.2, "左下标题", CORAL,  [...]),  # 左下
    (6.2, 4.2, "右下标题", CORAL,  [...]),  # 右下
]

for left, top, title, bar_color, lines in cards:
    add_card(slide, left, top, 5.3, 2.7, CARD_BG, top_bar_color=bar_color)
    add_textbox(slide, left + 0.25, top + 0.2, 4.8, 0.35, title,
                font_size=15, color=bar_color, bold=True)
    y = top + 0.65
    for line in lines:
        add_textbox(slide, left + 0.25, y, 4.8, 0.3, line,
                    font_size=12, color=WHITE)
        y += 0.33
```

**卡片布局的优势：**
- 内容分组清晰，视觉层次分明
- 每个卡片独立控制高度，不会溢出到页脚
- 顶部色条提供颜色编码，区分不同类别
- 卡片背景 `#3A4349` 与页面背景 `#313A40` 形成微妙层次感

**适用场景：**
- 优缺点对比（左列优势/右列劣势）
- 多维度对比（如平台支持 vs 分发方式）
- 激活机制等多属性并列对比
- 任何单列超过 8 条内容的页面

### 演示前检查清单

生成 PPT 后，必须逐页检查以下项目：

- [ ] **无溢出**：所有内容在 Y ≤ 6.6" 安全区域内，不与底部页脚重叠
- [ ] **高对比度**：深色背景上无 `#797979` 或更暗的文字颜色
- [ ] **可读性**：最小正文字号 ≥ 11pt，注释字号 ≥ 10pt
- [ ] **无截断**：长文本未被文本框边界截断
- [ ] **投影仪友好**：在降低屏幕亮度 50% 的情况下，所有文字仍可辨认

## 版式速查

| 索引 | 名称 | 占位符 | 推荐用途 |
|------|------|--------|----------|
| 0 | Title Slide | 12(标题) 13(副标题) 10(演讲者) 11(日期) | 封面 |
| 5 | Only Title | 0(标题) | 自定义内容页（卡片、统计、表格） ⭐ |
| 6 | Content | 0(标题) 10(内容) | 单列文本（≤8 条内容） |
| 7 | Two Content | 0(标题) 10(左) 11(右) | 双列对比（每列 ≤8 条） |
| 8 | Three Content | 0(标题) 10(左) 11(中) 12(右) | 三列展示（每列 ≤6 条） ⭐ |
| 11 | Goals | 0(标题) 18-21(四宫格) | 四要点 |
| 15 | Thanks | 0(标题) 10(联系信息) | 结束页 ⭐ |

其他版式：1(目录) 2(总结) 3(章节) 4(空白) 9(主标题) 10(内容标题) 12(流程) 13(分析表格) 14(定位图)

**版式选择决策树：**
```
内容条目数 ≤ 4 → [11] Goals 四宫格
内容条目数 ≤ 6 且单列 → [6] Content
内容条目数 ≤ 6 且需对比 → [7] Two Content
内容条目数 ≤ 6 且三维度 → [8] Three Content
内容条目数 7-15 → [5] Only Title + 卡片布局 ⭐
内容条目数 > 15 → 拆分为多页
```

## 推荐幻灯片组合

### 标准技术演示（6-9页）
```
[0] 封面 → [5] 数据亮点(自定义) → [8] 三列内容 → [5] 四宫格卡片
→ [7] 双列对比 → [8] 解决方案 → [6] 客户生态 → [5] 卡片布局 → [15] 结束
```

### 产品介绍（8页）
```
[0] 封面 → [6] 产品概述 → [8] 功能特性 → [11] 核心价值
→ [5] 竞品对比(卡片) → [7] 应用场景 → [6] 客户案例 → [15] 结束
```

### 对比分析（10-16页，v3.0 推荐）
```
[0] 封面 → [3] 章节页 → [7] 简短双列对比 → [3] 章节页
→ [5] 四宫格卡片对比 → [3] 章节页 → [8] 三列对比
→ [5] 合作伙伴卡片 → [5] 优缺点四宫格 → [5] 总结表格 → [15] 结束
```

## 视觉增强技巧（源自前端设计）

### 装饰元素
- **左侧色条**: 0.15" 宽的 NWCD Orange 竖条贯穿每页左边缘，统一视觉 motif（封面和结束页除外）
- **卡片色块**: 圆角矩形 + 彩色顶部色条，用于信息分组
- **图标圆圈**: 小圆形色块作为列表项或卡片的视觉锚点
- **顶部色条**: 卡片顶部 0.06" 高的彩色细线，区分不同类别

### 大数字统计展示（Stat Callouts）
```python
# 48pt 大数字 + 13pt 小标签，居中对齐
# word_wrap = False 防止换行
# 根据内容长度动态调整字号：≤4字符用48pt，更长用38pt
def add_stat_number(slide, number, label, left, top, box_width=2.5):
    font_size = 48 if len(number) <= 4 else 38
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(box_width), Inches(0.8))
    tf = txBox.text_frame
    tf.word_wrap = False  # 关键：禁止换行
    p = tf.paragraphs[0]
    p.text = number
    p.font.size = Pt(font_size)
    p.font.bold = True
```

### 卡片布局（v3.0 推荐）
```python
def add_card(slide, left, top, width, height, fill_color, border_color=None, top_bar_color=None):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color  # 推荐 #3A4349
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(1.5)
    else:
        shape.line.fill.background()
    shape.adjustments[0] = 0.05  # 圆角
    # 顶部色条
    if top_bar_color:
        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(left), Inches(top), Inches(width), Inches(0.06)
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = top_bar_color
        bar.line.fill.background()
    return shape
```

## 技术工作流

### 方式一：python-pptx（推荐用于模板）
```python
from pptx import Presentation
prs = Presentation('Deck_Template_NWCD_dark_202103.pptx')

# 删除模板示例页
for i in range(len(prs.slides) - 1, -1, -1):
    rId = prs.slides._sldIdLst[i].get(
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    prs.part.drop_rel(rId)
    prs.slides._sldIdLst.remove(prs.slides._sldIdLst[i])

# 添加新页面
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.placeholders[12].text = "标题"
prs.save('output.pptx')
```

### 方式二：PptxGenJS（适合从零创建）
```javascript
const pptxgen = require("pptxgenjs");
let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
let slide = pres.addSlide();
slide.background = { color: "313A40" };
slide.addText("标题", { x: 1, y: 0.3, w: 10, h: 0.8, fontSize: 36, color: "FFBD50", bold: true });
pres.writeFile({ fileName: "output.pptx" });
```

### 方式三：XML 编辑（适合精细修改）
```bash
python scripts/office/unpack.py template.pptx unpacked/
# 编辑 XML
python scripts/clean.py unpacked/
python scripts/office/pack.py unpacked/ output.pptx --original template.pptx
```

## QA 流程（必须执行）

1. **内容验证**: `python -m markitdown output.pptx` 检查文本完整性
2. **残留检查**: `grep -iE "lorem|ipsum|click to edit|fill in"` 确认无模板残留
3. **溢出检查**: 逐页确认内容未超出 Y=6.6" 安全线，未与底部页脚重叠
4. **对比度检查**: 确认无 `#797979` 或更暗的灰色用于深色背景上的文字
5. **视觉检查**: 转换为图片后逐页检查
   - 元素重叠、文本溢出、换行异常
   - 间距不均、边距不足（< 0.5"）
   - 低对比度文字、列未对齐
6. **投影仪模拟**: 降低屏幕亮度 50%，确认所有文字仍可辨认
7. **修复后重新验证**: 一次修复可能引入新问题

## 常见错误

- ❌ 占位符索引猜错 → 先打印 `slide.placeholders` 确认
- ❌ `paragraph.font.color = RGBColor(...)` → 应该是 `.color.rgb =`
- ❌ 所有统计数字用固定宽度文本框 → 根据内容长度调整
- ❌ 每页都用相同版式 → 交替使用不同布局
- ❌ 纯文本无视觉元素 → 每页至少一个形状/色块/图标
- ❌ PptxGenJS 颜色带 `#` → 只用 6 位 hex 如 `"FFBD50"`
- ❌ 复用 PptxGenJS option 对象 → 每次调用创建新对象
- ❌ **使用 `#797979` 作为文字颜色** → 在深色背景上对比度严重不足，投影仪下不可见，改用 `#A0A8B0`
- ❌ **内容溢出到底部页脚** → 内容必须在 Y ≤ 6.6" 安全区域内，超过 8 条内容改用卡片布局
- ❌ **单列塞太多内容** → 超过 8 条时改用 [5] Only Title + 四宫格卡片，而不是缩小字号硬塞
- ❌ **所有次要文字用同一种灰色** → 正文用白色 `#FFFFFF`，仅注释/脚注用 `#A0A8B0`

## v3.0 变更日志

相比 v2.0 的主要变更：

| 变更项 | v2.0 | v3.0 |
|--------|------|------|
| 能力维度 | 3 个（模板+设计+技术） | 4 个（+演示适配性） |
| 辅助灰色值 | `#797979`（对比度不足） | `#A0A8B0`（投影仪友好） |
| 内容溢出防护 | 仅在 QA 中提及 | 新增安全区域定义 + 内容量评估表 + 版式决策树 |
| 卡片布局 | 作为可选装饰 | 升级为内容密集页的推荐方案，提供标准参数 |
| 色彩体系 | 8 色 | 11 色（新增卡片背景、正文白、浅灰白） |
| QA 流程 | 4 步 | 7 步（新增溢出检查、对比度检查、投影仪模拟） |
| 常见错误 | 7 条 | 11 条（新增对比度、溢出、卡片布局相关） |

## 相关文件

- `template_layouts.md` - 16 种版式详细说明和代码示例
- `color_scheme.md` - 完整颜色方案和使用指南
- `best_practices.md` - PPT 生成最佳实践
- `code_examples.md` - Python 完整代码示例
- `common_patterns.md` - 10 种常见使用模式
