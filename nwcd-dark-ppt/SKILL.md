---
name: NWCD Dark PPT Generation
version: 2.0.0
description: 生成符合 NWCD Dark 模板风格的 PowerPoint 演示文稿。融合了前端设计美学和 PPTX 技术能力，提供从模板分析、视觉设计到内容填充的完整工作流。
---

# NWCD Dark PPT Generation Skill (v2.0)

融合了三个能力维度：
1. **NWCD 模板知识** — 版式结构、占位符映射、品牌色彩
2. **前端设计美学** — 视觉层次、色彩理论、排版节奏、装饰元素
3. **PPTX 技术能力** — python-pptx / PptxGenJS / XML 编辑工作流

## 模板信息

- **模板**: `Deck_Template_NWCD_dark_202103.pptx`
- **版式数**: 16 种（索引 0-15）
- **风格**: 深色主题，专业商务
- **主色调**: 橙色 `#FFBD50`、深灰 `#313A40`、中灰 `#797979`

## 设计思维（源自前端设计美学）

生成 PPT 前，先确定设计方向：

- **色彩主导**: 一种颜色占60-70%视觉权重，1-2种辅助色，一种锐利强调色。不要平均分配。
- **视觉 motif**: 选定一个贯穿全部幻灯片的装饰元素（左侧色条、圆形图标背景、卡片边框等）。
- **排版层次**: 标题36-44pt，节标题20-24pt，正文14-16pt，注释10-12pt。大小对比要明显。
- **每页都要有视觉元素**: 色块、图标、形状、图表。纯文本页面是失败的。
- **布局多样性**: 不要每页都用相同版式。交替使用单列、双列、三列、四宫格、统计数字等。
- **字体配对**: 标题用有个性的字体（Cambria、Georgia），正文用清晰的字体（Calibri）。避免Arial。

### NWCD 色彩体系

| 角色 | 颜色 | HEX | 用途 |
|------|------|-----|------|
| 品牌主色 | NWCD Orange | `#FFBD50` | 标题、强调、装饰 |
| 背景深色 | Dark Gray | `#313A40` | 幻灯片背景 |
| 辅助灰 | Medium Gray | `#797979` | 次要文本、注释 |
| 科技感 | Cyan | `#00D9FF` | 技术内容、链接 |
| 成功/增长 | Teal | `#00A896` | 正面指标 |
| 警示/重点 | Coral | `#FF6B6B` | 关键警告 |
| 创新/AI | Purple | `#8B5CF6` | 高级功能 |
| 提示/亮点 | Gold | `#FFC832` | 新功能标记 |

### 设计规则

1. **视觉焦点元素不换行**: 大字号（36pt+）的统计数字、日期、百分比等必须确保文本框宽度足够，设置 `word_wrap = False`，根据内容长度动态调整字号。
2. **不要在标题下加装饰线**: 这是 AI 生成幻灯片的典型特征，用留白或背景色代替。
3. **左对齐正文**: 段落和列表左对齐，只有标题可以居中。
4. **间距一致**: 选定 0.3" 或 0.5" 间距并全局统一使用。
5. **最小边距 0.5"**: 内容不要贴边。
6. **深色背景用浅色文字**: 确保对比度足够。

## 版式速查

| 索引 | 名称 | 占位符 | 推荐用途 |
|------|------|--------|----------|
| 0 | Title Slide | 12(标题) 13(副标题) 10(演讲者) 11(日期) | 封面 |
| 5 | Only Title | 0(标题) | 自定义内容页（统计、卡片） |
| 6 | Content | 0(标题) 10(内容) | 单列文本 |
| 7 | Two Content | 0(标题) 10(左) 11(右) | 双列对比 |
| 8 | Three Content | 0(标题) 10(左) 11(中) 12(右) | 三列展示 ⭐ |
| 11 | Goals | 0(标题) 18-21(四宫格) | 四要点 |
| 15 | Thanks | 0(标题) 10(联系信息) | 结束页 ⭐ |

其他版式：1(目录) 2(总结) 3(章节) 4(空白) 9(主标题) 10(内容标题) 12(流程) 13(分析表格) 14(定位图)

## 推荐幻灯片组合

### 标准技术演示（6-9页）
```
[0] 封面 → [5] 数据亮点(自定义) → [8] 三列内容 → [5] 四宫格优势(自定义)
→ [7] 双列对比 → [8] 解决方案 → [6] 客户生态 → [7] 基础设施 → [15] 结束
```

### 产品介绍（8页）
```
[0] 封面 → [6] 产品概述 → [8] 功能特性 → [11] 核心价值
→ [13] 竞品对比 → [7] 应用场景 → [6] 客户案例 → [15] 结束
```

## 视觉增强技巧（源自前端设计）

### 装饰元素
- **左侧色条**: 0.15" 宽的 NWCD Orange 竖条贯穿每页左边缘，统一视觉 motif
- **卡片色块**: 圆角矩形 + 彩色边框，用于信息分组
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

### 卡片布局
```python
def add_card(slide, left, top, width, height, fill_color, border_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, ...)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(1.5)
    shape.adjustments[0] = 0.05  # 圆角
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
3. **视觉检查**: 转换为图片后逐页检查
   - 元素重叠、文本溢出、换行异常
   - 间距不均、边距不足（< 0.5"）
   - 低对比度文字、列未对齐
4. **修复后重新验证**: 一次修复可能引入新问题

## 常见错误

- ❌ 占位符索引猜错 → 先打印 `slide.placeholders` 确认
- ❌ `paragraph.font.color = RGBColor(...)` → 应该是 `.color.rgb =`
- ❌ 所有统计数字用固定宽度文本框 → 根据内容长度调整
- ❌ 每页都用相同版式 → 交替使用不同布局
- ❌ 纯文本无视觉元素 → 每页至少一个形状/色块/图标
- ❌ PptxGenJS 颜色带 `#` → 只用 6 位 hex 如 `"FFBD50"`
- ❌ 复用 PptxGenJS option 对象 → 每次调用创建新对象

## 相关文件

- `template_layouts.md` - 16 种版式详细说明和代码示例
- `color_scheme.md` - 完整颜色方案和使用指南
- `best_practices.md` - PPT 生成最佳实践
- `code_examples.md` - Python 完整代码示例
- `common_patterns.md` - 10 种常见使用模式
