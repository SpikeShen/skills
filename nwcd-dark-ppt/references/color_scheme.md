# NWCD Dark 模板颜色方案

## 主色调

### 品牌橙色 - NWCD Orange
- **HEX**: `#FFBD50`
- **RGB**: (255, 189, 80)
- **用途**: 强调色、标题、重点内容
- **使用场景**: 
  - 主标题文字
  - 重要图标
  - 按钮和链接
  - 数据可视化的主色

### 深灰色 - Dark Gray
- **HEX**: `#313A40`
- **RGB**: (49, 58, 64)
- **用途**: 背景色、深色区域
- **使用场景**:
  - 幻灯片背景
  - 深色卡片
  - 分隔区域

### 中灰色 - Medium Gray
- **HEX**: `#797979`
- **RGB**: (121, 121, 121)
- **用途**: 次要文本、辅助信息
- **使用场景**:
  - 副标题
  - 说明文字
  - 图表标签

## 辅助色彩（可选）

### 科技蓝 - Cyan
- **HEX**: `#00D9FF`
- **RGB**: (0, 217, 255)
- **用途**: 技术感、现代感
- **使用场景**:
  - 技术架构图
  - 数据流向
  - 代码块高亮

### 霓虹粉 - Neon Pink
- **HEX**: `#FF006E`
- **RGB**: (255, 0, 110)
- **用途**: 警示、特别强调
- **使用场景**:
  - 重要警告
  - 关键指标
  - 特殊标记

### 紫色 - Purple
- **HEX**: `#8B5CF6`
- **RGB**: (139, 92, 246)
- **用途**: 创新、高级感
- **使用场景**:
  - AI/ML 相关内容
  - 高级功能
  - 渐变效果

### 黄色 - Yellow
- **HEX**: `#FFD700`
- **RGB**: (255, 215, 0)
- **用途**: 提示、注意
- **使用场景**:
  - 提示信息
  - 新功能标记
  - 亮点展示

## 文本颜色规范

### 主标题
- **颜色**: NWCD Orange (#FFBD50)
- **字号**: 44pt
- **字体**: 微软雅黑 Bold / Arial Bold

### 副标题
- **颜色**: Medium Gray (#797979)
- **字号**: 28pt
- **字体**: 微软雅黑 / Arial

### 正文
- **颜色**: 白色 (#FFFFFF) 或浅灰 (#E0E0E0)
- **字号**: 18-24pt
- **字体**: 微软雅黑 / Arial

### 强调文本
- **颜色**: NWCD Orange (#FFBD50) 或 Cyan (#00D9FF)
- **字号**: 与正文相同
- **字体**: 加粗

## Python 代码实现

### 设置文本颜色

```python
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# NWCD 颜色定义
NWCD_ORANGE = RGBColor(255, 189, 80)
DARK_GRAY = RGBColor(49, 58, 64)
MEDIUM_GRAY = RGBColor(121, 121, 121)
CYAN = RGBColor(0, 217, 255)
NEON_PINK = RGBColor(255, 0, 110)
PURPLE = RGBColor(139, 92, 246)
YELLOW = RGBColor(255, 215, 0)
WHITE = RGBColor(255, 255, 255)

# 应用到标题
def set_title_style(shape, text):
    """设置标题样式"""
    shape.text = text
    text_frame = shape.text_frame
    paragraph = text_frame.paragraphs[0]
    paragraph.font.size = Pt(44)
    paragraph.font.bold = True
    paragraph.font.color.rgb = NWCD_ORANGE
    paragraph.font.name = '微软雅黑'
    paragraph.alignment = PP_ALIGN.LEFT

# 应用到副标题
def set_subtitle_style(shape, text):
    """设置副标题样式"""
    shape.text = text
    text_frame = shape.text_frame
    paragraph = text_frame.paragraphs[0]
    paragraph.font.size = Pt(28)
    paragraph.font.color.rgb = MEDIUM_GRAY
    paragraph.font.name = '微软雅黑'
    paragraph.alignment = PP_ALIGN.LEFT

# 应用到正文
def set_body_style(shape, text, color=WHITE):
    """设置正文样式"""
    shape.text = text
    text_frame = shape.text_frame
    paragraph = text_frame.paragraphs[0]
    paragraph.font.size = Pt(20)
    paragraph.font.color.rgb = color
    paragraph.font.name = '微软雅黑'
    paragraph.alignment = PP_ALIGN.LEFT

# 添加带颜色的段落
def add_colored_paragraph(text_frame, text, color=WHITE, level=0):
    """添加带颜色的段落"""
    p = text_frame.add_paragraph()
    p.text = text
    p.level = level
    p.font.size = Pt(18)
    p.font.color.rgb = color
    p.font.name = '微软雅黑'
    return p
```

### 完整示例

```python
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor

# 颜色定义
NWCD_ORANGE = RGBColor(255, 189, 80)
CYAN = RGBColor(0, 217, 255)
WHITE = RGBColor(255, 255, 255)

# 加载模板
prs = Presentation('Deck_Template_NWCD_dark_202103.pptx')

# 创建标题页
slide = prs.slides.add_slide(prs.slide_layouts[0])
title = slide.placeholders[12]
subtitle = slide.placeholders[13]

# 设置标题（橙色）
title.text = "Kiro MCP 架构自动化"
title.text_frame.paragraphs[0].font.color.rgb = NWCD_ORANGE
title.text_frame.paragraphs[0].font.size = Pt(44)
title.text_frame.paragraphs[0].font.bold = True

# 设置副标题（青色）
subtitle.text = "基于 Model Context Protocol 的智能架构生成"
subtitle.text_frame.paragraphs[0].font.color.rgb = CYAN
subtitle.text_frame.paragraphs[0].font.size = Pt(28)

# 创建内容页
slide = prs.slides.add_slide(prs.slide_layouts[8])  # Three Content
title = slide.placeholders[0]
left = slide.placeholders[10]
middle = slide.placeholders[11]
right = slide.placeholders[12]

# 标题（橙色）
title.text = "技术栈"
title.text_frame.paragraphs[0].font.color.rgb = NWCD_ORANGE

# 左侧内容（白色）
tf = left.text_frame
tf.clear()
p = tf.paragraphs[0]
p.text = "前端"
p.font.color.rgb = NWCD_ORANGE
p.font.size = Pt(24)
p.font.bold = True

p = tf.add_paragraph()
p.text = "• React"
p.font.color.rgb = WHITE
p.font.size = Pt(18)

p = tf.add_paragraph()
p.text = "• TypeScript"
p.font.color.rgb = WHITE
p.font.size = Pt(18)

# 保存
prs.save('output.pptx')
```

## 颜色搭配建议

### 技术类演示
- **主色**: NWCD Orange + Cyan
- **背景**: Dark Gray
- **文本**: White
- **强调**: Neon Pink

### 商务类演示
- **主色**: NWCD Orange
- **背景**: Dark Gray
- **文本**: White
- **强调**: Medium Gray

### 创新类演示
- **主色**: Purple + Cyan
- **背景**: Dark Gray
- **文本**: White
- **强调**: NWCD Orange

## 渐变效果（高级）

```python
from pptx.enum.dml import MSO_FILL

def set_gradient_fill(shape, color1, color2):
    """设置渐变填充"""
    fill = shape.fill
    fill.gradient()
    fill.gradient_angle = 90.0
    
    # 第一个颜色
    fill.gradient_stops[0].color.rgb = color1
    
    # 第二个颜色
    fill.gradient_stops[1].color.rgb = color2

# 使用示例
# set_gradient_fill(shape, NWCD_ORANGE, PURPLE)
```

## 注意事项

1. **对比度**: 深色背景上使用浅色文字，确保可读性
2. **一致性**: 同类内容使用相同颜色
3. **层次感**: 标题、副标题、正文使用不同颜色区分
4. **强调**: 重要内容使用 NWCD Orange 或 Cyan
5. **适度**: 不要在一页中使用超过 3-4 种颜色
