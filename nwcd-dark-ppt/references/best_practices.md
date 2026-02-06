# PPT 生成最佳实践

## 1. 模板使用原则

### 正确加载模板
```python
from pptx import Presentation

# ✅ 正确：从模板创建
prs = Presentation('Deck_Template_NWCD_dark_202103.pptx')

# ❌ 错误：创建空白演示文稿
prs = Presentation()
```

### 使用正确的版式索引
```python
# ✅ 正确：使用版式索引
slide = prs.slides.add_slide(prs.slide_layouts[0])  # Title Slide

# ❌ 错误：硬编码或猜测
slide = prs.slides.add_slide(prs.slide_layouts[99])
```

## 2. 占位符操作

### 查找占位符
```python
# 打印所有占位符信息
for placeholder in slide.placeholders:
    print(f"索引: {placeholder.placeholder_format.idx}")
    print(f"类型: {placeholder.placeholder_format.type}")
    print(f"名称: {placeholder.name}")
```

### 安全访问占位符
```python
# ✅ 正确：检查占位符是否存在
if 12 in [ph.placeholder_format.idx for ph in slide.placeholders]:
    title = slide.placeholders[12]
    title.text = "标题"

# ❌ 错误：直接访问可能不存在的占位符
title = slide.placeholders[12]  # 可能抛出 KeyError
```

## 3. 文本样式设置

### 统一样式函数
```python
def set_text_style(text_frame, font_size=20, color=WHITE, bold=False):
    """统一的文本样式设置"""
    for paragraph in text_frame.paragraphs:
        paragraph.font.size = Pt(font_size)
        paragraph.font.color.rgb = color
        paragraph.font.bold = bold
        paragraph.font.name = '微软雅黑'
```

### 避免重复代码
```python
# ✅ 正确：使用辅助函数
set_text_style(title.text_frame, font_size=44, color=NWCD_ORANGE, bold=True)

# ❌ 错误：重复代码
title.text_frame.paragraphs[0].font.size = Pt(44)
title.text_frame.paragraphs[0].font.color.rgb = NWCD_ORANGE
title.text_frame.paragraphs[0].font.bold = True
```

## 4. 颜色使用规范

### 定义颜色常量
```python
# ✅ 正确：在文件顶部定义
NWCD_ORANGE = RGBColor(255, 189, 80)
CYAN = RGBColor(0, 217, 255)

# ❌ 错误：到处硬编码
color = RGBColor(255, 189, 80)  # 重复多次
```

### 颜色层次
```python
# 标题 - 橙色
title.text_frame.paragraphs[0].font.color.rgb = NWCD_ORANGE

# 副标题 - 青色或灰色
subtitle.text_frame.paragraphs[0].font.color.rgb = CYAN

# 正文 - 白色
body.text_frame.paragraphs[0].font.color.rgb = WHITE
```

## 5. 内容组织

### 幻灯片数量建议
- **技术演示**: 5-8 页
- **产品介绍**: 8-12 页
- **详细培训**: 15-20 页
- **快速汇报**: 3-5 页

### 推荐结构
```python
# 标准 6 页结构
1. Title Slide (封面)
2. Three Content (技术栈/概述)
3. Content (详细说明)
4. Goals (核心优势)
5. Two Content (应用场景)
6. Thanks (结束)
```

## 6. 文本内容规范

### 标题长度
- **主标题**: 10-20 字
- **副标题**: 15-30 字
- **内容标题**: 5-15 字

### 项目符号
```python
# ✅ 正确：简洁明了
points = [
    "• 自动化架构生成",
    "• 智能图标选择",
    "• 实时预览"
]

# ❌ 错误：过长
points = [
    "• 通过 AI 技术实现自动化的架构图生成功能，大大提高了工作效率..."
]
```

### 每页内容量
- **标题**: 1 行
- **要点**: 3-5 个
- **每个要点**: 1-2 行
- **总字数**: 50-100 字

## 7. 版式选择指南

### 根据内容选择版式
```python
# 封面/结束
layouts[0]  # Title Slide
layouts[15]  # Thanks

# 文字内容
layouts[6]  # Content (单列)
layouts[7]  # Two Content (双列)
layouts[8]  # Three Content (三列)

# 特殊用途
layouts[11]  # Goals (四宫格)
layouts[13]  # Analysis (表格)
layouts[5]  # Only Title (图片/图表)
```

### 版式组合建议
```python
# 技术演示
[0, 8, 6, 11, 7, 15]

# 产品介绍
[0, 1, 3, 6, 11, 13, 7, 15]

# 架构说明
[0, 5, 8, 6, 15]
```

## 8. 性能优化

### 批量操作
```python
# ✅ 正确：一次性创建多个幻灯片
slides_data = [...]
for data in slides_data:
    create_slide(prs, data)

# ❌ 错误：频繁保存
for data in slides_data:
    create_slide(prs, data)
    prs.save('temp.pptx')  # 不要这样做
```

### 资源管理
```python
# ✅ 正确：最后保存一次
prs = Presentation('template.pptx')
# ... 创建所有幻灯片
prs.save('output.pptx')

# ❌ 错误：多次加载模板
for i in range(10):
    prs = Presentation('template.pptx')  # 浪费资源
```

## 9. 错误处理

### 安全的占位符访问
```python
def safe_set_placeholder(slide, idx, text):
    """安全设置占位符文本"""
    try:
        placeholder = slide.placeholders[idx]
        placeholder.text = text
        return True
    except KeyError:
        print(f"警告：占位符 {idx} 不存在")
        return False
```

### 验证模板
```python
def validate_template(template_path):
    """验证模板是否有效"""
    try:
        prs = Presentation(template_path)
        if len(prs.slide_layouts) < 16:
            print("警告：模板版式数量不足")
            return False
        return True
    except Exception as e:
        print(f"错误：无法加载模板 - {e}")
        return False
```

## 10. 调试技巧

### 打印版式信息
```python
def print_layout_info(prs):
    """打印所有版式信息"""
    for i, layout in enumerate(prs.slide_layouts):
        print(f"\n[{i}] {layout.name}")
        print(f"  占位符数量: {len(layout.placeholders)}")
        for ph in layout.placeholders:
            print(f"    - [{ph.placeholder_format.idx}] {ph.name}")
```

### 打印幻灯片信息
```python
def print_slide_info(slide):
    """打印幻灯片信息"""
    print(f"版式: {slide.slide_layout.name}")
    print(f"形状数量: {len(slide.shapes)}")
    for shape in slide.shapes:
        if shape.has_text_frame:
            print(f"  文本: {shape.text[:50]}...")
```

## 11. 代码组织

### 模块化结构
```
project/
├── config.py          # 颜色、字体等配置
├── utils.py           # 辅助函数
├── slide_creators.py  # 幻灯片创建函数
└── main.py           # 主程序
```

### config.py 示例
```python
from pptx.dml.color import RGBColor

# 颜色配置
COLORS = {
    'orange': RGBColor(255, 189, 80),
    'cyan': RGBColor(0, 217, 255),
    'white': RGBColor(255, 255, 255),
}

# 字体配置
FONTS = {
    'title': {'size': 44, 'bold': True},
    'subtitle': {'size': 28, 'bold': False},
    'body': {'size': 20, 'bold': False},
}

# 版式配置
LAYOUTS = {
    'title': 0,
    'three_content': 8,
    'content': 6,
    'goals': 11,
    'two_content': 7,
    'thanks': 15,
}
```

## 12. 测试建议

### 单元测试
```python
import unittest

class TestPPTGeneration(unittest.TestCase):
    def test_create_title_slide(self):
        prs = Presentation('template.pptx')
        slide = create_title_slide(prs, "测试标题", "测试副标题")
        self.assertIsNotNone(slide)
        self.assertEqual(len(prs.slides), 1)
    
    def test_color_values(self):
        self.assertEqual(NWCD_ORANGE.rgb, (255, 189, 80))
```

### 视觉检查清单
- [ ] 标题颜色正确（橙色）
- [ ] 文本可读性良好
- [ ] 版式对齐整齐
- [ ] 字体大小合适
- [ ] 内容不超出边界
- [ ] 颜色搭配协调

## 13. 常见错误

### 错误 1：占位符索引错误
```python
# ❌ 错误
title = slide.placeholders[0]  # 可能不是标题

# ✅ 正确
title = slide.placeholders[12]  # 根据分析结果使用正确索引
```

### 错误 2：文本框清空方式
```python
# ❌ 错误
text_frame.text = ""  # 可能导致格式丢失

# ✅ 正确
text_frame.clear()  # 正确清空
```

### 错误 3：颜色设置
```python
# ❌ 错误
paragraph.font.color = RGBColor(255, 189, 80)

# ✅ 正确
paragraph.font.color.rgb = RGBColor(255, 189, 80)
```

## 14. 版本控制

### Git 忽略文件
```gitignore
# .gitignore
*.pptx
!Deck_Template_NWCD_dark_202103.pptx  # 保留模板
__pycache__/
*.pyc
venv/
```

### 文件命名规范
```
Kiro_MCP_NWCD_Professional_v1.0.pptx
Kiro_MCP_NWCD_Professional_v1.1.pptx
```

## 15. 文档化

### 代码注释
```python
def create_title_slide(prs, title_text, subtitle_text):
    """
    创建标题页
    
    Args:
        prs: Presentation 对象
        title_text: 主标题文本
        subtitle_text: 副标题文本
    
    Returns:
        Slide: 创建的幻灯片对象
    
    Example:
        >>> create_title_slide(prs, "标题", "副标题")
    """
    pass
```

### README 文档
```markdown
# PPT 生成工具

## 使用方法
1. 确保模板文件存在
2. 运行 `python main.py`
3. 输出文件在当前目录

## 依赖
- python-pptx >= 0.6.21
```
