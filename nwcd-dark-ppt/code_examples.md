# PPT 生成代码示例

## 完整示例：生成 Kiro MCP 演示文稿

```python
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ============ 颜色定义 ============
NWCD_ORANGE = RGBColor(255, 189, 80)
DARK_GRAY = RGBColor(49, 58, 64)
MEDIUM_GRAY = RGBColor(121, 121, 121)
CYAN = RGBColor(0, 217, 255)
NEON_PINK = RGBColor(255, 0, 110)
PURPLE = RGBColor(139, 92, 246)
YELLOW = RGBColor(255, 215, 0)
WHITE = RGBColor(255, 255, 255)

# ============ 辅助函数 ============

def set_text_style(text_frame, font_size=20, color=WHITE, bold=False, font_name='微软雅黑'):
    """设置文本样式"""
    for paragraph in text_frame.paragraphs:
        paragraph.font.size = Pt(font_size)
        paragraph.font.color.rgb = color
        paragraph.font.bold = bold
        paragraph.font.name = font_name
        paragraph.alignment = PP_ALIGN.LEFT

def add_bullet_points(text_frame, points, color=WHITE, font_size=18):
    """添加项目符号列表"""
    text_frame.clear()
    for i, point in enumerate(points):
        if i == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        p.text = point
        p.level = 0
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.name = '微软雅黑'

def create_title_slide(prs, title_text, subtitle_text, presenter="技术团队", date="2026年2月"):
    """创建标题页"""
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    
    # 主标题
    title = slide.placeholders[12]
    title.text = title_text
    set_text_style(title.text_frame, font_size=44, color=NWCD_ORANGE, bold=True)
    
    # 副标题
    subtitle = slide.placeholders[13]
    subtitle.text = subtitle_text
    set_text_style(subtitle.text_frame, font_size=28, color=CYAN)
    
    # 演讲者
    if len(slide.placeholders) > 10:
        presenter_ph = slide.placeholders[10]
        presenter_ph.text = presenter
        set_text_style(presenter_ph.text_frame, font_size=18, color=MEDIUM_GRAY)
    
    # 日期
    if len(slide.placeholders) > 11:
        date_ph = slide.placeholders[11]
        date_ph.text = date
        set_text_style(date_ph.text_frame, font_size=18, color=MEDIUM_GRAY)
    
    return slide

def create_three_content_slide(prs, title_text, left_content, middle_content, right_content):
    """创建三列内容页"""
    slide = prs.slides.add_slide(prs.slide_layouts[8])
    
    # 标题
    title = slide.placeholders[0]
    title.text = title_text
    set_text_style(title.text_frame, font_size=36, color=NWCD_ORANGE, bold=True)
    
    # 左侧内容
    left = slide.placeholders[10]
    add_bullet_points(left.text_frame, left_content, color=WHITE)
    
    # 中间内容
    middle = slide.placeholders[11]
    add_bullet_points(middle.text_frame, middle_content, color=WHITE)
    
    # 右侧内容
    right = slide.placeholders[12]
    add_bullet_points(right.text_frame, right_content, color=WHITE)
    
    return slide

def create_content_slide(prs, title_text, content_points):
    """创建单列内容页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 标题
    title = slide.placeholders[0]
    title.text = title_text
    set_text_style(title.text_frame, font_size=36, color=NWCD_ORANGE, bold=True)
    
    # 内容
    content = slide.placeholders[10]
    add_bullet_points(content.text_frame, content_points, color=WHITE)
    
    return slide

def create_goals_slide(prs, title_text, goals):
    """创建四宫格目标页"""
    slide = prs.slides.add_slide(prs.slide_layouts[11])
    
    # 标题
    title = slide.placeholders[0]
    title.text = title_text
    set_text_style(title.text_frame, font_size=36, color=NWCD_ORANGE, bold=True)
    
    # 四个目标
    goal_placeholders = [18, 19, 20, 21]
    colors = [CYAN, NEON_PINK, PURPLE, YELLOW]
    
    for i, (goal_text, ph_idx, color) in enumerate(zip(goals, goal_placeholders, colors)):
        if ph_idx in [ph.placeholder_format.idx for ph in slide.placeholders]:
            goal = slide.placeholders[ph_idx]
            goal.text = goal_text
            set_text_style(goal.text_frame, font_size=20, color=color, bold=True)
    
    return slide

def create_two_content_slide(prs, title_text, left_content, right_content):
    """创建双列内容页"""
    slide = prs.slides.add_slide(prs.slide_layouts[7])
    
    # 标题
    title = slide.placeholders[0]
    title.text = title_text
    set_text_style(title.text_frame, font_size=36, color=NWCD_ORANGE, bold=True)
    
    # 左侧内容
    left = slide.placeholders[10]
    add_bullet_points(left.text_frame, left_content, color=WHITE)
    
    # 右侧内容
    right = slide.placeholders[11]
    add_bullet_points(right.text_frame, right_content, color=WHITE)
    
    return slide

def create_thanks_slide(prs, thanks_text="Thanks！", contact_info=""):
    """创建结束页"""
    slide = prs.slides.add_slide(prs.slide_layouts[15])
    
    # 感谢标题
    title = slide.placeholders[0]
    title.text = thanks_text
    set_text_style(title.text_frame, font_size=60, color=NWCD_ORANGE, bold=True)
    
    # 联系信息
    if contact_info and len(slide.placeholders) > 10:
        contact = slide.placeholders[10]
        contact.text = contact_info
        set_text_style(contact.text_frame, font_size=24, color=CYAN)
    
    return slide

# ============ 主程序 ============

def main():
    # 加载模板
    prs = Presentation('Deck_Template_NWCD_dark_202103.pptx')
    
    # 第1页：标题页
    create_title_slide(
        prs,
        title_text="Kiro MCP 架构自动化",
        subtitle_text="基于 Model Context Protocol 的智能架构生成",
        presenter="技术团队",
        date="2026年2月"
    )
    
    # 第2页：技术栈（三列）
    create_three_content_slide(
        prs,
        title_text="技术栈",
        left_content=[
            "前端",
            "• React",
            "• TypeScript",
            "• Vite"
        ],
        middle_content=[
            "后端",
            "• Python",
            "• FastAPI",
            "• MCP SDK"
        ],
        right_content=[
            "基础设施",
            "• AWS",
            "• Docker",
            "• GitHub Actions"
        ]
    )
    
    # 第3页：工作流程
    create_content_slide(
        prs,
        title_text="工作流程",
        content_points=[
            "1. 用户通过 Kiro 输入架构需求",
            "2. MCP 协议连接 Draw.io 服务器",
            "3. AI 分析需求并选择合适的 AWS 图标",
            "4. 自动生成架构图并保存",
            "5. 支持实时预览和迭代优化"
        ]
    )
    
    # 第4页：核心优势（四宫格）
    create_goals_slide(
        prs,
        title_text="核心优势",
        goals=[
            "1. 自动化\n快速生成专业架构图\n节省 80% 时间",
            "2. 智能化\nAI 理解需求\n自动选择最佳方案",
            "3. 可扩展\n支持多种图标库\n灵活定制",
            "4. 易集成\nMCP 标准协议\n无缝对接"
        ]
    )
    
    # 第5页：使用场景（双列）
    create_two_content_slide(
        prs,
        title_text="使用场景",
        left_content=[
            "开发环境",
            "• 快速原型设计",
            "• 技术方案评审",
            "• 架构文档生成",
            "• 团队协作讨论"
        ],
        right_content=[
            "生产环境",
            "• 系统架构设计",
            "• 容量规划",
            "• 灾备方案",
            "• 安全审计"
        ]
    )
    
    # 第6页：结束页
    create_thanks_slide(
        prs,
        thanks_text="Thanks！",
        contact_info="联系方式：tech@example.com"
    )
    
    # 保存
    prs.save('Kiro_MCP_NWCD_Professional.pptx')
    print("✅ PPT 生成成功！")

if __name__ == '__main__':
    main()
```

## 高级示例：带图表的内容页

```python
def create_content_with_chart(prs, title_text, content_points, chart_data):
    """创建带图表的内容页"""
    slide = prs.slides.add_slide(prs.slide_layouts[7])  # Two Content
    
    # 标题
    title = slide.placeholders[0]
    title.text = title_text
    set_text_style(title.text_frame, font_size=36, color=NWCD_ORANGE, bold=True)
    
    # 左侧：文字内容
    left = slide.placeholders[10]
    add_bullet_points(left.text_frame, content_points, color=WHITE)
    
    # 右侧：图表（这里可以添加图表代码）
    # 注意：python-pptx 支持添加图表，但需要额外配置
    
    return slide
```

## 示例：自定义形状和文本框

```python
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE

def add_custom_shape(slide, text, left, top, width, height, color=CYAN):
    """添加自定义形状"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height)
    )
    
    # 设置填充颜色
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    
    # 设置文本
    text_frame = shape.text_frame
    text_frame.text = text
    set_text_style(text_frame, font_size=18, color=WHITE, bold=True)
    
    # 居中对齐
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    return shape

# 使用示例
slide = prs.slides.add_slide(prs.slide_layouts[5])  # Only Title
title = slide.placeholders[0]
title.text = "自定义布局"

# 添加三个彩色方块
add_custom_shape(slide, "前端", 1, 2, 3, 1.5, CYAN)
add_custom_shape(slide, "后端", 4.5, 2, 3, 1.5, NEON_PINK)
add_custom_shape(slide, "数据库", 8, 2, 3, 1.5, PURPLE)
```

## 示例：添加图片

```python
def add_image_to_slide(prs, title_text, image_path):
    """添加图片到幻灯片"""
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Only Title
    
    # 标题
    title = slide.placeholders[0]
    title.text = title_text
    set_text_style(title.text_frame, font_size=36, color=NWCD_ORANGE, bold=True)
    
    # 添加图片（居中）
    left = Inches(2)
    top = Inches(2)
    width = Inches(8)
    
    pic = slide.shapes.add_picture(image_path, left, top, width=width)
    
    return slide

# 使用示例
# add_image_to_slide(prs, "系统架构图", "architecture.png")
```

## 示例：表格

```python
def create_comparison_table(prs, title_text, headers, rows):
    """创建对比表格"""
    slide = prs.slides.add_slide(prs.slide_layouts[13])  # Analysis
    
    # 标题
    title = slide.placeholders[0]
    title.text = title_text
    set_text_style(title.text_frame, font_size=36, color=NWCD_ORANGE, bold=True)
    
    # 表格
    table_placeholder = slide.placeholders[10]
    graphic_frame = table_placeholder.insert_table(rows=len(rows)+1, cols=len(headers))
    table = graphic_frame.table
    
    # 设置表头
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = header
        set_text_style(cell.text_frame, font_size=16, color=NWCD_ORANGE, bold=True)
    
    # 填充数据
    for row_idx, row_data in enumerate(rows, start=1):
        for col_idx, cell_data in enumerate(row_data):
            cell = table.cell(row_idx, col_idx)
            cell.text = str(cell_data)
            set_text_style(cell.text_frame, font_size=14, color=WHITE)
    
    return slide

# 使用示例
# create_comparison_table(
#     prs,
#     "功能对比",
#     headers=["功能", "方案A", "方案B"],
#     rows=[
#         ["性能", "高", "中"],
#         ["成本", "低", "高"],
#         ["维护", "简单", "复杂"]
#     ]
# )
```

## 完整工作流

```python
def generate_presentation(template_path, output_path, content_data):
    """
    生成完整演示文稿
    
    Args:
        template_path: 模板文件路径
        output_path: 输出文件路径
        content_data: 内容数据字典
    """
    prs = Presentation(template_path)
    
    # 标题页
    if 'title' in content_data:
        create_title_slide(prs, **content_data['title'])
    
    # 内容页
    if 'slides' in content_data:
        for slide_data in content_data['slides']:
            slide_type = slide_data.get('type')
            
            if slide_type == 'three_content':
                create_three_content_slide(prs, **slide_data['content'])
            elif slide_type == 'content':
                create_content_slide(prs, **slide_data['content'])
            elif slide_type == 'goals':
                create_goals_slide(prs, **slide_data['content'])
            elif slide_type == 'two_content':
                create_two_content_slide(prs, **slide_data['content'])
    
    # 结束页
    if 'thanks' in content_data:
        create_thanks_slide(prs, **content_data['thanks'])
    
    prs.save(output_path)
    return output_path

# 使用示例
content = {
    'title': {
        'title_text': "Kiro MCP 架构自动化",
        'subtitle_text': "基于 Model Context Protocol 的智能架构生成"
    },
    'slides': [
        {
            'type': 'three_content',
            'content': {
                'title_text': "技术栈",
                'left_content': ["前端", "• React", "• TypeScript"],
                'middle_content': ["后端", "• Python", "• FastAPI"],
                'right_content': ["基础设施", "• AWS", "• Docker"]
            }
        }
    ],
    'thanks': {
        'thanks_text': "Thanks！",
        'contact_info': "tech@example.com"
    }
}

# generate_presentation('template.pptx', 'output.pptx', content)
```
