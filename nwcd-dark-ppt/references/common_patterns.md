# 常见使用模式

## 模式 1：快速生成标准演示

适用于：技术分享、产品介绍、项目汇报

```python
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

# 配置
TEMPLATE = 'Deck_Template_NWCD_dark_202103.pptx'
OUTPUT = 'output.pptx'

# 颜色
NWCD_ORANGE = RGBColor(255, 189, 80)
CYAN = RGBColor(0, 217, 255)
WHITE = RGBColor(255, 255, 255)

# 创建演示
prs = Presentation(TEMPLATE)

# 1. 封面
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.placeholders[12].text = "项目标题"
slide.placeholders[13].text = "项目副标题"

# 2. 内容页
slide = prs.slides.add_slide(prs.slide_layouts[8])
slide.placeholders[0].text = "技术栈"
slide.placeholders[10].text = "前端\n• React\n• TypeScript"
slide.placeholders[11].text = "后端\n• Python\n• FastAPI"
slide.placeholders[12].text = "基础设施\n• AWS\n• Docker"

# 3. 结束页
slide = prs.slides.add_slide(prs.slide_layouts[15])
slide.placeholders[0].text = "Thanks！"

prs.save(OUTPUT)
```

## 模式 2：数据驱动生成

适用于：批量生成、模板化内容

```python
# 数据结构
presentation_data = {
    'title': {
        'main': 'Kiro MCP 架构自动化',
        'sub': '基于 Model Context Protocol 的智能架构生成',
        'presenter': '技术团队',
        'date': '2026年2月'
    },
    'slides': [
        {
            'layout': 8,  # Three Content
            'title': '技术栈',
            'content': {
                'left': ['前端', '• React', '• TypeScript'],
                'middle': ['后端', '• Python', '• FastAPI'],
                'right': ['基础设施', '• AWS', '• Docker']
            }
        },
        {
            'layout': 6,  # Content
            'title': '工作流程',
            'content': [
                '1. 用户输入需求',
                '2. AI 分析架构',
                '3. 自动生成图表',
                '4. 实时预览',
                '5. 导出保存'
            ]
        }
    ],
    'thanks': {
        'text': 'Thanks！',
        'contact': 'tech@example.com'
    }
}

# 生成函数
def generate_from_data(template_path, data, output_path):
    prs = Presentation(template_path)
    
    # 标题页
    if 'title' in data:
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.placeholders[12].text = data['title']['main']
        slide.placeholders[13].text = data['title']['sub']
    
    # 内容页
    for slide_data in data.get('slides', []):
        layout_idx = slide_data['layout']
        slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
        
        # 设置标题
        if 'title' in slide_data:
            slide.placeholders[0].text = slide_data['title']
        
        # 根据版式类型填充内容
        if layout_idx == 8:  # Three Content
            content = slide_data['content']
            slide.placeholders[10].text = '\n'.join(content['left'])
            slide.placeholders[11].text = '\n'.join(content['middle'])
            slide.placeholders[12].text = '\n'.join(content['right'])
        elif layout_idx == 6:  # Content
            slide.placeholders[10].text = '\n'.join(slide_data['content'])
    
    # 结束页
    if 'thanks' in data:
        slide = prs.slides.add_slide(prs.slide_layouts[15])
        slide.placeholders[0].text = data['thanks']['text']
        if len(slide.placeholders) > 10:
            slide.placeholders[10].text = data['thanks']['contact']
    
    prs.save(output_path)

# 使用
generate_from_data(TEMPLATE, presentation_data, OUTPUT)
```

## 模式 3：JSON 配置文件

适用于：配置管理、团队协作

```json
{
  "template": "Deck_Template_NWCD_dark_202103.pptx",
  "output": "Kiro_MCP_Presentation.pptx",
  "metadata": {
    "title": "Kiro MCP 架构自动化",
    "subtitle": "基于 Model Context Protocol 的智能架构生成",
    "author": "技术团队",
    "date": "2026年2月"
  },
  "slides": [
    {
      "type": "title",
      "content": {
        "title": "Kiro MCP 架构自动化",
        "subtitle": "基于 Model Context Protocol 的智能架构生成",
        "presenter": "技术团队",
        "date": "2026年2月"
      }
    },
    {
      "type": "three_content",
      "content": {
        "title": "技术栈",
        "columns": [
          {
            "header": "前端",
            "items": ["React", "TypeScript", "Vite"]
          },
          {
            "header": "后端",
            "items": ["Python", "FastAPI", "MCP SDK"]
          },
          {
            "header": "基础设施",
            "items": ["AWS", "Docker", "GitHub Actions"]
          }
        ]
      }
    },
    {
      "type": "thanks",
      "content": {
        "text": "Thanks！",
        "contact": "tech@example.com"
      }
    }
  ]
}
```

```python
import json

def generate_from_json(json_path):
    with open(json_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    prs = Presentation(config['template'])
    
    for slide_config in config['slides']:
        slide_type = slide_config['type']
        content = slide_config['content']
        
        if slide_type == 'title':
            create_title_slide(prs, **content)
        elif slide_type == 'three_content':
            create_three_content_slide(prs, **content)
        elif slide_type == 'thanks':
            create_thanks_slide(prs, **content)
    
    prs.save(config['output'])

# 使用
generate_from_json('presentation_config.json')
```

## 模式 4：命令行工具

适用于：自动化脚本、CI/CD 集成

```python
import argparse

def main():
    parser = argparse.ArgumentParser(description='生成 NWCD 风格 PPT')
    parser.add_argument('--template', required=True, help='模板文件路径')
    parser.add_argument('--config', required=True, help='配置文件路径')
    parser.add_argument('--output', required=True, help='输出文件路径')
    parser.add_argument('--verbose', action='store_true', help='详细输出')
    
    args = parser.parse_args()
    
    if args.verbose:
        print(f"加载模板: {args.template}")
        print(f"读取配置: {args.config}")
    
    generate_from_json(args.config)
    
    if args.verbose:
        print(f"✅ 生成成功: {args.output}")

if __name__ == '__main__':
    main()
```

使用方式：
```bash
python generate_ppt.py \
  --template Deck_Template_NWCD_dark_202103.pptx \
  --config config.json \
  --output output.pptx \
  --verbose
```

## 模式 5：Web API 服务

适用于：在线生成、微服务架构

```python
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import List, Dict, Any
import tempfile
import os

app = FastAPI()

class SlideContent(BaseModel):
    type: str
    content: Dict[str, Any]

class PresentationRequest(BaseModel):
    template: str
    slides: List[SlideContent]
    output_name: str

@app.post("/generate")
async def generate_presentation(request: PresentationRequest):
    try:
        # 生成临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp:
            output_path = tmp.name
        
        # 生成 PPT
        prs = Presentation(request.template)
        
        for slide_data in request.slides:
            if slide_data.type == 'title':
                create_title_slide(prs, **slide_data.content)
            elif slide_data.type == 'three_content':
                create_three_content_slide(prs, **slide_data.content)
        
        prs.save(output_path)
        
        return {
            "status": "success",
            "file_path": output_path,
            "message": "PPT 生成成功"
        }
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 运行: uvicorn api:app --reload
```

## 模式 6：批量生成

适用于：多个演示文稿、A/B 测试

```python
def batch_generate(template_path, configs_dir, output_dir):
    """批量生成多个演示文稿"""
    import glob
    import os
    
    config_files = glob.glob(os.path.join(configs_dir, '*.json'))
    
    for config_file in config_files:
        print(f"处理: {config_file}")
        
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        output_name = os.path.basename(config_file).replace('.json', '.pptx')
        output_path = os.path.join(output_dir, output_name)
        
        try:
            generate_from_data(template_path, config, output_path)
            print(f"✅ 成功: {output_path}")
        except Exception as e:
            print(f"❌ 失败: {config_file} - {e}")

# 使用
batch_generate(
    'Deck_Template_NWCD_dark_202103.pptx',
    'configs/',
    'outputs/'
)
```

## 模式 7：模板继承

适用于：多个相似演示、品牌一致性

```python
class PPTGenerator:
    """PPT 生成器基类"""
    
    def __init__(self, template_path):
        self.template_path = template_path
        self.prs = None
        self.colors = {
            'orange': RGBColor(255, 189, 80),
            'cyan': RGBColor(0, 217, 255),
            'white': RGBColor(255, 255, 255)
        }
    
    def create(self):
        """创建演示文稿"""
        self.prs = Presentation(self.template_path)
        self.add_title_slide()
        self.add_content_slides()
        self.add_thanks_slide()
        return self.prs
    
    def add_title_slide(self):
        """添加标题页 - 子类实现"""
        raise NotImplementedError
    
    def add_content_slides(self):
        """添加内容页 - 子类实现"""
        raise NotImplementedError
    
    def add_thanks_slide(self):
        """添加结束页"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[15])
        slide.placeholders[0].text = "Thanks！"
    
    def save(self, output_path):
        """保存文件"""
        self.prs.save(output_path)

class TechPresentationGenerator(PPTGenerator):
    """技术演示生成器"""
    
    def __init__(self, template_path, title, tech_stack):
        super().__init__(template_path)
        self.title = title
        self.tech_stack = tech_stack
    
    def add_title_slide(self):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[0])
        slide.placeholders[12].text = self.title
    
    def add_content_slides(self):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[8])
        slide.placeholders[0].text = "技术栈"
        slide.placeholders[10].text = '\n'.join(self.tech_stack['frontend'])
        slide.placeholders[11].text = '\n'.join(self.tech_stack['backend'])
        slide.placeholders[12].text = '\n'.join(self.tech_stack['infra'])

# 使用
generator = TechPresentationGenerator(
    'template.pptx',
    'Kiro MCP',
    {
        'frontend': ['React', 'TypeScript'],
        'backend': ['Python', 'FastAPI'],
        'infra': ['AWS', 'Docker']
    }
)
generator.create()
generator.save('output.pptx')
```

## 模式 8：增量更新

适用于：修改现有演示、版本迭代

```python
def update_existing_presentation(pptx_path, updates):
    """更新现有演示文稿"""
    prs = Presentation(pptx_path)
    
    for update in updates:
        slide_idx = update['slide_index']
        placeholder_idx = update['placeholder_index']
        new_text = update['new_text']
        
        slide = prs.slides[slide_idx]
        slide.placeholders[placeholder_idx].text = new_text
    
    prs.save(pptx_path)

# 使用
updates = [
    {'slide_index': 0, 'placeholder_index': 12, 'new_text': '新标题'},
    {'slide_index': 1, 'placeholder_index': 0, 'new_text': '更新的内容'}
]
update_existing_presentation('existing.pptx', updates)
```

## 模式 9：多语言支持

适用于：国际化、多地区

```python
translations = {
    'zh': {
        'title': 'Kiro MCP 架构自动化',
        'thanks': '谢谢！',
        'contact': '联系方式'
    },
    'en': {
        'title': 'Kiro MCP Architecture Automation',
        'thanks': 'Thanks！',
        'contact': 'Contact'
    }
}

def generate_multilingual(template_path, content, lang='zh'):
    """生成多语言演示"""
    t = translations[lang]
    
    prs = Presentation(template_path)
    
    # 标题页
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.placeholders[12].text = t['title']
    
    # 结束页
    slide = prs.slides.add_slide(prs.slide_layouts[15])
    slide.placeholders[0].text = t['thanks']
    
    prs.save(f'output_{lang}.pptx')

# 生成中英文版本
generate_multilingual(TEMPLATE, content, 'zh')
generate_multilingual(TEMPLATE, content, 'en')
```

## 模式 10：测试驱动

适用于：质量保证、自动化测试

```python
import unittest

class TestPPTGeneration(unittest.TestCase):
    def setUp(self):
        self.template = 'Deck_Template_NWCD_dark_202103.pptx'
        self.output = 'test_output.pptx'
    
    def test_create_title_slide(self):
        prs = Presentation(self.template)
        slide = create_title_slide(prs, "测试", "测试副标题")
        self.assertEqual(len(prs.slides), 1)
    
    def test_slide_count(self):
        prs = Presentation(self.template)
        create_title_slide(prs, "标题", "副标题")
        create_three_content_slide(prs, "内容", [], [], [])
        create_thanks_slide(prs)
        self.assertEqual(len(prs.slides), 3)
    
    def tearDown(self):
        if os.path.exists(self.output):
            os.remove(self.output)

if __name__ == '__main__':
    unittest.main()
```
