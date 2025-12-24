# Docx文档解析器

这是一个Python程序，用于从docx文档URL中提取结构化数据，包括文本内容、段落样式、字符样式、表格和图片等信息。

## 功能特点

- **文本内容提取**：提取文档中的所有文本内容
- **段落样式解析**：获取对齐方式、行距、缩进等段落格式
- **字符样式解析**：获取字体、字号、颜色、粗体/斜体等字符格式
- **表格数据提取**：提取表格的结构和内容
- **图片信息提取**：获取图片的尺寸、位置和描述信息

## 安装依赖

```bash
pip install python-docx requests
```

或者使用requirements.txt安装所有依赖：

```bash
pip install -r requirements.txt
```

## 使用方法

### 基本用法

```python
from docx_parser import DocxParser

# 创建解析器实例
parser = DocxParser()

# 解析文档
result = parser.parse_docx("https://example.com/document.docx")

# 处理结果
print(f"段落数: {result['metadata']['paragraphs_count']}")
print(f"表格数: {result['metadata']['tables_count']}")
print(f"图片数: {result['metadata']['images_count']}")
```

### 命令行使用

```bash
python docx_parser.py https://example.com/document.docx
```

### 示例脚本

运行example_usage.py查看完整示例：

```bash
python example_usage.py
```

## 输出结构

解析结果是一个包含以下字段的字典：

```python
{
    'document_elements': [
        {
            'element_type': 'paragraph',  # 'paragraph', 'table', 'image'
            'content': [...],              # 元素内容
            'paragraph_style': {...}       # 段落样式（仅段落元素）
        },
        # ...更多元素
    ],
    'images': [...],                      # 图片信息列表
    'metadata': {
        'paragraphs_count': int,         # 段落总数
        'tables_count': int,              # 表格总数
        'images_count': int               # 图片总数
    }
}
```

### 段落元素结构

```python
{
    'element_type': 'paragraph',
    'content': [
        {
            'text': '文本内容',
            'style': {
                'font': '字体名称',
                'size': 字体大小,
                'color': '#RRGGBB',
                'bold': True/False,
                'italic': True/False,
                'underline': True/False
            }
        },
        # ...更多文本片段
    ],
    'paragraph_style': {
        'alignment': '对齐方式',          # 'left', 'center', 'right', 'justify'
        'line_spacing': 行距,
        'space_before': 段前间距,
        'space_after': 段后间距,
        'left_indent': 左缩进,
        'right_indent': 右缩进,
        'first_line_indent': 首行缩进
    }
}
```

### 表格元素结构

```python
{
    'element_type': 'table',
    'content': {
        'rows': 行数,
        'columns': 列数,
        'cells': [
            ['单元格1', '单元格2', ...],  # 第一行
            ['单元格1', '单元格2', ...],  # 第二行
            # ...更多行
        ]
    }
}
```

### 图片元素结构

```python
{
    'element_type': 'image',
    'content': {
        'width': 宽度(mm),
        'height': 高度(mm),
        'inline': 是否内联,
        'alt_text': '图片描述'
    }
}
```

## 注意事项

1. 程序会自动下载URL指向的docx文件到临时目录，解析完成后会自动清理
2. 支持的URL必须是直接指向docx文件的链接
3. 对于大型文档，解析可能需要一些时间
4. 图片提取基于docx文件内部结构，可能无法获取所有图片信息

## 文件说明

- `docx_parser.py` - 主解析器模块
- `example_usage.py` - 使用示例脚本
- `requirements.txt` - 依赖包列表
- `README.md` - 本说明文件