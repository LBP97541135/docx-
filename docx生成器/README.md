# JSON转Docx生成器

这是一个Python程序，用于将解析出的JSON数据重新生成为docx文件，能够还原文本内容、段落样式、字符样式、表格和图片等元素。

## 功能特点

- **文本内容还原**：将JSON中的文本内容重新生成到docx文件中
- **段落样式还原**：还原对齐方式、行距、缩进等段落格式
- **字符样式还原**：还原字体、字号、颜色、粗体/斜体等字符格式
- **表格结构还原**：重建表格的结构和内容
- **图片占位符**：为图片创建占位符，包含尺寸和描述信息

## 安装依赖

```bash
pip install python-docx
```

或者使用requirements.txt安装所有依赖：

```bash
pip install -r requirements.txt
```

## 使用方法

### 基本用法

```python
from json_to_docx import JsonToDocxConverter

# 创建转换器实例
converter = JsonToDocxConverter()

# 读取JSON数据
with open('parsed_docx_result.json', 'r', encoding='utf-8') as f:
    json_data = json.load(f)

# 转换为docx文件
output_path = converter.convert(json_data, 'regenerated_document.docx')
```

### 命令行使用

```bash
python json_to_docx.py input.json output.docx
```

### 示例脚本

运行example_json_to_docx.py查看完整示例：

```bash
python example_json_to_docx.py
```

## 输入JSON格式

程序期望的JSON输入格式应与docx解析器的输出格式一致，包含以下字段：

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

## 转换说明

### 段落元素转换

程序会还原以下段落样式：
- 对齐方式（左对齐、居中、右对齐、两端对齐）
- 行距
- 段前/段后间距
- 左/右缩进
- 首行缩进

以及以下字符样式：
- 字体名称
- 字号
- 文字颜色
- 粗体
- 斜体
- 下划线

### 表格元素转换

程序会根据JSON中的表格数据重建表格结构，包括：
- 表格行数和列数
- 每个单元格的内容

### 图片元素转换

由于原始图片数据不包含在JSON中，程序会为每个图片创建一个占位符，包含：
- 图片描述（如果有）
- 图片尺寸信息

## 注意事项

1. 程序只能还原文本和格式信息，无法还原原始图片
2. 某些高级格式可能无法完全还原
3. 生成的docx文件可以使用Microsoft Word或其他兼容软件打开
4. 对于大型文档，转换可能需要一些时间

## 文件说明

- `json_to_docx.py` - 主转换器模块
- `example_json_to_docx.py` - 使用示例脚本
- `requirements.txt` - 依赖包列表
- `README.md` - 本说明文件

## 与docx解析器的配合使用

这个程序与docx解析器配合使用，可以实现以下工作流程：

1. 使用docx解析器从URL解析docx文件，生成JSON数据
2. 对JSON数据进行处理或修改（可选）
3. 使用JSON转Docx生成器将JSON数据重新生成为docx文件

这样您就可以在保留文档格式的同时，对文档内容进行程序化处理。