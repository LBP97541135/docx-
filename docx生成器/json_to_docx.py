#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
JSON转Docx生成器
用于将解析出的JSON数据重新生成为docx文件，包括：
- 文本内容
- 段落样式（对齐方式、行距、缩进）
- 字符样式（字体、字号、颜色、粗体/斜体）
- 表格、图片等元素
"""

import os
import json
import base64
import io
from typing import Dict, List, Any, Optional
from dataclasses import dataclass

try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor, Mm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml.shared import OxmlElement
except ImportError:
    print("需要安装python-docx库: pip install python-docx")
    exit(1)


class JsonToDocxConverter:
    """JSON转Docx转换器"""
    
    def __init__(self):
        self.doc = Document()
        self.MAX_EMU = 2147483647  # 32位有符号整数的最大值，Word中EMU的上限
    
    def safe_pt(self, pt_value: float) -> Pt:
        """安全地转换点值为Pt对象，防止溢出"""
        if pt_value is None:
            return Pt(0)
        # 1 Pt = 12700 EMU
        emu_value = pt_value * 12700
        if emu_value > self.MAX_EMU:
            return Pt(self.MAX_EMU / 12700)
        if emu_value < -self.MAX_EMU:
            return Pt(-self.MAX_EMU / 12700)
        return Pt(pt_value)

    def safe_mm(self, mm_value: float) -> Mm:
        """安全地转换毫米值为Mm对象，防止溢出"""
        if mm_value is None:
            return Mm(0)
        # 1 Mm = 36000 EMU
        emu_value = mm_value * 36000
        if emu_value > self.MAX_EMU:
            return Mm(self.MAX_EMU / 36000)
        if emu_value < 0: # 长度不应为负
            return Mm(0)
        return Mm(mm_value)

    def safe_line_spacing(self, spacing: Any) -> Any:
        """安全地处理行距"""
        if spacing is None:
            return None
        
        from docx.shared import Emu, Length
        
        # 如果已经是 Length 对象，确保其值合法
        if isinstance(spacing, Length):
            return Emu(min(max(int(spacing), -self.MAX_EMU), self.MAX_EMU))
            
        # 从 JSON 中读取的通常是 int 或 float
        if isinstance(spacing, (int, float)):
            # 区分倍数和绝对值
            # 经验法则：如果数值很大（> 1000），通常是解析出的原始 EMU 值（绝对行距）
            # 如果数值较小，通常是行距倍数（如 1.0, 1.5）
            if spacing > 1000:
                emu_value = int(min(spacing, self.MAX_EMU))
                return Emu(emu_value)
            else:
                # 限制倍数在合理范围内（0 到 50 倍）
                return float(min(max(spacing, 0), 50))
        
        return spacing

    def convert_alignment(self, alignment_str: str):
        """转换对齐方式字符串为docx枚举"""
        alignment_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "distribute": WD_ALIGN_PARAGRAPH.DISTRIBUTE
        }
        return alignment_map.get(alignment_str, WD_ALIGN_PARAGRAPH.LEFT)
    
    def set_paragraph_style(self, paragraph, style_data: Dict[str, Any]):
        """设置段落样式"""
        if not style_data:
            return
        
        paragraph_format = paragraph.paragraph_format
        
        # 设置对齐方式
        try:
            if style_data.get('alignment'):
                paragraph_format.alignment = self.convert_alignment(style_data['alignment'])
        except Exception as e:
            print(f"设置对齐方式失败: {e}, value: {style_data.get('alignment')}")
        
        # 设置行距
        try:
            if style_data.get('line_spacing'):
                paragraph_format.line_spacing = self.safe_line_spacing(style_data['line_spacing'])
        except Exception as e:
            print(f"设置行距失败: {e}, value: {style_data.get('line_spacing')}")
        
        # 设置段前间距
        try:
            if style_data.get('space_before'):
                paragraph_format.space_before = self.safe_pt(style_data['space_before'])
        except Exception as e:
            print(f"设置段前间距失败: {e}, value: {style_data.get('space_before')}")
        
        # 设置段后间距
        try:
            if style_data.get('space_after'):
                paragraph_format.space_after = self.safe_pt(style_data['space_after'])
        except Exception as e:
            print(f"设置段后间距失败: {e}, value: {style_data.get('space_after')}")
        
        # 设置左缩进
        try:
            if style_data.get('left_indent'):
                paragraph_format.left_indent = self.safe_pt(style_data['left_indent'])
        except Exception as e:
            print(f"设置左缩进失败: {e}, value: {style_data.get('left_indent')}")
        
        # 设置右缩进
        try:
            if style_data.get('right_indent'):
                paragraph_format.right_indent = self.safe_pt(style_data['right_indent'])
        except Exception as e:
            print(f"设置右缩进失败: {e}, value: {style_data.get('right_indent')}")
        
        # 设置首行缩进
        try:
            if style_data.get('first_line_indent'):
                paragraph_format.first_line_indent = self.safe_pt(style_data['first_line_indent'])
        except Exception as e:
            print(f"设置首行缩进失败: {e}, value: {style_data.get('first_line_indent')}")
    
    def set_run_style(self, run, style_data: Dict[str, Any]):
        """设置字符样式"""
        if not style_data:
            return
        
        font = run.font
        
        # 设置字体
        try:
            if style_data.get('font'):
                font_name = style_data['font']
                font.name = font_name
                # 必须设置西文和东亚字体，否则中文可能无法正确显示
                rPr = run._element.get_or_add_rPr()
                rFonts = rPr.get_or_add_rFonts()
                rFonts.set(qn('w:eastAsia'), font_name)
        except Exception as e:
            print(f"设置字体失败: {e}")
        
        # 设置字号
        try:
            if style_data.get('size'):
                font.size = self.safe_pt(style_data['size'])
        except Exception as e:
            print(f"设置字号失败: {e}, value: {style_data.get('size')}")
        
        # 设置颜色
        try:
            if style_data.get('color'):
                color_str = style_data['color'].lstrip('#')
                if len(color_str) == 6:
                    r = int(color_str[0:2], 16)
                    g = int(color_str[2:4], 16)
                    b = int(color_str[4:6], 16)
                    font.color.rgb = RGBColor(r, g, b)
        except Exception as e:
            print(f"设置颜色失败: {e}")
        
        # 设置粗体
        try:
            if style_data.get('bold') is not None:
                font.bold = style_data['bold']
        except Exception as e:
            print(f"设置粗体失败: {e}")
        
        # 设置斜体
        try:
            if style_data.get('italic') is not None:
                font.italic = style_data['italic']
        except Exception as e:
            print(f"设置斜体失败: {e}")
        
        # 设置下划线
        try:
            if style_data.get('underline') is not None:
                font.underline = style_data['underline']
        except Exception as e:
            print(f"设置下划线失败: {e}")

        # 设置上标
        try:
            if style_data.get('superscript') is not None:
                font.superscript = style_data['superscript']
        except Exception as e:
            print(f"设置上标失败: {e}")

        # 设置下标
        try:
            if style_data.get('subscript') is not None:
                font.subscript = style_data['subscript']
        except Exception as e:
            print(f"设置下标失败: {e}")

        # 设置突出显示 (Highlight)
        try:
            if style_data.get('highlight'):
                from docx.enum.text import WD_COLOR_INDEX
                highlight_map = {
                    'yellow': WD_COLOR_INDEX.YELLOW,
                    'green': WD_COLOR_INDEX.BRIGHT_GREEN,
                    'cyan': WD_COLOR_INDEX.TURQUOISE,
                    'magenta': WD_COLOR_INDEX.PINK,
                    'blue': WD_COLOR_INDEX.BLUE,
                    'red': WD_COLOR_INDEX.RED,
                    'darkBlue': WD_COLOR_INDEX.DARK_BLUE,
                    'darkCyan': WD_COLOR_INDEX.TEAL,
                    'darkGreen': WD_COLOR_INDEX.GREEN,
                    'darkMagenta': WD_COLOR_INDEX.VIOLET,
                    'darkRed': WD_COLOR_INDEX.DARK_RED,
                    'darkYellow': WD_COLOR_INDEX.DARK_YELLOW,
                    'darkGray': WD_COLOR_INDEX.GRAY_50,
                    'lightGray': WD_COLOR_INDEX.GRAY_25,
                    'black': WD_COLOR_INDEX.BLACK
                }
                color_index = highlight_map.get(style_data['highlight'])
                if color_index is not None:
                    font.highlight_color = color_index
        except Exception as e:
            print(f"设置突出显示失败: {e}")
    
    def add_paragraph(self, element_data: Dict[str, Any]):
        """添加段落"""
        paragraph = self.doc.add_paragraph()
        
        # 设置段落样式
        self.set_paragraph_style(paragraph, element_data.get('paragraph_style', {}))
        
        # 添加文本内容和样式
        content = element_data.get('content', [])
        for text_part in content:
            if text_part.get('text'):
                run = paragraph.add_run(text_part['text'])
                self.set_run_style(run, text_part.get('style', {}))
    
    def add_table(self, element_data: Dict[str, Any]):
        """添加表格，支持合并单元格还原"""
        table_data = element_data.get('content', {})
        rows_count = table_data.get('rows', 0)
        cols_count = table_data.get('columns', 0)
        cells = table_data.get('cells', [])
        
        if rows_count == 0 or cols_count == 0:
            return
        
        # 创建表格
        table = self.doc.add_table(rows=rows_count, cols=cols_count)
        table.style = 'Table Grid'  # 默认添加边框
        
        # 用于记录已经合并过的区域，避免重复合并报错
        merged_regions = set()
        
        # 填充单元格内容并处理合并
        for r, row_data in enumerate(cells):
            if r >= rows_count:
                break
            for c, cell_info in enumerate(row_data):
                if c >= cols_count or cell_info is None:
                    continue
                
                # 获取当前单元格
                current_cell = table.cell(r, c)
                
                # 处理合并单元格
                row_span = cell_info.get('row_span', 1)
                col_span = cell_info.get('col_span', 1)
                
                if (row_span > 1 or col_span > 1) and (r, c) not in merged_regions:
                    # 计算结束坐标
                    end_r = r + row_span - 1
                    end_c = c + col_span - 1
                    
                    # 确保坐标不越界
                    end_r = min(end_r, rows_count - 1)
                    end_c = min(end_c, cols_count - 1)
                    
                    # 执行合并
                    other_cell = table.cell(end_r, end_c)
                    current_cell.merge(other_cell)
                    
                    # 标记该区域已合并
                    for dr in range(row_span):
                        for dc in range(col_span):
                            merged_regions.add((r + dr, c + dc))
                
                # 设置单元格文本
                current_cell.text = cell_info.get('text', '')
    
    def add_image(self, element_data: Dict[str, Any]):
        """添加图片（优先支持本地缓存路径，插入后删除）"""
        img_data = element_data.get('content', {})
        width_mm = img_data.get('width', 0)
        height_mm = img_data.get('height', 0)
        image_path = img_data.get('image_path')
        binary_content = img_data.get('binary_content')
        alignment = img_data.get('alignment')
        
        inserted = False
        last_paragraph = None
        
        # 1. 优先尝试从本地缓存路径加载
        if image_path and os.path.exists(image_path):
            try:
                if width_mm > 0 and height_mm > 0:
                    self.doc.add_picture(image_path, width=self.safe_mm(width_mm), height=self.safe_mm(height_mm))
                else:
                    self.doc.add_picture(image_path)
                inserted = True
                last_paragraph = self.doc.paragraphs[-1]
                
            except Exception as e:
                print(f"通过路径插入图片失败: {e}")

        # 2. 如果路径方式失败，尝试使用 base64 (回退方案)
        if not inserted and binary_content:
            try:
                image_bytes = base64.b64decode(binary_content)
                image_stream = io.BytesIO(image_bytes)
                if width_mm > 0 and height_mm > 0:
                    self.doc.add_picture(image_stream, width=self.safe_mm(width_mm), height=self.safe_mm(height_mm))
                else:
                    self.doc.add_picture(image_stream)
                inserted = True
                last_paragraph = self.doc.paragraphs[-1]
            except Exception as e:
                print(f"通过Base64插入图片失败: {e}")

        # 3. 如果成功插入，设置对齐方式
        if inserted and last_paragraph and alignment:
            last_paragraph.alignment = self.convert_alignment(alignment)

        # 4. 如果都失败了，使用占位符
        if not inserted:
            alt_text = img_data.get('alt_text', '图片')
            paragraph = self.doc.add_paragraph()
            run = paragraph.add_run(f"[图片插入失败: {alt_text}]")
            run.font.italic = True
    
    def set_document_default_font(self, font_name: str = "宋体"):
        """设置文档默认字体"""
        # 1. 设置正文样式的默认字体
        try:
            style = self.doc.styles['Normal']
            style.font.name = font_name
            rPr = style.element.get_or_add_rPr()
            rFonts = rPr.get_or_add_rFonts()
            rFonts.set(qn('w:eastAsia'), font_name)
        except Exception:
            pass
        
        # 2. 也可以通过修改文档默认设置来实现更底层的控制
        try:
            rPrDefault = self.doc.styles.element.xpath('w:docDefaults/w:rPrDefault')[0]
            rPr = rPrDefault.get_or_add_rPr()
            rFonts = rPr.get_or_add_rFonts()
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:eastAsia'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
        except Exception:
            pass

    def convert(self, json_data: Dict[str, Any], output_path: str):
        """将JSON数据转换为docx文件"""
        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        # 清空文档
        self.doc = Document()
        
        # 设置默认字体为宋体
        self.set_document_default_font("宋体")
        
        # 处理文档元素
        document_elements = json_data.get('document_elements', [])
        
        for element in document_elements:
            element_type = element.get('element_type')
            
            if element_type == 'paragraph':
                self.add_paragraph(element)
            elif element_type == 'table':
                self.add_table(element)
            elif element_type == 'image':
                self.add_image(element)
        
        # 保存文档
        self.doc.save(output_path)
        return output_path


def main():
    """主函数，用于命令行使用"""
    import sys
    
    if len(sys.argv) < 3:
        print("用法: python json_to_docx.py <json_file> <output_docx>")
        return
    
    json_file = sys.argv[1]
    output_file = sys.argv[2]
    
    if not os.path.exists(json_file):
        print(f"错误: 文件 '{json_file}' 不存在")
        return
    
    try:
        # 读取JSON数据
        with open(json_file, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        
        # 转换为docx
        converter = JsonToDocxConverter()
        output_path = converter.convert(json_data, output_file)
        
        print(f"成功生成docx文件: {output_path}")
        
    except Exception as e:
        print(f"转换失败: {str(e)}")


if __name__ == "__main__":
    main()