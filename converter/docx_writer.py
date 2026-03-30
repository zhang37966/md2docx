"""
DOCX 文档生成模块
接收 mistune AST 中间结构，根据样式模式生成格式化的 Word 文档。
"""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

from styles.docx_styles import (
    get_heading_map,
    FONT_CONFIG,
    NORMAL_LINE_SPACING,
    NORMAL_SPACE_BEFORE,
    NORMAL_SPACE_AFTER,
    BODY_LINE_SPACING,
    BODY_SPACE_BEFORE_LINES,
    BODY_SPACE_AFTER,
    FIRST_LINE_INDENT,
    HEADING_LINE_SPACING,
    HEADING_SPACE_BEFORE_LINES,
    HEADING_SPACE_AFTER_LINES,
    HEADING_CENTER_LEVELS,
    get_heading_font_sizes,
    FOOTER_PAGE_NUMBER,
    FOOTER_FONT_NAME,
    FOOTER_FONT_SIZE,
    CODE_BLOCK_BG_COLOR,
)
from converter.md_parser import extract_text_from_children
from docx.enum.style import WD_STYLE_TYPE


class DocxWriter:
    """Word 文档生成器"""

    def __init__(self, style_mode: str = 'A'):
        """
        Args:
            style_mode: 'A'（标题为Title）或 'B'（标题为Heading1）
        """
        self.style_mode = style_mode
        self.heading_map = get_heading_map(style_mode)
        self.heading_font_sizes = get_heading_font_sizes(style_mode)
        self.doc = Document()
        self._setup_default_styles()

    def _setup_default_styles(self):
        """设置文档默认样式。"""
        # 1. 基础 "正文" (Normal) 样式：配置字体
        normal_style = self.doc.styles['Normal']
        font = normal_style.font
        font.name = FONT_CONFIG['body']['name_en']
        font.size = FONT_CONFIG['body']['size']
        normal_style.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_CONFIG['body']['name_cn'])
        normal_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        normal_style.paragraph_format.line_spacing = NORMAL_LINE_SPACING
        normal_style.paragraph_format.space_before = NORMAL_SPACE_BEFORE
        normal_style.paragraph_format.space_after = NORMAL_SPACE_AFTER

        # 2. "正文文本" (Body Text) 样式：应用于具体段落
        # 配置行距、段前段后和首行缩进
        body_style = self.doc.styles['Body Text']
        pf = body_style.paragraph_format
        pf.line_spacing = BODY_LINE_SPACING
        pf.space_after = BODY_SPACE_AFTER
        pf.first_line_indent = FIRST_LINE_INDENT
        
        # 赋予正文文本严格的“段前 0.5 行” 选项而不仅是基于像素高度估算
        pPr = body_style._element.get_or_add_pPr()
        spacing = pPr.get_or_add_spacing()
        spacing.set(qn('w:beforeLines'), str(BODY_SPACE_BEFORE_LINES))
        if spacing.get(qn('w:before')) is not None:
            del spacing.attrib[qn('w:before')]
            
        # 设置正文两端对齐
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # 修改全局标题颜色（转为自动/黑色）并约束为统一种类字体，设定段前段后行距
        title_cfg = FONT_CONFIG.get('title', {})
        # 获取当前样式模式下一级标题对应的 Word 样式名
        first_heading_style = self.heading_map.get(1, 'Title')
        
        for s in self.doc.styles:
            if s.type == WD_STYLE_TYPE.PARAGRAPH and (s.name.startswith('Heading') or s.name == 'Title'):
                pf = s.paragraph_format
                pf.line_spacing = HEADING_LINE_SPACING
                
                # 判断是否是当前模式的一级标题（# 标题）
                is_first_heading = (s.name == first_heading_style)
                
                if is_first_heading:
                    # 一级标题专属配置：直接修改样式定义
                    s.font.color.rgb = RGBColor(0, 0, 0)
                    s.font.name = title_cfg.get('name_en', FONT_CONFIG['heading']['name_en'])
                    s.font.bold = title_cfg.get('bold', False)
                    if s.font.element.rPr is not None:
                        s.font.element.rPr.rFonts.set(qn('w:eastAsia'), title_cfg.get('name_cn', FONT_CONFIG['heading']['name_cn']))
                    
                    # 段前段后
                    if 'space_before' in title_cfg:
                        pf.space_before = title_cfg['space_before']
                    if 'space_after' in title_cfg:
                        pf.space_after = title_cfg['space_after']
                    
                    # 去除边框
                    if title_cfg.get('no_border', False):
                        pPr_title = s._element.get_or_add_pPr()
                        existing_bdr = pPr_title.find(qn('w:pBdr'))
                        if existing_bdr is not None:
                            pPr_title.remove(existing_bdr)
                        pBdr = parse_xml(
                            f'<w:pBdr {nsdecls("w")}>'
                            f'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                            f'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                            f'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                            f'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                            f'</w:pBdr>'
                        )
                        pPr_title.append(pBdr)
                    
                    # 剔除基于行的间距（因为一级标题用的是绝对磅值）
                    pPr_title = s._element.get_or_add_pPr()
                    spacing = pPr_title.get_or_add_spacing()
                    for attr_name in ('w:beforeLines', 'w:afterLines'):
                        if spacing.get(qn(attr_name)) is not None:
                            del spacing.attrib[qn(attr_name)]
                else:
                    # 普通标题样式（宋体）
                    s.font.color.rgb = RGBColor(0, 0, 0)
                    s.font.name = FONT_CONFIG['heading']['name_en']
                    if s.font.element.rPr is not None:
                        s.font.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_CONFIG['heading']['name_cn'])
                    
                    # 设置标题段前 1 行，段后 0.5 行
                    pPr = s._element.get_or_add_pPr()
                    spacing = pPr.get_or_add_spacing()
                    spacing.set(qn('w:beforeLines'), str(HEADING_SPACE_BEFORE_LINES))
                    spacing.set(qn('w:afterLines'), str(HEADING_SPACE_AFTER_LINES))
                    # 剔除可能干扰行距计算的硬编码磅值
                    for attr_name in ('w:before', 'w:after'):
                        if spacing.get(qn(attr_name)) is not None:
                            del spacing.attrib[qn(attr_name)]

        # 遍历所有段落样式，关闭网格对齐和自动调整右缩进
        for s in self.doc.styles:
            if s.type == WD_STYLE_TYPE.PARAGRAPH:
                pPr = s._element.get_or_add_pPr()
                # 移除旧的属性（如果存在）
                for tag in ('w:snapToGrid', 'w:adjustRightInd'):
                    existing = pPr.find(qn(tag))
                    if existing is not None:
                        pPr.remove(existing)
                # 强制关闭：如果定义了文档网格，则与网格对齐(W) & 自动调整右缩进(D)
                pPr.append(parse_xml(f'<w:snapToGrid {nsdecls("w")} w:val="0"/>'))
                pPr.append(parse_xml(f'<w:adjustRightInd {nsdecls("w")} w:val="0"/>'))

    def convert(self, ast_nodes: list, output_path: str):
        """将 AST 节点列表转换为 DOCX 并保存。

        Args:
            ast_nodes: mistune AST 节点列表
            output_path: 输出文件路径
        """
        for node in ast_nodes:
            self._process_node(node)

        # 设置页脚页码
        if FOOTER_PAGE_NUMBER:
            self._setup_footer()

        # 确保输出目录存在
        os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
        self.doc.save(output_path)

    def _process_node(self, node: dict, list_level: int = 0):
        """处理单个 AST 节点。

        Args:
            node: AST 节点字典
            list_level: 列表嵌套层级
        """
        node_type = node.get('type', '')

        if node_type == 'heading':
            self._add_heading(node)
        elif node_type in ('paragraph', 'block_text'):
            self._add_paragraph(node)
        elif node_type == 'list':
            self._add_list(node, list_level)
        elif node_type == 'block_code':
            self._add_code_block(node)
        elif node_type == 'table':
            self._add_table(node)
        elif node_type == 'thematic_break':
            self._add_horizontal_rule()
        elif node_type == 'block_quote':
            self._add_block_quote(node)
        elif node_type == 'blank_line':
            pass  # 忽略空行
        else:
            # 对于未知类型，尝试提取文本
            if 'children' in node:
                for child in node['children']:
                    if isinstance(child, dict):
                        self._process_node(child, list_level)

    def _add_heading(self, node: dict):
        """添加标题。"""
        level = node.get('attrs', {}).get('level', 1)
        children = node.get('children', [])
        text = extract_text_from_children(children)

        style_name = self.heading_map.get(level, f'Heading {min(level, 9)}')

        paragraph = self.doc.add_paragraph(style=style_name)
        self._add_inline_content(paragraph, children)

        # 根据配置决定是否居中对齐
        if level in HEADING_CENTER_LEVELS:
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 设置标题字号和字体（覆盖 _add_inline_content 中 _apply_body_font 设置的正文字体）
        font_size = self.heading_font_sizes.get(style_name)
        title_cfg = FONT_CONFIG.get('title', {})
        is_first = (level == 1)
        
        for run in paragraph.runs:
            if font_size:
                run.font.size = font_size
            run.font.color.rgb = RGBColor(0, 0, 0)
            
            if is_first:
                # 一级标题：黑体
                run.font.name = title_cfg.get('name_en', FONT_CONFIG['heading']['name_en'])
                run.font.element.rPr.rFonts.set(
                    qn('w:eastAsia'), title_cfg.get('name_cn', FONT_CONFIG['heading']['name_cn'])
                )
                if title_cfg.get('bold', False):
                    run.font.bold = True
            else:
                # 其余标题：宋体
                run.font.name = FONT_CONFIG['heading']['name_en']
                run.font.element.rPr.rFonts.set(
                    qn('w:eastAsia'), FONT_CONFIG['heading']['name_cn']
                )

    def _add_paragraph(self, node: dict):
        """添加段落。检测软/硬回车并拆分为多个独立的 Word 段落"""
        children = node.get('children', [])
        
        current_group = []
        groups = [current_group]
        
        for child in children:
            if isinstance(child, dict) and child.get('type') in ('softbreak', 'linebreak'):
                current_group = []
                groups.append(current_group)
            elif isinstance(child, dict) and child.get('type') == 'text' and '\n' in (child.get('raw', '') or child.get('text', '')):
                # 兼容文本中直接包含换行符的情况
                text = child.get('raw', '') or child.get('text', '')
                parts = text.split('\n')
                for i, part in enumerate(parts):
                    if i > 0:
                        current_group = []
                        groups.append(current_group)
                    if part:
                        # 构造一个新的纯文本节点推入组内
                        current_group.append({'type': 'text', 'raw': part})
            else:
                current_group.append(child)
                
        for group in groups:
            if group:  # 只生成非空的段落
                paragraph = self.doc.add_paragraph(style='Body Text')
                self._add_inline_content(paragraph, group)

    def _add_inline_content(self, paragraph, children: list):
        """向段落中添加行内内容（文本、加粗、斜体、代码等）。"""
        if not children:
            return

        for child in children:
            if isinstance(child, str):
                run = paragraph.add_run(child)
                self._apply_body_font(run)
            elif isinstance(child, dict):
                child_type = child.get('type', '')

                if child_type == 'text':
                    raw = child.get('raw', '') or child.get('text', '')
                    run = paragraph.add_run(raw)
                    self._apply_body_font(run)

                elif child_type == 'strong':
                    inner_children = child.get('children', [])
                    text = extract_text_from_children(inner_children)
                    run = paragraph.add_run(text)
                    run.bold = True
                    self._apply_body_font(run)

                elif child_type == 'emphasis':
                    inner_children = child.get('children', [])
                    text = extract_text_from_children(inner_children)
                    run = paragraph.add_run(text)
                    run.italic = True
                    self._apply_body_font(run)

                elif child_type == 'codespan':
                    raw = child.get('raw', '') or child.get('text', '')
                    run = paragraph.add_run(raw)
                    run.font.name = FONT_CONFIG['code']['name_en']
                    run.font.size = FONT_CONFIG['code']['size']
                    # 给行内代码加灰色背景
                    shd = parse_xml(
                        f'<w:shd {nsdecls("w")} w:fill="F0F0F0" w:val="clear"/>'
                    )
                    run.element.rPr.append(shd)

                elif child_type == 'link':
                    inner_children = child.get('children', [])
                    text = extract_text_from_children(inner_children)
                    link = child.get('link', '') or child.get('attrs', {}).get('url', '')
                    run = paragraph.add_run(text)
                    run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
                    run.underline = True
                    self._apply_body_font(run)

                elif child_type == 'softbreak':
                    run = paragraph.add_run('\n')

                elif child_type == 'linebreak':
                    run = paragraph.add_run('\n')

                elif child_type == 'image':
                    alt = child.get('alt', '') or child.get('attrs', {}).get('alt', '')
                    src = child.get('src', '') or child.get('attrs', {}).get('url', '')
                    # 图片引入——如果是本地路径且存在
                    if src and os.path.isfile(src):
                        try:
                            paragraph.add_run().add_picture(src, width=Inches(5))
                        except Exception:
                            run = paragraph.add_run(f'[图片: {alt or src}]')
                            self._apply_body_font(run)
                    else:
                        run = paragraph.add_run(f'[图片: {alt or src}]')
                        self._apply_body_font(run)

                elif 'children' in child:
                    self._add_inline_content(paragraph, child['children'])

    def _add_list(self, node: dict, level: int = 0):
        """添加列表（有序/无序）。"""
        ordered = node.get('attrs', {}).get('ordered', False)
        children = node.get('children', [])
        counter = node.get('attrs', {}).get('start', 1) or 1

        for item in children:
            if not isinstance(item, dict):
                continue
            item_type = item.get('type', '')
            if item_type == 'list_item':
                item_children = item.get('children', [])
                for sub_node in item_children:
                    if not isinstance(sub_node, dict):
                        continue
                    if sub_node.get('type') in ('paragraph', 'block_text'):
                        # 生成列表前缀
                        if ordered:
                            prefix = f'{counter}. '
                            counter += 1
                        else:
                            prefix = '• '

                        indent = '    ' * level
                        paragraph = self.doc.add_paragraph()
                        # 设置缩进
                        paragraph.paragraph_format.left_indent = Cm(1.27 * (level + 1))

                        run = paragraph.add_run(indent + prefix)
                        self._apply_body_font(run)
                        self._add_inline_content(paragraph, sub_node.get('children', []))

                    elif sub_node.get('type') == 'list':
                        # 嵌套列表
                        self._add_list(sub_node, level + 1)

    def _add_code_block(self, node: dict):
        """添加代码块。"""
        raw = node.get('raw', '') or node.get('text', '')
        # 移除末尾多余的换行
        raw = raw.rstrip('\n')

        paragraph = self.doc.add_paragraph()

        # 设置代码块背景
        pPr = paragraph._element.get_or_add_pPr()
        shd = parse_xml(
            f'<w:shd {nsdecls("w")} w:fill="F5F5F5" w:val="clear"/>'
        )
        pPr.append(shd)

        # 设置边框
        pBdr = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:top w:val="single" w:sz="4" w:space="4" w:color="DDDDDD"/>'
            f'  <w:left w:val="single" w:sz="4" w:space="4" w:color="DDDDDD"/>'
            f'  <w:bottom w:val="single" w:sz="4" w:space="4" w:color="DDDDDD"/>'
            f'  <w:right w:val="single" w:sz="4" w:space="4" w:color="DDDDDD"/>'
            f'</w:pBdr>'
        )
        pPr.append(pBdr)

        # 设置段落间距
        paragraph.paragraph_format.space_before = Pt(6)
        paragraph.paragraph_format.space_after = Pt(6)

        run = paragraph.add_run(raw)
        run.font.name = FONT_CONFIG['code']['name_en']
        run.font.size = FONT_CONFIG['code']['size']

    def _add_table(self, node: dict):
        """添加表格。"""
        children = node.get('children', [])
        if not children:
            return

        # 收集所有行数据：每行是一个 cell 列表
        all_rows = []
        head_row_count = 0

        for child in children:
            if not isinstance(child, dict):
                continue
            if child.get('type') == 'table_head':
                # table_head 直接包含 table_cell（它本身就是表头行）
                cells = child.get('children', [])
                if cells:
                    all_rows.append(cells)
                    head_row_count = 1
            elif child.get('type') == 'table_body':
                # table_body 包含 table_row，每个 table_row 包含 table_cell
                for row_node in child.get('children', []):
                    if isinstance(row_node, dict) and row_node.get('type') == 'table_row':
                        all_rows.append(row_node.get('children', []))

        if not all_rows:
            return

        # 计算列数
        cols = max(len(row) for row in all_rows) if all_rows else 0
        if cols == 0:
            return

        # 创建表格
        table = self.doc.add_table(rows=len(all_rows), cols=cols, style='Table Grid')

        for row_idx, row_cells in enumerate(all_rows):
            for col_idx, cell_node in enumerate(row_cells):
                if col_idx < cols and isinstance(cell_node, dict):
                    cell = table.cell(row_idx, col_idx)
                    cell_text = extract_text_from_children(cell_node.get('children', []))
                    cell.text = cell_text

                    # 表头加粗
                    if row_idx < head_row_count:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.bold = True

    def _add_horizontal_rule(self):
        """添加分隔线。"""
        paragraph = self.doc.add_paragraph()
        pPr = paragraph._element.get_or_add_pPr()
        pBdr = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:bottom w:val="single" w:sz="12" w:space="1" w:color="AAAAAA"/>'
            f'</w:pBdr>'
        )
        pPr.append(pBdr)

    def _add_block_quote(self, node: dict):
        """添加引用说明块。左侧带有竖线，底色为淡蓝色"""
        children = node.get('children', [])
        for child in children:
            if isinstance(child, dict):
                if child.get('type') in ('paragraph', 'block_text'):
                    paragraph = self.doc.add_paragraph()
                    # 添加左侧蓝色边框
                    pPr = paragraph._element.get_or_add_pPr()
                    pBdr = parse_xml(
                        f'<w:pBdr {nsdecls("w")}>'
                        f'  <w:left w:val="single" w:sz="18" w:space="8" w:color="4A90D9"/>'
                        f'</w:pBdr>'
                    )
                    pPr.append(pBdr)

                    # 设置灰色背景
                    shd = parse_xml(
                        f'<w:shd {nsdecls("w")} w:fill="F0F5FA" w:val="clear"/>'
                    )
                    pPr.append(shd)

                    # 设置缩进
                    paragraph.paragraph_format.left_indent = Cm(1)

                    self._add_inline_content(paragraph, child.get('children', []))
                else:
                    self._process_node(child)

    def _apply_body_font(self, run):
        """给 run 应用正文字体。"""
        run.font.name = FONT_CONFIG['body']['name_en']
        run.font.size = FONT_CONFIG['body']['size']
        run.font.element.rPr.rFonts.set(
            qn('w:eastAsia'), FONT_CONFIG['body']['name_cn']
        )

    def _setup_footer(self):
        """设置页脚页码：第 X 页 共 Y 页（居中）。"""
        from docx.oxml import OxmlElement

        section = self.doc.sections[0]
        footer = section.footer
        footer.is_linked_to_previous = False

        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        def add_run_text(text):
            run = paragraph.add_run(text)
            run.font.name = FOOTER_FONT_NAME
            run.font.size = FOOTER_FONT_SIZE
            run.font.element.rPr.rFonts.set(qn('w:eastAsia'), FOOTER_FONT_NAME)
            return run

        def add_field(field_code):
            """Insert a Word field code (e.g. PAGE, NUMPAGES)."""
            run = paragraph.add_run()
            run.font.name = FOOTER_FONT_NAME
            run.font.size = FOOTER_FONT_SIZE
            run.font.element.rPr.rFonts.set(qn('w:eastAsia'), FOOTER_FONT_NAME)

            fldChar_begin = OxmlElement('w:fldChar')
            fldChar_begin.set(qn('w:fldCharType'), 'begin')
            run.element.append(fldChar_begin)

            instrText = OxmlElement('w:instrText')
            instrText.set(qn('xml:space'), 'preserve')
            instrText.text = f' {field_code} '
            run.element.append(instrText)

            fldChar_separate = OxmlElement('w:fldChar')
            fldChar_separate.set(qn('w:fldCharType'), 'separate')
            run.element.append(fldChar_separate)

            fldChar_end = OxmlElement('w:fldChar')
            fldChar_end.set(qn('w:fldCharType'), 'end')
            run.element.append(fldChar_end)

        # “第 X 页 共 Y 页”
        add_run_text('第 ')
        add_field('PAGE')
        add_run_text(' 页 共 ')
        add_field('NUMPAGES')
        add_run_text(' 页')
