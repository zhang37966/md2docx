"""
样式定义模块
定义 Word 文档的字体、字号、间距等全局样式，
以及两种转换模式的标题级别映射。
"""

from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


# ============================================================
# 样式 A：文档标题作为 Title
# ============================================================
STYLE_A_HEADING_MAP = {
    1: 'Title',       # # 标题  → Title
    2: 'Heading 1',   # ## 标题 → Heading 1
    3: 'Heading 2',   # ### 标题 → Heading 2
    4: 'Heading 3',   # #### 标题 → Heading 3
    5: 'Heading 4',   # ##### 标题 → Heading 4
    6: 'Heading 5',   # ###### 标题 → Heading 5
}

# ============================================================
# 样式 B：文档标题作为 Heading 1
# ============================================================
STYLE_B_HEADING_MAP = {
    1: 'Heading 1',   # # 标题  → Heading 1
    2: 'Heading 2',   # ## 标题 → Heading 2
    3: 'Heading 3',   # ### 标题 → Heading 3
    4: 'Heading 4',   # #### 标题 → Heading 4
    5: 'Heading 5',   # ##### 标题 → Heading 5
    6: 'Heading 6',   # ###### 标题 → Heading 6
}


def get_heading_map(style_mode: str) -> dict:
    """根据样式模式返回标题映射表。

    Args:
        style_mode: 'A' 或 'B'

    Returns:
        标题级别到 Word 样式名的映射字典
    """
    if style_mode == 'A':
        return STYLE_A_HEADING_MAP
    return STYLE_B_HEADING_MAP


# ============================================================
# 字体配置
# ============================================================
FONT_CONFIG = {
    'body': {
        'name_cn': '仿宋_GB2312',
        'name_en': '仿宋_GB2312',
        'size': Pt(12),       # 小四号
    },
    'heading': {
        'name_cn': '宋体',
        'name_en': '宋体',
    },
    'code': {
        'name_cn': 'Consolas',
        'name_en': 'Consolas',
        'size': Pt(10),
    },
    # 一级标题 (#) 专用字体（样式A和B的 # 均使用）
    'title': {
        'name_cn': '黑体',
        'name_en': '黑体',
        'bold': True,
        'no_border': True,  # 去除 Title 样式默认的段落边框
        'space_before': Pt(0),   # 段前 0
        'space_after': Pt(16),   # 段后 16磅
    },
}

# ============================================================
# 页脚配置
# ============================================================
# 页码格式："第 X 页 共 Y 页"
FOOTER_PAGE_NUMBER = True
FOOTER_FONT_NAME = '宋体'
FOOTER_FONT_SIZE = Pt(9)

# ============================================================
# 段落排版配置
# ============================================================

# 1. 基础正文 (Normal) 样式配置
NORMAL_LINE_SPACING = 1.0
NORMAL_SPACE_BEFORE = Pt(0)
NORMAL_SPACE_AFTER = Pt(0)

# 2. 正文文本 (Body Text) 样式配置 (实际文章段落渲染所用的首发样式)
BODY_LINE_SPACING = 1.2
BODY_SPACE_BEFORE_LINES = 50 # 段前 0.5 行 (基于 XML beforeLines 参数，50 = 0.5行)
BODY_SPACE_AFTER = Pt(0)
FIRST_LINE_INDENT = Pt(24) # 首行缩进（2字符，12pt * 2 = 24pt）

# 3. 标题 (Heading) 样式配置
HEADING_LINE_SPACING = 1.0
HEADING_SPACE_BEFORE_LINES = 100 # 段前 1 行 (基于 XML beforeLines 参数，100 = 1行)
HEADING_SPACE_AFTER_LINES = 50   # 段后 0.5 行 (基于 XML afterLines 参数，50 = 0.5行)

# 需要居中对齐的标题等级 (对应 Markdown 的 # 等级数)
HEADING_CENTER_LEVELS = [1]  # # 一级标题居中

# 4. 列表 (List Paragraph) 样式配置
LIST_LEFT_INDENT = Cm(0.3)           # 文本之前（左缩进）0.3 厘米
LIST_HANGING_INDENT = Cm(0.7)        # 悬挂缩进 0.7 厘米
LIST_LINE_SPACING = 1.1              # 多倍行距 1.1 倍
LIST_SPACE_BEFORE_LINES = 50         # 段前 0.5 行 (基于 XML beforeLines 参数，50 = 0.5行)
LIST_SPACE_AFTER = Pt(0)             # 段后 0

# 标题字号映射
# 样式 B 的字号基准
HEADING_FONT_SIZES_B = {
    'Heading 1': Pt(18),   # 小二
    'Heading 2': Pt(16),   # 三号
    'Heading 3': Pt(14),   # 四号
    'Heading 4': Pt(12),   # 小四
    'Heading 5': Pt(11),
    'Heading 6': Pt(10.5),
}

# 样式 A 的每级标题采用样式 B 的下一级字号
# Title → B的H1(18pt)，H1 → B的H2(16pt)，H2 → B的H3(14pt) ...
HEADING_FONT_SIZES_A = {
    'Title': Pt(18),       # 对应样式B的 Heading 1 字号
    'Heading 1': Pt(16),   # 对应样式B的 Heading 2 字号
    'Heading 2': Pt(14),   # 对应样式B的 Heading 3 字号
    'Heading 3': Pt(12),   # 对应样式B的 Heading 4 字号
    'Heading 4': Pt(11),   # 对应样式B的 Heading 5 字号
    'Heading 5': Pt(10.5), # 对应样式B的 Heading 6 字号
}


def get_heading_font_sizes(style_mode: str) -> dict:
    """根据样式模式返回对应的标题字号映射表。"""
    if style_mode == 'A':
        return HEADING_FONT_SIZES_A
    return HEADING_FONT_SIZES_B

# 代码块背景色
CODE_BLOCK_BG_COLOR = RGBColor(0xF5, 0xF5, 0xF5)
CODE_BLOCK_BORDER_COLOR = RGBColor(0xDD, 0xDD, 0xDD)
