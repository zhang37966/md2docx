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
        'name_cn': '宋体',
        'name_en': 'Times New Roman',
        'size': Pt(12),       # 小四号
    },
    'heading': {
        'name_cn': '黑体',
        'name_en': 'Arial',
    },
    'code': {
        'name_cn': 'Consolas',
        'name_en': 'Consolas',
        'size': Pt(10),
    },
}

# 行间距
LINE_SPACING = 1.5

# 段前段后间距（磅）
SPACE_BEFORE = Pt(6)
SPACE_AFTER = Pt(6)

# 标题字号映射
HEADING_FONT_SIZES = {
    'Title': Pt(22),       # 二号
    'Heading 1': Pt(18),   # 小二
    'Heading 2': Pt(16),   # 三号
    'Heading 3': Pt(14),   # 四号
    'Heading 4': Pt(12),   # 小四
    'Heading 5': Pt(11),
    'Heading 6': Pt(10.5),
}

# 代码块背景色
CODE_BLOCK_BG_COLOR = RGBColor(0xF5, 0xF5, 0xF5)
CODE_BLOCK_BORDER_COLOR = RGBColor(0xDD, 0xDD, 0xDD)
