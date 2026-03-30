"""
Markdown 解析模块
使用 mistune 的 AST 模式解析 Markdown 文件，
输出统一的中间结构供 docx_writer 消费。
"""

import re
import mistune


def parse_markdown(md_text: str) -> list:
    """将 Markdown 文本解析为 AST 节点列表。

    Args:
        md_text: Markdown 源文本

    Returns:
        mistune AST 节点列表
    """
    # 预处理：修补缺失的空行
    # 1. 确保标题上方有空行（如果上一行不是空行）
    md_text = re.sub(r'([^\n])\n(#{1,6}\s+)', r'\1\n\n\2', md_text)
    # 2. 确保标题下方有空行（如果下一行不是空行）
    md_text = re.sub(r'(^|\n)(#{1,6}\s+.*?)\n([^\n])', r'\1\2\n\n\3', md_text)

    markdown = mistune.create_markdown(
        renderer='ast',
        plugins=['table', 'strikethrough']
    )
    ast = markdown(md_text)
    return ast


def read_md_file(file_path: str) -> str:
    """读取 Markdown 文件内容。

    Args:
        file_path: 文件路径

    Returns:
        文件内容字符串
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()


def extract_text_from_children(children: list) -> str:
    """从 AST children 中递归提取纯文本。

    Args:
        children: AST 子节点列表

    Returns:
        拼接后的纯文本
    """
    text = ''
    if not children:
        return text
    for child in children:
        if isinstance(child, str):
            text += child
        elif isinstance(child, dict):
            child_type = child.get('type', '')
            if child_type == 'text':
                text += child.get('raw', '') or child.get('text', '')
            elif child_type == 'codespan':
                text += child.get('raw', '') or child.get('text', '')
            elif child_type in ('emphasis', 'strong', 'link'):
                text += extract_text_from_children(child.get('children', []))
            elif 'children' in child:
                text += extract_text_from_children(child['children'])
            elif 'raw' in child:
                text += child['raw']
            elif 'text' in child:
                text += child['text']
    return text
