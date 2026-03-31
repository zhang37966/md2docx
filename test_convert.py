"""
CLI 转换测试脚本
用于验证转换引擎在无 GUI 环境下的正确性。
"""

import sys
import os

# 确保 Skill 目录（核心模块所在地）在 sys.path 中
SKILL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         '.agents', 'skills', 'md2docx-converter')
sys.path.insert(0, SKILL_DIR)

from converter.md_parser import parse_markdown, read_md_file
from converter.docx_writer import DocxWriter


def test_convert(md_path: str, style: str = 'A'):
    """测试转换功能。"""
    print(f"=== 测试样式 {style} ===")
    print(f"输入文件: {md_path}")

    # 读取 MD
    md_text = read_md_file(md_path)
    print(f"MD 文件大小: {len(md_text)} 字符")

    # 解析为 AST
    ast_nodes = parse_markdown(md_text)
    print(f"AST 节点数: {len(ast_nodes)}")

    # 打印节点类型概览
    for i, node in enumerate(ast_nodes):
        node_type = node.get('type', 'unknown')
        if node_type == 'heading':
            level = node.get('attrs', {}).get('level', '?')
            from converter.md_parser import extract_text_from_children
            text = extract_text_from_children(node.get('children', []))
            print(f"  [{i}] {node_type} (level={level}): {text}")
        else:
            print(f"  [{i}] {node_type}")

    # 生成 DOCX
    output_name = os.path.splitext(os.path.basename(md_path))[0]
    output_path = os.path.join(os.path.dirname(md_path), f'{output_name}_style{style}.docx')

    writer = DocxWriter(style_mode=style)
    writer.convert(ast_nodes, output_path)
    print(f"输出文件: {output_path}")
    print(f"文件大小: {os.path.getsize(output_path)} bytes")
    print()


if __name__ == '__main__':
    md_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_sample.md')

    if len(sys.argv) > 1:
        md_file = sys.argv[1]

    test_convert(md_file, 'A')
    test_convert(md_file, 'B')
    print("✅ 所有测试完成！")
