"""
MD2DOCX CLI 转换脚本
供 Antigravity Skill 调用，不依赖 GUI。

用法:
    python convert.py --input input.md --output output.docx [--style A|B]
    python convert.py --text "# 标题\n正文" --output output.docx [--style A|B]
    echo "# 标题" | python convert.py --stdin --output output.docx [--style A|B]
"""

import argparse
import sys
import os

# 将 Skill 根目录加入 sys.path，使 converter/styles/utils 可正常导入
SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, SKILL_ROOT)

from converter.md_parser import parse_markdown, read_md_file
from converter.docx_writer import DocxWriter
from utils.font_installer import check_and_install_font


def main():
    parser = argparse.ArgumentParser(
        description='MD2DOCX CLI - 将 Markdown 转换为格式化的 Word 文档'
    )

    # 输入源（三选一）
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument(
        '--input', '-i',
        help='输入的 Markdown 文件路径'
    )
    input_group.add_argument(
        '--text', '-t',
        help='直接传入的 Markdown 文本内容'
    )
    input_group.add_argument(
        '--stdin',
        action='store_true',
        help='从标准输入读取 Markdown 内容'
    )

    # 输出路径（必填）
    parser.add_argument(
        '--output', '-o',
        required=True,
        help='输出的 .docx 文件路径'
    )

    # 样式模式
    parser.add_argument(
        '--style', '-s',
        choices=['A', 'B'],
        default='A',
        help='转换样式模式：A=Title模式（默认），B=Heading模式'
    )

    # 跳过字体安装（用于已确认字体存在的场景）
    parser.add_argument(
        '--skip-font-check',
        action='store_true',
        help='跳过字体检测与自动安装'
    )

    args = parser.parse_args()

    # 1. 字体检查
    if not args.skip_font_check:
        try:
            check_and_install_font()
        except Exception as e:
            print(f"警告: 字体检查失败 ({e})，继续转换...", file=sys.stderr)

    # 2. 获取 Markdown 内容
    try:
        if args.input:
            if not os.path.isfile(args.input):
                print(f"错误: 输入文件不存在: {args.input}", file=sys.stderr)
                sys.exit(1)
            md_text = read_md_file(args.input)
        elif args.text:
            md_text = args.text
        elif args.stdin:
            md_text = sys.stdin.read()
        
        if not md_text or not md_text.strip():
            print("错误: Markdown 内容为空", file=sys.stderr)
            sys.exit(1)
    except Exception as e:
        print(f"错误: 读取输入失败: {e}", file=sys.stderr)
        sys.exit(1)

    # 3. 解析并转换
    try:
        ast_nodes = parse_markdown(md_text)
        writer = DocxWriter(style_mode=args.style)
        writer.convert(ast_nodes, args.output)
    except PermissionError:
        print(f"错误: 无法写入文件 '{args.output}'，可能被其他程序占用。请关闭后重试。", file=sys.stderr)
        sys.exit(2)
    except Exception as e:
        print(f"错误: 转换失败: {e}", file=sys.stderr)
        sys.exit(1)

    # 4. 输出结果
    abs_path = os.path.abspath(args.output)
    print(f"转换成功: {abs_path}")
    sys.exit(0)


if __name__ == '__main__':
    main()
