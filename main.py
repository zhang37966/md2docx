"""
MD → DOCX 转换工具
应用入口
"""

import sys
import os

# 核心转换模块已迁移至 Skill 目录，将其加入模块搜索路径
SKILL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         '.agents', 'skills', 'md2docx-converter')
sys.path.insert(0, SKILL_DIR)

from PyQt6.QtWidgets import QApplication
from ui.main_window import MainWindow
from utils.font_installer import check_and_install_font

def main():
    # 提前检查并安装可能缺少的字体
    check_and_install_font()
    
    app = QApplication(sys.argv)

    # 设置应用信息
    app.setApplicationName("MD2DOCX")
    app.setApplicationDisplayName("MD → DOCX 转换工具")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == '__main__':
    main()
