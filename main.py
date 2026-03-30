"""
MD → DOCX 转换工具
应用入口
"""

import sys
from PyQt6.QtWidgets import QApplication
from ui.main_window import MainWindow


def main():
    app = QApplication(sys.argv)

    # 设置应用信息
    app.setApplicationName("MD2DOCX")
    app.setApplicationDisplayName("MD → DOCX 转换工具")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == '__main__':
    main()
