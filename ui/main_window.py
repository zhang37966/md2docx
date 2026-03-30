"""
PyQt6 主窗口模块
提供文件拖拽/选择、样式切换、转换按钮等 UI 交互。
"""

import os
import sys
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QRadioButton, QButtonGroup,
    QFileDialog, QMessageBox, QFrame, QGroupBox,
    QProgressBar, QApplication,
)
from PyQt6.QtCore import Qt, QMimeData
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QFont, QIcon

from converter.md_parser import parse_markdown, read_md_file
from converter.docx_writer import DocxWriter


class DropArea(QFrame):
    """文件拖拽区域"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.file_path = None
        self.parent_window = parent

        self.setMinimumHeight(180)
        self.setStyleSheet("""
            DropArea {
                border: 2px dashed #AABBCC;
                border-radius: 12px;
                background-color: #F8FAFC;
            }
            DropArea:hover {
                border-color: #5B8DEF;
                background-color: #EEF2FF;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # 图标文字
        self.icon_label = QLabel("📄")
        self.icon_label.setFont(QFont("Segoe UI Emoji", 36))
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.icon_label)

        self.label = QLabel("将 Markdown 文件拖到这里\n或点击下方按钮选择文件")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setFont(QFont("Microsoft YaHei", 11))
        self.label.setStyleSheet("color: #666; border: none; background: transparent;")
        layout.addWidget(self.label)

        self.file_label = QLabel("")
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_label.setFont(QFont("Microsoft YaHei", 10))
        self.file_label.setStyleSheet("color: #5B8DEF; font-weight: bold; border: none; background: transparent;")
        self.file_label.setWordWrap(True)
        layout.addWidget(self.file_label)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            for url in urls:
                if url.toLocalFile().lower().endswith('.md'):
                    event.acceptProposedAction()
                    self.setStyleSheet("""
                        DropArea {
                            border: 2px solid #5B8DEF;
                            border-radius: 12px;
                            background-color: #E8EFFF;
                        }
                    """)
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet("""
            DropArea {
                border: 2px dashed #AABBCC;
                border-radius: 12px;
                background-color: #F8FAFC;
            }
            DropArea:hover {
                border-color: #5B8DEF;
                background-color: #EEF2FF;
            }
        """)

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        for url in urls:
            file_path = url.toLocalFile()
            if file_path.lower().endswith('.md'):
                self.set_file(file_path)
                break

        self.setStyleSheet("""
            DropArea {
                border: 2px dashed #AABBCC;
                border-radius: 12px;
                background-color: #F8FAFC;
            }
            DropArea:hover {
                border-color: #5B8DEF;
                background-color: #EEF2FF;
            }
        """)

    def set_file(self, file_path: str):
        """设置选中的文件。"""
        self.file_path = file_path
        filename = os.path.basename(file_path)
        self.file_label.setText(f"✅ {filename}")
        self.icon_label.setText("📝")
        self.label.setText("已选择文件：")
        if self.parent_window and hasattr(self.parent_window, 'on_file_selected'):
            self.parent_window.on_file_selected()


class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("MD → DOCX 转换工具")
        self.setMinimumSize(520, 480)
        self.resize(520, 480)

        # 主容器
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(24, 24, 24, 24)
        main_layout.setSpacing(16)

        # 应用全局样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #FFFFFF;
            }
            QGroupBox {
                font-family: "Microsoft YaHei";
                font-size: 12px;
                font-weight: bold;
                border: 1px solid #E0E0E0;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 16px;
                background-color: #FAFBFC;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
                color: #333;
            }
            QRadioButton {
                font-family: "Microsoft YaHei";
                font-size: 12px;
                spacing: 8px;
                padding: 6px 4px;
            }
            QRadioButton::indicator {
                width: 16px;
                height: 16px;
            }
        """)

        # ===== 标题 =====
        title_label = QLabel("MD → DOCX 转换工具")
        title_label.setFont(QFont("Microsoft YaHei", 16, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #1A1A2E; margin-bottom: 4px;")
        main_layout.addWidget(title_label)

        subtitle = QLabel("将 Markdown 文件转换为格式化的 Word 文档")
        subtitle.setFont(QFont("Microsoft YaHei", 9))
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("color: #888;")
        main_layout.addWidget(subtitle)

        # ===== 拖拽区域 =====
        self.drop_area = DropArea(self)
        main_layout.addWidget(self.drop_area)

        # ===== 选择文件按钮 =====
        select_btn = QPushButton("📂  选择 Markdown 文件")
        select_btn.setFont(QFont("Microsoft YaHei", 10))
        select_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        select_btn.setStyleSheet("""
            QPushButton {
                background-color: #F0F4FF;
                color: #5B8DEF;
                border: 1px solid #C8D8F8;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #E0EAFF;
                border-color: #5B8DEF;
            }
            QPushButton:pressed {
                background-color: #D0DFFF;
            }
        """)
        select_btn.clicked.connect(self.select_file)
        main_layout.addWidget(select_btn)

        # ===== 样式选择 =====
        style_group = QGroupBox("选择转换样式")
        style_layout = QVBoxLayout(style_group)
        style_layout.setSpacing(4)

        self.style_btn_group = QButtonGroup(self)

        self.radio_a = QRadioButton("样式 A：文档标题作为「标题」（Title）\n    # → Title,  ## → Heading 1,  ### → Heading 2")
        self.radio_b = QRadioButton("样式 B：文档标题作为「标题 1」（Heading 1）\n    # → Heading 1,  ## → Heading 2,  ### → Heading 3")
        self.radio_a.setChecked(True)

        self.style_btn_group.addButton(self.radio_a, 0)
        self.style_btn_group.addButton(self.radio_b, 1)

        style_layout.addWidget(self.radio_a)
        style_layout.addWidget(self.radio_b)
        main_layout.addWidget(style_group)

        # ===== 转换按钮 =====
        self.convert_btn = QPushButton("🚀  开始转换")
        self.convert_btn.setFont(QFont("Microsoft YaHei", 12, QFont.Weight.Bold))
        self.convert_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.convert_btn.setEnabled(False)
        self.convert_btn.setStyleSheet("""
            QPushButton {
                background-color: #5B8DEF;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px 24px;
                min-height: 20px;
            }
            QPushButton:hover {
                background-color: #4A7CE0;
            }
            QPushButton:pressed {
                background-color: #3A6CD0;
            }
            QPushButton:disabled {
                background-color: #CCC;
                color: #999;
            }
        """)
        self.convert_btn.clicked.connect(self.do_convert)
        main_layout.addWidget(self.convert_btn)

        # ===== 状态栏 =====
        self.status_label = QLabel("")
        self.status_label.setFont(QFont("Microsoft YaHei", 9))
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("color: #888;")
        main_layout.addWidget(self.status_label)

        main_layout.addStretch()

    def on_file_selected(self):
        """文件选中后启用转换按钮。"""
        self.convert_btn.setEnabled(True)
        self.status_label.setText("")

    def select_file(self):
        """弹出文件选择对话框。"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Markdown 文件",
            "",
            "Markdown 文件 (*.md);;所有文件 (*.*)"
        )
        if file_path:
            self.drop_area.set_file(file_path)

    def do_convert(self):
        """执行转换。"""
        file_path = self.drop_area.file_path
        if not file_path:
            QMessageBox.warning(self, "提示", "请先选择一个 Markdown 文件！")
            return

        if not os.path.isfile(file_path):
            QMessageBox.warning(self, "错误", f"文件不存在：\n{file_path}")
            return

        # 确定样式模式
        style_mode = 'A' if self.radio_a.isChecked() else 'B'

        # 选择保存路径
        default_name = os.path.splitext(os.path.basename(file_path))[0] + '.docx'
        default_dir = os.path.dirname(file_path)
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存 DOCX 文件",
            os.path.join(default_dir, default_name),
            "Word 文档 (*.docx)"
        )
        if not save_path:
            return

        self.status_label.setText("⏳ 正在转换...")
        self.status_label.setStyleSheet("color: #5B8DEF;")
        QApplication.processEvents()

        try:
            # 读取并解析 MD
            md_text = read_md_file(file_path)
            ast_nodes = parse_markdown(md_text)

            # 生成 DOCX
            writer = DocxWriter(style_mode=style_mode)
            writer.convert(ast_nodes, save_path)

            self.status_label.setText(f"✅ 转换成功！已保存至：{os.path.basename(save_path)}")
            self.status_label.setStyleSheet("color: #22C55E;")

            QMessageBox.information(
                self,
                "转换成功",
                f"文件已保存至：\n{save_path}"
            )

        except Exception as e:
            self.status_label.setText(f"❌ 转换失败：{str(e)}")
            self.status_label.setStyleSheet("color: #EF4444;")
            QMessageBox.critical(
                self,
                "转换失败",
                f"转换过程中出现错误：\n{str(e)}"
            )
