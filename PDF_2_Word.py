import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QFileDialog,
                             QListWidget, QMessageBox, QProgressBar)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon, QDragEnterEvent, QDropEvent
from pdf2docx import Converter


# --- 工作线程：用于后台转换，防止界面卡死 ---
class ConversionWorker(QThread):
    # 定义信号：更新进度条、更新状态文字、完成通知
    progress_signal = pyqtSignal(int)
    status_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, file_list):
        super().__init__()
        self.file_list = file_list

    def run(self):
        total_files = len(self.file_list)
        for index, pdf_path in enumerate(self.file_list):
            try:
                # 生成 Word 文件路径 (同目录下，后缀改为 .docx)
                docx_path = os.path.splitext(pdf_path)[0] + ".docx"
                file_name = os.path.basename(pdf_path)

                self.status_signal.emit(f"正在转换: {file_name} ...")

                # 开始转换
                cv = Converter(pdf_path)
                cv.convert(docx_path, start=0, end=None)
                cv.close()

                # 更新进度
                progress = int(((index + 1) / total_files) * 100)
                self.progress_signal.emit(progress)

            except Exception as e:
                # 如果出错，打印错误但不中断整个循环
                print(f"Error converting {pdf_path}: {e}")
                self.status_signal.emit(f"转换失败: {file_name}")

        self.finished_signal.emit()


# --- 主界面 ---
class PDFConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # 基本窗口设置
        self.setWindowTitle("PDF转Word - 作者:小庄-Python办公")
        self.resize(600, 450)
        self.setAcceptDrops(True)  # 开启拖拽支持

        # 初始化UI
        self.init_ui()

        # 存储待转换文件列表
        self.pdf_files = []

    def init_ui(self):
        # 主容器
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)

        # 1. 顶部说明与操作区
        top_layout = QHBoxLayout()

        self.btn_add_file = QPushButton("选择文件")
        self.btn_add_folder = QPushButton("选择文件夹")
        self.btn_clear = QPushButton("清空列表")

        # 设置按钮样式
        self.setup_styles()

        top_layout.addWidget(self.btn_add_file)
        top_layout.addWidget(self.btn_add_folder)
        top_layout.addWidget(self.btn_clear)
        main_layout.addLayout(top_layout)

        # 2. 中间列表区 (显示待转换文件)
        self.list_widget = QListWidget()
        self.list_widget.setAlternatingRowColors(True)
        # [修正] 删除了 setPlaceholderText，改为使用 ToolTip
        self.list_widget.setToolTip("请将 PDF 文件拖拽到此处，或点击上方按钮选择...")
        main_layout.addWidget(self.list_widget)

        # 3. 底部状态与执行区
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        main_layout.addWidget(self.progress_bar)

        self.lbl_status = QLabel("准备就绪 (支持拖拽文件到窗口)")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.lbl_status)

        self.btn_start = QPushButton("开始转换")
        self.btn_start.setMinimumHeight(40)
        self.btn_start.setStyleSheet("background-color: #28a745; color: white; font-weight: bold; font-size: 14px;")
        main_layout.addWidget(self.btn_start)

        # 4. 信号绑定
        self.btn_add_file.clicked.connect(self.select_file)
        self.btn_add_folder.clicked.connect(self.select_folder)
        self.btn_clear.clicked.connect(self.clear_list)
        self.btn_start.clicked.connect(self.start_conversion)

    def setup_styles(self):
        # 简单的美化
        style = """
            QPushButton { padding: 6px; border-radius: 4px; background-color: #007bff; color: white; }
            QPushButton:hover { background-color: #0069d9; }
            QPushButton#clear_btn { background-color: #dc3545; }
            QListWidget { border: 2px dashed #aaa; border-radius: 5px; padding: 5px; font-size: 13px; }
        """
        self.btn_clear.setObjectName("clear_btn")  # 给清空按钮单独ID以应用红色样式
        self.setStyleSheet(style)

    # --- 拖拽功能实现 ---
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        files_added = False
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isfile(path) and path.lower().endswith('.pdf'):
                self.add_file_to_list(path)
                files_added = True
            elif os.path.isdir(path):
                self.scan_folder(path)
                files_added = True

        if files_added:
            self.lbl_status.setText(f"已添加文件，共 {len(self.pdf_files)} 个")

    # --- 文件操作逻辑 ---
    def select_file(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择PDF文件", "", "PDF Files (*.pdf)")
        if files:
            for f in files:
                self.add_file_to_list(f)
            self.lbl_status.setText(f"已添加文件，共 {len(self.pdf_files)} 个")

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            self.scan_folder(folder)

    def scan_folder(self, folder_path):
        count = 0
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    full_path = os.path.join(root, file)
                    self.add_file_to_list(full_path)
                    count += 1
        self.lbl_status.setText(f"文件夹扫描完成，添加了 {count} 个文件")

    def add_file_to_list(self, path):
        if path not in self.pdf_files:
            self.pdf_files.append(path)
            self.list_widget.addItem(path)

    def clear_list(self):
        self.pdf_files.clear()
        self.list_widget.clear()
        self.progress_bar.setValue(0)
        self.lbl_status.setText("列表已清空")

    # --- 转换逻辑 ---
    def start_conversion(self):
        if not self.pdf_files:
            QMessageBox.warning(self, "提示", "请先添加 PDF 文件！")
            return

        self.btn_start.setEnabled(False)
        self.btn_add_file.setEnabled(False)
        self.btn_add_folder.setEnabled(False)
        self.btn_clear.setEnabled(False)
        self.setAcceptDrops(False)  # 转换时禁止拖拽

        # 启动线程
        self.worker = ConversionWorker(self.pdf_files)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.status_signal.connect(self.update_status)
        self.worker.finished_signal.connect(self.conversion_finished)
        self.worker.start()

    def update_progress(self, val):
        self.progress_bar.setValue(val)

    def update_status(self, text):
        self.lbl_status.setText(text)

    def conversion_finished(self):
        self.lbl_status.setText("所有转换已完成！")
        QMessageBox.information(self, "成功", "所有 PDF 已成功转换为 Word！\n保存在原文件目录下。")

        # 恢复界面状态
        self.btn_start.setEnabled(True)
        self.btn_add_file.setEnabled(True)
        self.btn_add_folder.setEnabled(True)
        self.btn_clear.setEnabled(True)
        self.setAcceptDrops(True)
        self.progress_bar.setValue(100)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFConverterApp()
    window.show()
    sys.exit(app.exec_())