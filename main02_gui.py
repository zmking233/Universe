import sys
import os
import subprocess
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QListWidget, QLabel, QFileDialog, QPushButton, QSpacerItem, QSizePolicy
)
from PyQt5.QtCore import Qt
from main02 import clean_text, parse_fields, extract_text_from_pdf, extract_text_from_docx, generate_wechat_html

class DropWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("自动排版工具")
        self.setAcceptDrops(True)
        self.resize(600, 400)
        layout = QVBoxLayout(self)
        self.label = QLabel("拖入案例库PDF文件（可多选）", self)
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)
        self.listWidget = QListWidget(self)
        layout.addWidget(self.listWidget)
        self.btn_select = QPushButton("手动选择文件", self)
        layout.addWidget(self.btn_select)
        self.btn_select.clicked.connect(self.open_file_dialog)
        self.output_dir = os.path.join(os.getcwd(), "output")
        os.makedirs(self.output_dir, exist_ok=True)
        self.listWidget.itemDoubleClicked.connect(self.open_file)

        # 添加弹性空间和右下角标签
        layout.addSpacerItem(QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding))
        self.author_label = QLabel("By LeClaire", self)
        self.author_label.setAlignment(Qt.AlignRight | Qt.AlignBottom)
        layout.addWidget(self.author_label)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        self.process_files(files)

    def open_file_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择PDF文件", "", "PDF(*.pdf)")
        if files:
            self.process_files(files)

    def process_files(self, files):
        for path in files:
            base_name = os.path.splitext(os.path.basename(path))[0]
            if path.endswith(".pdf"):
                raw_text = extract_text_from_pdf(path)
            else:
                continue
            cleaned = clean_text(raw_text)
            parsed = parse_fields(cleaned)
            html_output = generate_wechat_html(parsed)
            html_file = os.path.join(self.output_dir, f"{base_name}-公众号格式.html")
            with open(html_file, "w", encoding="utf-8") as f:
                f.write(html_output)
            self.listWidget.addItem(html_file)

    def open_file(self, item):
        path = item.text()
        if sys.platform.startswith('win'):
            os.startfile(path)
        elif sys.platform.startswith('darwin'):
            subprocess.call(['open', path])
        else:
            subprocess.call(['xdg-open', path])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = DropWidget()
    w.show()
    sys.exit(app.exec_())