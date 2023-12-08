import sys
import os
import time
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QTextEdit, QFileDialog, QDesktopWidget, QTabWidget
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5 import QtGui
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QProgressBar
from PyQt5.QtWidgets import QMessageBox

import jpype
import fitz  # pip install PyMuPDF
import json



def start_jvm_if_not_started():
    if not jpype.isJVMStarted():
        current_dir = os.path.dirname(os.path.realpath(__file__))
        jre_path = os.path.join(current_dir, "jre")  # JRE在应用程序目录中的jre文件夹中
        jvm_path = os.path.join(jre_path, "bin", "server", "jvm.dll")  # Windows系统
        # jvm_path = os.path.join(jre_path, "lib", "amd64", "server", "libjvm.so")  # Linux系统
        aspose_jar_path = os.path.join(current_dir, "asposecells", "lib", "aspose-cells-23.10.jar")
        if not os.path.exists(aspose_jar_path):
            raise FileNotFoundError("Aspose jar file not found: {}".format(aspose_jar_path))
        classpath = "-Djava.class.path={}".format(aspose_jar_path)
        jpype.startJVM(jvm_path, classpath)


start_jvm_if_not_started()

# import asposecells # pip install aspose-cells
from asposecells.api import Workbook, PdfSaveOptions

def convert(file_path, output_dir):
    workbook = Workbook(file_path)
    
    saveOptions = PdfSaveOptions()
    saveOptions.setOnePagePerSheet(True)
    out_path = file_path.replace(".xlsx", "_demo.pdf")
    out_pathf = os.path.join(output_dir, os.path.basename(out_path.replace("_demo.pdf", ".pdf")))
    workbook.save(out_path, saveOptions)
    
    modify_pdf(out_path, out_pathf, "CreatedwithAspose.CellsforPython")
    return out_pathf

def modify_pdf(pdf_file, output_file, keyword):
    doc = fitz.open(pdf_file)
    for page in doc:
        info = json.loads(page.get_text('json'))
        for block in info['blocks']:
            try:
                for line in block['lines']:
                    for span in line['spans']:
                        text = span.get('text', '').replace(' ', '')
                        if  keyword in text:
                            key = span
                            page.add_redact_annot(key['bbox'])
                            page.apply_redactions()
            except KeyError:
                continue
    doc.save(output_file)
    doc.close()
 
    os.remove(pdf_file)

def main(file_paths, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for file_path in file_paths:
        if file_path.endswith(".xlsx"):
            print(f"Converting: {file_path}")
            convert(file_path, output_dir)




class ConvertThread(QThread):
    # 定义一个信号,完成时发出
    finished = pyqtSignal(str)

    def __init__(self, file_paths, output_dir):
        super().__init__()
        self.file_paths = file_paths
        self.output_dir = output_dir

    def run(self):
        total_files = len(self.file_paths)
        current_dir = os.path.dirname(sys.executable)  # 获取exe所在的运行目录
        start_time = time.time()  # 记录开始时间

        for index, file_path in enumerate(self.file_paths):
            if self.isInterruptionRequested():
                break
            self.finished.emit(f"Converting: {file_path}")
            output_file_path = convert(file_path, self.output_dir)  # 获取转换后的文件路径
            abs_output_file_path = os.path.abspath(os.path.join(current_dir, output_file_path))  # 获取绝对路径
            progress = int((index + 1) / total_files * 100)
            self.finished.emit(f"Finished converting: {abs_output_file_path}, Progress: {progress}%")  # 输出转换后的文件绝对路径

        end_time = time.time()  # 记录结束时间
        elapsed_time = end_time - start_time  # 计算耗时
        self.finished.emit(f"Conversion completed! Time elapsed: {elapsed_time:.2f} seconds")  # 发送 "Conversion completed!" 信号，同时输出耗费时间


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.stop_conversion = False

    def initUI(self):
        layout = QVBoxLayout()
        self.tab_widget = QTabWidget()

        self.page1 = QWidget()
        self.page1_layout = QVBoxLayout(self.page1)

        if getattr(sys, 'frozen', False):
            current_dir = sys._MEIPASS 
        else:
            current_dir = os.path.dirname(os.path.realpath(__file__)) 
        Qicon_path = os.path.join(current_dir, "assets", "favicon.ico")
        self.setWindowIcon(QIcon(Qicon_path))
        self.path_label = QLabel()

        self.open_file_button = QPushButton('Select or Drag File(s)')
        self.open_file_button.clicked.connect(self.open_file)

        self.open_directory_button = QPushButton('Select or Drag Directory')
        self.open_directory_button.clicked.connect(self.open_directory)

        self.output_textedit = QTextEdit()
        self.output_textedit.setReadOnly(True)

        self.convert_button = QPushButton('Convert Files')
        self.convert_button.clicked.connect(self.toggle_conversion)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)

        self.page1_layout.addWidget(self.path_label)
        self.page1_layout.addWidget(self.open_file_button)
        self.page1_layout.addWidget(self.open_directory_button)
        self.page1_layout.addWidget(self.convert_button)
        self.page1_layout.addWidget(self.output_textedit)
        self.page1_layout.addStretch()  # 添加弹性空间
        self.page1_layout.addWidget(self.progress_bar)  # 将进度条添加到布局底部

        self.page2 = QWidget()
        self.page2_layout = QVBoxLayout(self.page2)
        self.about_label = QLabel("""This software is a Graphical User Interface (GUI) application based on PyQt5. It is used to convert Excel files (.xlsx) to PDF files.
    Features:
    - Only supports Excel files in .xlsx format.
    - Supports selecting or dragging files or a single folder.
    - The output files are saved in the "pdf_output" directory.""")
        self.about_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.about_label.setWordWrap(True)
        self.about_label.setAlignment(Qt.AlignLeft)
        self.page2_layout.addWidget(self.about_label)

        self.tab_widget.addTab(self.page1, "Converter")
        self.tab_widget.addTab(self.page2, "About")

        layout.addWidget(self.tab_widget)
        self.setLayout(layout)

        self.setWindowTitle('Xlsx To Pdf')

        screen = QDesktopWidget().screenGeometry()
        self.setGeometry(0, 0, int(screen.width() / 2), int(screen.height() / 2))
        self.move(int((screen.width() - self.width()) / 2), int((screen.height() - self.height()) / 2))

        # Set Minimum and Maximum Size of the Window
        self.setMinimumSize(400, 300)

        self.setAcceptDrops(True)

    def show_message(self, message):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText(message)
        msg.setWindowTitle("Conversion Completed")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()    

    def open_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.ExistingFiles)
        dialog.setOption(QFileDialog.ReadOnly)
        dialog.setNameFilter("Excel Files (*.xlsx);;All Files (*)")  # 添加文件类型过滤器
        if dialog.exec_():
            files = dialog.selectedFiles()
            if len(files) == 1:
                file = files[0]
                if os.path.isfile(file) and file.endswith(".xlsx"):
                    self.path_label.setText(f"Selected file: {file}")
                else:
                    self.path_label.setText(f"Selected files: {', '.join(files)}")
            else:
                self.path_label.setText(f"Selected files: {', '.join(files)}")


    def open_directory(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly

        directory = QFileDialog.getExistingDirectory(self, "Select Directory", options=options)
        if directory:
            self.path_label.setText(f"Selected directory: {directory}")

    def toggle_conversion(self):
        if self.stop_conversion:
            self.stop_conversion = False
            self.convert_button.setText('Convert Files')
            self.thread.requestInterruption()
        else:
            path = self.path_label.text().replace("Selected file: ", "").replace("Selected directory: ", "").replace("Selected files: ", "")
            if not path:
                QMessageBox.warning(self, 'Warning', 'No file or directory selected.')
                return
            self.stop_conversion = True
            self.convert_button.setText('Stop')
            self.convert_files()

    def convert_files(self):
        self.output_textedit.clear()
        output_dir = 'pdf_output'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        path = self.path_label.text().replace("Selected file: ", "").replace("Selected directory: ", "").replace("Selected files: ", "")
        if os.path.isfile(path):
            file_paths = [path]
        elif os.path.isdir(path):
            file_paths = [os.path.join(path, filename) for filename in os.listdir(path) if filename.endswith(".xlsx")]
        else:
            file_paths = path.split(', ')

        self.thread = ConvertThread(file_paths, output_dir)
        self.thread.finished.connect(self.update_ui)
        self.thread.start()

    def update_ui(self, text):
        if text.startswith("Conversion completed!"):
            self.stop_conversion = False
            self.convert_button.setText('Convert Files')
            self.output_textedit.append(text)
            self.show_message(text)
        elif text.startswith("Converting: "):
            self.output_textedit.append(text)
        elif "Progress: " in text:
            progress = int(text.split(", Progress: ")[1].replace("%", ""))
            self.progress_bar.setValue(progress)
            self.output_textedit.append(text.split(", Progress: ")[0])


    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        dropped_paths = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            dropped_paths.append(path)
        if len(dropped_paths) == 1:
            path = dropped_paths[0]
            if os.path.isfile(path) and path.endswith(".xlsx"):
                self.path_label.setText(f"Selected file: {path}")
            elif os.path.isdir(path):
                self.path_label.setText(f"Selected directory: {path}")
        else:
            self.path_label.setText(f"Selected files: {', '.join(dropped_paths)}")

class StdoutRedirect:
    def __init__(self, output_textedit):
        self.output_textedit = output_textedit

    def write(self, text):
        cursor = self.output_textedit.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.output_textedit.setTextCursor(cursor)
        self.output_textedit.ensureCursorVisible()

    def flush(self):
        pass

def start_app():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.stdout = StdoutRedirect(window.output_textedit)
    sys.exit(app.exec_())

if __name__ == "__main__":
    start_app()
