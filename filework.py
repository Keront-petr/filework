from pathlib import *
from PySide6.QtWidgets import *
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import Qt
from PySide6.QtGui import  QDragEnterEvent, QDropEvent
import sys
import shutil
class DragnDrop:
    def __init__(self, window):
        self.window = window
        self.target_folder = Path("input_files")
        self.setup()
    def setup(self):
        self.target_folder.mkdir(exist_ok=True)
        if hasattr(self.window, 'f_frame'):
            self.window.f_frame.QdragEnterEvent = self.drag_enter
            self.window.f_frame.QdropEvent = self.drop_files
        if hasattr(self.window, 'fileListWidget'):
            self.window.fileListWidget.itemClicked.connect(self.show_content)
    def drag_enter(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()
    def drop_files(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                source = Path(url.toLocalFile())
                if source.is_file():
                    dest = self.target_folder / source.name
                    shutil.copy2(source, dest)
            self.update_file_list()
            self.window.f_frame.setStyleSheet("")
        event.accept()
    def update_file_list(self):
        if hasattr(self.window, 'fileListWidget'):
            self.window.fileListWidget.clear()
            for file_path in self.target_folder.iterdir():
                if file_path.is_file():
                    item = QListWidgetItem(file_path.name)
                    item.setData(Qt.UserRole, str(file_path))
                    self.window.fileListWidget.addItem(item)
    def show_content(self, item):
        if hasattr(self.window, 'contentTextEdit'):
            file_path = Path(item.data(Qt.UserRole))
            try:
                if file_path.suffix.lower() in ['.txt','.md','.docx']:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    self.window.contentTextEdit.setPlainText(content)
                else:
                    self.window.contentTextEdit.setPlainText(f"Файл: {file_path.name}\n")
            except Exception as e:
                self.window.contentTextEdit.setPlainText(f"Ошибка: {e}")
def main():
    app = QApplication(sys.argv)
    ui_path = Path("filework.ui")
    window = QUiLoader().load(ui_path)
    window.show()
    sys.exit(app.exec())
if __name__ == "__main__":
    main()
