from pathlib import Path
from PySide6.QtWidgets import QApplication, QFileDialog
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import Qt, QEvent, QObject
from PySide6.QtGui import QDragEnterEvent, QDropEvent
import sys
import shutil

class DragnDrop(QObject):
    def __init__(self, window):
        super().__init__()
        self.window = window
        self.files_folder = Path("input_files")
        self.templates_folder = Path("templates")
        self.setup()

    def setup(self):
        self.files_folder.mkdir(exist_ok=True)
        self.templates_folder.mkdir(exist_ok=True)
        if hasattr(self.window, 'leftDropZone'):
            self.window.leftDropZone.setAcceptDrops(True)
            self.window.leftDropZone.installEventFilter(self)
        if hasattr(self.window, 'rightDropZone'):
            self.window.rightDropZone.setAcceptDrops(True)
            self.window.rightDropZone.installEventFilter(self)

        if hasattr(self.window, 'Files'):
            self.window.Files.triggered.connect(self.load_file_from_menu)
        if hasattr(self.window, 'Templates'):
            self.window.Templates.triggered.connect(self.load_template_from_menu)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.DragEnter:
            if obj == self.window.leftDropZone:
                self.drag_enter_files(event)
            elif obj == self.window.rightDropZone:
                self.drag_enter_templates(event)
            return True
        elif event.type() == QEvent.Drop:
            if obj == self.window.leftDropZone:
                self.drop_files(event)
            elif obj == self.window.rightDropZone:
                self.drop_templates(event)
            return True
        return super().eventFilter(obj, event)

    def drag_enter_files(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def drop_files(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                source = Path(url.toLocalFile())
                if source.is_file():
                    dest = self.files_folder / source.name
                    shutil.copy2(source, dest)
                    self.show_content_in_edit(dest, self.window.leftContentEdit)
            self.window.statusbar.showMessage("Файлы добавлены в input_files", 2000)
        event.acceptProposedAction()

    def drag_enter_templates(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def drop_templates(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                source = Path(url.toLocalFile())
                if source.is_file():
                    dest = self.templates_folder / source.name
                    shutil.copy2(source, dest)
                    self.show_content_in_edit(dest, self.window.rightContentEdit)
            self.window.statusbar.showMessage("Шаблоны добавлены в templates", 2000)
        event.acceptProposedAction()

    def show_content_in_edit(self, file_path: Path, text_edit):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            text_edit.setPlainText(content)
        except Exception as e:
            text_edit.setPlainText(f"Не удалось прочитать файл:\n{e}")

    def copy_file_to_folder(self, source_path: Path, target_folder: Path) -> Path:
        source_path = source_path.resolve()
        target_folder = target_folder.resolve()
        if source_path.parent == target_folder:
            return source_path
        dest_path = target_folder / source_path.name
        if dest_path.exists():
            stem = source_path.stem
            suffix = source_path.suffix
            counter = 1
            while dest_path.exists():
                dest_path = target_folder / f"{stem}_{counter}{suffix}"
                counter += 1
        shutil.copy2(source_path, dest_path)
        return dest_path

    def load_file_from_menu(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self.window,
            "Выберите файл",
            str(self.files_folder),
            "Все файлы (*.*)"
        )
        if file_path:
            dest_path = self.copy_file_to_folder(Path(file_path), self.files_folder)
            self.show_content_in_edit(dest_path, self.window.leftContentEdit)
            self.window.statusbar.showMessage(f"Файл скопирован в {self.files_folder}", 2000)

    def load_template_from_menu(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self.window,
            "Выберите шаблон",
            str(self.templates_folder),
            "Все файлы (*.*)"
        )
        if file_path:
            dest_path = self.copy_file_to_folder(Path(file_path), self.templates_folder)
            self.show_content_in_edit(dest_path, self.window.rightContentEdit)
            self.window.statusbar.showMessage(f"Шаблон скопирован в {self.templates_folder}", 2000)

def main():
    app = QApplication(sys.argv)
    ui_path = Path("filework.ui")
    window = QUiLoader().load(ui_path)
    window.show()
    dragdrop = DragnDrop(window)
    sys.exit(app.exec())

if __name__ == "__main__":
    main()