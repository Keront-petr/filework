from pathlib import Path
from PySide6.QtWidgets import(QApplication, QFileDialog, QMainWindow, QVBoxLayout,
                               QWidget, QTextEdit, QListWidget, QSplitter, QLabel,
                               QPushButton,QAbstractItemView,QTableWidgetItem,QTableWidget, QMessageBox, QHBoxLayout, QComboBox,QListWidgetItem)
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import Qt,QEvent, QObject
import sys
import shutil
import re
from docx import Document
import pdfplumber
import pandas

class DragnDrop(QObject):
    def __init__(self, window):
        super().__init__()
        self.window = window
        self.files_folder = Path("input_files")
        self.templates_folder = Path("templates")
        self.current_file_path = None
        self.current_template_path = None
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
        if hasattr(self.window, 'File'):
            self.window.File.triggered.connect(self.load_file_from_menu)
        if hasattr(self.window, 'Template'):
            self.window.Template.triggered.connect(self.load_template_from_menu)

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
                    dest = self.copy_file_to_folder(source, self.files_folder)
                    self.current_file_path = dest
                    self.show_content_in_edit(dest, self.window.leftContentEdit)
            self.window.statusbar.showMessage("Файлы добавлены в input_files", 2000)
        event.acceptProposedAction()

    def drop_templates(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                source = Path(url.toLocalFile())
                if source.is_file():
                    dest = self.copy_file_to_folder(source, self.templates_folder)
                    self.current_template_path = dest
                    self.show_content_in_edit(dest, self.window.rightContentEdit)
            self.window.statusbar.showMessage("Шаблоны добавлены в templates", 2000)
        event.acceptProposedAction()

    def show_content_in_edit(self, file_path: Path, text_edit):
        suffix = file_path.suffix.lower()
        content = ""
        try:
            if suffix == '.docx':
                doc = Document(file_path)
                paragraphs = [para.text for para in doc.paragraphs]
                content = "\n".join(paragraphs)
            elif suffix == '.pdf':
                with pdfplumber.open(file_path) as pdf:
                    text = ""
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text + "\n"
                    content = text
            else:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
        except UnicodeDecodeError:
            try:
                with open(file_path, 'r', encoding='cp1251') as f:
                    content = f.read()
            except Exception as e:
                content = f"Не удалось прочитать файл как текст: {e}"
        except Exception as e:
            content = f"Не удалось прочитать файл:\n{e}"
        text_edit.setPlainText(content)

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
            self.current_file_path = dest_path
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
            self.current_template_path = dest_path
            self.show_content_in_edit(dest_path, self.window.rightContentEdit)
            self.window.statusbar.showMessage(f"Шаблон скопирован в {self.templates_folder}", 2000)

class FillWindowTemplate(QMainWindow):
    def __init__(self, path, parent=None):
        super().__init__(parent)
        self.path = path
        self.markers = []
        self.variant_widgets = []
        self.generated_text = ""
        self.setWindowTitle(f"Заполнение шаблона: {path.name}")
        self.resize(1000, 700)
        self.setup_ui()
        self.load_template()

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        btn_layout = QHBoxLayout()
        self.generate_btn = QPushButton("Сгенерировать")
        self.save_btn = QPushButton("Сохранить как")
        btn_layout.addWidget(self.generate_btn)
        btn_layout.addWidget(self.save_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter)

        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        splitter.addWidget(self.text_edit)

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.addWidget(QLabel("Маркеры и варианты:"))
        self.markers_list = QListWidget()
        right_layout.addWidget(self.markers_list)
        splitter.addWidget(right_widget)

        splitter.setSizes([600, 300])

        self.generate_btn.clicked.connect(self.generate_filled_document)
        self.save_btn.clicked.connect(self.save_as)

    def load_template(self):
        suffix = self.path.suffix.lower()
        if suffix == '.docx':
            self.load_docx()
        elif suffix == '.pdf':
            self.load_pdf()
        else:
            self.load_text()
        self.display_text_with_markers()
        self.populate_variants_panel()

    def load_docx(self):
        doc = Document(self.path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        self.raw_text = "\n".join(full_text)
        self.find_markers_in_text(self.raw_text)

    def load_pdf(self):
        with pdfplumber.open(self.path) as pdf:
            full_text = []
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text.append(text)
        self.raw_text = "\n".join(full_text)
        self.find_markers_in_text(self.raw_text)

    def load_text(self):
        try:
            with open(self.path, 'r', encoding='utf-8') as f:
                self.raw_text = f.read()
        except Exception as e:
            self.raw_text = f"Ошибка чтения: {e}"
        self.find_markers_in_text(self.raw_text)

    def find_markers_in_text(self, text):
        underscore_pattern = r'_{2,}'
        index_pattern = r'<\d+>'

        markers = []
        for match in re.finditer(underscore_pattern, text):
            start, end = match.span()
            word = self.find_word(text, start)
            markers.append({
                'type': 'underscore',
                'start': start,
                'end': end,
                'text': match.group(),
                'word': word,
                'variants': self.generate_variants(word)  
            })
        for match in re.finditer(index_pattern, text):
            start, end = match.span()
            word = self.find_word(text, start)
            markers.append({
                'type': 'index',
                'start': start,
                'end': end,
                'text': match.group(),
                'word': word,
                'variants': self.generate_variants(word)
            })
        markers.sort(key=lambda x: x['start'])
        self.markers = markers

    def find_word(self, text, pos) :
        start = pos - 1
        while start >= 0 and (text[start].isalnum() or text[start] == '_'):
            start -= 1
        start += 1
        return text[start:pos]

    def generate_variants(self, word):
        if not word:
            return ["[пусто]", "[введите значение]"]
        return [word, f"новый_{word}", f"вариант_{word}", ""]

    def display_text_with_markers(self):
        self.text_edit.setPlainText(self.raw_text)

    def populate_variants_panel(self):
        self.markers_list.clear()
        self.variant_widgets.clear()
        for idx, marker in enumerate(self.markers):
            widget = QWidget()
            layout = QHBoxLayout(widget)
            label = QLabel(f"{marker['type']}: '{marker['text']}' (перед: {marker['word']})")
            combo = QComboBox()
            combo.addItems(marker['variants'])
            for i in range(combo.count()):
                if combo.itemText(i) == "":
                    combo.setItemText(i, "(пусто)")
            layout.addWidget(label)
            layout.addWidget(combo)
            item = QListWidgetItem(self.markers_list)
            item.setSizeHint(widget.sizeHint())
            self.markers_list.addItem(item)
            self.markers_list.setItemWidget(item, widget)
            self.variant_widgets.append(combo)

    def generate_filled_document(self):
        if not self.markers:
            QMessageBox.information(self, "Информация", "Маркеры не найдены")
            return
        selected_values = [cb.currentText() for cb in self.variant_widgets]
        new_text = self.raw_text
        for idx, marker in enumerate(reversed(self.markers)):
            value = selected_values[len(self.markers) - 1 - idx]
            if value == "(пусто)":
                value = ""
            new_text = new_text[:marker['start']] + value + new_text[marker['end']:]
        self.generated_text = new_text
        result_window = QTextEdit()
        result_window.setWindowTitle("Результат заполнения")
        result_window.setPlainText(new_text)
        result_window.resize(800, 600)
        result_window.show()

    def save_as(self):
        if not self.generated_text:
            QMessageBox.warning(self, "Предупреждение", "Сначала сгенерируйте документ!")
            return
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить результат как...",
            str(self.path.parent / f"{self.path.stem}_filled{self.path.suffix}"),
            "Текстовые файлы (*.txt);;Все файлы (*.*)"
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.generated_text)
                QMessageBox.information(self, "Сохранено", f"Результат сохранён в {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить: {e}")

class FillWindowFile(QMainWindow):
    def __init__(self, file_path, template_path, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.template_path = template_path
        self.setWindowTitle(f"Просмотр данных: {file_path.name}")
        self.resize(1000, 700)
        self.setup_ui()
        self.load_file()

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        btn_layout = QHBoxLayout()
        self.save_as_btn = QPushButton("Сохранить как")
        btn_layout.addWidget(self.save_as_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.stacked_widget = QWidget()
        self.stacked_layout = QVBoxLayout(self.stacked_widget)
        layout.addWidget(self.stacked_widget)

        self.save_as_btn.clicked.connect(self.save_as)

    def clear_stacked(self):
        while self.stacked_layout.count():
            child = self.stacked_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def load_file(self):
        suffix = self.file_path.suffix.lower()
        if suffix in ['.xlsx', '.xls']:
            self.load_excel()
        elif suffix == '.csv':
            self.load_csv()
        else:
            self.load_text()

    def load_excel(self):
        try:
            self.dataframe = pandas.read_excel(self.file_path,
                                               engine='openpyxl' if self.file_path.suffix == '.xlsx' else 'xlrd')
            self.display_dataframe()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить Excel:\n{e}")

    def load_csv(self):
        try:
            self.dataframe = pandas.read_csv(self.file_path)
            self.display_dataframe()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить CSV:\n{e}")

    def display_dataframe(self):
        table = QTableWidget()
        table.setRowCount(self.dataframe.shape[0])
        table.setColumnCount(self.dataframe.shape[1])
        table.setHorizontalHeaderLabels(self.dataframe.columns.astype(str))
        for i in range(self.dataframe.shape[0]):
            for j in range(self.dataframe.shape[1]):
                value = self.dataframe.iat[i, j]
                item = QTableWidgetItem(str(value) if not pandas.isna(value) else "")
                table.setItem(i, j, item)
        table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.clear_stacked()
        self.stacked_layout.addWidget(table)
        self.table_widget = table

    def load_text(self):
        try:
            with open(self.file_path, 'r', encoding='utf-8') as f:
                self.text_content = f.read()
        except UnicodeDecodeError:
            try:
                with open(self.file_path, 'r', encoding='cp1251') as f:
                    self.text_content = f.read()
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось прочитать текст: {e}")
                return
        text_edit = QTextEdit()
        text_edit.setPlainText(self.text_content)
        self.clear_stacked()
        self.stacked_layout.addWidget(text_edit)
        self.text_edit = text_edit

    def save_as(self):
        QMessageBox.information(self, "Информация", "Сохранение данных пока не реализовано")

def main():
    app = QApplication(sys.argv)
    ui_path = Path("filework.ui")
    window = QUiLoader().load(ui_path)
    window.show()

    dragdrop = DragnDrop(window)

    def on_fill():
        if dragdrop.current_template_path is None:
            QMessageBox.warning(window, "Предупреждение", "Сначала загрузите шаблон!")
            return
        if dragdrop.current_file_path is None:
            QMessageBox.warning(window, "Предупреждение", "Сначала загрузите файл данных!")
            return
        window.fill_file_window = FillWindowFile(
            dragdrop.current_file_path,
            dragdrop.current_template_path
        )
        window.fill_file_window.show()

    window.Fill.triggered.connect(on_fill)
    sys.exit(app.exec())

if __name__ == "__main__":
    main()