import os
import shutil
from pathlib import Path
from docxtpl import DocxTemplate
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton,
                             QLabel, QListWidget, QLineEdit, QFormLayout, QFileDialog,
                             QMessageBox, QHBoxLayout)
from PyQt5.QtCore import Qt


class TemplateGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Генератор документов из шаблонов")
        self.setGeometry(100, 100, 800, 600)

        # Папка для хранения шаблонов
        self.templates_dir = os.path.join(os.path.expanduser("~"), "DocumentTemplates")
        if not os.path.exists(self.templates_dir):
            os.makedirs(self.templates_dir)

        self.init_ui()
        self.current_template = None
        self.template_fields = {}

    def init_ui(self):
        main_widget = QWidget()
        main_layout = QVBoxLayout()

        # Верхняя панель: загрузка шаблонов
        upload_layout = QHBoxLayout()
        self.upload_btn = QPushButton("Загрузить шаблоны")
        self.upload_btn.clicked.connect(self.upload_templates)
        upload_layout.addWidget(self.upload_btn)

        #Удаление шаблонов
        self.delete_btn = QPushButton("Удалить шаблон")
        self.delete_btn.clicked.connect(self.delete_template)
        upload_layout.addWidget(self.delete_btn)

        self.status_label = QLabel("Готов к работе")
        upload_layout.addWidget(self.status_label)

        main_layout.addLayout(upload_layout)

        # Список доступных шаблонов
        self.templates_list = QListWidget()
        self.templates_list.itemClicked.connect(self.load_template)
        self.refresh_templates_list()
        main_layout.addWidget(QLabel("Доступные шаблоны:"))
        main_layout.addWidget(self.templates_list)

        # Область для заполнения полей
        self.form_layout = QFormLayout()
        self.fields_widget = QWidget()
        self.fields_widget.setLayout(self.form_layout)
        main_layout.addWidget(QLabel("Заполните поля:"))
        main_layout.addWidget(self.fields_widget)

        # Кнопка генерации документа
        self.generate_btn = QPushButton("Сгенерировать документ")
        self.generate_btn.clicked.connect(self.generate_document)
        self.generate_btn.setEnabled(False)
        main_layout.addWidget(self.generate_btn)

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

    def upload_templates(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Выберите шаблоны документов", "", "Word Documents (*.docx)"
        )

        if files:
            for file_path in files:
                file_name = os.path.basename(file_path)
                dest_path = os.path.join(self.templates_dir, file_name)
                shutil.copy2(file_path, dest_path)

            self.refresh_templates_list()
            self.status_label.setText(f"Загружено {len(files)} шаблонов")

    def delete_template(self):
        selected_item = self.templates_list.currentItem()  # Получаем выбранный шаблон
        if selected_item:
            template_name = selected_item.text()
            template_path = os.path.join(self.templates_dir, template_name)
            os.remove(template_path)  # Удаляем файл
            self.refresh_templates_list()  # Обновляем список

    def refresh_templates_list(self):
        self.templates_list.clear()
        for file_name in os.listdir(self.templates_dir):
            if file_name.endswith(".docx"):
                self.templates_list.addItem(file_name)

    def load_template(self, item):
        template_name = item.text()
        template_path = os.path.join(self.templates_dir, template_name)

        try:
            doc = DocxTemplate(template_path)
            context = doc.get_undeclared_template_variables()

            # Очищаем предыдущие поля
            self.clear_fields()

            # Создаем поля для заполнения
            self.template_fields = {}
            for field in context:
                label = QLabel(field.replace("_", " ").capitalize() + ":")
                line_edit = QLineEdit()
                self.form_layout.addRow(label, line_edit)
                self.template_fields[field] = line_edit

            self.current_template = template_path
            self.generate_btn.setEnabled(True)
            self.status_label.setText(f"Загружен шаблон: {template_name}")

        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить шаблон: {str(e)}")

    def clear_fields(self):
        # Удаляем все поля формы
        while self.form_layout.rowCount() > 0:
            self.form_layout.removeRow(0)

        self.template_fields = {}

    def generate_document(self):
        if not self.current_template:
            return

        # Собираем данные из полей
        context = {}
        for field_name, field_widget in self.template_fields.items():
            context[field_name] = field_widget.text()

        # Проверяем, что все поля заполнены
        empty_fields = [field for field, value in context.items() if not value.strip()]
        if empty_fields:
            QMessageBox.warning(
                self,
                "Пустые поля",
                f"Пожалуйста, заполните все поля. Пустые поля: {', '.join(empty_fields)}"
            )
            return

        # Выбираем место для сохранения
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить документ",
            "",
            "Word Documents (*.docx)"
        )

        if save_path:
            try:
                doc = DocxTemplate(self.current_template)
                doc.render(context)
                doc.save(save_path)
                self.status_label.setText(f"Документ сохранен: {os.path.basename(save_path)}")
                QMessageBox.information(self, "Успех", "Документ успешно сгенерирован и сохранен!")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить документ: {str(e)}")


if __name__ == "__main__":
    app = QApplication([])
    window = TemplateGeneratorApp()
    window.show()
    app.exec_()