# widgets/suppliers_widget.py

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QDialog, QLineEdit, QLabel
)
from database.repair_db import get_all_suppliers, delete_supplier
from dialogs.supplier_dialogs import SupplierDetailsDialog

class SuppliersWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)

        # Поиск
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Поиск:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Введите текст для поиска...")
        self.search_edit.textChanged.connect(self.filter_table)
        search_layout.addWidget(self.search_edit)
        search_layout.addStretch()
        layout.addLayout(search_layout)

        # Панель кнопок
        btn_layout = QHBoxLayout()
        self.btn_refresh = QPushButton("Обновить")
        self.btn_add = QPushButton("Добавить поставщика")
        self.btn_edit = QPushButton("Редактировать")
        self.btn_delete = QPushButton("Удалить")
        self.btn_select_all = QPushButton("Выбрать всех")
        self.btn_clear_all = QPushButton("Снять выделение")

        btn_layout.addWidget(self.btn_refresh)
        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_edit)
        btn_layout.addWidget(self.btn_delete)
        btn_layout.addWidget(self.btn_select_all)
        btn_layout.addWidget(self.btn_clear_all)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # Таблица: чекбокс, ID (скрыт), наименование, телефон, email, адрес
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["", "ID", "Наименование", "Телефон", "Email", "Адрес"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 30)  # минимальная ширина для чекбокса
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.hideColumn(1)  # скрываем ID
        layout.addWidget(self.table)

        # Храним все данные для фильтрации
        self.all_suppliers = []

        # Подключение сигналов
        self.btn_refresh.clicked.connect(self.load_suppliers)
        self.btn_add.clicked.connect(self.add_supplier)
        self.btn_edit.clicked.connect(self.edit_supplier)
        self.btn_delete.clicked.connect(self.delete_supplier)
        self.btn_select_all.clicked.connect(self.select_all)
        self.btn_clear_all.clicked.connect(self.clear_all)

        self.load_suppliers()

    def load_suppliers(self):
        self.all_suppliers = get_all_suppliers()
        self.filter_table()

    def filter_table(self):
        text = self.search_edit.text().strip().lower()
        filtered = []
        for s in self.all_suppliers:
            if (text in s.name.lower() or
                text in s.phone.lower() or
                text in s.email.lower() or
                text in s.address.lower()):
                filtered.append(s)

        self.table.setRowCount(len(filtered))
        for i, s in enumerate(filtered):
            chk = QTableWidgetItem()
            chk.setCheckState(Qt.Unchecked)
            self.table.setItem(i, 0, chk)
            self.table.setItem(i, 1, QTableWidgetItem(str(s.id)))
            self.table.setItem(i, 2, QTableWidgetItem(s.name))
            self.table.setItem(i, 3, QTableWidgetItem(s.phone))
            self.table.setItem(i, 4, QTableWidgetItem(s.email))
            self.table.setItem(i, 5, QTableWidgetItem(s.address))

    def add_supplier(self):
        dlg = SupplierDetailsDialog()
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.load_suppliers()

    def edit_supplier(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Внимание", "Выберите поставщика для редактирования")
            return
        supplier_id = int(self.table.item(row, 1).text())
        suppliers = get_all_suppliers()
        supplier = next((s for s in suppliers if s.id == supplier_id), None)
        if not supplier:
            return
        dlg = SupplierDetailsDialog(supplier, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.load_suppliers()

    def delete_supplier(self):
        selected_rows = []
        for row in range(self.table.rowCount()):
            if self.table.item(row, 0).checkState() == Qt.Checked:
                selected_rows.append(row)
        if not selected_rows:
            QMessageBox.warning(self, "Внимание", "Выберите хотя бы одного поставщика для удаления")
            return
        confirm = QMessageBox.question(self, "Подтверждение", f"Удалить {len(selected_rows)} поставщик(ов)?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm != QMessageBox.StandardButton.Yes:
            return
        for row in selected_rows:
            supplier_id = int(self.table.item(row, 1).text())
            try:
                delete_supplier(supplier_id)
            except ValueError as e:
                QMessageBox.warning(self, "Ошибка", str(e))
                return
        self.load_suppliers()


    def select_all(self):
        for row in range(self.table.rowCount()):
            self.table.item(row, 0).setCheckState(Qt.Checked)

    def clear_all(self):
        for row in range(self.table.rowCount()):
            self.table.item(row, 0).setCheckState(Qt.Unchecked)

    def populate_menu(self, menu_bar):
        pass

    def maybe_save(self):
        return True