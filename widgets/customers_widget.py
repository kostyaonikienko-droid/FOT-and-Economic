# widgets/customers_widget.py

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QDialog, QLineEdit, QLabel
)
from database.repair_db import get_all_customers, delete_customer
from dialogs.customer_dialogs import CustomerEditDialog

class CustomersWidget(QWidget):
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
        self.btn_add = QPushButton("Добавить")
        self.btn_edit = QPushButton("Изменить")
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
        self.table.setColumnWidth(0, 30)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.hideColumn(1)
        layout.addWidget(self.table)

        self.all_customers = []

        self.btn_refresh.clicked.connect(self.load_customers)
        self.btn_add.clicked.connect(self.add_customer)
        self.btn_edit.clicked.connect(self.edit_customer)
        self.btn_delete.clicked.connect(self.delete_customer)
        self.btn_select_all.clicked.connect(self.select_all)
        self.btn_clear_all.clicked.connect(self.clear_all)

        self.load_customers()

    def load_customers(self):
        self.all_customers = get_all_customers()
        self.filter_table()

    def filter_table(self):
        text = self.search_edit.text().strip().lower()
        filtered = []
        for c in self.all_customers:
            if (text in c.name.lower() or
                text in c.phone.lower() or
                text in c.email.lower() or
                text in c.address.lower()):
                filtered.append(c)

        self.table.setRowCount(len(filtered))
        for i, c in enumerate(filtered):
            chk = QTableWidgetItem()
            chk.setCheckState(Qt.Unchecked)
            self.table.setItem(i, 0, chk)
            self.table.setItem(i, 1, QTableWidgetItem(str(c.id)))
            self.table.setItem(i, 2, QTableWidgetItem(c.name))
            self.table.setItem(i, 3, QTableWidgetItem(c.phone))
            self.table.setItem(i, 4, QTableWidgetItem(c.email))
            self.table.setItem(i, 5, QTableWidgetItem(c.address))

    def add_customer(self):
        dlg = CustomerEditDialog()
        if dlg.exec() == QDialog.DialogCode.Accepted:
            from database.repair_db import add_customer
            add_customer(dlg.get_customer())
            self.load_customers()

    def edit_customer(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Внимание", "Выберите заказчика для редактирования")
            return
        customer_id = int(self.table.item(row, 1).text())
        customers = get_all_customers()
        customer = next((c for c in customers if c.id == customer_id), None)
        if not customer:
            return
        dlg = CustomerEditDialog(customer)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            from database.repair_db import update_customer
            update_customer(dlg.get_customer())
            self.load_customers()

    def delete_customer(self):
        selected_rows = []
        for row in range(self.table.rowCount()):
            if self.table.item(row, 0).checkState() == Qt.Checked:
                selected_rows.append(row)
        if not selected_rows:
            QMessageBox.warning(self, "Внимание", "Выберите хотя бы одного заказчика для удаления")
            return
        confirm = QMessageBox.question(self, "Подтверждение", f"Удалить {len(selected_rows)} заказчик(ов)?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm != QMessageBox.StandardButton.Yes:
            return
        for row in selected_rows:
            customer_id = int(self.table.item(row, 1).text())
            delete_customer(customer_id)
        self.load_customers()

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