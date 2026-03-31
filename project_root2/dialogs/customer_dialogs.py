# dialogs/customer_dialogs.py

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout,
    QLineEdit, QTextEdit, QDialogButtonBox,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QPushButton, QMessageBox
)
from models.repair_models import Customer
from database.repair_db import (
    get_all_customers, get_customer_by_id, add_customer, update_customer, delete_customer
)

class CustomerEditDialog(QDialog):
    def __init__(self, customer: Customer = None, parent=None):
        super().__init__(parent)
        self.customer = customer if customer else Customer()
        self.setWindowTitle("Редактирование заказчика" if customer else "Новый заказчик")
        self.resize(400, 300)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.name_edit = QLineEdit(self.customer.name)
        form.addRow("Наименование:", self.name_edit)

        self.phone_edit = QLineEdit(self.customer.phone)
        form.addRow("Телефон:", self.phone_edit)

        self.email_edit = QLineEdit(self.customer.email)
        form.addRow("Email:", self.email_edit)

        self.address_edit = QLineEdit(self.customer.address)
        form.addRow("Адрес:", self.address_edit)

        self.notes_edit = QTextEdit(self.customer.notes)
        self.notes_edit.setMaximumHeight(80)
        form.addRow("Примечание:", self.notes_edit)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_customer(self) -> Customer:
        self.customer.name = self.name_edit.text()
        self.customer.phone = self.phone_edit.text()
        self.customer.email = self.email_edit.text()
        self.customer.address = self.address_edit.text()
        self.customer.notes = self.notes_edit.toPlainText()
        return self.customer

class CustomersDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Справочник заказчиков")
        self.resize(700, 400)

        layout = QVBoxLayout(self)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Наименование", "Телефон", "Email", "Адрес"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(self.table)

        btn_layout = QHBoxLayout()
        self.btn_add = QPushButton("Добавить")
        self.btn_edit = QPushButton("Изменить")
        self.btn_delete = QPushButton("Удалить")
        self.btn_close = QPushButton("Закрыть")
        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_edit)
        btn_layout.addWidget(self.btn_delete)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_close)
        layout.addLayout(btn_layout)

        self.btn_add.clicked.connect(self.add_customer)
        self.btn_edit.clicked.connect(self.edit_customer)
        self.btn_delete.clicked.connect(self.delete_customer)
        self.btn_close.clicked.connect(self.accept)

        self.load_customers()

    def load_customers(self):
        customers = get_all_customers()
        self.table.setRowCount(len(customers))
        for i, c in enumerate(customers):
            self.table.setItem(i, 0, QTableWidgetItem(c.name))
            self.table.setItem(i, 1, QTableWidgetItem(c.phone))
            self.table.setItem(i, 2, QTableWidgetItem(c.email))
            self.table.setItem(i, 3, QTableWidgetItem(c.address))

    def add_customer(self):
        dlg = CustomerEditDialog()
        if dlg.exec() == QDialog.DialogCode.Accepted:
            add_customer(dlg.get_customer())
            self.load_customers()

    def edit_customer(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Внимание", "Выберите заказчика для редактирования")
            return
        name = self.table.item(row, 0).text()
        customers = get_all_customers()
        customer = next((c for c in customers if c.name == name), None)
        if not customer:
            return
        dlg = CustomerEditDialog(customer)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            update_customer(dlg.get_customer())
            self.load_customers()

    def delete_customer(self):
        row = self.table.currentRow()
        if row < 0:
            return
        confirm = QMessageBox.question(self, "Подтверждение", "Удалить выбранного заказчика?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.Yes:
            name = self.table.item(row, 0).text()
            customers = get_all_customers()
            customer = next((c for c in customers if c.name == name), None)
            if customer:
                delete_customer(customer.id)
                self.load_customers()