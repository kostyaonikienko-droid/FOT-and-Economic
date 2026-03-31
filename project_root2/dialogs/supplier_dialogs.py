# dialogs/supplier_dialogs.py

import datetime
import os
from PySide6.QtCore import Qt, QDate
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QDoubleSpinBox, QDialogButtonBox,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QPushButton, QMessageBox, QComboBox, QDateEdit,
    QTabWidget, QLabel, QTextEdit, QWidget, QFileDialog
)
from models.repair_models import Supplier, Agreement, Document
from database.repair_db import (
    get_all_suppliers, get_supplier_by_id, add_supplier, update_supplier, delete_supplier,
    get_agreements_by_supplier, get_agreement_by_id, add_agreement, update_agreement, delete_agreement,
    get_purchases_by_agreement, get_purchase_by_id, add_purchase, update_purchase, delete_purchase,
    get_documents_by_supplier, add_document, delete_document
)
from dialogs.common import select_item, QuantityPriceDialog
from .order_dialogs import PurchaseEditDialog  # будет определён позже

class SupplierDetailsDialog(QDialog):
    def __init__(self, supplier: Supplier = None, parent=None):
        super().__init__(parent)
        self.supplier = supplier if supplier else Supplier()
        self.setWindowTitle("Редактирование поставщика" if supplier else "Новый поставщик")
        self.resize(1000, 700)

        layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        # Вкладка "Основное"
        self.info_tab = QWidget()
        self.tabs.addTab(self.info_tab, "Основное")
        self.setup_info_tab()

        # Вкладка "Договоры"
        self.agreements_tab = QWidget()
        self.tabs.addTab(self.agreements_tab, "Договоры")
        self.setup_agreements_tab()

        # Вкладка "Счета"
        self.invoices_tab = QWidget()
        self.tabs.addTab(self.invoices_tab, "Счета")
        self.setup_invoices_tab()

        # Вкладка "Документы"
        self.docs_tab = QWidget()
        self.tabs.addTab(self.docs_tab, "Документы")
        self.setup_docs_tab()

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        if self.supplier.id:
            self.load_agreements()
            self.load_invoices()
            self.load_documents()

    def setup_info_tab(self):
        form = QFormLayout(self.info_tab)
        self.name_edit = QLineEdit(self.supplier.name)
        form.addRow("Наименование:", self.name_edit)
        self.phone_edit = QLineEdit(self.supplier.phone)
        form.addRow("Телефон:", self.phone_edit)
        self.email_edit = QLineEdit(self.supplier.email)
        form.addRow("Email:", self.email_edit)
        self.address_edit = QLineEdit(self.supplier.address)
        form.addRow("Адрес:", self.address_edit)
        self.notes_edit = QTextEdit(self.supplier.notes)
        self.notes_edit.setMaximumHeight(80)
        form.addRow("Примечание:", self.notes_edit)

    def setup_agreements_tab(self):
        layout = QVBoxLayout(self.agreements_tab)
        self.agreements_table = QTableWidget(0, 6)
        self.agreements_table.setHorizontalHeaderLabels(["Номер", "Дата", "Сумма", "Израсходовано", "Остаток", "Статус"])
        self.agreements_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.agreements_table)

        btn_layout = QHBoxLayout()
        self.btn_add_agreement = QPushButton("Добавить договор")
        self.btn_edit_agreement = QPushButton("Изменить")
        self.btn_delete_agreement = QPushButton("Удалить")
        btn_layout.addWidget(self.btn_add_agreement)
        btn_layout.addWidget(self.btn_edit_agreement)
        btn_layout.addWidget(self.btn_delete_agreement)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.btn_add_agreement.clicked.connect(self.add_agreement)
        self.btn_edit_agreement.clicked.connect(self.edit_agreement)
        self.btn_delete_agreement.clicked.connect(self.delete_agreement)

    def load_agreements(self):
        agreements = get_agreements_by_supplier(self.supplier.id)
        self.agreements_table.setRowCount(len(agreements))
        for i, a in enumerate(agreements):
            self.agreements_table.setItem(i, 0, QTableWidgetItem(a.agreement_number))
            self.agreements_table.setItem(i, 1, QTableWidgetItem(a.date.strftime("%d.%m.%Y")))
            self.agreements_table.setItem(i, 2, QTableWidgetItem(f"{a.total_amount:.2f}"))
            self.agreements_table.setItem(i, 3, QTableWidgetItem(f"{a.spent_amount:.2f}"))
            self.agreements_table.setItem(i, 4, QTableWidgetItem(f"{a.remaining_amount:.2f}"))
            self.agreements_table.setItem(i, 5, QTableWidgetItem(a.status))
            self.agreements_table.item(i, 0).setData(Qt.UserRole, a.id)

    def add_agreement(self):
        # Если поставщик ещё не сохранён, сначала сохраняем его
        if not self.supplier.id:
            # Сохраняем основную информацию
            self.supplier.name = self.name_edit.text()
            self.supplier.phone = self.phone_edit.text()
            self.supplier.email = self.email_edit.text()
            self.supplier.address = self.address_edit.text()
            self.supplier.notes = self.notes_edit.toPlainText()
            new_id = add_supplier(self.supplier)
            if not new_id:
                QMessageBox.warning(self, "Ошибка", "Не удалось сохранить поставщика")
                return
            self.supplier.id = new_id
        dlg = AgreementEditDialog(parent=self)
        # Теперь можно передать supplier_id
        if dlg.exec() == QDialog.DialogCode.Accepted:
            agreement = dlg.get_agreement()
            agreement.supplier_id = self.supplier.id
            add_agreement(agreement)
            self.load_agreements()
            self.load_invoices()

    def edit_agreement(self):
        row = self.agreements_table.currentRow()
        if row < 0:
            return
        agreement_id = self.agreements_table.item(row, 0).data(Qt.UserRole)
        agreement = get_agreement_by_id(agreement_id)
        if not agreement:
            return
        dlg = AgreementEditDialog(agreement, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            update_agreement(dlg.get_agreement())
            self.load_agreements()
            self.load_invoices()

    def delete_agreement(self):
        row = self.agreements_table.currentRow()
        if row < 0:
            return
        agreement_id = self.agreements_table.item(row, 0).data(Qt.UserRole)
        confirm = QMessageBox.question(self, "Подтверждение", "Удалить договор?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.Yes:
            delete_agreement(agreement_id)
            self.load_agreements()
            self.load_invoices()
            self.load_documents()

    def setup_invoices_tab(self):
        layout = QVBoxLayout(self.invoices_tab)
        self.invoices_table = QTableWidget(0, 5)
        self.invoices_table.setHorizontalHeaderLabels(["Номер счёта", "Дата", "Сумма", "Договор", "Примечание"])
        self.invoices_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.invoices_table)

        btn_layout = QHBoxLayout()
        self.btn_add_invoice = QPushButton("Добавить счёт")
        self.btn_edit_invoice = QPushButton("Изменить")
        self.btn_delete_invoice = QPushButton("Удалить")
        btn_layout.addWidget(self.btn_add_invoice)
        btn_layout.addWidget(self.btn_edit_invoice)
        btn_layout.addWidget(self.btn_delete_invoice)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.btn_add_invoice.clicked.connect(self.add_invoice)
        self.btn_edit_invoice.clicked.connect(self.edit_invoice)
        self.btn_delete_invoice.clicked.connect(self.delete_invoice)

    def load_invoices(self):
        agreements = get_agreements_by_supplier(self.supplier.id)
        all_purchases = []
        for a in agreements:
            purchases = get_purchases_by_agreement(a.id)
            for p in purchases:
                p.agreement_number = a.agreement_number
                all_purchases.append(p)
        all_purchases.sort(key=lambda x: x.date, reverse=True)
        self.invoices_table.setRowCount(len(all_purchases))
        for i, p in enumerate(all_purchases):
            total = sum(item.quantity * item.purchase_price for item in p.items)
            self.invoices_table.setItem(i, 0, QTableWidgetItem(p.invoice_number))
            self.invoices_table.setItem(i, 1, QTableWidgetItem(p.date.strftime("%d.%m.%Y")))
            self.invoices_table.setItem(i, 2, QTableWidgetItem(f"{total:.2f}"))
            self.invoices_table.setItem(i, 3, QTableWidgetItem(p.agreement_number))
            self.invoices_table.setItem(i, 4, QTableWidgetItem(p.notes))
            self.invoices_table.item(i, 0).setData(Qt.UserRole, p.id)

    def add_invoice(self):
        agreements = get_agreements_by_supplier(self.supplier.id)
        if not agreements:
            QMessageBox.warning(self, "Внимание", "Сначала добавьте договор для этого поставщика")
            return
        selected = select_item(
            self,
            "Выберите договор",
            agreements,
            ["Номер", "Дата", "Сумма", "Остаток"],
            lambda a: [a.agreement_number, a.date.strftime("%d.%m.%Y"),
                       f"{a.total_amount:.2f}", f"{a.remaining_amount:.2f}"]
        )
        if selected == "NEW":
            # Здесь можно было бы создать новый договор, но пока пропустим
            QMessageBox.information(self, "Информация",
                                    "Функция создания нового договора в этом окне не предусмотрена. Используйте вкладку 'Договоры'.")
            return
        if selected:
            dlg = PurchaseEditDialog(agreement_id=selected.id, parent=self)


            if dlg.exec() == QDialog.DialogCode.Accepted:
                try:
                    add_purchase(dlg.purchase)
                    self.load_invoices()
                    self.load_agreements()
                except ValueError as e:
                    QMessageBox.warning(self, "Ошибка", str(e))

    def edit_invoice(self):
        row = self.invoices_table.currentRow()
        if row < 0:
            return
        purchase_id = self.invoices_table.item(row, 0).data(Qt.UserRole)
        purchase = get_purchase_by_id(purchase_id)  # эту функцию нужно добавить в database
        if not purchase:
            return
        dlg = PurchaseEditDialog(purchase=purchase, parent=self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            update_purchase(dlg.purchase)
            self.load_invoices()
            self.load_agreements()

    def delete_invoice(self):
        row = self.invoices_table.currentRow()
        if row < 0:
            return
        purchase_id = self.invoices_table.item(row, 0).data(Qt.UserRole)
        confirm = QMessageBox.question(self, "Подтверждение", "Удалить счёт?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.Yes:
            delete_purchase(purchase_id)
            self.load_invoices()
            self.load_agreements()

    def setup_docs_tab(self):
        layout = QVBoxLayout(self.docs_tab)
        self.docs_table = QTableWidget(0, 5)
        self.docs_table.setHorizontalHeaderLabels(["Тип", "Название файла", "Дата загрузки", "Описание", "ID"])
        self.docs_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.docs_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.docs_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.docs_table.hideColumn(4)
        layout.addWidget(self.docs_table)

        btn_layout = QHBoxLayout()
        self.btn_add_doc = QPushButton("Добавить документ")
        self.btn_open_doc = QPushButton("Открыть")
        self.btn_delete_doc = QPushButton("Удалить")
        btn_layout.addWidget(self.btn_add_doc)
        btn_layout.addWidget(self.btn_open_doc)
        btn_layout.addWidget(self.btn_delete_doc)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.btn_add_doc.clicked.connect(self.add_document)
        self.btn_open_doc.clicked.connect(self.open_document)
        self.btn_delete_doc.clicked.connect(self.delete_document)

    def load_documents(self):
        docs = get_documents_by_supplier(self.supplier.id)
        self.docs_table.setRowCount(len(docs))
        for i, d in enumerate(docs):
            self.docs_table.setItem(i, 0, QTableWidgetItem(d.document_type))
            self.docs_table.setItem(i, 1, QTableWidgetItem(d.file_name))
            self.docs_table.setItem(i, 2, QTableWidgetItem(d.uploaded_date.strftime("%d.%m.%Y") if d.uploaded_date else ""))
            self.docs_table.setItem(i, 3, QTableWidgetItem(d.description))
            self.docs_table.setItem(i, 4, QTableWidgetItem(str(d.id)))

    def add_document(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите файл", "", "Все файлы (*.*)")
        if not file_path:
            return
        file_name = os.path.basename(file_path)
        # Диалог для ввода описания и типа
        dlg = QDialog(self)
        dlg.setWindowTitle("Информация о документе")
        layout = QVBoxLayout(dlg)
        form = QFormLayout()
        type_combo = QComboBox()
        type_combo.addItems(["договор", "счет", "спецификация", "акт", "иное"])
        form.addRow("Тип документа:", type_combo)
        desc_edit = QLineEdit()
        form.addRow("Описание:", desc_edit)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        doc = Document(
            supplier_id=self.supplier.id,
            document_type=type_combo.currentText(),
            file_path=file_path,
            file_name=file_name,
            description=desc_edit.text()
        )
        add_document(doc)
        self.load_documents()

    def open_document(self):
        row = self.docs_table.currentRow()
        if row < 0:
            return
        doc_id = int(self.docs_table.item(row, 4).text())
        docs = get_documents_by_supplier(self.supplier.id)
        doc = next((d for d in docs if d.id == doc_id), None)
        if doc and os.path.exists(doc.file_path):
            os.startfile(doc.file_path)
        else:
            QMessageBox.warning(self, "Ошибка", "Файл не найден")

    def delete_document(self):
        row = self.docs_table.currentRow()
        if row < 0:
            return
        doc_id = int(self.docs_table.item(row, 4).text())
        confirm = QMessageBox.question(self, "Подтверждение", "Удалить документ?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.Yes:
            delete_document(doc_id)
            self.load_documents()

    def accept(self):
        self.supplier.name = self.name_edit.text()
        self.supplier.phone = self.phone_edit.text()
        self.supplier.email = self.email_edit.text()
        self.supplier.address = self.address_edit.text()
        self.supplier.notes = self.notes_edit.toPlainText()
        if not self.supplier.id:
            new_id = add_supplier(self.supplier)
            self.supplier.id = new_id
        else:
            update_supplier(self.supplier)
        super().accept()

class AgreementEditDialog(QDialog):
    def __init__(self, agreement: Agreement = None, parent=None):
        super().__init__(parent)
        self.agreement = agreement if agreement else Agreement()
        self.setWindowTitle("Редактирование договора" if agreement else "Новый договор")
        self.resize(500, 450)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        if self.agreement.supplier_id:
            supplier = get_supplier_by_id(self.agreement.supplier_id)
            supplier_text = supplier.name if supplier else ""
            self.supplier_label = QLabel(supplier_text)
            self.supplier_id = self.agreement.supplier_id
            form.addRow("Поставщик:", self.supplier_label)
        else:
            self.supplier_display = QLineEdit()
            self.supplier_display.setReadOnly(True)
            self.supplier_display.setPlaceholderText("Нажмите кнопку для выбора")
            self.btn_select_supplier = QPushButton("...")
            self.btn_select_supplier.setMaximumWidth(30)
            supplier_layout = QHBoxLayout()
            supplier_layout.addWidget(self.supplier_display)
            supplier_layout.addWidget(self.btn_select_supplier)
            form.addRow("Поставщик:", supplier_layout)
            self.supplier_id = 0
            self.btn_select_supplier.clicked.connect(self.select_supplier)

        self.number_edit = QLineEdit(self.agreement.agreement_number)
        form.addRow("Номер договора:", self.number_edit)

        self.date_edit = QDateEdit()
        if self.agreement.date:
            self.date_edit.setDate(QDate(self.agreement.date.year, self.agreement.date.month, self.agreement.date.day))
        else:
            self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        form.addRow("Дата:", self.date_edit)

        self.amount_spin = QDoubleSpinBox()
        self.amount_spin.setRange(0, 1e9)
        self.amount_spin.setValue(self.agreement.total_amount)
        form.addRow("Общая сумма:", self.amount_spin)

        self.status_combo = QComboBox()
        self.status_combo.addItems(["активен", "завершён", "расторгнут"])
        self.status_combo.setCurrentText(self.agreement.status or "активен")
        form.addRow("Статус:", self.status_combo)

        self.notes_edit = QTextEdit(self.agreement.notes)
        self.notes_edit.setMaximumHeight(80)
        form.addRow("Примечание:", self.notes_edit)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def select_supplier(self):
        suppliers = get_all_suppliers()
        if not suppliers:
            QMessageBox.information(self, "Информация", "Нет поставщиков в справочнике")
            return
        selected = select_item(
            self,
            "Выберите поставщика",
            suppliers,
            ["ID", "Наименование", "Телефон", "Email"],
            lambda s: [s.id, s.name, s.phone, s.email]
        )
        if selected:
            self.supplier_display.setText(selected.name)
            self.supplier_id = selected.id

    def get_agreement(self) -> Agreement:
        self.agreement.supplier_id = self.supplier_id
        self.agreement.agreement_number = self.number_edit.text()
        self.agreement.date = self.date_edit.date().toPython()
        self.agreement.total_amount = self.amount_spin.value()
        self.agreement.status = self.status_combo.currentText()
        self.agreement.notes = self.notes_edit.toPlainText()
        return self.agreement