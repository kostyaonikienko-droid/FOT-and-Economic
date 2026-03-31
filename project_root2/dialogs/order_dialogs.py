# dialogs/order_dialogs.py

import datetime
from PySide6.QtCore import Qt, QDate
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QDoubleSpinBox, QDialogButtonBox,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QPushButton, QMessageBox, QComboBox, QDateEdit,
    QLabel, QTextEdit, QGroupBox, QInputDialog
)
from models.repair_models import (
    Purchase, PurchaseItem, Material, MaterialBatch,
    Order, WorkItem, OrderMaterial
)
from database.repair_db import (
    get_all_materials, get_material_by_id,
    get_agreement_by_id, get_available_batches,
    add_purchase, update_purchase, delete_purchase,
    get_all_customers, get_customer_by_id,
    add_material   # добавлено
)
from dialogs.common import select_item, QuantityPriceDialog
from dialogs.material_dialogs import MaterialEditDialog  # добавлено

# ----------------------------------------------------------------------
# Диалог редактирования счёта (входящего)
# ----------------------------------------------------------------------
class PurchaseEditDialog(QDialog):
    def __init__(self, agreement_id: int = None, purchase: Purchase = None, parent=None):
        super().__init__(parent)

        self.purchase = purchase if purchase else Purchase()
        self.agreement_id = agreement_id or (purchase.agreement_id if purchase else None)
        self.setWindowTitle("Редактирование счёта" if purchase else "Новый счёт")
        self.resize(700, 500)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        # Договор (только для чтения)
        if self.agreement_id:
            agreement = get_agreement_by_id(self.agreement_id)
            self.agreement_label = QLabel(f"{agreement.agreement_number} ({agreement.supplier_name})")
            form.addRow("Договор:", self.agreement_label)

        self.date_edit = QDateEdit()
        if self.purchase.date:
            self.date_edit.setDate(QDate(self.purchase.date.year, self.purchase.date.month, self.purchase.date.day))
        else:
            self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        form.addRow("Дата:", self.date_edit)

        self.invoice_edit = QLineEdit(self.purchase.invoice_number)
        form.addRow("Номер счёта:", self.invoice_edit)

        self.notes_edit = QTextEdit(self.purchase.notes)
        self.notes_edit.setMaximumHeight(60)
        form.addRow("Примечание:", self.notes_edit)

        layout.addLayout(form)

        # Таблица позиций
        self.items_table = QTableWidget(0, 4)
        self.items_table.setHorizontalHeaderLabels(["Материал", "Количество", "Цена", "Сумма"])
        self.items_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.items_table)

        # Кнопки управления позициями
        btn_layout = QHBoxLayout()
        self.btn_add_item = QPushButton("Добавить материал")
        self.btn_edit_item = QPushButton("Изменить")
        self.btn_delete_item = QPushButton("Удалить")
        btn_layout.addWidget(self.btn_add_item)
        btn_layout.addWidget(self.btn_edit_item)
        btn_layout.addWidget(self.btn_delete_item)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.btn_add_item.clicked.connect(self.add_item)
        self.btn_edit_item.clicked.connect(self.edit_item)
        self.btn_delete_item.clicked.connect(self.delete_item)

        # Итоговая сумма
        self.total_label = QLabel("Сумма счёта: 0.00")
        layout.addWidget(self.total_label)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.load_items()

    def load_items(self):
        self.items_table.setRowCount(len(self.purchase.items))
        total = 0.0
        materials = get_all_materials()
        mat_dict = {m.id: m for m in materials}
        for i, item in enumerate(self.purchase.items):
            mat = mat_dict.get(item.material_id)
            name = mat.name if mat else f"ID {item.material_id}"
            self.items_table.setItem(i, 0, QTableWidgetItem(name))
            self.items_table.setItem(i, 1, QTableWidgetItem(f"{item.quantity:.3f}"))
            self.items_table.setItem(i, 2, QTableWidgetItem(f"{item.purchase_price:.2f}"))
            sum_item = item.quantity * item.purchase_price
            self.items_table.setItem(i, 3, QTableWidgetItem(f"{sum_item:.2f}"))
            total += sum_item
        self.total_label.setText(f"Сумма счёта: {total:.2f}")

    def add_item(self):
        materials = get_all_materials()
        if not materials:
            reply = QMessageBox.question(self, "Информация", "Нет материалов в справочнике. Создать новый?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                dlg_mat = MaterialEditDialog()
                if dlg_mat.exec() == QDialog.DialogCode.Accepted:
                    new_mat = dlg_mat.get_material()
                    add_material(new_mat)
                    self.add_item()
                return
            else:
                return

        selected = select_item(
            self,
            "Выберите материал",
            materials,
            ["ID", "Наименование", "Ед.", "Закуп", "Продажа", "Остаток"],
            lambda m: [m.id, m.name, m.unit, m.purchase_price, m.sale_price, m.stock],
            enable_new=True
        )
        if selected == "NEW":
            dlg_mat = MaterialEditDialog()
            if dlg_mat.exec() == QDialog.DialogCode.Accepted:
                new_mat = dlg_mat.get_material()
                add_material(new_mat)
                self.add_item()
            return
        if selected is None:
            return

        dlg_qty = QuantityPriceDialog("Параметры позиции",
                                      label_qty="Количество:", label_price="Цена закупки:",
                                      default_qty=1.0, default_price=selected.purchase_price,
                                      qty_max=1e9, price_max=1e9)
        if dlg_qty.exec() == QDialog.DialogCode.Accepted:
            qty, price = dlg_qty.get_values()
            self.purchase.items.append(PurchaseItem(
                material_id=selected.id, quantity=qty, purchase_price=price
            ))
            self.load_items()


    def edit_item(self):
        row = self.items_table.currentRow()
        if row < 0:
            return
        item = self.purchase.items[row]
        dlg = QuantityPriceDialog("Редактирование позиции",
                                  label_qty="Количество:", label_price="Цена закупки:",
                                  default_qty=item.quantity, default_price=item.purchase_price,
                                  qty_max=1e9, price_max=1e9)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            qty, price = dlg.get_values()
            item.quantity = qty
            item.purchase_price = price
            self.load_items()

    def delete_item(self):
        row = self.items_table.currentRow()
        if row >= 0:
            del self.purchase.items[row]
            self.load_items()

    def accept(self):
        agreement_id = self.agreement_id or self.purchase.agreement_id
        if not self.agreement_id:
            QMessageBox.warning(self, "Ошибка", "Не выбран договор")
            return
        if not self.purchase.items:
            QMessageBox.warning(self, "Ошибка", "Добавьте хотя бы одну позицию")
            return
        self.purchase.agreement_id = agreement_id
        self.purchase.date = self.date_edit.date().toPython()
        self.purchase.invoice_number = self.invoice_edit.text()
        self.purchase.notes = self.notes_edit.toPlainText()
        super().accept()

# ----------------------------------------------------------------------
# Диалог выбора партий
# ----------------------------------------------------------------------
class BatchSelectionDialog(QDialog):
    def __init__(self, material: Material, available_batches: list, parent=None):
        super().__init__(parent)
        self.material = material
        self.batches = available_batches
        self.selected = []  # список кортежей (batch_id, quantity)
        self.setWindowTitle(f"Выбор партий материала: {material.name}")
        self.resize(800, 400)

        layout = QVBoxLayout(self)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Поиск по номеру партии, поставщику...")
        self.search_edit.textChanged.connect(self.filter_table)
        layout.addWidget(self.search_edit)

        self.table = QTableWidget(len(self.batches), 5)
        self.table.setHorizontalHeaderLabels(["Выбрать", "Партия", "Дата", "Остаток", "Цена закупки"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(self.table)

        for i, batch in enumerate(self.batches):
            chk = QTableWidgetItem()
            chk.setCheckState(Qt.Unchecked)
            self.table.setItem(i, 0, chk)
            self.table.setItem(i, 1, QTableWidgetItem(batch.batch_number or str(batch.id)))
            self.table.setItem(i, 2, QTableWidgetItem(batch.date.strftime("%d.%m.%Y")))
            self.table.setItem(i, 3, QTableWidgetItem(f"{batch.quantity:.3f}"))
            self.table.setItem(i, 4, QTableWidgetItem(f"{batch.purchase_price:.2f}"))
            self.table.item(i, 0).setData(Qt.UserRole, batch.id)

        btn_layout = QHBoxLayout()
        self.btn_ok = QPushButton("Списать выбранные")
        self.btn_cancel = QPushButton("Отмена")
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_ok)
        btn_layout.addWidget(self.btn_cancel)
        layout.addLayout(btn_layout)

        self.btn_ok.clicked.connect(self.accept_selection)
        self.btn_cancel.clicked.connect(self.reject)

    def filter_table(self):
        text = self.search_edit.text().strip().lower()
        for i, batch in enumerate(self.batches):
            visible = (text in (batch.batch_number or "").lower() or
                       text in str(batch.id).lower())
            self.table.setRowHidden(i, not visible)

    def accept_selection(self):
        for i in range(self.table.rowCount()):
            if self.table.isRowHidden(i):
                continue
            chk = self.table.item(i, 0)
            if chk and chk.checkState() == Qt.Checked:
                batch_id = chk.data(Qt.UserRole)
                batch = next(b for b in self.batches if b.id == batch_id)
                qty, ok = QInputDialog.getDouble(self, "Количество",
                                                 f"Сколько взять из партии {batch.batch_number or batch.id}?",
                                                 batch.quantity, 0.01, batch.quantity, 3)
                if ok and qty > 0:
                    self.selected.append((batch_id, qty))
        if not self.selected:
            QMessageBox.warning(self, "Внимание", "Не выбрано ни одной партии")
            return
        self.accept()

    def get_selected(self):
        return self.selected

# ----------------------------------------------------------------------
# Диалог работы (с материалами и партиями)
# ----------------------------------------------------------------------
class WorkItemDialog(QDialog):
    def __init__(self, work: WorkItem = None, parent=None):
        super().__init__(parent)
        self.work = work if work else WorkItem()
        self.setWindowTitle("Редактирование работы" if work else "Новая работа")
        self.resize(600, 500)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.desc_edit = QLineEdit(self.work.description)
        form.addRow("Описание работы:", self.desc_edit)

        self.hours_spin = QDoubleSpinBox()
        self.hours_spin.setRange(0, 1000)
        self.hours_spin.setValue(self.work.hours)
        form.addRow("Трудоёмкость (часы):", self.hours_spin)

        self.price_spin = QDoubleSpinBox()
        self.price_spin.setRange(0, 1e9)
        self.price_spin.setValue(self.work.price)
        form.addRow("Цена работы:", self.price_spin)

        layout.addLayout(form)

        self.materials_table = QTableWidget(0, 5)
        self.materials_table.setHorizontalHeaderLabels(["Материал", "Количество", "Цена продажи", "Сумма", "Партии"])
        self.materials_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.materials_table)

        btn_layout = QHBoxLayout()
        self.btn_add_mat = QPushButton("Добавить материал")
        self.btn_edit_mat = QPushButton("Изменить")
        self.btn_delete_mat = QPushButton("Удалить")
        btn_layout.addWidget(self.btn_add_mat)
        btn_layout.addWidget(self.btn_edit_mat)
        btn_layout.addWidget(self.btn_delete_mat)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.btn_add_mat.clicked.connect(self.add_material)
        self.btn_edit_mat.clicked.connect(self.edit_material)
        self.btn_delete_mat.clicked.connect(self.delete_material)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.load_materials()

    def load_materials(self):
        self.materials_table.setRowCount(len(self.work.materials))
        total = 0.0
        materials = get_all_materials()
        mat_dict = {m.id: m for m in materials}
        for i, om in enumerate(self.work.materials):
            mat = mat_dict.get(om.material_id)
            name = mat.name if mat else f"ID {om.material_id}"
            self.materials_table.setItem(i, 0, QTableWidgetItem(name))
            self.materials_table.setItem(i, 1, QTableWidgetItem(f"{om.quantity:.3f}"))
            self.materials_table.setItem(i, 2, QTableWidgetItem(f"{om.sale_price:.2f}"))
            summ = om.quantity * om.sale_price
            self.materials_table.setItem(i, 3, QTableWidgetItem(f"{summ:.2f}"))
            if hasattr(om, 'temp_batches') and om.temp_batches:
                batch_info = ", ".join(f"{qty:.3f} (партия {bid})" for bid, qty in om.temp_batches)
            else:
                batch_info = ", ".join(f"{b.quantity:.3f} (партия {b.material_batch_id})" for b in om.batches)
            self.materials_table.setItem(i, 4, QTableWidgetItem(batch_info))
            total += summ

    def add_material(self):
        materials = get_all_materials()
        if not materials:
            QMessageBox.information(self, "Информация", "Сначала добавьте материалы в справочник")
            return
        selected = select_item(
            self,
            "Выберите материал",
            materials,
            ["ID", "Наименование", "Ед.", "Закуп", "Продажа", "Остаток"],
            lambda m: [m.id, m.name, m.unit, m.purchase_price, m.sale_price, m.stock]
        )
        if selected:
            batches = get_available_batches(selected.id)
            if not batches:
                QMessageBox.warning(self, "Ошибка", "Нет доступных партий этого материала")
                return
            batch_dlg = BatchSelectionDialog(selected, batches, self)
            if batch_dlg.exec() == QDialog.DialogCode.Accepted:
                selected_batches = batch_dlg.get_selected()
                total_qty = sum(qty for _, qty in selected_batches)
                dlg = QuantityPriceDialog("Параметры продажи",
                                          label_qty="Общее количество:", label_price="Цена продажи:",
                                          default_qty=total_qty, default_price=selected.sale_price,
                                          qty_max=total_qty, price_max=1e9)
                if dlg.exec() == QDialog.DialogCode.Accepted:
                    qty, price = dlg.get_values()
                    om = OrderMaterial(material_id=selected.id,
                                       quantity=qty,
                                       sale_price=price)
                    om.temp_batches = selected_batches
                    self.work.materials.append(om)
                    self.load_materials()

    def edit_material(self):
        row = self.materials_table.currentRow()
        if row < 0:
            return
        om = self.work.materials[row]
        # Для упрощения: удаляем и предлагаем добавить заново
        self.work.materials.pop(row)
        self.load_materials()
        # Можно вызвать add_material, но это уже сложнее

    def delete_material(self):
        row = self.materials_table.currentRow()
        if row >= 0:
            del self.work.materials[row]
            self.load_materials()

    def get_work(self) -> WorkItem:
        self.work.description = self.desc_edit.text()
        self.work.hours = self.hours_spin.value()
        self.work.price = self.price_spin.value()
        return self.work

# ----------------------------------------------------------------------
# Диалог заказа (использует WorkItemDialog)
# ----------------------------------------------------------------------
class OrderEditDialog(QDialog):
    def __init__(self, order: Order = None, parent=None):
        super().__init__(parent)
        self.order = order if order else Order()
        self.setWindowTitle("Редактирование заказа" if order else "Новый заказ")
        self.resize(800, 600)

        layout = QVBoxLayout(self)

        # Основная информация
        main_group = QGroupBox("Основные данные")
        form = QFormLayout(main_group)

        # Выбор заказчика
        self.customer_display = QLineEdit()
        self.customer_display.setReadOnly(True)
        self.customer_display.setPlaceholderText("Нажмите кнопку для выбора")
        self.btn_select_customer = QPushButton("...")
        self.btn_select_customer.setMaximumWidth(30)
        cust_layout = QHBoxLayout()
        cust_layout.addWidget(self.customer_display)
        cust_layout.addWidget(self.btn_select_customer)
        form.addRow("Заказчик:", cust_layout)
        self.customer_id = self.order.customer_id
        if self.customer_id:
            cust = get_customer_by_id(self.customer_id)
            if cust:
                self.customer_display.setText(cust.name)
        self.btn_select_customer.clicked.connect(self.select_customer)

        self.number_edit = QLineEdit(self.order.order_number)
        form.addRow("Номер счёта:", self.number_edit)

        self.date_edit = QDateEdit()
        if self.order.date:
            self.date_edit.setDate(QDate(self.order.date.year, self.order.date.month, self.order.date.day))
        else:
            self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        form.addRow("Дата:", self.date_edit)

        self.status_combo = QComboBox()
        self.status_combo.addItems(["в работе", "выполнен", "оплачен"])
        if self.order.status:
            self.status_combo.setCurrentText(self.order.status)
        form.addRow("Статус:", self.status_combo)

        self.notes_edit = QTextEdit(self.order.notes)
        self.notes_edit.setMaximumHeight(60)
        form.addRow("Примечание:", self.notes_edit)

        layout.addWidget(main_group)

        # Таблица работ
        works_group = QGroupBox("Работы")
        works_layout = QVBoxLayout(works_group)

        self.works_table = QTableWidget(0, 4)
        self.works_table.setHorizontalHeaderLabels(["Описание", "Часы", "Цена", "Материалы"])
        self.works_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        works_layout.addWidget(self.works_table)

        btn_layout = QHBoxLayout()
        self.btn_add_work = QPushButton("Добавить работу")
        self.btn_edit_work = QPushButton("Изменить")
        self.btn_delete_work = QPushButton("Удалить")
        btn_layout.addWidget(self.btn_add_work)
        btn_layout.addWidget(self.btn_edit_work)
        btn_layout.addWidget(self.btn_delete_work)
        btn_layout.addStretch()
        works_layout.addLayout(btn_layout)

        layout.addWidget(works_group)

        self.btn_add_work.clicked.connect(self.add_work)
        self.btn_edit_work.clicked.connect(self.edit_work)
        self.btn_delete_work.clicked.connect(self.delete_work)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.load_works()

    def select_customer(self):
        customers = get_all_customers()
        if not customers:
            QMessageBox.information(self, "Информация", "Нет заказчиков в справочнике")
            return
        selected = select_item(
            self,
            "Выберите заказчика",
            customers,
            ["ID", "Наименование", "Телефон", "Email"],
            lambda c: [c.id, c.name, c.phone, c.email]
        )
        if selected:
            self.customer_display.setText(selected.name)
            self.customer_id = selected.id

    def load_works(self):
        self.works_table.setRowCount(len(self.order.work_items))
        for i, work in enumerate(self.order.work_items):
            self.works_table.setItem(i, 0, QTableWidgetItem(work.description))
            self.works_table.setItem(i, 1, QTableWidgetItem(f"{work.hours:.2f}"))
            self.works_table.setItem(i, 2, QTableWidgetItem(f"{work.price:.2f}"))
            mats = ", ".join([f"{om.quantity} {get_material_by_id(om.material_id).unit}" for om in work.materials])
            self.works_table.setItem(i, 3, QTableWidgetItem(mats))

    def add_work(self):
        dlg = WorkItemDialog()
        if dlg.exec() == QDialog.DialogCode.Accepted:
            work = dlg.get_work()
            self.order.work_items.append(work)
            self.load_works()

    def edit_work(self):
        row = self.works_table.currentRow()
        if row < 0:
            return
        work = self.order.work_items[row]
        dlg = WorkItemDialog(work)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            updated = dlg.get_work()
            self.order.work_items[row] = updated
            self.load_works()

    def delete_work(self):
        row = self.works_table.currentRow()
        if row >= 0:
            del self.order.work_items[row]
            self.load_works()

    def accept(self):
        if not self.customer_id:
            QMessageBox.warning(self, "Ошибка", "Выберите заказчика")
            return
        if not self.order.work_items:
            QMessageBox.warning(self, "Ошибка", "Добавьте хотя бы одну работу")
            return
        self.order.customer_id = self.customer_id
        self.order.customer_name = self.customer_display.text()
        self.order.order_number = self.number_edit.text()
        self.order.date = self.date_edit.date().toPython()
        self.order.status = self.status_combo.currentText()
        self.order.notes = self.notes_edit.toPlainText()
        super().accept()