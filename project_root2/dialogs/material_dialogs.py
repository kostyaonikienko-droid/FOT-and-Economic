# dialogs/material_dialogs.py

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QDoubleSpinBox, QDialogButtonBox,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QPushButton, QMessageBox, QComboBox, QLabel
)
from models.repair_models import Material
from database.repair_db import (
    get_all_materials, get_material_by_id, add_material, update_material, delete_material
)

class MaterialEditDialog(QDialog):
    def __init__(self, material: Material = None, parent=None):
        super().__init__(parent)
        self.material = material if material else Material()
        self.setWindowTitle("Редактирование" if material else "Новый")
        self.resize(500, 450)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.name_edit = QLineEdit(self.material.name)
        form.addRow("Наименование:", self.name_edit)

        self.inventory_edit = QLineEdit(self.material.inventory_number)
        form.addRow("Инвентарный номер:", self.inventory_edit)

        self.unit_edit = QLineEdit(self.material.unit)
        form.addRow("Ед. изм.:", self.unit_edit)

        self.type_combo = QComboBox()
        self.type_combo.addItems(["материал", "гсо", "расходник"])
        self.type_combo.setCurrentText(self.material.type)
        form.addRow("Тип:", self.type_combo)

        self.purchase_price_spin = QDoubleSpinBox()
        self.purchase_price_spin.setRange(0, 1e9)
        self.purchase_price_spin.setValue(self.material.purchase_price)
        form.addRow("Цена закупки:", self.purchase_price_spin)

        self.sale_price_spin = QDoubleSpinBox()
        self.sale_price_spin.setRange(0, 1e9)
        self.sale_price_spin.setValue(self.material.sale_price)
        form.addRow("Цена продажи:", self.sale_price_spin)

        self.stock_spin = QDoubleSpinBox()
        self.stock_spin.setRange(0, 1e9)
        self.stock_spin.setValue(self.material.stock)
        self.stock_spin.setReadOnly(True)
        form.addRow("Текущий остаток:", self.stock_spin)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_material(self) -> Material:
        self.material.name = self.name_edit.text()
        self.material.inventory_number = self.inventory_edit.text()
        self.material.unit = self.unit_edit.text()
        self.material.type = self.type_combo.currentText()
        self.material.purchase_price = self.purchase_price_spin.value()
        self.material.sale_price = self.sale_price_spin.value()
        return self.material

class MaterialsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Справочник материалов")
        self.resize(900, 500)

        layout = QVBoxLayout(self)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["Наименование", "Инв. №", "Ед.", "Тип", "Закуп", "Продажа", "Остаток"])
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

        self.btn_add.clicked.connect(self.add_material)
        self.btn_edit.clicked.connect(self.edit_material)
        self.btn_delete.clicked.connect(self.delete_material)
        self.btn_close.clicked.connect(self.accept)

        self.load_materials()

    def load_materials(self):
        materials = get_all_materials()
        self.table.setRowCount(len(materials))
        for i, m in enumerate(materials):
            self.table.setItem(i, 0, QTableWidgetItem(m.name))
            self.table.setItem(i, 1, QTableWidgetItem(m.inventory_number))
            self.table.setItem(i, 2, QTableWidgetItem(m.unit))
            self.table.setItem(i, 3, QTableWidgetItem(m.type))  # новый столбец
            self.table.setItem(i, 4, QTableWidgetItem(f"{m.purchase_price:.2f}"))
            self.table.setItem(i, 5, QTableWidgetItem(f"{m.sale_price:.2f}"))
            self.table.setItem(i, 6, QTableWidgetItem(f"{m.stock:.2f}"))

    def add_material(self):
        dlg = MaterialEditDialog()
        if dlg.exec() == QDialog.DialogCode.Accepted:
            add_material(dlg.get_material())
            self.load_materials()

    def edit_material(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Внимание", "Выберите материал для редактирования")
            return
        name = self.table.item(row, 0).text()
        inv = self.table.item(row, 1).text()
        materials = get_all_materials()
        material = next((m for m in materials if m.name == name and m.inventory_number == inv), None)
        if not material:
            return
        dlg = MaterialEditDialog(material)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            update_material(dlg.get_material())
            self.load_materials()

    def delete_material(self):
        row = self.table.currentRow()
        if row < 0:
            return
        confirm = QMessageBox.question(self, "Подтверждение", "Удалить выбранный материал?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.Yes:
            name = self.table.item(row, 0).text()
            inv = self.table.item(row, 1).text()
            materials = get_all_materials()
            material = next((m for m in materials if m.name == name and m.inventory_number == inv), None)
            if material:
                delete_material(material.id)
                self.load_materials()