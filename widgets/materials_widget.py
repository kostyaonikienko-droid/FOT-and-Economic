# widgets/materials_widget.py

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QDialog, QLineEdit, QLabel
)
from database.repair_db import get_all_materials, delete_material
from dialogs.material_dialogs import MaterialEditDialog
from dialogs.write_off_dialog import WriteOffDialog

class MaterialsWidget(QWidget):
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
        self.btn_write_off = QPushButton("Акт списания")
        self.btn_write_off.clicked.connect(self.open_write_off)
        btn_layout.addWidget(self.btn_write_off)

        btn_layout.addWidget(self.btn_refresh)
        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_edit)
        btn_layout.addWidget(self.btn_delete)
        btn_layout.addWidget(self.btn_select_all)
        btn_layout.addWidget(self.btn_clear_all)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # Таблица: чекбокс, ID (скрыт), наименование, инв.№, ед., тип, закуп, продажа, остаток
        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels(["", "ID", "Наименование", "Инв. №", "Ед.", "Тип", "Закуп", "Продажа", "Остаток"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 30)  # минимальная ширина для чекбокса
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.hideColumn(1)  # скрываем ID
        layout.addWidget(self.table)

        self.all_materials = []

        self.btn_refresh.clicked.connect(self.load_materials)
        self.btn_add.clicked.connect(self.add_material)
        self.btn_edit.clicked.connect(self.edit_material)
        self.btn_delete.clicked.connect(self.delete_material)
        self.btn_select_all.clicked.connect(self.select_all)
        self.btn_clear_all.clicked.connect(self.clear_all)

        self.load_materials()

    def open_write_off(self):
        from dialogs.write_off_dialog import WriteOffDialog
        dlg = WriteOffDialog(self)
        dlg.exec()

    def load_materials(self):
        self.all_materials = get_all_materials()
        self.filter_table()

    def filter_table(self):
        text = self.search_edit.text().strip().lower()
        filtered = []
        for m in self.all_materials:
            if (text in m.name.lower() or
                text in m.inventory_number.lower() or
                text in m.unit.lower() or
                text in m.type.lower() or
                text in str(m.purchase_price) or
                text in str(m.sale_price) or
                text in str(m.stock)):
                filtered.append(m)

        self.table.setRowCount(len(filtered))
        for i, m in enumerate(filtered):
            chk = QTableWidgetItem()
            chk.setCheckState(Qt.Unchecked)
            self.table.setItem(i, 0, chk)
            self.table.setItem(i, 1, QTableWidgetItem(str(m.id)))
            self.table.setItem(i, 2, QTableWidgetItem(m.name))
            self.table.setItem(i, 3, QTableWidgetItem(m.inventory_number))
            self.table.setItem(i, 4, QTableWidgetItem(m.unit))
            self.table.setItem(i, 5, QTableWidgetItem(m.type))
            self.table.setItem(i, 6, QTableWidgetItem(f"{m.purchase_price:.2f}"))
            self.table.setItem(i, 7, QTableWidgetItem(f"{m.sale_price:.2f}"))
            self.table.setItem(i, 8, QTableWidgetItem(f"{m.stock:.2f}"))

    def add_material(self):
        dlg = MaterialEditDialog()
        if dlg.exec() == QDialog.DialogCode.Accepted:
            from database.repair_db import add_material
            add_material(dlg.get_material())
            self.load_materials()

    def edit_material(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Внимание", "Выберите материал для редактирования")
            return
        material_id = int(self.table.item(row, 1).text())
        materials = get_all_materials()
        material = next((m for m in materials if m.id == material_id), None)
        if not material:
            return
        dlg = MaterialEditDialog(material)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            from database.repair_db import update_material
            update_material(dlg.get_material())
            self.load_materials()

    def delete_material(self):
        selected_rows = []
        for row in range(self.table.rowCount()):
            if self.table.item(row, 0).checkState() == Qt.Checked:
                selected_rows.append(row)
        if not selected_rows:
            QMessageBox.warning(self, "Внимание", "Выберите хотя бы один материал для удаления")
            return
        confirm = QMessageBox.question(self, "Подтверждение", f"Удалить {len(selected_rows)} материал(ов)?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm != QMessageBox.StandardButton.Yes:
            return
        for row in selected_rows:
            material_id = int(self.table.item(row, 1).text())
            delete_material(material_id)
        self.load_materials()

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