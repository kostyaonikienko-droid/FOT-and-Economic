# dialogs/common.py

from PySide6.QtWidgets import QDialog, QVBoxLayout, QLineEdit, QTableWidget, QTableWidgetItem, QHBoxLayout, QPushButton, QHeaderView, QMessageBox, QFormLayout, QDoubleSpinBox, QDialogButtonBox
from PySide6.QtCore import Qt

class SelectionDialog(QDialog):
    def __init__(self, title: str, items: list, headers: list, get_row_func, parent=None, enable_new=False):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(700, 500)
        self.items = items
        self.get_row_func = get_row_func
        self.selected_item = None
        self.enable_new = enable_new

        layout = QVBoxLayout(self)
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Введите текст для поиска...")
        self.search_edit.textChanged.connect(self.filter_items)
        layout.addWidget(self.search_edit)

        self.table = QTableWidget()
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.doubleClicked.connect(self.accept_selection)
        layout.addWidget(self.table)

        btn_layout = QHBoxLayout()
        # Добавляем кнопку "Новый"
        if enable_new:
            self.btn_new = QPushButton("Новый")
            self.btn_new.clicked.connect(self.new_item)
            btn_layout.addWidget(self.btn_new)
        btn_ok = QPushButton("Выбрать")
        btn_ok.clicked.connect(self.accept_selection)
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

        self.filtered_items = items.copy()
        self.refresh_table()

    def new_item(self):
        # Возвращаем специальный флаг
        self.selected_item = "NEW"
        self.accept()

    def refresh_table(self):
        self.table.setRowCount(len(self.filtered_items))
        for i, item in enumerate(self.filtered_items):
            row_data = self.get_row_func(item)
            for j, val in enumerate(row_data):
                self.table.setItem(i, j, QTableWidgetItem(str(val)))
            self.table.item(i, 0).setData(Qt.UserRole, item)

    def filter_items(self, text):
        text = text.strip().lower()
        if not text:
            self.filtered_items = self.items.copy()
        else:
            self.filtered_items = [item for item in self.items if
                                   any(text in str(getattr(item, field, '')).lower()
                                       for field in ['name', 'phone', 'email', 'address', 'agreement_number'])]
        self.refresh_table()

    def accept_selection(self):
        selected = self.table.currentRow()
        if selected >= 0:
            self.selected_item = self.filtered_items[selected]
            self.accept()
        else:
            QMessageBox.warning(self, "Внимание", "Выберите запись")

    def get_selected(self):
        return self.selected_item

def select_item(parent, title, items, headers, get_row_func, enable_new=False):
    dlg = SelectionDialog(title, items, headers, get_row_func, parent, enable_new)
    if dlg.exec() == QDialog.DialogCode.Accepted:
        return dlg.get_selected()
    return None

class QuantityPriceDialog(QDialog):
    def __init__(self, title: str, label_qty: str = "Количество:", label_price: str = "Цена:",
                 default_qty=1.0, default_price=0.0, qty_max=1e9, price_max=1e9, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(300, 150)
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.qty_spin = QDoubleSpinBox()
        self.qty_spin.setRange(0.01, qty_max)
        self.qty_spin.setValue(default_qty)
        self.qty_spin.setDecimals(3)
        self.price_spin = QDoubleSpinBox()
        self.price_spin.setRange(0, price_max)
        self.price_spin.setValue(default_price)
        self.price_spin.setDecimals(2)
        form.addRow(label_qty, self.qty_spin)
        form.addRow(label_price, self.price_spin)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_values(self):
        return self.qty_spin.value(), self.price_spin.value()


