# widgets/repair_widget.py

import datetime
from PySide6.QtCore import Qt, QDate
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QFileDialog, QMenu, QLabel, QComboBox,
    QLineEdit, QDialog, QFormLayout, QDateEdit, QDialogButtonBox
)
from PySide6.QtGui import QAction, QColor, QFont
from database.repair_db import get_all_orders, delete_order, add_order, update_order, get_order_by_id, get_material_by_id
from dialogs.order_dialogs import OrderEditDialog
from dialogs.material_dialogs import MaterialsDialog
from dialogs.customer_dialogs import CustomersDialog
from .suppliers_widget import SuppliersWidget  # если нужно

class RepairWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)

        # Фильтры
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Год:"))
        self.filter_year_combo = QComboBox()
        current_year = datetime.date.today().year
        for y in range(2020, 2031):
            self.filter_year_combo.addItem(str(y), y)
        self.filter_year_combo.setCurrentText(str(current_year))
        filter_layout.addWidget(self.filter_year_combo)

        filter_layout.addWidget(QLabel("Месяц:"))
        self.filter_month_combo = QComboBox()
        self.filter_month_combo.addItem("Все", 0)
        months = ["Январь","Февраль","Март","Апрель","Май","Июнь",
                  "Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"]
        for i, m in enumerate(months, 1):
            self.filter_month_combo.addItem(m, i)
        filter_layout.addWidget(self.filter_month_combo)

        self.filter_customer_edit = QLineEdit()
        self.filter_customer_edit.setPlaceholderText("Фильтр по заказчику")
        filter_layout.addWidget(self.filter_customer_edit)

        self.btn_reset = QPushButton("Сброс")
        self.btn_reset.clicked.connect(self.reset_filters)
        filter_layout.addWidget(self.btn_reset)

        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        # Панель кнопок
        btn_layout = QHBoxLayout()

        self.btn_refresh = QPushButton("Обновить")
        self.btn_add = QPushButton("Добавить заказ")
        self.btn_edit = QPushButton("Редактировать")
        self.btn_delete = QPushButton("Удалить")
        self.btn_materials = QPushButton("Материалы")
        self.btn_customers = QPushButton("Заказчики")
        self.btn_act = QPushButton("Акт списания")

        btn_layout.addWidget(self.btn_act)
        btn_layout.addWidget(self.btn_refresh)
        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_edit)
        btn_layout.addWidget(self.btn_delete)
        btn_layout.addWidget(self.btn_materials)
        btn_layout.addWidget(self.btn_customers)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # Таблица заказов (10 столбцов)
        self.table = QTableWidget(0, 10)
        self.table.setHorizontalHeaderLabels([
            "Номер счёта", "Дата", "Заказчик", "Статус", "", "Доход",
            "Расход", "Прибыль", "Себестоимость", "ID"
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.hideColumn(9)  # скрываем ID
        layout.addWidget(self.table)

        # Контекстное меню
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        # Подключение сигналов
        self.btn_act.clicked.connect(self.open_write_off_dialog)
        self.btn_refresh.clicked.connect(self.load_orders)
        self.btn_add.clicked.connect(self.add_order)
        self.btn_edit.clicked.connect(self.edit_order)
        self.btn_delete.clicked.connect(self.delete_order)
        self.btn_materials.clicked.connect(self.open_materials)
        self.btn_customers.clicked.connect(self.open_customers)

        self.filter_year_combo.currentIndexChanged.connect(self.load_orders)
        self.filter_month_combo.currentIndexChanged.connect(self.load_orders)
        self.filter_customer_edit.textChanged.connect(self.load_orders)

        self.load_orders()

    def open_write_off_dialog(self):
        from dialogs.select_orders_dialog import SelectOrdersDialog
        dlg = SelectOrdersDialog(self)
        dlg.exec()

    def reset_filters(self):
        self.filter_year_combo.setCurrentText(str(datetime.date.today().year))
        self.filter_month_combo.setCurrentIndex(0)
        self.filter_customer_edit.clear()
        self.load_orders()

    def load_orders(self):
        all_orders = get_all_orders()
        year = int(self.filter_year_combo.currentText())
        month_idx = self.filter_month_combo.currentData()
        customer_filter = self.filter_customer_edit.text().strip().lower()

        filtered = []
        for order in all_orders:
            if order.date.year != year:
                continue
            if month_idx and order.date.month != month_idx:
                continue
            if customer_filter and customer_filter not in order.customer_name.lower():
                continue
            filtered.append(order)

        self.table.setRowCount(len(filtered))
        for i, order in enumerate(filtered):
            # Расчёт дохода, себестоимости, прибыли
            income = 0.0
            cost = 0.0
            for work in order.work_items:
                income += work.price
                for om in work.materials:
                    income += om.quantity * om.sale_price
                    for batch in om.batches:
                        cost += batch.quantity * batch.purchase_price
            profit = income - cost

            self.table.setItem(i, 0, QTableWidgetItem(order.order_number))
            self.table.setItem(i, 1, QTableWidgetItem(order.date.strftime("%d.%m.%Y")))
            self.table.setItem(i, 2, QTableWidgetItem(order.customer_name))
            self.table.setItem(i, 3, QTableWidgetItem(order.status))

            # Индикатор статуса
            indicator = QTableWidgetItem("●")
            indicator.setTextAlignment(Qt.AlignCenter)
            if order.status == "в работе":
                indicator.setForeground(QColor(255, 165, 0))
                indicator.setToolTip("в работе")
            elif order.status == "выполнен":
                indicator.setForeground(QColor(0, 128, 0))
                indicator.setToolTip("выполнен")
            elif order.status == "оплачен":
                indicator.setForeground(QColor(0, 0, 255))
                indicator.setToolTip("оплачен")
            else:
                indicator.setForeground(QColor(128, 128, 128))
            self.table.setItem(i, 4, indicator)

            self.table.setItem(i, 5, QTableWidgetItem(f"{income:.2f}"))
            self.table.setItem(i, 6, QTableWidgetItem(f"{cost:.2f}"))
            self.table.setItem(i, 7, QTableWidgetItem(f"{profit:.2f}"))
            self.table.setItem(i, 8, QTableWidgetItem(f"{cost:.2f}"))
            self.table.setItem(i, 9, QTableWidgetItem(str(order.id)))

    def add_order(self):
        dlg = OrderEditDialog()
        if dlg.exec() == QDialog.DialogCode.Accepted:
            add_order(dlg.order)
            self.load_orders()

    def edit_order(self):
        rows = set(idx.row() for idx in self.table.selectedIndexes())
        if len(rows) != 1:
            QMessageBox.warning(self, "Внимание", "Выберите один заказ для редактирования")
            return
        row = list(rows)[0]
        order_id = int(self.table.item(row, 9).text())
        order = get_order_by_id(order_id)
        if not order:
            return
        dlg = OrderEditDialog(order)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            update_order(dlg.order)
            self.load_orders()

    def delete_order(self):
        rows = set(idx.row() for idx in self.table.selectedIndexes())
        if not rows:
            return
        confirm = QMessageBox.question(self, "Подтверждение",
                                       f"Удалить {len(rows)} заказ(ов)?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm != QMessageBox.StandardButton.Yes:
            return
        for row in rows:
            order_id = int(self.table.item(row, 9).text())
            delete_order(order_id)
        self.load_orders()

    def open_materials(self):
        dlg = MaterialsDialog(self)
        dlg.exec()

    def open_customers(self):
        dlg = CustomersDialog(self)
        dlg.exec()

    def show_context_menu(self, position):
        menu = QMenu()
        add_action = menu.addAction("Добавить заказ")
        add_action.triggered.connect(self.add_order)

        if self.table.selectedIndexes():
            edit_action = menu.addAction("Редактировать")
            edit_action.triggered.connect(self.edit_order)

            delete_action = menu.addAction("Удалить")
            delete_action.triggered.connect(self.delete_order)

        menu.exec(self.table.viewport().mapToGlobal(position))

    def populate_menu(self, menu_bar):
        file_menu = menu_bar.addMenu("Файл")
        new_order = QAction("Новый заказ", self)
        new_order.triggered.connect(self.add_order)
        file_menu.addAction(new_order)

        ref_menu = menu_bar.addMenu("Справочники")
        mat_action = QAction("Материалы", self)
        mat_action.triggered.connect(self.open_materials)
        ref_menu.addAction(mat_action)

        cust_action = QAction("Заказчики", self)
        cust_action.triggered.connect(self.open_customers)
        ref_menu.addAction(cust_action)

        # Можно добавить другие пункты
        pass

    def maybe_save(self):
        return True