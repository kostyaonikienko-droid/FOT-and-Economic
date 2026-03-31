from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget,
    QTableWidgetItem, QHeaderView, QMessageBox, QComboBox, QDateEdit,
    QLabel, QInputDialog, QLineEdit
)
from PySide6.QtCore import Qt, QDate
from . import db_manager

class DatabaseEditorDialog(QDialog):
    def __init__(self, fot_widget=None, parent=None):
        super().__init__(parent)
        self.fot_widget = fot_widget
        self.setWindowTitle("Редактор базы данных (спирт)")
        self.resize(1100, 600)
        layout = QVBoxLayout(self)

        # Панель фильтров
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Год:"))
        self.year_combo = QComboBox()
        self.year_combo.addItem("Все", None)
        for y in range(2020, 2031):
            self.year_combo.addItem(str(y), y)
        filter_layout.addWidget(self.year_combo)

        filter_layout.addWidget(QLabel("Месяц:"))
        self.month_combo = QComboBox()
        self.month_combo.addItem("Все", 0)
        months = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
                  "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
        for i, m in enumerate(months, 1):
            self.month_combo.addItem(m, i)
        filter_layout.addWidget(self.month_combo)

        filter_layout.addWidget(QLabel("Исполнитель (таб. №):"))
        self.tab_combo = QComboBox()
        self.tab_combo.addItem("Все", None)
        filter_layout.addWidget(self.tab_combo)

        filter_layout.addWidget(QLabel("Поиск:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Поиск по любому столбцу")
        filter_layout.addWidget(self.search_edit)

        self.btn_filter = QPushButton("Применить")
        self.btn_filter.clicked.connect(self.load_data)
        filter_layout.addWidget(self.btn_filter)
        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        # Таблица
        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels(["", "ID", "Дата вып", "Дата сч", "№ сч", "Работа", "Кол-во", "Таб. №", "Исполнитель", "Группа"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked)
        self.table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.table.cellDoubleClicked.connect(self.edit_quantity_cell)
        self.table.itemClicked.connect(self.on_item_clicked)
        self.table.setStyleSheet("QTableWidget::indicator { width: 16px; height: 16px; }")

        layout.addWidget(self.table)

        # Панель кнопок
        btn_layout = QHBoxLayout()
        self.btn_select_all = QPushButton("Выбрать все")
        self.btn_clear_all = QPushButton("Снять выделение")
        self.btn_edit = QPushButton("Редактировать количество")
        self.btn_delete = QPushButton("Удалить выбранные")
        self.btn_refresh = QPushButton("Обновить")
        btn_layout.addWidget(self.btn_select_all)
        btn_layout.addWidget(self.btn_clear_all)
        btn_layout.addWidget(self.btn_edit)
        btn_layout.addWidget(self.btn_delete)
        btn_layout.addWidget(self.btn_refresh)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.btn_select_all.clicked.connect(self.select_all)
        self.btn_clear_all.clicked.connect(self.clear_all)
        self.btn_edit.clicked.connect(self.edit_quantity)
        self.btn_delete.clicked.connect(self.delete_selected)
        self.btn_refresh.clicked.connect(self.load_data)

        self.load_data()

    def load_data(self):
        works = db_manager.get_all_works()
        year = self.year_combo.currentData()
        month = self.month_combo.currentData()
        tab_filter = self.tab_combo.currentData()
        search_text = self.search_edit.text().strip().lower()
        filtered = []
        for w in works:
            if year:
                try:
                    w_year = int(w['date_work'].split('.')[-1])
                    if w_year != year:
                        continue
                except:
                    continue
            if month:
                try:
                    w_month = int(w['date_work'].split('.')[1])
                    if w_month != month:
                        continue
                except:
                    continue
            if tab_filter and w.get('tab_number') != tab_filter:
                continue
            if search_text:
                row_text = f"{w['date_work']} {w['date_invoice']} {w['invoice_num']} {w['work_name']} {w.get('tab_number','')}".lower()
                if search_text not in row_text:
                    continue
            filtered.append(w)

        self.table.setRowCount(len(filtered))
        for i, w in enumerate(filtered):
            # Чекбокс (стандартный, как в справочниках)
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemFlag.ItemIsUserCheckable)
            chk.setCheckState(Qt.Unchecked)
            self.table.setItem(i, 0, chk)
            # ID
            self.table.setItem(i, 1, QTableWidgetItem(str(w['id'])))
            # Дата вып
            self.table.setItem(i, 2, QTableWidgetItem(w['date_work']))
            # Дата сч
            self.table.setItem(i, 3, QTableWidgetItem(w['date_invoice']))
            # № сч
            self.table.setItem(i, 4, QTableWidgetItem(w['invoice_num']))
            # Работа
            self.table.setItem(i, 5, QTableWidgetItem(w['work_name']))
            # Кол-во
            qty_item = QTableWidgetItem(str(w.get('quantity', 1)))
            qty_item.setData(Qt.UserRole, w['id'])
            self.table.setItem(i, 6, qty_item)
            # Таб. №
            self.table.setItem(i, 7, QTableWidgetItem(w.get('tab_number', '')))
            # Исполнитель
            emp_name = self.get_employee_name(w.get('tab_number', ''))
            self.table.setItem(i, 8, QTableWidgetItem(emp_name))
            # Группа
            self.table.setItem(i, 9, QTableWidgetItem(w.get('group_name', '')))

        # Заполнить комбобокс табельных номеров
        tab_numbers = set(w.get('tab_number') for w in works if w.get('tab_number'))
        self.tab_combo.clear()
        self.tab_combo.addItem("Все", None)
        for tn in sorted(tab_numbers):
            emp_name = self.get_employee_name(tn)
            self.tab_combo.addItem(f"{tn} - {emp_name}", tn)

    def on_item_clicked(self, item):
        """Обработчик клика по ячейке: переключает чекбокс, если клик в первом столбце."""
        if item.column() == 0:
            # Переключаем состояние чекбокса
            if item.checkState() == Qt.Checked:
                item.setCheckState(Qt.Unchecked)
            else:
                item.setCheckState(Qt.Checked)

    def get_employee_name(self, tab_number):
        if not tab_number:
            return ""
        if self.fot_widget and hasattr(self.fot_widget, 'project'):
            for emp in self.fot_widget.project.employees:
                if emp.tab_num == tab_number:
                    return emp.fio
        return ""

    def select_all(self):
        for row in range(self.table.rowCount()):
            self.table.item(row, 0).setCheckState(Qt.Checked)

    def clear_all(self):
        for row in range(self.table.rowCount()):
            self.table.item(row, 0).setCheckState(Qt.Unchecked)

    def edit_quantity_cell(self, row, col):
        if col == 6:
            self.edit_quantity()

    def edit_quantity(self):
        current = self.table.currentRow()
        if current < 0:
            QMessageBox.warning(self, "Внимание", "Выберите запись для редактирования")
            return
        qty_item = self.table.item(current, 6)
        if not qty_item:
            return
        work_id = int(qty_item.data(Qt.UserRole))
        old_qty = int(qty_item.text())
        new_qty, ok = QInputDialog.getInt(self, "Редактирование количества",
                                          "Введите новое количество:", old_qty, 1, 1000000)
        if ok:
            db_manager.update_work_quantity(work_id, new_qty)
            qty_item.setText(str(new_qty))
            QMessageBox.information(self, "Успех", "Количество обновлено")

    def delete_selected(self):
        selected = []
        for row in range(self.table.rowCount()):
            if self.table.item(row, 0).checkState() == Qt.Checked:
                work_id = int(self.table.item(row, 1).text())
                selected.append(work_id)
        if not selected:
            QMessageBox.warning(self, "Внимание", "Выберите записи для удаления")
            return
        confirm = QMessageBox.question(self, "Подтверждение", f"Удалить {len(selected)} записей?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.Yes:
            db_manager.delete_works_by_ids(selected)
            self.load_data()
            QMessageBox.information(self, "Успех", f"Удалено {len(selected)} записей")