import os
from datetime import datetime
import calendar
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QLabel, QLineEdit, QTreeWidget, QTreeWidgetItem, QMessageBox,
    QSplitter, QDateEdit, QDialog, QFormLayout, QDialogButtonBox,
    QInputDialog, QRadioButton, QDoubleSpinBox, QComboBox, QGroupBox,
    QApplication, QTextEdit, QMenu
)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QAction
from . import excel_parser
from . import db_manager
from . import grouping
from . import act_generator

class ActWidget(QWidget):
    def __init__(self, parent=None, fot_widget=None):
        super().__init__(parent)
        self.fot_widget = fot_widget
        self.current_works = []
        self.filtered_works = []
        self.grouped_works = {}
        self.init_ui()
        db_manager.init_db()

    def init_ui(self):
        main_layout = QVBoxLayout()

        top_panel = QHBoxLayout()
        self.btn_load = QPushButton("Загрузить Excel")
        self.btn_load.clicked.connect(self.load_excel)
        top_panel.addWidget(self.btn_load)

        self.label_file = QLabel("Файл не выбран")
        top_panel.addWidget(self.label_file)

        top_panel.addStretch()

        self.label_keywords = QLabel("Ключевые слова (через запятую):")
        top_panel.addWidget(self.label_keywords)
        self.edit_keywords = QLineEdit()
        top_panel.addWidget(self.edit_keywords)

        self.btn_filter = QPushButton("Фильтровать")
        self.btn_filter.clicked.connect(self.filter_works)
        top_panel.addWidget(self.btn_filter)

        self.btn_clear = QPushButton("Очистить")
        self.btn_clear.clicked.connect(self.clear_filter)
        top_panel.addWidget(self.btn_clear)

        main_layout.addLayout(top_panel)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Вкл.", "Группа / Работа", "Дата вып", "Дата сч", "№ сч", "Сумма", "Кол-во"])
        self.tree.setColumnWidth(0, 50)  # ширина для чекбокса
        self.tree.setColumnWidth(1, 300)
        self.tree.setEditTriggers(QTreeWidget.EditTrigger.DoubleClicked)
        self.tree.itemDoubleClicked.connect(self.edit_quantity)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.show_context_menu)
        splitter.addWidget(self.tree)

        right_panel = QVBoxLayout()
        self.btn_save_db = QPushButton("Сохранить выбранное в базу")
        self.btn_save_db.setFixedWidth(180)
        self.btn_save_db.clicked.connect(self.save_selected_to_db)
        right_panel.addWidget(self.btn_save_db)

        self.btn_generate_act = QPushButton("Сформировать акт из базы")
        self.btn_generate_act.clicked.connect(self.generate_act_from_db)
        right_panel.addWidget(self.btn_generate_act)

        self.btn_edit_db = QPushButton("Редактировать базу данных")
        self.btn_edit_db.clicked.connect(self.open_database_editor)
        right_panel.addWidget(self.btn_edit_db)

        self.btn_settings = QPushButton("Настройки акта")
        self.btn_settings.clicked.connect(self.open_settings)
        right_panel.addWidget(self.btn_settings)

        right_widget = QWidget()
        right_widget.setLayout(right_panel)
        splitter.addWidget(right_widget)

        splitter.setStretchFactor(0, 8)
        splitter.setStretchFactor(1, 2)

        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите Excel-файл", "", "Excel files (*.xls *.xlsx)")
        if not file_path:
            return
        try:
            works = excel_parser.parse_excel(file_path)
            if not works:
                QMessageBox.warning(self, "Предупреждение", "В файле не найдено данных работ.")
                return
            self.current_works = works
            self.label_file.setText(os.path.basename(file_path))
            QMessageBox.information(self, "Успех", f"Загружено {len(works)} записей.")
            self.filter_works()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось обработать файл:\n{str(e)}")

    def filter_works(self):
        if not self.current_works:
            QMessageBox.warning(self, "Предупреждение", "Сначала загрузите файл.")
            return
        keywords = self.edit_keywords.text().strip()
        if not keywords:
            filtered = self.current_works[:]
        else:
            kw_list = [kw.strip().lower() for kw in keywords.split(',')]
            filtered = []
            for w in self.current_works:
                work_name = w['work_name'].lower()
                if any(kw in work_name for kw in kw_list):
                    filtered.append(w)
        self.filtered_works = filtered
        self.grouped_works = grouping.group_works(self.filtered_works)
        self.update_preview()

    def clear_filter(self):
        self.edit_keywords.clear()
        self.filter_works()

    def update_preview(self):
        self.tree.clear()
        for group_name, works in self.grouped_works.items():
            group_item = QTreeWidgetItem(self.tree)
            group_item.setText(1, group_name)
            group_item.setFlags(group_item.flags() | Qt.ItemFlag.ItemIsAutoTristate)
            for w in works:
                work_text = w['work_name'][:50]
                item = QTreeWidgetItem(group_item)
                # Чекбокс
                item.setCheckState(0, Qt.Checked)
                item.setText(1, work_text)
                item.setText(2, w['date_work'])
                item.setText(3, w['date_invoice'])
                item.setText(4, w['invoice_num'])
                item.setText(5, f"{w['amount']:.2f}" if isinstance(w['amount'], (int, float)) else str(w['amount']))
                item.setText(6, str(w.get('quantity', 1)))
                item.setData(0, Qt.ItemDataRole.UserRole, w)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
            group_item.setExpanded(True)
        self.tree.resizeColumnToContents(1)

    def edit_quantity(self, item, column):
        if column == 6:  # колонка "Кол-во"
            w = item.data(0, Qt.ItemDataRole.UserRole)
            if w:
                old_qty = w.get('quantity', 1)
                new_qty, ok = QInputDialog.getInt(self, "Редактирование количества",
                                                  "Введите новое количество:", old_qty, 1, 1000000)
                if ok:
                    w['quantity'] = new_qty
                    item.setText(6, str(new_qty))
                    for fw in self.filtered_works:
                        if (fw['date_work'] == w['date_work'] and
                            fw['date_invoice'] == w['date_invoice'] and
                            fw['invoice_num'] == w['invoice_num'] and
                            fw['work_name'] == w['work_name']):
                            fw['quantity'] = new_qty
                            break
                    for cw in self.current_works:
                        if (cw['date_work'] == w['date_work'] and
                            cw['date_invoice'] == w['date_invoice'] and
                            cw['invoice_num'] == w['invoice_num'] and
                            cw['work_name'] == w['work_name']):
                            cw['quantity'] = new_qty
                            break

    def show_context_menu(self, pos):
        item = self.tree.itemAt(pos)
        if not item or item.parent() is None:
            return
        menu = QMenu()
        edit_action = menu.addAction("Редактировать количество")
        action = menu.exec(self.tree.viewport().mapToGlobal(pos))
        if action == edit_action:
            self.edit_quantity(item, 6)

    def save_selected_to_db(self):
        selected = []
        root = self.tree.invisibleRootItem()
        for i in range(root.childCount()):
            group_item = root.child(i)
            for j in range(group_item.childCount()):
                item = group_item.child(j)
                if item.checkState(0) == Qt.Checked:
                    w = item.data(0, Qt.ItemDataRole.UserRole)
                    if w:
                        selected.append(w)
        if not selected:
            QMessageBox.warning(self, "Предупреждение", "Нет выбранных записей для сохранения.")
            return
        for w in selected:
            w['group_name'] = grouping.determine_group(w['work_name']) or ''
        import_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        db_manager.save_works(selected, import_date)
        QMessageBox.information(self, "Успех", f"Сохранено {len(selected)} записей в базу данных.")

    def open_database_editor(self):
        from .database_editor import DatabaseEditorDialog
        dlg = DatabaseEditorDialog(fot_widget=self.fot_widget, parent=self)
        dlg.exec()

    def adjust_quantities_for_remainder(self, works, remainder, total_spirit):
        factor = remainder / total_spirit
        new_works = []
        for w in works:
            new_qty = int(w['quantity'] * factor)
            if new_qty < 1 and w['quantity'] > 0:
                new_qty = 1
            new_w = w.copy()
            new_w['quantity'] = new_qty
            new_works.append(new_w)

        new_total = 0
        for w in new_works:
            group = grouping.determine_group(w['work_name'])
            if group:
                norm = db_manager.get_spirit_norm(group)
                new_total += w['quantity'] * norm

        diff = remainder - new_total
        if diff > 0:
            max_norm = 0
            best_idx = -1
            for i, w in enumerate(new_works):
                group = grouping.determine_group(w['work_name'])
                if group:
                    norm = db_manager.get_spirit_norm(group)
                    if norm > max_norm:
                        max_norm = norm
                        best_idx = i
            if best_idx >= 0:
                add_units = int(diff / max_norm) + 1
                new_works[best_idx]['quantity'] += add_units
        return new_works

    def generate_act_from_db(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Формирование акта")
        dlg.resize(400, 280)
        layout = QVBoxLayout(dlg)

        radio_month = QRadioButton("По работам за месяц")
        radio_remainder = QRadioButton("По остатку спирта")
        radio_month.setChecked(True)
        layout.addWidget(radio_month)
        layout.addWidget(radio_remainder)

        month_group = QGroupBox("Параметры месяца")
        month_layout = QFormLayout(month_group)
        year_combo = QComboBox()
        current_year = QDate.currentDate().year()
        for y in range(2020, 2031):
            year_combo.addItem(str(y), y)
        year_combo.setCurrentText(str(current_year))
        month_layout.addRow("Год:", year_combo)
        month_combo = QComboBox()
        months = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
                  "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
        for i, m in enumerate(months, 1):
            month_combo.addItem(m, i)
        month_combo.setCurrentIndex(QDate.currentDate().month() - 1)
        month_layout.addRow("Месяц:", month_combo)
        layout.addWidget(month_group)

        remainder_group = QGroupBox("Параметры остатка")
        remainder_layout = QFormLayout(remainder_group)
        remainder_spin = QDoubleSpinBox()
        remainder_spin.setRange(0, 1000)
        remainder_spin.setValue(0.0)
        remainder_layout.addRow("Остаток (л):", remainder_spin)
        layout.addWidget(remainder_group)

        def on_mode_changed():
            is_month = radio_month.isChecked()
            month_group.setEnabled(is_month)
            remainder_group.setEnabled(not is_month)
        radio_month.toggled.connect(on_mode_changed)
        radio_remainder.toggled.connect(on_mode_changed)
        on_mode_changed()

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        year = year_combo.currentData()
        month = month_combo.currentData()
        start_str = f"01.{month:02d}.{year}"
        last_day = calendar.monthrange(year, month)[1]
        end_str = f"{last_day:02d}.{month:02d}.{year}"
        works = db_manager.get_works_by_period(start_str, end_str)

        if not works:
            QMessageBox.warning(self, "Предупреждение", f"Нет работ за {month:02d}.{year}")
            return

        if radio_remainder.isChecked():
            remainder = remainder_spin.value()
            if remainder <= 0:
                QMessageBox.warning(self, "Ошибка", "Остаток должен быть больше 0")
                return

            total_spirit = 0
            for w in works:
                group = grouping.determine_group(w['work_name'])
                if group:
                    norm = db_manager.get_spirit_norm(group)
                    total_spirit += w.get('quantity', 1) * norm

            if total_spirit > remainder:
                reply = QMessageBox.question(self, "Превышение остатка",
                                             f"Суммарный спирт за месяц ({total_spirit:.2f} л) превышает остаток ({remainder} л).\n"
                                             "Пропорционально уменьшить количество штук?",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if reply == QMessageBox.StandardButton.Yes:
                    works = self.adjust_quantities_for_remainder(works, remainder, total_spirit)
                else:
                    return

        grouped = grouping.group_works(works)
        output_path, _ = QFileDialog.getSaveFileName(self, "Сохранить акт", "", "Excel files (*.xlsx)")
        if output_path:
            act_generator.generate_act(grouped, start_str, end_str, output_path)
            os.startfile(output_path)

    def open_settings(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Настройки акта")
        dialog.resize(500, 500)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        settings = db_manager.get_all_settings()

        fields = {}
        org_edit = QLineEdit(settings.get("organization", ""))
        form.addRow("Организация:", org_edit)
        fields["organization"] = org_edit

        act_num_edit = QLineEdit(settings.get("act_number", ""))
        form.addRow("Номер акта:", act_num_edit)
        fields["act_number"] = act_num_edit

        act_date_edit = QDateEdit()
        act_date_edit.setDate(QDate.fromString(settings.get("act_date", QDate.currentDate().toString("dd.MM.yyyy")), "dd.MM.yyyy"))
        act_date_edit.setCalendarPopup(True)
        form.addRow("Дата акта:", act_date_edit)
        fields["act_date"] = act_date_edit

        chairman_edit = QLineEdit(settings.get("commission_chairman", ""))
        form.addRow("Председатель комиссии:", chairman_edit)
        fields["commission_chairman"] = chairman_edit

        members_edit = QTextEdit()
        members_edit.setPlainText(settings.get("commission_members", ""))
        members_edit.setMaximumHeight(100)
        form.addRow("Члены комиссии:", members_edit)
        fields["commission_members"] = members_edit

        responsible_edit = QLineEdit(settings.get("responsible_person", ""))
        form.addRow("Материально-ответственное лицо:", responsible_edit)
        fields["responsible_person"] = responsible_edit

        economist_edit = QLineEdit(settings.get("economist", ""))
        form.addRow("Экономист:", economist_edit)
        fields["economist"] = economist_edit

        ticket_edit = QLineEdit(settings.get("antiseptic_solution_ticket", ""))
        form.addRow("Реквизиты антисептика:", ticket_edit)
        fields["antiseptic_solution_ticket"] = ticket_edit

        inv_edit = QLineEdit(settings.get("antiseptic_inv_number", ""))
        form.addRow("Инв. № антисептика:", inv_edit)
        fields["antiseptic_inv_number"] = inv_edit

        amount_edit = QLineEdit(settings.get("antiseptic_amount", ""))
        form.addRow("Количество антисептика (л):", amount_edit)
        fields["antiseptic_amount"] = amount_edit

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            db_manager.set_setting("organization", org_edit.text())
            db_manager.set_setting("act_number", act_num_edit.text())
            db_manager.set_setting("act_date", act_date_edit.date().toString("dd.MM.yyyy"))
            db_manager.set_setting("commission_chairman", chairman_edit.text())
            db_manager.set_setting("commission_members", members_edit.toPlainText())
            db_manager.set_setting("responsible_person", responsible_edit.text())
            db_manager.set_setting("economist", economist_edit.text())
            db_manager.set_setting("antiseptic_solution_ticket", ticket_edit.text())
            db_manager.set_setting("antiseptic_inv_number", inv_edit.text())
            db_manager.set_setting("antiseptic_amount", amount_edit.text())
            QMessageBox.information(self, "Успех", "Настройки сохранены.")

    def populate_menu(self, menu_bar):
        pass

    def maybe_save(self):
        return True