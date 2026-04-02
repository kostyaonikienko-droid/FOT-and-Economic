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
        self.tree.setColumnWidth(0, 80)
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

        # self.btn_settings = QPushButton("Настройки")
        # self.btn_settings.clicked.connect(self.open_settings)
        # right_panel.addWidget(self.btn_settings)

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
        if column == 6:
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

    def trim_works_to_remainder(self, works, remainder):
        """
        Жёстко обрезает список работ так, чтобы суммарный литраж стал <= remainder.
        Сначала уменьшает количество у самых дорогих работ (по норме),
        если количество достигает 1, удаляет работу целиком.
        """
        # Добавляем норму и литры
        for w in works:
            group = grouping.determine_group(w['work_name'])
            norm = db_manager.get_spirit_norm(group) if group else 0
            w['norm'] = norm
            w['liters'] = w.get('quantity', 1) * norm

        remaining = [w.copy() for w in works]
        total = sum(w['liters'] for w in remaining)

        if total <= remainder:
            # Уже в лимите
            for w in remaining:
                del w['norm'], w['liters']
            return remaining

        # Сортируем по убыванию нормы
        remaining.sort(key=lambda x: x['norm'], reverse=True)

        while total > remainder and remaining:
            # Берём самую дорогую
            w = remaining[0]
            if w['quantity'] > 1:
                # Уменьшаем количество на 1
                w['quantity'] -= 1
                old_liters = w['liters']
                w['liters'] = w['quantity'] * w['norm']
                total -= (old_liters - w['liters'])
            else:
                # Количество = 1, удаляем работу
                total -= w['liters']
                remaining.pop(0)
            # После изменения (уменьшения или удаления) пересортировываем,
            # чтобы самая дорогая работа снова была в начале
            remaining.sort(key=lambda x: x['norm'], reverse=True)

        # Удаляем временные поля
        for w in remaining:
            del w['norm'], w['liters']
        return remaining


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
        dlg.resize(450, 400)
        layout = QVBoxLayout(dlg)

        radio_month = QRadioButton("По работам за месяц")
        radio_remainder = QRadioButton("По остатку спирта")
        radio_month.setChecked(True)
        layout.addWidget(radio_month)
        layout.addWidget(radio_remainder)

        # --- Блок параметров месяца (активен всегда) ---
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

        # --- Блок остатка ---
        remainder_group = QGroupBox("Остаток спирта (л)")
        remainder_layout = QFormLayout(remainder_group)
        remainder_spin = QDoubleSpinBox()
        remainder_spin.setRange(0, 1000)
        remainder_spin.setValue(0.0)
        remainder_layout.addRow("Остаток:", remainder_spin)
        layout.addWidget(remainder_group)

        # --- Блок формата ---
        format_group = QGroupBox("Формат акта")
        format_layout = QHBoxLayout(format_group)
        radio_word = QRadioButton("Word (для служебной)")
        radio_excel = QRadioButton("Excel (для архива)")
        radio_word.setChecked(True)
        format_layout.addWidget(radio_word)
        format_layout.addWidget(radio_excel)
        layout.addWidget(format_group)

        def on_mode_changed():
            is_month = radio_month.isChecked()
            month_group.setEnabled(True)  # месяц всегда доступен
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

        # Режим остатка: жёсткая обрезка
        if radio_remainder.isChecked():
            remainder = remainder_spin.value()
            if remainder <= 0:
                QMessageBox.warning(self, "Ошибка", "Остаток должен быть больше 0")
                return

            works = self.trim_works_to_remainder(works, remainder)

        grouped = grouping.group_works(works)
        format_type = 'word' if radio_word.isChecked() else 'excel'
        filter_str = "Word документ (*.docx)" if format_type == 'word' else "Excel файлы (*.xlsx)"
        output_path, _ = QFileDialog.getSaveFileName(self, "Сохранить акт", "", filter_str)
        if output_path:
            # Добавляем расширение, если его нет
            if format_type == 'word' and not output_path.endswith('.docx'):
                output_path += '.docx'
            elif format_type == 'excel' and not output_path.endswith('.xlsx'):
                output_path += '.xlsx'

            try:
                act_generator.generate_act(grouped, start_str, end_str, output_path, fot_widget=self.fot_widget,
                                           format=format_type)
                os.startfile(output_path)
            except PermissionError:
                QMessageBox.warning(self, "Ошибка",
                                    f"Не удалось сохранить файл: {output_path}\n"
                                    "Возможно, файл уже открыт в другой программе. Закройте его и повторите попытку.")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сформировать акт:\n{str(e)}")

    def open_settings(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Настройки")
        dialog.resize(500, 200)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        settings = db_manager.get_all_settings()

        ticket_edit = QLineEdit(settings.get("antiseptic_solution_ticket", ""))
        form.addRow("Реквизиты антисептика (тр., дата, кол-во):", ticket_edit)
        inv_edit = QLineEdit(settings.get("antiseptic_inv_number", ""))
        form.addRow("Инв. № антисептика:", inv_edit)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            db_manager.set_setting("antiseptic_solution_ticket", ticket_edit.text())
            db_manager.set_setting("antiseptic_inv_number", inv_edit.text())
            QMessageBox.information(self, "Успех", "Настройки сохранены.")


    def populate_menu(self, menu_bar):
        pass

    def show_instruction(self):
        instruction_text = (
            "ИНСТРУКЦИЯ ПО РАБОТЕ С МОДУЛЕМ «АКТ НА СПИРТ»\n\n"
            "1. Загрузите Excel-файл с данными (кнопка «Загрузить Excel»).\n"
            "   Программа автоматически отфильтрует строки с ошибочными датами (1900-01-01) и нулевыми счетами/количеством.\n\n"
            "2. При необходимости отфильтруйте записи по ключевым словам.\n"
            "   Например, для вискозиметров ВЗ‑246 и ВЗ‑4 введите «ВЗ-246, ВЗ-4» (через запятую).\n\n"
            "3. В дереве слева вы можете:\n"
            "   - Отметить/снять галочки у нужных записей.\n"
            "   - Дважды кликнуть по количеству, чтобы изменить его вручную.\n"
            "   - Использовать контекстное меню (правой кнопкой мыши) для редактирования.\n\n"
            "4. Сохраните выбранные записи в базу данных (кнопка «Сохранить выбранное в базу»).\n\n"
            "5. Для формирования акта нажмите «Сформировать акт из базы».\n"
            "   - Выберите период (месяц/год) или режим «По остатку спирта».\n"
            "   - В режиме остатка укажите желаемое количество литров. Программа автоматически\n"
            "     обрежет список записей, чтобы итоговый литраж не превышал остаток.\n"
            "   - Выберите формат: Word (для служебной) или Excel (для архива).\n"
            "   - Укажите имя файла для сохранения.\n\n"
            "6. При необходимости отредактируйте базу данных (кнопка «Редактировать базу данных»).\n"
            "   Там можно изменить количество, удалить записи, отфильтровать по году, месяцу, исполнителю.\n\n"
            "7. Настройки акта (кнопка «Настройки акта») позволяют задать:\n"
            "   - Реквизиты антисептического раствора (тр., инв. №, количество).\n"
            "   - Остальные параметры (организация, комиссия и т.д.) – они сейчас не используются\n"
            "     в итоговом документе, но могут быть полезны для будущих версий.\n\n"
            "8. Сгенерированный акт (Word) и журнал (Excel) соответствуют утверждённой форме.\n"
            "   В Excel-журнале добавлены столбцы «ФИО» и «Общее кол-во литров по счету».\n\n"
            "При возникновении ошибок убедитесь, что файл сохранения не открыт в другой программе."
        )
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QTextEdit, QPushButton
        from PySide6.QtCore import Qt
        dlg = QDialog(self)
        dlg.setWindowTitle("Инструкция по модулю «Акт на спирт»")
        dlg.resize(700, 500)
        layout = QVBoxLayout(dlg)
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText(instruction_text)
        layout.addWidget(text_edit)
        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(dlg.accept)
        layout.addWidget(btn_close, alignment=Qt.AlignmentFlag.AlignCenter)
        dlg.exec()

    def maybe_save(self):
        return True

