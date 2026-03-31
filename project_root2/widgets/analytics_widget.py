import datetime
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QFileDialog, QDateEdit, QFormLayout, QDialog, QDialogButtonBox,
    QGroupBox, QGridLayout, QComboBox, QCheckBox
)
import pandas as pd
from database.repair_db import (
    get_all_orders, get_material_by_id, get_all_suppliers, get_agreements_by_supplier,
    get_all_materials
)

class AnalyticsWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.load_data()
        # Автоматическое обновление каждые 5 минут (опционально)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.load_data)
        self.timer.start(300000)  # 5 минут

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # Панель выбора периода
        period_layout = QHBoxLayout()
        period_layout.addWidget(QLabel("Начало:"))
        self.start_date = QDateEdit()
        self.start_date.setDate(datetime.date.today().replace(day=1))
        self.start_date.setCalendarPopup(True)
        period_layout.addWidget(self.start_date)

        period_layout.addWidget(QLabel("Конец:"))
        self.end_date = QDateEdit()
        self.end_date.setDate(datetime.date.today())
        self.end_date.setCalendarPopup(True)
        period_layout.addWidget(self.end_date)

        self.btn_update = QPushButton("Применить")
        self.btn_update.clicked.connect(self.load_data)
        period_layout.addWidget(self.btn_update)
        period_layout.addStretch()
        layout.addLayout(period_layout)

        # Вкладки
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        self.setup_dashboard_tab()
        self.setup_dynamics_tab()
        self.setup_structure_tab()
        self.setup_suppliers_tab()
        self.setup_customers_tab()
        self.setup_materials_tab()
        self.setup_orders_tab()
        self.setup_forecast_tab()

        # Кнопка экспорта всех данных
        export_btn = QPushButton("Экспорт всех данных в Excel")
        export_btn.clicked.connect(self.export_all)
        layout.addWidget(export_btn, alignment=Qt.AlignRight)

    # ---------- Вспомогательные методы ----------
    def get_orders_in_period(self, start, end):
        all_orders = get_all_orders()
        return [o for o in all_orders if start <= o.date <= end]

    def compute_order_metrics(self, orders):
        total_income = 0.0
        total_cost = 0.0
        total_count = len(orders)
        for order in orders:
            income = 0.0
            cost = 0.0
            for work in order.work_items:
                income += work.price
                for om in work.materials:
                    income += om.quantity * om.sale_price
                    for batch in om.batches:
                        cost += batch.quantity * batch.purchase_price
            total_income += income
            total_cost += cost
        profit = total_income - total_cost
        margin = (profit / total_income * 100) if total_income != 0 else 0
        avg_check = total_income / total_count if total_count else 0
        return {
            'income': total_income,
            'cost': total_cost,
            'profit': profit,
            'margin': margin,
            'count': total_count,
            'avg_check': avg_check
        }

    # ---------- Вкладка "Главная панель" ----------
    def setup_dashboard_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, "Главная")
        layout = QVBoxLayout(tab)

        # Группа текущих метрик
        current_group = QGroupBox("Текущий период")
        current_layout = QGridLayout(current_group)
        self.lbl_income = QLabel("Доход: 0.00")
        self.lbl_cost = QLabel("Себестоимость: 0.00")
        self.lbl_profit = QLabel("Прибыль: 0.00")
        self.lbl_margin = QLabel("Рентабельность: 0%")
        self.lbl_count = QLabel("Кол-во заказов: 0")
        self.lbl_avg_check = QLabel("Средний чек: 0.00")
        current_layout.addWidget(self.lbl_income, 0, 0)
        current_layout.addWidget(self.lbl_cost, 0, 1)
        current_layout.addWidget(self.lbl_profit, 0, 2)
        current_layout.addWidget(self.lbl_margin, 1, 0)
        current_layout.addWidget(self.lbl_count, 1, 1)
        current_layout.addWidget(self.lbl_avg_check, 1, 2)
        layout.addWidget(current_group)

        # Группа сравнения с предыдущим периодом
        prev_group = QGroupBox("Сравнение с предыдущим периодом")
        prev_layout = QGridLayout(prev_group)
        self.lbl_prev_income = QLabel("Доход: 0.00 (0%)")
        self.lbl_prev_cost = QLabel("Себестоимость: 0.00 (0%)")
        self.lbl_prev_profit = QLabel("Прибыль: 0.00 (0%)")
        self.lbl_prev_count = QLabel("Заказов: 0 (0%)")
        self.lbl_prev_avg = QLabel("Средний чек: 0.00 (0%)")
        prev_layout.addWidget(self.lbl_prev_income, 0, 0)
        prev_layout.addWidget(self.lbl_prev_cost, 0, 1)
        prev_layout.addWidget(self.lbl_prev_profit, 0, 2)
        prev_layout.addWidget(self.lbl_prev_count, 1, 0)
        prev_layout.addWidget(self.lbl_prev_avg, 1, 1)
        layout.addWidget(prev_group)

    # ---------- Вкладка "Динамика" ----------
    def setup_dynamics_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, "Динамика")
        layout = QVBoxLayout(tab)

        self.dynamics_fig = Figure(figsize=(11, 6))
        self.dynamics_canvas = FigureCanvas(self.dynamics_fig)
        layout.addWidget(self.dynamics_canvas)

    def update_dynamics(self, start, end, orders):
        monthly = {}
        for order in orders:
            key = order.date.strftime("%Y-%m")
            if key not in monthly:
                monthly[key] = {'income': 0.0, 'cost': 0.0, 'count': 0}
            inc, cost = 0.0, 0.0
            for work in order.work_items:
                inc += work.price
                for om in work.materials:
                    inc += om.quantity * om.sale_price
                    for batch in om.batches:
                        cost += batch.quantity * batch.purchase_price
            monthly[key]['income'] += inc
            monthly[key]['cost'] += cost
            monthly[key]['count'] += 1

        months = sorted(monthly.keys())
        incomes = [monthly[m]['income'] for m in months]
        costs = [monthly[m]['cost'] for m in months]
        counts = [monthly[m]['count'] for m in months]
        profits = [incomes[i] - costs[i] for i in range(len(months))]
        margins = [p / incomes[i] * 100 if incomes[i] != 0 else 0 for i, p in enumerate(profits)]
        avg_profit = [profits[i] / counts[i] if counts[i] != 0 else 0 for i in range(len(months))]

        self.dynamics_fig.clear()
        ax1 = self.dynamics_fig.add_subplot(311)
        x = np.arange(len(months))
        width = 0.35
        ax1.bar(x - width/2, incomes, width, label='Доход', color='#27ae60')
        ax1.bar(x + width/2, costs, width, label='Себестоимость', color='#e67e22')
        ax1.set_xticks(x)
        ax1.set_xticklabels(months, rotation=45)
        ax1.legend()
        ax1.set_title("Доходы и себестоимость")
        for i, (inc, cost) in enumerate(zip(incomes, costs)):
            ax1.text(x[i] - width/2, inc, f'{inc:,.0f}', ha='center', va='bottom', fontsize=8)
            ax1.text(x[i] + width/2, cost, f'{cost:,.0f}', ha='center', va='bottom', fontsize=8)

        ax2 = self.dynamics_fig.add_subplot(312)
        ax2.bar(x, counts, color='#3498db')
        ax2.set_xticks(x)
        ax2.set_xticklabels(months, rotation=45)
        ax2.set_title("Количество заказов")
        for i, cnt in enumerate(counts):
            ax2.text(x[i], cnt, f'{cnt}', ha='center', va='bottom', fontsize=8)

        ax3 = self.dynamics_fig.add_subplot(313)
        ax3.bar(x, avg_profit, color='#e67e22')
        ax3.set_xticks(x)
        ax3.set_xticklabels(months, rotation=45)
        ax3.set_title("Средняя прибыль на заказ")
        for i, ap in enumerate(avg_profit):
            ax3.text(x[i], ap, f'{ap:,.0f}', ha='center', va='bottom', fontsize=8)

        self.dynamics_fig.tight_layout()
        self.dynamics_canvas.draw()

    # ---------- Вкладка "Структура" ----------
    def setup_structure_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, "Структура")
        layout = QVBoxLayout(tab)

        self.structure_fig = Figure(figsize=(11, 5))
        self.structure_canvas = FigureCanvas(self.structure_fig)
        layout.addWidget(self.structure_canvas)

    def update_structure(self, orders):
        work_cost = 0.0
        mat_cost = 0.0
        for order in orders:
            for work in order.work_items:
                work_cost += work.price
                for om in work.materials:
                    for batch in om.batches:
                        mat_cost += batch.quantity * batch.purchase_price

        self.structure_fig.clear()
        ax = self.structure_fig.add_subplot(111)
        labels = ['Стоимость работ', 'Материалы']
        values = [work_cost, mat_cost]
        ax.pie(values, labels=labels, autopct='%1.1f%%', startangle=90)
        ax.axis('equal')
        ax.set_title("Структура себестоимости")
        self.structure_fig.tight_layout()
        self.structure_canvas.draw()

    # ---------- Вкладка "Поставщики" ----------
    def setup_suppliers_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, "Поставщики")
        layout = QVBoxLayout(tab)

        self.suppliers_fig = Figure(figsize=(11, 4))
        self.suppliers_canvas = FigureCanvas(self.suppliers_fig)
        layout.addWidget(self.suppliers_canvas)

        self.suppliers_table = QTableWidget()
        self.suppliers_table.setHorizontalHeaderLabels(["Поставщик", "Израсходовано", "Остаток по договорам"])
        self.suppliers_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.suppliers_table)

    def update_suppliers(self):
        suppliers = get_all_suppliers()
        data = []
        for sup in suppliers:
            agreements = get_agreements_by_supplier(sup.id)
            spent = sum(a.spent_amount for a in agreements)
            total = sum(a.total_amount for a in agreements)
            remaining = total - spent
            data.append((sup.name, spent, remaining))
        data.sort(key=lambda x: x[1], reverse=True)

        self.suppliers_fig.clear()
        ax = self.suppliers_fig.add_subplot(111)
        names = [d[0] for d in data[:10]]
        spent = [d[1] for d in data[:10]]
        ax.barh(names, spent, color='#e67e22')
        ax.set_title("ТОП-10 поставщиков по израсходованным средствам")
        self.suppliers_fig.tight_layout()
        self.suppliers_canvas.draw()

        self.suppliers_table.setRowCount(len(data))
        for i, (name, spent, remaining) in enumerate(data):
            self.suppliers_table.setItem(i, 0, QTableWidgetItem(name))
            self.suppliers_table.setItem(i, 1, QTableWidgetItem(f"{spent:.2f}"))
            self.suppliers_table.setItem(i, 2, QTableWidgetItem(f"{remaining:.2f}"))

    # ---------- Вкладка "Заказчики" ----------
    def setup_customers_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, "Заказчики")
        layout = QVBoxLayout(tab)

        self.customers_fig = Figure(figsize=(11, 4))
        self.customers_canvas = FigureCanvas(self.customers_fig)
        layout.addWidget(self.customers_canvas)

        self.customers_table = QTableWidget()
        self.customers_table.setHorizontalHeaderLabels(["Заказчик", "Выручка", "Средний чек", "Кол-во заказов"])
        self.customers_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.customers_table)

    def update_customers(self, orders):
        customer_data = {}
        for order in orders:
            name = order.customer_name
            income = 0.0
            for work in order.work_items:
                income += work.price
                for om in work.materials:
                    income += om.quantity * om.sale_price
            if name not in customer_data:
                customer_data[name] = {'income': 0, 'count': 0}
            customer_data[name]['income'] += income
            customer_data[name]['count'] += 1

        data = [(name, d['income'], d['income'] / d['count'], d['count']) for name, d in customer_data.items()]
        data.sort(key=lambda x: x[1], reverse=True)

        self.customers_fig.clear()
        ax = self.customers_fig.add_subplot(111)
        names = [d[0] for d in data[:10]]
        incomes = [d[1] for d in data[:10]]
        ax.barh(names, incomes, color='#27ae60')
        ax.set_title("ТОП-10 заказчиков по выручке")
        self.customers_fig.tight_layout()
        self.customers_canvas.draw()

        self.customers_table.setRowCount(len(data))
        for i, (name, inc, avg, cnt) in enumerate(data):
            self.customers_table.setItem(i, 0, QTableWidgetItem(name))
            self.customers_table.setItem(i, 1, QTableWidgetItem(f"{inc:.2f}"))
            self.customers_table.setItem(i, 2, QTableWidgetItem(f"{avg:.2f}"))
            self.customers_table.setItem(i, 3, QTableWidgetItem(str(cnt)))

    # ---------- Вкладка "Материалы" ----------
    def setup_materials_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, "Материалы")
        layout = QVBoxLayout(tab)

        self.materials_fig = Figure(figsize=(11, 4))
        self.materials_canvas = FigureCanvas(self.materials_fig)
        layout.addWidget(self.materials_canvas)

        self.materials_table = QTableWidget()
        self.materials_table.setHorizontalHeaderLabels(["Материал", "Тип", "Продано", "Выручка", "Прибыль"])
        self.materials_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.materials_table)

    def update_materials(self, orders):
        mat_stats = {}
        for order in orders:
            for work in order.work_items:
                for om in work.materials:
                    mat = get_material_by_id(om.material_id)
                    if not mat:
                        continue
                    key = (mat.name, mat.type)
                    if key not in mat_stats:
                        mat_stats[key] = {'qty': 0, 'revenue': 0, 'cost': 0}
                    for batch in om.batches:
                        mat_stats[key]['qty'] += batch.quantity
                        revenue = batch.quantity * om.sale_price
                        mat_stats[key]['revenue'] += revenue
                        mat_stats[key]['cost'] += batch.quantity * batch.purchase_price

        data = [(name, typ, stats['qty'], stats['revenue'], stats['revenue'] - stats['cost'])
                for (name, typ), stats in mat_stats.items()]
        data.sort(key=lambda x: x[4], reverse=True)

        # График ТОП-10 по прибыли
        self.materials_fig.clear()
        ax = self.materials_fig.add_subplot(111)
        top10 = data[:10]
        names = [d[0] for d in top10]
        profits = [d[4] for d in top10]
        ax.barh(names, profits, color='#9b59b6')
        ax.set_title("ТОП-10 материалов по прибыли")
        self.materials_fig.tight_layout()
        self.materials_canvas.draw()

        self.materials_table.setRowCount(len(data))
        for i, (name, typ, qty, rev, profit) in enumerate(data):
            self.materials_table.setItem(i, 0, QTableWidgetItem(name))
            self.materials_table.setItem(i, 1, QTableWidgetItem(typ))
            self.materials_table.setItem(i, 2, QTableWidgetItem(f"{qty:.2f}"))
            self.materials_table.setItem(i, 3, QTableWidgetItem(f"{rev:.2f}"))
            self.materials_table.setItem(i, 4, QTableWidgetItem(f"{profit:.2f}"))

    # ---------- Вкладка "Заказы" ----------
    def setup_orders_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, "Заказы")
        layout = QVBoxLayout(tab)

        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Статус:"))
        self.order_status_combo = QComboBox()
        self.order_status_combo.addItems(["Все", "в работе", "выполнен", "оплачен"])
        filter_layout.addWidget(self.order_status_combo)

        filter_layout.addWidget(QLabel("Заказчик:"))
        self.order_customer_combo = QComboBox()
        filter_layout.addWidget(self.order_customer_combo)

        self.order_filter_btn = QPushButton("Применить")
        self.order_filter_btn.clicked.connect(self.load_orders_table)
        filter_layout.addWidget(self.order_filter_btn)
        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        self.orders_table = QTableWidget()
        self.orders_table.setHorizontalHeaderLabels(["Номер", "Дата", "Заказчик", "Статус", "Доход", "Себестоимость", "Прибыль"])
        self.orders_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.orders_table)

        export_btn = QPushButton("Экспорт таблицы в Excel")
        export_btn.clicked.connect(self.export_orders_table)
        layout.addWidget(export_btn, alignment=Qt.AlignRight)

    def load_orders_table(self):
        start = self.start_date.date().toPython()
        end = self.end_date.date().toPython()
        orders = self.get_orders_in_period(start, end)

        status = self.order_status_combo.currentText()
        customer = self.order_customer_combo.currentText()

        if status != "Все":
            orders = [o for o in orders if o.status == status]
        if customer and customer != "":
            orders = [o for o in orders if o.customer_name == customer]

        self.orders_table.setRowCount(len(orders))
        for i, order in enumerate(orders):
            income = 0.0
            cost = 0.0
            for work in order.work_items:
                income += work.price
                for om in work.materials:
                    income += om.quantity * om.sale_price
                    for batch in om.batches:
                        cost += batch.quantity * batch.purchase_price
            profit = income - cost
            self.orders_table.setItem(i, 0, QTableWidgetItem(order.order_number))
            self.orders_table.setItem(i, 1, QTableWidgetItem(order.date.strftime("%d.%m.%Y")))
            self.orders_table.setItem(i, 2, QTableWidgetItem(order.customer_name))
            self.orders_table.setItem(i, 3, QTableWidgetItem(order.status))
            self.orders_table.setItem(i, 4, QTableWidgetItem(f"{income:.2f}"))
            self.orders_table.setItem(i, 5, QTableWidgetItem(f"{cost:.2f}"))
            self.orders_table.setItem(i, 6, QTableWidgetItem(f"{profit:.2f}"))

        # Заполним комбобокс заказчиков
        customers = sorted(set(o.customer_name for o in orders))
        self.order_customer_combo.clear()
        self.order_customer_combo.addItem("")
        for c in customers:
            self.order_customer_combo.addItem(c)

    def export_orders_table(self):
        fname, _ = QFileDialog.getSaveFileName(self, "Сохранить заказы", "Заказы.xlsx", "Excel Files (*.xlsx)")
        if not fname:
            return
        data = []
        for row in range(self.orders_table.rowCount()):
            row_data = []
            for col in range(self.orders_table.columnCount()):
                item = self.orders_table.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)
        headers = [self.orders_table.horizontalHeaderItem(i).text() for i in range(self.orders_table.columnCount())]
        df = pd.DataFrame(data, columns=headers)
        df.to_excel(fname, index=False)
        QMessageBox.information(self, "Экспорт", "Таблица сохранена.")

    # ---------- Вкладка "Прогноз" ----------
    def setup_forecast_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, "Прогноз")
        layout = QVBoxLayout(tab)

        self.forecast_fig = Figure(figsize=(11, 5))
        self.forecast_canvas = FigureCanvas(self.forecast_fig)
        layout.addWidget(self.forecast_canvas)

    def update_forecast(self, orders):
        # Собираем доходы за последние 12 месяцев
        today = datetime.date.today()
        start = today - datetime.timedelta(days=365)
        monthly = {}
        for order in orders:
            if order.date >= start:
                key = order.date.strftime("%Y-%m")
                if key not in monthly:
                    monthly[key] = 0.0
                inc = 0.0
                for work in order.work_items:
                    inc += work.price
                    for om in work.materials:
                        inc += om.quantity * om.sale_price
                monthly[key] += inc
        months = sorted(monthly.keys())
        if len(months) < 3:
            # Недостаточно данных
            self.forecast_fig.clear()
            ax = self.forecast_fig.add_subplot(111)
            ax.text(0.5, 0.5, "Недостаточно данных для прогноза", ha='center', va='center')
            self.forecast_canvas.draw()
            return

        x = np.arange(len(months))
        y = np.array([monthly[m] for m in months])
        coeff = np.polyfit(x, y, 1)
        trend = np.poly1d(coeff)
        future_x = np.arange(len(months), len(months) + 6)
        future_y = trend(future_x)

        last_month = datetime.date.fromisoformat(months[-1] + "-01")
        future_labels = []
        for i in range(6):
            next_month = last_month.replace(day=1) + datetime.timedelta(days=32)
            next_month = next_month.replace(day=1)
            future_labels.append(next_month.strftime("%Y-%m"))
            last_month = next_month

        self.forecast_fig.clear()
        ax = self.forecast_fig.add_subplot(111)
        ax.plot(x, y, 'o-', label='Исторические данные')
        ax.plot(future_x, future_y, 'o--', label='Прогноз')
        ax.axvline(x=len(months)-0.5, color='gray', linestyle='--', alpha=0.5)
        ax.set_xticks(list(x) + list(future_x))
        ax.set_xticklabels(months + future_labels, rotation=45)
        ax.legend()
        ax.set_title("Прогноз доходов на 6 месяцев")
        for i, (xi, yi) in enumerate(zip(x, y)):
            ax.annotate(f'{yi:,.0f}', (xi, yi), textcoords="offset points", xytext=(0,10), ha='center', fontsize=8)
        for i, (xi, yi) in enumerate(zip(future_x, future_y)):
            ax.annotate(f'{yi:,.0f}', (xi, yi), textcoords="offset points", xytext=(0,10), ha='center', fontsize=8)
        self.forecast_fig.tight_layout()
        self.forecast_canvas.draw()

    # ---------- Основной метод загрузки данных ----------
    def load_data(self):
        start = self.start_date.date().toPython()
        end = self.end_date.date().toPython()
        orders = self.get_orders_in_period(start, end)

        # Расчёт метрик для текущего и предыдущего периода
        current = self.compute_order_metrics(orders)

        # Предыдущий период (такой же длины)
        duration = (end - start).days
        prev_start = start - datetime.timedelta(days=duration)
        prev_end = start - datetime.timedelta(days=1)
        prev_orders = self.get_orders_in_period(prev_start, prev_end)
        prev = self.compute_order_metrics(prev_orders)

        # Обновление главной панели
        self.lbl_income.setText(f"Доход: {current['income']:.2f}")
        self.lbl_cost.setText(f"Себестоимость: {current['cost']:.2f}")
        self.lbl_profit.setText(f"Прибыль: {current['profit']:.2f}")
        self.lbl_margin.setText(f"Рентабельность: {current['margin']:.1f}%")
        self.lbl_count.setText(f"Кол-во заказов: {current['count']}")
        self.lbl_avg_check.setText(f"Средний чек: {current['avg_check']:.2f}")

        # Сравнение с предыдущим периодом
        def change(val, prev_val):
            if prev_val == 0:
                return "∞" if val > 0 else "0%"
            return f"{((val - prev_val) / prev_val * 100):+.1f}%"
        self.lbl_prev_income.setText(f"Доход: {current['income']:.2f} ({change(current['income'], prev['income'])})")
        self.lbl_prev_cost.setText(f"Себестоимость: {current['cost']:.2f} ({change(current['cost'], prev['cost'])})")
        self.lbl_prev_profit.setText(f"Прибыль: {current['profit']:.2f} ({change(current['profit'], prev['profit'])})")
        self.lbl_prev_count.setText(f"Заказов: {current['count']} ({change(current['count'], prev['count'])})")
        self.lbl_prev_avg.setText(f"Средний чек: {current['avg_check']:.2f} ({change(current['avg_check'], prev['avg_check'])})")

        # Обновление других вкладок
        self.update_dynamics(start, end, orders)
        self.update_structure(orders)
        self.update_suppliers()
        self.update_customers(orders)
        self.update_materials(orders)
        self.load_orders_table()
        self.update_forecast(orders)

    # ---------- Экспорт всех данных ----------
    def export_all(self):
        fname, _ = QFileDialog.getSaveFileName(self, "Сохранить отчёт", "Аналитика.xlsx", "Excel Files (*.xlsx)")
        if not fname:
            return
        with pd.ExcelWriter(fname, engine='openpyxl') as writer:
            # Главные метрики
            metrics = [
                ["Показатель", "Значение"],
                ["Доход", self.lbl_income.text().split(": ")[1]],
                ["Себестоимость", self.lbl_cost.text().split(": ")[1]],
                ["Прибыль", self.lbl_profit.text().split(": ")[1]],
                ["Рентабельность", self.lbl_margin.text().split(": ")[1]],
                ["Количество заказов", self.lbl_count.text().split(": ")[1]],
                ["Средний чек", self.lbl_avg_check.text().split(": ")[1]],
            ]
            df = pd.DataFrame(metrics[1:], columns=metrics[0])
            df.to_excel(writer, sheet_name="Сводка", index=False)

            # Таблица заказов
            orders_data = []
            for row in range(self.orders_table.rowCount()):
                row_data = []
                for col in range(self.orders_table.columnCount()):
                    item = self.orders_table.item(row, col)
                    row_data.append(item.text() if item else "")
                orders_data.append(row_data)
            headers = [self.orders_table.horizontalHeaderItem(i).text() for i in range(self.orders_table.columnCount())]
            df_orders = pd.DataFrame(orders_data, columns=headers)
            df_orders.to_excel(writer, sheet_name="Заказы", index=False)

            # ТОП материалов
            mat_data = []
            for row in range(self.materials_table.rowCount()):
                row_data = []
                for col in range(self.materials_table.columnCount()):
                    item = self.materials_table.item(row, col)
                    row_data.append(item.text() if item else "")
                mat_data.append(row_data)
            headers_mat = [self.materials_table.horizontalHeaderItem(i).text() for i in range(self.materials_table.columnCount())]
            df_mat = pd.DataFrame(mat_data, columns=headers_mat)
            df_mat.to_excel(writer, sheet_name="Материалы", index=False)

            # ТОП заказчиков
            cust_data = []
            for row in range(self.customers_table.rowCount()):
                row_data = []
                for col in range(self.customers_table.columnCount()):
                    item = self.customers_table.item(row, col)
                    row_data.append(item.text() if item else "")
                cust_data.append(row_data)
            headers_cust = [self.customers_table.horizontalHeaderItem(i).text() for i in range(self.customers_table.columnCount())]
            df_cust = pd.DataFrame(cust_data, columns=headers_cust)
            df_cust.to_excel(writer, sheet_name="Заказчики", index=False)

            # ТОП поставщиков
            sup_data = []
            for row in range(self.suppliers_table.rowCount()):
                row_data = []
                for col in range(self.suppliers_table.columnCount()):
                    item = self.suppliers_table.item(row, col)
                    row_data.append(item.text() if item else "")
                sup_data.append(row_data)
            headers_sup = [self.suppliers_table.horizontalHeaderItem(i).text() for i in range(self.suppliers_table.columnCount())]
            df_sup = pd.DataFrame(sup_data, columns=headers_sup)
            df_sup.to_excel(writer, sheet_name="Поставщики", index=False)

        QMessageBox.information(self, "Экспорт", f"Отчёт сохранён в {fname}")

    def populate_menu(self, menu_bar):
        anal_menu = menu_bar.addMenu("Аналитика")
        dummy = anal_menu.addAction("(Не жмякай куда не просят!!!)")
        dummy.setEnabled(False)

    def maybe_save(self):
        return True