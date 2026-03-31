# widgets/directories_widget.py

from PySide6.QtWidgets import QWidget, QVBoxLayout, QTabWidget
from .suppliers_widget import SuppliersWidget
from .customers_widget import CustomersWidget
from .materials_widget import MaterialsWidget  # переименуем старый MaterialsDialog в виджет

class DirectoriesWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        self.suppliers_widget = SuppliersWidget(self)
        self.customers_widget = CustomersWidget(self)
        self.materials_widget = MaterialsWidget(self)
        self.tabs.addTab(self.suppliers_widget, "Поставщики")
        self.tabs.addTab(self.customers_widget, "Заказчики")
        self.tabs.addTab(self.materials_widget, "Материалы")
        layout.addWidget(self.tabs)

        # При переключении подвкладки обновляем данные
        self.tabs.currentChanged.connect(self.on_tab_changed)

    def on_tab_changed(self, index):
        if index == 0:
            self.suppliers_widget.load_suppliers()
        elif index == 1:
            self.customers_widget.load_customers()
        elif index == 2:
            self.materials_widget.load_materials()

    def maybe_save(self):
        return True

