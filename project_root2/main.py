import sys
import os
from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget, QMessageBox
from PySide6.QtGui import QIcon
from PySide6.QtCore import Qt
from widgets.repair_widget import RepairWidget
from widgets.suppliers_widget import SuppliersWidget
from fot_module.fot_main import FOTWidget
from gso_module.gso_widget import GSOWidget
from analytics_module.analytics_widget import AnalyticsWidget
from database.repair_db import init_db  # импортируем init_db

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
from spirit_module.act_widget import ActWidget
from widgets.directories_widget import DirectoriesWidget
from widgets.repair_widget import RepairWidget
from fot_module.fot_main import FOTWidget
from widgets.analytics_widget import AnalyticsWidget


class MainAppWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        init_db()
        self.setWindowTitle("Экономика и ФОТ")
        self.setMinimumSize(1300, 700)

        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)

        self.fot_widget = FOTWidget(self)
        self.repair_widget = RepairWidget(self)
        self.directories_widget = DirectoriesWidget(self)
        self.analytics_widget = AnalyticsWidget(self)

        self.tab_widget.addTab(self.fot_widget, "ФОТ")
        self.tab_widget.addTab(self.repair_widget, "Учёт материалов")
        self.tab_widget.addTab(self.directories_widget, "Справочники")
          # без fot_widget
        self.tab_widget.addTab(self.analytics_widget, "Аналитика")
        self.spirit_widget = ActWidget(self, fot_widget=self.fot_widget)
        self.tab_widget.addTab(self.spirit_widget, "Акт на спирт")
        self.menu_bar = self.menuBar()
        self.tab_widget.currentChanged.connect(self.update_menu)
        self.update_menu(0)

    def update_menu(self, index):
        self.menu_bar.clear()
        if index == 0:
            self.fot_widget.populate_menu(self.menu_bar)
        elif index == 1:
            self.repair_widget.populate_menu(self.menu_bar)
        elif index == 2:
            # Заглушка для вкладки Справочники
            dir_menu = self.menu_bar.addMenu("Справочники")
            dummy = dir_menu.addAction("(Не жмякай куда не просят!!!)")
            dummy.setEnabled(False)
        elif index == 3:
            # Заглушка для вкладки Аналитика
            anal_menu = self.menu_bar.addMenu("Аналитика")
            dummy = anal_menu.addAction("(Не жмякай куда не просят!!!)")
            dummy.setEnabled(False)

        elif index == 4:
            # Для вкладки "Акт на спирт" можно добавить заглушку меню или свои пункты
            spirit_menu = self.menu_bar.addMenu("Акт на спирт")
            dummy = spirit_menu.addAction("(Не жмякай куда не просят!!!)")
            dummy.setEnabled(False)
            
    def closeEvent(self, event):
        if (self.fot_widget.maybe_save() and
                self.repair_widget.maybe_save() and
                self.directories_widget.maybe_save() and
                self.spirit_widget.maybe_save() and
                self.analytics_widget.maybe_save()):
            event.accept()
        else:
            event.ignore()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName("Экономика и ФОТ")

    icon_path = resource_path("icon.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    window = MainAppWindow()
    if os.path.exists(icon_path):
        window.setWindowIcon(QIcon(icon_path))

    window.show()
    sys.exit(app.exec())