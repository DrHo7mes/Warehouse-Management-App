import sys
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QPushButton, QLabel, QStackedWidget, QFrame, QButtonGroup
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

from database import DatabaseManager
from calculations import get_today_jalali_str
from ui_pages import create_all_pages

FONT_TEXT = "B Nazanin"
FONT_TITLE = "B Titr"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = DatabaseManager()
        self.update_window_title()
        self.resize(1260, 780)
        self.setMinimumSize(1024, 650)
        self.setLayoutDirection(Qt.RightToLeft)

        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: #f4f6f9;
            }}
            QWidget {{
                font-family: '{FONT_TEXT}', 'Tahoma';
                font-size: 14px;
            }}
        """)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self.sidebar = self.build_sidebar()
        main_layout.addWidget(self.sidebar)

        content_wrapper = QWidget()
        content_layout = QVBoxLayout(content_wrapper)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        self.header = self.build_header()
        content_layout.addWidget(self.header)

        self.stack = QStackedWidget()
        content_layout.addWidget(self.stack)

        main_layout.addWidget(content_wrapper, 1)

        self.pages = create_all_pages(self.db, on_settings_saved_callback=self.update_window_title)
        self.menu_keys = [
            "dashboard", "items", "categories", "transactions",
            "usage_report", "depletion", "reports", "backup", "settings"
        ]
        for key in self.menu_keys:
            self.stack.addWidget(self.pages[key])

        self.nav_buttons[0].setChecked(True)
        self.switch_page(0)

    def update_window_title(self):
        warehouse_name = self.db.get_setting("warehouse_name", "انبار مرکزی")
        self.setWindowTitle(f"سامانه جامع مدیریت انبار و موجودی - {warehouse_name}")

    def build_sidebar(self):
        sidebar = QFrame()
        sidebar.setFixedWidth(260)
        sidebar.setStyleSheet(f"""
            QFrame {{
                background-color: #1e272e;
                border-left: 1px solid #0fb9b1;
            }}
            QLabel#AppTitle {{
                color: #ffffff;
                font-family: '{FONT_TITLE}', 'Tahoma';
                font-size: 16px;
                font-weight: bold;
                padding: 18px 10px;
                background-color: #0fb9b1;
                border-bottom: 2px solid #0a8d87;
            }}
            QPushButton.NavBtn {{
                color: #d2dae2;
                background-color: transparent;
                border: none;
                border-radius: 8px;
                padding: 10px 16px;
                text-align: right;
                font-family: '{FONT_TEXT}', 'Tahoma';
                font-size: 16px;
                font-weight: bold;
                margin: 3px 10px;
            }}
            QPushButton.NavBtn:hover {{
                background-color: #2f3640;
                color: #ffffff;
            }}
            QPushButton.NavBtn:checked {{
                background-color: #0fb9b1;
                color: #ffffff;
            }}
        """)

        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(0, 0, 0, 15)
        layout.setSpacing(4)

        lbl_logo = QLabel("📦 سامانه هوشمند انبارداری")
        lbl_logo.setObjectName("AppTitle")
        lbl_logo.setAlignment(Qt.AlignCenter)
        layout.addWidget(lbl_logo)

        layout.addSpacing(10)

        menu_items = [
            ("📊  داشبورد مدیریتی", "dashboard"),
            ("📦  مدیریت کالاها", "items"),
            ("🗂️  دسته‌بندی‌ها", "categories"),
            ("🔄  ثبت ورود و خروج", "transactions"),
            ("📉  گزارش مصرف", "usage_report"),
            ("⏳  پیش‌بینی اتمام موجودی", "depletion"),
            ("📑  گزارش‌ها (Excel/PDF)", "reports"),
            ("💾  پشتیبان‌گیری و بازیابی", "backup"),
            ("⚙️  تنظیمات", "settings"),
        ]

        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        self.nav_buttons = []

        for idx, (title, key) in enumerate(menu_items):
            btn = QPushButton(title)
            btn.setProperty("class", "NavBtn")
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            self.nav_group.addButton(btn, idx)
            self.nav_buttons.append(btn)
            layout.addWidget(btn)

        self.nav_group.idClicked.connect(self.switch_page)

        layout.addStretch()

        lbl_footer = QLabel("نسخه ۱.۰ - ذخیره‌سازی محلی امن")
        lbl_footer.setAlignment(Qt.AlignCenter)
        lbl_footer.setStyleSheet(f"color: #808e9b; font-family: '{FONT_TEXT}'; font-size: 13px; padding: 6px;")
        layout.addWidget(lbl_footer)

        return sidebar

    def build_header(self):
        header = QFrame()
        header.setFixedHeight(64)
        header.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border-bottom: 1px solid #dfe4ea;
            }
        """)
        layout = QHBoxLayout(header)
        layout.setContentsMargins(25, 0, 25, 0)

        self.lbl_page_title = QLabel("داشبورد مدیریتی")
        self.lbl_page_title.setFont(QFont(FONT_TITLE, 15, QFont.Bold))
        self.lbl_page_title.setStyleSheet(f"color: #2f3542; font-family: '{FONT_TITLE}'; border: none;")

        self.lbl_date = QLabel(f"📅 امروز: {get_today_jalali_str()}")
        self.lbl_date.setFont(QFont(FONT_TEXT, 13, QFont.Bold))
        self.lbl_date.setStyleSheet(f"""
            color: #0fb9b1;
            background-color: #e8f8f5;
            font-family: '{FONT_TEXT}';
            font-size: 15px;
            font-weight: bold;
            padding: 6px 14px;
            border-radius: 6px;
            border: 1px solid #a3e4d7;
        """)

        layout.addWidget(self.lbl_page_title)
        layout.addStretch()
        layout.addWidget(self.lbl_date)

        return header

    def switch_page(self, index):
        if 0 <= index < len(self.menu_keys):
            btn_text = self.nav_buttons[index].text()
            self.lbl_page_title.setText(btn_text)
            self.stack.setCurrentIndex(index)
            current_widget = self.stack.currentWidget()
            if hasattr(current_widget, "refresh_data"):
                current_widget.refresh_data()

    def closeEvent(self, event):
        if self.db.get_setting("auto_backup", "1") == "1":
            self.db.create_backup()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.RightToLeft)
    app.setFont(QFont(FONT_TEXT, 12))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())