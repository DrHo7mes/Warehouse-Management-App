import os
from datetime import datetime
import jdatetime

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame, QPushButton,
    QLineEdit, QComboBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QDialog, QFormLayout, QDoubleSpinBox, QSpinBox, QTextEdit, QMessageBox,
    QAbstractItemView, QInputDialog, QFileDialog, QCheckBox
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QColor, QBrush

from calculations import (
    enrich_item_with_calculations, iso_to_jalali, jalali_to_gregorian_str
)
from reports import export_to_excel, export_to_pdf

FONT_TEXT = "B Nazanin"
FONT_TITLE = "B Titr"

TABLE_STYLE = f"""
    QTableWidget {{
        background-color: #ffffff;
        alternate-background-color: #f8f9fa;
        border: 1px solid #dfe4ea;
        border-radius: 8px;
        font-family: '{FONT_TEXT}', 'Tahoma';
        font-size: 15px;
        gridline-color: #ecf0f1;
    }}
    QHeaderView::section {{
        background-color: #2f3542;
        color: #ffffff;
        font-family: '{FONT_TITLE}', 'Tahoma';
        font-weight: bold;
        padding: 9px;
        border: none;
        font-size: 13px;
    }}
"""


# =====================================================================
# صفحه ۰: داشبورد مدیریتی و جدول هشدار موجودی (DashboardPage)
# =====================================================================
class DashboardPage(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.build_ui()
        self.refresh_data()

    def build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(16)

        cards_layout = QHBoxLayout()
        cards_layout.setSpacing(10)

        self.card_total_items = self.create_kpi_card("📦 کل اقلام انبار", "0", "#e8f8f5", "#0fb9b1", "#0a8d87")
        self.card_total_cats = self.create_kpi_card("🗂️ دسته‌بندی‌ها", "0", "#eef2ff", "#3867d6", "#2d55b8")
        self.card_normal = self.create_kpi_card("✅ موجودی عادی", "0", "#d4edda", "#28a745", "#155724")
        self.card_warning = self.create_kpi_card("⚠️ کمبود موجودی", "0", "#fff3cd", "#f39c12", "#856404")
        self.card_critical = self.create_kpi_card("🚨 وضعیت بحرانی", "0", "#f8d7da", "#eb3b5a", "#721c24")
        self.card_out = self.create_kpi_card("🚫 تمام‌شده (صفر)", "0", "#dcdde1", "#2f3640", "#1e272e")

        for c in (
            self.card_total_items, self.card_total_cats, self.card_normal,
            self.card_warning, self.card_critical, self.card_out
        ):
            cards_layout.addWidget(c["frame"])

        layout.addLayout(cards_layout)

        header_bar = QFrame()
        header_bar.setStyleSheet(f"""
            QFrame {{
                background-color: #ffffff;
                border: 1px solid #dfe4ea;
                border-radius: 10px;
                padding: 10px 14px;
            }}
            QLineEdit, QComboBox {{
                padding: 6px 10px;
                border: 1px solid #ced6e0;
                border-radius: 6px;
                background-color: #f8f9fa;
                font-family: '{FONT_TEXT}';
                font-size: 15px;
            }}
        """)
        hb_layout = QHBoxLayout(header_bar)
        hb_layout.setSpacing(10)

        lbl_table_title = QLabel("🔔 جدول هشدار موجودی و فهرست کالاهایی که زودتر تمام می‌شوند")
        lbl_table_title.setFont(QFont(FONT_TITLE, 13, QFont.Bold))
        lbl_table_title.setStyleSheet(f"color: #2f3542; font-family: '{FONT_TITLE}'; border: none;")
        hb_layout.addWidget(lbl_table_title)

        hb_layout.addStretch()

        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("🔍 جست‌وجوی سریع کالا...")
        self.txt_search.setMinimumWidth(190)
        self.txt_search.textChanged.connect(self.load_dashboard_table)
        hb_layout.addWidget(self.txt_search)

        hb_layout.addWidget(QLabel("نمایش:"))
        self.cmb_mode = QComboBox()
        self.cmb_mode.addItem("همه کالاها (به ترتیب زودترین تاریخ اتمام)", "ALL_SORTED")
        self.cmb_mode.addItem("🚨 فقط هشدارها (تمام‌شده + بحرانی + کمبود موجودی)", "ALERTS_ONLY")
        self.cmb_mode.addItem("🚫 فقط کالاهای تمام‌شده", "OUT_OF_STOCK")
        self.cmb_mode.addItem("🚨 فقط کالاهای بحرانی", "CRITICAL")
        self.cmb_mode.addItem("⚠️ فقط کالاهای دارای کمبود موجودی", "WARNING")
        self.cmb_mode.addItem("✅ فقط کالاهای با وضعیت عادی", "NORMAL")
        self.cmb_mode.currentIndexChanged.connect(self.load_dashboard_table)
        hb_layout.addWidget(self.cmb_mode)

        btn_refresh = QPushButton("🔄 بروزرسانی")
        btn_refresh.setCursor(Qt.PointingHandCursor)
        btn_refresh.setStyleSheet(f"""
            QPushButton {{
                background-color: #0fb9b1; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 7px 14px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #0a8d87; }}
        """)
        btn_refresh.clicked.connect(self.refresh_data)
        hb_layout.addWidget(btn_refresh)

        layout.addWidget(header_bar)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "نام کالا", "دسته‌بندی", "موجودی", "مصرف روزانه",
            "روزهای باقی‌مانده", "تاریخ اتمام", "وضعیت"
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setStyleSheet(TABLE_STYLE)
        layout.addWidget(self.table, 1)

        self.lbl_summary_banner = QLabel("")
        self.lbl_summary_banner.setFont(QFont(FONT_TEXT, 13, QFont.Bold))
        self.lbl_summary_banner.setStyleSheet(f"""
            background-color: #ffffff;
            color: #2f3542;
            font-family: '{FONT_TEXT}';
            font-size: 15px;
            font-weight: bold;
            padding: 10px 15px;
            border-radius: 8px;
            border: 1px solid #dfe4ea;
        """)
        layout.addWidget(self.lbl_summary_banner)

    def create_kpi_card(self, title, value, bg_hex, border_hex, text_hex):
        frame = QFrame()
        frame.setStyleSheet(f"""
            QFrame {{
                background-color: {bg_hex};
                border: 1px solid {border_hex};
                border-top: 4px solid {border_hex};
                border-radius: 10px;
                padding: 10px 8px;
            }}
            QLabel {{
                border: none;
                color: {text_hex};
                background: transparent;
            }}
        """)
        v_layout = QVBoxLayout(frame)
        v_layout.setSpacing(6)
        v_layout.setContentsMargins(6, 6, 6, 6)

        lbl_title = QLabel(title)
        lbl_title.setFont(QFont(FONT_TITLE, 11, QFont.Bold))
        lbl_title.setStyleSheet(f"font-family: '{FONT_TITLE}'; font-size: 13px;")
        lbl_title.setAlignment(Qt.AlignCenter)

        lbl_value = QLabel(value)
        lbl_value.setFont(QFont(FONT_TEXT, 17, QFont.Bold))
        lbl_value.setStyleSheet(f"font-family: '{FONT_TEXT}'; font-size: 19px; font-weight: bold;")
        lbl_value.setAlignment(Qt.AlignCenter)

        v_layout.addWidget(lbl_title)
        v_layout.addWidget(lbl_value)
        return {"frame": frame, "title": lbl_title, "value": lbl_value}

    def refresh_data(self):
        self.load_dashboard_table()

    def load_dashboard_table(self):
        crit_days = int(self.db.get_setting("critical_days", 7))
        warn_days = int(self.db.get_setting("warning_days", 14))

        all_items = self.db.get_items(include_inactive=False)
        all_categories = self.db.get_categories()

        enriched_all = [enrich_item_with_calculations(it, crit_days, warn_days) for it in all_items]

        cnt_normal = sum(1 for x in enriched_all if x["status_code"] == "NORMAL")
        cnt_warning = sum(1 for x in enriched_all if x["status_code"] == "WARNING")
        cnt_critical = sum(1 for x in enriched_all if x["status_code"] == "CRITICAL")
        cnt_out = sum(1 for x in enriched_all if x["status_code"] == "OUT_OF_STOCK")

        self.card_total_items["value"].setText(f"{len(all_items)} قلم")
        self.card_total_cats["value"].setText(f"{len(all_categories)} دسته")
        self.card_normal["value"].setText(f"{cnt_normal} قلم")
        self.card_warning["value"].setText(f"{cnt_warning} قلم")
        self.card_critical["value"].setText(f"{cnt_critical} قلم")
        self.card_out["value"].setText(f"{cnt_out} قلم")

        search_txt = self.txt_search.text().strip().lower()
        mode = self.cmb_mode.currentData()

        filtered = enriched_all
        if search_txt:
            filtered = [
                x for x in filtered
                if search_txt in x["name"].lower()
                or search_txt in x["item_code"].lower()
                or search_txt in (x.get("category_name") or "").lower()
            ]

        if mode == "ALERTS_ONLY":
            filtered = [x for x in filtered if x["status_code"] in ("OUT_OF_STOCK", "CRITICAL", "WARNING")]
        elif mode in ("OUT_OF_STOCK", "CRITICAL", "WARNING", "NORMAL"):
            filtered = [x for x in filtered if x["status_code"] == mode]

        filtered.sort(key=lambda x: (x["sort_priority"], x["name"]))

        self.table.setRowCount(len(filtered))
        for row_idx, info in enumerate(filtered):
            unit = info["unit"]
            row_vals = [
                f"{info['name']} ({info['item_code']})",
                info.get("category_name", "-"),
                f"{info['stock']:g} {unit}",
                f"{info['daily_usage']:g} {unit}",
                info["days_display"],
                info["depletion_date"],
                info["status_text"]
            ]

            for col_idx, val in enumerate(row_vals):
                cell = QTableWidgetItem(str(val))
                cell.setTextAlignment(Qt.AlignCenter)

                if col_idx in (4, 5, 6):
                    cell.setBackground(QBrush(QColor(info["color_bg"])))
                    cell.setForeground(QBrush(QColor(info["color_fg"])))
                    f = cell.font()
                    f.setBold(True)
                    cell.setFont(f)

                self.table.setItem(row_idx, col_idx, cell)

            self.table.setRowHeight(row_idx, 38)

        urgent_count = cnt_out + cnt_critical + cnt_warning
        if len(all_items) == 0:
            self.lbl_summary_banner.setText(
                "👋 به سامانه انبارداری خوش آمدید! هنوز کالایی ثبت نشده است. از منوی «مدیریت کالاها» اولین کالای انبار را ثبت کنید."
            )
        elif urgent_count > 0:
            self.lbl_summary_banner.setText(
                f"🔔 خلاصه وضعیت انبار: در حال حاضر {urgent_count} قلم کالا در وضعیت هشدار "
                f"({cnt_out} تمام‌شده، {cnt_critical} بحرانی، {cnt_warning} کم‌موجودی) قرار دارند و نیازمند شارژ هستند."
            )
        else:
            self.lbl_summary_banner.setText(
                "✅ وضعیت موجودی تمام اقلام انبار پایدار و عادی است و هیچ کمبود فوری وجود ندارد."
            )


# =====================================================================
# پنجره پاپ‌آپ افزودن / ویرایش کالا
# =====================================================================
class ItemDialog(QDialog):
    def __init__(self, db, item_data=None, parent=None):
        super().__init__(parent)
        self.db = db
        self.item_data = item_data
        self.is_edit = item_data is not None

        self.setWindowTitle("ویرایش اطلاعات کالا" if self.is_edit else "افزودن کالای جدید به انبار")
        self.setMinimumWidth(480)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setStyleSheet(f"""
            QDialog {{ background-color: #ffffff; font-family: '{FONT_TEXT}'; }}
            QLabel {{ font-family: '{FONT_TEXT}'; font-size: 15px; font-weight: bold; color: #2c3e50; }}
            QLineEdit, QComboBox, QDoubleSpinBox, QTextEdit {{
                padding: 7px; border: 1px solid #bdc3c7; border-radius: 6px;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                background-color: #fcfcfc;
            }}
            QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus, QTextEdit:focus {{
                border: 1px solid #0fb9b1; background-color: #ffffff;
            }}
        """)

        layout = QVBoxLayout(self)

        lbl_dlg_header = QLabel("✏️ ویرایش اطلاعات کالا" if self.is_edit else "➕ افزودن کالای جدید به انبار")
        lbl_dlg_header.setFont(QFont(FONT_TITLE, 13, QFont.Bold))
        lbl_dlg_header.setStyleSheet(f"font-family: '{FONT_TITLE}'; font-size: 15px; color: #0fb9b1; padding: 10px 15px 0 15px;")
        layout.addWidget(lbl_dlg_header)

        form = QFormLayout()
        form.setSpacing(12)
        form.setContentsMargins(15, 10, 15, 15)

        self.txt_code = QLineEdit()
        if self.is_edit:
            self.txt_code.setText(self.item_data["item_code"])
        else:
            self.txt_code.setText(self.db.generate_next_item_code())

        self.txt_name = QLineEdit()
        self.txt_name.setPlaceholderText("مثال: دستکش یکبارمصرف لاتکس")
        if self.is_edit:
            self.txt_name.setText(self.item_data["name"])

        self.cmb_category = QComboBox()
        categories = self.db.get_categories()
        for c in categories:
            self.cmb_category.addItem(c["name"], c["id"])
        if self.is_edit:
            idx = self.cmb_category.findData(self.item_data["category_id"])
            if idx >= 0:
                self.cmb_category.setCurrentIndex(idx)

        self.spn_stock = QDoubleSpinBox()
        self.spn_stock.setRange(0, 999999999)
        self.spn_stock.setDecimals(2)
        if self.is_edit:
            self.spn_stock.setValue(float(self.item_data["stock"]))
            self.spn_stock.setEnabled(False)
            self.spn_stock.setToolTip("تغییر موجودی پس از تعریف اولیه، فقط از بخش «ثبت ورود و خروج» انجام می‌شود.")

        self.cmb_unit = QComboBox()
        self.cmb_unit.setEditable(True)
        standard_units = ["عدد", "بسته", "جعبه", "کارتن", "کیلوگرم", "گرم", "لیتر", "متر", "جفت", "رول", "دستگاه"]
        self.cmb_unit.addItems(standard_units)
        if self.is_edit:
            self.cmb_unit.setCurrentText(self.item_data["unit"])

        self.spn_min_stock = QDoubleSpinBox()
        self.spn_min_stock.setRange(0, 999999999)
        self.spn_min_stock.setDecimals(2)
        if self.is_edit:
            self.spn_min_stock.setValue(float(self.item_data["min_stock"]))

        self.spn_daily_usage = QDoubleSpinBox()
        self.spn_daily_usage.setRange(0, 999999999)
        self.spn_daily_usage.setDecimals(2)
        if self.is_edit:
            self.spn_daily_usage.setValue(float(self.item_data["daily_usage"]))

        self.txt_desc = QTextEdit()
        self.txt_desc.setFixedHeight(70)
        self.txt_desc.setPlaceholderText("توضیحات اختیاری (محل نگهداری، برند، ویژگی‌ها و...)")
        if self.is_edit:
            self.txt_desc.setPlainText(self.item_data.get("description", ""))

        form.addRow("شناسه اختصاصی کالا *:", self.txt_code)
        form.addRow("نام کالا *:", self.txt_name)
        form.addRow("دسته‌بندی *:", self.cmb_category)
        stock_label = "موجودی فعلی (قفل):" if self.is_edit else "موجودی اولیه *:"
        form.addRow(stock_label, self.spn_stock)
        form.addRow("واحد شمارش *:", self.cmb_unit)
        form.addRow("حداقل موجودی مجاز:", self.spn_min_stock)
        form.addRow("میزان مصرف روزانه:", self.spn_daily_usage)
        form.addRow("توضیحات اختیاری:", self.txt_desc)

        layout.addLayout(form)

        if self.is_edit:
            lbl_hint = QLabel("💡 نکته: برای افزایش یا کاهش تعداد موجودی، از منوی «ثبت ورود و خروج» استفاده کنید.")
            lbl_hint.setStyleSheet(f"color: #7f8c8d; font-family: '{FONT_TEXT}'; font-size: 13px; font-weight: normal; padding: 0 15px;")
            layout.addWidget(lbl_hint)

        btn_layout = QHBoxLayout()
        btn_layout.setContentsMargins(15, 10, 15, 15)

        btn_save = QPushButton("💾 ذخیره اطلاعات")
        btn_save.setCursor(Qt.PointingHandCursor)
        btn_save.setStyleSheet(f"""
            QPushButton {{
                background-color: #0fb9b1; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 9px 22px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #0a8d87; }}
        """)
        btn_save.clicked.connect(self.save_item)

        btn_cancel = QPushButton("انصراف")
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_cancel.setStyleSheet(f"""
            QPushButton {{
                background-color: #dfe4ea; color: #2f3542; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 9px 20px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #ced6e0; }}
        """)
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def save_item(self):
        code = self.txt_code.text().strip()
        name = self.txt_name.text().strip()
        cat_id = self.cmb_category.currentData()
        stock = self.spn_stock.value()
        unit = self.cmb_unit.currentText().strip()
        min_stock = self.spn_min_stock.value()
        daily_usage = self.spn_daily_usage.value()
        desc = self.txt_desc.toPlainText().strip()

        if not code or not name or not unit:
            QMessageBox.warning(self, "خطای اعتبارسنجی", "لطفاً شناسه کالا، نام کالا و واحد شمارش را وارد کنید.")
            return

        if self.is_edit:
            ok, msg = self.db.update_item(
                self.item_data["id"], code, name, cat_id, unit, min_stock, daily_usage, desc
            )
        else:
            ok, msg = self.db.add_item(
                code, name, cat_id, stock, unit, min_stock, daily_usage, desc
            )

        if ok:
            QMessageBox.information(self, "عملیات موفق", msg)
            self.accept()
        else:
            QMessageBox.critical(self, "خطا در ثبت", msg)


# =====================================================================
# صفحه ۱: مدیریت کالاها (ItemsManagementPage)
# =====================================================================
class ItemsManagementPage(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.build_ui()
        self.refresh_data()

    def build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        toolbar = QFrame()
        toolbar.setStyleSheet(f"""
            QFrame {{
                background-color: #ffffff;
                border: 1px solid #dfe4ea;
                border-radius: 10px;
                padding: 10px;
            }}
            QLineEdit, QComboBox {{
                padding: 7px 10px;
                border: 1px solid #ced6e0;
                border-radius: 6px;
                font-family: '{FONT_TEXT}';
                font-size: 15px;
                background-color: #f8f9fa;
            }}
        """)
        tb_layout = QHBoxLayout(toolbar)
        tb_layout.setSpacing(10)

        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("🔍 جست‌وجوی نام، کد کالا یا توضیحات...")
        self.txt_search.textChanged.connect(self.load_items)

        self.cmb_filter_cat = QComboBox()
        self.cmb_filter_cat.setMinimumWidth(160)
        self.cmb_filter_cat.currentIndexChanged.connect(self.load_items)

        self.cmb_sort = QComboBox()
        self.cmb_sort.setMinimumWidth(180)
        self.cmb_sort.addItem("مرتب‌سازی: نام کالا (الف تا ی)", "name_asc")
        self.cmb_sort.addItem("مرتب‌سازی: نام کالا (ی تا الف)", "name_desc")
        self.cmb_sort.addItem("مرتب‌سازی: کمترین موجودی", "stock_asc")
        self.cmb_sort.addItem("مرتب‌سازی: بیشترین موجودی", "stock_desc")
        self.cmb_sort.addItem("مرتب‌سازی: آخرین به‌روزرسانی", "updated_desc")
        self.cmb_sort.addItem("مرتب‌سازی: کد کالا", "code_asc")
        self.cmb_sort.currentIndexChanged.connect(self.load_items)

        btn_add = QPushButton("➕ افزودن کالای جدید")
        btn_add.setCursor(Qt.PointingHandCursor)
        btn_add.setStyleSheet(f"""
            QPushButton {{
                background-color: #20bf6b; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 8px 16px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #1eb980; }}
        """)
        btn_add.clicked.connect(self.open_add_dialog)

        btn_edit = QPushButton("✏️ ویرایش کالا")
        btn_edit.setCursor(Qt.PointingHandCursor)
        btn_edit.setStyleSheet(f"""
            QPushButton {{
                background-color: #3867d6; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 8px 16px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #2d55b8; }}
        """)
        btn_edit.clicked.connect(self.open_edit_dialog)

        btn_delete = QPushButton("🗑️ حذف کالا")
        btn_delete.setCursor(Qt.PointingHandCursor)
        btn_delete.setStyleSheet(f"""
            QPushButton {{
                background-color: #eb3b5a; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 8px 16px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #d1304d; }}
        """)
        btn_delete.clicked.connect(self.delete_selected_item)

        tb_layout.addWidget(self.txt_search, 2)
        tb_layout.addWidget(QLabel("دسته:"))
        tb_layout.addWidget(self.cmb_filter_cat, 1)
        tb_layout.addWidget(self.cmb_sort, 1)
        tb_layout.addWidget(btn_add)
        tb_layout.addWidget(btn_edit)
        tb_layout.addWidget(btn_delete)

        layout.addWidget(toolbar)

        self.table = QTableWidget()
        self.table.setColumnCount(11)
        self.table.setHorizontalHeaderLabels([
            "شناسه", "کد کالا", "نام کالا", "دسته‌بندی", "موجودی فعلی",
            "واحد", "حداقل مجاز", "مصرف روزانه", "روزهای مانده",
            "تاریخ اتمام (شمسی)", "وضعیت"
        ])
        self.table.setColumnHidden(0, True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.table.setColumnWidth(2, 200)
        self.table.doubleClicked.connect(self.open_edit_dialog)
        self.table.setStyleSheet(TABLE_STYLE)
        layout.addWidget(self.table, 1)

        footer = QHBoxLayout()
        self.lbl_total_count = QLabel("تعداد کل اقلام ثبت‌شده: ۰ مورد")
        self.lbl_total_count.setFont(QFont(FONT_TEXT, 13, QFont.Bold))
        self.lbl_total_count.setStyleSheet(f"color: #2f3542; font-family: '{FONT_TEXT}'; font-size: 15px; font-weight: bold;")

        lbl_tip = QLabel("💡 راهنما: برای ویرایش سریع هر کالا، می‌توانید روی سطر آن دبل‌کلیک کنید.")
        lbl_tip.setStyleSheet(f"color: #7f8c8d; font-family: '{FONT_TEXT}'; font-size: 14px;")

        footer.addWidget(self.lbl_total_count)
        footer.addStretch()
        footer.addWidget(lbl_tip)
        layout.addLayout(footer)

    def refresh_data(self):
        current_cat = self.cmb_filter_cat.currentData()
        self.cmb_filter_cat.blockSignals(True)
        self.cmb_filter_cat.clear()
        self.cmb_filter_cat.addItem("همه دسته‌بندی‌ها", 0)
        for c in self.db.get_categories():
            self.cmb_filter_cat.addItem(c["name"], c["id"])
        if current_cat is not None:
            idx = self.cmb_filter_cat.findData(current_cat)
            if idx >= 0:
                self.cmb_filter_cat.setCurrentIndex(idx)
        self.cmb_filter_cat.blockSignals(False)
        self.load_items()

    def load_items(self):
        search_txt = self.txt_search.text()
        cat_id = self.cmb_filter_cat.currentData()
        sort_by = self.cmb_sort.currentData()

        crit_days = int(self.db.get_setting("critical_days", 7))
        warn_days = int(self.db.get_setting("warning_days", 14))

        raw_items = self.db.get_items(search_txt, cat_id, sort_by)
        self.table.setRowCount(len(raw_items))

        for row_idx, item in enumerate(raw_items):
            info = enrich_item_with_calculations(item, crit_days, warn_days)

            values = [
                str(info["id"]),
                info["item_code"],
                info["name"],
                info.get("category_name", "-"),
                f"{info['stock']:g}",
                info["unit"],
                f"{info['min_stock']:g}",
                f"{info['daily_usage']:g}",
                info["days_display"],
                info["depletion_date"],
                info["status_text"]
            ]

            for col_idx, val in enumerate(values):
                cell = QTableWidgetItem(str(val))
                cell.setTextAlignment(Qt.AlignCenter)
                if col_idx == 2:
                    cell.setToolTip(info.get("description") or info["name"])

                if col_idx in (8, 9, 10):
                    cell.setBackground(QBrush(QColor(info["color_bg"])))
                    cell.setForeground(QBrush(QColor(info["color_fg"])))
                    font = cell.font()
                    font.setBold(True)
                    cell.setFont(font)

                self.table.setItem(row_idx, col_idx, cell)

            self.table.setRowHeight(row_idx, 38)

        all_active_count = len(self.db.get_items())
        self.lbl_total_count.setText(
            f"📦 تعداد کل اقلام فعال در انبار: {all_active_count} قلم  |  تعداد نمایش داده شده در جدول: {len(raw_items)} قلم"
        )

    def get_selected_item_id(self):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return None
        row = selected_rows[0].row()
        return int(self.table.item(row, 0).text())

    def open_add_dialog(self):
        dlg = ItemDialog(self.db, None, self)
        if dlg.exec() == QDialog.Accepted:
            self.load_items()

    def open_edit_dialog(self):
        item_id = self.get_selected_item_id()
        if not item_id:
            QMessageBox.warning(self, "انتخاب کالا", "لطفاً ابتدا یک کالا را از داخل جدول انتخاب کنید.")
            return
        item_data = self.db.get_item_by_id(item_id)
        if not item_data:
            return
        dlg = ItemDialog(self.db, item_data, self)
        if dlg.exec() == QDialog.Accepted:
            self.load_items()

    def delete_selected_item(self):
        item_id = self.get_selected_item_id()
        if not item_id:
            QMessageBox.warning(self, "انتخاب کالا", "لطفاً ابتدا کالای مورد نظر را از جدول انتخاب کنید.")
            return
        item_data = self.db.get_item_by_id(item_id)
        if not item_data:
            return

        reply = QMessageBox.question(
            self,
            "تأیید حذف کالا",
            f"آیا از حذف کالای «{item_data['name']}» (کد: {item_data['item_code']}) اطمینان دارید؟\n\n"
            "توجه: این کالا به صورت امن (حذف منطقی) از فهرست فعال خارج می‌شود تا سوابق تراکنش‌های قبلی آن در تاریخچه انبار از بین نرود.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            ok, msg = self.db.soft_delete_item(item_id)
            if ok:
                QMessageBox.information(self, "حذف موفق", msg)
                self.load_items()
            else:
                QMessageBox.critical(self, "خطا", msg)


# =====================================================================
# پنجره انتخاب دسته جایگزین هنگام حذف دسته‌بندی دارای کالا
# =====================================================================
class ReplaceCategoryDialog(QDialog):
    def __init__(self, cat_to_delete, other_categories, parent=None):
        super().__init__(parent)
        self.setWindowTitle("انتخاب دسته‌بندی جایگزین")
        self.setMinimumWidth(440)
        self.setLayoutDirection(Qt.RightToLeft)
        self.selected_replacement_id = None

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        lbl_info = QLabel(
            f"⚠️ در دسته‌بندی «{cat_to_delete['name']}» تعداد {cat_to_delete['total_items_count']} کالا (فعال یا بایگانی شده) ثبت شده است.\n\n"
            "برای جلوگیری از خراب شدن اطلاعات کالاها، لطفاً مشخص کنید کالاهای این دسته به کدام دسته‌بندی منتقل شوند:"
        )
        lbl_info.setWordWrap(True)
        lbl_info.setStyleSheet(f"color: #2c3e50; font-family: '{FONT_TEXT}'; font-size: 15px; font-weight: bold;")
        layout.addWidget(lbl_info)

        self.cmb_replacement = QComboBox()
        self.cmb_replacement.setStyleSheet(f"padding: 8px; font-family: '{FONT_TEXT}'; font-size: 15px;")
        for c in other_categories:
            self.cmb_replacement.addItem(c["name"], c["id"])
        layout.addWidget(self.cmb_replacement)

        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("✅ انتقال کالاها و حذف دسته")
        btn_ok.setStyleSheet(f"background-color: #eb3b5a; color: white; font-family: '{FONT_TEXT}'; font-size: 15px; font-weight: bold; padding: 8px 16px; border-radius: 6px;")
        btn_ok.clicked.connect(self.confirm)

        btn_cancel = QPushButton("انصراف")
        btn_cancel.setStyleSheet(f"background-color: #dfe4ea; color: #2f3542; font-family: '{FONT_TEXT}'; font-size: 15px; font-weight: bold; padding: 8px 16px; border-radius: 6px;")
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def confirm(self):
        self.selected_replacement_id = self.cmb_replacement.currentData()
        self.accept()


# =====================================================================
# صفحه ۲: مدیریت دسته‌بندی‌ها (CategoriesPage)
# =====================================================================
class CategoriesPage(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.editing_cat_id = None
        self.build_ui()
        self.refresh_data()

    def build_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)

        form_card = QFrame()
        form_card.setFixedWidth(340)
        form_card.setStyleSheet(f"""
            QFrame {{
                background-color: #ffffff;
                border: 1px solid #dfe4ea;
                border-radius: 10px;
                padding: 15px;
            }}
            QLabel {{ border: none; font-family: '{FONT_TEXT}'; font-size: 15px; font-weight: bold; color: #2f3542; }}
            QLineEdit, QTextEdit {{
                padding: 8px; border: 1px solid #ced6e0; border-radius: 6px;
                background-color: #f8f9fa; font-family: '{FONT_TEXT}'; font-size: 15px;
            }}
        """)
        form_layout = QVBoxLayout(form_card)
        form_layout.setSpacing(12)

        self.lbl_form_title = QLabel("➕ ایجاد دسته‌بندی جدید")
        self.lbl_form_title.setFont(QFont(FONT_TITLE, 13, QFont.Bold))
        self.lbl_form_title.setStyleSheet(f"color: #0fb9b1; font-family: '{FONT_TITLE}'; font-size: 15px; padding-bottom: 5px;")
        form_layout.addWidget(self.lbl_form_title)

        form_layout.addWidget(QLabel("نام دسته‌بندی *:"))
        self.txt_cat_name = QLineEdit()
        self.txt_cat_name.setPlaceholderText("مثال: قطعات الکترونیکی")
        form_layout.addWidget(self.txt_cat_name)

        form_layout.addWidget(QLabel("توضیحات دسته‌بندی:"))
        self.txt_cat_desc = QTextEdit()
        self.txt_cat_desc.setFixedHeight(90)
        self.txt_cat_desc.setPlaceholderText("توضیح کوتاه درباره اقلام این دسته...")
        form_layout.addWidget(self.txt_cat_desc)

        self.btn_save_cat = QPushButton("💾 ثبت دسته‌بندی جدید")
        self.btn_save_cat.setCursor(Qt.PointingHandCursor)
        self.btn_save_cat.setStyleSheet(f"""
            QPushButton {{
                background-color: #20bf6b; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 10px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #1eb980; }}
        """)
        self.btn_save_cat.clicked.connect(self.save_category)
        form_layout.addWidget(self.btn_save_cat)

        self.btn_cancel_edit = QPushButton("انصراف از ویرایش")
        self.btn_cancel_edit.setCursor(Qt.PointingHandCursor)
        self.btn_cancel_edit.setVisible(False)
        self.btn_cancel_edit.setStyleSheet(f"""
            QPushButton {{
                background-color: #dfe4ea; color: #2f3542; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 8px; border-radius: 6px; border: none;
            }}
        """)
        self.btn_cancel_edit.clicked.connect(self.reset_form)
        form_layout.addWidget(self.btn_cancel_edit)

        form_layout.addStretch()

        table_container = QVBoxLayout()
        table_container.setSpacing(12)

        top_bar = QHBoxLayout()
        lbl_list_title = QLabel("🗂️ فهرست دسته‌بندی‌های تعریف‌شده در انبار")
        lbl_list_title.setFont(QFont(FONT_TITLE, 13, QFont.Bold))
        lbl_list_title.setStyleSheet(f"font-family: '{FONT_TITLE}'; font-size: 14px; color: #2f3542;")

        btn_edit = QPushButton("✏️ ویرایش دسته انتخاب‌شده")
        btn_edit.setCursor(Qt.PointingHandCursor)
        btn_edit.setStyleSheet(f"""
            QPushButton {{
                background-color: #3867d6; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 8px 14px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #2d55b8; }}
        """)
        btn_edit.clicked.connect(self.start_edit_selected)

        btn_delete = QPushButton("🗑️ حذف دسته انتخاب‌شده")
        btn_delete.setCursor(Qt.PointingHandCursor)
        btn_delete.setStyleSheet(f"""
            QPushButton {{
                background-color: #eb3b5a; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 8px 14px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #d1304d; }}
        """)
        btn_delete.clicked.connect(self.delete_selected_category)

        top_bar.addWidget(lbl_list_title)
        top_bar.addStretch()
        top_bar.addWidget(btn_edit)
        top_bar.addWidget(btn_delete)
        table_container.addLayout(top_bar)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["شناسه", "نام دسته‌بندی", "تعداد کالاهای فعال", "توضیحات"])
        self.table.setColumnHidden(0, True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.doubleClicked.connect(self.start_edit_selected)
        self.table.setStyleSheet(TABLE_STYLE)
        table_container.addWidget(self.table, 1)

        layout.addWidget(form_card)
        layout.addLayout(table_container, 1)

    def refresh_data(self):
        self.categories_cache = self.db.get_categories()
        self.table.setRowCount(len(self.categories_cache))
        for idx, cat in enumerate(self.categories_cache):
            row_vals = [
                str(cat["id"]),
                cat["name"],
                f"{cat['active_items_count']} کالا",
                cat.get("description", "")
            ]
            for col, val in enumerate(row_vals):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(idx, col, item)
            self.table.setRowHeight(idx, 38)

    def reset_form(self):
        self.editing_cat_id = None
        self.lbl_form_title.setText("➕ ایجاد دسته‌بندی جدید")
        self.btn_save_cat.setText("💾 ثبت دسته‌بندی جدید")
        self.txt_cat_name.clear()
        self.txt_cat_desc.clear()
        self.btn_cancel_edit.setVisible(False)

    def save_category(self):
        name = self.txt_cat_name.text().strip()
        desc = self.txt_cat_desc.toPlainText().strip()

        if not name:
            QMessageBox.warning(self, "خطا", "لطفاً نام دسته‌بندی را وارد کنید.")
            return

        if self.editing_cat_id is None:
            ok, msg = self.db.add_category(name, desc)
        else:
            ok, msg = self.db.update_category(self.editing_cat_id, name, desc)

        if ok:
            QMessageBox.information(self, "موفق", msg)
            self.reset_form()
            self.refresh_data()
        else:
            QMessageBox.critical(self, "خطا", msg)

    def get_selected_category(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return None
        cat_id = int(self.table.item(rows[0].row(), 0).text())
        for c in self.categories_cache:
            if c["id"] == cat_id:
                return c
        return None

    def start_edit_selected(self):
        cat = self.get_selected_category()
        if not cat:
            QMessageBox.warning(self, "انتخاب دسته‌بندی", "لطفاً یک دسته‌بندی را از جدول انتخاب کنید.")
            return
        self.editing_cat_id = cat["id"]
        self.lbl_form_title.setText(f"✏️ ویرایش دسته: {cat['name']}")
        self.btn_save_cat.setText("💾 ذخیره تغییرات")
        self.txt_cat_name.setText(cat["name"])
        self.txt_cat_desc.setPlainText(cat.get("description", ""))
        self.btn_cancel_edit.setVisible(True)

    def delete_selected_category(self):
        cat = self.get_selected_category()
        if not cat:
            QMessageBox.warning(self, "انتخاب دسته‌بندی", "لطفاً ابتدا دسته‌بندی مورد نظر را از جدول انتخاب کنید.")
            return

        if len(self.categories_cache) <= 1:
            QMessageBox.warning(self, "خطا", "حداقل یک دسته‌بندی باید در سیستم باقی بماند.")
            return

        if cat["total_items_count"] > 0:
            other_cats = [c for c in self.categories_cache if c["id"] != cat["id"]]
            dlg = ReplaceCategoryDialog(cat, other_cats, self)
            if dlg.exec() == QDialog.Accepted and dlg.selected_replacement_id:
                ok, msg = self.db.delete_category(cat["id"], dlg.selected_replacement_id)
                if ok:
                    QMessageBox.information(self, "موفق", msg)
                    self.reset_form()
                    self.refresh_data()
                else:
                    QMessageBox.critical(self, "خطا", msg)
        else:
            reply = QMessageBox.question(
                self,
                "تأیید حذف دسته‌بندی",
                f"آیا از حذف دسته‌بندی «{cat['name']}» اطمینان دارید؟",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                ok, msg = self.db.delete_category(cat["id"])
                if ok:
                    QMessageBox.information(self, "موفق", msg)
                    self.reset_form()
                    self.refresh_data()
                else:
                    QMessageBox.critical(self, "خطا", msg)


# =====================================================================
# صفحه ۳: سیستم ثبت ورود و خروج کالا (TransactionsPage)
# =====================================================================
class TransactionsPage(QWidget):
    TX_LABELS = {
        "IN": ("📥 ورود به انبار", "#d4edda", "#155724"),
        "OUT": ("📤 خروج از انبار", "#f8d7da", "#721c24"),
        "USAGE": ("📉 ثبت مصرف روزانه", "#fff3cd", "#856404"),
        "ADJUST_IN": ("➕ اصلاحیه افزایشی", "#d1ecf1", "#0c5460"),
        "ADJUST_OUT": ("➖ اصلاحیه کاهشی", "#f5c6cb", "#721c24"),
    }

    def __init__(self, db):
        super().__init__()
        self.db = db
        self.items_map = {}
        self.transactions_cache = []
        self.build_ui()
        self.refresh_data()

    def build_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)

        form_card = QFrame()
        form_card.setFixedWidth(360)
        form_card.setStyleSheet(f"""
            QFrame {{
                background-color: #ffffff;
                border: 1px solid #dfe4ea;
                border-radius: 10px;
                padding: 15px;
            }}
            QLabel {{ border: none; font-family: '{FONT_TEXT}'; font-weight: bold; color: #2f3542; font-size: 15px; }}
            QLineEdit, QComboBox, QDoubleSpinBox, QTextEdit {{
                padding: 8px; border: 1px solid #ced6e0; border-radius: 6px;
                background-color: #f8f9fa; font-family: '{FONT_TEXT}'; font-size: 15px;
            }}
            QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus, QTextEdit:focus {{
                border: 1px solid #0fb9b1; background-color: #ffffff;
            }}
        """)
        form_layout = QVBoxLayout(form_card)
        form_layout.setSpacing(11)

        lbl_title = QLabel("🔄 ثبت تراکنش جدید انبار")
        lbl_title.setFont(QFont(FONT_TITLE, 13, QFont.Bold))
        lbl_title.setStyleSheet(f"color: #0fb9b1; font-family: '{FONT_TITLE}'; font-size: 15px; padding-bottom: 4px;")
        form_layout.addWidget(lbl_title)

        form_layout.addWidget(QLabel("انتخاب کالا *:"))
        self.cmb_item = QComboBox()
        self.cmb_item.currentIndexChanged.connect(self.update_selected_item_info)
        form_layout.addWidget(self.cmb_item)

        self.lbl_stock_info = QLabel("موجودی فعلی: -")
        self.lbl_stock_info.setStyleSheet(f"""
            background-color: #e8f8f5;
            color: #0a8d87;
            font-family: '{FONT_TEXT}';
            font-size: 15px;
            padding: 8px;
            border-radius: 6px;
            border: 1px solid #a3e4d7;
            font-weight: bold;
        """)
        self.lbl_stock_info.setAlignment(Qt.AlignCenter)
        form_layout.addWidget(self.lbl_stock_info)

        form_layout.addWidget(QLabel("نوع تراکنش *:"))
        self.cmb_tx_type = QComboBox()
        self.cmb_tx_type.addItem("📥 ورود کالا به انبار (افزایش موجودی)", "IN")
        self.cmb_tx_type.addItem("📤 خروج کالا از انبار (کاهش موجودی)", "OUT")
        self.cmb_tx_type.addItem("📉 ثبت مصرف واقعی روزانه (کاهش موجودی)", "USAGE")
        self.cmb_tx_type.addItem("➕ سند اصلاحی افزایشی (رفع مغایرت)", "ADJUST_IN")
        self.cmb_tx_type.addItem("➖ سند اصلاحی کاهشی (رفع مغایرت)", "ADJUST_OUT")
        form_layout.addWidget(self.cmb_tx_type)

        form_layout.addWidget(QLabel("تعداد / مقدار تراکنش *:"))
        self.spn_qty = QDoubleSpinBox()
        self.spn_qty.setRange(0.01, 999999999)
        self.spn_qty.setDecimals(2)
        self.spn_qty.setValue(1.0)
        form_layout.addWidget(self.spn_qty)

        form_layout.addWidget(QLabel("توضیحات (تحویل‌گیرنده / شماره فاکتور / علت):"))
        self.txt_desc = QTextEdit()
        self.txt_desc.setFixedHeight(75)
        self.txt_desc.setPlaceholderText("مثال: تحویل به واحد فنی / خرید طبق فاکتور ۱۲۴...")
        form_layout.addWidget(self.txt_desc)

        btn_submit = QPushButton("✅ ثبت نهایی تراکنش در انبار")
        btn_submit.setCursor(Qt.PointingHandCursor)
        btn_submit.setStyleSheet(f"""
            QPushButton {{
                background-color: #20bf6b; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 16px;
                padding: 11px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #1eb980; }}
        """)
        btn_submit.clicked.connect(self.submit_transaction)
        form_layout.addWidget(btn_submit)

        lbl_rule = QLabel(
            "🛡️ قانون ایمنی انبار:\n"
            "تمام تراکنش‌ها به‌صورت دائمی ثبت می‌شوند و موجودی هیچ کالایی منفی نخواهد شد. "
            "اگر تراکنشی را به اشتباه ثبت کردید، آن را از جدول سمت چپ انتخاب کرده و دکمه «ثبت سند اصلاحی» را بزنید."
        )
        lbl_rule.setWordWrap(True)
        lbl_rule.setStyleSheet(f"color: #7f8c8d; font-family: '{FONT_TEXT}'; font-size: 13px; font-weight: normal; padding-top: 6px;")
        form_layout.addWidget(lbl_rule)

        form_layout.addStretch()

        right_panel = QVBoxLayout()
        right_panel.setSpacing(12)

        filter_bar = QFrame()
        filter_bar.setStyleSheet(f"""
            QFrame {{
                background-color: #ffffff;
                border: 1px solid #dfe4ea;
                border-radius: 10px;
                padding: 10px;
            }}
            QComboBox {{
                padding: 6px 10px; border: 1px solid #ced6e0; border-radius: 6px;
                background-color: #f8f9fa; font-family: '{FONT_TEXT}'; font-size: 15px;
            }}
        """)
        fb_layout = QHBoxLayout(filter_bar)

        fb_layout.addWidget(QLabel("فیلتر کالا:"))
        self.cmb_filter_item = QComboBox()
        self.cmb_filter_item.setMinimumWidth(180)
        self.cmb_filter_item.currentIndexChanged.connect(self.load_transactions)
        fb_layout.addWidget(self.cmb_filter_item)

        fb_layout.addWidget(QLabel("نوع عملیات:"))
        self.cmb_filter_type = QComboBox()
        self.cmb_filter_type.addItem("همه تراکنش‌ها", "ALL")
        self.cmb_filter_type.addItem("فقط ورودی‌ها (ورود + اصلاحیه افزایشی)", "ALL_IN")
        self.cmb_filter_type.addItem("فقط خروجی‌ها (خروج + مصرف + اصلاحیه کاهشی)", "ALL_OUT")
        self.cmb_filter_type.addItem("ورود به انبار", "IN")
        self.cmb_filter_type.addItem("خروج از انبار", "OUT")
        self.cmb_filter_type.addItem("ثبت مصرف روزانه", "USAGE")
        self.cmb_filter_type.currentIndexChanged.connect(self.load_transactions)
        fb_layout.addWidget(self.cmb_filter_type)

        fb_layout.addStretch()

        btn_reverse = QPushButton("↩️ ثبت سند اصلاحی (معکوس تراکنش انتخابی)")
        btn_reverse.setCursor(Qt.PointingHandCursor)
        btn_reverse.setStyleSheet(f"""
            QPushButton {{
                background-color: #f39c12; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 8px 14px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #e67e22; }}
        """)
        btn_reverse.clicked.connect(self.reverse_selected_transaction)
        fb_layout.addWidget(btn_reverse)

        right_panel.addWidget(filter_bar)

        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels([
            "شناسه", "تاریخ و ساعت (شمسی)", "کد کالا", "نام کالا",
            "نوع تراکنش", "تعداد", "موجودی قبل", "موجودی بعد", "توضیحات"
        ])
        self.table.setColumnHidden(0, True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Interactive)
        self.table.setColumnWidth(1, 140)
        self.table.horizontalHeader().setSectionResizeMode(8, QHeaderView.Interactive)
        self.table.setColumnWidth(8, 220)
        self.table.setStyleSheet(TABLE_STYLE)
        right_panel.addWidget(self.table, 1)

        self.lbl_tx_count = QLabel("تعداد تراکنش‌های نمایش داده شده: ۰ مورد")
        self.lbl_tx_count.setStyleSheet(f"color: #2f3542; font-family: '{FONT_TEXT}'; font-size: 15px; font-weight: bold;")
        right_panel.addWidget(self.lbl_tx_count)

        layout.addWidget(form_card)
        layout.addLayout(right_panel, 1)

    def refresh_data(self):
        current_item_id = self.cmb_item.currentData()
        filter_item_id = self.cmb_filter_item.currentData()

        active_items = self.db.get_items(include_inactive=False)
        all_items = self.db.get_items(include_inactive=True)
        self.items_map = {it["id"]: it for it in all_items}

        self.cmb_item.blockSignals(True)
        self.cmb_item.clear()
        for it in active_items:
            self.cmb_item.addItem(f"{it['name']} ({it['item_code']})", it["id"])
        if current_item_id:
            idx = self.cmb_item.findData(current_item_id)
            if idx >= 0:
                self.cmb_item.setCurrentIndex(idx)
        self.cmb_item.blockSignals(False)

        self.cmb_filter_item.blockSignals(True)
        self.cmb_filter_item.clear()
        self.cmb_filter_item.addItem("همه کالاها (فعال و بایگانی‌شده)", 0)
        for it in all_items:
            status_suffix = "" if it["is_active"] else " [حذف شده]"
            self.cmb_filter_item.addItem(f"{it['name']} ({it['item_code']}){status_suffix}", it["id"])
        if filter_item_id:
            idx = self.cmb_filter_item.findData(filter_item_id)
            if idx >= 0:
                self.cmb_filter_item.setCurrentIndex(idx)
        self.cmb_filter_item.blockSignals(False)

        self.update_selected_item_info()
        self.load_transactions()

    def update_selected_item_info(self):
        item_id = self.cmb_item.currentData()
        if not item_id or item_id not in self.items_map:
            self.lbl_stock_info.setText("لطفاً ابتدا یک کالا در بخش «مدیریت کالاها» تعریف کنید.")
            return
        item = self.db.get_item_by_id(item_id)
        if item:
            self.items_map[item_id] = item
            self.lbl_stock_info.setText(
                f"📦 موجودی فعلی انبار: {item['stock']:g} {item['unit']}   |   مصرف روزانه: {item['daily_usage']:g} {item['unit']}"
            )

    def submit_transaction(self):
        item_id = self.cmb_item.currentData()
        if not item_id:
            QMessageBox.warning(self, "انتخاب کالا", "لطفاً ابتدا یک کالا انتخاب کنید.")
            return

        tx_type = self.cmb_tx_type.currentData()
        qty = self.spn_qty.value()
        desc = self.txt_desc.toPlainText().strip()

        if tx_type in ("ADJUST_IN", "ADJUST_OUT") and not desc:
            QMessageBox.warning(self, "توضیحات الزامی", "برای ثبت سند اصلاحی، لطفاً دلیل اصلاح را در کادر توضیحات بنویسید.")
            return

        ok, msg = self.db.record_transaction(item_id, tx_type, qty, desc)
        if ok:
            QMessageBox.information(self, "ثبت موفق", msg)
            self.spn_qty.setValue(1.0)
            self.txt_desc.clear()
            self.refresh_data()
        else:
            QMessageBox.critical(self, "خطا در ثبت تراکنش", msg)

    def load_transactions(self):
        filter_item = self.cmb_filter_item.currentData()
        filter_type = self.cmb_filter_type.currentData()

        self.transactions_cache = self.db.get_transactions(item_id=filter_item, tx_type=filter_type, limit=500)
        self.table.setRowCount(len(self.transactions_cache))

        for row_idx, tx in enumerate(self.transactions_cache):
            tx_label, bg_hex, fg_hex = self.TX_LABELS.get(
                tx["tx_type"], (tx["tx_type"], "#ffffff", "#2f3542")
            )
            unit = tx.get("unit", "")
            item_name = tx["item_name"] + ("" if tx.get("is_active", 1) else " (بایگانی)")

            row_vals = [
                str(tx["id"]),
                iso_to_jalali(tx["created_at"], include_time=True),
                tx["item_code"],
                item_name,
                tx_label,
                f"{tx['quantity']:g} {unit}",
                f"{tx['stock_before']:g} {unit}",
                f"{tx['stock_after']:g} {unit}",
                tx.get("description", "")
            ]

            for col_idx, val in enumerate(row_vals):
                cell = QTableWidgetItem(str(val))
                cell.setTextAlignment(Qt.AlignCenter)
                if col_idx == 8:
                    cell.setToolTip(str(val))
                if col_idx == 4:
                    cell.setBackground(QBrush(QColor(bg_hex)))
                    cell.setForeground(QBrush(QColor(fg_hex)))
                    f = cell.font()
                    f.setBold(True)
                    cell.setFont(f)
                self.table.setItem(row_idx, col_idx, cell)

            self.table.setRowHeight(row_idx, 38)

        self.lbl_tx_count.setText(f"📋 تعداد تراکنش‌های نمایش داده شده در تاریخچه: {len(self.transactions_cache)} مورد")

    def reverse_selected_transaction(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.warning(self, "انتخاب تراکنش", "لطفاً ابتدا تراکنشی که می‌خواهید اصلاح (معکوس) شود را از جدول انتخاب کنید.")
            return

        tx_id = int(self.table.item(rows[0].row(), 0).text())
        target_tx = next((t for t in self.transactions_cache if t["id"] == tx_id), None)
        if not target_tx:
            return

        if target_tx["tx_type"] in ("IN", "ADJUST_IN"):
            rev_type = "ADJUST_OUT"
            rev_action_name = "سند اصلاحی کاهشی (کسر از موجودی)"
        else:
            rev_type = "ADJUST_IN"
            rev_action_name = "سند اصلاحی افزایشی (بازگشت به موجودی)"

        reply = QMessageBox.question(
            self,
            "تأیید ثبت سند اصلاحی (معکوس)",
            f"آیا می‌خواهید تراکنش شماره #{target_tx['id']} مربوط به کالای «{target_tx['item_name']}» "
            f"به تعداد {target_tx['quantity']:g} {target_tx['unit']} را معکوس کنید؟\n\n"
            f"عملیات انجام‌شدنی: {rev_action_name}\n"
            "توجه: برای حفظ قابلیت ردیابی، تراکنش قبلی در تاریخچه باقی می‌ماند و یک سند اصلاحی رسمی جدید ثبت خواهد شد.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            desc = f"سند اصلاحی (معکوس تراکنش #{target_tx['id']} مورخ {iso_to_jalali(target_tx['created_at'])})"
            ok, msg = self.db.record_transaction(target_tx["item_id"], rev_type, target_tx["quantity"], desc)
            if ok:
                QMessageBox.information(self, "اصلاح موفق", msg)
                self.refresh_data()
            else:
                QMessageBox.critical(self, "خطا در ثبت اصلاحیه", msg)


# =====================================================================
# صفحه ۴: گزارش مصرف و تحلیل نرخ مصرف واقعی (UsageReportPage)
# =====================================================================
class UsageReportPage(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.rows_cache = []
        self.build_ui()
        self.refresh_data()

    def build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        toolbar = QFrame()
        toolbar.setStyleSheet(f"""
            QFrame {{
                background-color: #ffffff;
                border: 1px solid #dfe4ea;
                border-radius: 10px;
                padding: 10px;
            }}
            QLineEdit, QComboBox {{
                padding: 7px 10px; border: 1px solid #ced6e0; border-radius: 6px;
                background-color: #f8f9fa; font-family: '{FONT_TEXT}'; font-size: 15px;
            }}
        """)
        tb_layout = QHBoxLayout(toolbar)
        tb_layout.setSpacing(10)

        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("🔍 جست‌وجوی نام یا کد کالا...")
        self.txt_search.textChanged.connect(self.load_usage_data)

        tb_layout.addWidget(self.txt_search, 2)
        tb_layout.addWidget(QLabel("بازه تحلیل مصرف واقعی:"))

        self.cmb_period = QComboBox()
        self.cmb_period.addItem("۷ روز گذشته (هفتگی)", 7)
        self.cmb_period.addItem("۳۰ روز گذشته (ماهانه)", 30)
        self.cmb_period.addItem("۹۰ روز گذشته (فصلی)", 90)
        self.cmb_period.addItem("۳۶۵ روز گذشته (سالانه)", 365)
        self.cmb_period.setCurrentIndex(1)
        self.cmb_period.currentIndexChanged.connect(self.load_usage_data)
        tb_layout.addWidget(self.cmb_period, 1)

        btn_edit_rate = QPushButton("✏️ ویرایش دستی مصرف روزانه")
        btn_edit_rate.setCursor(Qt.PointingHandCursor)
        btn_edit_rate.setStyleSheet(f"""
            QPushButton {{
                background-color: #3867d6; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 8px 14px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #2d55b8; }}
        """)
        btn_edit_rate.clicked.connect(self.edit_daily_usage_manually)
        tb_layout.addWidget(btn_edit_rate)

        btn_sync_actual = QPushButton("🔄 جایگزینی مصرف روزانه با میانگین واقعی")
        btn_sync_actual.setCursor(Qt.PointingHandCursor)
        btn_sync_actual.setStyleSheet(f"""
            QPushButton {{
                background-color: #0fb9b1; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 8px 14px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #0a8d87; }}
        """)
        btn_sync_actual.clicked.connect(self.sync_selected_with_actual)
        tb_layout.addWidget(btn_sync_actual)

        layout.addWidget(toolbar)

        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels([
            "شناسه", "کد کالا", "نام کالا", "دسته‌بندی", "موجودی فعلی",
            "مصرف روزانه ثبت‌شده", "کل خروجی در بازه", "میانگین مصرف واقعی روزانه",
            "تخمین مصرف ماهانه (۳۰ روز)", "وضعیت انحراف مصرف"
        ])
        self.table.setColumnHidden(0, True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.doubleClicked.connect(self.edit_daily_usage_manually)
        self.table.setStyleSheet(TABLE_STYLE)
        layout.addWidget(self.table, 1)

        lbl_help = QLabel(
            "💡 راهنما: ستون «میانگین مصرف واقعی روزانه» بر اساس تراکنش‌های خروج و مصرف ثبت‌شده در بازه زمانی انتخابی محاسبه می‌شود. "
            "برای تغییر سریع مصرف روزانه هر کالا می‌توانید روی سطر آن دبل‌کلیک کنید."
        )
        lbl_help.setStyleSheet(f"color: #7f8c8d; font-family: '{FONT_TEXT}'; font-size: 14px;")
        layout.addWidget(lbl_help)

    def refresh_data(self):
        self.load_usage_data()

    def load_usage_data(self):
        days = int(self.cmb_period.currentData() or 30)
        search_txt = self.txt_search.text().strip()
        items = self.db.get_items(search_text=search_txt, sort_by="name_asc")

        self.rows_cache = []
        self.table.setRowCount(len(items))

        for row_idx, it in enumerate(items):
            avg_actual = self.db.get_actual_daily_usage_avg(it["id"], days=days)
            total_out_period = round(avg_actual * days, 2)
            est_daily = float(it["daily_usage"])
            monthly_est = round(est_daily * 30, 2)
            unit = it["unit"]

            if est_daily == 0 and avg_actual == 0:
                diff_text = "بدون مصرف"
                bg, fg = "#f1f2f6", "#57606f"
            elif avg_actual > est_daily * 1.15:
                diff_text = "⚠️ مصرف واقعی بیشتر از برآورد"
                bg, fg = "#fff3cd", "#856404"
            elif est_daily > 0 and avg_actual < est_daily * 0.85:
                diff_text = "ℹ️ مصرف واقعی کمتر از برآورد"
                bg, fg = "#d1ecf1", "#0c5460"
            else:
                diff_text = "✅ متناسب با برآورد"
                bg, fg = "#d4edda", "#155724"

            row_info = dict(it)
            row_info["avg_actual"] = avg_actual
            self.rows_cache.append(row_info)

            vals = [
                str(it["id"]),
                it["item_code"],
                it["name"],
                it.get("category_name", "-"),
                f"{it['stock']:g} {unit}",
                f"{est_daily:g} {unit}",
                f"{total_out_period:g} {unit}",
                f"{avg_actual:g} {unit}",
                f"{monthly_est:g} {unit}",
                diff_text
            ]

            for col_idx, val in enumerate(vals):
                cell = QTableWidgetItem(str(val))
                cell.setTextAlignment(Qt.AlignCenter)
                if col_idx == 9:
                    cell.setBackground(QBrush(QColor(bg)))
                    cell.setForeground(QBrush(QColor(fg)))
                    f = cell.font()
                    f.setBold(True)
                    cell.setFont(f)
                self.table.setItem(row_idx, col_idx, cell)

            self.table.setRowHeight(row_idx, 38)

    def get_selected_row_info(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return None
        row_idx = rows[0].row()
        if 0 <= row_idx < len(self.rows_cache):
            return self.rows_cache[row_idx]
        return None

    def edit_daily_usage_manually(self):
        info = self.get_selected_row_info()
        if not info:
            QMessageBox.warning(self, "انتخاب کالا", "لطفاً ابتدا یک کالا را از جدول انتخاب کنید.")
            return

        new_val, ok = QInputDialog.getDouble(
            self,
            "ویرایش میزان مصرف روزانه",
            f"میزان مصرف روزانه برای کالای «{info['name']}» (برحسب {info['unit']}) را وارد کنید:",
            float(info["daily_usage"]),
            0.0,
            999999999.0,
            2
        )
        if ok:
            success, msg = self.db.update_item(
                info["id"], info["item_code"], info["name"], info["category_id"],
                info["unit"], info["min_stock"], new_val, info.get("description", "")
            )
            if success:
                self.load_usage_data()
            else:
                QMessageBox.critical(self, "خطا", msg)

    def sync_selected_with_actual(self):
        info = self.get_selected_row_info()
        if not info:
            QMessageBox.warning(self, "انتخاب کالا", "لطفاً ابتدا یک کالا را از جدول انتخاب کنید.")
            return

        avg_actual = info["avg_actual"]
        reply = QMessageBox.question(
            self,
            "تأیید همگام‌سازی نرخ مصرف",
            f"آیا می‌خواهید مصرف روزانه ثبت‌شده برای «{info['name']}» از {info['daily_usage']:g} به میانگین مصرف واقعی ({avg_actual:g} {info['unit']}) تغییر کند؟",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            success, msg = self.db.update_item(
                info["id"], info["item_code"], info["name"], info["category_id"],
                info["unit"], info["min_stock"], avg_actual, info.get("description", "")
            )
            if success:
                QMessageBox.information(self, "به‌روزرسانی موفق", "نرخ مصرف روزانه با موفقیت به‌روزرسانی شد.")
                self.load_usage_data()
            else:
                QMessageBox.critical(self, "خطا", msg)


# =====================================================================
# صفحه ۵: پیش‌بینی اتمام موجودی (DepletionForecastPage)
# =====================================================================
class DepletionForecastPage(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.build_ui()
        self.refresh_data()

    def build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        cards_layout = QHBoxLayout()
        cards_layout.setSpacing(12)

        self.card_out = self.create_stat_card("🚫 اتمام موجودی (صفر)", "0", "#dcdde1", "#2f3640")
        self.card_crit = self.create_stat_card("🚨 هشدار بحرانی (< ۷ روز)", "0", "#f8d7da", "#721c24")
        self.card_warn = self.create_stat_card("⚠️ کمبود موجودی (۷ تا ۱۴ روز)", "0", "#fff3cd", "#856404")
        self.card_norm = self.create_stat_card("✅ وضعیت عادی (> ۱۴ روز)", "0", "#d4edda", "#155724")

        cards_layout.addWidget(self.card_out["frame"])
        cards_layout.addWidget(self.card_crit["frame"])
        cards_layout.addWidget(self.card_warn["frame"])
        cards_layout.addWidget(self.card_norm["frame"])
        layout.addLayout(cards_layout)

        filter_bar = QFrame()
        filter_bar.setStyleSheet(f"""
            QFrame {{
                background-color: #ffffff;
                border: 1px solid #dfe4ea;
                border-radius: 10px;
                padding: 10px;
            }}
            QLineEdit, QComboBox {{
                padding: 7px 10px; border: 1px solid #ced6e0; border-radius: 6px;
                background-color: #f8f9fa; font-family: '{FONT_TEXT}'; font-size: 15px;
            }}
        """)
        fb_layout = QHBoxLayout(filter_bar)
        fb_layout.setSpacing(10)

        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("🔍 جست‌وجوی نام یا کد کالا...")
        self.txt_search.textChanged.connect(self.load_forecast_table)
        fb_layout.addWidget(self.txt_search, 2)

        fb_layout.addWidget(QLabel("فیلتر وضعیت:"))
        self.cmb_status = QComboBox()
        self.cmb_status.addItem("همه وضعیت‌ها (مرتب‌شده بر اساس اولویت اتمام)", "ALL")
        self.cmb_status.addItem("🚨 فقط نیازمند اقدام فوری (تمام‌شده + بحرانی + کمبود)", "ACTION_NEEDED")
        self.cmb_status.addItem("🚫 اتمام موجودی (موجودی صفر)", "OUT_OF_STOCK")
        self.cmb_status.addItem("🚨 هشدار بحرانی", "CRITICAL")
        self.cmb_status.addItem("⚠️ کمبود موجودی", "WARNING")
        self.cmb_status.addItem("✅ وضعیت عادی", "NORMAL")
        self.cmb_status.addItem("ℹ️ بدون مصرف روزانه (مصرف صفر)", "NO_USAGE")
        self.cmb_status.currentIndexChanged.connect(self.load_forecast_table)
        fb_layout.addWidget(self.cmb_status, 2)

        fb_layout.addWidget(QLabel("دسته‌بندی:"))
        self.cmb_category = QComboBox()
        self.cmb_category.setMinimumWidth(150)
        self.cmb_category.currentIndexChanged.connect(self.load_forecast_table)
        fb_layout.addWidget(self.cmb_category, 1)

        layout.addWidget(filter_bar)

        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels([
            "کد کالا", "نام کالا", "دسته‌بندی", "موجودی فعلی",
            "مصرف روزانه", "روزهای باقی‌مانده", "تاریخ پیش‌بینی اتمام (شمسی)",
            "وضعیت موجودی", "پیشنهاد خرید (شارژ ۳۰ روز)"
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setStyleSheet(TABLE_STYLE)
        layout.addWidget(self.table, 1)

        lbl_formula = QLabel(
            "📐 فرمول محاسبه: تعداد روزهای باقی‌مانده = (موجودی فعلی ÷ مصرف روزانه) گرد شده به بالا  |  "
            "تاریخ پیش‌بینی اتمام = تاریخ امروز سیستم + تعداد روزهای باقی‌مانده"
        )
        lbl_formula.setStyleSheet(f"color: #7f8c8d; font-family: '{FONT_TEXT}'; font-size: 14px;")
        layout.addWidget(lbl_formula)

    def create_stat_card(self, title, value, bg_color, fg_color):
        frame = QFrame()
        frame.setStyleSheet(f"""
            QFrame {{
                background-color: {bg_color};
                border: 1px solid {fg_color};
                border-radius: 10px;
                padding: 10px;
            }}
            QLabel {{ border: none; color: {fg_color}; }}
        """)
        v_layout = QVBoxLayout(frame)
        v_layout.setSpacing(4)

        lbl_title = QLabel(title)
        lbl_title.setFont(QFont(FONT_TITLE, 12, QFont.Bold))
        lbl_title.setStyleSheet(f"font-family: '{FONT_TITLE}'; font-size: 13px;")
        lbl_title.setAlignment(Qt.AlignCenter)

        lbl_val = QLabel(value)
        lbl_val.setFont(QFont(FONT_TEXT, 17, QFont.Bold))
        lbl_val.setStyleSheet(f"font-family: '{FONT_TEXT}'; font-size: 19px; font-weight: bold;")
        lbl_val.setAlignment(Qt.AlignCenter)

        v_layout.addWidget(lbl_title)
        v_layout.addWidget(lbl_val)
        return {"frame": frame, "title": lbl_title, "value": lbl_val}

    def refresh_data(self):
        current_cat = self.cmb_category.currentData()
        self.cmb_category.blockSignals(True)
        self.cmb_category.clear()
        self.cmb_category.addItem("همه دسته‌بندی‌ها", 0)
        for c in self.db.get_categories():
            self.cmb_category.addItem(c["name"], c["id"])
        if current_cat is not None:
            idx = self.cmb_category.findData(current_cat)
            if idx >= 0:
                self.cmb_category.setCurrentIndex(idx)
        self.cmb_category.blockSignals(False)

        self.load_forecast_table()

    def load_forecast_table(self):
        crit_days = int(self.db.get_setting("critical_days", 7))
        warn_days = int(self.db.get_setting("warning_days", 14))

        self.card_crit["title"].setText(f"🚨 هشدار بحرانی (کمتر از {crit_days} روز)")
        self.card_warn["title"].setText(f"⚠️ کمبود موجودی ({crit_days} تا {warn_days} روز)")
        self.card_norm["title"].setText(f"✅ وضعیت عادی (بیش از {warn_days} روز)")

        all_items = self.db.get_items()
        counts = {"OUT_OF_STOCK": 0, "CRITICAL": 0, "WARNING": 0, "NORMAL": 0}
        for it in all_items:
            inf = enrich_item_with_calculations(it, crit_days, warn_days)
            st = inf["status_code"]
            if st in counts:
                counts[st] += 1

        self.card_out["value"].setText(f"{counts['OUT_OF_STOCK']} قلم")
        self.card_crit["value"].setText(f"{counts['CRITICAL']} قلم")
        self.card_warn["value"].setText(f"{counts['WARNING']} قلم")
        self.card_norm["value"].setText(f"{counts['NORMAL']} قلم")

        search_txt = self.txt_search.text().strip()
        cat_id = self.cmb_category.currentData()
        status_filter = self.cmb_status.currentData()

        filtered_raw = self.db.get_items(search_text=search_txt, category_id=cat_id)
        enriched_list = [enrich_item_with_calculations(it, crit_days, warn_days) for it in filtered_raw]

        if status_filter and status_filter != "ALL":
            if status_filter == "ACTION_NEEDED":
                enriched_list = [
                    x for x in enriched_list
                    if x["status_code"] in ("OUT_OF_STOCK", "CRITICAL", "WARNING")
                ]
            else:
                enriched_list = [x for x in enriched_list if x["status_code"] == status_filter]

        enriched_list.sort(key=lambda x: (x["sort_priority"], x["name"]))

        self.table.setRowCount(len(enriched_list))
        for row_idx, info in enumerate(enriched_list):
            unit = info["unit"]
            daily = float(info["daily_usage"])
            stock = float(info["stock"])

            target_30_days = daily * 30.0
            if daily <= 0:
                suggest_order = "-"
            elif stock < target_30_days:
                needed = round(target_30_days - stock, 2)
                suggest_order = f"{needed:g} {unit}"
            else:
                suggest_order = "کافی برای ۳۰ روز"

            row_vals = [
                info["item_code"],
                info["name"],
                info.get("category_name", "-"),
                f"{stock:g} {unit}",
                f"{daily:g} {unit}",
                info["days_display"],
                info["depletion_date"],
                info["status_text"],
                suggest_order
            ]

            for col_idx, val in enumerate(row_vals):
                cell = QTableWidgetItem(str(val))
                cell.setTextAlignment(Qt.AlignCenter)
                if col_idx in (5, 6, 7):
                    cell.setBackground(QBrush(QColor(info["color_bg"])))
                    cell.setForeground(QBrush(QColor(info["color_fg"])))
                    f = cell.font()
                    f.setBold(True)
                    cell.setFont(f)
                self.table.setItem(row_idx, col_idx, cell)

            self.table.setRowHeight(row_idx, 38)


# =====================================================================
# صفحه ۶: گزارش‌ها و خروجی Excel و PDF (ReportsPage)
# =====================================================================
class ReportsPage(QWidget):
    TX_TYPE_NAMES = {
        "IN": "ورود به انبار",
        "OUT": "خروج از انبار",
        "USAGE": "ثبت مصرف روزانه",
        "ADJUST_IN": "سند اصلاحی افزایشی",
        "ADJUST_OUT": "سند اصلاحی کاهشی",
    }

    def __init__(self, db):
        super().__init__()
        self.db = db
        self.current_headers = []
        self.current_rows = []
        self.current_report_title = "گزارش موجودی فعلی انبار"
        self.build_ui()
        self.refresh_data()

    def build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        ctrl_card = QFrame()
        ctrl_card.setStyleSheet(f"""
            QFrame {{
                background-color: #ffffff;
                border: 1px solid #dfe4ea;
                border-radius: 10px;
                padding: 12px;
            }}
            QLabel {{ border: none; font-family: '{FONT_TEXT}'; font-weight: bold; color: #2f3542; font-size: 15px; }}
            QComboBox, QLineEdit {{
                padding: 7px 10px; border: 1px solid #ced6e0; border-radius: 6px;
                background-color: #f8f9fa; font-family: '{FONT_TEXT}'; font-size: 15px;
            }}
        """)
        ctrl_layout = QVBoxLayout(ctrl_card)
        ctrl_layout.setSpacing(10)

        row1 = QHBoxLayout()
        row1.setSpacing(10)

        row1.addWidget(QLabel("📑 انتخاب نوع گزارش:"))
        self.cmb_report_type = QComboBox()
        self.cmb_report_type.setMinimumWidth(290)
        self.cmb_report_type.addItem("۱. گزارش موجودی فعلی کل انبار", "CURRENT_STOCK")
        self.cmb_report_type.addItem("۲. گزارش کالاهای کم‌موجودی و بحرانی", "LOW_STOCK")
        self.cmb_report_type.addItem("۳. گزارش کالاهای تمام‌شده (موجودی صفر)", "OUT_OF_STOCK")
        self.cmb_report_type.addItem("۴. گزارش ورود کالا به انبار (در بازه زمانی)", "TX_IN")
        self.cmb_report_type.addItem("۵. گزارش خروج و مصرف کالا (در بازه زمانی)", "TX_OUT")
        self.cmb_report_type.addItem("۶. گزارش تحلیلی مصرف روزانه، هفتگی و ماهانه", "USAGE_SUMMARY")
        self.cmb_report_type.addItem("۷. گزارش کاردکس (تاریخچه کامل یک کالای مشخص)", "ITEM_HISTORY")
        self.cmb_report_type.currentIndexChanged.connect(self.on_report_type_changed)
        row1.addWidget(self.cmb_report_type, 2)

        self.lbl_item_filter = QLabel("انتخاب کالا:")
        row1.addWidget(self.lbl_item_filter)
        self.cmb_item = QComboBox()
        self.cmb_item.setMinimumWidth(200)
        self.cmb_item.currentIndexChanged.connect(self.generate_preview)
        row1.addWidget(self.cmb_item, 1)

        self.lbl_from = QLabel("از تاریخ شمسی:")
        row1.addWidget(self.lbl_from)
        self.txt_date_from = QLineEdit()
        self.txt_date_from.setPlaceholderText("مثال: 1405/01/01")
        self.txt_date_from.setFixedWidth(120)
        row1.addWidget(self.txt_date_from)

        self.lbl_to = QLabel("تا تاریخ شمسی:")
        row1.addWidget(self.lbl_to)
        self.txt_date_to = QLineEdit()
        self.txt_date_to.setPlaceholderText("مثال: 1405/12/29")
        self.txt_date_to.setFixedWidth(120)
        row1.addWidget(self.txt_date_to)

        btn_preview = QPushButton("🔍 نمایش گزارش")
        btn_preview.setCursor(Qt.PointingHandCursor)
        btn_preview.setStyleSheet(f"""
            QPushButton {{
                background-color: #3867d6; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 8px 16px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #2d55b8; }}
        """)
        btn_preview.clicked.connect(self.generate_preview)
        row1.addWidget(btn_preview)

        ctrl_layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.lbl_report_summary = QLabel("تعداد رکوردهای گزارش: ۰ مورد")
        self.lbl_report_summary.setStyleSheet(f"color: #0a8d87; font-family: '{FONT_TITLE}'; font-size: 14px; font-weight: bold;")
        row2.addWidget(self.lbl_report_summary)

        row2.addStretch()

        btn_excel = QPushButton("📊 خروجی فایل Excel (.xlsx)")
        btn_excel.setCursor(Qt.PointingHandCursor)
        btn_excel.setStyleSheet(f"""
            QPushButton {{
                background-color: #20bf6b; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 9px 18px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #1eb980; }}
        """)
        btn_excel.clicked.connect(self.export_current_to_excel)
        row2.addWidget(btn_excel)

        btn_pdf = QPushButton("📄 خروجی فایل PDF (.pdf)")
        btn_pdf.setCursor(Qt.PointingHandCursor)
        btn_pdf.setStyleSheet(f"""
            QPushButton {{
                background-color: #eb3b5a; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 9px 18px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #d1304d; }}
        """)
        btn_pdf.clicked.connect(self.export_current_to_pdf)
        row2.addWidget(btn_pdf)

        ctrl_layout.addLayout(row2)
        layout.addWidget(ctrl_card)

        self.table = QTableWidget()
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setStyleSheet(TABLE_STYLE)
        layout.addWidget(self.table, 1)

    def refresh_data(self):
        selected_item = self.cmb_item.currentData()
        all_items = self.db.get_items(include_inactive=True)
        self.cmb_item.blockSignals(True)
        self.cmb_item.clear()
        for it in all_items:
            suffix = "" if it["is_active"] else " [بایگانی‌شده]"
            self.cmb_item.addItem(f"{it['name']} ({it['item_code']}){suffix}", it["id"])
        if selected_item:
            idx = self.cmb_item.findData(selected_item)
            if idx >= 0:
                self.cmb_item.setCurrentIndex(idx)
        self.cmb_item.blockSignals(False)

        self.on_report_type_changed()

    def on_report_type_changed(self):
        rep_code = self.cmb_report_type.currentData()
        needs_dates = rep_code in ("TX_IN", "TX_OUT", "ITEM_HISTORY")
        needs_item = rep_code == "ITEM_HISTORY"

        self.lbl_from.setVisible(needs_dates)
        self.txt_date_from.setVisible(needs_dates)
        self.lbl_to.setVisible(needs_dates)
        self.txt_date_to.setVisible(needs_dates)

        self.lbl_item_filter.setVisible(needs_item)
        self.cmb_item.setVisible(needs_item)

        self.generate_preview()

    def generate_preview(self):
        rep_code = self.cmb_report_type.currentData()
        crit_days = int(self.db.get_setting("critical_days", 7))
        warn_days = int(self.db.get_setting("warning_days", 14))

        start_g = None
        end_g = None
        if rep_code in ("TX_IN", "TX_OUT", "ITEM_HISTORY"):
            raw_from = self.txt_date_from.text().strip()
            raw_to = self.txt_date_to.text().strip()
            if raw_from:
                start_g = jalali_to_gregorian_str(raw_from)
                if not start_g:
                    QMessageBox.warning(self, "فرمت تاریخ", "فرمت «از تاریخ» معتبر نیست. مثال صحیح: 1405/01/01")
                    return
            if raw_to:
                end_g = jalali_to_gregorian_str(raw_to)
                if not end_g:
                    QMessageBox.warning(self, "فرمت تاریخ", "فرمت «تا تاریخ» معتبر نیست. مثال صحیح: 1405/12/29")
                    return

        headers = []
        rows = []

        if rep_code == "CURRENT_STOCK":
            self.current_report_title = "گزارش موجودی فعلی کل انبار"
            headers = [
                "ردیف", "کد کالا", "نام کالا", "دسته‌بندی", "موجودی فعلی",
                "واحد", "حداقل مجاز", "مصرف روزانه", "روزهای مانده", "تاریخ اتمام (شمسی)", "وضعیت"
            ]
            items = self.db.get_items(sort_by="name_asc")
            for idx, it in enumerate(items, start=1):
                inf = enrich_item_with_calculations(it, crit_days, warn_days)
                rows.append([
                    idx, inf["item_code"], inf["name"], inf.get("category_name", "-"),
                    f"{inf['stock']:g}", inf["unit"], f"{inf['min_stock']:g}",
                    f"{inf['daily_usage']:g}", inf["days_display"], inf["depletion_date"], inf["status_text"]
                ])

        elif rep_code == "LOW_STOCK":
            self.current_report_title = "گزارش کالاهای کم‌موجودی و بحرانی"
            headers = [
                "ردیف", "کد کالا", "نام کالا", "دسته‌بندی", "موجودی فعلی",
                "حداقل مجاز", "مصرف روزانه", "روزهای مانده", "تاریخ اتمام (شمسی)", "وضعیت"
            ]
            items = self.db.get_items(sort_by="stock_asc")
            idx_counter = 1
            for it in items:
                inf = enrich_item_with_calculations(it, crit_days, warn_days)
                is_low_min = float(inf["stock"]) <= float(inf["min_stock"]) and float(inf["min_stock"]) > 0
                if inf["status_code"] in ("CRITICAL", "WARNING", "OUT_OF_STOCK") or is_low_min:
                    rows.append([
                        idx_counter, inf["item_code"], inf["name"], inf.get("category_name", "-"),
                        f"{inf['stock']:g} {inf['unit']}", f"{inf['min_stock']:g} {inf['unit']}",
                        f"{inf['daily_usage']:g} {inf['unit']}", inf["days_display"],
                        inf["depletion_date"], inf["status_text"]
                    ])
                    idx_counter += 1

        elif rep_code == "OUT_OF_STOCK":
            self.current_report_title = "گزارش کالاهای تمام‌شده (موجودی صفر)"
            headers = ["ردیف", "کد کالا", "نام کالا", "دسته‌بندی", "واحد", "مصرف روزانه", "آخرین به‌روزرسانی"]
            items = [it for it in self.db.get_items() if float(it["stock"]) <= 0]
            for idx, it in enumerate(items, start=1):
                rows.append([
                    idx, it["item_code"], it["name"], it.get("category_name", "-"),
                    it["unit"], f"{it['daily_usage']:g}", iso_to_jalali(it["updated_at"], True)
                ])

        elif rep_code in ("TX_IN", "TX_OUT"):
            tx_filter = "ALL_IN" if rep_code == "TX_IN" else "ALL_OUT"
            self.current_report_title = "گزارش ورود کالا به انبار" if rep_code == "TX_IN" else "گزارش خروج و مصرف کالا"
            headers = [
                "ردیف", "تاریخ و ساعت (شمسی)", "کد کالا", "نام کالا",
                "نوع عملیات", "تعداد", "موجودی قبل", "موجودی بعد", "توضیحات"
            ]
            tx_list = self.db.get_transactions(tx_type=tx_filter, start_date=start_g, end_date=end_g, limit=2000)
            for idx, tx in enumerate(tx_list, start=1):
                u = tx.get("unit", "")
                rows.append([
                    idx, iso_to_jalali(tx["created_at"], True), tx["item_code"], tx["item_name"],
                    self.TX_TYPE_NAMES.get(tx["tx_type"], tx["tx_type"]),
                    f"{tx['quantity']:g} {u}", f"{tx['stock_before']:g} {u}",
                    f"{tx['stock_after']:g} {u}", tx.get("description", "")
                ])

        elif rep_code == "USAGE_SUMMARY":
            self.current_report_title = "گزارش تحلیلی مصرف روزانه، هفتگی و ماهانه"
            headers = [
                "ردیف", "کد کالا", "نام کالا", "دسته‌بندی", "موجودی فعلی",
                "مصرف روزانه برآوردی", "مصرف هفتگی (۷ روز)", "مصرف ماهانه (۳۰ روز)", "میانگین مصرف واقعی (۳۰ روز اخیر)"
            ]
            items = self.db.get_items(sort_by="name_asc")
            for idx, it in enumerate(items, start=1):
                u = it["unit"]
                d_rate = float(it["daily_usage"])
                w_rate = round(d_rate * 7, 2)
                m_rate = round(d_rate * 30, 2)
                actual_avg = self.db.get_actual_daily_usage_avg(it["id"], days=30)
                rows.append([
                    idx, it["item_code"], it["name"], it.get("category_name", "-"),
                    f"{it['stock']:g} {u}", f"{d_rate:g} {u}",
                    f"{w_rate:g} {u}", f"{m_rate:g} {u}", f"{actual_avg:g} {u}"
                ])

        elif rep_code == "ITEM_HISTORY":
            item_id = self.cmb_item.currentData()
            item_obj = self.db.get_item_by_id(item_id) if item_id else None
            item_title = item_obj["name"] if item_obj else "کالا"
            self.current_report_title = f"گزارش تاریخچه و کاردکس کالای «{item_title}»"
            headers = [
                "ردیف", "تاریخ و ساعت (شمسی)", "نوع تراکنش", "تعداد",
                "موجودی قبل", "موجودی بعد", "توضیحات"
            ]
            if item_id:
                tx_list = self.db.get_transactions(item_id=item_id, start_date=start_g, end_date=end_g, limit=2000)
                for idx, tx in enumerate(tx_list, start=1):
                    u = tx.get("unit", "")
                    rows.append([
                        idx, iso_to_jalali(tx["created_at"], True),
                        self.TX_TYPE_NAMES.get(tx["tx_type"], tx["tx_type"]),
                        f"{tx['quantity']:g} {u}", f"{tx['stock_before']:g} {u}",
                        f"{tx['stock_after']:g} {u}", tx.get("description", "")
                    ])

        self.current_headers = headers
        self.current_rows = rows

        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(len(rows))

        for r_idx, r_data in enumerate(rows):
            for c_idx, val in enumerate(r_data):
                cell = QTableWidgetItem(str(val))
                cell.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(r_idx, c_idx, cell)
            self.table.setRowHeight(r_idx, 38)

        self.lbl_report_summary.setText(
            f"📌 عنوان: {self.current_report_title}   |   تعداد رکوردهای یافت‌شده: {len(rows)} مورد"
        )

    def export_current_to_excel(self):
        if not self.current_headers:
            return
        default_name = f"{self.current_report_title.replace(' ', '_')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "ذخیره گزارش در قالب Excel", default_name, "Excel Files (*.xlsx)"
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        wh_name = self.db.get_setting("warehouse_name", "انبار مرکزی")
        ok, msg = export_to_excel(file_path, self.current_report_title, self.current_headers, self.current_rows, wh_name)
        if ok:
            QMessageBox.information(self, "خروجی اکسل موفق", msg)
        else:
            QMessageBox.critical(self, "خطا", msg)

    def export_current_to_pdf(self):
        if not self.current_headers:
            return
        default_name = f"{self.current_report_title.replace(' ', '_')}.pdf"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "ذخیره گزارش در قالب PDF", default_name, "PDF Files (*.pdf)"
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".pdf"):
            file_path += ".pdf"

        wh_name = self.db.get_setting("warehouse_name", "انبار مرکزی")
        ok, msg = export_to_pdf(file_path, self.current_report_title, self.current_headers, self.current_rows, wh_name)
        if ok:
            QMessageBox.information(self, "خروجی PDF موفق", msg)
        else:
            QMessageBox.critical(self, "خطا", msg)


# =====================================================================
# صفحه ۷: پشتیبان‌گیری و بازیابی اطلاعات (BackupRestorePage)
# =====================================================================
class BackupRestorePage(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.backups_list = []
        self.build_ui()
        self.refresh_data()

    def build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        top_card = QFrame()
        top_card.setStyleSheet(f"""
            QFrame {{
                background-color: #ffffff;
                border: 1px solid #dfe4ea;
                border-radius: 10px;
                padding: 15px;
            }}
            QLabel {{ border: none; color: #2f3542; font-family: '{FONT_TEXT}'; }}
        """)
        top_layout = QVBoxLayout(top_card)
        top_layout.setSpacing(12)

        lbl_title = QLabel("💾 مرکز امنیت، پشتیبان‌گیری (Backup) و بازیابی اطلاعات (Restore)")
        lbl_title.setFont(QFont(FONT_TITLE, 13, QFont.Bold))
        lbl_title.setStyleSheet(f"color: #0fb9b1; font-family: '{FONT_TITLE}'; font-size: 15px;")
        top_layout.addWidget(lbl_title)

        self.lbl_paths = QLabel("")
        self.lbl_paths.setStyleSheet(f"color: #57606f; font-family: '{FONT_TEXT}'; font-size: 14px;")
        top_layout.addWidget(self.lbl_paths)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        btn_quick_backup = QPushButton("➕ تهیه نسخه پشتیبان فوری (در پوشه برنامه)")
        btn_quick_backup.setCursor(Qt.PointingHandCursor)
        btn_quick_backup.setStyleSheet(f"""
            QPushButton {{
                background-color: #20bf6b; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 10px 18px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #1eb980; }}
        """)
        btn_quick_backup.clicked.connect(self.create_instant_backup)

        btn_external_backup = QPushButton("📁 ذخیره پشتیبان در مسیر دلخواه (فلش/درایو دیگر)")
        btn_external_backup.setCursor(Qt.PointingHandCursor)
        btn_external_backup.setStyleSheet(f"""
            QPushButton {{
                background-color: #3867d6; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 10px 18px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #2d55b8; }}
        """)
        btn_external_backup.clicked.connect(self.create_external_backup)

        btn_restore_selected = QPushButton("♻️ بازیابی از نسخه انتخاب‌شده در جدول")
        btn_restore_selected.setCursor(Qt.PointingHandCursor)
        btn_restore_selected.setStyleSheet(f"""
            QPushButton {{
                background-color: #f39c12; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 10px 18px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #e67e22; }}
        """)
        btn_restore_selected.clicked.connect(self.restore_from_selected_row)

        btn_restore_file = QPushButton("📂 انتخاب و بازیابی فایل پشتیبان از کامپیوتر")
        btn_restore_file.setCursor(Qt.PointingHandCursor)
        btn_restore_file.setStyleSheet(f"""
            QPushButton {{
                background-color: #eb3b5a; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 10px 18px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #d1304d; }}
        """)
        btn_restore_file.clicked.connect(self.restore_from_external_file)

        btn_row.addWidget(btn_quick_backup)
        btn_row.addWidget(btn_external_backup)
        btn_row.addWidget(btn_restore_selected)
        btn_row.addWidget(btn_restore_file)
        btn_row.addStretch()

        top_layout.addLayout(btn_row)
        layout.addWidget(top_card)

        lbl_list = QLabel("📋 فهرست نسخه‌های پشتیبان ذخیره‌شده در پوشه Backups:")
        lbl_list.setFont(QFont(FONT_TITLE, 12, QFont.Bold))
        lbl_list.setStyleSheet(f"font-family: '{FONT_TITLE}'; font-size: 14px; color: #2f3542;")
        layout.addWidget(lbl_list)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels([
            "نام فایل پشتیبان", "تاریخ و ساعت ایجاد (شمسی)", "حجم فایل", "مسیر کامل فایل"
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setStyleSheet(TABLE_STYLE)
        layout.addWidget(self.table, 1)

        lbl_safe = QLabel(
            "🛡️ تضمین امنیت اطلاعات: قبل از هر بار بازیابی اطلاعات (Restore)، نرم‌افزار به صورت خودکار یک نسخه پشتیبان از وضعیت فعلی تهیه می‌کند تا هیچ داده‌ای به اشتباه از دست نرود."
        )
        lbl_safe.setStyleSheet(f"color: #7f8c8d; font-family: '{FONT_TEXT}'; font-size: 14px;")
        layout.addWidget(lbl_safe)

    def refresh_data(self):
        self.lbl_paths.setText(
            f"📍 مسیر پایگاه داده اصلی: {self.db.db_path}\n"
            f"📁 مسیر پوشه پشتیبان‌ها: {self.db.backup_dir}"
        )
        self.load_backups_list()

    def load_backups_list(self):
        self.backups_list = []
        os.makedirs(self.db.backup_dir, exist_ok=True)

        files = [
            f for f in os.listdir(self.db.backup_dir)
            if f.lower().endswith(".db")
        ]

        items_info = []
        for fname in files:
            fpath = os.path.join(self.db.backup_dir, fname)
            try:
                stat = os.stat(fpath)
                dt = datetime.fromtimestamp(stat.st_mtime)
                jdt = jdatetime.datetime.fromgregorian(datetime=dt).strftime("%Y/%m/%d - %H:%M:%S")
                size_kb = max(1, round(stat.st_size / 1024, 1))
                items_info.append({
                    "name": fname,
                    "jalali_time": jdt,
                    "size_str": f"{size_kb} کیلوبایت",
                    "path": fpath,
                    "mtime": stat.st_mtime
                })
            except Exception:
                continue

        items_info.sort(key=lambda x: x["mtime"], reverse=True)
        self.backups_list = items_info

        self.table.setRowCount(len(items_info))
        for r_idx, b in enumerate(items_info):
            vals = [b["name"], b["jalali_time"], b["size_str"], b["path"]]
            for c_idx, val in enumerate(vals):
                cell = QTableWidgetItem(str(val))
                cell.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(r_idx, c_idx, cell)
            self.table.setRowHeight(r_idx, 38)

    def create_instant_backup(self):
        ok, result = self.db.create_backup()
        if ok:
            QMessageBox.information(
                self, "پشتیبان‌گیری موفق", f"نسخه پشتیبان جدید با موفقیت ساخته شد:\n{result}"
            )
            self.load_backups_list()
        else:
            QMessageBox.critical(self, "خطا", result)

    def create_external_backup(self):
        folder = QFileDialog.getExistingDirectory(
            self, "انتخاب پوشه برای ذخیره فایل پشتیبان (مثلاً فلش مموری)"
        )
        if not folder:
            return
        ok, result = self.db.create_backup(target_folder=folder)
        if ok:
            QMessageBox.information(
                self, "پشتیبان‌گیری موفق", f"نسخه پشتیبان با موفقیت در مسیر انتخابی ذخیره شد:\n{result}"
            )
            self.load_backups_list()
        else:
            QMessageBox.critical(self, "خطا", result)

    def restore_from_selected_row(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.warning(self, "انتخاب نسخه پشتیبان", "لطفاً ابتدا یک نسخه پشتیبان را از جدول انتخاب کنید.")
            return
        row_idx = rows[0].row()
        if 0 <= row_idx < len(self.backups_list):
            target = self.backups_list[row_idx]
            self.confirm_and_restore(target["path"])

    def restore_from_external_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "انتخاب فایل پشتیبان پایگاه داده", "", "SQLite Database (*.db)"
        )
        if file_path:
            self.confirm_and_restore(file_path)

    def confirm_and_restore(self, file_path):
        reply = QMessageBox.question(
            self,
            "تأیید بازیابی اطلاعات (Restore)",
            f"آیا از جایگزینی اطلاعات فعلی انبار با فایل پشتیبان زیر اطمینان دارید؟\n\n{file_path}\n\n"
            "توجه: قبل از بازیابی، یک نسخه پشتیبان خودکار از اطلاعات فعلی شما در پوشه backups ذخیره می‌شود.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            ok, msg = self.db.restore_backup(file_path)
            if ok:
                QMessageBox.information(self, "بازیابی موفق", msg)
                self.load_backups_list()
            else:
                QMessageBox.critical(self, "خطا در بازیابی", msg)


# =====================================================================
# صفحه ۸: تنظیمات نرم‌افزار و مدیریت کالاهای بایگانی‌شده (SettingsPage)
# =====================================================================
class SettingsPage(QWidget):
    def __init__(self, db, on_settings_saved_callback=None):
        super().__init__()
        self.db = db
        self.on_settings_saved_callback = on_settings_saved_callback
        self.archived_items = []
        self.build_ui()
        self.refresh_data()

    def build_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)

        settings_card = QFrame()
        settings_card.setFixedWidth(400)
        settings_card.setStyleSheet(f"""
            QFrame {{
                background-color: #ffffff;
                border: 1px solid #dfe4ea;
                border-radius: 10px;
                padding: 18px;
            }}
            QLabel {{ border: none; font-family: '{FONT_TEXT}'; font-weight: bold; color: #2f3542; font-size: 15px; }}
            QLineEdit, QSpinBox {{
                padding: 8px; border: 1px solid #ced6e0; border-radius: 6px;
                background-color: #f8f9fa; font-family: '{FONT_TEXT}'; font-size: 15px;
            }}
        """)
        s_layout = QVBoxLayout(settings_card)
        s_layout.setSpacing(14)

        lbl_title = QLabel("⚙️ تنظیمات عمومی و پارامترهای هشدار")
        lbl_title.setFont(QFont(FONT_TITLE, 13, QFont.Bold))
        lbl_title.setStyleSheet(f"color: #0fb9b1; font-family: '{FONT_TITLE}'; font-size: 15px;")
        s_layout.addWidget(lbl_title)

        s_layout.addWidget(QLabel("نام انبار / سازمان (نمایش در عنوان و گزارش‌ها):"))
        self.txt_warehouse_name = QLineEdit()
        s_layout.addWidget(self.txt_warehouse_name)

        s_layout.addWidget(QLabel("🚨 مرز هشدار بحرانی (تعداد روز باقی‌مانده کمتر از):"))
        self.spn_critical_days = QSpinBox()
        self.spn_critical_days.setRange(1, 365)
        self.spn_critical_days.setSuffix(" روز")
        s_layout.addWidget(self.spn_critical_days)

        s_layout.addWidget(QLabel("⚠️ مرز هشدار کمبود موجودی (تعداد روز باقی‌مانده تا):"))
        self.spn_warning_days = QSpinBox()
        self.spn_warning_days.setRange(2, 365)
        self.spn_warning_days.setSuffix(" روز")
        s_layout.addWidget(self.spn_warning_days)

        self.chk_auto_backup = QCheckBox("تهیه نسخه پشتیبان خودکار هنگام بستن نرم‌افزار")
        self.chk_auto_backup.setStyleSheet(f"font-family: '{FONT_TEXT}'; font-size: 15px; font-weight: bold; color: #2f3542; padding: 6px 0;")
        s_layout.addWidget(self.chk_auto_backup)

        btn_save = QPushButton("💾 ذخیره تنظیمات")
        btn_save.setCursor(Qt.PointingHandCursor)
        btn_save.setStyleSheet(f"""
            QPushButton {{
                background-color: #20bf6b; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 16px;
                padding: 11px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #1eb980; }}
        """)
        btn_save.clicked.connect(self.save_settings)
        s_layout.addWidget(btn_save)

        s_layout.addStretch()

        archive_layout = QVBoxLayout()
        archive_layout.setSpacing(12)

        top_bar = QHBoxLayout()
        lbl_arch = QLabel("🗄️ فهرست کالاهای بایگانی‌شده (حذف منطقی شده)")
        lbl_arch.setFont(QFont(FONT_TITLE, 12, QFont.Bold))
        lbl_arch.setStyleSheet(f"font-family: '{FONT_TITLE}'; font-size: 14px; color: #2f3542;")

        btn_restore_item = QPushButton("♻️ بازگردانی کالای انتخاب‌شده به انبار فعال")
        btn_restore_item.setCursor(Qt.PointingHandCursor)
        btn_restore_item.setStyleSheet(f"""
            QPushButton {{
                background-color: #3867d6; color: white; font-weight: bold;
                font-family: '{FONT_TEXT}'; font-size: 15px;
                padding: 8px 16px; border-radius: 6px; border: none;
            }}
            QPushButton:hover {{ background-color: #2d55b8; }}
        """)
        btn_restore_item.clicked.connect(self.reactivate_selected_item)

        top_bar.addWidget(lbl_arch)
        top_bar.addStretch()
        top_bar.addWidget(btn_restore_item)
        archive_layout.addLayout(top_bar)

        self.table_archived = QTableWidget()
        self.table_archived.setColumnCount(5)
        self.table_archived.setHorizontalHeaderLabels([
            "شناسه", "کد کالا", "نام کالا", "دسته‌بندی", "موجودی زمان بایگانی"
        ])
        self.table_archived.setColumnHidden(0, True)
        self.table_archived.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_archived.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table_archived.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_archived.setAlternatingRowColors(True)
        self.table_archived.verticalHeader().setVisible(False)
        self.table_archived.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_archived.setStyleSheet(TABLE_STYLE)
        archive_layout.addWidget(self.table_archived, 1)

        lbl_arch_note = QLabel(
            "💡 توضیح: کالاهایی که در بخش «مدیریت کالاها» حذف می‌شوند، برای حفظ تاریخچه تراکنش‌ها در اینجا بایگانی می‌شوند. "
            "هر زمان بخواهید می‌توانید آن‌ها را دوباره فعال کنید."
        )
        lbl_arch_note.setStyleSheet(f"color: #7f8c8d; font-family: '{FONT_TEXT}'; font-size: 14px;")
        archive_layout.addWidget(lbl_arch_note)

        layout.addWidget(settings_card)
        layout.addLayout(archive_layout, 1)

    def refresh_data(self):
        wh_name = self.db.get_setting("warehouse_name", "انبار مرکزی")
        crit = int(self.db.get_setting("critical_days", 7))
        warn = int(self.db.get_setting("warning_days", 14))
        auto_bkp = self.db.get_setting("auto_backup", "1") == "1"

        self.txt_warehouse_name.setText(wh_name)
        self.spn_critical_days.setValue(crit)
        self.spn_warning_days.setValue(warn)
        self.chk_auto_backup.setChecked(auto_bkp)

        self.load_archived_items()

    def load_archived_items(self):
        all_items = self.db.get_items(include_inactive=True)
        self.archived_items = [it for it in all_items if int(it.get("is_active", 1)) == 0]

        self.table_archived.setRowCount(len(self.archived_items))
        for r_idx, it in enumerate(self.archived_items):
            vals = [
                str(it["id"]),
                it["item_code"],
                it["name"],
                it.get("category_name", "-"),
                f"{it['stock']:g} {it['unit']}"
            ]
            for c_idx, val in enumerate(vals):
                cell = QTableWidgetItem(str(val))
                cell.setTextAlignment(Qt.AlignCenter)
                self.table_archived.setItem(r_idx, c_idx, cell)
            self.table_archived.setRowHeight(r_idx, 38)

    def save_settings(self):
        wh_name = self.txt_warehouse_name.text().strip() or "انبار مرکزی"
        crit = self.spn_critical_days.value()
        warn = self.spn_warning_days.value()

        if warn <= crit:
            QMessageBox.warning(
                self, "خطا در اعداد هشدار",
                "مرز هشدار کمبود موجودی باید بزرگ‌تر از مرز هشدار بحرانی باشد (مثلاً بحرانی ۷ روز و کمبود ۱۴ روز)."
            )
            return

        self.db.set_setting("warehouse_name", wh_name)
        self.db.set_setting("critical_days", str(crit))
        self.db.set_setting("warning_days", str(warn))
        self.db.set_setting("auto_backup", "1" if self.chk_auto_backup.isChecked() else "0")

        if self.on_settings_saved_callback:
            self.on_settings_saved_callback()

        QMessageBox.information(self, "ذخیره موفق", "تنظیمات نرم‌افزار با موفقیت ذخیره و اعمال شد.")

    def reactivate_selected_item(self):
        rows = self.table_archived.selectionModel().selectedRows()
        if not rows:
            QMessageBox.warning(self, "انتخاب کالا", "لطفاً ابتدا یک کالای بایگانی‌شده را از جدول انتخاب کنید.")
            return
        item_id = int(self.table_archived.item(rows[0].row(), 0).text())
        with self.db.get_connection() as conn:
            conn.execute(
                "UPDATE items SET is_active = 1, updated_at = ? WHERE id = ?",
                (self.db.now_iso(), item_id)
            )
            conn.commit()
        QMessageBox.information(self, "فعال‌سازی مجدد", "کالای انتخاب‌شده با موفقیت به فهرست فعال انبار بازگردانده شد.")
        self.load_archived_items()


def create_all_pages(db, on_settings_saved_callback=None):
    pages = {
        "dashboard": DashboardPage(db),
        "items": ItemsManagementPage(db),
        "categories": CategoriesPage(db),
        "transactions": TransactionsPage(db),
        "usage_report": UsageReportPage(db),
        "depletion": DepletionForecastPage(db),
        "reports": ReportsPage(db),
        "backup": BackupRestorePage(db),
        "settings": SettingsPage(db, on_settings_saved_callback),
    }
    return pages