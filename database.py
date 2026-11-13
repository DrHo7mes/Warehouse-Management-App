import sqlite3
import os
import sys
import shutil
from datetime import datetime
import jdatetime


def get_base_dir():
    """تعیین مسیر دقیق پوشه برنامه (سازگار با حالت کد و حالت فایل EXE)"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class DatabaseManager:
    def __init__(self, db_name="inventory.db"):
        self.base_dir = get_base_dir()
        self.db_path = os.path.join(self.base_dir, db_name)
        self.backup_dir = os.path.join(self.base_dir, "backups")
        os.makedirs(self.backup_dir, exist_ok=True)
        self.init_database()

    def get_connection(self):
        """ایجاد ارتباط امن با پایگاه داده SQLite"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON;")
        return conn

    def now_iso(self):
        """تاریخ و ساعت فعلی با فرمت استاندارد برای ذخیره داخلی"""
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def init_database(self):
        """ساخت جدول‌ها و اطلاعات اولیه در اولین اجرای برنامه"""
        with self.get_connection() as conn:
            cursor = conn.cursor()

            # ۱. جدول دسته‌بندی‌ها
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS categories (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    description TEXT DEFAULT ''
                )
            """)

            # ۲. جدول کالاها (همراه با ستون is_active برای حذف منطقی)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    item_code TEXT NOT NULL UNIQUE,
                    name TEXT NOT NULL,
                    category_id INTEGER NOT NULL,
                    stock REAL NOT NULL DEFAULT 0 CHECK(stock >= 0),
                    unit TEXT NOT NULL,
                    min_stock REAL NOT NULL DEFAULT 0,
                    daily_usage REAL NOT NULL DEFAULT 0 CHECK(daily_usage >= 0),
                    updated_at TEXT NOT NULL,
                    description TEXT DEFAULT '',
                    is_active INTEGER NOT NULL DEFAULT 1,
                    FOREIGN KEY (category_id) REFERENCES categories(id)
                )
            """)

            # ۳. جدول تراکنش‌ها (ورود، خروج، مصرف واقعی، اصلاحیه)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS transactions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    item_id INTEGER NOT NULL,
                    tx_type TEXT NOT NULL,
                    quantity REAL NOT NULL CHECK(quantity > 0),
                    stock_before REAL NOT NULL,
                    stock_after REAL NOT NULL,
                    created_at TEXT NOT NULL,
                    description TEXT DEFAULT '',
                    FOREIGN KEY (item_id) REFERENCES items(id)
                )
            """)

            # ۴. جدول تنظیمات
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
            """)

            # افزودن دسته‌بندی‌های پیش‌فرض در صورت خالی بودن جدول
            cursor.execute("SELECT COUNT(*) as cnt FROM categories")
            if cursor.fetchone()["cnt"] == 0:
                default_categories = [
                    ("تجهیزات اداری", "ملزومات و دستگاه‌های اداری"),
                    ("لوازم مصرفی", "اقلام با مصرف روزانه و مستمر"),
                    ("ابزارآلات", "ابزارهای دستی و برقی"),
                    ("تجهیزات ایمنی", "لوازم حفاظت فردی و ایمنی (PPE)"),
                    ("مواد شوینده", "مواد بهداشتی و نظافتی"),
                    ("تجهیزات فنی", "قطعات یدکی و تجهیزات تخصصی"),
                    ("سایر موارد", "دسته‌بندی عمومی برای سایر اقلام")
                ]
                cursor.executemany(
                    "INSERT INTO categories (name, description) VALUES (?, ?)",
                    default_categories
                )

            # تنظیمات پیش‌فرض
            default_settings = [
                ("warehouse_name", "انبار مرکزی"),
                ("critical_days", "7"),
                ("warning_days", "14"),
                ("auto_backup", "1")
            ]
            cursor.executemany(
                "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)",
                default_settings
            )

            conn.commit()

    # ==========================================
    # بخش مدیریت دسته‌بندی‌ها
    # ==========================================
    def get_categories(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT c.id, c.name, c.description,
                       COUNT(CASE WHEN i.is_active = 1 THEN 1 END) as active_items_count,
                       COUNT(i.id) as total_items_count
                FROM categories c
                LEFT JOIN items i ON c.id = i.category_id
                GROUP BY c.id
                ORDER BY c.id ASC
            """)
            return [dict(row) for row in cursor.fetchall()]

    def add_category(self, name, description=""):
        name = name.strip()
        if not name:
            return False, "نام دسته‌بندی نمی‌تواند خالی باشد."
        try:
            with self.get_connection() as conn:
                conn.execute(
                    "INSERT INTO categories (name, description) VALUES (?, ?)",
                    (name, description.strip())
                )
                conn.commit()
            return True, "دسته‌بندی جدید با موفقیت ثبت شد."
        except sqlite3.IntegrityError:
            return False, "این نام دسته‌بندی قبلاً ثبت شده است."

    def update_category(self, category_id, name, description=""):
        name = name.strip()
        if not name:
            return False, "نام دسته‌بندی نمی‌تواند خالی باشد."
        try:
            with self.get_connection() as conn:
                conn.execute(
                    "UPDATE categories SET name = ?, description = ? WHERE id = ?",
                    (name, description.strip(), category_id)
                )
                conn.commit()
            return True, "اطلاعات دسته‌بندی با موفقیت ویرایش شد."
        except sqlite3.IntegrityError:
            return False, "دسته‌بندی دیگری با این نام وجود دارد."

    def delete_category(self, category_id, replacement_category_id=None):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) as cnt FROM categories")
            if cursor.fetchone()["cnt"] <= 1:
                return False, "حداقل یک دسته‌بندی باید در سیستم باقی بماند."

            cursor.execute("SELECT COUNT(*) as cnt FROM items WHERE category_id = ?", (category_id,))
            item_count = cursor.fetchone()["cnt"]

            if item_count > 0:
                if not replacement_category_id or int(replacement_category_id) == int(category_id):
                    return False, f"در این دسته {item_count} کالا ثبت شده است. لطفاً ابتدا دسته جایگزین را مشخص کنید."
                cursor.execute(
                    "UPDATE items SET category_id = ?, updated_at = ? WHERE category_id = ?",
                    (replacement_category_id, self.now_iso(), category_id)
                )

            cursor.execute("DELETE FROM categories WHERE id = ?", (category_id,))
            conn.commit()
            return True, "دسته‌بندی با موفقیت حذف شد."

    # ==========================================
    # بخش مدیریت کالاها
    # ==========================================
    def generate_next_item_code(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) as cnt FROM items")
            next_num = cursor.fetchone()["cnt"] + 1001
            while True:
                code = f"ITM-{next_num}"
                cursor.execute("SELECT 1 FROM items WHERE item_code = ?", (code,))
                if not cursor.fetchone():
                    return code
                next_num += 1

    def add_item(self, item_code, name, category_id, stock, unit, min_stock, daily_usage, description=""):
        item_code = item_code.strip()
        name = name.strip()
        unit = unit.strip()

        if not item_code or not name or not unit:
            return False, "شناسه، نام کالا و واحد شمارش اجباری هستند."
        if stock < 0 or min_stock < 0 or daily_usage < 0:
            return False, "مقادیر موجودی و مصرف نمی‌توانند منفی باشند."

        now = self.now_iso()
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO items (item_code, name, category_id, stock, unit, min_stock, daily_usage, updated_at, description, is_active)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
                """, (item_code, name, category_id, stock, unit, min_stock, daily_usage, now, description.strip()))

                item_id = cursor.lastrowid

                # اگر موجودی اولیه بیشتر از صفر بود، یک تراکنش موجودی اولیه ثبت شود
                if stock > 0:
                    cursor.execute("""
                        INSERT INTO transactions (item_id, tx_type, quantity, stock_before, stock_after, created_at, description)
                        VALUES (?, 'IN', ?, 0, ?, ?, 'ثبت موجودی اولیه هنگام تعریف کالا')
                    """, (item_id, stock, stock, now))

                conn.commit()
            return True, "کالای جدید با موفقیت در انبار ثبت شد."
        except sqlite3.IntegrityError:
            return False, "شناسه اختصاصی (کد کالا) تکراری است. لطفاً کد دیگری وارد کنید."

    def update_item(self, item_id, item_code, name, category_id, unit, min_stock, daily_usage, description=""):
        item_code = item_code.strip()
        name = name.strip()
        unit = unit.strip()
        if not item_code or not name or not unit:
            return False, "شناسه، نام کالا و واحد شمارش اجباری هستند."
        if min_stock < 0 or daily_usage < 0:
            return False, "حداقل موجودی و مصرف روزانه نمی‌توانند منفی باشند."

        try:
            with self.get_connection() as conn:
                conn.execute("""
                    UPDATE items
                    SET item_code = ?, name = ?, category_id = ?, unit = ?,
                        min_stock = ?, daily_usage = ?, description = ?, updated_at = ?
                    WHERE id = ?
                """, (item_code, name, category_id, unit, min_stock, daily_usage, description.strip(), self.now_iso(), item_id))
                conn.commit()
            return True, "اطلاعات کالا با موفقیت به‌روزرسانی شد."
        except sqlite3.IntegrityError:
            return False, "این شناسه کالا قبلاً برای قلم دیگری ثبت شده است."

    def soft_delete_item(self, item_id):
        """حذف منطقی کالا: غیرفعال شدن بدون پاک شدن تاریخچه تراکنش‌ها"""
        with self.get_connection() as conn:
            conn.execute(
                "UPDATE items SET is_active = 0, updated_at = ? WHERE id = ?",
                (self.now_iso(), item_id)
            )
            conn.commit()
        return True, "کالا با موفقیت حذف (بایگانی) شد و سوابق تاریخی آن محفوظ ماند."

    def get_items(self, search_text="", category_id=None, sort_by="name_asc", include_inactive=False):
        query = """
            SELECT i.*, c.name as category_name
            FROM items i
            LEFT JOIN categories c ON i.category_id = c.id
            WHERE 1=1
        """
        params = []

        if not include_inactive:
            query += " AND i.is_active = 1"

        if search_text and search_text.strip():
            q = f"%{search_text.strip()}%"
            query += " AND (i.name LIKE ? OR i.item_code LIKE ? OR i.description LIKE ?)"
            params.extend([q, q, q])

        if category_id and int(category_id) > 0:
            query += " AND i.category_id = ?"
            params.append(int(category_id))

        sort_map = {
            "name_asc": "i.name ASC",
            "name_desc": "i.name DESC",
            "stock_asc": "i.stock ASC",
            "stock_desc": "i.stock DESC",
            "updated_desc": "i.updated_at DESC",
            "code_asc": "i.item_code ASC"
        }
        order_clause = sort_map.get(sort_by, "i.name ASC")
        query += f" ORDER BY {order_clause}"

        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_item_by_id(self, item_id):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT i.*, c.name as category_name
                FROM items i
                LEFT JOIN categories c ON i.category_id = c.id
                WHERE i.id = ?
            """, (item_id,))
            row = cursor.fetchone()
            return dict(row) if row else None

    # ==========================================
    # بخش ثبت ورود و خروج و تراکنش‌های انبار
    # ==========================================
    def record_transaction(self, item_id, tx_type, quantity, description="", custom_datetime=None):
        """
        tx_type:
          'IN'          -> ورود به انبار (افزایش موجودی)
          'OUT'         -> خروج از انبار (کاهش موجودی)
          'USAGE'       -> ثبت مصرف واقعی روزانه (کاهش موجودی)
          'ADJUST_IN'   -> اصلاحیه افزایشی (افزایش موجودی)
          'ADJUST_OUT'  -> اصلاحیه کاهشی (کاهش موجودی)
        """
        try:
            quantity = float(quantity)
        except (ValueError, TypeError):
            return False, "تعداد وارد شده معتبر نیست."

        if quantity <= 0:
            return False, "تعداد تراکنش باید بزرگ‌تر از صفر باشد."

        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT stock, name, unit FROM items WHERE id = ?", (item_id,))
            item = cursor.fetchone()
            if not item:
                return False, "کالای مورد نظر در پایگاه داده یافت نشد."

            stock_before = float(item["stock"])

            if tx_type in ("IN", "ADJUST_IN"):
                stock_after = round(stock_before + quantity, 4)
            elif tx_type in ("OUT", "USAGE", "ADJUST_OUT"):
                if stock_before < quantity:
                    return False, f"موجودی کافی نیست! موجودی فعلی «{item['name']}» برابر با {stock_before:g} {item['unit']} است."
                stock_after = round(stock_before - quantity, 4)
            else:
                return False, "نوع تراکنش نامعتبر است."

            tx_time = custom_datetime if custom_datetime else self.now_iso()

            cursor.execute("""
                INSERT INTO transactions (item_id, tx_type, quantity, stock_before, stock_after, created_at, description)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (item_id, tx_type, quantity, stock_before, stock_after, tx_time, description.strip()))

            cursor.execute("""
                UPDATE items SET stock = ?, updated_at = ? WHERE id = ?
            """, (stock_after, self.now_iso(), item_id))

            conn.commit()
            return True, f"تراکنش با موفقیت ثبت شد. موجودی جدید: {stock_after:g} {item['unit']}"

    def get_transactions(self, item_id=None, tx_type=None, start_date=None, end_date=None, limit=500):
        query = """
            SELECT t.*, i.item_code, i.name as item_name, i.unit, i.is_active, c.name as category_name
            FROM transactions t
            JOIN items i ON t.item_id = i.id
            LEFT JOIN categories c ON i.category_id = c.id
            WHERE 1=1
        """
        params = []

        if item_id and int(item_id) > 0:
            query += " AND t.item_id = ?"
            params.append(int(item_id))

        if tx_type and tx_type != "ALL":
            if tx_type == "ALL_IN":
                query += " AND t.tx_type IN ('IN', 'ADJUST_IN')"
            elif tx_type == "ALL_OUT":
                query += " AND t.tx_type IN ('OUT', 'USAGE', 'ADJUST_OUT')"
            else:
                query += " AND t.tx_type = ?"
                params.append(tx_type)

        if start_date:
            query += " AND substr(t.created_at, 1, 10) >= ?"
            params.append(start_date)

        if end_date:
            query += " AND substr(t.created_at, 1, 10) <= ?"
            params.append(end_date)

        query += " ORDER BY t.created_at DESC, t.id DESC"
        if limit:
            query += f" LIMIT {int(limit)}"

        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_actual_daily_usage_avg(self, item_id, days=30):
        """محاسبه میانگین مصرف واقعی کالا بر اساس خروجی‌های ثبت شده در بازه زمانی"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT COALESCE(SUM(quantity), 0) as total_out
                FROM transactions
                WHERE item_id = ?
                  AND tx_type IN ('OUT', 'USAGE')
                  AND date(created_at) >= date('now', 'localtime', ?)
            """, (item_id, f"-{int(days)} days"))
            total_out = cursor.fetchone()["total_out"]
            return round(total_out / max(1, int(days)), 2)

    # ==========================================
    # بخش تنظیمات و پشتیبان‌گیری (Backup & Restore)
    # ==========================================
    def get_setting(self, key, default=""):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT value FROM settings WHERE key = ?", (key,))
            row = cursor.fetchone()
            return row["value"] if row else default

    def set_setting(self, key, value):
        with self.get_connection() as conn:
            conn.execute(
                "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
                (key, str(value))
            )
            conn.commit()

    def create_backup(self, target_folder=None):
        """تهیه نسخه پشتیبان با نام حاوی تاریخ و ساعت شمسی"""
        folder = target_folder if target_folder else self.backup_dir
        os.makedirs(folder, exist_ok=True)
        j_now = jdatetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        backup_filename = f"Inventory_Backup_{j_now}.db"
        backup_path = os.path.join(folder, backup_filename)

        try:
            with self.get_connection() as src_conn:
                with sqlite3.connect(backup_path) as dst_conn:
                    src_conn.backup(dst_conn)
            return True, backup_path
        except Exception as e:
            return False, f"خطا در تهیه نسخه پشتیبان: {str(e)}"

    def restore_backup(self, backup_file_path):
        """بررسی سلامت فایل پشتیبان و جایگزینی امن آن"""
        if not os.path.exists(backup_file_path):
            return False, "فایل پشتیبان انتخاب شده یافت نشد."

        try:
            # بررسی اینکه فایل واقعاً پایگاه داده معتبر انبار ماست
            test_conn = sqlite3.connect(backup_file_path)
            cursor = test_conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='items'")
            if not cursor.fetchone():
                test_conn.close()
                return False, "فایل انتخاب شده یک پایگاه داده معتبر انبارداری نیست."
            test_conn.close()

            # گرفتن یک پشتیبان خودکار از وضعیت فعلی قبل از بازگردانی
            self.create_backup()

            # جایگزینی فایل
            shutil.copy2(backup_file_path, self.db_path)
            return True, "اطلاعات انبار با موفقیت از نسخه پشتیبان بازیابی شد."
        except Exception as e:
            return False, f"خطا در بازیابی اطلاعات: {str(e)}"


# بخش آزمایش مستقیم فایل در مرحله چهارم
if __name__ == "__main__":
    db = DatabaseManager()
    cats = db.get_categories()
    print("=" * 55)
    print("✅ پایگاه داده SQLite با موفقیت ساخته و راه‌اندازی شد!")
    print(f"📁 مسیر فایل ذخیره شده: {db.db_path}")
    print(f"📦 تعداد دسته‌بندی‌های پیش‌فرض آماده: {len(cats)} دسته")
    for c in cats:
        print(f"   - {c['id']}: {c['name']} ({c['description']})")
    print("=" * 55)