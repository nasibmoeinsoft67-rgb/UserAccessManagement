# -*- coding: utf-8 -*-
import sys
import pyodbc
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QTableWidget, QTableWidgetItem,
    QMessageBox, QCheckBox, QLabel, QInputDialog
)
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QMessageBox
# ----------------------------- تنظیمات پیش‌فرض اتصال -----------------------------
# (طبق درخواستی که گفتی این مقادیر به‌صورت پیش‌فرض استفاده شوند)
SERVER = r".\Moein2012"
DATABASE = "Moein1"
USERNAME = "Sa"
PASSWORD = "arta0@"

# ----------------------------- پنجره‌ی اتصال -----------------------------
class ConnectWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("اتصال به دیتابیس Moein")
        self.setGeometry(400, 250, 350, 120)

        # نمایش پیام اولیه
        layout = QVBoxLayout()
        self.label = QLabel("در حال تلاش برای اتصال به دیتابیس ...")
        layout.addWidget(self.label)
        self.setLayout(layout)

        # ⚠️ نکته مهم:
        # به جای اینکه مستقیماً connect_to_db را صدا بزنیم، با تاخیر 100 میلی‌ثانیه صدا می‌زنیم
        QTimer.singleShot(100, self.connect_to_db)

    def connect_to_db(self):
        try:
            conn_str = (
                f"DRIVER={{SQL Server}};"
                f"SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}"
            )
            self.conn = pyodbc.connect(conn_str, autocommit=False)
            cursor = self.conn.cursor()

            # بررسی ستون IsActive
            cursor.execute("""
                SELECT COLUMN_NAME
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'UserAccess' AND COLUMN_NAME = 'IsActive'
            """)
            exists = cursor.fetchone()
            if not exists:
                cursor.execute("ALTER TABLE dbo.UserAccess ADD IsActive BIT DEFAULT 1;")
                self.conn.commit()

            QMessageBox.information(self, "نتیجه اتصال", "اتصال به دیتابیس موفق بود ✅")

            # ✅ بستن پنجره و رفتن به مرحله بعد
            self.close()
            self.main_window = MainWindow(self.conn)
            self.main_window.show()

        except Exception as e:
            QMessageBox.critical(self, "خطا در اتصال", f"اتصال به دیتابیس ناموفق بود ❌\n{str(e)}")
            self.close()

# ----------------------------- پنجره‌ی اصلی -----------------------------
class MainWindow(QWidget):
    def __init__(self, connection):
        super().__init__()
        # نگهداری اتصال پایگاه‌داده
        self.conn = connection

        self.setWindowTitle("مدیریت سطح دسترسی کاربران")
        self.setGeometry(250, 80, 900, 600)

        # چیدمان کلی عمودی
        layout = QVBoxLayout()

        # دکمه‌ها برای نمایش حالت‌های مختلف
        self.btn_show_allowed = QPushButton("✅ الف: نمایش فرم‌هایی که کاربر دسترسی ندارد")
        self.btn_show_denied = QPushButton("✅ ب: نمایش فرم‌هایی که کاربر دسترسی دارد")
        self.btn_show_all = QPushButton("نمایش همه فرم‌ها (ویرایش وضعیت دسترسی)")

        # افزودن دکمه‌ها به چیدمان
        layout.addWidget(self.btn_show_allowed)
        layout.addWidget(self.btn_show_denied)
        layout.addWidget(self.btn_show_all)

        # جدول برای نمایش اطلاعات
        self.table = QTableWidget()
        layout.addWidget(self.table)

        self.setLayout(layout)

        # اتصال رویداد کلیک دکمه‌ها به متدها
        self.btn_show_allowed.clicked.connect(self.show_allowed_forms)
        self.btn_show_denied.clicked.connect(self.show_denied_forms)
        self.btn_show_all.clicked.connect(self.show_all_forms)

    # ------------------ نمایش فرم‌هایی که کاربر دسترسی دارد ------------------
    def show_allowed_forms(self):
        """
        توجه: این کوئری اکنون ستون AccessStatus و FormButtonId را برمی‌گرداند
        تا load_data بتواند checkbox و عملیات بروز‌رسانی را انجام دهد.
        """
        user_id, ok = self.ask_user_id()
        if not ok:
            return

        query = f"""
            SELECT 
                ua.UserId,
                f.Name AS FormName,
                fb.Name AS ButtonName,
                ISNULL(ua.IsActive, 0) AS AccessStatus,
                fb.ID AS FormButtonId
            FROM dbo.UserAccess ua
            JOIN dbo.FormButtons fb ON ua.FormButtonsId = fb.ID
            JOIN dbo.Forms f ON fb.IDForm = f.ID
            WHERE ua.UserId = {user_id}
            ORDER BY f.Name;
        """
        # editable=True چون می‌خواهیم امکان فعال/غیرفعال شدن را داشته باشیم
        self.load_data(query, editable=True, user_id=user_id)

    # ------------------ نمایش فرم‌هایی که کاربر دسترسی ندارد ------------------
    def show_denied_forms(self):
        user_id, ok = self.ask_user_id()
        if not ok:
            return

        query = f"""
            SELECT 
                f.ID AS FormId,
                f.Name AS FormName
            FROM dbo.Forms f
            WHERE f.ID NOT IN (
                SELECT fb.IDForm 
                FROM dbo.FormButtons fb
                JOIN dbo.UserAccess ua ON fb.ID = ua.FormButtonsId
                WHERE ua.UserId = {user_id}
            );
        """
        # editable=False چون این حالت فقط نمایش است
        self.load_data(query, editable=False, user_id=user_id)

    # ------------------ نمایش همه فرم‌ها با وضعیت و قابلیت ویرایش ------------------
    def show_all_forms(self):
        user_id, ok = self.ask_user_id()
        if not ok:
            return

        query = f"""
            SELECT 
                f.Name AS FormName,
                fb.Name AS ButtonName,
                ISNULL(ua.IsActive, 0) AS AccessStatus,
                fb.ID AS FormButtonId
            FROM dbo.Forms f
            JOIN dbo.FormButtons fb ON fb.IDForm = f.ID
            LEFT JOIN dbo.UserAccess ua 
                ON ua.FormButtonsId = fb.ID 
                AND ua.UserId = {user_id}
            ORDER BY f.MenuOrder, fb.ButtonOrder;
        """
        self.load_data(query, editable=True, user_id=user_id)

    # ------------------ متد بارگذاری دیتا در جدول ------------------
    def load_data(self, query, editable=False, user_id=None):
        """
        اجرای کوئری، نمایش نتایج در QTableWidget و اضافه کردن checkbox برای ویرایش.
        این متد سعی می‌کند انعطاف‌پذیر باشد و اگر ستون‌های فنی (مثل FormButtonId) وجود
        داشته باشند آنها را مخفی کند و از اندیس ستون‌ها استفاده نماید.
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute(query)
            # نام ستون‌ها را از cursor.description می‌گیریم
            columns = [desc[0] for desc in cursor.description]
            rows = cursor.fetchall()

            # تعداد ستون‌ها (اضافه یک ستون برای checkbox در صورت editable)
            self.table.setColumnCount(len(columns) + (1 if editable else 0))
            self.table.setRowCount(len(rows))

            # تنظیم هدرها (اسم‌های ستون + "فعال؟" اگر ویرایش‌پذیر باشد)
            headers = columns + (["فعال؟"] if editable else [])
            self.table.setHorizontalHeaderLabels(headers)

            # پیدا کردن اندیس ستون‌های مورد نیاز (برای حالت editable)
            access_idx = None
            id_idx = None
            if editable:
                # سعی می‌کنیم نام‌های مختلف را بررسی کنیم تا مقاوم باشیم
                if "AccessStatus" in columns:
                    access_idx = columns.index("AccessStatus")
                elif "IsActive" in columns:
                    access_idx = columns.index("IsActive")

                # شناسه‌ی FormButton
                if "FormButtonId" in columns:
                    id_idx = columns.index("FormButtonId")
                elif "FormButtonsId" in columns:
                    id_idx = columns.index("FormButtonsId")
                # در صورت نبودن id_idx، عملیات بروز‌رسانی نمی‌تواند انجام شود.

            # پر کردن جدول با داده‌ها
            for i, row in enumerate(rows):
                # هر ستون را در جدول می‌گذاریم (خواندن با اندیس بهتر است از row.<name>)
                for j, val in enumerate(row):
                    text = "" if val is None else str(val)
                    item = QTableWidgetItem(text)
                    # جلوگیری از ویرایش مستقیم سلول‌ها (فقط از طریق checkbox تغییر انجام شود)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.table.setItem(i, j, item)

                # اگر editable باشیم، یک checkbox در ستون انتهایی قرار می‌دهیم
                if editable:
                    # مقدار checked را از access_idx می‌گیریم (در صورت وجود)
                    checked = False
                    if access_idx is not None:
                        checked = bool(row[access_idx])
                    # شناسه‌ی فرم-دکمه (برای update) را می‌گیریم
                    fid = None
                    if id_idx is not None:
                        fid = row[id_idx]

                    # ساخت checkbox و تنظیم وضعیت اولیه
                    chk = QCheckBox()
                    chk.setChecked(checked)

                    # اگر شناسه موجود باشد، اتصال سیگنال به تابع update_access
                    # از lambda با پارامتر پیش‌فرض استفاده می‌کنیم تا مقدار fid هر ردیف ثابت بماند
                    if fid is not None:
                        chk.stateChanged.connect(lambda state, fid=fid, uid=user_id: self.update_access(uid, fid, state))
                    else:
                        # اگر شناسه وجود نداشت، فقط چک‌باکس غیرفعال باشد (در عمل نباید رخ دهد)
                        chk.setEnabled(False)

                    # قرار دادن checkbox در سلول آخر (index = len(columns))
                    self.table.setCellWidget(i, len(columns), chk)

            # پس از پر کردن جدول، ستون‌های فنی را مخفی می‌کنیم تا کاربر نبیند
            if editable:
                for col_name in ("FormButtonId", "FormButtonsId", "AccessStatus", "IsActive"):
                    if col_name in columns:
                        idx = columns.index(col_name)
                        # مخفی کردن ستون فنی
                        self.table.setColumnHidden(idx, True)

        except Exception as e:
            QMessageBox.critical(self, "خطا در اجرا", f"خطا هنگام بارگذاری داده:\n{str(e)}")

    # ------------------ به‌روزرسانی وضعیت دسترسی ------------------
    def update_access(self, user_id, form_button_id, state):
        """
        وقتی چک‌باکس تغییر کرد، این متد اجرا می‌شود.
        اگر ردیف مربوط به UserId + FormButtonsId وجود نداشت، آن را INSERT می‌کند،
        در غیر این صورت مقدار IsActive را UPDATE می‌کند.
        """
        try:
            cursor = self.conn.cursor()
            # مقدار باینری (1 یا 0) براساس state
            active = 1 if state == Qt.Checked else 0

            # از یک بلوک IF NOT EXISTS استفاده می‌کنیم تا یا درج یا بروزرسانی انجام شود
            sql = """
                IF NOT EXISTS (
                    SELECT 1 FROM dbo.UserAccess WHERE UserId = ? AND FormButtonsId = ?
                )
                BEGIN
                    INSERT INTO dbo.UserAccess (UserId, FormButtonsId, IsActive)
                    VALUES (?, ?, ?)
                END
                ELSE
                BEGIN
                    UPDATE dbo.UserAccess SET IsActive = ? WHERE UserId = ? AND FormButtonsId = ?
                END
            """
            # ترتیب پارامترها مطابق علامت‌گذاری‌های ? در کوئری بالا
            params = (user_id, form_button_id, user_id, form_button_id, active, active, user_id, form_button_id)
            cursor.execute(sql, params)
            self.conn.commit()
        except Exception as e:
            QMessageBox.critical(self, "خطا در ذخیره تغییر", f"خطا هنگام ذخیره تغییر:\n{str(e)}")

    # ------------------ گرفتن UserId از کاربر ------------------
    def ask_user_id(self):
        """
        فعلاً به صورت ساده از کاربر یک UserId عددی گرفته می‌شود.
        در آینده بهتر است این قسمت را با ComboBox پر از کاربران واقعی جایگزین کنیم.
        """
        user_id, ok = QInputDialog.getInt(self, "انتخاب کاربر", "UserId را وارد کنید:")
        return user_id, ok

# ----------------------------- اجرای برنامه -----------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    connect_window = ConnectWindow()
    connect_window.show()
    sys.exit(app.exec_())
