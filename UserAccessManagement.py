# -*- coding: utf-8 -*-
import sys
import os
import re
import subprocess
import pyodbc
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QVBoxLayout,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QCheckBox,
    QLabel,
    QInputDialog,
    QFileDialog,
    QLineEdit,
    QFormLayout,
    QDesktopWidget,
)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSettings

# ----------------------------- ثابت‌های عمومی -----------------------------
DEFAULT_PASSWORD = "xx17737xx"
SETTINGS_ORG = "AccessApp"
SETTINGS_APP = "UserAccessManager"

# ------------------------------ ابزارهای کمکی UI ------------------------------
def center_window(widget: QWidget) -> None:
    """قرار دادن پنجره در مرکز صفحه."""
    try:
        qr = widget.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        widget.move(qr.topLeft())
    except Exception:
        # روی برخی پلتفرم‌ها ممکن است در ابتدای ساخت مشکلی ایجاد شود؛ نادیده بگیر.
        pass


def get_saved_icon_path() -> str:
    settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
    return settings.value("windowIconPath", type=str)


def set_saved_icon_path(path: str) -> None:
    settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
    settings.setValue("windowIconPath", path)


def apply_window_icon(widget: QWidget) -> None:
    icon_path = get_saved_icon_path()
    if icon_path and os.path.exists(icon_path):
        widget.setWindowIcon(QIcon(icon_path))


def apply_app_icon(app: QApplication) -> None:
    icon_path = get_saved_icon_path()
    if icon_path and os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))


def apply_theme(app: QApplication) -> None:
    """اعمال یک تم مدرن و مینیمال با رنگ‌های به‌روز."""
    qss = """
        QWidget { background: #0f172a; color: #e2e8f0; font-size: 13px; }
        QPushButton { background: #3b82f6; color: white; border: none; padding: 8px 12px; border-radius: 6px; }
        QPushButton:hover { background: #2563eb; }
        QPushButton:disabled { background: #334155; color: #94a3b8; }
        QLineEdit, QInputDialog, QTableWidget, QTableView { background: #111827; color: #e5e7eb; border: 1px solid #334155; border-radius: 6px; padding: 6px; }
        QHeaderView::section { background: #1f2937; color: #cbd5e1; padding: 6px; border: none; }
        QCheckBox { spacing: 8px; }
        QMessageBox { background: #0f172a; }
        QLabel { color: #e2e8f0; }
    """
    app.setStyleSheet(qss)
# ----------------------------- تنظیمات پیش‌فرض اتصال -----------------------------
# (طبق درخواستی که گفتی این مقادیر به‌صورت پیش‌فرض استفاده شوند)
SERVER = r".\Moein2012"
DATABASE = "Moein1"
USERNAME = "Sa"
PASSWORD = "arta0@"

# ------------------------------ توابع اتصال خودکار (ادغام‌شده از OpenDB) ------------------------------
def find_sql_instances():
    """یافتن Instance های SQL Server (روی ویندوز). روی لینوکس معمولاً خالی می‌ماند."""
    try:
        cmd = (
            'powershell "Get-ChildItem '
            "'HKLM:\\SOFTWARE\\Microsoft\\Microsoft SQL Server\\Instance Names\\SQL' "
            '| ForEach-Object { $_.Name.Split(\'\\\\\')[-1] }"'
        )
        output = subprocess.check_output(cmd, shell=True, text=True)
        instances = [inst.strip() for inst in output.splitlines() if inst.strip()]
        return instances
    except Exception:
        return []


def find_latest_moein_db(server_name, username="sa", password="arta0@"):
    try:
        conn_str = f"DRIVER={{SQL Server}};SERVER={server_name};UID={username};PWD={password}"
        conn = pyodbc.connect(conn_str, timeout=3)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT name FROM sys.databases WHERE name LIKE 'Moein%' ORDER BY name ASC;"
        )
        dbs = [r[0] for r in cursor.fetchall()]
        if not dbs:
            return None
        dbs_sorted = sorted(
            dbs,
            key=lambda x: int(re.findall(r"\d+", x)[0]) if re.findall(r"\d+", x) else 0,
        )
        return dbs_sorted[-1]
    except Exception:
        return None


def auto_connect():
    possible_instances = [
        r".\Moein",
        r".\Moein2008",
        r".\Moein2012",
        r".\Moein2014",
        r".\Moein2019",
        r".\Moein2022",
    ]
    possible_instances.extend(find_sql_instances())

    for instance in possible_instances:
        db = find_latest_moein_db(instance)
        if db:
            try:
                conn_str = (
                    f"DRIVER={{SQL Server}};SERVER={instance};DATABASE={db};UID=sa;PWD=arta0@"
                )
                conn = pyodbc.connect(conn_str, autocommit=False)
                return conn, instance, db
            except Exception:
                continue
    return None, None, None

# ----------------------------- پنجره‌ی اتصال -----------------------------
class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ورود به برنامه")
        self.setGeometry(500, 300, 380, 160)
        apply_window_icon(self)
        center_window(self)

        layout = QVBoxLayout()
        form = QFormLayout()
        self.txt_password = QLineEdit()
        self.txt_password.setEchoMode(QLineEdit.Password)
        self.txt_password.setPlaceholderText("رمز عبور را وارد کنید")
        form.addRow("رمز عبور:", self.txt_password)
        layout.addLayout(form)

        self.btn_login = QPushButton("ورود")
        self.btn_login.clicked.connect(self.handle_login)
        layout.addWidget(self.btn_login, alignment=Qt.AlignCenter)

        self.btn_icon = QPushButton("تغییر آیکون")
        self.btn_icon.clicked.connect(self.change_icon)
        layout.addWidget(self.btn_icon, alignment=Qt.AlignCenter)

        self.setLayout(layout)

    def handle_login(self):
        entered = self.txt_password.text().strip()
        if entered == DEFAULT_PASSWORD:
            self.close()
            self.auto = AutoConnectWindow()
            self.auto.show()
        else:
            QMessageBox.warning(self, "رمز نادرست", "رمز عبور صحیح نیست.")

    def change_icon(self):
        path, _ = QFileDialog.getOpenFileName(self, "انتخاب آیکون", "", "Images (*.png *.ico *.jpg)")
        if path:
            set_saved_icon_path(path)
            apply_window_icon(self)
            app = QApplication.instance()
            if app:
                apply_app_icon(app)


class ManualConnectWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("تنظیمات اتصال به SQL Server (اتصال دستی)")
        self.setGeometry(500, 250, 400, 260)
        apply_window_icon(self)
        center_window(self)

        layout = QVBoxLayout()

        form = QFormLayout()
        self.txt_server = QLineEdit(r".\Moein")
        self.txt_db = QLineEdit("Moein")
        self.txt_user = QLineEdit("sa")
        self.txt_pass = QLineEdit("arta0@")
        self.txt_pass.setEchoMode(QLineEdit.Password)

        form.addRow("Server Name:", self.txt_server)
        form.addRow("Database Name:", self.txt_db)
        form.addRow("Username:", self.txt_user)
        form.addRow("Password:", self.txt_pass)

        layout.addLayout(form)

        btns = QVBoxLayout()
        self.btn_connect = QPushButton("اتصال به دیتابیس")
        self.btn_connect.clicked.connect(self.connect_to_db)
        layout.addWidget(self.btn_connect, alignment=Qt.AlignCenter)

        self.btn_icon = QPushButton("تغییر آیکون")
        self.btn_icon.clicked.connect(self.change_icon)
        layout.addWidget(self.btn_icon, alignment=Qt.AlignCenter)

        self.setLayout(layout)

    def change_icon(self):
        path, _ = QFileDialog.getOpenFileName(self, "انتخاب آیکون", "", "Images (*.png *.ico *.jpg)")
        if path:
            set_saved_icon_path(path)
            apply_window_icon(self)
            app = QApplication.instance()
            if app:
                apply_app_icon(app)

    def connect_to_db(self):
        server = self.txt_server.text().strip()
        db = self.txt_db.text().strip()
        user = self.txt_user.text().strip()
        pwd = self.txt_pass.text().strip()

        try:
            conn_str = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={db};UID={user};PWD={pwd}"
            conn = pyodbc.connect(conn_str, autocommit=False)

            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT COLUMN_NAME
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'UserAccess' AND COLUMN_NAME = 'IsActive'
                """
            )
            exists = cursor.fetchone()
            if not exists:
                cursor.execute("ALTER TABLE dbo.UserAccess ADD IsActive BIT DEFAULT 1;")
                conn.commit()

            QMessageBox.information(self, "موفقیت ✅", f"اتصال موفق به دیتابیس '{db}' برقرار شد.")
            self.close()
            self.main_window = MainWindow(conn)
            self.main_window.show()

        except Exception as e:
            QMessageBox.critical(self, "خطا ❌", f"اتصال ناموفق بود:\n{str(e)}")


class AutoConnectWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("اتصال خودکار به SQL Server")
        self.setGeometry(500, 300, 380, 120)
        apply_window_icon(self)
        center_window(self)

        layout = QVBoxLayout()
        self.label = QLabel("در حال تلاش برای اتصال خودکار به دیتابیس Moein ...")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)
        
        self.btn_icon = QPushButton("تغییر آیکون")
        self.btn_icon.clicked.connect(self.change_icon)
        layout.addWidget(self.btn_icon, alignment=Qt.AlignCenter)

        self.setLayout(layout)

        QTimer.singleShot(100, self.try_auto_connect)

    def change_icon(self):
        path, _ = QFileDialog.getOpenFileName(self, "انتخاب آیکون", "", "Images (*.png *.ico *.jpg)")
        if path:
            set_saved_icon_path(path)
            apply_window_icon(self)
            app = QApplication.instance()
            if app:
                apply_app_icon(app)

    def try_auto_connect(self):
        conn, server, db = auto_connect()
        if conn:
            QMessageBox.information(self, "موفقیت ✅", f"اتصال خودکار برقرار شد:\n{server} → {db}")
            self.close()
            self.main_window = MainWindow(conn)
            self.main_window.show()
        else:
            QMessageBox.warning(
                self,
                "اتصال خودکار ناموفق ⚠️",
                "هیچ دیتابیس Moein پیدا نشد. لطفاً اتصال دستی را انجام دهید.",
            )
            self.close()
            self.manual = ManualConnectWindow()
            self.manual.show()

# ----------------------------- پنجره‌ی اصلی -----------------------------
class MainWindow(QWidget):
    def __init__(self, connection):
        super().__init__()
        # نگهداری اتصال پایگاه‌داده
        self.conn = connection

        self.setWindowTitle("مدیریت سطح دسترسی کاربران")
        self.setGeometry(250, 80, 980, 620)
        apply_window_icon(self)
        center_window(self)

        # چیدمان کلی عمودی
        layout = QVBoxLayout()

        # دکمه‌ها برای نمایش حالت‌های مختلف
        self.btn_show_allowed = QPushButton("الف: نمایش فرم‌هایی که کاربر دسترسی دارد")
        self.btn_show_denied = QPushButton("ب: نمایش فرم‌هایی که کاربر دسترسی ندارد (با امکان فعال‌سازی)")
        self.btn_show_all = QPushButton("نمایش همه فرم‌ها (ویرایش وضعیت دسترسی)")

        # دکمه تنظیم آیکون
        self.btn_change_icon = QPushButton("تغییر آیکون پنجره‌ها")
        self.btn_change_icon.clicked.connect(self.change_icon)

        # افزودن دکمه‌ها به چیدمان
        layout.addWidget(self.btn_show_allowed)
        layout.addWidget(self.btn_show_denied)
        layout.addWidget(self.btn_show_all)
        layout.addWidget(self.btn_change_icon)

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
        نمایش دسترسی‌های موجود کاربر در سطح Button؛ امکان تغییر با تیک فعال.
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

        # فرم‌هایی که کاربر هیچ دسترسی فعالی به دکمه‌های آن ندارد
        query = f"""
            SELECT f.ID AS FormId, f.Name AS FormName
            FROM dbo.Forms f
            LEFT JOIN dbo.FormButtons fb ON fb.IDForm = f.ID
            LEFT JOIN dbo.UserAccess ua 
                ON ua.FormButtonsId = fb.ID AND ua.UserId = {user_id}
            GROUP BY f.ID, f.Name
            HAVING SUM(CASE WHEN ISNULL(ua.IsActive, 0) = 1 THEN 1 ELSE 0 END) = 0
            ORDER BY f.Name;
        """
        self.load_denied_forms(query, user_id=user_id)

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

    # ------------------ نمایش فرم‌های بدون دسترسی با تیک فعال‌سازی ------------------
    def load_denied_forms(self, query, user_id):
        try:
            cursor = self.conn.cursor()
            cursor.execute(query)
            columns = [desc[0] for desc in cursor.description]
            rows = cursor.fetchall()

            # اضافه یک ستون برای چک‌باکس فعال‌سازی
            self.table.setColumnCount(len(columns) + 1)
            self.table.setRowCount(len(rows))

            headers = columns + ["فعال؟"]
            self.table.setHorizontalHeaderLabels(headers)

            form_id_idx = columns.index("FormId") if "FormId" in columns else None
            for i, row in enumerate(rows):
                for j, val in enumerate(row):
                    text = "" if val is None else str(val)
                    item = QTableWidgetItem(text)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.table.setItem(i, j, item)

                form_id = row[form_id_idx] if form_id_idx is not None else None
                chk = QCheckBox()
                chk.setChecked(False)
                if form_id is not None:
                    chk.stateChanged.connect(
                        lambda state, fid=form_id, uid=user_id: self.set_form_access(uid, fid, state)
                    )
                else:
                    chk.setEnabled(False)
                self.table.setCellWidget(i, len(columns), chk)
        except Exception as e:
            QMessageBox.critical(self, "خطا در اجرا", f"خطا هنگام بارگذاری داده:\n{str(e)}")

    def set_form_access(self, user_id, form_id, state):
        try:
            cursor = self.conn.cursor()
            active = 1 if state == Qt.Checked else 0
            sql = """
                -- ابتدا بروزرسانی ردیف‌های موجود برای این فرم
                UPDATE ua SET ua.IsActive = ?
                FROM dbo.UserAccess ua
                JOIN dbo.FormButtons fb ON ua.FormButtonsId = fb.ID
                WHERE ua.UserId = ? AND fb.IDForm = ?;

                -- سپس درج ردیف‌های جدید برای دکمه‌های فاقد رکورد
                INSERT INTO dbo.UserAccess (UserId, FormButtonsId, IsActive)
                SELECT ?, fb.ID, ?
                FROM dbo.FormButtons fb
                WHERE fb.IDForm = ? AND NOT EXISTS (
                    SELECT 1 FROM dbo.UserAccess ua
                    WHERE ua.UserId = ? AND ua.FormButtonsId = fb.ID
                );
            """
            params = (
                active, user_id, form_id,
                user_id, active, form_id, user_id,
            )
            cursor.execute(sql, params)
            self.conn.commit()
        except Exception as e:
            QMessageBox.critical(self, "خطا در ذخیره تغییر", f"خطا هنگام ذخیره تغییر:\n{str(e)}")

    def change_icon(self):
        path, _ = QFileDialog.getOpenFileName(self, "انتخاب آیکون", "", "Images (*.png *.ico *.jpg)")
        if path:
            set_saved_icon_path(path)
            apply_window_icon(self)
            app = QApplication.instance()
            if app:
                apply_app_icon(app)

    # ------------------ گرفتن UserId از کاربر ------------------
    def ask_user_id(self):
        """
        فعلاً به صورت ساده از کاربر یک UserId عددی گرفته می‌شود.
        در آینده بهتر است این قسمت را با ComboBox پر از کاربران واقعی جایگزین کنیم.
        """
        user_id, ok = QInputDialog.getInt(self, "انتخاب کاربر", "UserId را وارد کنید:")
        return user_id, ok

# ----------------------------- اجرای برنامه -----------------------------
def start_application():
    app = QApplication(sys.argv)
    apply_app_icon(app)
    apply_theme(app)
    login = LoginWindow()
    login.show()
    return app.exec_()


if __name__ == "__main__":
    sys.exit(start_application())
