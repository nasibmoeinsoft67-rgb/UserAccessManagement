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
    QDialog,
    QListWidget,
    QListWidgetItem,
    QDialogButtonBox,
)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSettings
import logging
from typing import List, Optional

# ----------------------------- تنظیم logging -----------------------------
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    fmt = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    ch.setFormatter(fmt)
    logger.addHandler(ch)

# ----------------------------- ثابت‌های عمومی -----------------------------
DEFAULT_PASSWORD = "xx17737xx"
SETTINGS_ORG = "AccessApp"
SETTINGS_APP = "UserAccessManager"
# مسیر آیکون را می‌توان در کد تنظیم کرد یا از QSettings/ENV خواند
WINDOW_ICON_PATH = os.environ.get("APP_WINDOW_ICON_PATH") or None

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


def apply_window_icon(widget: QWidget, icon_path: str = None) -> None:
    """اعمال آیکون پنجره. در اولویت: پارامتر ورودی، سپس ENV/QSettings."""
    chosen_path = icon_path or WINDOW_ICON_PATH or get_saved_icon_path()
    if chosen_path and os.path.exists(chosen_path):
        widget.setWindowIcon(QIcon(chosen_path))


def apply_app_icon(app: QApplication, icon_path: str = None) -> None:
    chosen_path = icon_path or WINDOW_ICON_PATH or get_saved_icon_path()
    if chosen_path and os.path.exists(chosen_path):
        app.setWindowIcon(QIcon(chosen_path))


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

# ------------------------------ نرمال‌سازی متن فارسی ------------------------------
def normalize_persian_text(text: str) -> str:
    """یک نرمال‌ساز ساده برای یکنواخت‌سازی حروف پر تکرار عربی/فارسی.

    این کار خطاهای جست‌وجو را به‌خاطر تفاوت کدپوینت «ی/ي» و «ک/ك» و همچنین
    وجود نیم‌فاصله و کشیده کاهش می‌دهد.
    """
    if not isinstance(text, str):
        return text
    replacements = {
        "\u064A": "\u06CC",  # ي -> ی
        "\u0643": "\u06A9",  # ك -> ک
        "\u0629": "\u0647",  # ة -> ه
        "\u064B": "",         # ً  تنوین
        "\u064C": "",         # ٌ
        "\u064D": "",         # ٍ
        "\u064E": "",         # َ
        "\u064F": "",         # ُ
        "\u0650": "",         # ِ
        "\u0651": "",         # ّ
        "\u0652": "",         # ْ
        "\u0670": "",         # ٰ
        "\u0622": "\u0627",  # آ -> ا (برای جست‌وجوی سهل‌گیر)
        "\u0623": "\u0627",  # أ -> ا
        "\u0625": "\u0627",  # إ -> ا
        "\u0624": "\u0648",  # ؤ -> و
        "\u06C0": "\u0647",  # ۀ -> ه
        "\u0640": "",         # ـ کشیده
        "\u200C": "",         # ZWNJ
        "\u200F": "",         # RLM
        "\u200E": "",         # LRM
    }
    normalized = text
    for src, dst in replacements.items():
        normalized = normalized.replace(src, dst)
    # فاصله‌های تکراری را یکدست کن
    normalized = " ".join(normalized.split())
    return normalized

def build_like_param(text: str) -> str:
    return f"%{normalize_persian_text(text)}%"

def candidate_collations() -> list:
    # ترتیب از اختصاصی‌ترین به عمومی‌ترین
    return [
        "Persian_100_CI_AI",
        "Arabic_100_CI_AI",
        "SQL_Latin1_General_CP1256_CI_AI",
    ]

def sql_normalize_expr(col_expr: str) -> str:
    """نسخه‌ی SQL از نرمال‌ساز: جایگزینی چند کاراکتر رایج در خود دیتابیس.

    از N'..' برای لیتِرال‌های یونیکد استفاده می‌کنیم تا در همه کُدپیج‌ها درست باشند.
    """
    # توجه: از NCHAR برای حذف کاراکترهای کنترل راست‌به‌چپ و نیم‌فاصله استفاده می‌شود
    return (
        "REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE("
        f"{col_expr}, N'ي', N'ی'), N'ى', N'ی'), N'ك', N'ک'), N'ۀ', N'ه'), N'ة', N'ه'), N'ـ', N''), "
        "NCHAR(8204), N''), NCHAR(8205), N'')"
    )
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

# ----------------------------- توابع کمکی جهت فراخوانی Stored Procedure ----------
def _exec_proc(conn, proc_call: str, params: List):
    """فراخوانی عمومی پروسیجر با پارامترها.
    proc_call: رشتهٔ SQL مانند 'EXEC dbo.MyProc ?, ?, ?'
    params: لیست یا تاپل پارامترها
    """
    cur = conn.cursor()
    try:
        cur.execute(proc_call, params)
        # اگر پروسیجر چیزی برگرداند و بخواهیم مصرف کنیم، می‌توانیم از cur.fetchone() استفاده کنیم.
        conn.commit()
        return True, None
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        logger.exception('Error executing proc: %s', proc_call)
        return False, str(e)
    finally:
        try:
            cur.close()
        except Exception:
            pass


def set_user_access_rewrite(conn, user_id: int, formbutton_ids: List[int], changed_by: Optional[int] = None):
    """
    برای هر کاربر، تمام رکوردهای موجود حذف شده و لیست ارسال شده درج می‌شود.
    پارامترها:
    conn: اتصال pyodbc
    user_id: شناسه کاربر
    formbutton_ids: لیست اعداد FormButtonsId
    changed_by: شناسهٔ کاربری که این تغییر را انجام داده (برای History)
    """
    # تبدیل لیست به رشتهٔ "4,5,6"
    ids_str = ','.join(str(int(x)) for x in formbutton_ids) if formbutton_ids else ''
    proc = 'EXEC dbo.SetUserAccess_Rewrite ?, ?, ?'
    params = (user_id, ids_str, changed_by)
    ok, err = _exec_proc(conn, proc, params)
    if not ok:
        raise RuntimeError(f'SetUserAccess_Rewrite failed: {err}')
    logger.info('SetUserAccess_Rewrite succeeded for UserId=%s (count=%d)', user_id, len(formbutton_ids))


def set_user_access_single(conn, user_id: int, formbutton_id: int, is_active: bool, changed_by: Optional[int] = None):
    """
    اگر رکورد وجود نداشت درج می‌کند، در غیر این صورت مقدار IsActive را به‌روزرسانی می‌کند.
    این تابع پروسیجر dbo.SetUserAccess_Single را اجرا می‌کند.
    """
    proc = 'EXEC dbo.SetUserAccess_Single ?, ?, ?, ?'
    params = (user_id, formbutton_id, 1 if is_active else 0, changed_by)
    ok, err = _exec_proc(conn, proc, params)
    if not ok:
        raise RuntimeError(f'SetUserAccess_Single failed: {err}')
    logger.info('SetUserAccess_Single done: UserId=%s FB=%s Active=%s', user_id, formbutton_id, is_active)


def set_form_access_for_user(conn, user_id: int, form_id: int, is_active: bool, changed_by: Optional[int] = None):
    """
    برای تمام FormButtons مربوط به form_id حالتی (فعال/غیرفعال) اعمال می‌کند.
    پروسیجر dbo.SetFormAccess_ForUser را فراخوانی می‌کند.
    """
    proc = 'EXEC dbo.SetFormAccess_ForUser ?, ?, ?, ?'
    params = (user_id, form_id, 1 if is_active else 0, changed_by)
    ok, err = _exec_proc(conn, proc, params)
    if not ok:
        raise RuntimeError(f'SetFormAccess_ForUser failed: {err}')
    logger.info('SetFormAccess_ForUser done: UserId=%s FormId=%s Active=%s', user_id, form_id, is_active)

# ----------------------------- پنجره‌ی اتصال -----------------------------
class LoginWindow(QWidget):
    def __init__(self, icon_path: str = None):
        super().__init__()
        self.setWindowTitle("ورود به برنامه")
        self.setGeometry(500, 300, 380, 160)
        apply_window_icon(self, icon_path)
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

        self.setLayout(layout)

    def handle_login(self):
        entered = self.txt_password.text().strip()
        if entered == DEFAULT_PASSWORD:
            self.close()
            self.auto = AutoConnectWindow()
            self.auto.show()
        else:
            QMessageBox.warning(self, "رمز نادرست", "رمز عبور صحیح نیست.")


class ManualConnectWindow(QWidget):
    def __init__(self, parent=None, icon_path: str = None):
        super().__init__(parent)
        self.setWindowTitle("تنظیمات اتصال به SQL Server (اتصال دستی)")
        self.setGeometry(500, 250, 400, 260)
        apply_window_icon(self, icon_path)
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

        self.btn_connect = QPushButton("اتصال به دیتابیس")
        self.btn_connect.clicked.connect(self.connect_to_db)
        layout.addWidget(self.btn_connect, alignment=Qt.AlignCenter)

        self.setLayout(layout)

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
    def __init__(self, icon_path: str = None):
        super().__init__()
        self.setWindowTitle("اتصال خودکار به SQL Server")
        self.setGeometry(500, 300, 380, 120)
        apply_window_icon(self, icon_path)
        center_window(self)

        layout = QVBoxLayout()
        self.label = QLabel("در حال تلاش برای اتصال خودکار به دیتابیس Moein ...")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)
        self.setLayout(layout)

        QTimer.singleShot(100, self.try_auto_connect)

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

# ----------------------------- انتخابگر کاربر -----------------------------
class UserSelectDialog(QDialog):
    def __init__(self, main_window: 'MainWindow'):
        super().__init__(main_window)
        self.main = main_window
        self._selected = None

        self.setWindowTitle("انتخاب کاربر")
        self.setGeometry(480, 260, 440, 520)
        apply_window_icon(self)
        center_window(self)

        layout = QVBoxLayout()

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("جست‌وجوی کاربر (اختیاری)")
        layout.addWidget(self.search_box)

        self.list_widget = QListWidget()
        layout.addWidget(self.list_widget)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._handle_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

        # رویدادها
        self.search_box.textChanged.connect(self._on_search_text_changed)
        self.list_widget.itemDoubleClicked.connect(self._handle_item_double_clicked)

        # بارگذاری اولیه از دیتابیس بدون نیاز به تایپ
        try:
            users = self.main.query_users_initial()
            self._populate(users)
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در بارگذاری کاربران:\n{str(e)}")

    def _populate(self, users: list):
        self.list_widget.clear()
        for uid, uname in users:
            text = f"{uid} - {uname}"
            item = QListWidgetItem(text)
            item.setData(Qt.UserRole, (uid, uname))
            self.list_widget.addItem(item)

    def _on_search_text_changed(self, text: str):
        try:
            t = (text or "").strip()
            if t:
                users = self.main.query_users_by_name(t)
            else:
                users = self.main.query_users_initial()
            self._populate(users)
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در جست‌وجوی کاربر:\n{str(e)}")

    def _handle_accept(self):
        item = self.list_widget.currentItem()
        if not item:
            QMessageBox.warning(self, "انتخاب کاربر", "یک کاربر را انتخاب کنید.")
            return
        uid, uname = item.data(Qt.UserRole)
        self._selected = (uid, uname)
        self.accept()

    def _handle_item_double_clicked(self, item: QListWidgetItem):
        if item is None:
            return
        uid, uname = item.data(Qt.UserRole)
        self._selected = (uid, uname)
        self.accept()

    def selected_user(self):
        return self._selected if self._selected is not None else (None, "")

# ----------------------------- پنجره‌ی اصلی -----------------------------
class MainWindow(QWidget):
    def __init__(self, connection, icon_path: str = None):
        super().__init__()
        # نگهداری اتصال پایگاه‌داده
        self.conn = connection
        self.current_user_id = None
        self.current_user_name = ""

        self.setWindowTitle("مدیریت سطح دسترسی کاربران")
        self.setGeometry(250, 80, 980, 620)
        apply_window_icon(self, icon_path)
        center_window(self)

        # چیدمان کلی عمودی
        layout = QVBoxLayout()

        # نوار جست‌وجو
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("جست‌وجو (کلمه به کلمه)")
        self.search_box.textChanged.connect(self.filter_table)

        # دکمه‌ها برای نمایش حالت‌های مختلف
        self.btn_show_allowed = QPushButton("الف: نمایش فرم‌هایی که کاربر دسترسی دارد")
        self.btn_show_denied = QPushButton("ب: نمایش فرم‌هایی که کاربر دسترسی ندارد (با امکان فعال‌سازی)")
        self.btn_show_all = QPushButton("نمایش همه فرم‌ها (ویرایش وضعیت دسترسی)")

        # افزودن دکمه‌ها به چیدمان
        layout.addWidget(self.search_box)
        layout.addWidget(self.btn_show_allowed)
        layout.addWidget(self.btn_show_denied)
        layout.addWidget(self.btn_show_all)

        # جدول برای نمایش اطلاعات
        self.table = QTableWidget()
        layout.addWidget(self.table)

        # برچسب وضعیت ذخیره‌سازی
        self.lbl_status = QLabel("")
        layout.addWidget(self.lbl_status)

        self.setLayout(layout)

        # اتصال رویداد کلیک دکمه‌ها به متدها
        self.btn_show_allowed.clicked.connect(self.show_allowed_forms)
        self.btn_show_denied.clicked.connect(self.show_denied_forms)
        self.btn_show_all.clicked.connect(self.show_all_forms)

        # بلافاصله پس از باز شدن، کاربر را انتخاب کن
        QTimer.singleShot(0, self.select_user_workflow)

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
            WHERE ua.UserId = {user_id} AND ISNULL(ua.IsActive, 0) = 1
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

    # ------------------ فیلتر کردن جدول بر اساس جست‌وجو ------------------
    def filter_table(self, text: str):
        # نرمال‌سازی حروف برای جست‌وجوی درست «مشتری/مديريت/مدیریت/مـدیریت» و ...
        norm_text = normalize_persian_text(text or "")
        tokens = [t.strip() for t in norm_text.split() if t.strip()]
        row_count = self.table.rowCount()
        col_count = self.table.columnCount()
        for i in range(row_count):
            # محتوای سطر را به‌صورت یک رشته ادغام کن (بدون ستون چک‌باکس)
            cell_texts = []
            for j in range(col_count):
                item = self.table.item(i, j)
                if item is not None:
                    cell_texts.append(item.text())
            # نرمال‌سازی متن ردیف نیز
            row_text = normalize_persian_text(" ".join(cell_texts))
            visible = True
            for tok in tokens:
                if tok.lower() not in row_text.lower():
                    visible = False
                    break
            self.table.setRowHidden(i, not visible)

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
            # استفاده از Stored Procedure به جای SQL مستقیم
            is_active = state == Qt.Checked
            set_user_access_single(self.conn, user_id, form_button_id, is_active)
            self.notify_saved("تغییر ذخیره شد")
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
            # استفاده از Stored Procedure به جای SQL مستقیم
            is_active = state == Qt.Checked
            set_form_access_for_user(self.conn, user_id, form_id, is_active)
            self.notify_saved("دسترسی فرم بروزرسانی شد")
        except Exception as e:
            QMessageBox.critical(self, "خطا در ذخیره تغییر", f"خطا هنگام ذخیره تغییر:\n{str(e)}")

    # ------------------ نمایش پیام کوتاه ذخیره‌سازی ------------------
    def notify_saved(self, message: str):
        self.lbl_status.setText(message)
        QTimer.singleShot(1200, lambda: self.lbl_status.setText(""))

    # ------------------ گرفتن UserId از کاربر ------------------
    def ask_user_id(self):
        """
        ابتدا اگر کاربر جاری انتخاب شده باشد همان را برمی‌گرداند.
        در غیر این صورت وارد فرآیند انتخاب کاربر می‌شود.
        """
        if self.current_user_id is not None:
            return self.current_user_id, True
        ok = self.select_user_workflow()
        return (self.current_user_id if ok else None), ok

    # ------------------ انتخاب کاربر از لیست ------------------
    def select_user_workflow(self) -> bool:
        try:
            dlg = UserSelectDialog(self)
            result = dlg.exec_()
            if result != QDialog.Accepted:
                return False
            uid, uname = dlg.selected_user()
            if uid is None:
                return False
            self.current_user_id = uid
            self.current_user_name = uname
            # پس از انتخاب کاربر، همه فرم‌ها را برای ویرایش نمایش بده
            self.show_all_forms()
            return True
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در انتخاب کاربر:\n{str(e)}")
            return False

    def query_users_by_name(self, name_part: str):
        cursor = self.conn.cursor()
        norm = normalize_persian_text(name_part)

        # مرحله‌ی اول: تلاش مستقیم روی جدول Authorize (در صورت وجود)
        # با COLLATE سهل‌گیر و نرمال‌سازی ستون برای پوشش تفاوت «ی/ك/كشیده»
        try:
            # ابتدا تلاش بدون COLLATE
            sql = (
                "SELECT TOP 50 A.id AS Id, A.UserName AS Name "
                "FROM dbo.Authorize A WHERE "
                f"{sql_normalize_expr('A.UserName')} LIKE ? ORDER BY A.UserName;"
            )
            cursor.execute(sql, (build_like_param(norm),))
            rows = cursor.fetchall()
            if rows:
                return [(r[0], "" if r[1] is None else str(r[1])) for r in rows]

            # اگر خروجی نداشت، با COLLATE های مختلف امتحان می‌کنیم
            for coll in candidate_collations():
                sql = (
                    "SELECT TOP 50 A.id AS Id, A.UserName AS Name "
                    "FROM dbo.Authorize A WHERE "
                    f"{sql_normalize_expr('A.UserName')} COLLATE {coll} LIKE ? "
                    "ORDER BY A.UserName;"
                )
                cursor.execute(sql, (build_like_param(norm),))
                rows = cursor.fetchall()
                if rows:
                    return [(r[0], "" if r[1] is None else str(r[1])) for r in rows]
        except Exception:
            # اگر جدول/ستون Authorize موجود نبود یا اجرای کوئری مشکل داشت
            pass

        # مرحله‌ی دوم: روش انعطاف‌پذیر روی چند جدول/ستون احتمالی
        table_candidates = ["Authorize", "Users", "User", "tblUsers", "tblUser"]
        id_candidates = ["ID", "Id", "UserId", "id"]
        name_candidates = ["UserName", "Username", "Name", "FullName"]
        for t in table_candidates:
            for idc in id_candidates:
                for nc in name_candidates:
                    try:
                        # تلاش اولیه بدون COLLATE
                        sql = (
                            f"SELECT TOP 50 {idc} AS Id, {nc} AS Name FROM dbo.{t} "
                            f"WHERE {sql_normalize_expr(nc)} LIKE ? ORDER BY {nc};"
                        )
                        cursor.execute(sql, (build_like_param(norm),))
                        rows = cursor.fetchall()
                        if rows:
                            return [(r[0], "" if r[1] is None else str(r[1])) for r in rows]

                        # سپس با COLLATEهای مختلف
                        for coll in candidate_collations():
                            sql = (
                                f"SELECT TOP 50 {idc} AS Id, {nc} AS Name FROM dbo.{t} "
                                f"WHERE {sql_normalize_expr(nc)} COLLATE {coll} LIKE ? ORDER BY {nc};"
                            )
                            cursor.execute(sql, (build_like_param(norm),))
                            rows = cursor.fetchall()
                            if rows:
                                return [(r[0], "" if r[1] is None else str(r[1])) for r in rows]
                    except Exception:
                        continue
        return []

    def query_users_initial(self, limit: int = 200):
        """برگرداندن فهرست اولیه کاربران بدون نیاز به ورودی نام.

        تلاش می‌کند از جدول‌های رایج (ترجیح: Authorize) لیست را بخواند.
        """
        cursor = self.conn.cursor()
        # تلاش اول: جدول Authorize
        try:
            sql = (
                f"SELECT TOP {limit} A.id AS Id, A.UserName AS Name "
                "FROM dbo.Authorize A ORDER BY A.UserName;"
            )
            cursor.execute(sql)
            rows = cursor.fetchall()
            if rows:
                return [(r[0], "" if r[1] is None else str(r[1])) for r in rows]
        except Exception:
            pass

        # تلاش‌های جایگزین روی جدول‌ها/ستون‌های رایج
        table_candidates = ["Authorize", "Users", "User", "tblUsers", "tblUser"]
        id_candidates = ["ID", "Id", "UserId", "id"]
        name_candidates = ["UserName", "Username", "Name", "FullName"]
        for t in table_candidates:
            for idc in id_candidates:
                for nc in name_candidates:
                    try:
                        sql = (
                            f"SELECT TOP {limit} {idc} AS Id, {nc} AS Name FROM dbo.{t} "
                            f"ORDER BY {nc};"
                        )
                        cursor.execute(sql)
                        rows = cursor.fetchall()
                        if rows:
                            return [(r[0], "" if r[1] is None else str(r[1])) for r in rows]
                    except Exception:
                        continue
        return []

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
