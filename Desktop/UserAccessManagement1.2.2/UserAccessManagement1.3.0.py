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
    QHBoxLayout,
    QTreeWidget,
    QTreeWidgetItem,
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
    QHeaderView,
)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QIcon, QBrush, QColor
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


def apply_app_icon(app: QApplication) -> None:
    chosen_path = WINDOW_ICON_PATH or get_saved_icon_path()
    if chosen_path and os.path.exists(chosen_path):
        app.setWindowIcon(QIcon(chosen_path))


def apply_theme(app: QApplication) -> None:
    """اعمال یک تم مدرن و مینیمال با رنگ‌های به‌روز."""
    qss = """
        QWidget { background: #0f172a; color: #e2e8f0; font-size: 13px; }
        QPushButton { background: #3b82f6; color: white; border: none; padding: 8px 12px; border-radius: 6px; }
        QPushButton:hover { background: #2563eb; }
        QPushButton:disabled { background: #334155; color: #94a3b8; }
        QLineEdit, QInputDialog, QTableWidget, QTableView, QTreeWidget { background: #111827; color: #e5e7eb; border: 1px solid #334155; border-radius: 6px; padding: 6px; }
        QHeaderView::section { background: #1f2937; color: #cbd5e1; padding: 6px; border: none; }
        QCheckBox { spacing: 8px; }
        QMessageBox { background: #0f172a; }
        QLabel { color: #e2e8f0; }
    """
    app.setStyleSheet(qss)

# ----------------------------- نرمال‌سازی متن فارسی ------------------------------
def normalize_persian_text(text: str) -> str:
    """یک نرمال‌ساز ساده برای یکنواخت‌سازی حروف پر تکرار عربی/فارسی."""
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
        "\u0622": "\u0627",  # آ -> ا
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
    normalized = " ".join(normalized.split())
    return normalized

def build_like_param(text: str) -> str:
    return f"%{normalize_persian_text(text)}%"

def candidate_collations() -> list:
    return [
        "Persian_100_CI_AI",
        "Arabic_100_CI_AI",
        "SQL_Latin1_General_CP1256_CI_AI",
    ]

def sql_normalize_expr(col_expr: str) -> str:
    """نسخه‌ی SQL از نرمال‌ساز: جایگزینی چند کاراکتر رایج در خود دیتابیس."""
    return (
        "REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE("
        f"{col_expr}, N'ي', N'ی'), N'ى', N'ی'), N'ك', N'ک'), N'ۀ', N'ه'), N'ة', N'ه'), N'ـ', N''), "
        "NCHAR(8204), N''), NCHAR(8205), N'')"
    )

# ----------------------------- تنظیمات پیش‌فرض اتصال -----------------------------
SERVER = r".\Moein2012"
DATABASE = "Moein1"
USERNAME = "Sa"
PASSWORD = "arta0@"

# ------------------------------ توابع اتصال خودکار ------------------------------
def find_sql_instances():
    """یافتن Instance های SQL Server (روی ویندوز)."""
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

# ----------------------------- اطمینان از وجود ستون IsActive -----------------------------
def ensure_isactive_column(conn) -> None:
    """ستون `IsActive` را در صورت نبود روی جدول `dbo.UserAccess` می‌سازد."""
    cur = conn.cursor()
    try:
        cur.execute(
            """
            SELECT 1
            FROM INFORMATION_SCHEMA.TABLES
            WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'UserAccess'
            """
        )
        tbl_exists = cur.fetchone() is not None
        if not tbl_exists:
            logger.warning("Table dbo.UserAccess not found; skip ensuring IsActive column.")
            return

        cur.execute(
            """
            SELECT 1
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'UserAccess' AND COLUMN_NAME = 'IsActive'
            """
        )
        col_exists = cur.fetchone() is not None

        if not col_exists:
            cur.execute(
                """
                ALTER TABLE dbo.UserAccess
                ADD IsActive BIT NOT NULL CONSTRAINT DF_UserAccess_IsActive DEFAULT (1) WITH VALUES;
                """
            )
            conn.commit()
            logger.info("IsActive column added to dbo.UserAccess with default 1.")
        else:
            cur.execute("UPDATE dbo.UserAccess SET IsActive = 1 WHERE IsActive IS NULL;")
            conn.commit()
            logger.info("IsActive nulls backfilled to 1 on dbo.UserAccess.")
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        logger.exception("Failed to ensure IsActive column on dbo.UserAccess")
        raise
    finally:
        try:
            cur.close()
        except Exception:
            pass

# ----------------------------- بررسی وجود Stored Procedure -----------------------------
def stored_procedure_exists(conn, schema_name: str, procedure_name: str) -> bool:
    """بررسی می‌کند که پروسیجر مشخص در دیتابیس جاری وجود دارد یا خیر."""
    cursor = conn.cursor()
    try:
        cursor.execute(
            """
            SELECT 1
            FROM sys.procedures
            WHERE name = ? AND schema_id = SCHEMA_ID(?)
            """,
            (procedure_name, schema_name),
        )
        return cursor.fetchone() is not None
    except Exception:
        return False
    finally:
        try:
            cursor.close()
        except Exception:
            pass

# ----------------------------- پیاده‌سازی‌های جایگزین (Fallback) ----------------------
def _fallback_set_user_access_single(conn, user_id: int, formbutton_id: int, is_active: bool) -> None:
    """اگر پروسیجر SetUserAccess_Single نبود، سطر مربوطه را upsert می‌کند."""
    cursor = conn.cursor()
    try:
        if is_active:
            cursor.execute(
                """
                IF NOT EXISTS (SELECT 1 FROM dbo.UserAccess WHERE UserId = ? AND FormButtonsId = ?)
                INSERT INTO dbo.UserAccess (UserId, FormButtonsId, IsActive) VALUES (?, ?, 1);
                ELSE
                UPDATE dbo.UserAccess SET IsActive = 1 WHERE UserId = ? AND FormButtonsId = ?;
                """,
                (user_id, formbutton_id, user_id, formbutton_id, user_id, formbutton_id),
            )
        else:
            cursor.execute(
                """
                DELETE FROM dbo.UserAccess
                WHERE UserId = ? AND FormButtonsId = ?;
                """,
                (user_id, formbutton_id),
            )
        conn.commit()
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        raise
    finally:
        try:
            cursor.close()
        except Exception:
            pass


def _fallback_set_form_access_for_user(conn, user_id: int, form_id: int, is_active: bool) -> None:
    """اگر SetFormAccess_ForUser نبود، برای تمام دکمه‌های فرم، دسترسی کاربر را تنظیم می‌کند."""
    cursor = conn.cursor()
    try:
        if is_active:
            cursor.execute(
                """
                UPDATE ua
                SET ua.IsActive = 1
                FROM dbo.UserAccess AS ua
                JOIN dbo.FormButtons AS fb ON fb.ID = ua.FormButtonsId
                WHERE ua.UserId = ? AND fb.IDForm = ?;
                """,
                (user_id, form_id),
            )

            cursor.execute(
                """
                INSERT INTO dbo.UserAccess (UserId, FormButtonsId, IsActive)
                SELECT ?, fb.ID, 1
                FROM dbo.FormButtons AS fb
                WHERE fb.IDForm = ?
                  AND NOT EXISTS (
                      SELECT 1 FROM dbo.UserAccess AS ua
                      WHERE ua.UserId = ? AND ua.FormButtonsId = fb.ID
                  );
                """,
                (user_id, form_id, user_id),
            )
        else:
            cursor.execute(
                """
                DELETE ua
                FROM dbo.UserAccess AS ua
                JOIN dbo.FormButtons AS fb ON fb.ID = ua.FormButtonsId
                WHERE ua.UserId = ? AND fb.IDForm = ?;
                """,
                (user_id, form_id),
            )
        conn.commit()
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        raise
    finally:
        try:
            cursor.close()
        except Exception:
            pass


def _fallback_set_user_access_rewrite(conn, user_id: int, formbutton_ids: List[int]) -> None:
    """اگر SetUserAccess_Rewrite نبود، دسترسی‌های کاربر را کامل بازنویسی می‌کند."""
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM dbo.UserAccess WHERE UserId = ?;", (user_id,))
        ids_to_insert = [(user_id, int(fid)) for fid in (formbutton_ids or [])]
        if ids_to_insert:
            cursor.executemany(
                "INSERT INTO dbo.UserAccess (UserId, FormButtonsId, IsActive) VALUES (?, ?, 1);",
                ids_to_insert,
            )
        conn.commit()
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        raise
    finally:
        try:
            cursor.close()
        except Exception:
            pass

# ----------------------------- توابع کمکی جهت فراخوانی Stored Procedure ----------
def _exec_proc(conn, proc_call: str, params: List):
    """فراخوانی عمومی پروسیجر با پارامترها."""
    cur = conn.cursor()
    try:
        cur.execute(proc_call, params)
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
    """برای هر کاربر، تمام رکوردهای موجود حذف شده و لیست ارسال شده درج می‌شود."""
    ids_str = ','.join(str(int(x)) for x in formbutton_ids) if formbutton_ids else ''
    if stored_procedure_exists(conn, 'dbo', 'SetUserAccess_Rewrite'):
        proc = 'EXEC dbo.SetUserAccess_Rewrite ?, ?, ?'
        params = (user_id, ids_str, changed_by)
        ok, err = _exec_proc(conn, proc, params)
        if not ok:
            raise RuntimeError(f'SetUserAccess_Rewrite failed: {err}')
    else:
        _fallback_set_user_access_rewrite(conn, user_id, formbutton_ids)
    logger.info('SetUserAccess_Rewrite succeeded for UserId=%s (count=%d)', user_id, len(formbutton_ids))


def set_user_access_single(conn, user_id: int, formbutton_id: int, is_active: bool, changed_by: Optional[int] = None):
    """اگر رکورد وجود نداشت درج می‌کند، در غیر این صورت مقدار IsActive را به‌روزرسانی می‌کند."""
    _fallback_set_user_access_single(conn, user_id, formbutton_id, is_active)
    logger.info('SetUserAccess_Single done: UserId=%s FB=%s Active=%s', user_id, formbutton_id, is_active)


def set_form_access_for_user(conn, user_id: int, form_id: int, is_active: bool, changed_by: Optional[int] = None):
    """برای تمام FormButtons مربوط به form_id حالتی (فعال/غیرفعال) اعمال می‌کند."""
    _fallback_set_form_access_for_user(conn, user_id, form_id, is_active)
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
            ensure_isactive_column(conn)

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
            try:
                ensure_isactive_column(conn)
            except Exception as e:
                QMessageBox.warning(self, "خطای سازگارسازی اسکیم", f"عدم موفقیت در تضمین ستون IsActive:\n{str(e)}")
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

        self.search_box.textChanged.connect(self._on_search_text_changed)
        self.list_widget.itemDoubleClicked.connect(self._handle_item_double_clicked)

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
        self.conn = connection
        self.current_user_id = None
        self.current_user_name = ""
        self.current_mode = None

        self.setWindowTitle("مدیریت سطح دسترسی کاربران")
        self.setGeometry(250, 80, 1200, 700)
        apply_window_icon(self, icon_path)
        center_window(self)

        layout = QVBoxLayout()

        # نوار جست‌وجو
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("جست‌وجو (کلمه به کلمه)")
        self.search_box.textChanged.connect(self.filter_tree_and_table)
        layout.addWidget(self.search_box)

        # دکمه‌ها
        button_layout = QVBoxLayout()
        self.btn_show_allowed = QPushButton("نمایش فرم های مجاز")
        self.btn_show_denied = QPushButton("نمایش فرم های غیرمجاز")
        self.btn_show_all = QPushButton("نمایش همه فرم‌ها")
        self.btn_save_all = QPushButton("ذخیره تغییرات")
        button_layout.addWidget(self.btn_show_allowed)
        button_layout.addWidget(self.btn_show_denied)
        button_layout.addWidget(self.btn_show_all)
        button_layout.addWidget(self.btn_save_all)

        # چیدمان دو ستونه: درخت + جدول
        main_layout = QHBoxLayout()
        
        # درخت سلسله مراتبی فرم‌ها
        self.tree_widget = QTreeWidget()
        self.tree_widget.setHeaderLabels(["دسترسی", "وضعیت"])
        self.tree_widget.setColumnWidth(0, 600)
        self.tree_widget.itemChanged.connect(self.on_tree_item_changed)
        main_layout.addWidget(self.tree_widget, 1)

        # جدول جزئیات
        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(2)
        self.table_widget.setHorizontalHeaderLabels(["کد", "نام"])
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        main_layout.addWidget(self.table_widget, 1)

        layout.addLayout(button_layout)
        layout.addLayout(main_layout)

        self.lbl_status = QLabel("")
        layout.addWidget(self.lbl_status)

        self.setLayout(layout)

        self.btn_show_allowed.clicked.connect(self.show_allowed_forms)
        self.btn_show_denied.clicked.connect(self.show_denied_forms)
        self.btn_show_all.clicked.connect(self.show_all_forms)
        self.btn_save_all.clicked.connect(self.save_all_changes)

        QTimer.singleShot(0, self.select_user_workflow)

    def on_tree_item_changed(self, item, column):
        """مدیریت تغییر وضعیت checkbox"""
        if column != 0:
            return
        
        # بررسی اینکه آیا آیتم یک دکمه فرم است یا خود فرم
        item_type = item.data(0, Qt.UserRole + 1)
        if item_type == "button":
            # این یک دکمه فرم است
            button_id = item.data(0, Qt.UserRole)
            is_checked = item.checkState(0) == Qt.Checked
            
            # به‌روزرسانی وضعیت والد (فرم)
            parent = item.parent()
            if parent:
                self.update_parent_state(parent)
        elif item_type == "form":
            # این یک فرم است - باید همه فرزندان را به‌روزرسانی کند
            form_id = item.data(0, Qt.UserRole)
            is_checked = item.checkState(0) == Qt.Checked
            
            # غیرفعال کردن سیگنال برای جلوگیری از فراخوانی مکرر
            self.tree_widget.blockSignals(True)
            for i in range(item.childCount()):
                child = item.child(i)
                # اطمینان از اینکه فرزندان نیز قابل تیک خوردن هستند
                if child.flags() & Qt.ItemIsUserCheckable:
                    child.setCheckState(0, Qt.Checked if is_checked else Qt.Unchecked)
            self.tree_widget.blockSignals(False)
            # پس از به‌روزرسانی فرزندان، وضعیت والد را مجدداً تنظیم کن
            self.update_parent_state(item)


    def update_parent_state(self, parent_item):
        """به‌روزرسانی وضعیت checkbox والد بر اساس فرزندان"""
        if not parent_item:
            return
        
        checked_count = 0
        partially_checked = False
        total_checkable_children = 0
        
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            # فقط فرزندانی که قابلیت تیک خوردن دارند را در نظر بگیر
            if child.flags() & Qt.ItemIsUserCheckable:
                total_checkable_children += 1
                state = child.checkState(0)
                if state == Qt.Checked:
                    checked_count += 1
                elif state == Qt.PartiallyChecked:
                    partially_checked = True

        # اگر هیچ فرزند قابل تیک خوردنی وجود ندارد، والد نباید قابل تیک خوردن باشد
        if total_checkable_children == 0:
             # غیرفعال کردن حالت تیک خوردن برای والد اگر فرزندی برای تیک زدن ندارد
             parent_item.setFlags(parent_item.flags() & ~Qt.ItemIsUserCheckable)
             parent_item.setCheckState(0, Qt.Unchecked)
             return
        else:
            # اطمینان از اینکه والد قابلیت تیک خوردن دارد اگر فرزندان دارد
            if not (parent_item.flags() & Qt.ItemIsUserCheckable):
                parent_item.setFlags(parent_item.flags() | Qt.ItemIsUserCheckable)
        
        # غیرفعال کردن سیگنال برای جلوگیری از فراخوانی مکرر
        self.tree_widget.blockSignals(True)
        if partially_checked:
            parent_item.setCheckState(0, Qt.PartiallyChecked)
        elif checked_count == 0:
            parent_item.setCheckState(0, Qt.Unchecked)
        elif checked_count == total_checkable_children:
            parent_item.setCheckState(0, Qt.Checked)
        else: # checked_count > 0 and checked_count < total_checkable_children
             parent_item.setCheckState(0, Qt.PartiallyChecked)
        self.tree_widget.blockSignals(False)


    def save_all_changes(self):
        """جمع‌آوری همه checkbox های checked و rewrite دسترسی‌ها."""
        if self.current_user_id is None:
            QMessageBox.warning(self, "انتخاب کاربر", "ابتدا یک کاربر انتخاب کنید.")
            return

        checked_button_ids = []
        for i in range(self.tree_widget.topLevelItemCount()):
            form_item = self.tree_widget.topLevelItem(i)
            # بررسی فرزندان (دکمه‌های فرم)
            for j in range(form_item.childCount()):
                button_item = form_item.child(j)
                # اطمینان از اینکه آیتم قابل تیک خوردن است و تیک خورده است
                if button_item.flags() & Qt.ItemIsUserCheckable and button_item.checkState(0) == Qt.Checked:
                    button_id = button_item.data(0, Qt.UserRole)
                    if button_id:
                        checked_button_ids.append(button_id)

        try:
            set_user_access_rewrite(self.conn, self.current_user_id, checked_button_ids)
            self.notify_saved("همه تغییرات ذخیره شد")
            self.refresh_main_app()
            QTimer.singleShot(100, self.reload_current_mode)
        except Exception as e:
            QMessageBox.critical(self, "خطا در ذخیره", f"خطا هنگام ذخیره تغییرات:\n{str(e)}")

    def refresh_main_app(self):
        """تلاش برای رفرش برنامه اصلی حسابداری (Moein)."""
        try:
            # تلاش برای بستن برنامه اصلی در صورت اجرا بودن
            subprocess.call(['taskkill', '/IM', 'Moein.exe', '/F'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            # سعی در یافتن مسیر برنامه اصلی
            # این مسیر باید بر اساس نصب واقعی برنامه تنظیم شود
            main_app_path = None
            possible_paths = [
                r"C:\Program Files (x86)\MoeinSystem\Moein.exe",
                r"C:\Program Files\MoeinSystem\Moein.exe",
                r"C:\Moein\Moein.exe", # مثال: مسیر نصب در درایو C
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    main_app_path = path
                    break
            
            if main_app_path:
                subprocess.Popen(main_app_path)
                logger.info(f"برنامه اصلی Moein.exe در مسیر {main_app_path} اجرا شد.")
            else:
                logger.warning("مسیر برنامه اصلی Moein.exe یافت نشد. لطفا مسیر را در کد تنظیم کنید.")
        except Exception as e:
            logger.exception("خطا در رفرش برنامه اصلی: %s", str(e))

    def show_allowed_forms(self):
        user_id, ok = self.ask_user_id()
        if not ok:
            return
        self.current_mode = 'allowed'
        self.load_hierarchy(user_id, True)

    def show_denied_forms(self):
        user_id, ok = self.ask_user_id()
        if not ok:
            return
        self.current_mode = 'denied'
        self.load_hierarchy(user_id, False)

    def show_all_forms(self):
        user_id, ok = self.ask_user_id()
        if not ok:
            return
        self.current_mode = 'all'
        self.load_hierarchy(user_id, None)

    def load_hierarchy(self, user_id, is_allowed):
        """بارگذاری درخت سلسله مراتبی فرم‌ها و دکمه‌های آنها"""
        self.tree_widget.clear()
        self.table_widget.setRowCount(0)
        
        self.tree_widget.blockSignals(True)
        
        try:
            cursor = self.conn.cursor()
            
            query = """
            SELECT 
                f.ID as FormId,
                f.Name as FormName,
                fb.ID as ButtonId,
                fb.Name as ButtonName,
                ISNULL(ua.IsActive, 0) as HasAccess
            FROM dbo.Forms f
            LEFT JOIN dbo.FormButtons fb ON fb.IDForm = f.ID
            LEFT JOIN dbo.UserAccess ua ON ua.FormButtonsId = fb.ID AND ua.UserId = ?
            ORDER BY f.Name, fb.Name;
            """
            cursor.execute(query, (user_id,))
            rows = cursor.fetchall()

            # گروه‌بندی داده‌ها بر اساس فرم
            forms_dict = {}
            for row in rows:
                form_id, form_name, button_id, button_name, has_access = row
                
                if form_id not in forms_dict:
                    forms_dict[form_id] = {
                        'name': form_name,
                        'buttons': []
                    }
                
                if button_id:
                    forms_dict[form_id]['buttons'].append({
                        'id': button_id,
                        'name': button_name,
                        'access': has_access
                    })

            # ساخت درخت
            for form_id, form_data in forms_dict.items():
                # بررسی فیلتر
                if is_allowed is not None:
                    # بررسی اینکه آیا حداقل یک دکمه دسترسی دارد
                    has_any_access = any(btn['access'] for btn in form_data['buttons'])
                    if is_allowed and not has_any_access:
                        continue
                    if not is_allowed and has_any_access:
                        continue

                # ساخت آیتم فرم (والد)
                form_item = QTreeWidgetItem(self.tree_widget, [form_data['name'], ""])
                form_item.setData(0, Qt.UserRole, form_id)
                form_item.setData(0, Qt.UserRole + 1, "form")
                
                # اضافه کردن checkbox در حالت 'all'
                if self.current_mode == 'all':
                    form_item.setFlags(form_item.flags() | Qt.ItemIsUserCheckable)
                    form_item.setCheckState(0, Qt.Unchecked) # پیش‌فرض Unchecked

                # اضافه کردن دکمه‌های فرم (فرزندان)
                checked_count = 0
                has_buttons = False
                for button in form_data['buttons']:
                    has_buttons = True
                    button_item = QTreeWidgetItem(form_item, [button['name'], ""])
                    button_item.setData(0, Qt.UserRole, button['id'])
                    button_item.setData(0, Qt.UserRole + 1, "button")
                    
                    # اضافه کردن checkbox در حالت 'all'
                    if self.current_mode == 'all':
                        button_item.setFlags(button_item.flags() | Qt.ItemIsUserCheckable)
                        button_item.setCheckState(0, Qt.Checked if button['access'] else Qt.Unchecked)
                        if button['access']:
                            checked_count += 1
                    
                    # رنگ‌آمیزی بر اساس دسترسی
                    if button['access']:
                        button_item.setForeground(0, QBrush(QColor("#10b981")))  # سبز
                    else:
                        button_item.setForeground(0, QBrush(QColor("#ef4444")))  # قرمز

                # تنظیم وضعیت checkbox والد بر اساس فرزندان
                if self.current_mode == 'all' and has_buttons:
                    total_buttons = len(form_data['buttons'])
                    if checked_count == 0:
                        form_item.setCheckState(0, Qt.Unchecked)
                    elif checked_count == total_buttons:
                        form_item.setCheckState(0, Qt.Checked)
                    else:
                        form_item.setCheckState(0, Qt.PartiallyChecked)
                elif self.current_mode == 'all' and not has_buttons:
                    # اگر فرم دکمه‌ای ندارد، والد نباید قابل تیک خوردن باشد
                    form_item.setFlags(form_item.flags() & ~Qt.ItemIsUserCheckable)
                    form_item.setCheckState(0, Qt.Unchecked)


            self.tree_widget.expandAll()

            button_count = sum(len(f['buttons']) for f in forms_dict.values())
            self.table_widget.setRowCount(button_count)
            row_idx = 0
            for form_id, form_data in forms_dict.items():
                for button in form_data['buttons']:
                    self.table_widget.setItem(row_idx, 0, QTableWidgetItem(str(button['id'])))
                    self.table_widget.setItem(row_idx, 1, QTableWidgetItem(f"{form_data['name']} - {button['name']}"))
                    row_idx += 1

        except Exception as e:
            QMessageBox.critical(self, "خطا در اجرا", f"خطا هنگام بارگذاری داده:\n{str(e)}")
        finally:
            self.tree_widget.blockSignals(False)

    def filter_tree_and_table(self, text: str):
        """فیلتر کردن درخت و جدول بر اساس جست‌وجو"""
        norm_text = normalize_persian_text(text or "")
        tokens = [t.strip() for t in norm_text.split() if t.strip()]
        
        # فیلتر درخت
        for i in range(self.tree_widget.topLevelItemCount()):
            item = self.tree_widget.topLevelItem(i)
            self.filter_tree_item(item, tokens)
            for j in range(item.childCount()):
                child = item.child(j)
                self.filter_tree_item(child, tokens)
        
        # فیلتر جدول
        for i in range(self.table_widget.rowCount()):
            cell_texts = []
            for j in range(self.table_widget.columnCount()):
                item = self.table_widget.item(i, j)
                if item is not None:
                    cell_texts.append(item.text())
            row_text = normalize_persian_text(" ".join(cell_texts))
            visible = all(tok.lower() in row_text.lower() for tok in tokens)
            self.table_widget.setRowHidden(i, not visible)

    def filter_tree_item(self, item, tokens):
        """فیلتر کردن یک آیتم درخت"""
        row_text = normalize_persian_text(item.text(0))
        visible = all(tok.lower() in row_text.lower() for tok in tokens)
        item.setHidden(not visible)

    def reload_current_mode(self):
        """بارگذاری مجدد حالت فعلی"""
        if self.current_mode == 'allowed':
            self.show_allowed_forms()
        elif self.current_mode == 'denied':
            self.show_denied_forms()
        elif self.current_mode == 'all':
            self.show_all_forms()

    def notify_saved(self, message: str):
        """نمایش پیام ذخیره‌سازی"""
        self.lbl_status.setText(message)
        QTimer.singleShot(1200, lambda: self.lbl_status.setText(""))

    def ask_user_id(self):
        """دریافت شناسه کاربر"""
        if self.current_user_id is not None:
            return self.current_user_id, True
        ok = self.select_user_workflow()
        return (self.current_user_id if ok else None), ok

    def select_user_workflow(self) -> bool:
        """انتخاب کاربر از دیالوگ"""
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
            self.show_all_forms()
            return True
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در انتخاب کاربر:\n{str(e)}")
            return False

    def query_users_by_name(self, name_part: str):
        """جست‌وجوی کاربران بر اساس نام - حفظ ساختار کوئری از 1.2.6"""
        cursor = self.conn.cursor()
        norm = normalize_persian_text(name_part)

        try:
            sql = (
                "SELECT TOP 50 A.id AS Id, A.UserName AS Name "
                "FROM dbo.Authorize A WHERE "
                f"{sql_normalize_expr('A.UserName')} LIKE ? ORDER BY A.UserName;"
            )
            cursor.execute(sql, (build_like_param(norm),))
            rows = cursor.fetchall()
            if rows:
                return [(r[0], "" if r[1] is None else str(r[1])) for r in rows]

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
            pass

        table_candidates = ["Authorize", "Users", "User", "tblUsers", "tblUser"]
        id_candidates = ["ID", "Id", "UserId", "id"]
        name_candidates = ["UserName", "Username", "Name", "FullName"]
        for t in table_candidates:
            for idc in id_candidates:
                for nc in name_candidates:
                    try:
                        sql = (
                            f"SELECT TOP 50 {idc} AS Id, {nc} AS Name FROM dbo.{t} "
                            f"WHERE {sql_normalize_expr(nc)} LIKE ? ORDER BY {nc};"
                        )
                        cursor.execute(sql, (build_like_param(norm),))
                        rows = cursor.fetchall()
                        if rows:
                            return [(r[0], "" if r[1] is None else str(r[1])) for r in rows]

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
        """بارگذاری اولیه کاربران - حفظ ساختار کوئری از 1.2.6"""
        cursor = self.conn.cursor()
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
