# -*- coding: utf-8 -*-
import sys, subprocess, re, pyodbc
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QPushButton,
    QLineEdit, QFormLayout, QMessageBox, QDesktopWidget
)
from PyQt5.QtCore import Qt, QTimer

# ------------------------------ تابع کمکی برای مرکز‌چین کردن پنجره ------------------------------
def center_window(widget):
    """پنجره را وسط صفحه قرار می‌دهد."""
    qr = widget.frameGeometry()
    cp = QDesktopWidget().availableGeometry().center()
    qr.moveCenter(cp)
    widget.move(qr.topLeft())

# ------------------------------ جستجوی SQL Instanceهای سیستم ------------------------------
def find_sql_instances():
    """بررسی رجیستری برای یافتن Instanceهای SQL نصب‌شده"""
    try:
        cmd = (
            'powershell "Get-ChildItem '
            '\'HKLM:\\SOFTWARE\\Microsoft\\Microsoft SQL Server\\Instance Names\\SQL\' '
            '| ForEach-Object { $_.Name.Split(\'\\\\\')[-1] }"'
        )
        output = subprocess.check_output(cmd, shell=True, text=True)
        instances = [inst.strip() for inst in output.splitlines() if inst.strip()]
        return instances
    except Exception:
        return []

# ------------------------------ جستجوی آخرین دیتابیس Moein ------------------------------
def find_latest_moein_db(server_name, username="sa", password="arta0@"):
    """در هر Instance دنبال آخرین دیتابیس Moein می‌گردد"""
    try:
        conn_str = f"DRIVER={{SQL Server}};SERVER={server_name};UID={username};PWD={password}"
        conn = pyodbc.connect(conn_str, timeout=3)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sys.databases WHERE name LIKE 'Moein%' ORDER BY name ASC;")
        dbs = [r[0] for r in cursor.fetchall()]
        if not dbs:
            return None
        # آخرین دیتابیس بر اساس عدد انتهای نام (مثل Moein20)
        dbs_sorted = sorted(dbs, key=lambda x: int(re.findall(r'\d+', x)[0]) if re.findall(r'\d+', x) else 0)
        return dbs_sorted[-1]
    except Exception:
        return None

# ------------------------------ تلاش برای اتصال خودکار ------------------------------
def auto_connect():
    """تلاش برای یافتن و اتصال خودکار به آخرین دیتابیس Moein"""
    possible_instances = [
        r".\Moein",
        r".\Moein2008",
        r".\Moein2012",
        r".\Moein2014",
        r".\Moein2019",
        r".\Moein2022",
    ]
    # اضافه کردن Instanceهای واقعی سیستم
    possible_instances.extend(find_sql_instances())

    for instance in possible_instances:
        db = find_latest_moein_db(instance)
        if db:
            try:
                conn_str = f"DRIVER={{SQL Server}};SERVER={instance};DATABASE={db};UID=sa;PWD=arta0@"
                conn = pyodbc.connect(conn_str, autocommit=False)
                return conn, instance, db
            except:
                continue
    return None, None, None

# ------------------------------ پنجره تنظیمات دستی ------------------------------
class ManualConnectWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("تنظیمات اتصال به SQL Server (اتصال دستی)")
        self.setGeometry(500, 250, 400, 250)
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
        """اتصال دستی به دیتابیس"""
        server = self.txt_server.text().strip()
        db = self.txt_db.text().strip()
        user = self.txt_user.text().strip()
        pwd = self.txt_pass.text().strip()

        try:
            conn_str = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={db};UID={user};PWD={pwd}"
            conn = pyodbc.connect(conn_str, autocommit=False)

            # بررسی ستون IsActive
            cursor = conn.cursor()
            cursor.execute("""
                SELECT COLUMN_NAME
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'UserAccess' AND COLUMN_NAME = 'IsActive'
            """)
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

# ------------------------------ پنجره اتصال خودکار ------------------------------
class AutoConnectWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("اتصال خودکار به SQL Server")
        self.setGeometry(500, 300, 350, 100)
        center_window(self)

        layout = QVBoxLayout()
        self.label = QLabel("در حال تلاش برای اتصال خودکار به دیتابیس Moein ...")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)
        self.setLayout(layout)

        # اجرای تابع اتصال با تاخیر برای جلوگیری از فریز شدن UI
        QTimer.singleShot(100, self.try_auto_connect)

    def try_auto_connect(self):
        conn, server, db = auto_connect()
        if conn:
            QMessageBox.information(self, "موفقیت ✅", f"اتصال خودکار برقرار شد:\n{server} → {db}")
            self.close()
            self.main_window = MainWindow(conn)
            self.main_window.show()
        else:
            QMessageBox.warning(self, "اتصال خودکار ناموفق ⚠️", "هیچ دیتابیس Moein پیدا نشد.\nلطفاً اتصال دستی را انجام دهید.")
            self.close()
            self.manual = ManualConnectWindow()
            self.manual.show()

# ------------------------------ پنجره اصلی برنامه ------------------------------
class MainWindow(QWidget):
    def __init__(self, connection):
        super().__init__()
        self.conn = connection
        self.setWindowTitle("مدیریت سطح دسترسی کاربران")
        self.setGeometry(300, 200, 600, 400)
        center_window(self)

        layout = QVBoxLayout()
        label = QLabel("✅ اتصال برقرار شد، در این بخش تنظیمات دسترسی نمایش داده می‌شود.")
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)

        self.setLayout(layout)

# ------------------------------ اجرای برنامه ------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = AutoConnectWindow()
    win.show()
    sys.exit(app.exec_())
