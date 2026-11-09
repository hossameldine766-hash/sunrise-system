import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from PIL import Image, ImageTk
import os
import shutil
import subprocess
import sys
from datetime import datetime
from tkcalendar import DateEntry
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl import Workbook, load_workbook
def close_splash():
    """إغلاق شاشة التحميل (تعمل فقط في EXE المبني بـ --splash)"""
    try:
        import pyi_splash
        if pyi_splash.is_alive():
            pyi_splash.close()
    except ImportError:
        pass

# ═══════════════════════════════════════════════════════════════════
# ثوابت التصميم الموحد
# ═══════════════════════════════════════════════════════════════════
import sys
import os

def resource_path(relative_path):
    """الحصول على المسار الصحيح للملفات داخل EXE أو وضع التطوير"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


# ─── الشعار والإعلان ───
COMPANY_LOGO_PATH = resource_path("m.png")

COMPANY_AD = ("🏢 شركة الزهراء للبرمجيات - المحاسب/حسام الدين سيد ربيع - "
              "أنظمة إدارة متكاملة | للاستفسار: 01003796984 📞 ")
AD_BG_COLOR = "#2c3e50"
AD_FG_COLOR = "#f1c40f"
AD_SPEED = 80
AD_DISPLAY_LENGTH = 70

BANNER_AD_LABELS = []

# ─── القائمة الجانبية ───
SIDEBAR_BG        = '#154360'
SIDEBAR_BTN_BG    = '#2471a3'
SIDEBAR_BTN_HOVER = '#f1c40f'
SIDEBAR_FG        = 'white'
SIDEBAR_ACCENT    = '#f1c40f'



def init_db():
    conn = sqlite3.connect('sunrise.db')
    c = conn.cursor()

    # جدول المبيعات
    c.execute('''
    CREATE TABLE IF NOT EXISTS sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL,
        place TEXT,
        cost INTEGER,
        euro INTEGER,
        dollar INTEGER,
        pound INTEGER,
        visa INTEGER,
        image_path TEXT,
        profit REAL DEFAULT 0
    )
    ''')

    # جدول المصاريف
    c.execute('''
    CREATE TABLE IF NOT EXISTS paids (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL,
        description TEXT NOT NULL,
        value INTEGER NOT NULL,
        department TEXT NOT NULL,
        place TEXT NOT NULL,
        source TEXT NOT NULL,
        sender TEXT NOT NULL,
        image_path TEXT,
        euro REAL DEFAULT 0,
        usd REAL DEFAULT 0,
        invoice_id INTEGER,
        payment_type TEXT

    )
    ''')

    # جدول الأرصدة المرحلة
    c.execute('''
    CREATE TABLE IF NOT EXISTS carried_balances (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        month_year TEXT NOT NULL,
        place TEXT NOT NULL,
        euro REAL DEFAULT 0,
        usd REAL DEFAULT 0,
        egp REAL DEFAULT 0,
        euro_rate REAL DEFAULT 51.0,
        usd_rate REAL DEFAULT 48.0,
        UNIQUE(month_year, place)
    )
    ''')

    # جدول الخزينة العامة
    c.execute('''
    CREATE TABLE IF NOT EXISTS general_treasury (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        payment_date TEXT NOT NULL,
        euro REAL DEFAULT 0,
        usd REAL DEFAULT 0,
        egp REAL DEFAULT 0,
        place TEXT
    )
    ''')

    # جدول المصاريف الدورية
    c.execute('''
    CREATE TABLE IF NOT EXISTS recurring_expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        description TEXT NOT NULL,
        value REAL DEFAULT 0,
        department TEXT
    )
    ''')

    # جدول ترحيل الأرباح / الصافي
    c.execute('''
    CREATE TABLE IF NOT EXISTS profit_transfers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        month_year TEXT NOT NULL UNIQUE,
        net_amount REAL DEFAULT 0.0,
        total_additional_expenses REAL DEFAULT 0.0,
        remaining_balance REAL DEFAULT 0.0,
        notes TEXT,
        transfer_date TEXT
    )
    ''')

    # جدول المصاريف الإضافية
    c.execute('''
    CREATE TABLE IF NOT EXISTS additional_expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        expense_date TEXT NOT NULL,
        description TEXT NOT NULL,
        amount REAL DEFAULT 0,
        image_path TEXT
    )
    ''')
    # جدول المبالغ المسلمة من صاحب المكان
    # ═══════════════════════════════════════════════════════════════════
    c.execute('''
    CREATE TABLE IF NOT EXISTS owner_money (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL,
        place TEXT,
        amount REAL NOT NULL DEFAULT 0,
        transaction_type TEXT NOT NULL,
        description TEXT,
        image_path TEXT,
        created_at TEXT
    )
    ''')
    # ═══════════════════════════════════════════════════════════════════
    # جداول حسابات التجار
    # ═══════════════════════════════════════════════════════════════════

    # جدول التجار
    c.execute('''
    CREATE TABLE IF NOT EXISTS traders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        phone TEXT,
        address TEXT,
        notes TEXT,
        created_at TEXT
    )
    ''')
    # جدول فواتير التجار
    c.execute('''
    CREATE TABLE IF NOT EXISTS trader_invoices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_no TEXT,
        date TEXT NOT NULL,
        trader_id INTEGER NOT NULL,
        total REAL NOT NULL DEFAULT 0,
        notes TEXT,
        image_path TEXT,
        created_at TEXT,
        payment_type TEXT DEFAULT 'أجل',
        paid_amount REAL DEFAULT 0,
        FOREIGN KEY (trader_id) REFERENCES traders(id) ON DELETE CASCADE
    )
    ''')
    # جدول بنود الفاتورة
    c.execute('''
    CREATE TABLE IF NOT EXISTS trader_invoice_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_id INTEGER NOT NULL,
        item_name TEXT NOT NULL,
        quantity REAL NOT NULL DEFAULT 1,
        price REAL NOT NULL DEFAULT 0,
        total REAL NOT NULL DEFAULT 0,
        FOREIGN KEY (invoice_id) REFERENCES trader_invoices(id) ON DELETE CASCADE
    )
    ''')
    # جدول مدفوعات التجار
    c.execute('''
    CREATE TABLE IF NOT EXISTS trader_payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL,
        trader_id INTEGER NOT NULL,
        amount REAL NOT NULL DEFAULT 0,
        payment_method TEXT,
        notes TEXT,
        image_path TEXT,
        created_at TEXT,
        paid_id INTEGER,
        FOREIGN KEY (trader_id) REFERENCES traders(id) ON DELETE CASCADE
    )
    ''')


    conn.commit()
    conn.close()

# ═══════════════════════════════════════════════════════════════════
# MIGRATION: دعم العملات الأجنبية في trader_payments
# ═══════════════════════════════════════════════════════════════════
def migrate_add_currency_to_trader_payments():
    """إضافة أعمدة العملة الأجنبية لجدول trader_payments"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        c.execute("PRAGMA table_info(trader_payments)")
        cols = [r[1] for r in c.fetchall()]

        if 'currency' not in cols:
            c.execute("ALTER TABLE trader_payments ADD COLUMN currency TEXT DEFAULT 'نقدي'")
            print("🔧 إضافة 'currency' لـ trader_payments")

        if 'foreign_amount' not in cols:
            c.execute("ALTER TABLE trader_payments ADD COLUMN foreign_amount REAL DEFAULT 0")
            print("🔧 إضافة 'foreign_amount' لـ trader_payments")

        if 'exchange_rate' not in cols:
            c.execute("ALTER TABLE trader_payments ADD COLUMN exchange_rate REAL DEFAULT 0")
            print("🔧 إضافة 'exchange_rate' لـ trader_payments")

        conn.commit()
        conn.close()
        print("✅ اكتملت تعديلات trader_payments")

    except Exception as e:
        print(f"⚠️ خطأ: {e}")



# ===================== دوال العمليات على قاعدة البيانات =====================
def add_sale(date, place, cost, euro, dollar, pound, visa, image_source_path, profit):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('''
            INSERT INTO sales (date, place, cost, euro, dollar, pound, visa, image_path, profit)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?,?)
        ''', (date, place, int(cost), int(euro), int(dollar), int(pound), int(visa), None, profit))

        sale_id = c.lastrowid
        if sale_id is None:
            c.execute('SELECT last_insert_rowid()')
            sale_id = c.fetchone()[0]

        conn.commit()

        final_path = None
        if image_source_path and os.path.exists(image_source_path):
            # ✅ استخدام مجلد sales_images/
            final_path = copy_image_to_project(image_source_path, sale_id, 'sales_images/')
            if final_path:
                c.execute('UPDATE sales SET image_path = ? WHERE id = ?', (final_path, sale_id))
                conn.commit()

        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ في add_sale: {e}")
        return False


def update_sale(sale_id, date, place, cost, euro, dollar, pound, visa, new_image_path, profit):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('SELECT image_path FROM sales WHERE id = ?', (sale_id,))
        row = c.fetchone()
        old_path = clean_path(row[0]) if row else None

        new_path = clean_path(new_image_path)

        final_path = old_path

        if new_path == "REMOVE":
            if old_path and os.path.exists(old_path):
                try:
                    os.remove(old_path)
                    print(f"🗑️ تم حذف الصورة: {old_path}")
                except Exception as e:
                    print(f"⚠️ فشل حذف الصورة: {e}")
            final_path = None

        elif new_path is not None and new_path != old_path:
            if old_path and os.path.exists(old_path):
                try:
                    os.remove(old_path)
                    print(f"🗑️ تم حذف الصورة القديمة: {old_path}")
                except Exception as e:
                    print(f"⚠️ فشل حذف الصورة القديمة: {e}")
            # ✅ استخدام مجلد sales_images/
            final_path = copy_image_to_project(new_path, sale_id, 'sales_images/')
            if final_path is None:
                final_path = None

        elif new_path is not None and old_path is None:
            final_path = copy_image_to_project(new_path, sale_id, 'sales_images/')
            if final_path is None:
                final_path = None

        elif new_path is None and old_path is not None:
            final_path = old_path

        c.execute('''
            UPDATE sales
            SET date=?, place=?, cost=?, euro=?, dollar=?, pound=?, visa=?, image_path=?, profit=?
            WHERE id=?
        ''', (date, place, int(cost), int(euro), int(dollar), int(pound), int(visa), final_path, profit, sale_id))

        affected = c.rowcount
        conn.commit()
        conn.close()
        return affected > 0

    except Exception as e:
        print(f"❌ خطأ في update_sale: {e}")
        return False


def delete_sale(sale_id):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('SELECT image_path FROM sales WHERE id = ?', (sale_id,))
        row = c.fetchone()
        image_path = clean_path(row[0]) if row else None

        c.execute('DELETE FROM sales WHERE id = ?', (sale_id,))
        affected = c.rowcount
        conn.commit()
        conn.close()

        if affected > 0:
            if image_path and os.path.exists(image_path):
                try:
                    os.remove(image_path)
                    print(f"🗑️ تم حذف الصورة: {image_path}")
                except Exception as e:
                    print(f"⚠️ فشل حذف الصورة: {e}")
            return True
        else:
            print(f"⚠️ لم يتم العثور على صف برقم {sale_id}")
            return False

    except Exception as e:
        print(f"❌ خطأ في delete_sale: {e}")
        return False


def get_all_sales():
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM sales ORDER BY id DESC')
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ في get_all_sales: {e}")
        return []


def search_sales(search_term=None, date=None, place=None):
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        query = "SELECT * FROM sales WHERE 1=1"
        params = []

        if search_term and search_term.strip():
            query += " AND (date LIKE ? OR place LIKE ? OR cost LIKE ? OR euro LIKE ? OR dollar LIKE ? OR pound LIKE ? OR visa LIKE ?)"
            term = f"%{search_term.strip()}%"
            params.extend([term, term, term, term, term, term, term])
        if date and date.strip():
            query += " AND date = ?"
            params.append(date.strip())
        if place and place.strip() and place != "الكل":
            query += " AND place = ?"
            params.append(place.strip())

        query += " ORDER BY id "
        c.execute(query, params)
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ في search_sales: {e}")
        return []



# =====================  جدول المصاريف دوال العمليات على قاعدة البيانات =====================



def copy_image_to_project(source_path, record_id, target_dir='bill_images/'):
    """نسخ الصورة إلى مجلد محدد (افتراضي bill_images)"""
    src = clean_path(source_path)
    if not src or not os.path.exists(src):
        return None

    if record_id is None:
        return None

    os.makedirs(target_dir, exist_ok=True)

    ext = os.path.splitext(src)[1]
    if not ext:
        return None

    target_filename = f"{record_id}{ext}"
    target_path = os.path.join(target_dir, target_filename)

    try:
        shutil.copy2(src, target_path)
        print(f"✅ تم نسخ الصورة إلى: {target_path}")
        return target_path
    except Exception as e:
        print(f"❌ فشل نسخ الصورة: {e}")
        return None


def clean_path(path):
    """تحويل أي قيمة إلى مسار نصي صالح أو None"""
    if path is None:
        return None
    if isinstance(path, tuple):
        path = path[0] if path else None
    if path is not None:
        path = str(path).strip()
        return path if path else None
    return None


def safe_float(value, default=0):
    """
    تحويل أي قيمة إلى float مع التعامل مع الأخطاء.
    - إذا كانت القيمة None أو نصاً غير رقمي، تُرجع القيمة الافتراضية (0.0).
    - إذا كانت القيمة رقمًا (int أو float) تُرجعها كـ float.
    """
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        value = value.strip()
        if value == '':
            return default
        try:
            return float(value)
        except ValueError:
            return default
    return default


def add_paid(date, description, value, department, place, source, sender, image_source_path, euro, usd):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('''
            INSERT INTO paids (date, description, value, department, place, source, sender, image_path,euro,usd)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?,?,?)
        ''', (date, description, value, department, place, source, sender, None, euro, usd))

        paid_id = c.lastrowid
        if paid_id is None:
            c.execute('SELECT last_insert_rowid()')
            paid_id = c.fetchone()[0]

        conn.commit()

        final_path = None
        if image_source_path and os.path.exists(image_source_path):
            final_path = copy_image_to_project(image_source_path, paid_id)  # نفس دالة نسخ الصور
            if final_path:
                c.execute('UPDATE paids SET image_path = ? WHERE id = ?', (final_path, paid_id))
                conn.commit()

        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ في add_paid: {e}")
        return False


def update_paid(paid_id, date, description, value, department, place, source, sender, new_image_path, euro, usd):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('SELECT image_path FROM paids WHERE id = ?', (paid_id,))
        row = c.fetchone()
        old_image_path = clean_path(row[0]) if row else None

        new_path = clean_path(new_image_path)

        final_path = old_image_path  # افتراضي

        # 1. إذا كانت "REMOVE" -> حذف الصورة القديمة
        if new_path == "REMOVE":
            if old_image_path and os.path.exists(old_image_path):
                try:
                    os.remove(old_image_path)
                    print(f"🗑️ تم حذف الصورة القديمة: {old_image_path}")
                except Exception as e:
                    print(f"⚠️ فشل حذف الصورة: {e}")
            final_path = None

        # 2. إذا كانت صورة جديدة مختلفة عن القديمة
        elif new_path is not None and new_path != old_image_path:
            if old_image_path and os.path.exists(old_image_path):
                try:
                    os.remove(old_image_path)
                    print(f"🗑️ تم حذف الصورة القديمة: {old_image_path}")
                except Exception as e:
                    print(f"⚠️ فشل حذف الصورة القديمة: {e}")
            # نسخ الصورة الجديدة
            final_path = copy_image_to_project(new_path, paid_id)
            if final_path is None:
                final_path = None

        # 3. إذا كانت صورة جديدة والقديمة None -> إضافة صورة
        elif new_path is not None and old_image_path is None:
            final_path = copy_image_to_project(new_path, paid_id)
            if final_path is None:
                final_path = None

        # 4. إذا new_path هو None والقديمة موجودة -> احتفظ بالقديمة
        elif new_path is None and old_image_path is not None:
            final_path = old_image_path  # لا تغيير

        # تحديث قاعدة البيانات
        c.execute('''
            UPDATE paids
            SET date=?, description=?, value=?, department=?, place=?, source=?, sender=?, image_path=?,euro=?,usd=?
            WHERE id=?
        ''', (date, description, value, department, place, source, sender, final_path, euro, usd, paid_id))

        affected = c.rowcount
        conn.commit()
        conn.close()
        return affected > 0

    except Exception as e:
        print(f"❌ خطأ في update_paid: {e}")
        return False





def delete_paid(paid_id):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        # جلب مسار الصورة قبل الحذف
        c.execute('SELECT image_path FROM paids WHERE id = ?', (paid_id,))
        row = c.fetchone()
        image_path = clean_path(row[0]) if row else None

        # حذف السجل من قاعدة البيانات
        c.execute('DELETE FROM paids WHERE id = ?', (paid_id,))
        affected = c.rowcount
        conn.commit()
        conn.close()

        if affected > 0:
            # محاولة حذف الملف (إذا فشل، نطبع تحذيراً ولكن نعتبر العملية ناجحة)
            if image_path and os.path.exists(image_path):
                try:
                    os.remove(image_path)
                    print(f"🗑️ تم حذف الصورة: {image_path}")
                except Exception as e:
                    print(f"⚠️ فشل حذف الصورة (لكن تم حذف السجل): {e}")
            return True
        else:
            print(f"⚠️ لم يتم العثور على مصروف برقم {paid_id}")
            return False

    except Exception as e:
        print(f"❌ خطأ في delete_paid: {e}")
        return False




def get_all_paids():
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM paids ORDER BY date DESC, id DESC')
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ في get_all_paids: {e}")
        return []


def search_paids(search_term=None, start_date=None, end_date=None, department=None, place=None, source=None):
    try:
        conn = sqlite3.connect('sunrise.db')  # استخدم DB_NAME أو 'sunrise.db'
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        query = "SELECT * FROM paids WHERE 1=1"
        params = []

        if search_term and search_term.strip():
            query += " AND (description LIKE ? OR department LIKE ? OR source LIKE ? OR sender LIKE ? OR value LIKE ? OR place LIKE ? OR date LIKE ? OR euro LIKE ? OR usd LIKE ?)"
            term = f"%{search_term.strip()}%"
            params.extend([term, term, term, term, term, term, term, term, term])

        if start_date and start_date.strip():
            query += " AND date >= ?"
            params.append(start_date.strip())
        if end_date and end_date.strip():
            query += " AND date <= ?"
            params.append(end_date.strip())
        if department and department.strip() and department != "الكل":
            query += " AND department = ?"
            params.append(department.strip())
        if place and place.strip() and place != "الكل":
            query += " AND place = ?"
            params.append(place.strip())
        if source and source.strip() and source != "الكل":
            query += " AND source = ?"
            params.append(source.strip())

        query += " ORDER BY date "
        c.execute(query, params)
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ في search_paids: {e}")
        return []





################  دوال المبيعات اليومية##################
def get_sales_data(start_date, end_date, place='الكل'):
    """
    جلب المبيعات اليومية مجمعة حسب التاريخ والمكان.
    ترجع قائمة من القواميس: [{'sale_date': '...', 'place': '...', 'euro': 0, ...}]
    """
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()

        query = '''
            SELECT 
                date,
                place,
                SUM(euro) AS euro,
                SUM(dollar) AS usd,
                SUM(pound) AS egp,
                SUM(visa) AS visa,
                SUM(cost) AS cost
            FROM sales
            WHERE date >= ? AND date <= ?
        '''
        params = [start_date, end_date]

        if place and place != 'الكل':
            query += " AND place = ?"
            params.append(place)

        query += " GROUP BY date, place ORDER BY date, place"

        c.execute(query, params)
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ في جلب بيانات المبيعات: {e}")
        return []


def get_sales_totals(start_date, end_date, place='الكل'):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        query = '''
            SELECT 
                SUM(euro) AS total_euro,
                SUM(dollar) AS total_usd,
                SUM(pound) AS total_egp,
                SUM(visa) AS total_visa,
                SUM(cost) AS total_cost
            FROM sales
            WHERE date >= ? AND date <= ?
        '''
        params = [start_date, end_date]
        if place and place != 'الكل':
            query += " AND place = ?"
            params.append(place)
        c.execute(query, params)
        row = c.fetchone()
        conn.close()
        if row:
            return {
                'euro': row[0] or 0,
                'usd': row[1] or 0,
                'egp': row[2] or 0,
                'visa': row[3] or 0,
                'cost': row[4] or 0
            }
        return {}
    except Exception as e:
        print(f"خطأ في حساب الإجماليات: {e}")
        return {}




#######تقرير الحساب الختامي#######
def get_sales_by_place(start_date, end_date):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        query = '''
            SELECT 
                place,
                SUM(euro) AS euro,
                SUM(dollar) AS usd,
                SUM(pound) AS egp,
                SUM(visa) AS visa
            FROM sales
            WHERE date >= ? AND date <= ?
            GROUP BY place
        '''
        c.execute(query, (start_date, end_date))
        rows = c.fetchall()
        conn.close()

        result = {}
        for place, euro, usd, egp, visa in rows:
            key = place.lower() if place else 'other'
            result[key] = {'euro': euro or 0.0, 'usd': usd or 0.0, 'egp': egp or 0.0, 'visa': visa or 0.0}
        return result
    except Exception as e:
        print(f"خطأ في جلب مبيعات الأماكن: {e}")
        return {}


def get_expenses_by_department(start_date, end_date, place=None):
    """
    ترجع إجمالي المصاريف مقسمة حسب القسم مع تفصيل العملات.
    """
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        query = '''
            SELECT 
                department,
                SUM(value) AS total_egp,
                SUM(euro) AS total_euro,
                SUM(usd) AS total_usd
            FROM paids
            WHERE date >= ? AND date <= ?
        '''
        params = [start_date, end_date]
        if place and place != 'الكل':
            query += " AND place = ?"
            params.append(place)
        query += " GROUP BY department"

        print(f"استعلام المصاريف: {query}")
        print(f"المعاملات: {params}")

        c.execute(query, params)
        rows = c.fetchall()
        conn.close()

        print("نتائج المصاريف الخام:", rows)  # طباعة لترى البيانات

        # تهيئة النتيجة
        result = {
            'purchases': {'egp': 0.0, 'euro': 0.0, 'usd': 0.0},
            'expenses': {'egp': 0.0, 'euro': 0.0, 'usd': 0.0},
            'other': {'egp': 0.0, 'euro': 0.0, 'usd': 0.0}
        }
        for dept, egp, euro, usd in rows:
            dept_lower = dept.lower() if dept else ''
            if 'مشتريات' in dept_lower or dept_lower == 'purchases':
                result['purchases'] = {'egp': egp or 0.0, 'euro': euro or 0.0, 'usd': usd or 0.0}
            elif 'مصاريف' in dept_lower or dept_lower == 'expenses':
                result['expenses'] = {'egp': egp or 0.0, 'euro': euro or 0.0, 'usd': usd or 0.0}
            else:
                # تجميع الأقسام الأخرى
                result['other']['egp'] += egp or 0.0
                result['other']['euro'] += euro or 0.0
                result['other']['usd'] += usd or 0.0

        print("نتيجة الدالة النهائية:", result)
        return result
    except Exception as e:
        print(f"خطأ في جلب المصاريف: {e}")
        return {'purchases': {'egp': 0, 'euro': 0, 'usd': 0},
                'expenses': {'egp': 0, 'euro': 0, 'usd': 0},
                'other': {'egp': 0, 'euro': 0, 'usd': 0}}


####دوال الرصيد المرحل########3
def get_carried_balance(month_year, place='الكل'):
    """
    جلب رصيد شهر ومكان معينين (بدون فيزا).
    """
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        if place == 'الكل':
            c.execute(
                'SELECT month_year, place, euro, usd, egp, euro_rate, usd_rate FROM carried_balances WHERE month_year = ?',
                (month_year,))
        else:
            c.execute(
                'SELECT month_year, place, euro, usd, egp, euro_rate, usd_rate FROM carried_balances WHERE month_year = ? AND place = ?',
                (month_year, place))
        rows = c.fetchall()
        conn.close()
        if rows:
            if place == 'الكل':
                result = {}
                for row in rows:
                    d = dict(row)
                    result[d['place']] = {
                        'euro': d['euro'], 'usd': d['usd'], 'egp': d['egp'],
                        'euro_rate': d['euro_rate'], 'usd_rate': d['usd_rate']
                    }
                return result
            else:
                d = dict(rows[0])
                return {
                    'euro': d['euro'], 'usd': d['usd'], 'egp': d['egp'],
                    'euro_rate': d['euro_rate'], 'usd_rate': d['usd_rate']
                }
        else:
            if place == 'الكل':
                return {}
            else:
                return {'euro': 0.0, 'usd': 0.0, 'egp': 0.0, 'euro_rate': 51.0, 'usd_rate': 48.0}
    except Exception as e:
        print(f"خطأ: {e}")
        return {} if place == 'الكل' else {'euro': 0.0, 'usd': 0.0, 'egp': 0.0, 'euro_rate': 51.0, 'usd_rate': 48.0}


def save_carried_balance(month_year, place, euro, usd, egp, euro_rate=51.0, usd_rate=48.0):
    """حفظ رصيد شهر ومكان (بدون فيزا)"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('''
            INSERT INTO carried_balances (month_year, place, euro, usd, egp, euro_rate, usd_rate)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(month_year, place) DO UPDATE SET
                euro = excluded.euro,
                usd = excluded.usd,
                egp = excluded.egp,
                euro_rate = excluded.euro_rate,
                usd_rate = excluded.usd_rate
        ''', (month_year, place, euro, usd, egp, euro_rate, usd_rate))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return False


def get_next_month(month_year):
    from datetime import datetime, timedelta
    dt = datetime.strptime(month_year, '%Y-%m')
    next_dt = dt.replace(day=1) + timedelta(days=32)
    next_dt = next_dt.replace(day=1)
    return next_dt.strftime('%Y-%m')


def get_total_carried_balance(month_year, place='الكل'):
    """ترجع مجموع أرصدة مكان معين (بدون فيزا)"""
    balances = get_carried_balance(month_year, place)
    if place == 'الكل':
        total = {'euro': 0.0, 'usd': 0.0, 'egp': 0.0}
        if balances:
            for p, data in balances.items():
                total['euro'] += data['euro']
                total['usd'] += data['usd']
                total['egp'] += data['egp']
            first = next(iter(balances.values()))
            total['euro_rate'] = first.get('euro_rate', 51.0)
            total['usd_rate'] = first.get('usd_rate', 48.0)
        else:
            total['euro_rate'] = 51.0
            total['usd_rate'] = 48.0
        return total
    else:
        return balances if balances else {
            'euro': 0.0, 'usd': 0.0, 'egp': 0.0,
            'euro_rate': 51.0, 'usd_rate': 48.0
        }


#######دالة التصدير الي الاكسيل المساعده###########


MASTER_FILE = "Master_Report.xlsx"
MASTER_SHEET = "كل التقارير"


def append_report_to_master(data, report_title, headers, row_colors=None):
    """
    إضافة تقرير جديد إلى الملف الرئيسي بشكل آمن.
    لا تعدل أي شيء في هذه الدالة.
    """
    try:
        # 1. التحقق من صحة البيانات
        if not data:
            print("❌ البيانات فارغة")
            return False
        if not headers:
            print("❌ العناوين فارغة")
            return False

        # 2. فتح أو إنشاء الملف
        if os.path.exists(MASTER_FILE):
            try:
                wb = load_workbook(MASTER_FILE)
            except Exception as e:
                print(f"⚠️ الملف تالف، سيتم إنشاء ملف جديد: {e}")
                wb = Workbook()
                wb.remove(wb.active)
        else:
            wb = Workbook()
            wb.remove(wb.active)

        # 3. التأكد من وجود الصفحة الرئيسية
        if MASTER_SHEET not in wb.sheetnames:
            ws = wb.create_sheet(title=MASTER_SHEET)
            ws.sheet_view.rightToLeft = True
        else:
            ws = wb[MASTER_SHEET]

        # 4. تحديد الصف التالي
        next_row = ws.max_row + 1

        # 5. إضافة عنوان التقرير
        title_cell = ws.cell(row=next_row, column=1, value=f"📊 {report_title}")
        title_cell.font = Font(bold=True, size=14)
        title_cell.alignment = Alignment(horizontal='right', vertical='center')
        next_row += 2

        # 6. كتابة العناوين
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=next_row, column=col_idx, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
            cell.alignment = Alignment(horizontal='center', vertical='center')
        next_row += 1

        # 7. كتابة البيانات (مع ضبط عدد الأعمدة تلقائياً)
        expected_cols = len(headers)
        for row_idx, row_data in enumerate(data):
            # تحويل إلى قائمة
            row_list = list(row_data)

            # ضبط الطول ليتطابق مع عدد العناوين
            if len(row_list) < expected_cols:
                row_list.extend([''] * (expected_cols - len(row_list)))
            elif len(row_list) > expected_cols:
                row_list = row_list[:expected_cols]

            # كتابة الخلايا
            for col_idx, value in enumerate(row_list, start=1):
                # معالجة القيم
                if value is None:
                    value = ''
                elif isinstance(value, (int, float)):
                    pass  # نتركها كما هي
                else:
                    value = str(value)
                cell = ws.cell(row=next_row, column=col_idx, value=value)
                cell.alignment = Alignment(horizontal='right', vertical='center')

            next_row += 1

        # 8. سطر فارغ بعد التقرير
        next_row += 1

        # 9. ضبط عرض الأعمدة (نسخة آمنة)
        for col in ws.columns:
            max_length = 0
            # أول خلية في العمود تعطينا الحرف
            if col and len(col) > 0:
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                # ضبط العرض بحد أقصى 35
                ws.column_dimensions[col_letter].width = min(max_length + 3, 35)

        # 10. حفظ الملف
        wb.save(MASTER_FILE)
        print(f"✅ تم إضافة التقرير بنجاح: {report_title}")
        return True

    except Exception as e:
        print(f"❌ خطأ فادح في append_report_to_master: {e}")
        import traceback
        traceback.print_exc()
        return False

    # ========== دوال الخزينة العامة ==========


def add_treasury_payment(payment_date, place, euro, usd, egp):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('''
            INSERT INTO general_treasury (payment_date, place, euro, usd, egp)
            VALUES (?, ?, ?, ?, ?)
        ''', (payment_date, place, euro, usd, egp))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"خطأ في إضافة مبلغ للخزينة: {e}")
        return False


def get_all_treasury_payments():
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM general_treasury ORDER BY payment_date DESC')
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ: {e}")
        return []


def search_treasury_payments(start_date, end_date, place='الكل'):
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        query = '''
            SELECT * FROM general_treasury
            WHERE payment_date >= ? AND payment_date <= ?
        '''
        params = [start_date, end_date]
        if place and place != 'الكل':
            query += " AND place = ?"
            params.append(place)
        query += " ORDER BY payment_date DESC"
        c.execute(query, params)
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ في البحث: {e}")
        return []


def update_treasury_payment(payment_id, payment_date, place, euro, usd, egp):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('''
            UPDATE general_treasury
            SET payment_date = ?, place = ?, euro = ?, usd = ?, egp = ?
            WHERE id = ?
        ''', (payment_date, place, euro, usd, egp, payment_id))
        affected = c.rowcount
        conn.commit()
        conn.close()
        return affected > 0
    except Exception as e:
        print(f"خطأ في التعديل: {e}")
        return False


def delete_treasury_payment(payment_id):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('DELETE FROM general_treasury WHERE id = ?', (payment_id,))
        affected = c.rowcount
        conn.commit()
        conn.close()
        return affected > 0
    except Exception as e:
        print(f"خطأ في الحذف: {e}")
        return False


def get_treasury_totals(start_date=None, end_date=None, place='الكل'):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        query = 'SELECT SUM(euro), SUM(usd), SUM(egp) FROM general_treasury WHERE 1=1'
        params = []
        if start_date and end_date:
            query += " AND payment_date >= ? AND payment_date <= ?"
            params.extend([start_date, end_date])
        if place and place != 'الكل':
            query += " AND place = ?"
            params.append(place)
        c.execute(query, params)
        row = c.fetchone()
        conn.close()
        if row:
            return {'euro': row[0] or 0.0, 'usd': row[1] or 0.0, 'egp': row[2] or 0.0}
        return {'euro': 0.0, 'usd': 0.0, 'egp': 0.0}
    except Exception as e:
        print(f"خطأ: {e}")
        return {'euro': 0.0, 'usd': 0.0, 'egp': 0.0}


######### ######### cash balance  #################


def get_sales_totals_by_place(start_date, end_date, place='الكل'):
    """
    ترجع إجمالي المبيعات (يورو، دولار، جنيه، فيزا) حسب الفترة والمكان.
    """
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        query = '''
            SELECT 
                SUM(euro) AS total_euro,
                SUM(dollar) AS total_usd,
                SUM(pound) AS total_egp
            FROM sales
            WHERE date >= ? AND date <= ?
        '''
        params = [start_date, end_date]
        if place and place != 'الكل':
            query += " AND place = ?"
            params.append(place)
        c.execute(query, params)
        row = c.fetchone()
        conn.close()
        if row:
            return {
                'euro': row[0] or 0,
                'dollar': row[1] or 0,
                'pound': row[2] or 0
            }
        return {'euro': 0, 'usd': 0, 'egp': 0, 'visa': 0}
    except Exception as e:
        print(f"خطأ في جلب إجمالي المبيعات: {e}")
        return {'euro': 0, 'usd': 0, 'egp': 0, 'visa': 0}


def get_treasury_totals_by_place(start_date, end_date, place='الكل'):
    """
    ترجع إجمالي المبالغ المسلمة للخزينة (يورو، دولار، جنيه) حسب الفترة والمكان.
    """
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        query = '''
            SELECT 
                SUM(euro) AS total_euro,
                SUM(usd) AS total_usd,
                SUM(egp) AS total_egp
            FROM general_treasury
            WHERE payment_date >= ? AND payment_date <= ?
        '''
        params = [start_date, end_date]
        if place and place != 'الكل':
            query += " AND place = ?"
            params.append(place)
        c.execute(query, params)
        row = c.fetchone()
        conn.close()
        if row:
            return {
                'euro': row[0] or 0,
                'usd': row[1] or 0,
                'egp': row[2] or 0
            }
        return {'euro': 0, 'usd': 0, 'egp': 0}
    except Exception as e:
        print(f"خطأ في جلب إجمالي الخزينة: {e}")
        return {'euro': 0, 'usd': 0, 'egp': 0}


def get_shop_expenses_total(start_date, end_date, place='الكل'):
    """
    ترجع إجمالي مصاريف المحل (source = 'المحل') مقسمة حسب العملات.
    المخرجات: {'euro': 0.0, 'usd': 0.0, 'egp': 0.0}
    """
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        query = '''
            SELECT SUM(euro) AS total_euro, SUM(usd) AS total_usd, SUM(value) AS total_egp
            FROM paids
            WHERE date >= ? AND date <= ?
            AND source = 'المحل'
        '''
        params = [start_date, end_date]
        if place and place != 'الكل':
            query += " AND place = ?"
            params.append(place)
        c.execute(query, params)
        row = c.fetchone()
        conn.close()
        if row:
            return {
                'euro': row[0] or 0,
                'usd': row[1] or 0,
                'egp': row[2] or 0
            }
        return {'euro': 0, 'usd': 0, 'egp': 0}
    except Exception as e:
        print(f"خطأ في جلب مصاريف المحل: {e}")
        return {'euro': 0, 'usd': 0, 'egp': 0}



# ========== دوال المصاريف الدورية (مبسطة) ==========

def add_recurring_expense(description, value, department):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('''
            INSERT INTO recurring_expenses (description, value, department)
            VALUES (?, ?, ?)
        ''', (description, value, department))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"خطأ: {e}")
        return False


def get_all_recurring():
    """جلب جميع المصاريف الدورية (الأعمدة الأساسية فقط)"""
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        # نقرأ فقط الأعمدة التي نريدها، حتى لو كانت هناك أعمدة إضافية في الجدول
        c.execute('SELECT id, description, value, department FROM recurring_expenses ORDER BY description')
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ: {e}")
        return []


def update_recurring_expense(exp_id, description, value, department):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('''
            UPDATE recurring_expenses
            SET description=?, value=?, department=?
            WHERE id=?
        ''', (description, value, department, exp_id))
        affected = c.rowcount
        conn.commit()
        conn.close()
        return affected > 0
    except Exception as e:
        print(f"خطأ: {e}")
        return False


def delete_recurring_expense(exp_id):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('DELETE FROM recurring_expenses WHERE id = ?', (exp_id,))
        affected = c.rowcount
        conn.commit()
        conn.close()
        return affected > 0
    except Exception as e:
        print(f"خطأ: {e}")
        return False


def apply_recurring_expenses(month_year):
    """ترحيل المصاريف الدورية إلى جدول paid (بـ euro=0, usd=0)"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        # جلب المصاريف الدورية
        c.execute('SELECT description, value, department FROM recurring_expenses')
        recurring_items = c.fetchall()
        added_count = 0

        for desc, value, dept in recurring_items:
            # التحقق من عدم التكرار في نفس الشهر
            c.execute('''
                SELECT id FROM paids 
                WHERE description = ? 
                AND strftime('%Y-%m', date) = ?
            ''', (desc, month_year))
            existing = c.fetchone()

            if not existing:
                from datetime import datetime
                first_day = f"{month_year}-01"
                # إدراج في جدول paid (مع تعيين euro=0, usd=0)
                c.execute('''
                    INSERT INTO paids (date, description, value, euro, usd, department, place, source, sender, image_path)
                    VALUES (?, ?, ?, 0, 0, ?, '', '', '', '')
                ''', (first_day, desc, value, dept))
                added_count += 1

        conn.commit()
        conn.close()
        return added_count
    except Exception as e:
        print(f"❌ خطأ في ترحيل المصاريف الدورية: {e}")
        return -1


#####   دالة حساب الربح  #################


def calculate_profit(euro, usd, egp, visa, cost, euro_rate, usd_rate):
    """
    حساب الربح = (يورو * سعر_اليورو) + (دولار * سعر_الدولار) + جنيه + (فيزا * 0.96) - الكوست
    """
    euro_egp = euro * euro_rate
    usd_egp = usd * usd_rate
    visa_net = visa * 0.96
    total_revenue = euro_egp + usd_egp + egp + visa_net
    profit = total_revenue - cost
    return int(profit)

    # ========== دوال ترحيل الأرباح ==========


def add_profit_transfer(month_year, net_amount, notes=""):
    """إدراج ترحيل جديد للصافي (أو تحديثه إذا كان الشهر موجوداً)"""
    from datetime import datetime
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('''
            INSERT INTO profit_transfers (month_year, net_amount, total_additional_expenses, remaining_balance, notes, transfer_date)
            VALUES (?, ?, 0.0, ?, ?, ?)
            ON CONFLICT(month_year) DO UPDATE SET
                net_amount = excluded.net_amount,
                remaining_balance = excluded.net_amount - total_additional_expenses,
                notes = excluded.notes,
                transfer_date = excluded.transfer_date
        ''', (month_year, net_amount, net_amount, notes, datetime.today().strftime('%Y-%m-%d')))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"خطأ في إضافة ترحيل الصافي: {e}")
        return False





def get_all_profit_transfers():
    """جلب جميع ترحيلات الأرباح"""
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM profit_transfers ORDER BY month_year DESC')
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ: {e}")
        return []







# ========== دوال المصاريف الإضافية (مستقلة) ==========

def add_additional_expense(expense_date, description, amount, image_source_path):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('''
            INSERT INTO additional_expenses (expense_date, description, amount, image_path)
            VALUES (?, ?, ?, ?)
        ''', (expense_date, description, amount, None))

        expense_id = c.lastrowid
        if expense_id is None:
            c.execute('SELECT last_insert_rowid()')
            expense_id = c.fetchone()[0]

        conn.commit()

        final_path = None
        if image_source_path and os.path.exists(image_source_path):
            # ✅ استخدام مجلد sales_images/
            final_path = copy_image_to_project(image_source_path, expense_id, 'final_image/')
            if final_path:
                c.execute('UPDATE additional_expenses SET image_path = ? WHERE id = ?', (final_path, expense_id))
                conn.commit()

        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"خطأ في إضافة المصروف الإضافي: {e}")
        return False


def get_all_additional_expenses():
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM additional_expenses ORDER BY expense_date DESC')
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ: {e}")
        return []


def update_additional_expense(expense_id, expense_date, description, amount, new_image_path):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('SELECT image_path FROM additional_expenses WHERE id = ?', (expense_id,))
        row = c.fetchone()
        old_path = clean_path(row[0]) if row else None

        new_path = clean_path(new_image_path)

        final_path = old_path

        if new_path == "REMOVE":
            if old_path and os.path.exists(old_path):
                try:
                    os.remove(old_path)
                    print(f"🗑️ تم حذف الصورة: {old_path}")
                except Exception as e:
                    print(f"⚠️ فشل حذف الصورة: {e}")
            final_path = None

        elif new_path is not None and new_path != old_path:
            if old_path and os.path.exists(old_path):
                try:
                    os.remove(old_path)
                    print(f"🗑️ تم حذف الصورة القديمة: {old_path}")
                except Exception as e:
                    print(f"⚠️ فشل حذف الصورة القديمة: {e}")
            # ✅ استخدام مجلد sales_images/
            final_path = copy_image_to_project(new_path, expense_id, 'final_images/')
            if final_path is None:
                final_path = None

        elif new_path is not None and old_path is None:
            final_path = copy_image_to_project(new_path, expense_id, 'final_images/')
            if final_path is None:
                final_path = None

        elif new_path is None and old_path is not None:
            final_path = old_path

        c.execute('''
            UPDATE additional_expenses
            SET expense_date = ?, description = ?, amount = ?, image_path = ?
            WHERE id = ?
        ''', (expense_date, description, amount, final_path, expense_id))
        affected = c.rowcount
        conn.commit()
        conn.close()
        return affected > 0
    except Exception as e:
        print(f"خطأ: {e}")
        return False


def delete_additional_expense(expense_id):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('DELETE FROM additional_expenses WHERE id = ?', (expense_id,))
        affected = c.rowcount
        conn.commit()
        conn.close()
        return affected > 0
    except Exception as e:
        print(f"خطأ: {e}")
        return False


def get_total_additional_expenses():
    """إجمالي جميع المصاريف الإضافية"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('SELECT SUM(amount) FROM additional_expenses')
        total = c.fetchone()[0]
        conn.close()
        return total if total is not None else 0.0
    except Exception as e:
        print(f"خطأ: {e}")
        return 0.0


def get_total_profits():
    """إجمالي جميع الأرباح المرحلة (net_amount)"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('SELECT SUM(net_amount) FROM profit_transfers')
        total = c.fetchone()[0]
        conn.close()
        return total if total is not None else 0.0
    except Exception as e:
        print(f"خطأ: {e}")
        return 0.0

# دوال المبالغ المسلمة من صاحب المكان
# ═══════════════════════════════════════════════════════════════════
def add_owner_money(date, place, amount, transaction_type,
                     description, image_source_path):
    """إضافة حركة جديدة"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('''
            INSERT INTO owner_money
            (date, place, amount, transaction_type, description,
             image_path, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (date, place, amount, transaction_type, description, None,
              datetime.now().strftime('%Y-%m-%d %H:%M:%S')))

        trans_id = c.lastrowid
        conn.commit()

        if image_source_path and os.path.exists(image_source_path):
            final_path = copy_image_to_project(image_source_path, trans_id,
                                                'owner_money_images/')
            if final_path:
                c.execute('UPDATE owner_money SET image_path=? WHERE id=?',
                          (final_path, trans_id))
                conn.commit()

        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return False

def update_owner_money(trans_id, date, place, amount, transaction_type,
                        description, new_image_path):
    """تعديل حركة"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        c.execute('SELECT image_path FROM owner_money WHERE id=?',
                  (trans_id,))
        row = c.fetchone()
        old_path = clean_path(row[0]) if row else None
        new_path = clean_path(new_image_path)
        final_path = old_path

        if new_path == "REMOVE":
            if old_path and os.path.exists(old_path):
                try: os.remove(old_path)
                except: pass
            final_path = None
        elif new_path is not None and new_path != old_path:
            if old_path and os.path.exists(old_path):
                try: os.remove(old_path)
                except: pass
            final_path = copy_image_to_project(new_path, trans_id,
                                                'owner_money_images/')

        c.execute('''
            UPDATE owner_money
            SET date=?, place=?, amount=?, transaction_type=?,
                description=?, image_path=?
            WHERE id=?
        ''', (date, place, amount, transaction_type, description,
              final_path, trans_id))

        affected = c.rowcount
        conn.commit()
        conn.close()
        return affected > 0
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return False

def delete_owner_money(trans_id):
    """حذف حركة"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        c.execute('SELECT image_path FROM owner_money WHERE id=?',
                  (trans_id,))
        row = c.fetchone()
        image_path = clean_path(row[0]) if row else None

        c.execute('DELETE FROM owner_money WHERE id=?', (trans_id,))
        affected = c.rowcount
        conn.commit()
        conn.close()

        if affected > 0 and image_path and os.path.exists(image_path):
            try:
                os.remove(image_path)
            except:
                pass

        return affected > 0
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return False

def get_all_owner_money():
    """جلب كل الحركات"""
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM owner_money ORDER BY date DESC, id DESC')
        rows = [dict(row) for row in c.fetchall()]
        conn.close()
        return rows
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return []

def search_owner_money(search_term=None, start_date=None, end_date=None,
                        place=None, transaction_type=None):
    """بحث في الحركات"""
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()

        query = "SELECT * FROM owner_money WHERE 1=1"
        params = []

        if search_term and search_term.strip():
            query += " AND (description LIKE ? OR date LIKE ? OR amount LIKE ?)"
            term = f"%{search_term.strip()}%"
            params.extend([term, term, term])

        if start_date and start_date.strip():
            query += " AND date >= ?"
            params.append(start_date.strip())

        if end_date and end_date.strip():
            query += " AND date <= ?"
            params.append(end_date.strip())

        if place and place.strip() and place != "الكل":
            query += " AND place = ?"
            params.append(place.strip())

        if (transaction_type and transaction_type.strip()
                and transaction_type != "الكل"):
            query += " AND transaction_type = ?"
            params.append(transaction_type.strip())

        query += " ORDER BY date DESC, id DESC"

        c.execute(query, params)
        rows = [dict(row) for row in c.fetchall()]
        conn.close()
        return rows
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return []

def get_owner_money_balance(place=None):
    """حساب الرصيد الحالي (استلام − رد)"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        query = '''SELECT
            COALESCE(SUM(CASE WHEN transaction_type='استلام'
                THEN amount ELSE 0 END), 0),
            COALESCE(SUM(CASE WHEN transaction_type='رد'
                THEN amount ELSE 0 END), 0)
            FROM owner_money WHERE 1=1'''
        params = []

        if place and place != "الكل":
            query += " AND place = ?"
            params.append(place)

        c.execute(query, params)
        received, returned = c.fetchone()
        conn.close()

        received = received or 0
        returned = returned or 0

        return {
            'received': received,
            'returned': returned,
            'balance': received - returned,
        }
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return {'received': 0, 'returned': 0, 'balance': 0}

def get_owner_money_totals(start_date, end_date, place=None):
    """جلب إجمالي المبالغ المسلمة خلال فترة (للتقرير)"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        query = '''SELECT
            COALESCE(SUM(CASE WHEN transaction_type='استلام'
                THEN amount ELSE 0 END), 0),
            COALESCE(SUM(CASE WHEN transaction_type='رد'
                THEN amount ELSE 0 END), 0)
            FROM owner_money
            WHERE date >= ? AND date <= ?'''
        params = [start_date, end_date]

        if place and place != "الكل":
            query += " AND place = ?"
            params.append(place)

        c.execute(query, params)
        received, returned = c.fetchone()
        conn.close()

        received = received or 0
        returned = returned or 0

        return {
            'received': received,
            'returned': returned,
            'net': received - returned,
        }
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return {'received': 0, 'returned': 0, 'net': 0}

# ═══════════════════════════════════════════════════════════════════
# دوال التجار
# ═══════════════════════════════════════════════════════════════════
def add_trader(name, phone='', address='', notes=''):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('''INSERT INTO traders (name, phone, address, notes, created_at)
                     VALUES (?, ?, ?, ?, ?)''',
                  (name, phone, address, notes,
                   datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        conn.commit()
        conn.close()
        return True
    except sqlite3.IntegrityError:
        return False
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return False

def update_trader(trader_id, name, phone, address, notes):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('''UPDATE traders SET name=?, phone=?, address=?, notes=?
                     WHERE id=?''', (name, phone, address, notes, trader_id))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return False

def delete_trader(trader_id):
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('DELETE FROM traders WHERE id=?', (trader_id,))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return False

def get_all_traders():
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM traders ORDER BY name')
        rows = [dict(row) for row in c.fetchall()]
        conn.close()
        return rows
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return []

def get_trader_balance(trader_id):
    """
    الرصيد = (إجمالي الفواتير − المدفوع وقت الفاتورة) − الدفعات اللاحقة
    """
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        c.execute('''SELECT
                        COALESCE(SUM(total), 0),
                        COALESCE(SUM(paid_amount), 0)
                     FROM trader_invoices
                     WHERE trader_id=?''', (trader_id,))
        invoices_total, invoices_paid = c.fetchone()
        invoices_total = invoices_total or 0
        invoices_paid = invoices_paid or 0

        c.execute('''SELECT COALESCE(SUM(amount), 0)
                     FROM trader_payments
                     WHERE trader_id=?''', (trader_id,))
        payments_total = c.fetchone()[0] or 0

        conn.close()
        return (invoices_total - invoices_paid) - payments_total
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return 0


def add_trader_invoice(invoice_no, date, trader_id, total, notes,
                       items, image_source_path=None):
    """إضافة فاتورة جديدة مع بنودها"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        c.execute('''INSERT INTO trader_invoices
                     (invoice_no, date, trader_id, total, notes,
                      image_path, created_at)
                     VALUES (?, ?, ?, ?, ?, ?, ?)''',
                  (invoice_no, date, trader_id, total, notes, None,
                   datetime.now().strftime('%Y-%m-%d %H:%M:%S')))

        invoice_id = c.lastrowid

        for item in items:
            c.execute('''INSERT INTO trader_invoice_items
                         (invoice_id, item_name, quantity, price, total)
                         VALUES (?, ?, ?, ?, ?)''',
                      (invoice_id, item['name'], item['qty'],
                       item['price'], item['total']))

        conn.commit()
        conn.close()
        return invoice_id
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return None

def delete_trader_invoice(invoice_id):
    """حذف الفاتورة + حذف قيدها من paids"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        # ─── حذف القيد المرتبط من paids ───
        c.execute('DELETE FROM paids WHERE invoice_id=?', (invoice_id,))

        # ─── حذف الفاتورة (والبنود بحكم CASCADE) ───
        c.execute('DELETE FROM trader_invoices WHERE id=?', (invoice_id,))

        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return False


def get_trader_invoice(invoice_id):
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('''SELECT ti.*, t.name as trader_name
                     FROM trader_invoices ti
                     JOIN traders t ON ti.trader_id = t.id
                     WHERE ti.id=?''', (invoice_id,))
        row = c.fetchone()
        conn.close()
        return dict(row) if row else None
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return None

def get_invoice_items(invoice_id):
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM trader_invoice_items WHERE invoice_id=?',
                  (invoice_id,))
        rows = [dict(row) for row in c.fetchall()]
        conn.close()
        return rows
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return []

def add_trader_invoice(invoice_no, date, trader_id, total, notes,
                        items, payment_type='أجل', paid_amount=0,
                        place='bazar'):
    """
    إضافة فاتورة + تسجيلها تلقائياً في paids.
    payment_type: 'كاش' | 'جزئي' | 'أجل'
    paid_amount: المبلغ المدفوع فعلياً
    """
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        # ─── جلب اسم التاجر ───
        c.execute('SELECT name FROM traders WHERE id=?', (trader_id,))
        t_row = c.fetchone()
        if not t_row:
            conn.close()
            return None
        trader_name = t_row[0]

        # ═══════════════════════════════════════════════════════
        # 1) إضافة الفاتورة
        # ═══════════════════════════════════════════════════════
        c.execute('''INSERT INTO trader_invoices
                     (invoice_no, date, trader_id, total, notes,
                      image_path, created_at, payment_type, paid_amount)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                  (invoice_no, date, trader_id, total, notes, None,
                   datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                   payment_type, paid_amount))

        invoice_id = c.lastrowid

        # ─── 2) إضافة البنود ───
        for item in items:
            c.execute('''INSERT INTO trader_invoice_items
                         (invoice_id, item_name, quantity, price, total)
                         VALUES (?, ?, ?, ?, ?)''',
                      (invoice_id, item['name'], item['qty'],
                       item['price'], item['total']))

        # ═══════════════════════════════════════════════════════
        # 3) تسجيل القيد في paids
        # ═══════════════════════════════════════════════════════
        remaining = total - paid_amount

        if payment_type == 'كاش':
            desc = f"فاتورة كاش رقم {invoice_no or invoice_id} - {trader_name}"
        elif payment_type == 'جزئي':
            desc = (f"فاتورة جزئي رقم {invoice_no or invoice_id} - "
                    f"{trader_name} (مدفوع: {paid_amount:,.2f} / "
                    f"إجمالي: {total:,.2f})")
        else:  # أجل
            desc = (f"فاتورة أجل رقم {invoice_no or invoice_id} - "
                    f"{trader_name} (الإجمالي: {total:,.2f} - غير مدفوعة)")

        c.execute('''INSERT INTO paids
                     (date, description, value, department, place,
                      source, sender, image_path, euro, usd,
                      invoice_id, payment_type)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                  (date, desc, paid_amount,
                   'مشتريات', place, 'المحل', trader_name,
                   None, 0, 0, invoice_id, payment_type))

        conn.commit()
        conn.close()
        return invoice_id

    except Exception as e:
        print(f"❌ خطأ في add_trader_invoice: {e}")
        return None


def delete_trader_payment(payment_id):
    """
    حذف دفعة التاجر + حذف المصروف المرتبط بها من صفحة المصروفات.
    """
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        # ─── جلب paid_id المرتبط ───
        c.execute('SELECT paid_id FROM trader_payments WHERE id=?',
                  (payment_id,))
        row = c.fetchone()

        # ─── حذف الدفعة ───
        c.execute('DELETE FROM trader_payments WHERE id=?', (payment_id,))

        # ─── حذف المصروف المرتبط ───
        if row and row[0]:
            c.execute('DELETE FROM paids WHERE id=?', (row[0],))

        conn.commit()
        conn.close()
        return True

    except Exception as e:
        print(f"❌ خطأ في delete_trader_payment: {e}")
        return False

def get_trader_payments(trader_id=None):
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()

        if trader_id:
            c.execute('''SELECT tp.*, t.name as trader_name
                         FROM trader_payments tp
                         JOIN traders t ON tp.trader_id = t.id
                         WHERE tp.trader_id=?
                         ORDER BY tp.date DESC, tp.id DESC''', (trader_id,))
        else:
            c.execute('''SELECT tp.*, t.name as trader_name
                         FROM trader_payments tp
                         JOIN traders t ON tp.trader_id = t.id
                         ORDER BY tp.date DESC, tp.id DESC''')

        rows = [dict(row) for row in c.fetchall()]
        conn.close()
        return rows
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return []

def get_trader_statement(trader_id):
    """
    كشف حساب كامل للتاجر (فواتير + دفعات عند الفاتورة + دفعات لاحقة).
    يعرض العملة الأجنبية في الدفعات.
    """
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()

        # ═══════════════════════════════════════════════════════
        # 1) الفواتير
        # ═══════════════════════════════════════════════════════
        c.execute('''SELECT id, date, invoice_no, total as amount,
                            COALESCE(paid_amount, 0) as paid_amount,
                            COALESCE(payment_type, 'أجل') as payment_type,
                            notes
                     FROM trader_invoices
                     WHERE trader_id = ?
                     ORDER BY date, id''', (trader_id,))
        invoices = c.fetchall()

        # ═══════════════════════════════════════════════════════
        # 2) الدفعات اللاحقة (مع كل الأعمدة الجديدة)
        # ═══════════════════════════════════════════════════════
        c.execute('''SELECT id, date, amount,
                            COALESCE(payment_method, '') as method,
                            COALESCE(notes, '') as notes,
                            COALESCE(currency, 'نقدي') as currency,
                            COALESCE(foreign_amount, 0) as foreign_amount,
                            COALESCE(exchange_rate, 0) as exchange_rate
                     FROM trader_payments
                     WHERE trader_id = ?
                     ORDER BY date, id''', (trader_id,))
        payments = c.fetchall()

        conn.close()

        # ═══════════════════════════════════════════════════════
        # بناء الحركات
        # ═══════════════════════════════════════════════════════
        movements = []

        # ─── الفواتير + الدفعات المرتبطة بها ───
        for inv in invoices:
            # 1) الفاتورة
            movements.append({
                'date': inv['date'],
                'type': 'فاتورة',
                'description': (f"فاتورة {inv['payment_type']} رقم "
                                f"{inv['invoice_no'] or inv['id']}"),
                'debit': 0,
                'credit': inv['amount'] or 0,
                'sort_key': f"{inv['date']}_1_{inv['id']}_1",
            })

            # 2) الدفعة الجزئية/الكاش عند الفاتورة
            paid_at_invoice = inv['paid_amount'] or 0
            if paid_at_invoice > 0:
                movements.append({
                    'date': inv['date'],
                    'type': 'دفعة عند الفاتورة',
                    'description': (f"دفعة على فاتورة "
                                    f"{inv['invoice_no'] or inv['id']} "
                                    f"({inv['payment_type']})"),
                    'debit': paid_at_invoice,
                    'credit': 0,
                    'sort_key': f"{inv['date']}_1_{inv['id']}_2",
                })

        # ─── الدفعات اللاحقة (مع العملة) ───
        for pay in payments:
            currency = pay['currency'] or 'نقدي'
            foreign = pay['foreign_amount'] or 0
            rate = pay['exchange_rate'] or 0
            amount_egp = pay['amount'] or 0

            # ✅ بناء الوصف حسب العملة
            if currency == 'يورو' and foreign > 0 and rate > 0:
                desc = (f"💰 دفعة {foreign:,.2f} € × {rate:,.2f} "
                        f"= {amount_egp:,.2f} ج.م")
            elif currency == 'دولار' and foreign > 0 and rate > 0:
                desc = (f"💰 دفعة {foreign:,.2f} $ × {rate:,.2f} "
                        f"= {amount_egp:,.2f} ج.م")
            else:
                desc = f"💰 دفعة {amount_egp:,.2f} ج.م"

            # إضافة طريقة الدفع
            method = pay['method'] or ''
            if method and method != 'نقدي':
                desc += f" [{method}]"

            # إضافة ملاحظات
            if pay['notes']:
                desc += f" - {pay['notes']}"

            movements.append({
                'date': pay['date'],
                'type': 'دفعة',
                'description': desc,
                'debit': amount_egp,
                'credit': 0,
                'sort_key': f"{pay['date']}_2_{pay['id']}",
            })

        # ─── ترتيب الحركات ───
        movements.sort(key=lambda x: x['sort_key'])

        # ─── حساب الرصيد التراكمي ───
        balance = 0
        for m in movements:
            balance += m['credit'] - m['debit']
            m['balance'] = balance

        return movements

    except Exception as e:
        print(f"❌ خطأ في get_trader_statement: {e}")
        import traceback
        traceback.print_exc()
        return []

# ═══════════════════════════════════════════════════════════════════
# دوال تعديل الفواتير والدفعات
# ═══════════════════════════════════════════════════════════════════

def update_trader_invoice(invoice_id, invoice_no, date, trader_id,
                           total, notes, items, payment_type='أجل',
                           paid_amount=0, place='bazar'):
    """تعديل فاتورة + تحديث قيدها في paids"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        # ─── جلب اسم التاجر ───
        c.execute('SELECT name FROM traders WHERE id=?', (trader_id,))
        t_row = c.fetchone()
        trader_name = t_row[0] if t_row else ''

        # ─── 1) تحديث الفاتورة ───
        c.execute('''UPDATE trader_invoices
                     SET invoice_no=?, date=?, trader_id=?, total=?,
                         notes=?, payment_type=?, paid_amount=?
                     WHERE id=?''',
                  (invoice_no, date, trader_id, total, notes,
                   payment_type, paid_amount, invoice_id))

        # ─── 2) حذف البنود القديمة وإضافة الجديدة ───
        c.execute('DELETE FROM trader_invoice_items WHERE invoice_id=?',
                  (invoice_id,))
        for item in items:
            c.execute('''INSERT INTO trader_invoice_items
                         (invoice_id, item_name, quantity, price, total)
                         VALUES (?, ?, ?, ?, ?)''',
                      (invoice_id, item['name'], item['qty'],
                       item['price'], item['total']))

        # ─── 3) تحديث القيد في paids ───
        remaining = total - paid_amount

        if payment_type == 'كاش':
            desc = f"فاتورة كاش رقم {invoice_no or invoice_id} - {trader_name}"
        elif payment_type == 'جزئي':
            desc = (f"فاتورة جزئي رقم {invoice_no or invoice_id} - "
                    f"{trader_name} (مدفوع: {paid_amount:,.2f} / "
                    f"إجمالي: {total:,.2f})")
        else:
            desc = (f"فاتورة أجل رقم {invoice_no or invoice_id} - "
                    f"{trader_name} (الإجمالي: {total:,.2f} - غير مدفوعة)")

        # ─── هل يوجد قيد موجود؟ ───
        c.execute('SELECT id FROM paids WHERE invoice_id=?', (invoice_id,))
        existing = c.fetchone()

        if existing:
            # تحديث القيد الموجود
            c.execute('''UPDATE paids
                         SET date=?, description=?, value=?,
                             place=?, sender=?, payment_type=?
                         WHERE id=?''',
                      (date, desc, paid_amount, place, trader_name,
                       payment_type, existing[0]))
        else:
            # إنشاء قيد جديد
            c.execute('''INSERT INTO paids
                         (date, description, value, department, place,
                          source, sender, image_path, euro, usd,
                          invoice_id, payment_type)
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                      (date, desc, paid_amount, 'مشتريات', place,
                       'المحل', trader_name, None, 0, 0,
                       invoice_id, payment_type))

        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return False

def get_trader_invoice(invoice_id):
    """جلب فاتورة واحدة بكل بياناتها"""
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('''SELECT ti.*, t.name as trader_name
                     FROM trader_invoices ti
                     JOIN traders t ON ti.trader_id = t.id
                     WHERE ti.id=?''', (invoice_id,))
        row = c.fetchone()
        conn.close()
        return dict(row) if row else None
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return None

def update_trader_payment(payment_id, date, trader_id, amount,
                           payment_method, notes, place='bazar',
                           currency='نقدي', foreign_amount=0,
                           exchange_rate=0):
    """تعديل دفعة + تحديث المصروف المرتبط بها"""
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        # جلب paid_id
        c.execute('SELECT paid_id FROM trader_payments WHERE id=?',
                  (payment_id,))
        row = c.fetchone()
        paid_id = row[0] if row else None

        # جلب اسم التاجر
        c.execute('SELECT name FROM traders WHERE id=?', (trader_id,))
        t_row = c.fetchone()
        trader_name = t_row[0] if t_row else ''

        # ═══ 1) تحديث الدفعة ═══
        c.execute('''UPDATE trader_payments
                     SET date=?, trader_id=?, amount=?,
                         payment_method=?, notes=?,
                         currency=?, foreign_amount=?, exchange_rate=?
                     WHERE id=?''',
                  (date, trader_id, amount, payment_method, notes,
                   currency, foreign_amount, exchange_rate, payment_id))

        # ═══ 2) تحديد الأعمدة ═══
        if currency == 'يورو':
            paids_value = 0
            paids_euro = foreign_amount
            paids_usd = 0
        elif currency == 'دولار':
            paids_value = 0
            paids_euro = 0
            paids_usd = foreign_amount
        else:
            paids_value = amount
            paids_euro = 0
            paids_usd = 0

        # ═══ 3) بناء الوصف ═══
        description = f"دفعة للتاجر: {trader_name}"
        if currency == 'يورو' and foreign_amount > 0:
            description += f" ({foreign_amount:,.2f} € × {exchange_rate:,.2f})"
        elif currency == 'دولار' and foreign_amount > 0:
            description += f" ({foreign_amount:,.2f} $ × {exchange_rate:,.2f})"
        if notes:
            description += f" - {notes}"

        # ═══ 4) تحديث المصروف المرتبط ═══
        if paid_id:
            c.execute('''UPDATE paids
                         SET date=?, description=?, value=?,
                             place=?, sender=?, euro=?, usd=?
                         WHERE id=?''',
                      (date, description, paids_value, place,
                       trader_name, paids_euro, paids_usd, paid_id))

        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return False


# ═══════════════════════════════════════════════════════════════════
# إضافة دفعة للتاجر + تسجيلها تلقائياً في صفحة المصروفات
# ═══════════════════════════════════════════════════════════════════
def add_trader_payment(date, trader_id, amount, payment_method, notes,
                       place='bazar', currency='نقدي',
                       foreign_amount=0, exchange_rate=0):
    """
    إضافة دفعة للتاجر + تسجيلها تلقائياً كمصروف.

    - amount: المبلغ بالجنيه (المحول)
    - currency: 'نقدي' | 'يورو' | 'دولار'
    - foreign_amount: المبلغ بالعملة الأجنبية
    - exchange_rate: سعر الصرف
    """
    try:
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()

        # جلب اسم التاجر
        c.execute('SELECT name FROM traders WHERE id=?', (trader_id,))
        row = c.fetchone()
        if not row:
            conn.close()
            return (False, None)
        trader_name = row[0]

        # ═══ 1) إضافة الدفعة في trader_payments ═══
        c.execute('''INSERT INTO trader_payments
                     (date, trader_id, amount, payment_method, notes,
                      image_path, created_at, currency, foreign_amount,
                      exchange_rate)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                  (date, trader_id, amount, payment_method, notes, None,
                   datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                   currency, foreign_amount, exchange_rate))
        payment_id = c.lastrowid

        # ═══ 2) بناء الوصف ═══
        description = f"دفعة للتاجر: {trader_name}"
        if currency == 'يورو' and foreign_amount > 0:
            description += f" ({foreign_amount:,.2f} € × {exchange_rate:,.2f})"
        elif currency == 'دولار' and foreign_amount > 0:
            description += f" ({foreign_amount:,.2f} $ × {exchange_rate:,.2f})"
        if notes:
            description += f" - {notes}"

        # ═══ 3) تحديد الأعمدة حسب العملة ═══
        # ✅ نضع المبلغ في العمود المناسب فقط (لتجنب الازدواج)
        if currency == 'يورو':
            paids_value = 0
            paids_euro = foreign_amount
            paids_usd = 0
        elif currency == 'دولار':
            paids_value = 0
            paids_euro = 0
            paids_usd = foreign_amount
        else:  # نقدي
            paids_value = amount
            paids_euro = 0
            paids_usd = 0

        # ═══ 4) تسجيل المصروف في paids ═══
        c.execute('''INSERT INTO paids
                     (date, description, value, department, place,
                      source, sender, image_path, euro, usd)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                  (date, description, paids_value,
                   'مشتريات', place, 'المحل', trader_name,
                   None, paids_euro, paids_usd))
        paid_id = c.lastrowid

        # ═══ 5) ربط الدفعة بالمصروف ═══
        c.execute('UPDATE trader_payments SET paid_id=? WHERE id=?',
                  (paid_id, payment_id))

        conn.commit()
        conn.close()
        return (True, paid_id)

    except Exception as e:
        print(f"❌ خطأ في add_trader_payment: {e}")
        return (False, None)


# ═══════════════════════════════════════════════════════════════════
# جلب كل فواتير التاجر (أو كل الفواتير)
# ═══════════════════════════════════════════════════════════════════
def get_trader_invoices(trader_id=None):
    """
    جلب فواتير تاجر معين أو كل الفواتير.
    إذا كان trader_id موجوداً → فواتير هذا التاجر فقط.
    إذا كان None → كل الفواتير.
    """
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()

        if trader_id:
            c.execute('''SELECT ti.*, t.name as trader_name
                         FROM trader_invoices ti
                         JOIN traders t ON ti.trader_id = t.id
                         WHERE ti.trader_id = ?
                         ORDER BY ti.date DESC, ti.id DESC''', (trader_id,))
        else:
            c.execute('''SELECT ti.*, t.name as trader_name
                         FROM trader_invoices ti
                         JOIN traders t ON ti.trader_id = t.id
                         ORDER BY ti.date DESC, ti.id DESC''')

        rows = [dict(row) for row in c.fetchall()]
        conn.close()
        return rows
    except Exception as e:
        print(f"❌ خطأ في get_trader_invoices: {e}")
        return []

def get_trader_payment(payment_id):
    """جلب دفعة واحدة"""
    try:
        conn = sqlite3.connect('sunrise.db')
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM trader_payments WHERE id=?', (payment_id,))
        row = c.fetchone()
        conn.close()
        return dict(row) if row else None
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return None

    # ===================== واجهة المستخدم (Tkinter) =====================
# ═══════════════════════════════════════════════════════════════════
# SALE FORM WINDOW
# ═══════════════════════════════════════════════════════════════════
class SaleFormWindow:
    """نافذة إدخال / تعديل المبيعات"""

    def __init__(self, parent_page, sale_id=None):
        self.parent_page = parent_page
        self.sale_id = sale_id
        self.is_edit = sale_id is not None
        self.image_path = None
        self.image_preview = None

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = "تعديل مبيعات" if self.is_edit else "إضافة مبيعات جديدة"
        self.window.title(title)
        self.window.geometry("760x560")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ─── العنوان ───
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icon = "✏" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'),
                 bg='#2c3e50', fg='white').pack(pady=15)

        # ─── الإطار الرئيسي ───
        main = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=20)
        main.pack(fill='both', expand=True)

        # ─── التاريخ ───
        tk.Label(main, text="التاريخ: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=0, column=0, padx=5, pady=8, sticky='e')

        self.entry_date = DateEntry(main, width=18,
                                     date_pattern='yyyy-mm-dd',
                                     font=('Arial', 11),
                                     justify='center')
        self.entry_date.grid(row=0, column=1, padx=5, pady=8, sticky='w')
        self.entry_date.set_date(datetime.now())

        # ─── المكان ───
        tk.Label(main, text="المكان: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=1, column=0, padx=5, pady=8, sticky='e')

        self.place_var = tk.StringVar(value="bazar")
        self.place_combo = ttk.Combobox(main,
                                         textvariable=self.place_var,
                                         values=('bazar', 'beach', 'other'),
                                         state='readonly', width=16,
                                         font=('Arial', 11),
                                         justify='center')
        self.place_combo.grid(row=1, column=1, padx=5, pady=8, sticky='w')

        # ─── المبلغ الفعلي (Cost) ───
        tk.Label(main, text="التكلفة (Cost):",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=2, column=0, padx=5, pady=8, sticky='e')

        self.entry_cost = tk.Entry(main, width=20,
                                    font=('Arial', 12),
                                    justify='center',
                                    bg='#fadbd8')
        self.entry_cost.grid(row=2, column=1, padx=5, pady=8, sticky='w')
        self.entry_cost.insert(0, "0")

        # ─── يورو ───
        tk.Label(main, text="يورو (€):",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#2980b9', anchor='e').grid(
                 row=0, column=2, padx=5, pady=8, sticky='e')

        self.entry_euro = tk.Entry(main, width=20,
                                    font=('Arial', 12),
                                    justify='center',
                                    bg='#e8f4f8')
        self.entry_euro.grid(row=0, column=3, padx=5, pady=8, sticky='w')
        self.entry_euro.insert(0, "0")

        # ─── دولار ───
        tk.Label(main, text="دولار ($):",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#16a085', anchor='e').grid(
                 row=1, column=2, padx=5, pady=8, sticky='e')

        self.entry_dollar = tk.Entry(main, width=20,
                                      font=('Arial', 12),
                                      justify='center',
                                      bg='#e8f8f5')
        self.entry_dollar.grid(row=1, column=3, padx=5, pady=8, sticky='w')
        self.entry_dollar.insert(0, "0")

        # ─── جنيه ───
        tk.Label(main, text="جنيه (L.E):",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#27ae60', anchor='e').grid(
                 row=2, column=2, padx=5, pady=8, sticky='e')

        self.entry_pound = tk.Entry(main, width=20,
                                     font=('Arial', 12),
                                     justify='center',
                                     bg='#eafaf1')
        self.entry_pound.grid(row=2, column=3, padx=5, pady=8, sticky='w')
        self.entry_pound.insert(0, "0")

        # ─── فيزا ───
        tk.Label(main, text="فيزا (Visa):",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#8e44ad', anchor='e').grid(
                 row=3, column=0, padx=5, pady=8, sticky='e')

        self.entry_visa = tk.Entry(main, width=20,
                                    font=('Arial', 12),
                                    justify='center',
                                    bg='#f4ecf7')
        self.entry_visa.grid(row=3, column=1, padx=5, pady=8, sticky='w')
        self.entry_visa.insert(0, "0")

        # ─── الصورة ───
        img_frame = tk.LabelFrame(main, text="صورة الفاتورة (اختياري)",
                                   font=('Arial', 10, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=8, pady=8)
        img_frame.grid(row=0, column=4, rowspan=6, padx=15, pady=5, sticky='n')

        img_container = tk.Frame(img_frame, width=140, height=140,
                                  bg='lightgray', relief='ridge', bd=2)
        img_container.pack()
        img_container.grid_propagate(False)

        self.image_label = tk.Label(img_container, bg='lightgray',
                                     text='\nلا توجد صورة\n',
                                     font=('Arial', 11))
        self.image_label.pack(fill='both', expand=True)
        self.image_label.bind("<Double-Button-1>", self._open_image)

        tk.Button(img_frame, text="📁 اختيار صورة",
                  command=self._select_image,
                  bg='#3498db', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=5).pack(pady=5)

        tk.Button(img_frame, text="🗑 إزالة الصورة",
                  command=self._remove_image,
                  bg='#e74c3c', fg='white',
                  font=('Arial', 9),
                  padx=10, pady=3).pack()

        # ─── أزرار الحفظ ───
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ",
                  command=self.save,
                  bg='#27ae60', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء",
                  command=self.window.destroy,
                  bg='#e74c3c', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        # ─── تحميل البيانات إن كان تعديلاً ───
        if self.is_edit:
            self._load_data()

        self.entry_cost.focus()
        self.entry_cost.select_range(0, tk.END)

    # ═══════════════════════════════════════════════════════════
    def _select_image(self):
        path = filedialog.askopenfilename(
            title="اختر الصورة",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.bmp"),
                       ("All files", "*.*")])
        if path:
            self.image_path = path
            self._show_preview(path)

    def _show_preview(self, path):
        try:
            from PIL import Image, ImageTk
            img = Image.open(path)
            img.thumbnail((140, 140), Image.Resampling.LANCZOS)
            self.image_preview = ImageTk.PhotoImage(img)
            self.image_label.config(image=self.image_preview, text='')
        except Exception as e:
            print(f"⚠️ خطأ في الصورة: {e}")

    def _remove_image(self):
        self.image_path = "REMOVE"
        self.image_label.config(image='', text='\nلا توجد صورة\n')
        self.image_preview = None

    def _open_image(self, event=None):
        if not self.image_path or self.image_path == "REMOVE":
            return
        if not os.path.exists(self.image_path):
            messagebox.showerror("خطأ", "الملف غير موجود")
            return
        try:
            os.startfile(self.image_path)
        except:
            import subprocess
            subprocess.Popen(['start', self.image_path], shell=True)

    def _load_data(self):
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute("""SELECT date, place, cost, euro, dollar,
                            pound, visa, image_path
                     FROM sales WHERE id=?""", (self.sale_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            messagebox.showerror("خطأ", "الصف غير موجود")
            self.window.destroy()
            return

        date, place, cost, euro, dollar, pound, visa, image_path = row

        try:
            if date:
                self.entry_date.set_date(datetime.strptime(date, '%Y-%m-%d'))
        except:
            pass

        self.place_var.set(place or 'bazar')

        for entry, val in [(self.entry_cost, cost),
                           (self.entry_euro, euro),
                           (self.entry_dollar, dollar),
                           (self.entry_pound, pound),
                           (self.entry_visa, visa)]:
            entry.delete(0, tk.END)
            entry.insert(0, str(val) if val not in (None, '') else '0')

        if image_path and os.path.exists(image_path):
            self.image_path = image_path
            self._show_preview(image_path)

    def save(self):
        date = self.entry_date.get().strip()
        place = self.place_var.get().strip()

        if not date:
            messagebox.showwarning("تنبيه", "التاريخ مطلوب")
            return

        try:
            cost = int(float(self.entry_cost.get() or 0))
            euro = int(float(self.entry_euro.get() or 0))
            dollar = int(float(self.entry_dollar.get() or 0))
            pound = int(float(self.entry_pound.get() or 0))
            visa = int(float(self.entry_visa.get() or 0))
        except ValueError:
            messagebox.showerror("خطأ", "القيم المالية يجب أن تكون أرقاماً")
            return

        # حساب الربح
        month_year = datetime.today().strftime('%Y-%m')
        rates = get_carried_balance(month_year, 'bazar')
        euro_rate = rates.get('euro_rate', 51.0) if isinstance(rates, dict) else 51.0
        usd_rate = rates.get('usd_rate', 48.0) if isinstance(rates, dict) else 48.0

        profit = calculate_profit(euro, dollar, pound, visa, cost,
                                   euro_rate, usd_rate)

        if self.is_edit:
            ok = update_sale(self.sale_id, date, place, cost, euro,
                             dollar, pound, visa, self.image_path, profit)
            msg_ok = "تم تعديل المبيعات"
            msg_err = "فشل التعديل"
        else:
            ok = add_sale(date, place, cost, euro, dollar, pound,
                          visa, self.image_path, profit)
            msg_ok = "تمت إضافة المبيعات"
            msg_err = "فشل الإضافة"

        if ok:
            messagebox.showinfo("نجاح", f"{msg_ok}\n💰 الربح: {profit} جنيه")
            self.window.destroy()
            self.parent_page.load_all_sales()
        else:
            messagebox.showerror("خطأ", msg_err)



class salesPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller
        self.selected_id = None

        # 1) الشريط العلوي
        add_program_banner(self)

        # 2) شريط الحالة
        self.status_bar = tk.Label(self, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#2c3e50', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # 3) منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        # العنوان
        tk.Label(content, text="💵 إدارة المبيعات",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=10)

        # ═══ إطار الأزرار ═══
        btn_frame = tk.LabelFrame(content, text="العمليات",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=15, pady=12)
        btn_frame.pack(fill='x', padx=15, pady=(0, 10))

        row1 = tk.Frame(btn_frame, bg='#ecf0f1')
        row1.pack(pady=3)

        tk.Button(row1, text="➕ إضافة مبيعات",
                  command=self.add_sale,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        self.btn_edit = tk.Button(row1, text="✏ تعديل",
                                   command=self.edit_sale,
                                   bg='#f39c12', fg='white',
                                   font=('Arial', 11, 'bold'),
                                   padx=20, pady=8, width=14,
                                   state='disabled')
        self.btn_edit.pack(side='left', padx=5)

        self.btn_delete = tk.Button(row1, text="🗑 حذف",
                                     command=self.delete_sale,
                                     bg='#e74c3c', fg='white',
                                     font=('Arial', 11, 'bold'),
                                     padx=20, pady=8, width=14,
                                     state='disabled')
        self.btn_delete.pack(side='left', padx=5)

        tk.Button(row1, text="🔄 تحديث",
                  command=self.load_all_sales,
                  bg='#95a5a6', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=14).pack(side='left', padx=5)

        tk.Button(row1, text="💾 تصدير Excel",
                  command=self.export_to_excel,
                  bg='#16a085', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        # ═══ إطار البحث ═══
        search_frame = tk.Frame(content, bg='#ecf0f1')
        search_frame.pack(fill='x', padx=15, pady=5)

        tk.Label(search_frame, text="🔍 بحث:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_entry = tk.Entry(search_frame, width=20,
                                      font=('Arial', 11))
        self.search_entry.pack(side='left', padx=5)

        tk.Label(search_frame, text="من:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_date_entry = DateEntry(search_frame, width=12,
                                            date_pattern='yyyy-mm-dd',
                                            font=('Arial', 10))
        self.search_date_entry.pack(side='left', padx=5)

        tk.Label(search_frame, text="المكان:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_place_var = tk.StringVar(value="الكل")
        ttk.Combobox(search_frame, textvariable=self.search_place_var,
                     values=("bazar", "beach", "الكل"),
                     state='readonly', width=10,
                     font=('Arial', 11)).pack(side='left', padx=5)

        tk.Button(search_frame, text="🔍 بحث",
                  command=self.search_sales,
                  bg='#9b59b6', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        tk.Button(search_frame, text="🗑 مسح",
                  command=self.clear_search,
                  bg='#bdc3c7', fg='#2c3e50',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        # ═══ الجدول ═══
        tree_frame = tk.LabelFrame(content, text="سجل المبيعات",
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 15))

        columns = ('ID', 'Date', 'Place', 'Cost', 'Euro', 'Dollar',
                   'L.E', 'Visa', 'Profit', 'Image')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                  show='headings', height=15)

        headers = {'ID': 'ID', 'Date': 'التاريخ', 'Place': 'المكان',
                   'Cost': 'التكلفة', 'Euro': 'يورو', 'Dollar': 'دولار',
                   'L.E': 'جنيه', 'Visa': 'فيزا', 'Profit': 'الربح',
                   'Image': 'الصورة'}
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('ID', width=50, anchor='center')
        self.tree.column('Date', width=100, anchor='center')
        self.tree.column('Place', width=80, anchor='center')
        self.tree.column('Cost', width=90, anchor='center')
        self.tree.column('Euro', width=80, anchor='center')
        self.tree.column('Dollar', width=80, anchor='center')
        self.tree.column('L.E', width=80, anchor='center')
        self.tree.column('Visa', width=80, anchor='center')
        self.tree.column('Profit', width=100, anchor='center')
        self.tree.column('Image', width=120, anchor='center')

        self.tree.tag_configure('odd_row', background='#f8f9fa')
        self.tree.tag_configure('even_row', background='#ffffff')

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', lambda e: self.edit_sale())
        # ✅ ربط النقر المزدوج
        self.tree.bind('<Double-Button-1>', self.on_tree_double_click)
        # ═══ الإجماليات ═══
        total_frame = tk.Frame(content, bg='#34495e')
        total_frame.pack(fill='x', padx=15, pady=(0, 10))

        self.total_cost_label = self._make_total_label(total_frame, "التكلفة", '#e74c3c')
        self.total_euro_label = self._make_total_label(total_frame, "يورو", '#3498db')
        self.total_dollar_label = self._make_total_label(total_frame, "دولار", '#16a085')
        self.total_le_label = self._make_total_label(total_frame, "جنيه", '#f1c40f')
        self.total_visa_label = self._make_total_label(total_frame, "فيزا", '#9b59b6')

        # تحميل
        self.load_all_sales()

    def _make_total_label(self, parent, title, color):
        card = tk.Frame(parent, bg='#34495e')
        card.pack(side='right', expand=True, padx=8, pady=8)
        tk.Label(card, text=title, font=('Arial', 10, 'bold'),
                 bg='#34495e', fg=color).pack()
        lbl = tk.Label(card, text="0", font=('Arial', 14, 'bold'),
                        bg='#34495e', fg='white')
        lbl.pack()
        return lbl

    def on_select(self, event):
        sel = self.tree.selection()
        if sel:
            try:
                self.selected_id = int(sel[0])
            except:
                self.selected_id = None
                return
            self.btn_edit.config(state='normal')
            self.btn_delete.config(state='normal')
        else:
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')
    def on_tree_double_click(self, event):
        """لو العمود اللي اتضغط عليه هو الصورة → افتحها، غير كده → افتح التعديل"""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        col = self.tree.identify_column(event.x)  # "#10" للعمود العاشر

        # ترتيب الأعمدة: ID=1, Date=2, Place=3, Cost=4, Euro=5,
        #                Dollar=6, L.E=7, Visa=8, Profit=9, Image=10
        if col == '#10':
            self.open_row_image()
        else:
            self.edit_sale()

    def open_row_image(self):
        """فتح صورة المبيعة المحددة"""
        sel = self.tree.selection()
        if not sel:
            return

        try:
            sale_id = int(sel[0])
        except (ValueError, TypeError):
            return

        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute("SELECT image_path FROM sales WHERE id=?", (sale_id,))
        row = c.fetchone()
        conn.close()

        if not row or not row[0]:
            messagebox.showinfo("معلومة", "لا توجد صورة لهذا السجل")
            return

        path = clean_path(row[0])

        if not path or not os.path.exists(path):
            messagebox.showerror("خطأ",
                                 f"ملف الصورة غير موجود:\n{path}")
            return

        try:
            os.startfile(path)
        except:
            import subprocess
            subprocess.Popen(['start', path], shell=True)
    def add_sale(self):
        SaleFormWindow(self, sale_id=None)

    def edit_sale(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر سجلاً أولاً")
            return
        SaleFormWindow(self, sale_id=self.selected_id)

    def delete_sale(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر سجلاً أولاً")
            return
        if not messagebox.askyesno("تأكيد",
                                    f"هل أنت متأكد من حذف السجل رقم {self.selected_id}؟"):
            return
        if delete_sale(self.selected_id):
            messagebox.showinfo("نجاح", "تم الحذف")
            self.selected_id = None
            self.load_all_sales()

    def load_all_sales(self):
        rows = get_all_sales()
        self._populate_tree(rows)
        self._update_totals(rows)
        self.status_bar.config(text=f"عدد السجلات: {len(rows)}")

    def _populate_tree(self, rows):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, row in enumerate(rows):
            img_path = row.get('image_path') or ''
            img_name = os.path.basename(img_path) if img_path else '---'
            tag = 'odd_row' if i % 2 == 0 else 'even_row'
            self.tree.insert('', 'end', iid=str(row.get('id')),
                              tags=(tag,),
                              values=(row.get('id', ''),
                                       row.get('date', ''),
                                       row.get('place', ''),
                                       row.get('cost', 0),
                                       row.get('euro', 0),
                                       row.get('dollar', 0),
                                       row.get('pound', 0),
                                       row.get('visa', 0),
                                       f"{row.get('profit', 0):,}",
                                       img_name))

    def _update_totals(self, rows):
        total_cost = sum(r.get('cost', 0) or 0 for r in rows)
        total_euro = sum(r.get('euro', 0) or 0 for r in rows)
        total_dollar = sum(r.get('dollar', 0) or 0 for r in rows)
        total_le = sum(r.get('pound', 0) or 0 for r in rows)
        total_visa = sum(r.get('visa', 0) or 0 for r in rows)

        self.total_cost_label.config(text=f"{total_cost:,}")
        self.total_euro_label.config(text=f"{total_euro:,}")
        self.total_dollar_label.config(text=f"{total_dollar:,}")
        self.total_le_label.config(text=f"{total_le:,}")
        self.total_visa_label.config(text=f"{total_visa:,}")

    def search_sales(self):
        term = self.search_entry.get().strip()
        date = self.search_date_entry.get().strip()
        place = self.search_place_var.get().strip()
        rows = search_sales(term, date, place)
        self._populate_tree(rows)
        self._update_totals(rows)

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        self.search_date_entry.delete(0, tk.END)
        self.search_place_var.set("الكل")
        self.load_all_sales()

    def export_to_excel(self):
        # نفس الكود القديم
        pass

# ═══════════════════════════════════════════════════════════════════
# PAID FORM WINDOW
# ═══════════════════════════════════════════════════════════════════
class PaidFormWindow:
    """نافذة إدخال / تعديل المصاريف"""

    def __init__(self, parent_page, paid_id=None):
        self.parent_page = parent_page
        self.paid_id = paid_id
        self.is_edit = paid_id is not None
        self.image_path = None
        self.image_preview = None

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = "تعديل مصروف" if self.is_edit else "إضافة مصروف جديد"
        self.window.title(title)
        self.window.geometry("820x620")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ─── العنوان ───
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icon = "✏" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'),
                 bg='#2c3e50', fg='white').pack(pady=15)

        # ─── الإطار الرئيسي ───
        main = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=15)
        main.pack(fill='both', expand=True)

        # ─── الصف 1: التاريخ + المكان ───
        tk.Label(main, text="التاريخ: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=0, column=0, padx=5, pady=8, sticky='e')

        self.entry_date = DateEntry(main, width=18,
                                     date_pattern='yyyy-mm-dd',
                                     font=('Arial', 11),
                                     justify='center')
        self.entry_date.grid(row=0, column=1, padx=5, pady=8, sticky='w')
        self.entry_date.set_date(datetime.now())

        tk.Label(main, text="المكان:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=0, column=2, padx=5, pady=8, sticky='e')

        self.place_var = tk.StringVar(value="bazar")
        ttk.Combobox(main, textvariable=self.place_var,
                     values=('bazar', 'beach'),
                     state='readonly', width=16,
                     font=('Arial', 11)).grid(
                     row=0, column=3, padx=5, pady=8, sticky='w')

        # ─── الصف 2: الوصف ───
        tk.Label(main, text="البيان: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=1, column=0, padx=5, pady=8, sticky='e')

        self.entry_description = tk.Entry(main, width=48,
                                           font=('Arial', 12))
        self.entry_description.grid(row=1, column=1, columnspan=3,
                                     padx=5, pady=8, sticky='w')

        # ─── الصف 3: القيمة + القسم ───
        tk.Label(main, text="القيمة (ج.م): *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#c0392b', anchor='e').grid(
                 row=2, column=0, padx=5, pady=8, sticky='e')

        self.entry_value = tk.Entry(main, width=20,
                                     font=('Arial', 13),
                                     justify='center',
                                     bg='#fadbd8')
        self.entry_value.grid(row=2, column=1, padx=5, pady=8, sticky='w')

        tk.Label(main, text="القسم: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=2, column=2, padx=5, pady=8, sticky='e')

        self.department_var = tk.StringVar(value="مشتريات")
        ttk.Combobox(main, textvariable=self.department_var,
                     values=('مشتريات', 'مصاريف'),
                     state='readonly', width=16,
                     font=('Arial', 11)).grid(
                     row=2, column=3, padx=5, pady=8, sticky='w')

        # ─── الصف 4: مصدر الدفع + المرسل ───
        tk.Label(main, text="مصدر الدفع:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=3, column=0, padx=5, pady=8, sticky='e')

        self.source_var = tk.StringVar(value="المحل")
        ttk.Combobox(main, textvariable=self.source_var,
                     values=('المحل', 'مدفوعات الحاج'),
                     state='readonly', width=18,
                     font=('Arial', 11)).grid(
                     row=3, column=1, padx=5, pady=8, sticky='w')

        tk.Label(main, text="المرسل:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=3, column=2, padx=5, pady=8, sticky='e')

        self.sender_var = tk.StringVar(value="البائع")
        ttk.Combobox(main, textvariable=self.sender_var,
                     values=('البائع', 'الحاج'),
                     state='readonly', width=16,
                     font=('Arial', 11)).grid(
                     row=3, column=3, padx=5, pady=8, sticky='w')

        # ─── الصف 5: يورو + دولار ───
        tk.Label(main, text="يورو (€):",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#2980b9', anchor='e').grid(
                 row=4, column=0, padx=5, pady=8, sticky='e')

        self.entry_euro = tk.Entry(main, width=20,
                                    font=('Arial', 12),
                                    justify='center',
                                    bg='#e8f4f8')
        self.entry_euro.grid(row=4, column=1, padx=5, pady=8, sticky='w')
        self.entry_euro.insert(0, "0")

        tk.Label(main, text="دولار ($):",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#16a085', anchor='e').grid(
                 row=4, column=2, padx=5, pady=8, sticky='e')

        self.entry_usd = tk.Entry(main, width=20,
                                   font=('Arial', 12),
                                   justify='center',
                                   bg='#e8f8f5')
        self.entry_usd.grid(row=4, column=3, padx=5, pady=8, sticky='w')
        self.entry_usd.insert(0, "0")

        # ─── الصورة ───
        img_frame = tk.LabelFrame(main, text="صورة الإيصال (اختياري)",
                                   font=('Arial', 10, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=8, pady=8)
        img_frame.grid(row=0, column=4, rowspan=5, padx=15, pady=5, sticky='n')

        img_container = tk.Frame(img_frame, width=140, height=140,
                                  bg='lightgray', relief='ridge', bd=2)
        img_container.pack()
        img_container.grid_propagate(False)

        self.image_label = tk.Label(img_container, bg='lightgray',
                                     text='\nلا توجد صورة\n',
                                     font=('Arial', 11))
        self.image_label.pack(fill='both', expand=True)
        self.image_label.bind("<Double-Button-1>", self._open_image)

        tk.Button(img_frame, text="📁 اختيار صورة",
                  command=self._select_image,
                  bg='#3498db', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=5).pack(pady=5)

        tk.Button(img_frame, text="🗑 إزالة",
                  command=self._remove_image,
                  bg='#e74c3c', fg='white',
                  font=('Arial', 9),
                  padx=10, pady=3).pack()

        # ─── أزرار الحفظ ───
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ",
                  command=self.save,
                  bg='#27ae60', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء",
                  command=self.window.destroy,
                  bg='#e74c3c', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        # ─── تحميل البيانات ───
        if self.is_edit:
            self._load_data()

        self.entry_description.focus()

    def _select_image(self):
        path = filedialog.askopenfilename(
            title="اختر الصورة",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.bmp"),
                       ("All files", "*.*")])
        if path:
            self.image_path = path
            self._show_preview(path)

    def _show_preview(self, path):
        try:
            from PIL import Image, ImageTk
            img = Image.open(path)
            img.thumbnail((140, 140), Image.Resampling.LANCZOS)
            self.image_preview = ImageTk.PhotoImage(img)
            self.image_label.config(image=self.image_preview, text='')
        except Exception as e:
            print(f"⚠️ خطأ في الصورة: {e}")

    def _remove_image(self):
        self.image_path = "REMOVE"
        self.image_label.config(image='', text='\nلا توجد صورة\n')
        self.image_preview = None

    def _open_image(self, event=None):
        if not self.image_path or self.image_path == "REMOVE":
            return
        if not os.path.exists(self.image_path):
            messagebox.showerror("خطأ", "الملف غير موجود")
            return
        try:
            os.startfile(self.image_path)
        except:
            import subprocess
            subprocess.Popen(['start', self.image_path], shell=True)

    def _load_data(self):
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute("""SELECT date, description, value, department,
                            place, source, sender, image_path, euro, usd
                     FROM paids WHERE id=?""", (self.paid_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            messagebox.showerror("خطأ", "الصف غير موجود")
            self.window.destroy()
            return

        (date, desc, val, dept, place, src, sender,
         img_path, euro, usd) = row

        try:
            if date:
                self.entry_date.set_date(datetime.strptime(date, '%Y-%m-%d'))
        except:
            pass

        self.entry_description.insert(0, desc or '')
        self.entry_value.insert(0, str(val) if val else '0')
        self.department_var.set(dept or 'مشتريات')
        self.place_var.set(place or 'bazar')
        self.source_var.set(src or 'المحل')
        self.sender_var.set(sender or 'البائع')
        self.entry_euro.delete(0, tk.END)
        self.entry_euro.insert(0, str(euro or 0))
        self.entry_usd.delete(0, tk.END)
        self.entry_usd.insert(0, str(usd or 0))

        if img_path and os.path.exists(img_path):
            self.image_path = img_path
            self._show_preview(img_path)

    def save(self):
        date = self.entry_date.get().strip()
        desc = self.entry_description.get().strip()

        if not desc:
            messagebox.showwarning("تنبيه", "البيان مطلوب")
            return

        try:
            value = float(self.entry_value.get() or 0)
            euro = float(self.entry_euro.get() or 0)
            usd = float(self.entry_usd.get() or 0)
        except ValueError:
            messagebox.showerror("خطأ", "القيم المالية يجب أن تكون أرقاماً")
            return

        dept = self.department_var.get()
        place = self.place_var.get()
        src = self.source_var.get()
        sender = self.sender_var.get()

        if self.is_edit:
            ok = update_paid(self.paid_id, date, desc, value, dept,
                             place, src, sender, self.image_path, euro, usd)
            msg_ok = "تم تعديل المصروف"
        else:
            ok = add_paid(date, desc, value, dept, place, src,
                          sender, self.image_path, euro, usd)
            msg_ok = "تمت إضافة المصروف"

        if ok:
            messagebox.showinfo("نجاح", msg_ok)
            self.window.destroy()
            self.parent_page.load_all_paids()
        else:
            messagebox.showerror("خطأ", "فشل الحفظ")


# ═══════════════════════════════════════════════════════════════════
# PAIDS PAGE
# ═══════════════════════════════════════════════════════════════════
class paidsPage(tk.Frame):
    """صفحة إدارة المصاريف"""

    def __init__(self, parent, controller):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller
        self.selected_id = None

        # ═══ 1) الشريط العلوي ═══
        add_program_banner(self)

        # ═══ 2) شريط الحالة (أسفل) ═══
        self.status_bar = tk.Label(self, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#2c3e50', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # ═══ 3) منطقة التمرير ═══
        content = make_scrollable(self)
        self._page_content = content

        # ═══ 4) العنوان ═══
        tk.Label(content, text="💸 إدارة المصاريف",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=10)

        # ═══ 5) إطار الأزرار ═══
        btn_frame = tk.LabelFrame(content, text="العمليات",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=15, pady=12)
        btn_frame.pack(fill='x', padx=15, pady=(0, 10))

        row1 = tk.Frame(btn_frame, bg='#ecf0f1')
        row1.pack(pady=3)

        tk.Button(row1, text="➕ إضافة مصروف",
                  command=self.add_paid,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        self.btn_edit = tk.Button(row1, text="✏ تعديل",
                                   command=self.edit_paid,
                                   bg='#f39c12', fg='white',
                                   font=('Arial', 11, 'bold'),
                                   padx=20, pady=8, width=14,
                                   state='disabled')
        self.btn_edit.pack(side='left', padx=5)

        self.btn_delete = tk.Button(row1, text="🗑 حذف",
                                     command=self.delete_paid,
                                     bg='#e74c3c', fg='white',
                                     font=('Arial', 11, 'bold'),
                                     padx=20, pady=8, width=14,
                                     state='disabled')
        self.btn_delete.pack(side='left', padx=5)

        tk.Button(row1, text="🔄 تحديث",
                  command=self.load_all_paids,
                  bg='#95a5a6', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=14).pack(side='left', padx=5)

        tk.Button(row1, text="🔄 ترحيل الدوري",
                  command=self.apply_recurring,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        # ═══ 6) إطار البحث ═══
        search_frame = tk.LabelFrame(content, text="فلاتر البحث",
                                      font=('Arial', 11, 'bold'),
                                      bg='#ecf0f1', fg='#34495e',
                                      padx=10, pady=10)
        search_frame.pack(fill='x', padx=15, pady=(0, 10))

        # الصف الأول: البحث + التواريخ
        row_s1 = tk.Frame(search_frame, bg='#ecf0f1')
        row_s1.pack(fill='x', pady=(0, 8))

        tk.Label(row_s1, text="🔍 بحث عام:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_entry = tk.Entry(row_s1, width=20,
                                      font=('Arial', 11),
                                      justify='center')
        self.search_entry.pack(side='left', padx=5)

        tk.Label(row_s1, text="من تاريخ:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.start_date_entry = DateEntry(row_s1, width=12,
                                           date_pattern='yyyy-mm-dd',
                                           font=('Arial', 10),
                                           justify='center')
        self.start_date_entry.pack(side='left', padx=5)

        tk.Label(row_s1, text="إلى تاريخ:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.end_date_entry = DateEntry(row_s1, width=12,
                                         date_pattern='yyyy-mm-dd',
                                         font=('Arial', 10),
                                         justify='center')
        self.end_date_entry.pack(side='left', padx=5)

        # الصف الثاني: الفلاتر
        row_s2 = tk.Frame(search_frame, bg='#ecf0f1')
        row_s2.pack(fill='x')

        tk.Label(row_s2, text="المكان:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_place_var = tk.StringVar(value="الكل")
        ttk.Combobox(row_s2, textvariable=self.search_place_var,
                     values=("bazar", "beach", "الكل"),
                     state='readonly', width=10,
                     font=('Arial', 11)).pack(side='left', padx=5)

        tk.Label(row_s2, text="مصدر الدفع:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_source_var = tk.StringVar(value="الكل")
        ttk.Combobox(row_s2, textvariable=self.search_source_var,
                     values=("المحل", "مدفوعات الحاج", "الكل"),
                     state='readonly', width=14,
                     font=('Arial', 11)).pack(side='left', padx=5)

        tk.Label(row_s2, text="القسم:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_department_var = tk.StringVar(value="الكل")
        ttk.Combobox(row_s2, textvariable=self.search_department_var,
                     values=("مشتريات", "مصاريف", "الكل"),
                     state='readonly', width=10,
                     font=('Arial', 11)).pack(side='left', padx=5)

        tk.Button(row_s2, text="🔍 بحث",
                  command=self.search_paids,
                  bg='#9b59b6', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        tk.Button(row_s2, text="🗑 مسح",
                  command=self.clear_search,
                  bg='#bdc3c7', fg='#2c3e50',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        tk.Button(row_s2, text="💾 تصدير Excel",
                  command=self.export_to_excel,
                  bg='#16a085', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        # ═══ 7) جدول العرض ═══
        tree_frame = tk.LabelFrame(content, text="سجل المصاريف",
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 5))

        columns = ('id', 'date', 'description', 'value', 'euro', 'usd',
                   'department', 'place', 'source', 'sender', 'image')

        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                  show='headings', height=10)

        headers = {
            'id': 'ID', 'date': 'التاريخ', 'description': 'البيان',
            'value': 'القيمة', 'euro': 'يورو', 'usd': 'دولار',
            'department': 'القسم', 'place': 'المكان',
            'source': 'مصدر الدفع', 'sender': 'المرسل', 'image': 'الصورة'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=40, anchor='center')
        self.tree.column('date', width=100, anchor='center')
        self.tree.column('description', width=220, anchor='center')
        self.tree.column('value', width=80, anchor='center')
        self.tree.column('euro', width=70, anchor='center')
        self.tree.column('usd', width=70, anchor='center')
        self.tree.column('department', width=90, anchor='center')
        self.tree.column('place', width=70, anchor='center')
        self.tree.column('source', width=110, anchor='center')
        self.tree.column('sender', width=80, anchor='center')
        self.tree.column('image', width=100, anchor='center')

        # ألوان الصفوف
        self.tree.tag_configure('odd_row', background='#f8f9fa')
        self.tree.tag_configure('even_row', background='#ffffff')
        self.tree.tag_configure('purchase_row', background='#d6eaf8')
        self.tree.tag_configure('expense_row', background='#fdebd0')

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', lambda e: self.edit_paid())
        # ✅ ربط النقر المزدوج
        self.tree.bind('<Double-Button-1>', self.on_tree_double_click)

        # ═══ 8) شريط الإجماليات ═══
        total_frame = tk.Frame(content, bg='#34495e')
        total_frame.pack(fill='x', padx=15, pady=(5, 15))

        self.total_value_label = self._make_total_card(
            total_frame, "الإجمالي (ج.م)", '#2ecc71')
        self.total_euro_label = self._make_total_card(
            total_frame, "إجمالي يورو", '#3498db')
        self.total_usd_label = self._make_total_card(
            total_frame, "إجمالي دولار", '#16a085')
        self.total_count_label = self._make_total_card(
            total_frame, "عدد السجلات", '#f1c40f')

        # ═══ تحميل البيانات ═══
        self.load_all_paids()

    # ═══════════════════════════════════════════════════════════
    def _make_total_card(self, parent, title, color):
        """إنشاء بطاقة إجمالي"""
        card = tk.Frame(parent, bg='#34495e')
        card.pack(side='right', expand=True, padx=8, pady=8)

        tk.Label(card, text=title, font=('Arial', 10, 'bold'),
                 bg='#34495e', fg=color).pack()

        lbl = tk.Label(card, text="0", font=('Arial', 14, 'bold'),
                        bg='#34495e', fg='white')
        lbl.pack()
        return lbl

    def on_tree_double_click(self, event):
        """لو العمود اللي اتضغط عليه هو الصورة → افتحها، غير كده → افتح التعديل"""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        col = self.tree.identify_column(event.x)

        # ترتيب الأعمدة في paidsPage:
        # id=1, date=2, description=3, value=4, euro=5, usd=6,
        # department=7, place=8, source=9, sender=10, image=11
        if col == '#11':
            self.open_row_image()
        else:
            self.edit_paid()

    def open_row_image(self):
        """فتح صورة المصروف المحدد"""
        sel = self.tree.selection()
        if not sel:
            return

        try:
            paid_id = int(sel[0])
        except (ValueError, TypeError):
            return

        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute("SELECT image_path FROM paids WHERE id=?", (paid_id,))
        row = c.fetchone()
        conn.close()

        if not row or not row[0]:
            messagebox.showinfo("معلومة", "لا توجد صورة لهذا السجل")
            return

        path = clean_path(row[0])

        if not path or not os.path.exists(path):
            messagebox.showerror("خطأ",
                                 f"ملف الصورة غير موجود:\n{path}")
            return

        try:
            os.startfile(path)
        except:
            import subprocess
            subprocess.Popen(['start', path], shell=True)

    # ═══════════════════════════════════════════════════════════
    # CRUD
    # ═══════════════════════════════════════════════════════════
    def add_paid(self):
        PaidFormWindow(self, paid_id=None)

    def edit_paid(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر سجلاً أولاً")
            return
        PaidFormWindow(self, paid_id=self.selected_id)

    def delete_paid(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر سجلاً أولاً")
            return
        if not messagebox.askyesno("تأكيد",
                                    f"هل أنت متأكد من حذف السجل رقم {self.selected_id}؟"):
            return
        if delete_paid(self.selected_id):
            messagebox.showinfo("نجاح", "تم الحذف")
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')
            self.load_all_paids()

    # ═══════════════════════════════════════════════════════════
    # تحميل وعرض البيانات
    # ═══════════════════════════════════════════════════════════
    def load_all_paids(self):
        rows = get_all_paids()
        self._populate_tree(rows)
        self._update_totals(rows)
        self.status_bar.config(text=f"عدد السجلات: {len(rows)}")

    def search_paids(self):
        term = self.search_entry.get().strip()
        start = self.start_date_entry.get().strip()
        end = self.end_date_entry.get().strip()
        dept = self.search_department_var.get().strip()
        place = self.search_place_var.get().strip()
        source = self.search_source_var.get().strip()

        rows = search_paids(term, start, end, dept, place, source)
        self._populate_tree(rows)
        self._update_totals(rows)
        self.status_bar.config(text=f"نتائج البحث: {len(rows)} سجل")

        if not rows:
            messagebox.showinfo("بحث", "لا توجد نتائج")

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        self.start_date_entry.delete(0, tk.END)
        self.end_date_entry.delete(0, tk.END)
        self.search_place_var.set("الكل")
        self.search_source_var.set("الكل")
        self.search_department_var.set("الكل")
        self.load_all_paids()

    def _populate_tree(self, rows):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not rows:
            return

        for i, row in enumerate(rows):
            img_path = row.get('image_path') or ''
            img_name = os.path.basename(img_path) if img_path else ''

            dept = row.get('department', '')
            if dept == 'مشتريات':
                tag = 'purchase_row'
            elif dept == 'مصاريف':
                tag = 'expense_row'
            else:
                tag = 'odd_row' if i % 2 == 0 else 'even_row'

            values = (
                row.get('id', ''),
                row.get('date', ''),
                row.get('description', ''),
                row.get('value', 0),
                row.get('euro', 0),
                row.get('usd', 0),
                row.get('department', ''),
                row.get('place', ''),
                row.get('source', ''),
                row.get('sender', ''),
                img_name,
            )

            self.tree.insert('', 'end', iid=str(row.get('id')),
                              tags=(tag,), values=values)

    def _update_totals(self, rows):
        total_value = sum(safe_float(r.get('value', 0)) for r in rows)
        total_euro = sum(safe_float(r.get('euro', 0)) for r in rows)
        total_usd = sum(safe_float(r.get('usd', 0)) for r in rows)

        self.total_value_label.config(text=f"{total_value:,.2f}")
        self.total_euro_label.config(text=f"{total_euro:,.2f}")
        self.total_usd_label.config(text=f"{total_usd:,.2f}")
        self.total_count_label.config(text=str(len(rows)))

    # ═══════════════════════════════════════════════════════════
    # التحديد
    # ═══════════════════════════════════════════════════════════
    def on_select(self, event):
        sel = self.tree.selection()
        if sel:
            try:
                self.selected_id = int(sel[0])
            except (ValueError, TypeError):
                self.selected_id = None
                return
            self.btn_edit.config(state='normal')
            self.btn_delete.config(state='normal')
        else:
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')

    # ═══════════════════════════════════════════════════════════
    # ترحيل المصاريف الدورية
    # ═══════════════════════════════════════════════════════════
    def apply_recurring(self):
        month_year = datetime.today().strftime('%Y-%m')

        if not messagebox.askyesno("تأكيد الترحيل",
                                    f"هل تريد ترحيل المصاريف الدورية للشهر {month_year}؟\n"
                                    f"(سيتم إضافتها إلى قائمة المصاريف)"):
            return

        added = apply_recurring_expenses(month_year)

        if added == -1:
            messagebox.showerror("خطأ", "حدث خطأ أثناء الترحيل")
        elif added == 0:
            messagebox.showinfo("تنبيه",
                                f"لا توجد مصاريف دورية جديدة للشهر {month_year}")
        else:
            messagebox.showinfo("نجاح",
                                f"تم ترحيل {added} مصروف دوري للشهر {month_year}")
            self.load_all_paids()

    # ═══════════════════════════════════════════════════════════
    # التصدير إلى Excel
    # ═══════════════════════════════════════════════════════════
    def export_to_excel(self):
        try:
            search_term = self.search_entry.get().strip()
            start_date = self.start_date_entry.get().strip()
            end_date = self.end_date_entry.get().strip()
            department = self.search_department_var.get().strip()
            place = self.search_place_var.get().strip()
            source = self.search_source_var.get().strip()

            if not start_date or not end_date:
                messagebox.showwarning("تنبيه",
                                        "يرجى إدخال تاريخ البداية والنهاية")
                return

            rows = search_paids(search_term, start_date, end_date,
                                 department, place, source)

            if not rows:
                messagebox.showwarning("تنبيه", "لا توجد مصاريف في هذه الفترة")
                return

            total = sum(safe_float(row.get('value', 0)) for row in rows)
            total_euro = sum(safe_float(row.get('euro', 0)) for row in rows)
            total_usd = sum(safe_float(row.get('usd', 0)) for row in rows)

            data_rows = []
            for r in rows:
                img_path = r.get('image_path') or ''
                img_name = os.path.basename(img_path) if img_path else ''

                data_rows.append([
                    r.get('date', '') or '',
                    r.get('description', '') or '',
                    safe_float(r.get('value', 0)),
                    safe_float(r.get('euro', 0)),
                    safe_float(r.get('usd', 0)),
                    r.get('department', '') or '',
                    r.get('place', '') or '',
                    r.get('source', '') or '',
                    img_name
                ])

            data_rows.append([
                "الإجمالي", "", total, total_euro, total_usd,
                "", "", "", ""
            ])

            headers = ["التاريخ", "البيان", "القيمة", "يورو", "دولار",
                       "القسم", "المكان", "مصدر الدفع", "الصورة"]

            report_title = (f"تقرير المصاريف ({start_date} إلى {end_date} "
                            f"قسم {department} مكان {place} دفع {source})")

            success = append_report_to_master(data_rows, report_title, headers)

            if success:
                messagebox.showinfo("نجاح",
                                     f"تم إضافة التقرير إلى الملف الرئيسي:\n{MASTER_FILE}")
            else:
                messagebox.showerror("خطأ", "فشل التصدير")

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"حدث خطأ: {e}")


##########    report_sales_page  #################

# ═══════════════════════════════════════════════════════════════════
# REPORT SALES PAGE (تقرير المبيعات اليومية)
# ═══════════════════════════════════════════════════════════════════
class report_sales_page(tk.Frame):
    """صفحة تقرير المبيعات اليومية"""

    def __init__(self, parent, controller):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller

        # ═══ 1) الشريط العلوي ═══
        add_program_banner(self)

        # ═══ 2) شريط الحالة (أسفل) ═══
        self.status_bar = tk.Label(self, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#2c3e50', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # ═══ 3) منطقة التمرير ═══
        content = make_scrollable(self)
        self._page_content = content

        # ═══ 4) العنوان ═══
        tk.Label(content, text="📊 تقرير المبيعات اليومية",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=10)

        tk.Label(content, text="عرض تفصيلي للمبيعات مجمّعة حسب اليوم والمكان",
                 font=('Arial', 11, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=(0, 10))

        # ═══ 5) فلاتر الفترة ═══
        filter_frame = tk.LabelFrame(content, text="اختر الفترة",
                                      font=('Arial', 11, 'bold'),
                                      bg='#ecf0f1', fg='#34495e',
                                      padx=15, pady=12)
        filter_frame.pack(fill='x', padx=15, pady=(0, 10))

        row1 = tk.Frame(filter_frame, bg='#ecf0f1')
        row1.pack(fill='x', pady=(0, 8))

        tk.Label(row1, text="من تاريخ:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(
                 side='right', padx=18)

        self.start_date_entry = DateEntry(row1, width=13,
                                           date_pattern='yyyy-mm-dd',
                                           font=('Arial', 11))
        self.start_date_entry.pack(side='right', padx=18)

        tk.Label(row1, text="إلى تاريخ:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(
                 side='right', padx=5)

        self.end_date_entry = DateEntry(row1, width=13,
                                         date_pattern='yyyy-mm-dd',
                                         font=('Arial', 11))
        self.end_date_entry.pack(side='right', padx=5)
        self.end_date_entry.set_date(datetime.today())

        tk.Label(row1, text="المكان:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(
                 side='right', padx=5)

        self.place_var = tk.StringVar(value="الكل")
        ttk.Combobox(row1, textvariable=self.place_var,
                     values=("الكل", "bazar", "beach", "other"),
                     state='readonly', width=12,
                     font=('Arial', 11)).pack(side='right', padx=5)

        # ─── أزرار العمليات ───
        row2 = tk.Frame(filter_frame, bg='#ecf0f1')
        row2.pack(fill='x')

        tk.Button(row2, text="🔍 عرض التقرير",
                  command=self.show_report,
                  bg='#27ae60', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=20, pady=8).pack(side='right', padx=5)

        tk.Button(row2, text="💾 تصدير Excel",
                  command=self.export_to_excel,
                  bg='#2980b9', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=20, pady=8).pack(side='right', padx=5)

        # ─── أزرار سريعة ───
        quick_frame = tk.Frame(filter_frame, bg='#ecf0f1')
        quick_frame.pack(fill='x', pady=(8, 0))

        tk.Label(quick_frame, text="فترات سريعة:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#34495e').pack(side='right', padx=5)

        tk.Button(quick_frame, text="اليوم",
                  command=self.set_today,
                  bg='#3498db', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=4).pack(side='right', padx=3)

        tk.Button(quick_frame, text="أمس",
                  command=self.set_yesterday,
                  bg='#95a5a6', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=4).pack(side='right', padx=3)

        tk.Button(quick_frame, text="آخر 7 أيام",
                  command=self.set_last_7_days,
                  bg='#f39c12', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=4).pack(side='right', padx=3)

        tk.Button(quick_frame, text="هذا الشهر",
                  command=self.set_this_month,
                  bg='#16a085', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=4).pack(side='right', padx=3)

        # ═══ 6) جدول التقرير ═══
        tree_frame = tk.LabelFrame(content, text="تفاصيل المبيعات اليومية",
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 5))

        columns = ('date', 'place', 'cost', 'euro', 'usd', 'egp', 'visa')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                  show='headings', height=15)

        headers = {
            'date': 'التاريخ',
            'place': 'المكان',
            'cost': 'التكلفة',
            'euro': 'يورو (€)',
            'usd': 'دولار ($)',
            'egp': 'جنيه (ج.م)',
            'visa': 'فيزا'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('date', width=130, anchor='center')
        self.tree.column('place', width=120, anchor='center')
        self.tree.column('cost', width=120, anchor='center')
        self.tree.column('euro', width=120, anchor='center')
        self.tree.column('usd', width=120, anchor='center')
        self.tree.column('egp', width=140, anchor='center')
        self.tree.column('visa', width=120, anchor='center')

        # ─── ألوان الصفوف ───
        self.tree.tag_configure('odd_row', background='#f8f9fa')
        self.tree.tag_configure('even_row', background='#ffffff')
        self.tree.tag_configure('bazar_row', background='#e8f4f8')
        self.tree.tag_configure('beach_row', background='#fef5e7')
        self.tree.tag_configure('other_row', background='#f4ecf7')
        self.tree.tag_configure('total_row',
                                 background='#2874a6',
                                 foreground='#ffffff',
                                 font=('Arial', 12, 'bold'))

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        # ═══ 7) بطاقات الإجماليات ═══
        total_frame = tk.Frame(content, bg='#34495e')
        total_frame.pack(fill='x', padx=15, pady=(5, 15))

        self.total_cost_label = self._make_card(
            total_frame, "إجمالي التكلفة", '#e74c3c')
        self.total_euro_label = self._make_card(
            total_frame, "إجمالي يورو", '#3498db')
        self.total_usd_label = self._make_card(
            total_frame, "إجمالي دولار", '#16a085')
        self.total_egp_label = self._make_card(
            total_frame, "إجمالي جنيه", '#f1c40f')
        self.total_visa_label = self._make_card(
            total_frame, "إجمالي فيزا", '#9b59b6')

    # ═══════════════════════════════════════════════════════════
    def _make_card(self, parent, title, color):
        """إنشاء بطاقة إجمالي"""
        card = tk.Frame(parent, bg='#34495e')
        card.pack(side='right', expand=True, padx=8, pady=8)

        tk.Label(card, text=title, font=('Arial', 10, 'bold'),
                 bg='#34495e', fg=color).pack(pady=(5, 2))

        lbl = tk.Label(card, text="0", font=('Arial', 14, 'bold'),
                        bg='#34495e', fg='white')
        lbl.pack(pady=(0, 5))
        return lbl

    # ═══════════════════════════════════════════════════════════
    # الفترات السريعة
    # ═══════════════════════════════════════════════════════════
    def set_today(self):
        """اليوم"""
        today = datetime.today()
        self.start_date_entry.set_date(today)
        self.end_date_entry.set_date(today)
        self.show_report()

    def set_yesterday(self):
        """أمس"""
        from datetime import timedelta
        yesterday = datetime.today() - timedelta(days=1)
        self.start_date_entry.set_date(yesterday)
        self.end_date_entry.set_date(yesterday)
        self.show_report()

    def set_last_7_days(self):
        """آخر 7 أيام"""
        from datetime import timedelta
        today = datetime.today()
        self.start_date_entry.set_date(today - timedelta(days=6))
        self.end_date_entry.set_date(today)
        self.show_report()

    def set_this_month(self):
        """هذا الشهر"""
        today = datetime.today()
        self.start_date_entry.set_date(today.replace(day=1))
        self.end_date_entry.set_date(today)
        self.show_report()

    # ═══════════════════════════════════════════════════════════
    # عرض التقرير
    # ═══════════════════════════════════════════════════════════
    def show_report(self):
        """عرض تقرير المبيعات اليومية"""
        start = self.start_date_entry.get().strip()
        end = self.end_date_entry.get().strip()
        place = self.place_var.get()

        if not start or not end:
            messagebox.showwarning("تنبيه",
                                    "يرجى إدخال تاريخ البداية والنهاية")
            return

        try:
            d1 = datetime.strptime(start, '%Y-%m-%d')
            d2 = datetime.strptime(end, '%Y-%m-%d')
            if d2 < d1:
                messagebox.showerror("خطأ",
                                     "تاريخ النهاية يجب أن يكون بعد البداية")
                return
        except:
            messagebox.showerror("خطأ", "صيغة التاريخ غير صحيحة")
            return

        # ─── جلب البيانات ───
        rows = get_sales_data(start, end, place)
        totals = get_sales_totals(start, end, place)

        # ─── تفريغ الجدول ───
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not rows:
            self.status_bar.config(
                text=f"⚠️ لا توجد مبيعات في الفترة ({start} - {end})")
            self._reset_totals()
            return

        # ─── إدراج الصفوف ───
        for i, row in enumerate(rows):
            place_row = row.get('place', '')
            if place_row == 'bazar':
                tag = 'bazar_row'
            elif place_row == 'beach':
                tag = 'beach_row'
            elif place_row == 'other':
                tag = 'other_row'
            else:
                tag = 'odd_row' if i % 2 == 0 else 'even_row'

            self.tree.insert('', 'end', tags=(tag,), values=(
                row.get('date', ''),
                place_row,
                f"{row.get('cost', 0):,}",
                f"{row.get('euro', 0):,}",
                f"{row.get('usd', 0):,}",
                f"{row.get('egp', 0):,}",
                f"{row.get('visa', 0):,}",
            ))

        # ─── صف الإجمالي في نهاية الجدول ───
        self.tree.insert('', 'end', tags=('total_row',), values=(
            'الإجمالي',
            '',
            f"{totals.get('cost', 0):,}",
            f"{totals.get('euro', 0):,}",
            f"{totals.get('usd', 0):,}",
            f"{totals.get('egp', 0):,}",
            f"{totals.get('visa', 0):,}",
        ))

        # ─── تحديث بطاقات الإجماليات ───
        self.total_cost_label.config(
            text=f"{totals.get('cost', 0):,}")
        self.total_euro_label.config(
            text=f"{totals.get('euro', 0):,}")
        self.total_usd_label.config(
            text=f"{totals.get('usd', 0):,}")
        self.total_egp_label.config(
            text=f"{totals.get('egp', 0):,}")
        self.total_visa_label.config(
            text=f"{totals.get('visa', 0):,}")

        # ─── شريط الحالة ───
        self.status_bar.config(
            text=f"✅ التقرير ({start} إلى {end}) | "
                 f"عدد الصفوف: {len(rows)}")

        # ─── حفظ للتصدير ───
        self.last_data = {
            'start': start,
            'end': end,
            'place': place,
            'rows': rows,
            'totals': totals,
        }

    def _reset_totals(self):
        """إعادة الإجماليات للصفر"""
        for lbl in [self.total_cost_label, self.total_euro_label,
                    self.total_usd_label, self.total_egp_label,
                    self.total_visa_label]:
            lbl.config(text="0")

    # ═══════════════════════════════════════════════════════════
    # التصدير إلى Excel
    # ═══════════════════════════════════════════════════════════
    def export_to_excel(self):
        """تصدير التقرير إلى الملف الرئيسي"""
        start = self.start_date_entry.get().strip()
        end = self.end_date_entry.get().strip()
        place = self.place_var.get()

        if not start or not end:
            messagebox.showwarning("تنبيه",
                                    "يرجى إدخال تاريخ البداية والنهاية")
            return

        rows_dict = get_sales_data(start, end, place)
        totals = get_sales_totals(start, end, place)

        if not rows_dict:
            messagebox.showwarning("تنبيه",
                                    "لا توجد بيانات في هذه الفترة")
            return

        try:
            # ─── بناء صفوف البيانات ───
            data_rows = []
            for r in rows_dict:
                data_rows.append([
                    r.get('date', ''),
                    r.get('place', ''),
                    r.get('cost', 0),
                    r.get('euro', 0),
                    r.get('usd', 0),
                    r.get('egp', 0),
                    r.get('visa', 0),
                ])

            # ─── صف الإجمالي ───
            data_rows.append([
                "الإجمالي", "",
                totals.get('cost', 0),
                totals.get('euro', 0),
                totals.get('usd', 0),
                totals.get('egp', 0),
                totals.get('visa', 0),
            ])

            headers = ["التاريخ", "المكان", "التكلفة",
                       "يورو", "دولار", "جنيه", "فيزا"]

            report_title = f"تقرير المبيعات اليومية ({start} إلى {end})"

            success = append_report_to_master(data_rows,
                                               report_title, headers)

            if success:
                messagebox.showinfo(
                    "نجاح",
                    f"تم إضافة التقرير إلى:\n{MASTER_FILE}")
            else:
                messagebox.showerror("خطأ", "فشل التصدير")

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"حدث خطأ: {e}")




###############  تقرير البيان الختامي  #####################


# ═══════════════════════════════════════════════════════════════════
# FINAL STATEMENT PAGE
# ═══════════════════════════════════════════════════════════════════
class FinalStatementPage(tk.Frame):
    """صفحة البيان الختامي - تقرير شامل للمبيعات والمصاريف والصافي"""

    def __init__(self, parent, controller):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller

        # ═══ 1) الشريط العلوي ═══
        add_program_banner(self)

        # ═══ 2) شريط الحالة (أسفل) ═══
        self.status_bar = tk.Label(self, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#2c3e50', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # ═══ 3) منطقة التمرير ═══
        content = make_scrollable(self)
        self._page_content = content

        # ═══ 4) العنوان ═══
        tk.Label(content, text="📊 البيان الختامي",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=10)

        tk.Label(content, text="تقرير شامل يجمع المبيعات والمصاريف والأرصدة المرحلة",
                 font=('Arial', 11, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=(0, 10))

        # ═══ 5) فلاتر الفترة ═══
        filter_frame = tk.LabelFrame(content, text="اختر الفترة",
                                      font=('Arial', 11, 'bold'),
                                      bg='#ecf0f1', fg='#34495e',
                                      padx=15, pady=12)
        filter_frame.pack(fill='x', padx=15, pady=(0, 10))

        row1 = tk.Frame(filter_frame, bg='#ecf0f1')
        row1.pack(fill='x', pady=(0, 8))

        tk.Label(row1, text="من تاريخ:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(
                 side='right', padx=18)

        self.start_date_entry = DateEntry(row1, width=13,
                                           date_pattern='yyyy-mm-dd',
                                           font=('Arial', 11))
        self.start_date_entry.pack(side='right', padx=18)

        tk.Label(row1, text="إلى تاريخ:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(
                 side='right', padx=5)

        self.end_date_entry = DateEntry(row1, width=13,
                                         date_pattern='yyyy-mm-dd',
                                         font=('Arial', 11))
        self.end_date_entry.pack(side='right', padx=5)
        self.end_date_entry.set_date(datetime.today())

        # ─── نسبة خصم الفيزا ───
        tk.Label(row1, text="نسبة خصم الفيزا (%):",
                 font=('Arial', 12, 'bold'),
                 bg='#ecf0f1', fg='#8e44ad').pack(side='right', padx=5)

        self.visa_discount_entry = tk.Entry(row1, width=8,
                                             font=('Arial', 12),
                                             justify='center',
                                             bg='#f4ecf7')
        self.visa_discount_entry.pack(side='right', padx=5)
        self.visa_discount_entry.insert(0, "2.5")

        # ─── أزرار العمليات ───
        row2 = tk.Frame(filter_frame, bg='#ecf0f1')
        row2.pack(fill='x')

        tk.Button(row2, text="🔍 عرض التقرير",
                  command=self.show_report,
                  bg='#27ae60', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=20, pady=8).pack(side='right', padx=5)

        tk.Button(row2, text="💾 تصدير Excel",
                  command=self.export_to_excel,
                  bg='#2980b9', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=20, pady=8).pack(side='right', padx=5)

        tk.Button(row2, text="💰 ترحيل الصافي",
                  command=self.transfer_profit,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=20, pady=8).pack(side='right', padx=5)

        # ═══ 6) جدول البيان الختامي ═══
        tree_frame = tk.LabelFrame(content, text="ملخص البيان",
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 15))

        columns = ('label', 'euro', 'usd', 'egp', 'visa', 'total')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                  show='headings', height=16)

        headers = {
            'label': 'البيان',
            'euro': 'يورو (€)',
            'usd': 'دولار ($)',
            'egp': 'جنيه (ج.م)',
            'visa': 'فيزا',
            'total': 'الإجمالي'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('label', width=280, anchor='e')
        self.tree.column('euro', width=130, anchor='center')
        self.tree.column('usd', width=130, anchor='center')
        self.tree.column('egp', width=130, anchor='center')
        self.tree.column('visa', width=110, anchor='center')
        self.tree.column('total', width=150, anchor='center')

        # ─── ألوان الصفوف ───
        self.tree.tag_configure('normal',
                                 background='#ffffff',
                                 foreground='#2c3e50',
                                 font=('Arial', 11))
        self.tree.tag_configure('total',
                                 background='#2874a6',
                                 foreground='#ffffff',
                                 font=('Arial', 12, 'bold'))
        self.tree.tag_configure('highlight',
                                 background='#f39c12',
                                 foreground='#ffffff',
                                 font=('Arial', 12, 'bold'))
        self.tree.tag_configure('final_positive',
                                 background='#27ae60',
                                 foreground='#ffffff',
                                 font=('Arial', 13, 'bold'))
        self.tree.tag_configure('final_negative',
                                 background='#e74c3c',
                                 foreground='#ffffff',
                                 font=('Arial', 13, 'bold'))

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

    # ═══════════════════════════════════════════════════════════
    # دوال الحساب (من الكود الأصلي مع تحسينات)
    # ═══════════════════════════════════════════════════════════
    def _calc_all(self):
        """حساب كل قيم البيان الختامي"""
        start = self.start_date_entry.get().strip()
        end = self.end_date_entry.get().strip()

        if not start or not end:
            messagebox.showwarning("تنبيه",
                                    "يرجى إدخال تاريخ البداية والنهاية")
            return None

        try:
            visa_discount_percent = float(self.visa_discount_entry.get().strip() or 2.5)
        except ValueError:
            visa_discount_percent = 2.5

        visa_discount_factor = 1 - (visa_discount_percent / 100.0)

        # ─── الشهر الحالي والشهر التالي ───
        month = end[:7]
        next_month = get_next_month(month)

        # ─── الرصيد المرحل للشهر الحالي ───
        carried = get_total_carried_balance(month)
        carried_euro = carried.get('euro', 0.0)
        carried_usd = carried.get('usd', 0.0)
        carried_egp = carried.get('egp', 0.0)
        euro_rate = carried.get('euro_rate', 51.0)
        usd_rate = carried.get('usd_rate', 48.0)

        # ─── الرصيد المرحل للشهر التالي ───
        next_carried = get_total_carried_balance(next_month)
        next_euro = next_carried.get('euro', 0.0)
        next_usd = next_carried.get('usd', 0.0)
        next_egp = next_carried.get('egp', 0.0)

        # ─── المبيعات حسب المكان ───
        sales_data = get_sales_by_place(start, end)
        bazar = sales_data.get('bazar',
                                {'euro': 0, 'usd': 0, 'egp': 0, 'visa': 0})
        beach = sales_data.get('beach',
                                {'euro': 0, 'usd': 0, 'egp': 0, 'visa': 0})
        other = sales_data.get('other',
                                {'euro': 0, 'usd': 0, 'egp': 0, 'visa': 0})

        # ─── إجماليات المبيعات ───
        total_euro = bazar['euro'] + beach['euro'] + other['euro'] + carried_euro
        total_usd = bazar['usd'] + beach['usd'] + other['usd'] + carried_usd
        total_egp = bazar['egp'] + beach['egp'] + other['egp'] + carried_egp
        total_visa = bazar['visa'] + beach['visa'] + other['visa']

        # ─── التحويلات ───
        euro_converted = total_euro * euro_rate
        usd_converted = total_usd * usd_rate
        visa_net = total_visa * visa_discount_factor
        total_egp_converted = total_egp + euro_converted + usd_converted + visa_net

        # ─── المصاريف حسب القسم ───
        expenses = get_expenses_by_department(start, end)
        purchases_cur = expenses.get('purchases',
                                      {'egp': 0.0, 'euro': 0.0, 'usd': 0.0})
        expenses_cur = expenses.get('expenses',
                                     {'egp': 0.0, 'euro': 0.0, 'usd': 0.0})
        other_cur = expenses.get('other',
                                  {'egp': 0.0, 'euro': 0.0, 'usd': 0.0})

        total_expenses_cur = {
            'egp': expenses_cur['egp'] + other_cur['egp'],
            'euro': expenses_cur['euro'] + other_cur['euro'],
            'usd': expenses_cur['usd'] + other_cur['usd']
        }

        purchases_total = (purchases_cur['egp'] +
                           purchases_cur['euro'] * euro_rate +
                           purchases_cur['usd'] * usd_rate)
        expenses_total = (expenses_cur['egp'] +
                          expenses_cur['euro'] * euro_rate +
                          expenses_cur['usd'] * usd_rate)

        # ─── الحسابات النهائية ───
        net_sales = total_egp_converted - purchases_total
        sub_total = net_sales - expenses_total

        carried_egp_total = (next_egp +
                              next_euro * euro_rate +
                              next_usd * usd_rate)

        final_net = sub_total - carried_egp_total

        return {
            'start': start, 'end': end,
            'month': month, 'next_month': next_month,
            'visa_discount_percent': visa_discount_percent,
            'carried': {'euro': carried_euro, 'usd': carried_usd, 'egp': carried_egp},
            'euro_rate': euro_rate, 'usd_rate': usd_rate,
            'next_carried': {'euro': next_euro, 'usd': next_usd, 'egp': next_egp},
            'bazar': bazar, 'beach': beach, 'other': other,
            'total_euro': total_euro, 'total_usd': total_usd,
            'total_egp': total_egp, 'total_visa': total_visa,
            'euro_converted': euro_converted,
            'usd_converted': usd_converted,
            'visa_net': visa_net,
            'total_egp_converted': total_egp_converted,
            'purchases_cur': purchases_cur,
            'purchases_total': purchases_total,
            'expenses_cur': expenses_cur,
            'other_cur': other_cur,
            'total_expenses_cur': total_expenses_cur,
            'expenses_total': expenses_total,
            'net_sales': net_sales,
            'sub_total': sub_total,
            'carried_egp_total': carried_egp_total,
            'final_net': final_net,
        }

    # ═══════════════════════════════════════════════════════════
    # عرض البيان
    # ═══════════════════════════════════════════════════════════
    def show_report(self):
        """حساب وعرض البيان الختامي"""
        data = self._calc_all()
        if not data:
            return

        # ─── تفريغ الجدول ───
        for item in self.tree.get_children():
            self.tree.delete(item)

        # ─── دالة مساعدة لتنسيق الأرقام ───
        def fmt(v):
            if v is None or v == '':
                return ''
            try:
                return f"{float(v):,.2f}"
            except:
                return str(v)

        # ═══ 1) الرصيد المرحل من الشهر السابق ═══
        self.tree.insert('', 'end', tags=('normal',), values=(
            'رصيد مرحل من الشهر السابق',
            fmt(data['carried']['euro']),
            fmt(data['carried']['usd']),
            fmt(data['carried']['egp']),
            '',
            '',
        ))

        # ═══ 2) مبيعات البازار ═══
        self.tree.insert('', 'end', tags=('normal',), values=(
            'مبيعات البازار',
            fmt(data['bazar']['euro']),
            fmt(data['bazar']['usd']),
            fmt(data['bazar']['egp']),
            fmt(data['bazar'].get('visa', 0)),
            '',
        ))

        # ═══ 3) مبيعات البيتش ═══
        self.tree.insert('', 'end', tags=('normal',), values=(
            'مبيعات البيتش',
            fmt(data['beach']['euro']),
            fmt(data['beach']['usd']),
            fmt(data['beach']['egp']),
            fmt(data['beach'].get('visa', 0)),
            '',
        ))

        # ═══ 4) الإجمالي (شامل الرصيد) ═══
        self.tree.insert('', 'end', tags=('total',), values=(
            'إجمالي (بما في ذلك الرصيد)',
            fmt(data['total_euro']),
            fmt(data['total_usd']),
            fmt(data['total_egp']),
            fmt(data['total_visa']),
            '',
        ))

        # ═══ 5) سعر الصرف / الخصم ═══
        self.tree.insert('', 'end', tags=('normal',), values=(
            'سعر الصرف / الخصم',
            fmt(data['euro_rate']),
            fmt(data['usd_rate']),
            '',
            f"خصم {data['visa_discount_percent']:.1f}%",
            '',
        ))

        # ═══ 6) المبيعات المحولة ═══
        self.tree.insert('', 'end', tags=('total',), values=(
            'المبيعات المحولة (بالجنيه)',
            fmt(data['euro_converted']),
            fmt(data['usd_converted']),
            fmt(data['total_egp']),
            fmt(data['visa_net']),
            fmt(data['total_egp_converted']),
        ))

        # ═══ 7) المشتريات ═══
        self.tree.insert('', 'end', tags=('normal',), values=(
            'المشتريات',
            fmt(data['purchases_cur']['euro']),
            fmt(data['purchases_cur']['usd']),
            fmt(data['purchases_cur']['egp']),
            '',
            fmt(data['purchases_total']),
        ))

        # ═══ 8) صافي المبيعات ═══
        self.tree.insert('', 'end', tags=('total',), values=(
            'صافي المبيعات',
            '', '', '', '',
            fmt(data['net_sales']),
        ))

        # ═══ 9) المصاريف ═══
        self.tree.insert('', 'end', tags=('normal',), values=(
            'المصاريف',
            fmt(data['expenses_cur']['euro']),
            fmt(data['expenses_cur']['usd']),
            fmt(data['expenses_cur']['egp']),
            '',
            fmt(data['expenses_total']),
        ))

        # ═══ 10) صافي (صافي المبيعات - المصاريف) ═══
        self.tree.insert('', 'end', tags=('total',), values=(
            'صافي (صافي المبيعات - المصاريف)',
            '', '', '', '',
            fmt(data['sub_total']),
        ))

        # ═══ 11) رصيد مرحل للشهر التالي ═══
        self.tree.insert('', 'end', tags=('normal',), values=(
            'رصيد مرحل للشهر التالي',
            fmt(data['next_carried']['euro']),
            fmt(data['next_carried']['usd']),
            fmt(data['next_carried']['egp']),
            '',
            fmt(data['carried_egp_total']),
        ))

        # ═══ 12) الصافي النهائي ═══
        final = data['final_net']
        tag = 'final_positive' if final >= 0 else 'final_negative'
        self.tree.insert('', 'end', tags=(tag,), values=(
            '⭐ الصافي النهائي ⭐',
            '', '', '', '',
            fmt(final),
        ))

        # ─── حفظ للتصدير ───
        self.last_data = data

        # ─── شريط الحالة ───
        self.status_bar.config(
            text=f"البيان من {data['start']} إلى {data['end']} | "
                 f"الصافي النهائي: {final:,.2f} ج.م")

    # ═══════════════════════════════════════════════════════════
    # ترحيل الصافي
    # ═══════════════════════════════════════════════════════════
    def transfer_profit(self):
        """ترحيل قيمة الصافي النهائي إلى صفحة إدارة الأرباح"""
        if not hasattr(self, 'last_data') or not self.last_data:
            messagebox.showwarning("تنبيه", "اعرض البيان أولاً")
            return

        data = self.last_data
        month = data['month']
        final_net = data['final_net']

        if not messagebox.askyesno(
                "تأكيد الترحيل",
                f"سيتم ترحيل الصافي للشهر {month}\n\n"
                f"القيمة: {final_net:,.2f} ج.م\n\n"
                f"هل تريد المتابعة؟"):
            return

        if add_profit_transfer(month, final_net):
            messagebox.showinfo(
                "نجاح",
                f"✅ تم ترحيل الصافي للشهر {month}\n"
                f"القيمة: {final_net:,.2f} ج.م")

            # الانتقال إلى صفحة إدارة الأرباح
            self.controller.show_page(ProfitTransferPage)
            page_ref = self.controller.pages.get('ProfitTransferPage')
            if page_ref and hasattr(page_ref, 'load_all_data'):
                page_ref.load_all_data()
        else:
            messagebox.showerror("خطأ", "فشل الترحيل")

    # ═══════════════════════════════════════════════════════════
    # التصدير إلى Excel
    # ═══════════════════════════════════════════════════════════
    def export_to_excel(self):
        """تصدير البيان الختامي إلى الملف الرئيسي"""
        if not hasattr(self, 'last_data') or not self.last_data:
            messagebox.showwarning("تنبيه", "اعرض البيان أولاً")
            return

        try:
            data = self.last_data

            def fmt(v):
                if v is None or v == '':
                    return ''
                try:
                    return round(float(v), 2)
                except:
                    return str(v)

            # ═══ بناء صفوف البيانات ═══
            data_rows = [
                ('رصيد مرحل من الشهر السابق',
                 fmt(data['carried']['euro']),
                 fmt(data['carried']['usd']),
                 fmt(data['carried']['egp']),
                 '', ''),

                ('مبيعات البازار',
                 fmt(data['bazar']['euro']),
                 fmt(data['bazar']['usd']),
                 fmt(data['bazar']['egp']),
                 fmt(data['bazar'].get('visa', 0)), ''),

                ('مبيعات البيتش',
                 fmt(data['beach']['euro']),
                 fmt(data['beach']['usd']),
                 fmt(data['beach']['egp']),
                 fmt(data['beach'].get('visa', 0)), ''),

                ('إجمالي (بما في ذلك الرصيد)',
                 fmt(data['total_euro']),
                 fmt(data['total_usd']),
                 fmt(data['total_egp']),
                 fmt(data['total_visa']), ''),

                ('سعر الصرف / الخصم',
                 fmt(data['euro_rate']),
                 fmt(data['usd_rate']),
                 '',
                 f"خصم {data['visa_discount_percent']:.1f}%",
                 ''),

                ('المبيعات المحولة (بالجنيه)',
                 fmt(data['euro_converted']),
                 fmt(data['usd_converted']),
                 fmt(data['total_egp']),
                 fmt(data['visa_net']),
                 fmt(data['total_egp_converted'])),

                ('المشتريات',
                 fmt(data['purchases_cur']['euro']),
                 fmt(data['purchases_cur']['usd']),
                 fmt(data['purchases_cur']['egp']),
                 '',
                 fmt(data['purchases_total'])),

                ('صافي المبيعات',
                 '', '', '', '',
                 fmt(data['net_sales'])),

                ('المصاريف',
                 fmt(data['expenses_cur']['euro']),
                 fmt(data['expenses_cur']['usd']),
                 fmt(data['expenses_cur']['egp']),
                 '',
                 fmt(data['expenses_total'])),

                ('صافي (صافي المبيعات - المصاريف)',
                 '', '', '', '',
                 fmt(data['sub_total'])),

                ('رصيد مرحل للشهر التالي',
                 fmt(data['next_carried']['euro']),
                 fmt(data['next_carried']['usd']),
                 fmt(data['next_carried']['egp']),
                 '',
                 fmt(data['carried_egp_total'])),

                ('⭐ الصافي النهائي ⭐',
                 '', '', '', '',
                 fmt(data['final_net'])),
            ]

            headers = ["البيان", "يورو", "دولار", "جنيه", "فيزا", "الإجمالي"]

            title = f"البيان الختامي ({data['start']} إلى {data['end']})"

            success = append_report_to_master(data_rows, title, headers)

            if success:
                messagebox.showinfo(
                    "نجاح",
                    f"تم إضافة التقرير إلى:\n{MASTER_FILE}")
            else:
                messagebox.showerror("خطأ", "فشل التصدير")

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"حدث خطأ: {e}")

# ═══════════════════════════════════════════════════════════════════
# CARRIED BALANCE FORM WINDOW
# ═══════════════════════════════════════════════════════════════════
class CarriedBalanceFormWindow:
    """نافذة إدخال / تعديل رصيد مرحّل"""

    def __init__(self, parent_page, month_year=None, place=None):
        self.parent_page = parent_page
        self.is_edit = bool(month_year and place)

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = ("تعديل رصيد مرحّل" if self.is_edit
                 else "إضافة رصيد مرحّل جديد")
        self.window.title(title)
        self.window.geometry("620x580")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ─── العنوان ───
        title_frame = tk.Frame(self.window, bg='#8B4513', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icon = "✏" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'),
                 bg='#8B4513', fg='white').pack(pady=15)

        # ─── الإطار الرئيسي ───
        main = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=15)
        main.pack(fill='both', expand=True)

        # ─── الشهر والمكان ───
        row0 = tk.Frame(main, bg='#ecf0f1')
        row0.pack(fill='x', pady=(0, 10))

        tk.Label(row0, text="الشهر (YYYY-MM): *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.entry_month = tk.Entry(row0, width=12,
                                     font=('Arial', 12),
                                     justify='center',
                                     bg='#fef9e7')
        self.entry_month.pack(side='left', padx=5)
        self.entry_month.insert(0, datetime.today().strftime('%Y-%m'))

        tk.Label(row0, text="المكان: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=15)

        self.place_var = tk.StringVar(value="bazar")
        ttk.Combobox(row0, textvariable=self.place_var,
                     values=('bazar', 'beach', 'other'),
                     state='readonly', width=12,
                     font=('Arial', 11)).pack(side='left', padx=5)

        # ═══ قسم: رصيد بداية الشهر ═══
        start_frame = tk.LabelFrame(main,
                                     text="📥 رصيد بداية الشهر (مرحّل من السابق)",
                                     font=('Arial', 11, 'bold'),
                                     bg='#e8f4f8', fg='#1a5276',
                                     padx=15, pady=10)
        start_frame.pack(fill='x', pady=(0, 10))

        self.start_entries = {}
        for idx, (label, key, color) in enumerate([
                ("يورو (€)", "euro", "#2980b9"),
                ("دولار ($)", "usd", "#16a085"),
                ("جنيه (ج.م)", "egp", "#27ae60")]):

            row = tk.Frame(start_frame, bg='#e8f4f8')
            row.pack(fill='x', pady=3)

            tk.Label(row, text=label + ":",
                     font=('Arial', 11, 'bold'),
                     bg='#e8f4f8', fg=color,
                     width=12, anchor='e').pack(side='left', padx=5)

            entry = tk.Entry(row, width=20,
                             font=('Arial', 12),
                             justify='center')
            entry.pack(side='left', padx=5)
            entry.insert(0, "0.00")

            self.start_entries[key] = entry

        # ─── سعر الصرف ───
        rates_frame = tk.LabelFrame(main,
                                     text="💱 سعر الصرف",
                                     font=('Arial', 11, 'bold'),
                                     bg='#fff9e6', fg='#8B4513',
                                     padx=15, pady=10)
        rates_frame.pack(fill='x', pady=(0, 10))

        row_r = tk.Frame(rates_frame, bg='#fff9e6')
        row_r.pack(fill='x', pady=3)

        tk.Label(row_r, text="سعر اليورو:",
                 font=('Arial', 11, 'bold'),
                 bg='#fff9e6', fg='#2980b9',
                 width=12, anchor='e').pack(side='left', padx=5)

        self.euro_rate_entry = tk.Entry(row_r, width=15,
                                         font=('Arial', 12),
                                         justify='center')
        self.euro_rate_entry.pack(side='left', padx=5)
        self.euro_rate_entry.insert(0, "51.00")

        tk.Label(row_r, text="سعر الدولار:",
                 font=('Arial', 11, 'bold'),
                 bg='#fff9e6', fg='#16a085',
                 width=12, anchor='e').pack(side='left', padx=15)

        self.usd_rate_entry = tk.Entry(row_r, width=15,
                                        font=('Arial', 12),
                                        justify='center')
        self.usd_rate_entry.pack(side='left', padx=5)
        self.usd_rate_entry.insert(0, "48.00")

        # ═══ قسم: رصيد نهاية الشهر ═══
        end_frame = tk.LabelFrame(main,
                                   text="📤 رصيد نهاية الشهر (مرحّل للشهر التالي)",
                                   font=('Arial', 11, 'bold'),
                                   bg='#fdedec', fg='#c0392b',
                                   padx=15, pady=10)
        end_frame.pack(fill='x', pady=(0, 10))

        self.end_entries = {}
        for idx, (label, key, color) in enumerate([
                ("يورو (€)", "euro", "#2980b9"),
                ("دولار ($)", "usd", "#16a085"),
                ("جنيه (ج.م)", "egp", "#27ae60")]):

            row = tk.Frame(end_frame, bg='#fdedec')
            row.pack(fill='x', pady=3)

            tk.Label(row, text=label + ":",
                     font=('Arial', 11, 'bold'),
                     bg='#fdedec', fg=color,
                     width=12, anchor='e').pack(side='left', padx=5)

            entry = tk.Entry(row, width=20,
                             font=('Arial', 12),
                             justify='center')
            entry.pack(side='left', padx=5)
            entry.insert(0, "0.00")

            self.end_entries[key] = entry

        # ─── ملاحظة ───
        tk.Label(main,
                 text="💡 عند الحفظ: يُحفظ رصيد البداية في هذا الشهر\n"
                      "ويُحفظ رصيد النهاية كرصيد بداية للشهر التالي تلقائياً",
                 font=('Arial', 9, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d',
                 justify='center').pack(pady=(5, 0))

        # ─── أزرار الحفظ ───
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ",
                  command=self.save,
                  bg='#27ae60', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء",
                  command=self.window.destroy,
                  bg='#e74c3c', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        # ─── تحميل البيانات إن كان تعديلاً ───
        if self.is_edit:
            self._load_data()

        self.entry_month.focus()

    # ═══════════════════════════════════════════════════════════
    def _load_data(self):
        """تحميل بيانات الرصيد الحالي"""
        month = self.entry_month.get().strip()
        place = self.place_var.get()

        bal = get_carried_balance(month, place)

        for key, entry in self.start_entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, f"{bal.get(key, 0):.2f}")

        self.euro_rate_entry.delete(0, tk.END)
        self.euro_rate_entry.insert(0, f"{bal.get('euro_rate', 51.0):.2f}")

        self.usd_rate_entry.delete(0, tk.END)
        self.usd_rate_entry.insert(0, f"{bal.get('usd_rate', 48.0):.2f}")

        # رصيد نهاية الشهر = رصيد بداية الشهر التالي
        next_month = get_next_month(month)
        next_bal = get_carried_balance(next_month, place)

        for key, entry in self.end_entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, f"{next_bal.get(key, 0):.2f}")

    def save(self):
        month = self.entry_month.get().strip()
        place = self.place_var.get().strip()

        if not month:
            messagebox.showwarning("تنبيه", "الشهر مطلوب")
            return

        try:
            # قراءة رصيد البداية
            start = {k: float(e.get() or 0) for k, e in self.start_entries.items()}
            euro_rate = float(self.euro_rate_entry.get() or 51.0)
            usd_rate = float(self.usd_rate_entry.get() or 48.0)

            # حفظ رصيد البداية
            save_carried_balance(month, place,
                                  start['euro'], start['usd'], start['egp'],
                                  euro_rate, usd_rate)

            # قراءة رصيد النهاية
            end = {k: float(e.get() or 0) for k, e in self.end_entries.items()}

            # حفظ رصيد النهاية كرصيد بداية للشهر التالي
            next_month = get_next_month(month)
            save_carried_balance(next_month, place,
                                  end['euro'], end['usd'], end['egp'],
                                  euro_rate, usd_rate)

            messagebox.showinfo(
                "نجاح",
                f"✅ تم حفظ أرصدة شهر {month} للمكان '{place}'\n\n"
                f"سيتم ترحيل رصيد النهاية إلى شهر {next_month}")

            self.window.destroy()
            self.parent_page.load_all()

        except ValueError:
            messagebox.showerror("خطأ", "يرجى إدخال أرقام صحيحة")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ: {e}")

# ═══════════════════════════════════════════════════════════════════
# CARRIED BALANCE PAGE (النسخة الأصلية بتصميم جديد)
# ═══════════════════════════════════════════════════════════════════
class CarriedBalancePage(tk.Frame):
    """صفحة إدارة الأرصدة المرحلة بين الشهور"""

    def __init__(self, parent, controller):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller

        # ═══ 1) الشريط العلوي ═══
        add_program_banner(self)

        # ═══ 2) شريط الحالة (أسفل) ═══
        self.status_bar = tk.Label(self, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#2c3e50', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # ═══ 3) منطقة التمرير ═══
        content = make_scrollable(self)
        self._page_content = content

        # ═══ 4) العنوان ═══
        tk.Label(content, text="💵 إدارة الأرصدة المرحلة",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=10)

        tk.Label(content, text="حدد الشهر والمكان، ثم اضغط تحميل لعرض الأرصدة",
                 font=('Arial', 11, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=(0, 10))

        # ═══════════════════════════════════════════════════════
        # 5) إطار اختيار الشهر والمكان
        # ═══════════════════════════════════════════════════════
        filter_frame = tk.LabelFrame(content, text="اختر الشهر والمكان",
                                      font=('Arial', 11, 'bold'),
                                      bg='#ecf0f1', fg='#34495e',
                                      padx=15, pady=12)
        filter_frame.pack(fill='x', padx=15, pady=(0, 10))

        row_filter = tk.Frame(filter_frame, bg='#ecf0f1')
        row_filter.pack(fill='x')

        tk.Label(row_filter, text="الشهر (YYYY-MM):",
                 font=('Arial', 12, 'bold'),
                 bg='#ecf0f1').pack(side='right', padx=5)

        self.month_entry = tk.Entry(row_filter, width=12,
                                     font=('Arial', 12),
                                     justify='center',
                                     bg='#fef9e7')
        self.month_entry.pack(side='right', padx=5)
        self.month_entry.insert(0, datetime.today().strftime('%Y-%m'))

        tk.Label(row_filter, text="المكان:",
                 font=('Arial', 12, 'bold'),
                 bg='#ecf0f1').pack(side='right', padx=15)

        self.place_var = tk.StringVar(value="bazar")
        ttk.Combobox(row_filter, textvariable=self.place_var,
                     values=('bazar', 'beach', 'other'),
                     state='readonly', width=12,
                     font=('Arial', 11)).pack(side='right', padx=5)

        tk.Button(row_filter, text="📥 تحميل",
                  command=self.load_balance,
                  bg='#3498db', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=20, pady=8).pack(side='right', padx=15)

        # ═══════════════════════════════════════════════════════
        # 6) قسم: رصيد بداية الشهر
        # ═══════════════════════════════════════════════════════
        start_frame = tk.LabelFrame(content,
                                     text="📥 رصيد بداية الشهر (مرحّل من الشهر السابق)",
                                     font=('Arial', 12, 'bold'),
                                     bg='#e8f4f8', fg='#1a5276',
                                     padx=20, pady=15)
        start_frame.pack(fill='x', padx=15, pady=(0, 10))

        self.start_entries = {}

        row_s = tk.Frame(start_frame, bg='#e8f4f8')
        row_s.pack(fill='x')

        for idx, (label, key, color) in enumerate([
                ("يورو (€)", "euro", "#2980b9"),
                ("دولار ($)", "usd", "#16a085"),
                ("جنيه (ج.م)", "egp", "#27ae60")]):

            cell = tk.Frame(row_s, bg='#e8f4f8')
            cell.pack(side='right', expand=True, padx=10)

            tk.Label(cell, text=label,
                     font=('Arial', 11, 'bold'),
                     bg='#e8f4f8', fg=color).pack()

            entry = tk.Entry(cell, width=15,
                             font=('Arial', 13, 'bold'),
                             justify='center',
                             bg='white',
                             relief='solid', bd=1)
            entry.pack(pady=5)
            entry.insert(0, "0.00")

            self.start_entries[key] = entry

        # ═══════════════════════════════════════════════════════
        # 7) قسم: سعر الصرف
        # ═══════════════════════════════════════════════════════
        rates_frame = tk.LabelFrame(content,
                                     text="💱 سعر الصرف",
                                     font=('Arial', 12, 'bold'),
                                     bg='#fff9e6', fg='#8B4513',
                                     padx=20, pady=15)
        rates_frame.pack(fill='x', padx=15, pady=(0, 10))

        row_r = tk.Frame(rates_frame, bg='#fff9e6')
        row_r.pack(fill='x')

        # سعر اليورو
        cell_e = tk.Frame(row_r, bg='#fff9e6')
        cell_e.pack(side='right', expand=True, padx=10)

        tk.Label(cell_e, text="سعر اليورو (€)",
                 font=('Arial', 11, 'bold'),
                 bg='#fff9e6', fg='#2980b9').pack()

        self.euro_rate_entry = tk.Entry(cell_e, width=15,
                                         font=('Arial', 13, 'bold'),
                                         justify='center',
                                         bg='white',
                                         relief='solid', bd=1)
        self.euro_rate_entry.pack(pady=5)
        self.euro_rate_entry.insert(0, "51.00")

        # سعر الدولار
        cell_d = tk.Frame(row_r, bg='#fff9e6')
        cell_d.pack(side='right', expand=True, padx=10)

        tk.Label(cell_d, text="سعر الدولار ($)",
                 font=('Arial', 11, 'bold'),
                 bg='#fff9e6', fg='#16a085').pack()

        self.usd_rate_entry = tk.Entry(cell_d, width=15,
                                        font=('Arial', 13, 'bold'),
                                        justify='center',
                                        bg='white',
                                        relief='solid', bd=1)
        self.usd_rate_entry.pack(pady=5)
        self.usd_rate_entry.insert(0, "48.00")

        # ═══════════════════════════════════════════════════════
        # 8) قسم: رصيد نهاية الشهر
        # ═══════════════════════════════════════════════════════
        end_frame = tk.LabelFrame(content,
                                   text="📤 رصيد نهاية الشهر (مرحّل للشهر التالي)",
                                   font=('Arial', 12, 'bold'),
                                   bg='#fdedec', fg='#c0392b',
                                   padx=20, pady=15)
        end_frame.pack(fill='x', padx=15, pady=(0, 10))

        self.end_entries = {}

        row_e = tk.Frame(end_frame, bg='#fdedec')
        row_e.pack(fill='x')

        for idx, (label, key, color) in enumerate([
                ("يورو (€)", "euro", "#2980b9"),
                ("دولار ($)", "usd", "#16a085"),
                ("جنيه (ج.م)", "egp", "#27ae60")]):

            cell = tk.Frame(row_e, bg='#fdedec')
            cell.pack(side='right', expand=True, padx=10)

            tk.Label(cell, text=label,
                     font=('Arial', 11, 'bold'),
                     bg='#fdedec', fg=color).pack()

            entry = tk.Entry(cell, width=15,
                             font=('Arial', 13, 'bold'),
                             justify='center',
                             bg='white',
                             relief='solid', bd=1)
            entry.pack(pady=5)
            entry.insert(0, "0.00")

            self.end_entries[key] = entry

        # ═══════════════════════════════════════════════════════
        # 9) ملاحظة
        # ═══════════════════════════════════════════════════════
        note_frame = tk.Frame(content, bg='#fef9e7',
                               relief='solid', bd=1)
        note_frame.pack(fill='x', padx=15, pady=(0, 10))

        tk.Label(note_frame,
                 text="💡 عند الحفظ: يُحفظ رصيد البداية في الشهر المحدد\n"
                      "ويُحفظ رصيد النهاية تلقائياً كرصيد بداية للشهر التالي",
                 font=('Arial', 10, 'italic'),
                 bg='#fef9e7', fg='#8B4513',
                 justify='center').pack(pady=10)

        # ═══════════════════════════════════════════════════════
        # 10) أزرار الحفظ
        # ═══════════════════════════════════════════════════════
        btn_frame = tk.Frame(content, bg='#ecf0f1')
        btn_frame.pack(pady=(10, 20))

        tk.Button(btn_frame, text="💾 حفظ",
                  command=self.save_balance,
                  bg='#27ae60', fg='white',
                  font=('Arial', 13, 'bold'),
                  padx=40, pady=12).pack(side='right', padx=10)

        tk.Button(btn_frame, text="🗑 مسح الحقول",
                  command=self.clear_fields,
                  bg='#95a5a6', fg='white',
                  font=('Arial', 13, 'bold'),
                  padx=40, pady=12).pack(side='right', padx=10)

    # ═══════════════════════════════════════════════════════════
    def load_balance(self):
        """تحميل أرصدة الشهر والمكان"""
        month = self.month_entry.get().strip()
        place = self.place_var.get()

        if not month:
            messagebox.showwarning("تنبيه", "يرجى إدخال الشهر")
            return

        # التحقق من صيغة الشهر
        try:
            datetime.strptime(month, '%Y-%m')
        except ValueError:
            messagebox.showerror("خطأ",
                                  "صيغة الشهر غير صحيحة\n"
                                  "استخدم الصيغة: YYYY-MM (مثل 2025-01)")
            return

        try:
            # ─── جلب رصيد البداية ───
            bal = get_carried_balance(month, place)

            if not isinstance(bal, dict) or not bal:
                bal = {'euro': 0.0, 'usd': 0.0, 'egp': 0.0,
                       'euro_rate': 51.0, 'usd_rate': 48.0}

            # ─── تعبئة حقول البداية ───
            for key, entry in self.start_entries.items():
                entry.delete(0, tk.END)
                entry.insert(0, f"{bal.get(key, 0):.2f}")

            self.euro_rate_entry.delete(0, tk.END)
            self.euro_rate_entry.insert(0, f"{bal.get('euro_rate', 51.0):.2f}")

            self.usd_rate_entry.delete(0, tk.END)
            self.usd_rate_entry.insert(0, f"{bal.get('usd_rate', 48.0):.2f}")

            # ─── جلب رصيد النهاية (من الشهر التالي) ───
            next_month = get_next_month(month)
            next_bal = get_carried_balance(next_month, place)

            if not isinstance(next_bal, dict) or not next_bal:
                next_bal = {'euro': 0.0, 'usd': 0.0, 'egp': 0.0}

            # ─── تعبئة حقول النهاية ───
            for key, entry in self.end_entries.items():
                entry.delete(0, tk.END)
                entry.insert(0, f"{next_bal.get(key, 0):.2f}")

            # ─── شريط الحالة ───
            self.status_bar.config(
                text=f"✅ تم تحميل رصيد شهر {month} للمكان '{place}' "
                     f"| سيُرحَّل إلى شهر {next_month}")

        except Exception as e:
            messagebox.showerror("خطأ", f"فشل تحميل البيانات: {e}")

    def save_balance(self):
        """حفظ الأرصدة"""
        month = self.month_entry.get().strip()
        place = self.place_var.get()

        if not month:
            messagebox.showwarning("تنبيه", "يرجى إدخال الشهر")
            return

        try:
            datetime.strptime(month, '%Y-%m')
        except ValueError:
            messagebox.showerror("خطأ", "صيغة الشهر غير صحيحة")
            return

        try:
            # ─── قراءة قيم البداية ───
            start = {}
            for key, entry in self.start_entries.items():
                start[key] = float(entry.get().strip() or 0)

            euro_rate = float(self.euro_rate_entry.get().strip() or 51.0)
            usd_rate = float(self.usd_rate_entry.get().strip() or 48.0)

            # ─── حفظ رصيد البداية ───
            save_carried_balance(month, place,
                                  start['euro'], start['usd'], start['egp'],
                                  euro_rate, usd_rate)

            # ─── قراءة قيم النهاية ───
            end = {}
            for key, entry in self.end_entries.items():
                end[key] = float(entry.get().strip() or 0)

            # ─── حفظ رصيد النهاية كرصيد بداية للشهر التالي ───
            next_month = get_next_month(month)
            save_carried_balance(next_month, place,
                                  end['euro'], end['usd'], end['egp'],
                                  euro_rate, usd_rate)

            messagebox.showinfo(
                "نجاح",
                f"✅ تم حفظ الأرصدة بنجاح\n\n"
                f"الشهر: {month}\n"
                f"المكان: {place}\n\n"
                f"📥 رصيد البداية: {start['egp']:,.2f} ج.م\n"
                f"📤 رصيد النهاية: {end['egp']:,.2f} ج.م\n\n"
                f"سيُرحَّل رصيد النهاية إلى شهر {next_month}")

            self.status_bar.config(
                text=f"✅ تم الحفظ بنجاح - رصيد {month} مرحَّل إلى {next_month}")

        except ValueError:
            messagebox.showerror("خطأ",
                                  "يرجى إدخال أرقام صحيحة في جميع الحقول")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل الحفظ: {e}")

    def clear_fields(self):
        """مسح جميع الحقول"""
        for entry in list(self.start_entries.values()) + \
                     list(self.end_entries.values()):
            entry.delete(0, tk.END)
            entry.insert(0, "0.00")

        self.euro_rate_entry.delete(0, tk.END)
        self.euro_rate_entry.insert(0, "51.00")

        self.usd_rate_entry.delete(0, tk.END)
        self.usd_rate_entry.insert(0, "48.00")

        self.status_bar.config(text="🗑 تم مسح الحقول")

    # ═══════════════════════════════════════════════════════════
    # التحميل والعرض
    # ═══════════════════════════════════════════════════════════
    def load_all(self):
        """تحميل كل الأرصدة"""
        rows = self._get_all_balances()
        self._populate_tree(rows)
        self._update_totals(rows)
        self.status_bar.config(text=f"عدد السجلات: {len(rows)}")

    def search_balances(self):
        """بحث بالأرصدة"""
        month = self.search_month_entry.get().strip()
        place = self.search_place_var.get().strip()

        if month and place != 'الكل':
            bal = get_carried_balance(month, place)
            if bal and isinstance(bal, dict) and bal.get('egp') is not None:
                rows = [{
                    'month_year': month,
                    'place': place,
                    'euro': bal.get('euro', 0),
                    'usd': bal.get('usd', 0),
                    'egp': bal.get('egp', 0),
                    'euro_rate': bal.get('euro_rate', 51.0),
                    'usd_rate': bal.get('usd_rate', 48.0),
                }]
            else:
                rows = []
        elif month:
            rows = self._get_balances_by_month(month)
        else:
            rows = self._get_all_balances()

        self._populate_tree(rows)
        self._update_totals(rows)
        self.status_bar.config(text=f"نتائج البحث: {len(rows)} سجل")

        if not rows:
            messagebox.showinfo("بحث", "لا توجد نتائج")

    def clear_search(self):
        """مسح البحث"""
        self.search_month_entry.delete(0, tk.END)
        self.search_month_entry.insert(0, datetime.today().strftime('%Y-%m'))
        self.search_place_var.set("الكل")
        self.load_all()

    def _get_all_balances(self):
        """جلب كل الأرصدة من قاعدة البيانات"""
        try:
            conn = sqlite3.connect('sunrise.db')
            conn.row_factory = sqlite3.Row
            c = conn.cursor()
            c.execute("""SELECT month_year, place, euro, usd, egp,
                                euro_rate, usd_rate
                         FROM carried_balances
                         ORDER BY month_year DESC, place""")
            rows = [dict(row) for row in c.fetchall()]
            conn.close()
            return rows
        except Exception as e:
            print(f"خطأ: {e}")
            return []

    def _get_balances_by_month(self, month):
        """جلب أرصدة شهر معين"""
        try:
            conn = sqlite3.connect('sunrise.db')
            conn.row_factory = sqlite3.Row
            c = conn.cursor()
            c.execute("""SELECT month_year, place, euro, usd, egp,
                                euro_rate, usd_rate
                         FROM carried_balances
                         WHERE month_year = ?
                         ORDER BY place""", (month,))
            rows = [dict(row) for row in c.fetchall()]
            conn.close()
            return rows
        except Exception as e:
            print(f"خطأ: {e}")
            return []

    def _populate_tree(self, rows):
        """عرض الصفوف في الجدول"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not rows:
            return

        for r in rows:
            place = r.get('place', '')
            if place == 'bazar':
                tag = 'bazar_row'
            elif place == 'beach':
                tag = 'beach_row'
            elif place == 'other':
                tag = 'other_row'
            else:
                tag = 'odd_row'

            iid = f"{r['month_year']}_{place}"

            self.tree.insert('', 'end', iid=iid, tags=(tag,), values=(
                r.get('month_year', ''),
                place,
                f"{r.get('euro', 0):,.2f}",
                f"{r.get('usd', 0):,.2f}",
                f"{r.get('egp', 0):,.2f}",
                f"{r.get('euro_rate', 51.0):,.2f}",
                f"{r.get('usd_rate', 48.0):,.2f}",
            ))

    def _update_totals(self, rows):
        """تحديث الإجماليات"""
        total_euro = sum(r.get('euro', 0) or 0 for r in rows)
        total_usd = sum(r.get('usd', 0) or 0 for r in rows)
        total_egp = sum(r.get('egp', 0) or 0 for r in rows)

        self.total_euro_label.config(text=f"{total_euro:,.2f}")
        self.total_usd_label.config(text=f"{total_usd:,.2f}")
        self.total_egp_label.config(text=f"{total_egp:,.2f}")
        self.total_count_label.config(text=str(len(rows)))

    # ═══════════════════════════════════════════════════════════
    # التحديد
    # ═══════════════════════════════════════════════════════════
    def on_select(self, event):
        """عند تحديد صف"""
        sel = self.tree.selection()
        if sel:
            iid = sel[0]
            try:
                month, place = iid.split('_', 1)
                self.selected_key = (month, place)
                self.btn_edit.config(state='normal')
                self.btn_delete.config(state='normal')
            except:
                self.selected_key = None
                self.btn_edit.config(state='disabled')
                self.btn_delete.config(state='disabled')
        else:
            self.selected_key = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')

    # ═══════════════════════════════════════════════════════════
    # التصدير إلى Excel
    # ═══════════════════════════════════════════════════════════
    def export_to_excel(self):
        """تصدير الأرصدة إلى الملف الرئيسي"""
        try:
            rows = self._get_all_balances()

            if not rows:
                messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
                return

            data_rows = []
            for r in rows:
                data_rows.append([
                    r.get('month_year', ''),
                    r.get('place', ''),
                    r.get('euro', 0),
                    r.get('usd', 0),
                    r.get('egp', 0),
                    r.get('euro_rate', 51.0),
                    r.get('usd_rate', 48.0),
                ])

            total_euro = sum(r.get('euro', 0) or 0 for r in rows)
            total_usd = sum(r.get('usd', 0) or 0 for r in rows)
            total_egp = sum(r.get('egp', 0) or 0 for r in rows)

            data_rows.append([
                "الإجمالي", "",
                total_euro, total_usd, total_egp,
                "", ""
            ])

            headers = ["الشهر", "المكان", "يورو", "دولار", "جنيه",
                       "سعر اليورو", "سعر الدولار"]

            title = "تقرير الأرصدة المرحلة"

            success = append_report_to_master(data_rows, title, headers)

            if success:
                messagebox.showinfo(
                    "نجاح",
                    f"تم إضافة التقرير إلى:\n{MASTER_FILE}")
            else:
                messagebox.showerror("خطأ", "فشل التصدير")

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"حدث خطأ: {e}")

        ####الخزينه العامه##############
# ═══════════════════════════════════════════════════════════════════
# TREASURY FORM WINDOW
# ═══════════════════════════════════════════════════════════════════
class TreasuryFormWindow:
    """نافذة إدخال / تعديل حركة خزينة"""

    def __init__(self, parent_page, payment_id=None):
        self.parent_page = parent_page
        self.payment_id = payment_id
        self.is_edit = payment_id is not None

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = "تعديل حركة خزينة" if self.is_edit else "إضافة حركة خزينة"
        self.window.title(title)
        self.window.geometry("600x400")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ─── العنوان ───
        title_frame = tk.Frame(self.window, bg='#8B4513', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icon = "✏" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'),
                 bg='#8B4513', fg='white').pack(pady=15)

        # ─── الإطار الرئيسي ───
        main = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=20)
        main.pack(fill='both', expand=True)

        # ─── التاريخ ───
        tk.Label(main, text="التاريخ: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=0, column=0, padx=5, pady=10, sticky='e')

        self.entry_date = DateEntry(main, width=20,
                                     date_pattern='yyyy-mm-dd',
                                     font=('Arial', 11),
                                     justify='center')
        self.entry_date.grid(row=0, column=1, padx=5, pady=10, sticky='w')
        self.entry_date.set_date(datetime.now())

        # ─── المكان ───
        tk.Label(main, text="المكان: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=1, column=0, padx=5, pady=10, sticky='e')

        self.place_var = tk.StringVar(value="bazar")
        ttk.Combobox(main, textvariable=self.place_var,
                     values=('bazar', 'beach', 'other'),
                     state='readonly', width=18,
                     font=('Arial', 11)).grid(
                     row=1, column=1, padx=5, pady=10, sticky='w')

        # ─── يورو ───
        tk.Label(main, text="يورو (€):",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#2980b9', anchor='e').grid(
                 row=2, column=0, padx=5, pady=10, sticky='e')

        self.entry_euro = tk.Entry(main, width=22,
                                    font=('Arial', 12),
                                    justify='center',
                                    bg='#e8f4f8')
        self.entry_euro.grid(row=2, column=1, padx=5, pady=10, sticky='w')
        self.entry_euro.insert(0, "0")

        # ─── دولار ───
        tk.Label(main, text="دولار ($):",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#16a085', anchor='e').grid(
                 row=3, column=0, padx=5, pady=10, sticky='e')

        self.entry_usd = tk.Entry(main, width=22,
                                   font=('Arial', 12),
                                   justify='center',
                                   bg='#e8f8f5')
        self.entry_usd.grid(row=3, column=1, padx=5, pady=10, sticky='w')
        self.entry_usd.insert(0, "0")

        # ─── جنيه ───
        tk.Label(main, text="جنيه (ج.م):",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#27ae60', anchor='e').grid(
                 row=4, column=0, padx=5, pady=10, sticky='e')

        self.entry_egp = tk.Entry(main, width=22,
                                   font=('Arial', 12),
                                   justify='center',
                                   bg='#eafaf1')
        self.entry_egp.grid(row=4, column=1, padx=5, pady=10, sticky='w')
        self.entry_egp.insert(0, "0")

        # ─── أزرار الحفظ ───
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ",
                  command=self.save,
                  bg='#27ae60', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء",
                  command=self.window.destroy,
                  bg='#e74c3c', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        if self.is_edit:
            self._load_data()

        self.entry_euro.focus()

    def _load_data(self):
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute("""SELECT payment_date, place, euro, usd, egp
                     FROM general_treasury WHERE id=?""", (self.payment_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            messagebox.showerror("خطأ", "السجل غير موجود")
            self.window.destroy()
            return

        date, place, euro, usd, egp = row

        try:
            if date:
                self.entry_date.set_date(datetime.strptime(date, '%Y-%m-%d'))
        except:
            pass

        self.place_var.set(place or 'bazar')
        for entry, val in [(self.entry_euro, euro),
                           (self.entry_usd, usd),
                           (self.entry_egp, egp)]:
            entry.delete(0, tk.END)
            entry.insert(0, str(val or 0))

    def save(self):
        date = self.entry_date.get().strip()
        place = self.place_var.get().strip()

        if not date or not place:
            messagebox.showwarning("تنبيه", "التاريخ والمكان مطلوبان")
            return

        try:
            euro = float(self.entry_euro.get() or 0)
            usd = float(self.entry_usd.get() or 0)
            egp = float(self.entry_egp.get() or 0)
        except ValueError:
            messagebox.showerror("خطأ", "القيم يجب أن تكون أرقاماً")
            return

        if self.is_edit:
            ok = update_treasury_payment(self.payment_id, date, place,
                                          euro, usd, egp)
            msg = "تم التعديل"
        else:
            ok = add_treasury_payment(date, place, euro, usd, egp)
            msg = "تمت الإضافة"

        if ok:
            messagebox.showinfo("نجاح", msg)
            self.window.destroy()
            self.parent_page.load_all()
        else:
            messagebox.showerror("خطأ", "فشل الحفظ")


# ═══════════════════════════════════════════════════════════════════
# GENERAL TREASURY PAGE
# ═══════════════════════════════════════════════════════════════════
class GeneralTreasuryPage(tk.Frame):
    """صفحة المبالغ المسلمة للخزينة العامة"""

    def __init__(self, parent, controller):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller
        self.selected_id = None

        # ═══ 1) الشريط العلوي ═══
        add_program_banner(self)

        # ═══ 2) شريط الحالة (أسفل) ═══
        self.status_bar = tk.Label(self, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#2c3e50', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # ═══ 3) منطقة التمرير ═══
        content = make_scrollable(self)
        self._page_content = content

        # ═══ 4) العنوان ═══
        tk.Label(content, text="🏦 المبالغ المسلمة للخزينة العامة",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=10)

        # ═══ 5) إطار الأزرار ═══
        btn_frame = tk.LabelFrame(content, text="العمليات",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=15, pady=12)
        btn_frame.pack(fill='x', padx=15, pady=(0, 10))

        row1 = tk.Frame(btn_frame, bg='#ecf0f1')
        row1.pack(pady=3)

        tk.Button(row1, text="➕ إضافة حركة",
                  command=self.add_payment,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        self.btn_edit = tk.Button(row1, text="✏ تعديل",
                                   command=self.edit_payment,
                                   bg='#f39c12', fg='white',
                                   font=('Arial', 11, 'bold'),
                                   padx=20, pady=8, width=14,
                                   state='disabled')
        self.btn_edit.pack(side='left', padx=5)

        self.btn_delete = tk.Button(row1, text="🗑 حذف",
                                     command=self.delete_payment,
                                     bg='#e74c3c', fg='white',
                                     font=('Arial', 11, 'bold'),
                                     padx=20, pady=8, width=14,
                                     state='disabled')
        self.btn_delete.pack(side='left', padx=5)

        tk.Button(row1, text="🔄 تحديث",
                  command=self.load_all,
                  bg='#95a5a6', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=14).pack(side='left', padx=5)

        tk.Button(row1, text="💾 تصدير Excel",
                  command=self.export_to_excel,
                  bg='#16a085', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        # ═══ 6) إطار البحث ═══
        search_frame = tk.LabelFrame(content, text="فلاتر البحث",
                                      font=('Arial', 11, 'bold'),
                                      bg='#ecf0f1', fg='#34495e',
                                      padx=10, pady=10)
        search_frame.pack(fill='x', padx=15, pady=(0, 10))

        row_s = tk.Frame(search_frame, bg='#ecf0f1')
        row_s.pack(fill='x')

        tk.Label(row_s, text="من تاريخ:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_start_entry = DateEntry(row_s, width=12,
                                             date_pattern='yyyy-mm-dd',
                                             font=('Arial', 10),
                                             justify='center')
        self.search_start_entry.pack(side='left', padx=5)

        tk.Label(row_s, text="إلى تاريخ:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_end_entry = DateEntry(row_s, width=12,
                                           date_pattern='yyyy-mm-dd',
                                           font=('Arial', 10),
                                           justify='center')
        self.search_end_entry.pack(side='left', padx=5)

        tk.Label(row_s, text="المكان:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_place_var = tk.StringVar(value="الكل")
        ttk.Combobox(row_s, textvariable=self.search_place_var,
                     values=("الكل", "bazar", "beach", "other"),
                     state='readonly', width=10,
                     font=('Arial', 11)).pack(side='left', padx=5)

        tk.Button(row_s, text="🔍 بحث",
                  command=self.search_payments,
                  bg='#9b59b6', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=5).pack(side='left', padx=5)

        tk.Button(row_s, text="🗑 مسح",
                  command=self.clear_search,
                  bg='#bdc3c7', fg='#2c3e50',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=5).pack(side='left', padx=5)

        # ═══ 7) جدول العرض ═══
        tree_frame = tk.LabelFrame(content, text="سجل حركات الخزينة",
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 5))

        columns = ('id', 'date', 'place', 'euro', 'usd', 'egp')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                  show='headings', height=15)

        headers = {
            'id': 'ID', 'date': 'التاريخ', 'place': 'المكان',
            'euro': 'يورو (€)', 'usd': 'دولار ($)', 'egp': 'جنيه (ج.م)'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=60, anchor='center')
        self.tree.column('date', width=130, anchor='center')
        self.tree.column('place', width=120, anchor='center')
        self.tree.column('euro', width=140, anchor='center')
        self.tree.column('usd', width=140, anchor='center')
        self.tree.column('egp', width=160, anchor='center')

        # ألوان الصفوف
        self.tree.tag_configure('odd_row', background='#f8f9fa')
        self.tree.tag_configure('even_row', background='#ffffff')

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', lambda e: self.edit_payment())

        # ═══ 8) بطاقات الإجماليات ═══
        total_frame = tk.Frame(content, bg='#34495e')
        total_frame.pack(fill='x', padx=15, pady=(5, 15))

        self.total_euro_label = self._make_total_card(
            total_frame, "إجمالي يورو (€)", '#3498db')
        self.total_usd_label = self._make_total_card(
            total_frame, "إجمالي دولار ($)", '#16a085')
        self.total_egp_label = self._make_total_card(
            total_frame, "إجمالي جنيه (ج.م)", '#f1c40f')
        self.total_count_label = self._make_total_card(
            total_frame, "عدد الحركات", '#e74c3c')

        # ─── تحميل ───
        self.load_all()

    # ═══════════════════════════════════════════════════════════
    def _make_total_card(self, parent, title, color):
        """إنشاء بطاقة إجمالي"""
        card = tk.Frame(parent, bg='#34495e')
        card.pack(side='right', expand=True, padx=8, pady=8)

        tk.Label(card, text=title, font=('Arial', 10, 'bold'),
                 bg='#34495e', fg=color).pack()

        lbl = tk.Label(card, text="0.00", font=('Arial', 14, 'bold'),
                        bg='#34495e', fg='white')
        lbl.pack()
        return lbl

    # ═══════════════════════════════════════════════════════════
    # CRUD
    # ═══════════════════════════════════════════════════════════
    def add_payment(self):
        TreasuryFormWindow(self, payment_id=None)

    def edit_payment(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر سجلاً أولاً")
            return
        TreasuryFormWindow(self, payment_id=self.selected_id)

    def delete_payment(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر سجلاً أولاً")
            return
        if not messagebox.askyesno("تأكيد",
                                    f"هل أنت متأكد من حذف السجل رقم {self.selected_id}؟"):
            return
        if delete_treasury_payment(self.selected_id):
            messagebox.showinfo("نجاح", "تم الحذف")
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')
            self.load_all()

    # ═══════════════════════════════════════════════════════════
    # التحميل والعرض
    # ═══════════════════════════════════════════════════════════
    def load_all(self):
        rows = get_all_treasury_payments()
        self._populate_tree(rows)
        self._update_totals(rows)
        self.status_bar.config(text=f"عدد الحركات: {len(rows)}")

    def search_payments(self):
        start = self.search_start_entry.get().strip()
        end = self.search_end_entry.get().strip()
        place = self.search_place_var.get().strip()

        if not start or not end:
            messagebox.showwarning("تنبيه",
                                    "يرجى إدخال تاريخ البداية والنهاية")
            return

        rows = search_treasury_payments(start, end, place)
        self._populate_tree(rows)
        self._update_totals(rows)
        self.status_bar.config(text=f"نتائج البحث: {len(rows)} سجل")

        if not rows:
            messagebox.showinfo("بحث", "لا توجد نتائج")

    def clear_search(self):
        self.search_start_entry.delete(0, tk.END)
        self.search_end_entry.delete(0, tk.END)
        self.search_place_var.set("الكل")
        self.load_all()

    def _populate_tree(self, rows):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not rows:
            return

        for i, row in enumerate(rows):
            tag = 'odd_row' if i % 2 == 0 else 'even_row'
            self.tree.insert('', 'end', iid=str(row['id']),
                              tags=(tag,), values=(
                                  row['id'],
                                  row.get('payment_date', ''),
                                  row.get('place', ''),
                                  f"{row.get('euro', 0):,.2f}",
                                  f"{row.get('usd', 0):,.2f}",
                                  f"{row.get('egp', 0):,.2f}",
                              ))

    def _update_totals(self, rows):
        total_euro = sum(r.get('euro', 0) or 0 for r in rows)
        total_usd = sum(r.get('usd', 0) or 0 for r in rows)
        total_egp = sum(r.get('egp', 0) or 0 for r in rows)

        self.total_euro_label.config(text=f"{total_euro:,.2f}")
        self.total_usd_label.config(text=f"{total_usd:,.2f}")
        self.total_egp_label.config(text=f"{total_egp:,.2f}")
        self.total_count_label.config(text=str(len(rows)))

    # ═══════════════════════════════════════════════════════════
    # التحديد
    # ═══════════════════════════════════════════════════════════
    def on_select(self, event):
        sel = self.tree.selection()
        if sel:
            try:
                self.selected_id = int(sel[0])
            except (ValueError, TypeError):
                self.selected_id = None
                return
            self.btn_edit.config(state='normal')
            self.btn_delete.config(state='normal')
        else:
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')

    # ═══════════════════════════════════════════════════════════
    # التصدير إلى Excel
    # ═══════════════════════════════════════════════════════════
    def export_to_excel(self):
        """تصدير بيانات الخزينة إلى الملف الرئيسي"""
        try:
            start = self.search_start_entry.get().strip()
            end = self.search_end_entry.get().strip()
            place = self.search_place_var.get()

            if start and end:
                month_year = end[:7]
                rows = search_treasury_payments(start, end, place)
                title = f"تقرير الخزينة ({start} إلى {end})"
            else:
                month_year = datetime.today().strftime('%Y-%m')
                rows = get_all_treasury_payments()
                title = "تقرير الخزينة (جميع السجلات)"

            if not rows:
                messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
                return

            # بناء صفوف البيانات
            data_rows = []
            for r in rows:
                data_rows.append([
                    r.get('payment_date', ''),
                    r.get('place', ''),
                    r.get('euro', 0),
                    r.get('usd', 0),
                    r.get('egp', 0),
                ])

            # حساب الإجماليات
            total_euro = sum(r.get('euro', 0) or 0 for r in rows)
            total_usd = sum(r.get('usd', 0) or 0 for r in rows)
            total_egp = sum(r.get('egp', 0) or 0 for r in rows)

            # جلب أسعار الصرف
            rates = get_carried_balance(month_year, 'bazar')
            euro_rate = rates.get('euro_rate', 51.0) if isinstance(rates, dict) else 51.0
            usd_rate = rates.get('usd_rate', 48.0) if isinstance(rates, dict) else 48.0

            total_converted = total_egp + (total_euro * euro_rate) + (total_usd * usd_rate)

            # صف الإجمالي
            data_rows.append([
                "الإجمالي", "",
                total_euro, total_usd, total_egp
            ])
            data_rows.append([
                "الإجمالي (محول للجنيه)", "",
                "", "", f"{total_converted:,.2f}"
            ])

            headers = ["التاريخ", "المكان", "يورو", "دولار", "جنيه"]

            success = append_report_to_master(data_rows, title, headers)

            if success:
                messagebox.showinfo("نجاح",
                                     f"تم إضافة التقرير إلى:\n{MASTER_FILE}")
            else:
                messagebox.showerror("خطأ", "فشل التصدير")

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"حدث خطأ: {e}")


# ═══════════════════════════════════════════════════════════════════
# CASH BALANCE PAGE (تقرير رصيد النقدية)
# ═══════════════════════════════════════════════════════════════════
class CashBalancePage(tk.Frame):
    """صفحة تقرير رصيد النقدية"""

    def __init__(self, parent, controller):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller

        # ═══ 1) الشريط العلوي ═══
        add_program_banner(self)

        # ═══ 2) شريط الحالة (أسفل) ═══
        self.status_bar = tk.Label(self, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#2c3e50', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # ═══ 3) منطقة التمرير ═══
        content = make_scrollable(self)
        self._page_content = content

        # ═══ 4) العنوان ═══
        tk.Label(content, text="💵 تقرير رصيد النقدية",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=10)

        tk.Label(content, text="يحسب الرصيد النقدي بعد كل الحركات والخصومات",
                 font=('Arial', 11, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=(0, 10))

        # ═══ 5) فلاتر الفترة ═══
        filter_frame = tk.LabelFrame(content, text="اختر الفترة",
                                      font=('Arial', 11, 'bold'),
                                      bg='#ecf0f1', fg='#34495e',
                                      padx=15, pady=12)
        filter_frame.pack(fill='x', padx=15, pady=(0, 10))

        row1 = tk.Frame(filter_frame, bg='#ecf0f1')
        row1.pack(fill='x', pady=(0, 8))

        tk.Label(row1, text="من تاريخ:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(
                 side='right', padx=18)

        self.start_date_entry = DateEntry(row1, width=13,
                                           date_pattern='yyyy-mm-dd',
                                           font=('Arial', 11))
        self.start_date_entry.pack(side='right', padx=18)

        tk.Label(row1, text="إلى تاريخ:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(
                 side='right', padx=5)

        self.end_date_entry = DateEntry(row1, width=13,
                                         date_pattern='yyyy-mm-dd',
                                         font=('Arial', 11))
        self.end_date_entry.pack(side='right', padx=5)
        self.end_date_entry.set_date(datetime.today())

        tk.Label(row1, text="المكان:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(
                 side='right', padx=5)

        self.place_var = tk.StringVar(value="الكل")
        ttk.Combobox(row1, textvariable=self.place_var,
                     values=("الكل", "bazar", "beach", "other"),
                     state='readonly', width=12,
                     font=('Arial', 11)).pack(side='right', padx=5)

        # ─── أزرار العمليات ───
        row2 = tk.Frame(filter_frame, bg='#ecf0f1')
        row2.pack(fill='x')

        tk.Button(row2, text="🔍 عرض التقرير",
                  command=self.show_report,
                  bg='#27ae60', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=20, pady=8).pack(side='right', padx=5)

        tk.Button(row2, text="💾 تصدير Excel",
                  command=self.export_to_excel,
                  bg='#2980b9', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=20, pady=8).pack(side='right', padx=5)

        # ─── أزرار سريعة ───
        quick_frame = tk.Frame(filter_frame, bg='#ecf0f1')
        quick_frame.pack(fill='x', pady=(8, 0))

        tk.Label(quick_frame, text="فترات سريعة:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#34495e').pack(side='right', padx=5)

        tk.Button(quick_frame, text="هذا الشهر",
                  command=self.set_this_month,
                  bg='#16a085', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=4).pack(side='right', padx=3)

        tk.Button(quick_frame, text="الشهر السابق",
                  command=self.set_prev_month,
                  bg='#95a5a6', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=4).pack(side='right', padx=3)

        tk.Button(quick_frame, text="آخر 30 يوم",
                  command=self.set_last_30_days,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=4).pack(side='right', padx=3)

        # ═══ 6) جدول التقرير ═══
        tree_frame = tk.LabelFrame(content, text="ملخص رصيد النقدية",
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 5))

        columns = ('label', 'euro', 'usd', 'egp')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                  show='headings', height=10)

        headers = {
            'label': 'البيان',
            'euro': 'يورو (€)',
            'usd': 'دولار ($)',
            'egp': 'جنيه (ج.م)'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('label', width=350, anchor='e')
        self.tree.column('euro', width=170, anchor='center')
        self.tree.column('usd', width=170, anchor='center')
        self.tree.column('egp', width=190, anchor='center')

        # ─── ألوان الصفوف ───
        self.tree.tag_configure('carried',
                                 background='#d6eaf8',
                                 foreground='#1a5276',
                                 font=('Arial', 11, 'bold'))
        self.tree.tag_configure('sales',
                                 background='#eafaf1',
                                 foreground='#145a32',
                                 font=('Arial', 11, 'bold'))
        self.tree.tag_configure('total',
                                 background='#2874a6',
                                 foreground='#ffffff',
                                 font=('Arial', 12, 'bold'))
        self.tree.tag_configure('treasury',
                                 background='#fef9e7',
                                 foreground='#7d6608',
                                 font=('Arial', 11, 'bold'))
        self.tree.tag_configure('expense',
                                 background='#fdedec',
                                 foreground='#a93226',
                                 font=('Arial', 11, 'bold'))
        self.tree.tag_configure('carry_next',
                                 background='#fff3cd',
                                 foreground='#7d6608',
                                 font=('Arial', 11, 'bold'))
        self.tree.tag_configure('final',
                                 background='#27ae60',
                                 foreground='#ffffff',
                                 font=('Arial', 13, 'bold'))
        self.tree.tag_configure('owner_money',
                                background='#fef9e7',
                                foreground='#7d6608',
                                font=('Arial', 11, 'bold'))

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        # ═══ 7) بطاقات الإجماليات ═══
        total_frame = tk.Frame(content, bg='#34495e')
        total_frame.pack(fill='x', padx=15, pady=(5, 15))

        self.total_initial_label = self._make_card(
            total_frame, "الرصيد الابتدائي", '#3498db')
        self.total_sales_label = self._make_card(
            total_frame, "إجمالي المبيعات", '#16a085')
        self.total_treasury_label = self._make_card(
            total_frame, "المسلّم للخزينة", '#f39c12')
        self.total_expenses_label = self._make_card(
            total_frame, "مصاريف المحل", '#e74c3c')
        self.total_final_label = self._make_card(
            total_frame, "المتبقي النهائي", '#2ecc71')

    # ═══════════════════════════════════════════════════════════
    def _make_card(self, parent, title, color):
        """إنشاء بطاقة إجمالي"""
        card = tk.Frame(parent, bg='#34495e')
        card.pack(side='right', expand=True, padx=5, pady=8)

        tk.Label(card, text=title, font=('Arial', 9, 'bold'),
                 bg='#34495e', fg=color).pack(pady=(5, 2))

        lbl = tk.Label(card, text="0", font=('Arial', 13, 'bold'),
                        bg='#34495e', fg='white')
        lbl.pack(pady=(0, 5))
        return lbl

    # ═══════════════════════════════════════════════════════════
    # الفترات السريعة
    # ═══════════════════════════════════════════════════════════
    def set_this_month(self):
        """هذا الشهر"""
        today = datetime.today()
        self.start_date_entry.set_date(today.replace(day=1))
        self.end_date_entry.set_date(today)
        self.show_report()

    def set_prev_month(self):
        """الشهر السابق"""
        from datetime import timedelta
        today = datetime.today()
        first_current = today.replace(day=1)
        last_prev = first_current - timedelta(days=1)
        first_prev = last_prev.replace(day=1)

        self.start_date_entry.set_date(first_prev)
        self.end_date_entry.set_date(last_prev)
        self.show_report()

    def set_last_30_days(self):
        """آخر 30 يوم"""
        from datetime import timedelta
        today = datetime.today()
        self.start_date_entry.set_date(today - timedelta(days=29))
        self.end_date_entry.set_date(today)
        self.show_report()

    # ═══════════════════════════════════════════════════════════
    # حساب وعرض التقرير
    # ═══════════════════════════════════════════════════════════
    def _calc_all(self):
        """حساب كل قيم رصيد النقدية"""
        start = self.start_date_entry.get().strip()
        end = self.end_date_entry.get().strip()
        place = self.place_var.get()

        if not start or not end:
            messagebox.showwarning("تنبيه",
                                    "يرجى إدخال تاريخ البداية والنهاية")
            return None

        try:
            d1 = datetime.strptime(start, '%Y-%m-%d')
            d2 = datetime.strptime(end, '%Y-%m-%d')
            if d2 < d1:
                messagebox.showerror("خطأ",
                                     "تاريخ النهاية يجب أن يكون بعد البداية")
                return None
        except:
            messagebox.showerror("خطأ", "صيغة التاريخ غير صحيحة")
            return None

        # ─── الشهر الحالي والشهر التالي ───
        month = end[:7]
        next_month = get_next_month(month)

        # ─── الرصيد المرحل من الشهر السابق ───
        carried = get_total_carried_balance(month, place)
        carried_euro = carried.get('euro', 0.0)
        carried_usd = carried.get('usd', 0.0)
        carried_egp = carried.get('egp', 0.0)

        # ─── الرصيد المرحل للشهر التالي ───
        next_carried = get_total_carried_balance(next_month, place)
        next_euro = next_carried.get('euro', 0.0)
        next_usd = next_carried.get('usd', 0.0)
        next_egp = next_carried.get('egp', 0.0)

        # ─── المبيعات (بدون فيزا) ───
        sales = get_sales_totals_by_place(start, end, place)
        sales_euro = sales.get('euro', 0.0)
        sales_usd = sales.get('dollar', 0.0)
        sales_egp = sales.get('pound', 0.0)

        # ─── المبالغ المسلمة من صاحب المكان (بفلتر المكان) ───
        owner_money = get_owner_money_totals(start, end, place)
        owner_money_net = owner_money.get('net', 0)

        # ─── الإجمالي الكلي (شامل أموال صاحب المكان) ───
        total_euro = carried_euro + sales_euro
        total_usd = carried_usd + sales_usd
        total_egp = carried_egp + sales_egp + owner_money_net

        # ─── الخزينة ───
        treasury = get_treasury_totals_by_place(start, end, place)
        treasury_euro = treasury.get('euro', 0.0)
        treasury_usd = treasury.get('usd', 0.0)
        treasury_egp = treasury.get('egp', 0.0)

        # ─── مصاريف المحل ───
        shop_expenses = get_shop_expenses_total(start, end, place)
        shop_euro = shop_expenses.get('euro', 0.0)
        shop_usd = shop_expenses.get('usd', 0.0)
        shop_egp = shop_expenses.get('egp', 0.0)

        # ─── المتبقي قبل الترحيل ───
        before_carry_euro = total_euro - treasury_euro - shop_euro
        before_carry_usd = total_usd - treasury_usd - shop_usd
        before_carry_egp = total_egp - treasury_egp - shop_egp

        # ─── المتبقي النهائي بعد خصم المرحل للشهر التالي ───
        final_euro = before_carry_euro - next_euro
        final_usd = before_carry_usd - next_usd
        final_egp = before_carry_egp - next_egp

        return {
            'start': start, 'end': end, 'place': place,
            'month': month, 'next_month': next_month,
            'carried': {'euro': carried_euro, 'usd': carried_usd, 'egp': carried_egp},  # ✅ dict
            'sales': {'euro': sales_euro, 'usd': sales_usd, 'egp': sales_egp},
            'owner_money': owner_money_net,
            'total': {'euro': total_euro, 'usd': total_usd, 'egp': total_egp},
            'treasury': {'euro': treasury_euro, 'usd': treasury_usd, 'egp': treasury_egp},
            'shop': {'euro': shop_euro, 'usd': shop_usd, 'egp': shop_egp},
            'next_carried': {'euro': next_euro, 'usd': next_usd, 'egp': next_egp},
            'final': {'euro': final_euro, 'usd': final_usd, 'egp': final_egp},
        }

    def show_report(self):
        """عرض تقرير رصيد النقدية"""
        data = self._calc_all()
        if not data:
            return

        # ─── تفريغ الجدول ───
        for item in self.tree.get_children():
            self.tree.delete(item)

        # ─── دالة تنسيق ───
        def fmt(v):
            try:
                return f"{float(v):,.2f}"
            except:
                return str(v)

        # ═══ 1) الرصيد المرحل من الشهر السابق ═══
        self.tree.insert('', 'end', tags=('carried',), values=(
            '📥 رصيد مرحل من الشهر السابق',
            fmt(data['carried']['euro']),
            fmt(data['carried']['usd']),
            fmt(data['carried']['egp']),
        ))

        # ═══ 2) إجمالي المبيعات ═══
        self.tree.insert('', 'end', tags=('sales',), values=(
            '💰 إجمالي المبيعات',
            fmt(data['sales']['euro']),
            fmt(data['sales']['usd']),
            fmt(data['sales']['egp']),
        ))

        # ═══ 3) المبالغ المسلمة من صاحب المكان ═══
        self.tree.insert('', 'end', tags=('owner_money',), values=(
            '🤝 المبالغ المسلمة من صاحب المكان',
            '',  # يورو
            '',  # دولار
            fmt(data['owner_money']),
        ))

        # ═══ 4) الإجمالي الكلي ═══
        self.tree.insert('', 'end', tags=('total',), values=(
            '📊 الإجمالي الكلي (رصيد + مبيعات + أموال صاحب المكان)',
            fmt(data['total']['euro']),
            fmt(data['total']['usd']),
            fmt(data['total']['egp']),
        ))

        # ═══ 4) المبالغ المسلمة للخزينة ═══
        self.tree.insert('', 'end', tags=('treasury',), values=(
            '🏦 المبالغ المسلمة للخزينة',
            fmt(data['treasury']['euro']),
            fmt(data['treasury']['usd']),
            fmt(data['treasury']['egp']),
        ))

        # ═══ 5) مصاريف المحل ═══
        self.tree.insert('', 'end', tags=('expense',), values=(
            '💸 مصاريف المحل (مخصومة)',
            fmt(data['shop']['euro']),
            fmt(data['shop']['usd']),
            fmt(data['shop']['egp']),
        ))

        # ═══ 6) رصيد مرحل للشهر التالي (مخصوم) ═══
        self.tree.insert('', 'end', tags=('carry_next',), values=(
            '📤 رصيد مرحل للشهر التالي (مخصوم)',
            fmt(data['next_carried']['euro']),
            fmt(data['next_carried']['usd']),
            fmt(data['next_carried']['egp']),
        ))

        # ═══ 7) المتبقي النهائي ═══
        self.tree.insert('', 'end', tags=('final',), values=(
            '⭐ المتبقي النهائي (رصيد النقدية)',
            fmt(data['final']['euro']),
            fmt(data['final']['usd']),
            fmt(data['final']['egp']),
        ))

        # ─── تحديث بطاقات الإجماليات ───
        self.total_initial_label.config(
            text=f"{data['carried']['egp']:,.2f}")
        self.total_sales_label.config(
            text=f"{data['sales']['egp']:,.2f}")
        self.total_treasury_label.config(
            text=f"{data['treasury']['egp']:,.2f}")
        self.total_expenses_label.config(
            text=f"{data['shop']['egp']:,.2f}")

        # المتبقي النهائي - لون حسب القيمة
        final_egp = data['final']['egp']
        self.total_final_label.config(
            text=f"{final_egp:,.2f}",
            fg='#2ecc71' if final_egp >= 0 else '#e74c3c')

        # ─── حفظ للتصدير ───
        self.last_data = data

        # ─── شريط الحالة ───
        self.status_bar.config(
            text=f"✅ التقرير ({data['start']} إلى {data['end']}) | "
                 f"المكان: {data['place']} | "
                 f"المتبقي النهائي: {final_egp:,.2f} ج.م")

    # ═══════════════════════════════════════════════════════════
    # التصدير إلى Excel
    # ═══════════════════════════════════════════════════════════
    def export_to_excel(self):
        """تصدير تقرير رصيد النقدية إلى الملف الرئيسي"""
        if not hasattr(self, 'last_data') or not self.last_data:
            messagebox.showwarning("تنبيه", "اعرض التقرير أولاً")
            return

        try:
            data = self.last_data

            def fmt(v):
                try:
                    return round(float(v), 2)
                except:
                    return v

            # ─── بناء صفوف البيانات ───
            data_rows = [
                ['📥 رصيد مرحل من الشهر السابق',
                 fmt(data['carried']['euro']),
                 fmt(data['carried']['usd']),
                 fmt(data['carried']['egp'])],

                ['💰 إجمالي المبيعات',
                 fmt(data['sales']['euro']),
                 fmt(data['sales']['usd']),
                 fmt(data['sales']['egp'])],

                ['📊 الإجمالي الكلي (رصيد + مبيعات)',
                 fmt(data['total']['euro']),
                 fmt(data['total']['usd']),
                 fmt(data['total']['egp'])],

                ['🏦 المبالغ المسلمة للخزينة',
                 fmt(data['treasury']['euro']),
                 fmt(data['treasury']['usd']),
                 fmt(data['treasury']['egp'])],

                ['💸 مصاريف المحل (مخصومة)',
                 fmt(data['shop']['euro']),
                 fmt(data['shop']['usd']),
                 fmt(data['shop']['egp'])],

                ['📤 رصيد مرحل للشهر التالي (مخصوم)',
                 fmt(data['next_carried']['euro']),
                 fmt(data['next_carried']['usd']),
                 fmt(data['next_carried']['egp'])],

                ['⭐ المتبقي النهائي (رصيد النقدية)',
                 fmt(data['final']['euro']),
                 fmt(data['final']['usd']),
                 fmt(data['final']['egp'])],
            ]

            headers = ["البيان", "يورو", "دولار", "جنيه"]

            title = (f"تقرير رصيد النقدية "
                     f"({data['start']} إلى {data['end']}) - {data['place']}")

            success = append_report_to_master(data_rows, title, headers)

            if success:
                messagebox.showinfo(
                    "نجاح",
                    f"تم إضافة التقرير إلى:\n{MASTER_FILE}")
            else:
                messagebox.showerror("خطأ", "فشل التصدير")

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"حدث خطأ: {e}")

# ═══════════════════════════════════════════════════════════════════
# RECURRING EXPENSE FORM WINDOW
# ═══════════════════════════════════════════════════════════════════
class RecurringExpenseFormWindow:
    """نافذة إدخال / تعديل مصروف دوري"""

    def __init__(self, parent_page, exp_id=None):
        self.parent_page = parent_page
        self.exp_id = exp_id
        self.is_edit = exp_id is not None

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = "تعديل مصروف دوري" if self.is_edit else "إضافة مصروف دوري"
        self.window.title(title)
        self.window.geometry("520x320")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ─── العنوان ───
        title_frame = tk.Frame(self.window, bg='#8e44ad', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icon = "✏" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'),
                 bg='#8e44ad', fg='white').pack(pady=15)

        # ─── الإطار الرئيسي ───
        main = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=20)
        main.pack(fill='both', expand=True)

        # ─── الوصف ───
        tk.Label(main, text="الوصف: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=0, column=0, padx=5, pady=12, sticky='e')

        self.entry_desc = tk.Entry(main, width=32,
                                    font=('Arial', 11))
        self.entry_desc.grid(row=0, column=1, padx=5, pady=12, sticky='w')

        # ─── القيمة ───
        tk.Label(main, text="القيمة (ج.م): *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#c0392b', anchor='e').grid(
                 row=1, column=0, padx=5, pady=12, sticky='e')

        self.entry_value = tk.Entry(main, width=32,
                                     font=('Arial', 13),
                                     justify='center',
                                     bg='#fadbd8')
        self.entry_value.grid(row=1, column=1, padx=5, pady=12, sticky='w')

        # ─── القسم ───
        tk.Label(main, text="القسم:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=2, column=0, padx=5, pady=12, sticky='e')

        self.dept_var = tk.StringVar(value="مشتريات")
        ttk.Combobox(main, textvariable=self.dept_var,
                     values=('مشتريات', 'مصاريف', 'رواتب'),
                     state='readonly', width=30,
                     font=('Arial', 11)).grid(
                     row=2, column=1, padx=5, pady=12, sticky='w')

        # ─── أزرار الحفظ ───
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ",
                  command=self.save,
                  bg='#27ae60', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء",
                  command=self.window.destroy,
                  bg='#e74c3c', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        if self.is_edit:
            self._load_data()

        self.entry_desc.focus()

    def _load_data(self):
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute("""SELECT description, value, department
                     FROM recurring_expenses WHERE id=?""", (self.exp_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            messagebox.showerror("خطأ", "السجل غير موجود")
            self.window.destroy()
            return

        desc, value, dept = row
        self.entry_desc.insert(0, desc or '')
        self.entry_value.insert(0, str(value or 0))
        self.dept_var.set(dept or 'مشتريات')

    def save(self):
        desc = self.entry_desc.get().strip()
        try:
            value = float(self.entry_value.get() or 0)
        except ValueError:
            messagebox.showerror("خطأ", "أدخل قيمة صحيحة")
            return

        if not desc:
            messagebox.showwarning("تنبيه", "الوصف مطلوب")
            return

        dept = self.dept_var.get()

        if self.is_edit:
            ok = update_recurring_expense(self.exp_id, desc, value, dept)
            msg = "تم التعديل"
        else:
            ok = add_recurring_expense(desc, value, dept)
            msg = "تمت الإضافة"

        if ok:
            messagebox.showinfo("نجاح", msg)
            self.window.destroy()
            self.parent_page.load_all()
        else:
            messagebox.showerror("خطأ", "فشل الحفظ")


# ═══════════════════════════════════════════════════════════════════
# RECURRING EXPENSES PAGE
# ═══════════════════════════════════════════════════════════════════
class RecurringExpensesPage(tk.Frame):
    """صفحة إدارة المصاريف الدورية (الثابتة)"""

    def __init__(self, parent, controller):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller
        self.selected_id = None

        # ═══ 1) الشريط العلوي ═══
        add_program_banner(self)

        # ═══ 2) شريط الحالة (أسفل) ═══
        self.status_bar = tk.Label(self, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#2c3e50', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # ═══ 3) منطقة التمرير ═══
        content = make_scrollable(self)
        self._page_content = content

        # ═══ 4) العنوان ═══
        tk.Label(content, text="🔄 إدارة المصاريف الدورية",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=10)

        tk.Label(content, text="المصاريف الثابتة التي تتكرر شهرياً",
                 font=('Arial', 11, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=(0, 10))

        # ═══ 5) إطار الأزرار ═══
        btn_frame = tk.LabelFrame(content, text="العمليات",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=15, pady=12)
        btn_frame.pack(fill='x', padx=15, pady=(0, 10))

        row1 = tk.Frame(btn_frame, bg='#ecf0f1')
        row1.pack(pady=3)

        tk.Button(row1, text="➕ إضافة مصروف دوري",
                  command=self.add_recurring,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=20).pack(side='left', padx=5)

        self.btn_edit = tk.Button(row1, text="✏ تعديل",
                                   command=self.edit_recurring,
                                   bg='#f39c12', fg='white',
                                   font=('Arial', 11, 'bold'),
                                   padx=20, pady=8, width=14,
                                   state='disabled')
        self.btn_edit.pack(side='left', padx=5)

        self.btn_delete = tk.Button(row1, text="🗑 حذف",
                                     command=self.delete_recurring,
                                     bg='#e74c3c', fg='white',
                                     font=('Arial', 11, 'bold'),
                                     padx=20, pady=8, width=14,
                                     state='disabled')
        self.btn_delete.pack(side='left', padx=5)

        tk.Button(row1, text="🔄 تحديث",
                  command=self.load_all,
                  bg='#95a5a6', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=14).pack(side='left', padx=5)

        tk.Button(row1, text="📥 ترحيل للشهر الحالي",
                  command=self.transfer_now,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=20).pack(side='left', padx=5)

        # ═══ 6) جدول العرض ═══
        tree_frame = tk.LabelFrame(content, text="قائمة المصاريف الدورية",
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 5))

        columns = ('id', 'description', 'value', 'department')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                  show='headings', height=15)

        headers = {
            'id': 'ID',
            'description': 'الوصف',
            'value': 'القيمة (ج.م)',
            'department': 'القسم'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=60, anchor='center')
        self.tree.column('description', width=400, anchor='e')
        self.tree.column('value', width=180, anchor='center')
        self.tree.column('department', width=180, anchor='center')

        # ألوان الصفوف
        self.tree.tag_configure('odd_row', background='#f8f9fa')
        self.tree.tag_configure('even_row', background='#ffffff')
        self.tree.tag_configure('purchase_row', background='#d6eaf8')
        self.tree.tag_configure('expense_row', background='#fdebd0')
        self.tree.tag_configure('salary_row', background='#e8f8f5')

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', lambda e: self.edit_recurring())

        # ═══ 7) بطاقات الإجماليات ═══
        total_frame = tk.Frame(content, bg='#34495e')
        total_frame.pack(fill='x', padx=15, pady=(5, 15))

        self.total_value_label = self._make_total_card(
            total_frame, "إجمالي القيم (ج.م)", '#2ecc71')
        self.total_count_label = self._make_total_card(
            total_frame, "عدد البنود", '#f1c40f')
        self.total_purchase_label = self._make_total_card(
            total_frame, "مشتريات", '#3498db')
        self.total_expense_label = self._make_total_card(
            total_frame, "مصاريف", '#e67e22')

        # ─── تحميل ───
        self.load_all()

    # ═══════════════════════════════════════════════════════════
    def _make_total_card(self, parent, title, color):
        """إنشاء بطاقة إجمالي"""
        card = tk.Frame(parent, bg='#34495e')
        card.pack(side='right', expand=True, padx=8, pady=8)

        tk.Label(card, text=title, font=('Arial', 10, 'bold'),
                 bg='#34495e', fg=color).pack()

        lbl = tk.Label(card, text="0.00", font=('Arial', 14, 'bold'),
                        bg='#34495e', fg='white')
        lbl.pack()
        return lbl

    # ═══════════════════════════════════════════════════════════
    # CRUD
    # ═══════════════════════════════════════════════════════════
    def add_recurring(self):
        RecurringExpenseFormWindow(self, exp_id=None)

    def edit_recurring(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر سجلاً أولاً")
            return
        RecurringExpenseFormWindow(self, exp_id=self.selected_id)

    def delete_recurring(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر سجلاً أولاً")
            return
        if not messagebox.askyesno(
                "تأكيد",
                f"هل أنت متأكد من حذف هذا المصروف الدوري؟"):
            return
        if delete_recurring_expense(self.selected_id):
            messagebox.showinfo("نجاح", "تم الحذف")
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')
            self.load_all()

    # ═══════════════════════════════════════════════════════════
    # التحميل والعرض
    # ═══════════════════════════════════════════════════════════
    def load_all(self):
        rows = get_all_recurring()
        self._populate_tree(rows)
        self._update_totals(rows)
        self.status_bar.config(text=f"عدد البنود: {len(rows)}")

    def _populate_tree(self, rows):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not rows:
            return

        for i, r in enumerate(rows):
            dept = r.get('department', '')
            if dept == 'مشتريات':
                tag = 'purchase_row'
            elif dept == 'مصاريف':
                tag = 'expense_row'
            elif dept == 'رواتب':
                tag = 'salary_row'
            else:
                tag = 'odd_row' if i % 2 == 0 else 'even_row'

            self.tree.insert('', 'end', iid=str(r['id']),
                              tags=(tag,), values=(
                                  r['id'],
                                  r.get('description', ''),
                                  f"{r.get('value', 0):,.2f}",
                                  dept,
                              ))

    def _update_totals(self, rows):
        total_value = sum(r.get('value', 0) or 0 for r in rows)
        total_purchases = sum(r.get('value', 0) or 0 for r in rows
                              if r.get('department') == 'مشتريات')
        total_expenses = sum(r.get('value', 0) or 0 for r in rows
                              if r.get('department') == 'مصاريف')

        self.total_value_label.config(text=f"{total_value:,.2f}")
        self.total_count_label.config(text=str(len(rows)))
        self.total_purchase_label.config(text=f"{total_purchases:,.2f}")
        self.total_expense_label.config(text=f"{total_expenses:,.2f}")

    # ═══════════════════════════════════════════════════════════
    # التحديد
    # ═══════════════════════════════════════════════════════════
    def on_select(self, event):
        sel = self.tree.selection()
        if sel:
            try:
                self.selected_id = int(sel[0])
            except (ValueError, TypeError):
                self.selected_id = None
                return
            self.btn_edit.config(state='normal')
            self.btn_delete.config(state='normal')
        else:
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')

    # ═══════════════════════════════════════════════════════════
    # الترحيل للشهر الحالي
    # ═══════════════════════════════════════════════════════════
    def transfer_now(self):
        """ترحيل كل المصاريف الدورية للشهر الحالي"""
        month_year = datetime.today().strftime('%Y-%m')

        if not messagebox.askyesno(
                "تأكيد الترحيل",
                f"هل تريد ترحيل المصاريف الدورية للشهر {month_year}؟\n\n"
                f"سيتم إضافتها إلى قائمة المصاريف.\n"
                f"(إذا كانت مضافة مسبقاً لن تُكرر)"):
            return

        added = apply_recurring_expenses(month_year)

        if added == -1:
            messagebox.showerror("خطأ", "حدث خطأ أثناء الترحيل")
        elif added == 0:
            messagebox.showinfo(
                "تنبيه",
                f"لا توجد مصاريف دورية جديدة للشهر {month_year}\n"
                f"(قد تكون مضافة مسبقاً)")
        else:
            messagebox.showinfo(
                "نجاح",
                f"✅ تم ترحيل {added} مصروف دوري للشهر {month_year}\n\n"
                f"يمكنك رؤيتها في صفحة المصاريف")
            self.load_all()

# ═══════════════════════════════════════════════════════════════════
# ADDITIONAL EXPENSE FORM WINDOW
# ═══════════════════════════════════════════════════════════════════
class AdditionalExpenseFormWindow:
    """نافذة إدخال / تعديل مصروف إضافي"""

    def __init__(self, parent_page, expense_id=None):
        self.parent_page = parent_page
        self.expense_id = expense_id
        self.is_edit = expense_id is not None
        self.image_path = None
        self.image_preview = None

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = "تعديل مصروف إضافي" if self.is_edit else "إضافة مصروف إضافي"
        self.window.title(title)
        self.window.geometry("700x450")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ─── العنوان ───
        title_frame = tk.Frame(self.window, bg='#8B4513', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icon = "✏" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'),
                 bg='#8B4513', fg='white').pack(pady=15)

        # ─── الإطار الرئيسي ───
        main = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=20)
        main.pack(fill='both', expand=True)

        # ─── التاريخ ───
        tk.Label(main, text="التاريخ: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=0, column=0, padx=5, pady=10, sticky='e')

        self.entry_date = DateEntry(main, width=20,
                                     date_pattern='yyyy-mm-dd',
                                     font=('Arial', 11),
                                     justify='center')
        self.entry_date.grid(row=0, column=1, padx=5, pady=10, sticky='w')
        self.entry_date.set_date(datetime.now())

        # ─── المبلغ ───
        tk.Label(main, text="المبلغ (ج.م): *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#c0392b', anchor='e').grid(
                 row=1, column=0, padx=5, pady=10, sticky='e')

        self.entry_amount = tk.Entry(main, width=22,
                                      font=('Arial', 13),
                                      justify='center',
                                      bg='#fadbd8')
        self.entry_amount.grid(row=1, column=1, padx=5, pady=10, sticky='w')

        # ─── الوصف ───
        tk.Label(main, text="الوصف: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=2, column=0, padx=5, pady=10, sticky='e')

        self.entry_desc = tk.Entry(main, width=45,
                                    font=('Arial', 11))
        self.entry_desc.grid(row=2, column=1, padx=5, pady=10, sticky='w')

        # ─── الصورة ───
        img_frame = tk.LabelFrame(main, text="صورة الإيصال",
                                   font=('Arial', 10, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=8, pady=8)
        img_frame.grid(row=0, column=2, rowspan=3, padx=15, pady=5, sticky='n')

        img_container = tk.Frame(img_frame, width=130, height=130,
                                  bg='lightgray', relief='ridge', bd=2)
        img_container.pack()
        img_container.grid_propagate(False)

        self.image_label = tk.Label(img_container, bg='lightgray',
                                     text='\nلا توجد صورة\n',
                                     font=('Arial', 11))
        self.image_label.pack(fill='both', expand=True)
        self.image_label.bind("<Double-Button-1>", self._open_image)

        tk.Button(img_frame, text="📁 اختيار",
                  command=self._select_image,
                  bg='#3498db', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=5).pack(pady=5)

        tk.Button(img_frame, text="🗑 إزالة",
                  command=self._remove_image,
                  bg='#e74c3c', fg='white',
                  font=('Arial', 9),
                  padx=10, pady=3).pack()

        # ─── أزرار الحفظ ───
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ",
                  command=self.save,
                  bg='#27ae60', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء",
                  command=self.window.destroy,
                  bg='#e74c3c', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        if self.is_edit:
            self._load_data()

        self.entry_amount.focus()

    def _select_image(self):
        path = filedialog.askopenfilename(
            title="اختر الصورة",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.bmp"),
                       ("All files", "*.*")])
        if path:
            self.image_path = path
            self._show_preview(path)

    def _show_preview(self, path):
        try:
            from PIL import Image, ImageTk
            img = Image.open(path)
            img.thumbnail((130, 130), Image.Resampling.LANCZOS)
            self.image_preview = ImageTk.PhotoImage(img)
            self.image_label.config(image=self.image_preview, text='')
        except Exception as e:
            print(f"⚠️ خطأ في الصورة: {e}")

    def _remove_image(self):
        self.image_path = "REMOVE"
        self.image_label.config(image='', text='\nلا توجد صورة\n')
        self.image_preview = None

    def _open_image(self, event=None):
        if not self.image_path or self.image_path == "REMOVE":
            return
        if not os.path.exists(self.image_path):
            return
        try:
            os.startfile(self.image_path)
        except:
            import subprocess
            subprocess.Popen(['start', self.image_path], shell=True)

    def _load_data(self):
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute("""SELECT expense_date, description, amount, image_path
                     FROM additional_expenses WHERE id=?""",
                  (self.expense_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            messagebox.showerror("خطأ", "السجل غير موجود")
            self.window.destroy()
            return

        date, desc, amount, img_path = row

        try:
            if date:
                self.entry_date.set_date(datetime.strptime(date, '%Y-%m-%d'))
        except:
            pass

        self.entry_amount.insert(0, str(amount or 0))
        self.entry_desc.insert(0, desc or '')

        if img_path and os.path.exists(img_path):
            self.image_path = img_path
            self._show_preview(img_path)

    def save(self):
        date = self.entry_date.get().strip()
        try:
            amount = float(self.entry_amount.get() or 0)
        except ValueError:
            messagebox.showerror("خطأ", "أدخل مبلغاً صحيحاً")
            return

        desc = self.entry_desc.get().strip()

        if not desc:
            messagebox.showwarning("تنبيه", "الوصف مطلوب")
            return

        if amount <= 0:
            messagebox.showwarning("تنبيه", "المبلغ يجب أن يكون أكبر من صفر")
            return

        if self.is_edit:
            ok = update_additional_expense(self.expense_id, date, desc,
                                            amount, self.image_path)
            msg = "تم التعديل"
        else:
            ok = add_additional_expense(date, desc, amount, self.image_path)
            msg = "تمت الإضافة"

        if ok:
            messagebox.showinfo("نجاح", msg)
            self.window.destroy()
            self.parent_page.load_all_data()
        else:
            messagebox.showerror("خطأ", "فشل الحفظ")

# ═══════════════════════════════════════════════════════════════════
# PROFIT TRANSFER PAGE
# ═══════════════════════════════════════════════════════════════════
class ProfitTransferPage(tk.Frame):
    """صفحة إدارة الأرباح المرحلة والمصاريف الإضافية"""

    def __init__(self, parent, controller):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller
        self.selected_transfer_id = None
        self.selected_expense_id = None

        # ═══ 1) الشريط العلوي ═══
        add_program_banner(self)

        # ═══ 2) شريط الحالة (أسفل) ═══
        self.status_bar = tk.Label(self, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#2c3e50', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # ═══ 3) منطقة التمرير ═══
        content = make_scrollable(self)
        self._page_content = content

        # ═══ 4) العنوان ═══
        tk.Label(content, text="💰 إدارة الأرباح والمصاريف الإضافية",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=10)

        # ═══ 5) بطاقات الإجماليات (3 بطاقات) ═══
        cards_frame = tk.Frame(content, bg='#34495e')
        cards_frame.pack(fill='x', padx=15, pady=(0, 15))

        self.total_profit_label = self._make_card(
            cards_frame, "💰 إجمالي الأرباح", '#2ecc71')
        self.total_expenses_label = self._make_card(
            cards_frame, "💸 إجمالي المصاريف الإضافية", '#e74c3c')
        self.remaining_label = self._make_card(
            cards_frame, "✅ المتبقي (صافي الأرباح)", '#f1c40f')

        # ═════════════════════════════════════════════════════════
        # 6) قسم: الأرباح المرحلة
        # ═════════════════════════════════════════════════════════
        transfers_frame = tk.LabelFrame(
            content, text="📊 الأرباح المرحلة (من البيان الختامي)",
            font=('Arial', 12, 'bold'),
            bg='#ecf0f1', fg='#8B4513',
            padx=10, pady=10)
        transfers_frame.pack(fill='x', padx=15, pady=(0, 15))

        # ─── جدول الأرباح ───
        tree1_frame = tk.Frame(transfers_frame, bg='#ecf0f1')
        tree1_frame.pack(fill='x')

        columns1 = ('id', 'month', 'net_amount', 'notes')
        self.tree_transfers = ttk.Treeview(tree1_frame,
                                            columns=columns1,
                                            show='headings', height=6)

        headers1 = {
            'id': 'ID',
            'month': 'الشهر',
            'net_amount': 'الصافي المرحل (ج.م)',
            'notes': 'ملاحظات'
        }
        for col, h in headers1.items():
            self.tree_transfers.heading(col, text=h)

        self.tree_transfers.column('id', width=50, anchor='center')
        self.tree_transfers.column('month', width=120, anchor='center')
        self.tree_transfers.column('net_amount', width=200, anchor='center')
        self.tree_transfers.column('notes', width=350, anchor='w')

        self.tree_transfers.tag_configure('odd_row', background='#f8f9fa')
        self.tree_transfers.tag_configure('even_row', background='#ffffff')

        sb1 = ttk.Scrollbar(tree1_frame, orient='vertical',
                             command=self.tree_transfers.yview)
        self.tree_transfers.configure(yscrollcommand=sb1.set)
        self.tree_transfers.pack(side='left', fill='x', expand=True)
        sb1.pack(side='right', fill='y')

        self.tree_transfers.bind('<<TreeviewSelect>>',
                                  self.on_transfer_select)

        # ═════════════════════════════════════════════════════════
        # 7) قسم: المصاريف الإضافية
        # ═════════════════════════════════════════════════════════
        expenses_frame = tk.LabelFrame(
            content, text="💸 المصاريف الإضافية (من الأرباح)",
            font=('Arial', 12, 'bold'),
            bg='#ecf0f1', fg='#c0392b',
            padx=10, pady=10)
        expenses_frame.pack(fill='x', padx=15, pady=(0, 10))

        # ─── إطار الأزرار ───
        btn_frame = tk.Frame(expenses_frame, bg='#ecf0f1')
        btn_frame.pack(fill='x', pady=(0, 8))

        tk.Button(btn_frame, text="➕ إضافة مصروف",
                  command=self.add_expense,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        self.btn_edit = tk.Button(btn_frame, text="✏ تعديل",
                                   command=self.edit_expense,
                                   bg='#f39c12', fg='white',
                                   font=('Arial', 11, 'bold'),
                                   padx=20, pady=8, width=14,
                                   state='disabled')
        self.btn_edit.pack(side='left', padx=5)

        self.btn_delete = tk.Button(btn_frame, text="🗑 حذف",
                                     command=self.delete_expense,
                                     bg='#e74c3c', fg='white',
                                     font=('Arial', 11, 'bold'),
                                     padx=20, pady=8, width=14,
                                     state='disabled')
        self.btn_delete.pack(side='left', padx=5)

        tk.Button(btn_frame, text="🔄 تحديث",
                  command=self.load_all_data,
                  bg='#95a5a6', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=14).pack(side='left', padx=5)

        # ─── جدول المصاريف ───
        tree2_frame = tk.Frame(expenses_frame, bg='#ecf0f1')
        tree2_frame.pack(fill='x')

        columns2 = ('id', 'date', 'description', 'amount', 'image')
        self.tree_expenses = ttk.Treeview(tree2_frame,
                                           columns=columns2,
                                           show='headings', height=6)

        headers2 = {
            'id': 'ID',
            'date': 'التاريخ',
            'description': 'الوصف',
            'amount': 'المبلغ (ج.م)',
            'image': 'الصورة'
        }
        for col, h in headers2.items():
            self.tree_expenses.heading(col, text=h)

        self.tree_expenses.column('id', width=50, anchor='center')
        self.tree_expenses.column('date', width=120, anchor='center')
        self.tree_expenses.column('description', width=350, anchor='w')
        self.tree_expenses.column('amount', width=150, anchor='center')
        self.tree_expenses.column('image', width=150, anchor='center')

        self.tree_expenses.tag_configure('odd_row', background='#f8f9fa')
        self.tree_expenses.tag_configure('even_row', background='#ffffff')

        sb2 = ttk.Scrollbar(tree2_frame, orient='vertical',
                             command=self.tree_expenses.yview)
        self.tree_expenses.configure(yscrollcommand=sb2.set)
        self.tree_expenses.pack(side='left', fill='x', expand=True)
        sb2.pack(side='right', fill='y')

        self.tree_expenses.bind('<<TreeviewSelect>>',
                                 self.on_expense_select)
        self.tree_expenses.bind('<Double-Button-1>',
                                 lambda e: self.edit_expense())

        # ─── تحميل أولي ───
        self.load_all_data()

    # ═══════════════════════════════════════════════════════════
    def _make_card(self, parent, title, color):
        """إنشاء بطاقة إجمالي"""
        card = tk.Frame(parent, bg='#34495e')
        card.pack(side='right', expand=True, padx=8, pady=8)

        tk.Label(card, text=title, font=('Arial', 11, 'bold'),
                 bg='#34495e', fg=color).pack(pady=(8, 2))

        lbl = tk.Label(card, text="0.00", font=('Arial', 16, 'bold'),
                        bg='#34495e', fg='white')
        lbl.pack(pady=(0, 8))
        return lbl

    # ═══════════════════════════════════════════════════════════
    # تحميل البيانات
    # ═══════════════════════════════════════════════════════════
    def load_all_data(self):
        """تحميل كل البيانات"""
        # 1) الأرباح المرحلة
        transfers = get_all_profit_transfers()
        for item in self.tree_transfers.get_children():
            self.tree_transfers.delete(item)

        for i, r in enumerate(transfers):
            tag = 'odd_row' if i % 2 == 0 else 'even_row'
            self.tree_transfers.insert('', 'end', iid=str(r['id']),
                                        tags=(tag,), values=(
                                            r['id'],
                                            r['month_year'],
                                            f"{r['net_amount']:,.2f}",
                                            r.get('notes', '') or ''
                                        ))

        # 2) المصاريف الإضافية
        expenses = get_all_additional_expenses()
        for item in self.tree_expenses.get_children():
            self.tree_expenses.delete(item)

        for i, r in enumerate(expenses):
            img_path = r.get('image_path') or ''
            img_name = os.path.basename(img_path) if img_path else ''
            tag = 'odd_row' if i % 2 == 0 else 'even_row'

            self.tree_expenses.insert('', 'end', iid=str(r['id']),
                                       tags=(tag,), values=(
                                           r['id'],
                                           r.get('expense_date', ''),
                                           r.get('description', ''),
                                           f"{r.get('amount', 0):,.2f}",
                                           img_name
                                       ))

        # 3) تحديث الإجماليات
        self._update_totals()

        # 4) شريط الحالة
        self.status_bar.config(
            text=f"عدد الأرباح: {len(transfers)} | "
                 f"عدد المصاريف: {len(expenses)}")

    def _update_totals(self):
        """تحديث الإجماليات"""
        total_profit = get_total_profits()
        total_expenses = get_total_additional_expenses()
        remaining = total_profit - total_expenses

        self.total_profit_label.config(text=f"{total_profit:,.2f}")
        self.total_expenses_label.config(text=f"{total_expenses:,.2f}")
        self.remaining_label.config(
            text=f"{remaining:,.2f}",
            fg='#2ecc71' if remaining >= 0 else '#e74c3c'
        )

    # ═══════════════════════════════════════════════════════════
    # اختيار الصفوف
    # ═══════════════════════════════════════════════════════════
    def on_transfer_select(self, event):
        """عند تحديد صف في جدول الأرباح"""
        sel = self.tree_transfers.selection()
        if sel:
            try:
                self.selected_transfer_id = int(sel[0])
            except:
                self.selected_transfer_id = None
        else:
            self.selected_transfer_id = None

    def on_expense_select(self, event):
        """عند تحديد صف في جدول المصاريف"""
        sel = self.tree_expenses.selection()
        if sel:
            try:
                self.selected_expense_id = int(sel[0])
            except:
                self.selected_expense_id = None
                return
            self.btn_edit.config(state='normal')
            self.btn_delete.config(state='normal')
        else:
            self.selected_expense_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')

    # ═══════════════════════════════════════════════════════════
    # CRUD للمصاريف الإضافية
    # ═══════════════════════════════════════════════════════════
    def add_expense(self):
        AdditionalExpenseFormWindow(self, expense_id=None)

    def edit_expense(self):
        if self.selected_expense_id is None:
            messagebox.showwarning("تنبيه", "اختر مصروفاً أولاً")
            return
        AdditionalExpenseFormWindow(self, expense_id=self.selected_expense_id)

    def delete_expense(self):
        if self.selected_expense_id is None:
            messagebox.showwarning("تنبيه", "اختر مصروفاً أولاً")
            return

        if not messagebox.askyesno(
                "تأكيد",
                f"هل أنت متأكد من حذف هذا المصروف؟"):
            return

        if delete_additional_expense(self.selected_expense_id):
            messagebox.showinfo("نجاح", "تم الحذف")
            self.selected_expense_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')
            self.load_all_data()
#☀️ Sunrise System
# ═══════════════════════════════════════════════════════════════════
# دالة الشريط العلوي (البانر)
# ═══════════════════════════════════════════════════════════════════


def add_program_banner(parent, bg='#1a5276', fg_title='#f1c40f',
                       logo_path=None, logo_size=70):

    """[شريط علوي احترافي: شعار يسار + إعلان منتصف + اسم البرنامج يمين]"""
    if logo_path is None:
        logo_path = COMPANY_LOGO_PATH

    # ═══════════════════════════════════════════════════════
    # الألوان
    # ═══════════════════════════════════════════════════════
    BG_MAIN     = '#1a5276'      # أزرق داكن (رئيسي)
    BG_LOGO     = '#154360'      # أغمق للشعار
    BG_AD       = '#154360'      # أغمق للمنتصف
    BG_TEXT     = '#154360'      # للمعلومات
    SEPARATOR   = '#f1c40f'      # فاصل ذهبي

    # ارتفاع البانر
    banner = tk.Frame(parent, bg=BG_MAIN, height=90)
    banner.pack(fill='x', side='top')
    banner.pack_propagate(False)

    # ═══════════════════════════════════════════════════════
    # 1) الشعار على اليسار (بخلفية داكنة قليلًا)
    # ═══════════════════════════════════════════════════════
    logo_frame = tk.Frame(banner, bg=BG_LOGO, width=120)
    logo_frame.pack(side='left', fill='y')
    logo_frame.pack_propagate(False)

    if logo_path and os.path.exists(logo_path):
        try:
            logo_img = Image.open(logo_path)
            if logo_img.mode != 'RGBA':
                logo_img = logo_img.convert('RGBA')
            logo_img.thumbnail((logo_size, logo_size), Image.Resampling.LANCZOS)
            logo_frame._logo_photo = ImageTk.PhotoImage(logo_img)
            tk.Label(logo_frame, image=logo_frame._logo_photo,
                     bg=BG_LOGO).pack(expand=True)
        except Exception as e:
            print(f"⚠️ [لم يتم تحميل اللوجو]: {e}")

    # ─── فاصل ذهبي ───
   # tk.Frame(banner, bg=SEPARATOR, width=2).pack(side='left', fill='y')

    # ═══════════════════════════════════════════════════════
    # 2) اسم البرنامج والشركة على اليمين
    # ═══════════════════════════════════════════════════════
    text_frame = tk.Frame(banner, bg=BG_TEXT)
    text_frame.pack(side='right', fill='y', padx=0)

    # سطر أول: اسم البرنامج
    tk.Label(text_frame, text="☀️ Sunrise System",
             font=('Arial', 17, 'bold'),
             bg=BG_TEXT, fg=fg_title).pack(anchor='c', pady=(15, 2))

    # خط فاصل أفقي رفيع
#    tk.Frame(text_frame, bg=fg_title, height=1).pack(
  #      fill='x', pady=4)

    # سطر ثاني: اسم الشركة
    tk.Label(text_frame,
             text="شركة الزهراء للبرمجيات - المحاسب / حسام الدين سيد ربيع 📞  01003796984",
             font=('Arial', 11, 'bold'),
             bg=BG_TEXT, fg='white').pack(anchor='e')

    # سطر ثالث: رقم الهاتف
    tk.Label(text_frame,
             text="📞  01003796984",
             font=('Arial', 10, 'bold'),
             bg=BG_TEXT, fg='#bdc3c7').pack(anchor='e', pady=(2, 15))

    # ─── فاصل ذهبي ───
   # tk.Frame(banner, bg=SEPARATOR, width=2).pack(side='right', fill='y')

    # ═══════════════════════════════════════════════════════
    # 3) الإعلان المتحرك في المنتصف
    # ═══════════════════════════════════════════════════════
    ad_frame = tk.Frame(banner, bg=BG_AD)
    ad_frame.pack(side='left', fill='both', expand=True)

    ad_label = tk.Label(
        ad_frame,
        text=" " * AD_DISPLAY_LENGTH,
        font=('Arial', 11, 'bold'),
        bg=BG_AD,
        fg=fg_title,
        anchor='center',
    )
    ad_label.pack(expand=True, fill='both')

    # ✅ سجّل الليبل لتحديثه من MainApp
    BANNER_AD_LABELS.append(ad_label)

    return banner



# ═══════════════════════════════════════════════════════════════════
# دالة منطقة التمرير
# ═══════════════════════════════════════════════════════════════════
def make_scrollable(page_frame, bg_color='#ecf0f1'):
    """يحوّل الإطار إلى منطقة قابلة للتمرير"""
    page_frame.configure(bg=bg_color)

    scroll_holder = tk.Frame(page_frame, bg=bg_color)
    scroll_holder.pack(fill='both', expand=True, side='top')

    canvas = tk.Canvas(scroll_holder, bg=bg_color,
                       highlightthickness=0, bd=0)
    vsb = ttk.Scrollbar(scroll_holder, orient='vertical',
                        command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)

    vsb.pack(side='right', fill='y')
    canvas.pack(side='left', fill='both', expand=True)

    content = tk.Frame(canvas, bg=bg_color)
    win_id = canvas.create_window((0, 0), window=content,
                                  anchor='nw', tags='content')

    def _sync_width(event):
        canvas.itemconfig(win_id, width=event.width)

    canvas.bind('<Configure>', _sync_width)

    def _update_scrollregion(event=None):
        canvas.configure(scrollregion=canvas.bbox('all'))

    content.bind('<Configure>', _update_scrollregion)

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

    canvas.bind_all('<MouseWheel>', _on_mousewheel)

    page_frame._canvas = canvas
    page_frame._content = content
    return content
# ═══════════════════════════════════════════════════════════════════
# OWNER MONEY FORM WINDOW
# ═══════════════════════════════════════════════════════════════════
class OwnerMoneyFormWindow:
    """نافذة إدخال/تعديل حركة مبلغ مسلّم من صاحب المكان"""

    def __init__(self, parent_page, trans_id=None,
                 default_type='استلام'):
        self.parent_page = parent_page
        self.trans_id = trans_id
        self.is_edit = trans_id is not None
        self.image_path = None
        self.image_preview = None

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = ("تعديل حركة" if self.is_edit
                 else "إضافة مبلغ من صاحب المكان")
        self.window.title(title)
        self.window.geometry("640x520")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ─── العنوان ───
        title_frame = tk.Frame(self.window, bg='#1a5276', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icon = "✏" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'),
                 bg='#1a5276', fg='#f1c40f').pack(pady=15)

        # ─── الإطار الرئيسي ───
        main = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=15)
        main.pack(fill='both', expand=True)

        # ─── التاريخ ───
        self.entry_date = DateEntry(main, width=18,
                                    date_pattern='yyyy-mm-dd',
                                    font=('Arial', 11),
                                    justify='center')
        self.entry_date.pack(side='right', padx=5)
        self.entry_date.set_date(datetime.now())

        # ─── المكان ───
        tk.Label(main, text="المكان: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='right', padx=15)

        self.place_var = tk.StringVar(value="bazar")
        ttk.Combobox(main, textvariable=self.place_var,
                     values=('bazar', 'beach', 'other'),
                     state='readonly', width=12,
                     font=('Arial', 11)).pack(side='right', padx=5)

        # ─── نوع الحركة ───
        row1 = tk.Frame(main, bg='#ecf0f1')
        row1.pack(fill='x', pady=(0, 10))

        tk.Label(row1, text="نوع الحركة: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#c0392b').pack(side='right', padx=5)

        self.type_var = tk.StringVar(value=default_type)
        type_combo = ttk.Combobox(row1, textvariable=self.type_var,
                                   values=('استلام', 'رد'),
                                   state='readonly', width=15,
                                   font=('Arial', 12, 'bold'))
        type_combo.pack(side='right', padx=5)
        type_combo.bind('<<ComboboxSelected>>', self._on_type_change)

        self.type_hint = tk.Label(main, text="",
                                   font=('Arial', 9, 'italic'),
                                   bg='#ecf0f1', fg='#7f8c8d')
        self.type_hint.pack(fill='x', pady=(0, 10))
        self._on_type_change()

        # ─── المبلغ ───
        row2 = tk.Frame(main, bg='#ecf0f1')
        row2.pack(fill='x', pady=(0, 10))

        tk.Label(row2, text="المبلغ (ج.م): *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#c0392b').pack(side='right', padx=5)

        self.entry_amount = tk.Entry(row2, width=20,
                                      font=('Arial', 14, 'bold'),
                                      justify='center',
                                      bg='#fadbd8')
        self.entry_amount.pack(side='right', padx=5)

        # ─── الوصف ───
        tk.Label(main, text="الوصف / البيان:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').pack(fill='x', pady=(0, 5))

        self.entry_desc = tk.Text(main, height=3,
                                   font=('Arial', 11),
                                   bg='white')
        self.entry_desc.pack(fill='x', pady=(0, 10))

        # ─── الصورة ───
        img_frame = tk.LabelFrame(main, text="صورة الإيصال (اختياري)",
                                   font=('Arial', 10, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=8, pady=8)
        img_frame.pack(fill='x', pady=(0, 10))

        img_row = tk.Frame(img_frame, bg='#ecf0f1')
        img_row.pack()

        img_container = tk.Frame(img_row, width=120, height=120,
                                  bg='lightgray', relief='ridge', bd=2)
        img_container.pack(side='right', padx=10)
        img_container.grid_propagate(False)

        self.image_label = tk.Label(img_container, bg='lightgray',
                                     text='\nلا توجد صورة\n',
                                     font=('Arial', 10))
        self.image_label.pack(fill='both', expand=True)
        self.image_label.bind("<Double-Button-1>", self._open_image)

        btn_col = tk.Frame(img_row, bg='#ecf0f1')
        btn_col.pack(side='right')

        tk.Button(btn_col, text="📁 اختيار صورة",
                  command=self._select_image,
                  bg='#3498db', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=5).pack(pady=3)

        tk.Button(btn_col, text="🗑 إزالة",
                  command=self._remove_image,
                  bg='#e74c3c', fg='white',
                  font=('Arial', 9),
                  padx=10, pady=3).pack()

        # ─── الأزرار ───
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ",
                  command=self.save,
                  bg='#27ae60', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء",
                  command=self.window.destroy,
                  bg='#e74c3c', fg='white',
                  padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        if self.is_edit:
            self._load_data()

        self.entry_amount.focus()

    def _on_type_change(self, event=None):
        hints = {
            'استلام': '💰 مبلغ استلمناه من صاحب المكان (يزيد الرصيد)',
            'رد':    '↩️ مبلغ رددناه لصاحب المكان (ينقص الرصيد)',
        }
        self.type_hint.config(text=hints.get(self.type_var.get(), ''))

    def _select_image(self):
        path = filedialog.askopenfilename(
            title="اختر الصورة",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.bmp"),
                       ("All files", "*.*")])
        if path:
            self.image_path = path
            self._show_preview(path)

    def _show_preview(self, path):
        try:
            img = Image.open(path)
            img.thumbnail((120, 120), Image.Resampling.LANCZOS)
            self.image_preview = ImageTk.PhotoImage(img)
            self.image_label.config(image=self.image_preview, text='')
        except Exception as e:
            print(f"⚠️ خطأ: {e}")

    def _remove_image(self):
        self.image_path = "REMOVE"
        self.image_label.config(image='', text='\nلا توجد صورة\n')
        self.image_preview = None

    def _open_image(self, event=None):
        if not self.image_path or self.image_path == "REMOVE":
            return
        if not os.path.exists(self.image_path):
            return
        try:
            os.startfile(self.image_path)
        except:
            import subprocess
            subprocess.Popen(['start', self.image_path], shell=True)

    def _load_data(self):
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute("""SELECT date, place, amount, transaction_type,
                            description, image_path
                     FROM owner_money WHERE id=?""", (self.trans_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            ...

        date, place, amount, ttype, desc, img_path = row

        try:
            if date:
                self.entry_date.set_date(datetime.strptime(date, '%Y-%m-%d'))
        except:
            pass

        self.type_var.set(ttype or 'استلام')
        self.place_var.set(place or 'bazar')

        self._on_type_change()

        self.entry_amount.delete(0, tk.END)
        self.entry_amount.insert(0, str(amount or 0))

        self.entry_desc.delete('1.0', tk.END)
        self.entry_desc.insert('1.0', desc or '')

        if img_path and os.path.exists(img_path):
            self.image_path = img_path
            self._show_preview(img_path)

    def save(self):
        date = self.entry_date.get().strip()
        place = self.place_var.get().strip()
        ttype = self.type_var.get().strip()
        desc = self.entry_desc.get('1.0', 'end').strip()

        if not place:
            messagebox.showwarning("تنبيه", "المكان مطلوب")
            return

        try:
            amount = float(self.entry_amount.get() or 0)
        except ValueError:
            messagebox.showerror("خطأ", "أدخل مبلغاً صحيحاً")
            return

        if not date:
            messagebox.showwarning("تنبيه", "التاريخ مطلوب")
            return

        if amount <= 0:
            messagebox.showwarning("تنبيه", "المبلغ يجب أن يكون أكبر من صفر")
            return

        # ✅ التعيين قبل الفحص
        if self.is_edit:
            ok = update_owner_money(self.trans_id, date, place, amount, ttype,
                                    desc, self.image_path)
            success_msg = "تم التعديل"
            error_msg = "فشل التعديل"
        else:
            ok = add_owner_money(date, place, amount, ttype,
                                 desc, self.image_path)
            success_msg = "تمت الإضافة"
            error_msg = "فشل الإضافة"

        # ✅ الآن safe
        if ok:
            messagebox.showinfo("نجاح", success_msg)
            self.window.destroy()
            self.parent_page.load_all()
        else:
            messagebox.showerror("خطأ", error_msg)


# ═══════════════════════════════════════════════════════════════════
# OWNER MONEY PAGE (المبالغ المسلمة من صاحب المكان)
# ═══════════════════════════════════════════════════════════════════
class OwnerMoneyPage(tk.Frame):
    """صفحة المبالغ المسلمة من صاحب المكان"""

    def __init__(self, parent, controller):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller
        self.selected_id = None

        # ═══ 1) البانر ═══
        add_program_banner(self)

        # ═══ 2) شريط الحالة ═══
        self.status_bar = tk.Label(self, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#2c3e50', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # ═══ 3) منطقة التمرير ═══
        content = make_scrollable(self)
        self._page_content = content

        # ═══ 4) العنوان ═══
        tk.Label(content, text="💰 المبالغ المسلمة من صاحب المكان",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=10)

        tk.Label(content,
                 text="تسجيل المبالغ المستلمة من صاحب المكان لاستخدامها في المصاريف",
                 font=('Arial', 11, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=(0, 10))

        # ═══════════════════════════════════════════════════════
        # 5) بطاقة الرصيد الحالي
        # ═══════════════════════════════════════════════════════
        balance_card = tk.Frame(content, bg='#1a5276', height=140)
        balance_card.pack(fill='x', padx=15, pady=(0, 15))
        balance_card.pack_propagate(False)

        tk.Label(balance_card,
                 text="💰 الرصيد الحالي من صاحب المكان",
                 font=('Arial', 13, 'bold'),
                 bg='#1a5276', fg='#f1c40f').pack(pady=(12, 5))

        self.lbl_balance = tk.Label(balance_card, text="0.00 ج.م",
                                     font=('Arial', 26, 'bold'),
                                     bg='#1a5276', fg='white')
        self.lbl_balance.pack()

        details_row = tk.Frame(balance_card, bg='#1a5276')
        details_row.pack(pady=(8, 0))

        self.lbl_received = tk.Label(details_row,
                                      text="إجمالي المستلم: 0.00",
                                      font=('Arial', 11, 'bold'),
                                      bg='#1a5276', fg='#2ecc71')
        self.lbl_received.pack(side='right', padx=15)

        self.lbl_returned = tk.Label(details_row,
                                      text="إجمالي المردود: 0.00",
                                      font=('Arial', 11, 'bold'),
                                      bg='#1a5276', fg='#85c1e9')
        self.lbl_returned.pack(side='right', padx=15)

        # ═══════════════════════════════════════════════════════
        # 6) إطار الأزرار
        # ═══════════════════════════════════════════════════════
        btn_frame = tk.LabelFrame(content, text="العمليات",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=15, pady=12)
        btn_frame.pack(fill='x', padx=15, pady=(0, 10))

        row1 = tk.Frame(btn_frame, bg='#ecf0f1')
        row1.pack(pady=3)

        tk.Button(row1, text="💰 استلام من صاحب المكان",
                  command=lambda: self.add_trans('استلام'),
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=22).pack(side='left', padx=3)

        tk.Button(row1, text="↩️ رد لصاحب المكان",
                  command=lambda: self.add_trans('رد'),
                  bg='#2980b9', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=18).pack(side='left', padx=3)

        self.btn_edit = tk.Button(row1, text="✏ تعديل",
                                   command=self.edit_trans,
                                   bg='#f39c12', fg='white',
                                   font=('Arial', 11, 'bold'),
                                   padx=15, pady=8, width=12,
                                   state='disabled')
        self.btn_edit.pack(side='left', padx=3)

        self.btn_delete = tk.Button(row1, text="🗑 حذف",
                                     command=self.delete_trans,
                                     bg='#e74c3c', fg='white',
                                     font=('Arial', 11, 'bold'),
                                     padx=15, pady=8, width=12,
                                     state='disabled')
        self.btn_delete.pack(side='left', padx=3)

        # ═══════════════════════════════════════════════════════
        # 7) إطار البحث
        # ═══════════════════════════════════════════════════════
        search_frame = tk.LabelFrame(content, text="فلاتر البحث",
                                      font=('Arial', 11, 'bold'),
                                      bg='#ecf0f1', fg='#34495e',
                                      padx=10, pady=10)
        search_frame.pack(fill='x', padx=15, pady=(0, 10))

        row_s = tk.Frame(search_frame, bg='#ecf0f1')
        row_s.pack(fill='x')

        tk.Label(row_s, text="🔍 بحث:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_entry = tk.Entry(row_s, width=18,
                                      font=('Arial', 11),
                                      justify='center')
        self.search_entry.pack(side='left', padx=5)

        tk.Label(row_s, text="من:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.start_date_entry = DateEntry(row_s, width=12,
                                           date_pattern='yyyy-mm-dd',
                                           font=('Arial', 10),
                                           justify='center')
        self.start_date_entry.pack(side='left', padx=5)

        tk.Label(row_s, text="إلى:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.end_date_entry = DateEntry(row_s, width=12,
                                         date_pattern='yyyy-mm-dd',
                                         font=('Arial', 10),
                                         justify='center')
        self.end_date_entry.pack(side='left', padx=5)

        tk.Label(row_s, text="المكان:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_place_var = tk.StringVar(value="الكل")
        ttk.Combobox(row_s, textvariable=self.search_place_var,
                     values=("الكل", "bazar", "beach", "other"),
                     state='readonly', width=10,
                     font=('Arial', 11)).pack(side='left', padx=5)

        tk.Label(row_s, text="النوع:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_type_var = tk.StringVar(value="الكل")
        ttk.Combobox(row_s, textvariable=self.search_type_var,
                     values=("الكل", "استلام", "رد"),
                     state='readonly', width=10,
                     font=('Arial', 11)).pack(side='left', padx=5)

        tk.Button(row_s, text="🔍 بحث",
                  command=self.search_trans,
                  bg='#9b59b6', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        tk.Button(row_s, text="🗑 مسح",
                  command=self.clear_search,
                  bg='#bdc3c7', fg='#2c3e50',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        tk.Button(row_s, text="📋 عرض الكل",
                  command=self.load_all,
                  bg='#3498db', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        tk.Button(row_s, text="💾 تصدير Excel",
                  command=self.export_to_excel,
                  bg='#16a085', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        # ═══════════════════════════════════════════════════════
        # 8) الجدول
        # ═══════════════════════════════════════════════════════
        tree_frame = tk.LabelFrame(content, text="سجل الحركات",
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 15))

        columns = ('id', 'date', 'place', 'type', 'amount',
                   'description', 'image')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                 show='headings', height=14)

        headers = {
            'id': 'ID',
            'date': 'التاريخ',
            'place': 'المكان',
            'type': 'نوع الحركة',
            'amount': 'المبلغ (ج.م)',
            'description': 'البيان',
            'image': 'الصورة'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=50, anchor='center')
        self.tree.column('date', width=110, anchor='center')
        self.tree.column('place', width=90, anchor='center')
        self.tree.column('type', width=160, anchor='center')
        self.tree.column('amount', width=140, anchor='center')
        self.tree.column('description', width=280, anchor='e')
        self.tree.column('image', width=120, anchor='center')

        self.tree.tag_configure('received',
                                 background='#e8f8f5',
                                 foreground='#145a32',
                                 font=('Arial', 11, 'bold'))
        self.tree.tag_configure('returned',
                                 background='#ebf5fb',
                                 foreground='#1a5276',
                                 font=('Arial', 11, 'bold'))
        # ✅ صف الإجمالي
        self.tree.tag_configure('total_row',
                                background='#2874a6',
                                foreground='#ffffff',
                                font=('Arial', 12, 'bold'))

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', self.on_double_click)

        # ─── تحميل ───
        self.load_all()

    # ═══════════════════════════════════════════════════════════
    def load_all(self):
        rows = get_all_owner_money()
        self._populate_tree(rows)
        self._update_balance()
        self.status_bar.config(text=f"عدد الحركات: {len(rows)}")

    def _populate_tree(self, rows):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not rows:
            return

        # ─── متغيرات الإجماليات ───
        total_received = 0
        total_returned = 0

        # ─── إدراج الصفوف ───
        for r in rows:
            ttype = r.get('transaction_type', '')
            amount = r.get('amount', 0) or 0

            if ttype == 'استلام':
                tag = 'received'
                type_display = '💰 استلام من صاحب المكان'
                amount_display = f"+ {amount:,.2f}"
                total_received += amount
            else:
                tag = 'returned'
                type_display = '↩️ رد لصاحب المكان'
                amount_display = f"- {amount:,.2f}"
                total_returned += amount

            img_path = r.get('image_path') or ''
            img_name = os.path.basename(img_path) if img_path else ''

            self.tree.insert('', 'end', iid=str(r['id']),
                             tags=(tag,), values=(
                    r['id'],
                    r.get('date', ''),
                    r.get('place', ''),
                    type_display,
                    amount_display,
                    r.get('description', '') or '',
                    img_name,
                ))

        # ═══════════════════════════════════════════════════════
        # ✅ صف الإجمالي في نهاية الجدول
        # ═══════════════════════════════════════════════════════
        net = total_received - total_returned

        self.tree.insert('', 'end', tags=('total_row',), values=(
            '',
            '',
            '',
            f'📊 الإجمالي ({len(rows)} حركة)',
            f"{net:+,.2f}",
            f"مستلم: {total_received:,.2f}  |  مردود: {total_returned:,.2f}",
            '',
        ))

    def _update_balance(self):
        bal = get_owner_money_balance()
        self.lbl_balance.config(text=f"{bal['balance']:,.2f} ج.م")
        self.lbl_received.config(
            text=f"إجمالي المستلم: {bal['received']:,.2f}")
        self.lbl_returned.config(
            text=f"إجمالي المردود: {bal['returned']:,.2f}")

    # ═══════════════════════════════════════════════════════════
    def add_trans(self, ttype):
        OwnerMoneyFormWindow(self, trans_id=None, default_type=ttype)

    def edit_trans(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر حركة أولاً")
            return
        OwnerMoneyFormWindow(self, trans_id=self.selected_id)

    def delete_trans(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر حركة أولاً")
            return
        if not messagebox.askyesno(
                "تأكيد",
                f"هل أنت متأكد من حذف الحركة رقم {self.selected_id}؟"):
            return

        if delete_owner_money(self.selected_id):
            messagebox.showinfo("نجاح", "تم الحذف")
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')
            self.load_all()

    def on_select(self, event):
        sel = self.tree.selection()
        if sel:
            try:
                self.selected_id = int(sel[0])
            except (ValueError, TypeError):
                self.selected_id = None
                return
            self.btn_edit.config(state='normal')
            self.btn_delete.config(state='normal')
        else:
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')

    def on_double_click(self, event):
        col = self.tree.identify_column(event.x)
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        if col == '#7':  # ← عمود الصورة الآن في الموضع السابع
            self._open_row_image()
        else:
            self.edit_trans()

    def _open_row_image(self):
        sel = self.tree.selection()
        if not sel:
            return
        try:
            trans_id = int(sel[0])
        except (ValueError, TypeError):
            return

        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute("SELECT image_path FROM owner_money WHERE id=?",
                  (trans_id,))
        row = c.fetchone()
        conn.close()

        if not row or not row[0]:
            messagebox.showinfo("معلومة", "لا توجد صورة لهذا السجل")
            return

        path = clean_path(row[0])
        if not path or not os.path.exists(path):
            messagebox.showerror("خطأ", "ملف الصورة غير موجود")
            return

        try:
            os.startfile(path)
        except:
            import subprocess
            subprocess.Popen(['start', path], shell=True)

    def search_trans(self):
        term = self.search_entry.get().strip()
        start = self.start_date_entry.get().strip()
        end = self.end_date_entry.get().strip()
        place = self.search_place_var.get().strip()
        ttype = self.search_type_var.get().strip()

        rows = search_owner_money(term, start, end, place, ttype)
        self._populate_tree(rows)
        self.status_bar.config(text=f"نتائج البحث: {len(rows)} سجل")
        if not rows:
            messagebox.showinfo("بحث", "لا توجد نتائج")

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        self.start_date_entry.delete(0, tk.END)
        self.end_date_entry.delete(0, tk.END)
        self.search_place_var.set("الكل")
        self.search_type_var.set("الكل")
        self.load_all()

    def export_to_excel(self):
        rows = get_all_owner_money()
        if not rows:
            messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
            return

        try:
            data_rows = []
            for r in rows:
                data_rows.append([
                    r.get('date', ''),
                    r.get('place', ''),
                    r.get('transaction_type', ''),
                    r.get('amount', 0),
                    r.get('description', '') or '',
                ])

            bal = get_owner_money_balance()
            data_rows.append(['الرصيد الحالي', '', bal['balance'], ''])

            headers = ["التاريخ", "المكان", "نوع الحركة", "المبلغ (ج.م)", "البيان"]
            title = "تقرير المبالغ المسلمة من صاحب المكان"

            success = append_report_to_master(data_rows, title, headers)
            if success:
                messagebox.showinfo("نجاح",
                                     f"تم إضافة التقرير إلى:\n{MASTER_FILE}")
            else:
                messagebox.showerror("خطأ", "فشل التصدير")
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"حدث خطأ: {e}")
# ═══════════════════════════════════════════════════════════════════
# TRADER FORM WINDOW
# ═══════════════════════════════════════════════════════════════════
class TraderFormWindow:
    def __init__(self, parent_page, trader_id=None):
        self.parent_page = parent_page
        self.trader_id = trader_id
        self.is_edit = trader_id is not None

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = "تعديل تاجر" if self.is_edit else "إضافة تاجر جديد"
        self.window.title(title)
        self.window.geometry("520x420")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ─── العنوان ───
        title_frame = tk.Frame(self.window, bg='#1a5276', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icon = "✏" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'),
                 bg='#1a5276', fg='#f1c40f').pack(pady=15)

        # ─── الإطار الرئيسي ───
        main = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=20)
        main.pack(fill='both', expand=True)

        # ─── اسم التاجر ───
        tk.Label(main, text="اسم التاجر: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=0, column=0, padx=5, pady=10, sticky='e')

        self.entry_name = tk.Entry(main, width=30, font=('Arial', 11))
        self.entry_name.grid(row=0, column=1, padx=5, pady=10, sticky='w')

        # ─── الهاتف ───
        tk.Label(main, text="الهاتف:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=1, column=0, padx=5, pady=10, sticky='e')

        self.entry_phone = tk.Entry(main, width=30, font=('Arial', 11))
        self.entry_phone.grid(row=1, column=1, padx=5, pady=10, sticky='w')

        # ─── العنوان ───
        tk.Label(main, text="العنوان:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=2, column=0, padx=5, pady=10, sticky='e')

        self.entry_address = tk.Entry(main, width=30, font=('Arial', 11))
        self.entry_address.grid(row=2, column=1, padx=5, pady=10, sticky='w')

        # ─── ملاحظات ───
        tk.Label(main, text="ملاحظات:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='ne').grid(
                 row=3, column=0, padx=5, pady=10, sticky='ne')

        self.entry_notes = tk.Text(main, width=30, height=4, font=('Arial', 11))
        self.entry_notes.grid(row=3, column=1, padx=5, pady=10, sticky='w')

        # ─── الأزرار ───
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ", command=self.save,
                  bg='#27ae60', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=30, pady=10).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء", command=self.window.destroy,
                  bg='#e74c3c', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=30, pady=10).pack(side='left', padx=5)

        if self.is_edit:
            self._load_data()

        self.entry_name.focus()

    def _load_data(self):
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('SELECT name, phone, address, notes FROM traders WHERE id=?',
                  (self.trader_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            messagebox.showerror("خطأ", "التاجر غير موجود")
            self.window.destroy()
            return

        self.entry_name.insert(0, row[0] or '')
        self.entry_phone.insert(0, row[1] or '')
        self.entry_address.insert(0, row[2] or '')
        self.entry_notes.insert('1.0', row[3] or '')

    def save(self):
        name = self.entry_name.get().strip()
        phone = self.entry_phone.get().strip()
        address = self.entry_address.get().strip()
        notes = self.entry_notes.get('1.0', 'end').strip()

        if not name:
            messagebox.showwarning("تنبيه", "اسم التاجر مطلوب")
            return

        if self.is_edit:
            ok = update_trader(self.trader_id, name, phone, address, notes)
            success_msg = "تم التعديل"
            error_msg = "فشل التعديل"
        else:
            ok = add_trader(name, phone, address, notes)
            success_msg = "تمت الإضافة"
            error_msg = "الاسم موجود مسبقاً"

        if ok:
            messagebox.showinfo("نجاح", success_msg)
            self.window.destroy()
            self.parent_page.load_all()
        else:
            messagebox.showerror("خطأ", error_msg)
# ═══════════════════════════════════════════════════════════════════
# TRADER INVOICE WINDOW (يدعم الإضافة والتعديل)
# ═══════════════════════════════════════════════════════════════════
class TraderInvoiceWindow:
    def __init__(self, parent_page, trader_id=None, invoice_id=None):
        self.parent_page = parent_page
        self.trader_id = trader_id
        self.invoice_id = invoice_id
        self.is_edit = invoice_id is not None
        self.items = []

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = "تعديل فاتورة تاجر" if self.is_edit else "إضافة فاتورة تاجر"
        self.window.title(title)
        self.window.geometry("820x730")
        self.window.grab_set()
        self.window.configure(bg='#ecf0f1')

        # ─── العنوان ───
        title_frame = tk.Frame(self.window, bg='#1a5276', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        icon = "✏" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'),
                 bg='#1a5276', fg='#f1c40f').pack(pady=15)

        # ─── الإطار الرئيسي ───
        main = tk.Frame(self.window, bg='#ecf0f1', padx=20, pady=15)
        main.pack(fill='both', expand=True)

        # ─── التاريخ + رقم الفاتورة ───
        row0 = tk.Frame(main, bg='#ecf0f1')
        row0.pack(fill='x', pady=(0, 10))

        tk.Label(row0, text="التاريخ: *",
                 font=('Arial', 11, 'bold'), bg='#ecf0f1').pack(side='right', padx=5)

        self.entry_date = DateEntry(row0, width=12,
                                     date_pattern='yyyy-mm-dd',
                                     font=('Arial', 11), justify='center')
        self.entry_date.pack(side='right', padx=5)
        self.entry_date.set_date(datetime.now())

        tk.Label(row0, text="رقم الفاتورة:",
                 font=('Arial', 11, 'bold'), bg='#ecf0f1').pack(side='right', padx=15)

        self.entry_invoice_no = tk.Entry(row0, width=15,
                                          font=('Arial', 11), justify='center')
        self.entry_invoice_no.pack(side='right', padx=5)

        # ─── التاجر + المكان ───
        row1 = tk.Frame(main, bg='#ecf0f1')
        row1.pack(fill='x', pady=(0, 10))

        tk.Label(row1, text="التاجر: *",
                 font=('Arial', 11, 'bold'), bg='#ecf0f1').pack(side='right', padx=5)

        self.trader_var = tk.StringVar()
        self.trader_combo = ttk.Combobox(row1, textvariable=self.trader_var,
                                          width=25, font=('Arial', 11),
                                          state='readonly')
        self.trader_combo.pack(side='right', padx=5)

        traders = get_all_traders()
        self.traders_map = {t['name']: t['id'] for t in traders}
        self.trader_combo['values'] = list(self.traders_map.keys())

        if not self.is_edit and trader_id:
            for t in traders:
                if t['id'] == trader_id:
                    self.trader_var.set(t['name'])
                    break
        elif traders and not self.is_edit:
            self.trader_var.set(traders[0]['name'])

        tk.Label(row1, text="المكان:",
                 font=('Arial', 11, 'bold'), bg='#ecf0f1').pack(side='right', padx=15)

        self.place_var = tk.StringVar(value="bazar")
        ttk.Combobox(row1, textvariable=self.place_var,
                     values=('bazar', 'beach', 'other'),
                     state='readonly', width=10,
                     font=('Arial', 11)).pack(side='right', padx=5)

        # ═══════════════════════════════════════════════════════
        # قسم: نوع الدفع
        # ═══════════════════════════════════════════════════════
        payment_frame = tk.LabelFrame(main, text="💰 نوع الدفع",
                                       font=('Arial', 11, 'bold'),
                                       bg='#fff9e6', fg='#8B4513',
                                       padx=15, pady=10)
        payment_frame.pack(fill='x', pady=(0, 10))

        # ─── الصف الأول: النوع ───
        row_p1 = tk.Frame(payment_frame, bg='#fff9e6')
        row_p1.pack(fill='x', pady=(0, 8))

        tk.Label(row_p1, text="نوع الدفع:",
                 font=('Arial', 11, 'bold'),
                 bg='#fff9e6', fg='#8B4513').pack(side='right', padx=5)

        self.payment_type_var = tk.StringVar(value='أجل')
        payment_combo = ttk.Combobox(row_p1, textvariable=self.payment_type_var,
                                      values=('كاش', 'جزئي', 'أجل'),
                                      state='readonly', width=15,
                                      font=('Arial', 12, 'bold'))
        payment_combo.pack(side='right', padx=5)
        payment_combo.bind('<<ComboboxSelected>>', self._on_payment_type_change)

        # ─── الصف الثاني: المبلغ المدفوع ───
        row_p2 = tk.Frame(payment_frame, bg='#fff9e6')
        row_p2.pack(fill='x')

        tk.Label(row_p2, text="المبلغ المدفوع:",
                 font=('Arial', 11, 'bold'),
                 bg='#fff9e6', fg='#8B4513').pack(side='right', padx=5)

        self.entry_paid_amount = tk.Entry(row_p2, width=15,
                                           font=('Arial', 12, 'bold'),
                                           justify='center',
                                           bg='#fadbd8')
        self.entry_paid_amount.pack(side='right', padx=5)
        self.entry_paid_amount.insert(0, "0")

        self.lbl_remaining = tk.Label(row_p2, text="",
                                       font=('Arial', 11, 'bold'),
                                       bg='#fff9e6', fg='#c0392b')
        self.lbl_remaining.pack(side='right', padx=15)

        # ═══════════════════════════════════════════════════════
        # البنود
        # ═══════════════════════════════════════════════════════
        items_frame = tk.LabelFrame(main, text="بنود الفاتورة",
                                     font=('Arial', 11, 'bold'),
                                     bg='#ecf0f1', fg='#34495e',
                                     padx=10, pady=10)
        items_frame.pack(fill='both', expand=True, pady=(0, 10))

        # ─── إدخال البند ───
        add_row = tk.Frame(items_frame, bg='#ecf0f1')
        add_row.pack(fill='x', pady=(0, 10))

        tk.Label(add_row, text="اسم الصنف:",
                 font=('Arial', 11, 'bold'), bg='#ecf0f1').pack(side='right', padx=3)

        self.entry_item_name = tk.Entry(add_row, width=18, font=('Arial', 11))
        self.entry_item_name.pack(side='right', padx=3)

        tk.Label(add_row, text="الكمية:",
                 font=('Arial', 11, 'bold'), bg='#ecf0f1').pack(side='right', padx=3)

        self.entry_item_qty = tk.Entry(add_row, width=8,
                                        font=('Arial', 11), justify='center')
        self.entry_item_qty.pack(side='right', padx=3)

        tk.Label(add_row, text="السعر:",
                 font=('Arial', 11, 'bold'), bg='#ecf0f1').pack(side='right', padx=3)

        self.entry_item_price = tk.Entry(add_row, width=10,
                                          font=('Arial', 11), justify='center')
        self.entry_item_price.pack(side='right', padx=3)

        tk.Button(add_row, text="➕ إضافة",
                  command=self.add_item,
                  bg='#27ae60', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=3).pack(side='right', padx=5)

        # ─── جدول البنود ───
        tree_frame = tk.Frame(items_frame, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True)

        columns = ('name', 'qty', 'price', 'total')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                  show='headings', height=6)
        headers = {'name': 'الصنف', 'qty': 'الكمية',
                   'price': 'السعر', 'total': 'الإجمالي'}
        for col, h in headers.items():
            self.tree.heading(col, text=h)
            self.tree.column(col, width=150, anchor='center')

        self.tree.tag_configure('odd_row', background='#f8f9fa')
        self.tree.tag_configure('even_row', background='#ffffff')

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        btn_items = tk.Frame(items_frame, bg='#ecf0f1')
        btn_items.pack(pady=5)

        tk.Button(btn_items, text="✏ تعديل البند",
                  command=self.edit_item,
                  bg='#f39c12', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=3).pack(side='right', padx=3)

        tk.Button(btn_items, text="🗑 حذف البند",
                  command=self.delete_item,
                  bg='#e74c3c', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=3).pack(side='right', padx=3)

        # ─── الإجمالي ───
        total_frame = tk.Frame(main, bg='#1a5276', height=50)
        total_frame.pack(fill='x')
        total_frame.pack_propagate(False)

        tk.Label(total_frame, text="الإجمالي:",
                 font=('Arial', 13, 'bold'),
                 bg='#1a5276', fg='#f1c40f').pack(side='right', padx=15, pady=10)

        self.lbl_total = tk.Label(total_frame, text="0.00 ج.م",
                                   font=('Arial', 16, 'bold'),
                                   bg='#1a5276', fg='white')
        self.lbl_total.pack(side='left', padx=15, pady=10)

        # ─── ملاحظات ───
        tk.Label(main, text="ملاحظات:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(anchor='e', pady=(5, 0))

        self.entry_notes = tk.Entry(main, font=('Arial', 11))
        self.entry_notes.pack(fill='x', pady=(0, 10))

        # ─── الأزرار ───
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        btn_text = "💾 حفظ التعديلات" if self.is_edit else "💾 حفظ الفاتورة"
        tk.Button(btn_frame, text=btn_text,
                  command=self.save,
                  bg='#27ae60', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=30, pady=10).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء",
                  command=self.window.destroy,
                  bg='#e74c3c', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=30, pady=10).pack(side='left', padx=5)

        if self.is_edit:
            self._load_invoice()

        self.entry_item_name.focus()

    # ═══════════════════════════════════════════════════════════
    def _on_payment_type_change(self, event=None):
        """تحديث حقل المبلغ المدفوع حسب النوع"""
        ptype = self.payment_type_var.get()

        if ptype == 'كاش':
            # ملء الإجمالي تلقائياً
            try:
                total = sum(it['total'] for it in self.items)
                self.entry_paid_amount.delete(0, tk.END)
                self.entry_paid_amount.insert(0, f"{total:.2f}")
            except:
                pass
            self.entry_paid_amount.config(state='disabled')

        elif ptype == 'أجل':
            self.entry_paid_amount.config(state='normal')
            self.entry_paid_amount.delete(0, tk.END)
            self.entry_paid_amount.insert(0, "0")
            self.entry_paid_amount.config(state='disabled')

        else:  # جزئي
            self.entry_paid_amount.config(state='normal')
            self.entry_paid_amount.focus()

        self._update_remaining()

    def _update_remaining(self):
        """تحديث المتبقي"""
        try:
            total = sum(it['total'] for it in self.items)
            paid = float(self.entry_paid_amount.get() or 0)
            remaining = total - paid
            self.lbl_remaining.config(
                text=f"المتبقي: {remaining:,.2f} ج.م",
                fg='#c0392b' if remaining > 0 else '#27ae60')
        except:
            self.lbl_remaining.config(text="")

    # ═══════════════════════════════════════════════════════════
    def _load_invoice(self):
        invoice = get_trader_invoice(self.invoice_id)
        if not invoice:
            messagebox.showerror("خطأ", "الفاتورة غير موجودة")
            self.window.destroy()
            return

        try:
            if invoice.get('date'):
                self.entry_date.set_date(
                    datetime.strptime(invoice['date'], '%Y-%m-%d'))
        except:
            pass

        self.entry_invoice_no.insert(0, invoice.get('invoice_no', '') or '')
        self.trader_var.set(invoice.get('trader_name', ''))
        self.place_var.set(invoice.get('place', 'bazar') or 'bazar')

        # نوع الدفع
        ptype = invoice.get('payment_type', 'أجل') or 'أجل'
        self.payment_type_var.set(ptype)

        # المبلغ المدفوع
        paid = invoice.get('paid_amount', 0) or 0

        # البنود
        items = get_invoice_items(self.invoice_id)
        for it in items:
            self.items.append({
                'id': it['id'],
                'name': it['item_name'],
                'qty': it['quantity'],
                'price': it['price'],
                'total': it['total'],
            })
        self._refresh_tree()

        # المبلغ المدفوع بعد تحميل البنود
        self.entry_paid_amount.config(state='normal')
        self.entry_paid_amount.delete(0, tk.END)
        self.entry_paid_amount.insert(0, f"{paid:.2f}")

        self._on_payment_type_change()
        self._update_remaining()

    # ═══════════════════════════════════════════════════════════
    def add_item(self):
        name = self.entry_item_name.get().strip()
        try:
            qty = float(self.entry_item_qty.get() or 0)
            price = float(self.entry_item_price.get() or 0)
        except ValueError:
            messagebox.showerror("خطأ", "الكمية والسعر يجب أن يكونا أرقاماً")
            return

        if not name:
            messagebox.showwarning("تنبيه", "اسم الصنف مطلوب")
            return

        if qty <= 0 or price <= 0:
            messagebox.showwarning("تنبيه", "الكمية والسعر يجب أن يكونا أكبر من صفر")
            return

        self.items.append({
            'name': name, 'qty': qty, 'price': price, 'total': qty * price,
        })
        self._refresh_tree()
        self._clear_item_fields()
        self._on_payment_type_change()

    def edit_item(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("تنبيه", "اختر بنداً من الجدول أولاً")
            return

        idx = self.tree.index(sel[0])

        name = self.entry_item_name.get().strip()
        try:
            qty = float(self.entry_item_qty.get() or 0)
            price = float(self.entry_item_price.get() or 0)
        except ValueError:
            messagebox.showerror("خطأ", "الكمية والسعر يجب أن يكونا أرقاماً")
            return

        if not name:
            messagebox.showwarning("تنبيه", "اسم الصنف مطلوب")
            return

        self.items[idx]['name'] = name
        self.items[idx]['qty'] = qty
        self.items[idx]['price'] = price
        self.items[idx]['total'] = qty * price

        self._refresh_tree()
        self._clear_item_fields()
        self._on_payment_type_change()

    def delete_item(self):
        sel = self.tree.selection()
        if not sel:
            return
        if not messagebox.askyesno("تأكيد", "هل تريد حذف هذا البند؟"):
            return
        idx = self.tree.index(sel[0])
        del self.items[idx]
        self._refresh_tree()
        self._clear_item_fields()
        self._on_payment_type_change()

    def _clear_item_fields(self):
        self.entry_item_name.delete(0, tk.END)
        self.entry_item_qty.delete(0, tk.END)
        self.entry_item_price.delete(0, tk.END)
        self.entry_item_name.focus()

    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        total = 0
        for i, it in enumerate(self.items):
            tag = 'odd_row' if i % 2 == 0 else 'even_row'
            self.tree.insert('', 'end', tags=(tag,), values=(
                it['name'],
                f"{it['qty']:,.2f}",
                f"{it['price']:,.2f}",
                f"{it['total']:,.2f}",
            ))
            total += it['total']

        self.lbl_total.config(text=f"{total:,.2f} ج.م")

    def save(self):
        date = self.entry_date.get().strip()
        invoice_no = self.entry_invoice_no.get().strip()
        trader_name = self.trader_var.get().strip()
        place = self.place_var.get().strip()
        notes = self.entry_notes.get().strip()
        ptype = self.payment_type_var.get().strip()

        if not date:
            messagebox.showwarning("تنبيه", "التاريخ مطلوب")
            return
        if not trader_name:
            messagebox.showwarning("تنبيه", "التاجر مطلوب")
            return
        if not self.items:
            messagebox.showwarning("تنبيه", "الفاتورة فارغة")
            return

        trader_id = self.traders_map.get(trader_name)
        if not trader_id:
            messagebox.showerror("خطأ", "التاجر غير موجود")
            return

        total = sum(it['total'] for it in self.items)

        # ─── حساب المبلغ المدفوع ───
        if ptype == 'كاش':
            paid_amount = total
        elif ptype == 'أجل':
            paid_amount = 0
        else:  # جزئي
            try:
                paid_amount = float(self.entry_paid_amount.get() or 0)
            except ValueError:
                messagebox.showerror("خطأ", "المبلغ المدفوع غير صحيح")
                return
            if paid_amount < 0 or paid_amount > total:
                messagebox.showerror("خطأ",
                                     f"المبلغ المدفوع يجب أن يكون بين 0 و {total:,.2f}")
                return

        # ─── الحفظ ───
        if self.is_edit:
            ok = update_trader_invoice(
                self.invoice_id, invoice_no, date, trader_id,
                total, notes, self.items, ptype, paid_amount, place)
            success_msg = f"✅ تم تعديل الفاتورة\nالمتبقي: {total - paid_amount:,.2f} ج.م"
            error_msg = "فشل التعديل"
        else:
            result = add_trader_invoice(
                invoice_no, date, trader_id, total, notes,
                self.items, ptype, paid_amount, place)
            ok = result is not None
            success_msg = (f"✅ تم حفظ الفاتورة\n"
                           f"الإجمالي: {total:,.2f}\n"
                           f"المدفوع: {paid_amount:,.2f}\n"
                           f"المتبقي: {total - paid_amount:,.2f}")
            error_msg = "فشل الحفظ"

        if ok:
            messagebox.showinfo("نجاح", success_msg)
            self.window.destroy()
            self.parent_page.load_all()

            # ✅ تحديث صفحة المصروفات
            if self.parent_page.controller:
                try:
                    paids_page = self.parent_page.controller.pages.get('paidsPage')
                    if paids_page and hasattr(paids_page, 'load_all_paids'):
                        paids_page.load_all_paids()
                except:
                    pass
        else:
            messagebox.showerror("خطأ", error_msg)

# ═══════════════════════════════════════════════════════════════════
# TRADER PAYMENT WINDOW (يدعم العملات الأجنبية)
# ═══════════════════════════════════════════════════════════════════
class TraderPaymentWindow:
    def __init__(self, parent_page, trader_id=None, payment_id=None):
        self.parent_page = parent_page
        self.trader_id = trader_id
        self.payment_id = payment_id
        self.is_edit = payment_id is not None

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = "تعديل دفعة" if self.is_edit else "دفعة للتاجر"
        self.window.title(title)
        self.window.geometry("700x620")

        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ═══ 1) العنوان (أعلى) ═══
        title_frame = tk.Frame(self.window, bg='#1a5276', height=60)
        title_frame.pack(fill='x', side='top')
        title_frame.pack_propagate(False)
        icon = "✏" if self.is_edit else "💸"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'),
                 bg='#1a5276', fg='#f1c40f').pack(pady=15)

        # ═══ 2) الأزرار (أسفل) — قبل المحتوى ═══
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(side='bottom', pady=12)

        btn_text = "💾 حفظ التعديلات" if self.is_edit else "💾 حفظ"
        tk.Button(btn_frame, text=btn_text,
                  command=self.save,
                  bg='#27ae60', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=30, pady=10).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء",
                  command=self.window.destroy,
                  bg='#e74c3c', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=30, pady=10).pack(side='left', padx=5)

        # ─── الإطار الرئيسي ───
        # إطار خارجي للمنتصف
        outer = tk.Frame(self.window, bg='#ecf0f1')
        outer.pack(fill='x', pady=2)

        # إطار المحتوى الداخلي (متمركز بـ pack)
        main = tk.Frame(outer, bg='#ecf0f1', padx=5, pady=15,
                        relief='solid', bd=1)
        main.pack()  # ← بدون fill يجعل الإطار بحجم محتواه

        # ─── التاريخ ───
        tk.Label(main, text="التاريخ: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=0, column=0, padx=5, pady=15, sticky='e')

        self.entry_date = DateEntry(main, width=20,
                                     date_pattern='yyyy-mm-dd',
                                     font=('Arial', 11), justify='center')
        self.entry_date.grid(row=0, column=1, padx=5, pady=10, sticky='w')
        self.entry_date.set_date(datetime.now())

        # ─── التاجر ───
        tk.Label(main, text="التاجر: *",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=1, column=0, padx=5, pady=10, sticky='e')

        self.trader_var = tk.StringVar()
        self.trader_combo = ttk.Combobox(main, textvariable=self.trader_var,
                                          width=22, font=('Arial', 11),
                                          state='readonly')
        self.trader_combo.grid(row=1, column=1, padx=5, pady=10, sticky='w')

        traders = get_all_traders()
        self.traders_map = {t['name']: t['id'] for t in traders}
        self.trader_combo['values'] = list(self.traders_map.keys())

        if not self.is_edit and trader_id:
            for t in traders:
                if t['id'] == trader_id:
                    self.trader_var.set(t['name'])
                    break
        elif traders and not self.is_edit:
            self.trader_var.set(traders[0]['name'])

        self.trader_combo.bind('<<ComboboxSelected>>',
                                self._update_balance_label)

        # ─── الرصيد الحالي ───
        self.lbl_balance = tk.Label(main, text="",
                                     font=('Arial', 11, 'bold'),
                                     bg='#fff9e6', fg='#8B4513',
                                     relief='solid', bd=1,
                                     padx=10, pady=8)
        self.lbl_balance.grid(row=2, column=0, columnspan=2,
                               padx=5, pady=(0, 10), sticky='ew')

        # ═══════════════════════════════════════════════════════
        # قسم العملة
        # ═══════════════════════════════════════════════════════
        currency_frame = tk.LabelFrame(main, text="💰 العملة",
                                        font=('Arial', 11, 'bold'),
                                        bg='#fff9e6', fg='#8B4513',
                                        padx=10, pady=10)
        currency_frame.grid(row=3, column=0, columnspan=2,
                             padx=5, pady=(0, 10), sticky='ew')

        tk.Label(currency_frame, text="نوع العملة:",
                 font=('Arial', 11, 'bold'),
                 bg='#fff9e6', fg='#8B4513').grid(
                 row=0, column=0, padx=5, pady=5, sticky='e')

        self.currency_var = tk.StringVar(value="نقدي")
        currency_combo = ttk.Combobox(currency_frame,
                                       textvariable=self.currency_var,
                                       values=('نقدي', 'يورو', 'دولار'),
                                       state='readonly', width=15,
                                       font=('Arial', 11, 'bold'))
        currency_combo.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        currency_combo.bind('<<ComboboxSelected>>',
                             self._on_currency_change)

        # ─── المبلغ بالعملة الأجنبية ───
        tk.Label(currency_frame, text="المبلغ بالعملة:",
                 font=('Arial', 11, 'bold'),
                 bg='#fff9e6', fg='#8B4513').grid(
                 row=1, column=0, padx=5, pady=5, sticky='e')

        self.entry_foreign = tk.Entry(currency_frame, width=15,
                                       font=('Arial', 12),
                                       justify='center',
                                       bg='white')
        self.entry_foreign.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        self.entry_foreign.insert(0, "0")
        self.entry_foreign.bind('<KeyRelease>', self._update_conversion)

        # ─── سعر الصرف ───
        tk.Label(currency_frame, text="سعر الصرف:",
                 font=('Arial', 11, 'bold'),
                 bg='#fff9e6', fg='#8B4513').grid(
                 row=2, column=0, padx=5, pady=5, sticky='e')

        self.entry_rate = tk.Entry(currency_frame, width=15,
                                    font=('Arial', 12),
                                    justify='center',
                                    bg='white')
        self.entry_rate.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        self.entry_rate.insert(0, "0")
        self.entry_rate.bind('<KeyRelease>', self._update_conversion)

        # ─── المبلغ بالجنيه (يُحسب تلقائياً) ───
        self.lbl_amount_label = tk.Label(currency_frame, text="المبلغ (ج.م):",
                                         font=('Arial', 11, 'bold'),
                                         bg='#fff9e6', fg='#8B4513')
        self.lbl_amount_label.grid(row=1, column=0, padx=5, pady=5, sticky='e')

        self.lbl_converted = tk.Label(currency_frame, text="0.00",
                                       font=('Arial', 14, 'bold'),
                                       bg='#e8f8f5', fg='#145a32',
                                       relief='solid', bd=1,
                                       padx=10, pady=5)
        self.lbl_converted.grid(row=3, column=1, padx=5, pady=5, sticky='w')

        # ═══════════════════════════════════════════════════════
        # طريقة الدفع + المكان
        # ═══════════════════════════════════════════════════════
        tk.Label(main, text="طريقة الدفع:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=4, column=0, padx=5, pady=10, sticky='e')

        self.method_var = tk.StringVar(value="نقدي")
        ttk.Combobox(main, textvariable=self.method_var,
                     values=('نقدي', 'تحويل بنكي', 'شيك', 'فودافون كاش'),
                     state='readonly', width=20,
                     font=('Arial', 11)).grid(
                     row=4, column=1, padx=5, pady=10, sticky='w')

        tk.Label(main, text="المكان:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=5, column=0, padx=5, pady=10, sticky='e')

        self.place_var = tk.StringVar(value="bazar")
        ttk.Combobox(main, textvariable=self.place_var,
                     values=('bazar', 'beach', 'other'),
                     state='readonly', width=20,
                     font=('Arial', 11)).grid(
                     row=5, column=1, padx=5, pady=10, sticky='w')

        # ─── ملاحظات ───
        tk.Label(main, text="ملاحظات:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=6, column=0, padx=5, pady=10, sticky='e')

        self.entry_notes = tk.Entry(main, width=22, font=('Arial', 11))
        self.entry_notes.grid(row=6, column=1, padx=5, pady=10, sticky='w')

        # ─── ملاحظة توضيحية ───
        tk.Label(main,
                 text="💡 سيتم تسجيل الدفعة تلقائياً في صفحة المصروفات",
                 font=('Arial', 9, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').grid(
                 row=7, column=0, columnspan=2, pady=(5, 0))



        # ─── تحميل البيانات ───
        if self.is_edit:
            self._load_payment()
        else:
            self._update_balance_label()
            self._on_currency_change()

        self.entry_foreign.focus()

    # ═══════════════════════════════════════════════════════════
    def _on_currency_change(self, event=None):
        """تحديث حقول العملة حسب الاختيار"""
        curr = self.currency_var.get()

        # ✅ تحديث نص الليبل حسب العملة
        if curr == 'نقدي':
            self.lbl_amount_label.config(text="المبلغ (ج.م):")
            self.entry_foreign.config(state='normal', bg='white')  # ← مفتوحة!
            self.entry_rate.config(state='disabled', bg='#ecf0f1')
            self.lbl_converted.config(text="(نقدي - لا يحتاج تحويل)")
        elif curr == 'يورو':
            self.lbl_amount_label.config(text="المبلغ (€):")
            self.entry_foreign.config(state='normal', bg='white')
            self.entry_rate.config(state='normal', bg='white')
            self._update_conversion()
        elif curr == 'دولار':
            self.lbl_amount_label.config(text="المبلغ ($):")
            self.entry_foreign.config(state='normal', bg='white')
            self.entry_rate.config(state='normal', bg='white')
            self._update_conversion()

    def _update_conversion(self, event=None):
        """حساب المبلغ بالجنيه"""
        curr = self.currency_var.get()

        try:
            amount = float(self.entry_foreign.get() or 0)
        except ValueError:
            self.lbl_converted.config(text="خطأ")
            return

        if curr == 'نقدي':
            # عند "نقدي": المبلغ بالجنيه هو نفسه
            self.lbl_converted.config(text=f"{amount:,.2f} ج.م")
        else:
            try:
                rate = float(self.entry_rate.get() or 0)
                egp = amount * rate
                self.lbl_converted.config(text=f"{egp:,.2f} ج.م")
            except ValueError:
                self.lbl_converted.config(text="خطأ في الإدخال")
    def _load_payment(self):
        """تحميل بيانات الدفعة"""
        payment = get_trader_payment(self.payment_id)
        if not payment:
            messagebox.showerror("خطأ", "الدفعة غير موجودة")
            self.window.destroy()
            return

        # التاريخ
        try:
            if payment.get('date'):
                self.entry_date.set_date(
                    datetime.strptime(payment['date'], '%Y-%m-%d'))
        except:
            pass

        # التاجر
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('SELECT name FROM traders WHERE id=?',
                  (payment.get('trader_id'),))
        row = c.fetchone()
        conn.close()
        if row:
            self.trader_var.set(row[0])

        # العملة
        currency = payment.get('currency', 'نقدي') or 'نقدي'
        self.currency_var.set(currency)

        # المبلغ بالعملة
        foreign = payment.get('foreign_amount', 0) or 0
        self.entry_foreign.delete(0, tk.END)
        self.entry_foreign.insert(0, str(foreign))

        # سعر الصرف
        rate = payment.get('exchange_rate', 0) or 0
        self.entry_rate.delete(0, tk.END)
        self.entry_rate.insert(0, str(rate))

        # طريقة الدفع
        self.method_var.set(payment.get('payment_method', 'نقدي') or 'نقدي')

        # الملاحظات
        self.entry_notes.insert(0, payment.get('notes', '') or '')

        self._on_currency_change()
        self._update_balance_label()

    def _update_balance_label(self, event=None):
        name = self.trader_var.get()
        trader_id = self.traders_map.get(name)
        if trader_id:
            bal = get_trader_balance(trader_id)
            if self.is_edit:
                try:
                    original = float(self.entry_amount.get() or 0)
                    bal += original
                except:
                    pass
            self.lbl_balance.config(
                text=f"💰 الرصيد الحالي للتاجر: {bal:,.2f} ج.م")

    def save(self):
        date = self.entry_date.get().strip()
        trader_name = self.trader_var.get().strip()
        method = self.method_var.get().strip()
        place = self.place_var.get().strip()
        notes = self.entry_notes.get().strip()
        currency = self.currency_var.get().strip()

        if not date:
            messagebox.showwarning("تنبيه", "التاريخ مطلوب")
            return
        if not trader_name:
            messagebox.showwarning("تنبيه", "التاجر مطلوب")
            return

        # ─── حساب المبلغ بالجنيه ───
        # ─── حساب المبلغ بالجنيه ───
        if currency == 'نقدي':
            try:
                amount_egp = float(self.entry_foreign.get() or 0)
            except ValueError:
                messagebox.showerror("خطأ", "المبلغ غير صحيح")
                return
            foreign_amount = 0
            exchange_rate = 0
        else:
            try:
                foreign_amount = float(self.entry_foreign.get() or 0)
                exchange_rate = float(self.entry_rate.get() or 0)
            except ValueError:
                messagebox.showerror("خطأ", "المبلغ أو سعر الصرف غير صحيح")
                return
            if foreign_amount <= 0:
                messagebox.showwarning("تنبيه", "أدخل المبلغ بالعملة الأجنبية")
                return
            if exchange_rate <= 0:
                messagebox.showwarning("تنبيه", "أدخل سعر الصرف")
                return
            amount_egp = foreign_amount * exchange_rate

        if amount_egp <= 0:
            messagebox.showwarning("تنبيه", "المبلغ يجب أن يكون أكبر من صفر")
            return

        trader_id = self.traders_map.get(trader_name)
        if not trader_id:
            messagebox.showerror("خطأ", "التاجر غير موجود")
            return

        # ─── الحفظ ───
        if self.is_edit:
            ok = update_trader_payment(
                self.payment_id, date, trader_id, amount_egp,
                method, notes, place,
                currency, foreign_amount, exchange_rate)
            success_msg = f"✅ تم تعديل الدفعة\nالمبلغ: {amount_egp:,.2f} ج.م"
            error_msg = "فشل التعديل"
        else:
            ok, paid_id = add_trader_payment(
                date, trader_id, amount_egp, method, notes, place,
                currency, foreign_amount, exchange_rate)
            success_msg = (f"✅ تم تسجيل دفعة {amount_egp:,.2f} ج.م\n\n"
                           f"📌 ملاحظة: تم تسجيلها كمصروف")
            error_msg = "فشل الحفظ"

        if ok:
            messagebox.showinfo("نجاح", success_msg)
            self.window.destroy()
            self.parent_page.load_all()

            # تحديث صفحة المصروفات
            if self.parent_page.controller:
                try:
                    paids_page = self.parent_page.controller.pages.get('paidsPage')
                    if paids_page and hasattr(paids_page, 'load_all_paids'):
                        paids_page.load_all_paids()
                except:
                    pass
        else:
            messagebox.showerror("خطأ", error_msg)

# ═══════════════════════════════════════════════════════════════════
# TRADER INVOICES LIST WINDOW
# ═══════════════════════════════════════════════════════════════════
class TraderInvoicesListWindow:
    """نافذة عرض كل فواتير التاجر مع إمكانية التعديل"""

    def __init__(self, parent_page, trader_id):
        self.parent_page = parent_page
        self.trader_id = trader_id
        self.selected_invoice_id = None

        # جلب اسم التاجر
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('SELECT name FROM traders WHERE id=?', (trader_id,))
        row = c.fetchone()
        conn.close()
        trader_name = row[0] if row else "غير معروف"

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title(f"فواتير - {trader_name}")
        self.window.geometry("900x600")
        self.window.grab_set()
        self.window.configure(bg='#ecf0f1')

        # ─── العنوان ───
        title_frame = tk.Frame(self.window, bg='#1a5276', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        tk.Label(title_frame, text=f"📄 فواتير: {trader_name}",
                 font=('Arial', 16, 'bold'),
                 bg='#1a5276', fg='#f1c40f').pack(pady=15)

        # ─── الأزرار ───
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(fill='x', padx=15, pady=10)

        tk.Button(btn_frame, text="✏ تعديل الفاتورة",
                  command=self.edit_invoice,
                  bg='#f39c12', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8).pack(side='right', padx=5)

        tk.Button(btn_frame, text="🗑 حذف الفاتورة",
                  command=self.delete_invoice,
                  bg='#e74c3c', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8).pack(side='right', padx=5)

        tk.Button(btn_frame, text="🔄 تحديث",
                  command=self.load_invoices,
                  bg='#3498db', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8).pack(side='right', padx=5)

        # ─── الجدول ───
        tree_frame = tk.Frame(self.window, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(0, 15))

        columns = ('id', 'date', 'invoice_no', 'total', 'notes')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                  show='headings', height=15)

        headers = {
            'id': 'ID', 'date': 'التاريخ', 'invoice_no': 'رقم الفاتورة',
            'total': 'الإجمالي', 'notes': 'ملاحظات'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=60, anchor='center')
        self.tree.column('date', width=120, anchor='center')
        self.tree.column('invoice_no', width=140, anchor='center')
        self.tree.column('total', width=160, anchor='center')
        self.tree.column('notes', width=380, anchor='e')

        self.tree.tag_configure('odd_row', background='#f8f9fa')
        self.tree.tag_configure('even_row', background='#ffffff')

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', lambda e: self.edit_invoice())

        self.load_invoices()

    def load_invoices(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        invoices = get_trader_invoices(self.trader_id)
        for i, inv in enumerate(invoices):
            tag = 'odd_row' if i % 2 == 0 else 'even_row'
            self.tree.insert('', 'end', iid=str(inv['id']), tags=(tag,), values=(
                inv['id'],
                inv.get('date', ''),
                inv.get('invoice_no', '') or '-',
                f"{inv.get('total', 0):,.2f}",
                inv.get('notes', '') or '',
            ))

    def on_select(self, event):
        sel = self.tree.selection()
        self.selected_invoice_id = int(sel[0]) if sel else None

    def edit_invoice(self):
        if self.selected_invoice_id is None:
            messagebox.showwarning("تنبيه", "اختر فاتورة أولاً")
            return
        TraderInvoiceWindow(self.parent_page,
                             invoice_id=self.selected_invoice_id)
        # ستُحدَّث القائمة بعد إغلاق النافذة
        self.window.after(500, self.load_invoices)

    def delete_invoice(self):
        if self.selected_invoice_id is None:
            messagebox.showwarning("تنبيه", "اختر فاتورة أولاً")
            return

        if not messagebox.askyesno("تأكيد",
                                    "هل أنت متأكد من حذف هذه الفاتورة؟"):
            return

        if delete_trader_invoice(self.selected_invoice_id):
            messagebox.showinfo("نجاح", "تم الحذف")
            self.selected_invoice_id = None
            self.load_invoices()


# ═══════════════════════════════════════════════════════════════════
# TRADER STATEMENT WINDOW
# ═══════════════════════════════════════════════════════════════════
class TraderStatementWindow:
    def __init__(self, parent_page, trader_id):
        self.parent_page = parent_page
        self.trader_id = trader_id

        # جلب اسم التاجر
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('SELECT name FROM traders WHERE id=?', (trader_id,))
        row = c.fetchone()
        conn.close()
        trader_name = row[0] if row else "غير معروف"

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title(f"كشف حساب - {trader_name}")
        self.window.geometry("1000x680")
        self.window.grab_set()
        self.window.configure(bg='#ecf0f1')

        # ─── العنوان ───
        title_frame = tk.Frame(self.window, bg='#1a5276', height=70)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        tk.Label(title_frame,
                 text=f"📄 كشف حساب: {trader_name}",
                 font=('Arial', 16, 'bold'),
                 bg='#1a5276', fg='#f1c40f').pack(pady=20)

        # ─── الجدول ───
        tree_frame = tk.Frame(self.window, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, padx=15, pady=15)

        columns = ('date', 'type', 'description', 'debit', 'credit', 'balance')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                  show='headings', height=18)

        headers = {
            'date': 'التاريخ', 'type': 'النوع',
            'description': 'البيان', 'debit': 'مدين (دفع)',
            'credit': 'دائن (فاتورة)', 'balance': 'الرصيد'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('date', width=110, anchor='center')
        self.tree.column('type', width=100, anchor='center')
        self.tree.column('description', width=350, anchor='e')
        self.tree.column('debit', width=120, anchor='center')
        self.tree.column('credit', width=120, anchor='center')
        self.tree.column('balance', width=130, anchor='center')

        self.tree.tag_configure('invoice', background='#fff9e6')
        self.tree.tag_configure('payment', background='#e8f8f5')

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        # ─── الملخص ───
        summary = tk.Frame(self.window, bg='#34495e', height=70)
        summary.pack(fill='x', padx=15, pady=(0, 15))
        summary.pack_propagate(False)

        self.lbl_invoices = tk.Label(summary, text="إجمالي الفواتير: 0.00",
                                      font=('Arial', 12, 'bold'),
                                      bg='#34495e', fg='#f1c40f')
        self.lbl_invoices.pack(side='right', padx=20, pady=20)

        self.lbl_payments = tk.Label(summary, text="إجمالي المدفوع: 0.00",
                                      font=('Arial', 12, 'bold'),
                                      bg='#34495e', fg='#2ecc71')
        self.lbl_payments.pack(side='right', padx=20, pady=20)

        self.lbl_balance = tk.Label(summary, text="الرصيد الحالي: 0.00",
                                     font=('Arial', 14, 'bold'),
                                     bg='#34495e', fg='#e74c3c')
        self.lbl_balance.pack(side='left', padx=20, pady=20)

        # ─── الأزرار ───
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="💾 تصدير Excel",
                  command=self.export_to_excel,
                  bg='#27ae60', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=20, pady=8).pack(side='left', padx=5)

        tk.Button(btn_frame, text="🖼️ تصدير صورة",
                  command=self.export_to_image,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=20, pady=8).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إغلاق",
                  command=self.window.destroy,
                  bg='#e74c3c', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=20, pady=8).pack(side='left', padx=5)

        self._load_statement()

    def _load_statement(self):
        movements = get_trader_statement(self.trader_id)

        total_invoices = 0
        total_payments = 0

        for m in movements:
            tag = 'invoice' if m['type'] == 'فاتورة' else 'payment'
            self.tree.insert('', 'end', tags=(tag,), values=(
                m['date'],
                m['type'],
                m['description'],
                f"{m['debit']:,.2f}" if m['debit'] else '',
                f"{m['credit']:,.2f}" if m['credit'] else '',
                f"{m['balance']:,.2f}",
            ))
            total_invoices += m['credit']
            total_payments += m['debit']

        balance = total_invoices - total_payments

        self.lbl_invoices.config(
            text=f"إجمالي الفواتير: {total_invoices:,.2f}")
        self.lbl_payments.config(
            text=f"إجمالي المدفوع: {total_payments:,.2f}")
        self.lbl_balance.config(
            text=f"الرصيد الحالي: {balance:,.2f} ج.م",
            fg='#e74c3c' if balance > 0 else '#2ecc71')

    # ═══════════════════════════════════════════════════════════
    # التصدير إلى Excel
    # ═══════════════════════════════════════════════════════════
    def export_to_excel(self):
        """تصدير كشف حساب التاجر إلى ملف Excel"""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import (Font, Alignment,
                                         PatternFill, Border, Side)
        except ImportError:
            messagebox.showerror("خطأ", "مكتبة openpyxl غير مثبتة")
            return

        from tkinter import filedialog

        # جلب اسم التاجر
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute("SELECT name, phone FROM traders WHERE id=?",
                  (self.trader_id,))
        row = c.fetchone()
        conn.close()
        if not row:
            messagebox.showerror("خطأ", "التاجر غير موجود")
            return
        trader_name = row[0]
        phone = row[1] or ''

        # اسم الملف الافتراضي
        safe_name = trader_name.replace(' ', '_').replace('/', '_')
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"كشف_حساب_{safe_name}.xlsx",
            title="حفظ كشف الحساب Excel"
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "كشف حساب"
            ws.sheet_view.rightToLeft = True

            # ═══ عرض الأعمدة ═══
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 45
            ws.column_dimensions['D'].width = 15
            ws.column_dimensions['E'].width = 15
            ws.column_dimensions['F'].width = 18

            # ═══ العنوان ═══
            ws.merge_cells('A1:F1')
            c1 = ws['A1']
            c1.value = f"📄 كشف حساب التاجر: {trader_name}"
            c1.font = Font(name='Arial', size=16, bold=True,
                           color='FFFFFF')
            c1.fill = PatternFill("solid", fgColor="1A5276")
            c1.alignment = Alignment(horizontal='center',
                                     vertical='center')
            ws.row_dimensions[1].height = 35

            # ═══ الفون ═══
            if phone:
                ws.merge_cells('A2:F2')
                c2 = ws['A2']
                c2.value = f"📞 {phone}"
                c2.font = Font(name='Arial', size=12, bold=True,
                               color='2C3E50')
                c2.fill = PatternFill("solid", fgColor="ECF0F1")
                c2.alignment = Alignment(horizontal='center',
                                         vertical='center')
                ws.row_dimensions[2].height = 22
                start_row = 4
            else:
                start_row = 3

            # ═══ الرؤوس ═══
            headers = ['التاريخ', 'النوع', 'البيان',
                       'مدين (دفع)', 'دائن (فاتورة)', 'الرصيد']
            for col_idx, h in enumerate(headers, 1):
                cell = ws.cell(row=start_row, column=col_idx, value=h)
                cell.font = Font(name='Arial', size=12, bold=True,
                                 color='FFFFFF')
                cell.fill = PatternFill("solid", fgColor="34495E")
                cell.alignment = Alignment(horizontal='center',
                                           vertical='center')
                cell.border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin'))
            ws.row_dimensions[start_row].height = 25

            # ═══ البيانات ═══
            row_num = start_row + 1
            total_invoices = 0
            total_payments = 0

            for item in self.tree.get_children():
                values = self.tree.item(item)['values']
                is_invoice = 'فاتورة' in str(values[1])
                bg = "FFF9E6" if is_invoice else "E8F8F5"

                for col_idx, val in enumerate(values, 1):
                    cell = ws.cell(row=row_num, column=col_idx, value=val)
                    cell.font = Font(name='Arial', size=11)
                    cell.fill = PatternFill("solid", fgColor=bg)
                    cell.alignment = Alignment(
                        horizontal='right' if col_idx == 3 else 'center',
                        vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin'))
                row_num += 1

            # ═══ الملخص ═══
            row_num += 1

            # إجمالي الفواتير
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=5)
            cell = ws.cell(row=row_num, column=1,
                           value="إجمالي الفواتير (دائن)")
            cell.font = Font(name='Arial', size=12, bold=True,
                             color='8B4513')
            cell.fill = PatternFill("solid", fgColor="FFF3CD")
            cell.alignment = Alignment(horizontal='right',
                                       vertical='center')

            inv_text = self.lbl_invoices.cget('text')
            inv_val = inv_text.split(':')[-1].strip() if ':' in inv_text else '0'
            cell = ws.cell(row=row_num, column=6, value=inv_val)
            cell.font = Font(name='Arial', size=12, bold=True,
                             color='C0392B')
            cell.fill = PatternFill("solid", fgColor="FFF3CD")
            cell.alignment = Alignment(horizontal='center',
                                       vertical='center')
            row_num += 1

            # إجمالي المدفوع
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=5)
            cell = ws.cell(row=row_num, column=1,
                           value="إجمالي المدفوع (مدين)")
            cell.font = Font(name='Arial', size=12, bold=True,
                             color='145A32')
            cell.fill = PatternFill("solid", fgColor="D5F4E6")
            cell.alignment = Alignment(horizontal='right',
                                       vertical='center')

            pay_text = self.lbl_payments.cget('text')
            pay_val = pay_text.split(':')[-1].strip() if ':' in pay_text else '0'
            cell = ws.cell(row=row_num, column=6, value=pay_val)
            cell.font = Font(name='Arial', size=12, bold=True,
                             color='145A32')
            cell.fill = PatternFill("solid", fgColor="D5F4E6")
            cell.alignment = Alignment(horizontal='center',
                                       vertical='center')
            row_num += 1

            # الرصيد الحالي
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=5)
            cell = ws.cell(row=row_num, column=1,
                           value="⭐ الرصيد الحالي")
            cell.font = Font(name='Arial', size=13, bold=True,
                             color='FFFFFF')
            cell.fill = PatternFill("solid", fgColor="1A5276")
            cell.alignment = Alignment(horizontal='right',
                                       vertical='center')

            bal_text = self.lbl_balance.cget('text')
            bal_val = bal_text.split(':')[-1].strip() if ':' in bal_text else '0'
            cell = ws.cell(row=row_num, column=6, value=bal_val)
            cell.font = Font(name='Arial', size=13, bold=True,
                             color='FFFFFF')
            cell.fill = PatternFill("solid", fgColor="1A5276")
            cell.alignment = Alignment(horizontal='center',
                                       vertical='center')
            ws.row_dimensions[row_num].height = 30

            # ═══ التذييل ═══
            row_num += 2
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=6)
            cell = ws.cell(row=row_num, column=1,
                           value="شركة الزهراء للبرمجيات - "
                                 "المحاسب / حسام الدين سيد ربيع - "
                                 "01003796984")
            cell.font = Font(name='Arial', size=10, italic=True,
                             color='7F8C8D')
            cell.alignment = Alignment(horizontal='center',
                                       vertical='center')

            # ═══ الحفظ ═══
            wb.save(file_path)
            messagebox.showinfo("نجاح", f"✅ تم الحفظ:\n{file_path}")

            if messagebox.askyesno("فتح الملف",
                                   "هل تريد فتح الملف الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"فشل التصدير:\n{e}")

    # ═══════════════════════════════════════════════════════════
    # التصدير إلى صورة
    # ═══════════════════════════════════════════════════════════
    def export_to_image(self):
        """تصدير كشف حساب التاجر إلى صورة PNG"""
        try:
            from PIL import Image, ImageDraw, ImageFont
            import arabic_reshaper
            from bidi.algorithm import get_display
        except ImportError:
            messagebox.showerror(
                "خطأ",
                "مكتبة مفقودة. نفذ:\n"
                "pip install Pillow arabic-reshaper python-bidi")
            return

        from tkinter import filedialog

        # جلب اسم التاجر
        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute("SELECT name, phone FROM traders WHERE id=?",
                  (self.trader_id,))
        row = c.fetchone()
        conn.close()
        if not row:
            messagebox.showerror("خطأ", "التاجر غير موجود")
            return
        trader_name = row[0]
        phone = row[1] or ''

        safe_name = trader_name.replace(' ', '_').replace('/', '_')
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=f"كشف_حساب_{safe_name}.png",
            title="حفظ كشف الحساب صورة"
        )
        if not file_path:
            return

        try:
            # ─── الخط العربي ───
            font_path = None
            for fp in ["C:/Windows/Fonts/arial.ttf",
                       "C:/Windows/Fonts/tahoma.ttf",
                       "C:/Windows/Fonts/segoeui.ttf"]:
                if os.path.exists(fp):
                    font_path = fp
                    break

            if not font_path:
                messagebox.showerror("خطأ", "لا يوجد خط عربي متاح")
                return

            def ar(text):
                if text is None:
                    return ''
                text = str(text)
                reshaped = arabic_reshaper.reshape(text)
                return get_display(reshaped)

            # ─── الخطوط ───
            font_title = ImageFont.truetype(font_path, 32)
            font_subtitle = ImageFont.truetype(font_path, 22)
            font_header = ImageFont.truetype(font_path, 20)
            font_body = ImageFont.truetype(font_path, 18)
            font_summary = ImageFont.truetype(font_path, 22)
            font_footer = ImageFont.truetype(font_path, 14)

            # ─── جلب البيانات من الجدول ───
            rows_data = []
            for item in self.tree.get_children():
                values = self.tree.item(item)['values']
                rows_data.append(values)

            # ─── الأبعاد ───
            # ═══ الأبعاد ═══
            W = 1200
            margin = 40
            row_h = 38
            header_h = 45
            top_h = 130
            summary_h = 3 * 45 + 40  # 3 صفوف ملخص
            footer_h = 60

            # ✅ الشعار
            logo_size = 90
            logo_img = None
            if COMPANY_LOGO_PATH and os.path.exists(COMPANY_LOGO_PATH):
                try:
                    logo_img = Image.open(COMPANY_LOGO_PATH)
                    if logo_img.mode != 'RGBA':
                        logo_img = logo_img.convert('RGBA')
                    logo_img.thumbnail((logo_size, logo_size),
                                       Image.Resampling.LANCZOS)
                except Exception as e:
                    print(f"⚠️ فشل تحميل الشعار: {e}")
                    logo_img = None

            table_h = header_h + (len(rows_data) * row_h) + 20
            H = top_h + 40 + table_h + summary_h + footer_h + 40

            # ─── الصورة ───
            img = Image.new('RGB', (W, H), 'white')
            draw = ImageDraw.Draw(img)

            # ─── الألوان ───
            COLOR_HEADER_BG = (26, 82, 118)
            COLOR_HEADER_FG = (241, 196, 15)
            COLOR_TABLE_HEADER = (52, 73, 94)
            COLOR_ROW_INVOICE = (255, 249, 230)
            COLOR_ROW_PAYMENT = (232, 248, 245)
            COLOR_BORDER = (189, 195, 199)
            COLOR_FOOTER = (127, 140, 141)

            y = 0

            # ═══ الهيدر ═══
            draw.rectangle([(0, 0), (W, top_h)], fill=COLOR_HEADER_BG)

            # ✅ الشعار على اليسار (قبل العنوان بقليل)
            if logo_img:
                logo_x = 30
                logo_y = (top_h - logo_img.height) // 2
                # لصق مع احترام الشفافية
                if logo_img.mode == 'RGBA':
                    img.paste(logo_img, (logo_x, logo_y), logo_img)
                else:
                    img.paste(logo_img, (logo_x, logo_y))

            # عنوان (يُزاح لليمين إذا كان الشعار موجوداً)
            title_text = ar(f"📄 كشف حساب التاجر: {trader_name}")
            bbox = draw.textbbox((0, 0), title_text, font=font_title)
            tw = bbox[2] - bbox[0]

            if logo_img:
                # إزاحة العنوان لتجنب التعارض مع الشعار
                center_x = (W + logo_size + 30) / 2
            else:
                center_x = W / 2

            draw.text((center_x - tw / 2, 25), title_text,
                      fill=COLOR_HEADER_FG, font=font_title)

            # فون
            if phone:
                phone_text = ar(f"📞 {phone}")
                bbox = draw.textbbox((0, 0), phone_text, font=font_subtitle)
                pw = bbox[2] - bbox[0]
                draw.text((center_x - pw / 2, 75), phone_text,
                          fill=(255, 255, 255), font=font_subtitle)

            # ═══ الرؤوس ═══
            y = top_h + 20
            headers = ['التاريخ', 'النوع', 'البيان',
                       'مدين', 'دائن', 'الرصيد']
            # أعمدة من اليمين لليسار
            num_cols = len(headers)
            avail_w = W - 2 * margin
            col_w = avail_w // num_cols

            draw.rectangle([(margin, y), (W - margin, y + header_h)],
                           fill=COLOR_TABLE_HEADER)

            for i, h in enumerate(headers):
                x_right = W - margin - (i * col_w)
                x_center = x_right - col_w // 2
                text = ar(h)
                bbox = draw.textbbox((0, 0), text, font=font_header)
                tw = bbox[2] - bbox[0]
                draw.text((x_center - tw / 2, y + 12), text,
                          fill=(255, 255, 255), font=font_header)

            y += header_h

            # ═══ الصفوف ═══
            for row_values in rows_data:
                is_invoice = 'فاتورة' in str(row_values[1])
                bg = COLOR_ROW_INVOICE if is_invoice else COLOR_ROW_PAYMENT

                draw.rectangle([(margin, y), (W - margin, y + row_h)],
                               fill=bg, outline=COLOR_BORDER)

                for i, val in enumerate(row_values):
                    x_right = W - margin - (i * col_w)
                    x_center = x_right - col_w // 2
                    text = ar(val if val is not None else '')
                    bbox = draw.textbbox((0, 0), text, font=font_body)
                    tw = bbox[2] - bbox[0]
                    draw.text((x_center - tw / 2, y + 8), text,
                              fill=(44, 62, 80), font=font_body)

                y += row_h

            # ═══ الملخص ═══
            y += 30

            # إجمالي الفواتير
            draw.rectangle([(margin, y), (W - margin, y + 45)],
                           fill=(255, 243, 205),
                           outline=(139, 69, 19), width=2)

            inv_text = self.lbl_invoices.cget('text')
            inv_val = inv_text.split(':')[-1].strip() if ':' in inv_text else '0'
            label_text = ar("إجمالي الفواتير (دائن)")
            value_text = ar(inv_val)

            draw.text((W - margin - 20, y + 12), label_text,
                      fill=(139, 69, 19), font=font_summary, anchor='ra')
            draw.text((margin + 20, y + 12), value_text,
                      fill=(192, 57, 43), font=font_summary, anchor='la')
            y += 45

            # إجمالي المدفوع
            draw.rectangle([(margin, y), (W - margin, y + 45)],
                           fill=(213, 244, 230),
                           outline=(20, 90, 50), width=2)

            pay_text = self.lbl_payments.cget('text')
            pay_val = pay_text.split(':')[-1].strip() if ':' in pay_text else '0'
            label_text = ar("إجمالي المدفوع (مدين)")
            value_text = ar(pay_val)

            draw.text((W - margin - 20, y + 12), label_text,
                      fill=(20, 90, 50), font=font_summary, anchor='ra')
            draw.text((margin + 20, y + 12), value_text,
                      fill=(20, 90, 50), font=font_summary, anchor='la')
            y += 45

            # الرصيد الحالي
            draw.rectangle([(margin, y), (W - margin, y + 55)],
                           fill=(26, 82, 118),
                           outline=(26, 82, 118), width=2)

            bal_text = self.lbl_balance.cget('text')
            bal_val = bal_text.split(':')[-1].strip() if ':' in bal_text else '0'
            label_text = ar("⭐ الرصيد الحالي")
            value_text = ar(bal_val)

            draw.text((W - margin - 20, y + 15), label_text,
                      fill=(255, 255, 255), font=font_summary, anchor='ra')
            draw.text((margin + 20, y + 15), value_text,
                      fill=(255, 255, 255), font=font_summary, anchor='la')
            y += 55 + 20

            # ═══ التذييل ═══
            footer = ar("شركة الزهراء للبرمجيات - المحاسب / "
                        "حسام الدين سيد ربيع - 01003796984")
            bbox = draw.textbbox((0, 0), footer, font=font_footer)
            fw = bbox[2] - bbox[0]
            draw.text(((W - fw) / 2, y), footer,
                      fill=COLOR_FOOTER, font=font_footer)

            # ═══ الحفظ ═══
            img.save(file_path, 'PNG')
            messagebox.showinfo("نجاح", f"✅ تم الحفظ:\n{file_path}")

            if messagebox.askyesno("فتح الصورة",
                                   "هل تريد فتح الصورة الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"فشل التصدير:\n{e}")


# ═══════════════════════════════════════════════════════════════════
# TRADERS PAGE
# ═══════════════════════════════════════════════════════════════════
class TradersPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller
        self.selected_id = None

        # 1) البانر
        add_program_banner(self)

        # 2) شريط الحالة
        self.status_bar = tk.Label(self, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#2c3e50', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # 3) منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        # 4) العنوان
        tk.Label(content, text="🧑‍💼 إدارة حسابات التجار",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=10)

        # 5) الأزرار
        btn_frame = tk.LabelFrame(content, text="العمليات",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=15, pady=12)
        btn_frame.pack(fill='x', padx=15, pady=(0, 10))

        row1 = tk.Frame(btn_frame, bg='#ecf0f1')
        row1.pack(pady=3)

        tk.Button(row1, text="➕ إضافة تاجر",
                  command=self.add_trader,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=13).pack(side='left', padx=3)

        self.btn_edit = tk.Button(row1, text="✏ تعديل",
                                   command=self.edit_trader,
                                   bg='#f39c12', fg='white',
                                   font=('Arial', 11, 'bold'),
                                   padx=15, pady=8, width=12,
                                   state='disabled')
        self.btn_edit.pack(side='left', padx=3)

        self.btn_delete = tk.Button(row1, text="🗑 حذف",
                                     command=self.delete_trader,
                                     bg='#e74c3c', fg='white',
                                     font=('Arial', 11, 'bold'),
                                     padx=15, pady=8, width=12,
                                     state='disabled')
        self.btn_delete.pack(side='left', padx=3)

        self.btn_invoice = tk.Button(row1, text="📄 إضافة فاتورة",
                                      command=self.add_invoice,
                                      bg='#2980b9', fg='white',
                                      font=('Arial', 11, 'bold'),
                                      padx=15, pady=8, width=13,
                                      state='disabled')
        self.btn_invoice.pack(side='left', padx=3)

        self.btn_payment = tk.Button(row1, text="💸 دفعة للتاجر",
                                      command=self.add_payment,
                                      bg='#16a085', fg='white',
                                      font=('Arial', 11, 'bold'),
                                      padx=15, pady=8, width=13,
                                      state='disabled')
        self.btn_payment.pack(side='left', padx=3)

        self.btn_statement = tk.Button(row1, text="📊 كشف حساب",
                                        command=self.show_statement,
                                        bg='#8e44ad', fg='white',
                                        font=('Arial', 11, 'bold'),
                                        padx=15, pady=8, width=13,
                                        state='disabled')
        self.btn_statement.pack(side='left', padx=3)
        self.btn_invoices_list = tk.Button(row1, text="📋 عرض الفواتير",
                                           command=self.show_invoices_list,
                                           bg='#c0392b', fg='white',
                                           font=('Arial', 11, 'bold'),
                                           padx=15, pady=8, width=13,
                                           state='disabled')
        self.btn_invoices_list.pack(side='left', padx=3)
        self.btn_payments_list = tk.Button(row1, text="💳 عرض الدفعات",
                                           command=self.show_payments_list,
                                           bg='#16a085', fg='white',
                                           font=('Arial', 11, 'bold'),
                                           padx=15, pady=8, width=13,
                                           state='disabled')
        self.btn_payments_list.pack(side='left', padx=3)

        # 6) الجدول
        tree_frame = tk.LabelFrame(content, text="قائمة التجار",
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 15))

        columns = ('id', 'name', 'phone', 'address', 'balance')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                  show='headings', height=15)

        headers = {
            'id': 'ID', 'name': 'اسم التاجر', 'phone': 'الهاتف',
            'address': 'العنوان', 'balance': 'الرصيد المستحق'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=60, anchor='center')
        self.tree.column('name', width=250, anchor='e')
        self.tree.column('phone', width=150, anchor='center')
        self.tree.column('address', width=250, anchor='e')
        self.tree.column('balance', width=160, anchor='center')

        self.tree.tag_configure('owe',
                                 background='#fdebd0',
                                 foreground='#7d6608',
                                 font=('Arial', 11, 'bold'))
        self.tree.tag_configure('paid',
                                 background='#e8f8f5',
                                 foreground='#145a32',
                                 font=('Arial', 11, 'bold'))

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', lambda e: self.show_invoices_list())

        # ─── تحميل ───
        self.load_all()

    def show_payments_list(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر تاجراً أولاً")
            return
        TraderPaymentsListWindow(self, self.selected_id)

    def show_invoices_list(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر تاجراً أولاً")
            return
        TraderInvoicesListWindow(self, self.selected_id)

    def load_all(self):
        rows = get_all_traders()
        self._populate_tree(rows)

        # إجمالي المديونيات
        total_due = 0
        for t in rows:
            bal = get_trader_balance(t['id'])
            if bal > 0:
                total_due += bal

        self.status_bar.config(
            text=f"عدد التجار: {len(rows)} | "
                 f"إجمالي المستحقات: {total_due:,.2f} ج.م")

    def _populate_tree(self, rows):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for t in rows:
            bal = get_trader_balance(t['id'])
            tag = 'owe' if bal > 0.01 else 'paid'
            self.tree.insert('', 'end', iid=str(t['id']), tags=(tag,), values=(
                t['id'],
                t['name'],
                t.get('phone', '') or '',
                t.get('address', '') or '',
                f"{bal:,.2f}",
            ))

    def on_select(self, event):
        sel = self.tree.selection()
        if sel:
            try:
                self.selected_id = int(sel[0])
            except (ValueError, TypeError):
                self.selected_id = None
                return
            for btn in (self.btn_edit, self.btn_delete,
                        self.btn_invoice, self.btn_payment,
                        self.btn_statement, self.btn_invoices_list,
                        self.btn_payments_list):
                btn.config(state='normal')


        else:
            self.selected_id = None
            for btn in (self.btn_edit, self.btn_delete,
                        self.btn_invoice, self.btn_payment,
                        self.btn_statement, self.btn_invoices_list):
                btn.config(state='disabled')

    # ═══ CRUD ═══
    def add_trader(self):
        TraderFormWindow(self)

    def edit_trader(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر تاجراً أولاً")
            return
        TraderFormWindow(self, trader_id=self.selected_id)

    def delete_trader(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر تاجراً أولاً")
            return

        if not messagebox.askyesno("تأكيد",
                                    "⚠️ سيتم حذف كل فواتير ومدفوعات هذا التاجر.\n"
                                    "هل أنت متأكد؟"):
            return

        if delete_trader(self.selected_id):
            messagebox.showinfo("نجاح", "تم الحذف")
            self.selected_id = None
            self.load_all()

    def add_invoice(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر تاجراً أولاً")
            return
        TraderInvoiceWindow(self, trader_id=self.selected_id)

    def add_payment(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر تاجراً أولاً")
            return
        TraderPaymentWindow(self, trader_id=self.selected_id)

    def show_statement(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر تاجراً أولاً")
            return
        TraderStatementWindow(self, self.selected_id)
# ═══════════════════════════════════════════════════════════════════
# TRADER PAYMENTS LIST WINDOW
# ═══════════════════════════════════════════════════════════════════
class TraderPaymentsListWindow:
    """نافذة عرض كل دفعات التاجر مع إمكانية التعديل"""

    def __init__(self, parent_page, trader_id):
        self.parent_page = parent_page
        self.trader_id = trader_id
        self.selected_payment_id = None

        conn = sqlite3.connect('sunrise.db')
        c = conn.cursor()
        c.execute('SELECT name FROM traders WHERE id=?', (trader_id,))
        row = c.fetchone()
        conn.close()
        trader_name = row[0] if row else "غير معروف"

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title(f"دفعات - {trader_name}")
        self.window.geometry("900x600")
        self.window.grab_set()
        self.window.configure(bg='#ecf0f1')

        # ─── العنوان ───
        title_frame = tk.Frame(self.window, bg='#16a085', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        tk.Label(title_frame, text=f"💸 دفعات: {trader_name}",
                 font=('Arial', 16, 'bold'),
                 bg='#16a085', fg='white').pack(pady=15)

        # ─── الأزرار ───
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(fill='x', padx=15, pady=10)

        tk.Button(btn_frame, text="✏ تعديل الدفعة",
                  command=self.edit_payment,
                  bg='#f39c12', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8).pack(side='right', padx=5)

        tk.Button(btn_frame, text="🗑 حذف الدفعة",
                  command=self.delete_payment,
                  bg='#e74c3c', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8).pack(side='right', padx=5)

        tk.Button(btn_frame, text="🔄 تحديث",
                  command=self.load_payments,
                  bg='#3498db', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8).pack(side='right', padx=5)

        # ─── الجدول ───
        tree_frame = tk.Frame(self.window, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(0, 15))

        columns = ('id', 'date', 'currency', 'foreign_amount',
                   'rate', 'amount', 'method', 'notes')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                 show='headings', height=15)

        headers = {
            'id': 'ID',
            'date': 'التاريخ',
            'currency': 'العملة',
            'foreign_amount': 'المبلغ بالعملة',
            'rate': 'سعر الصرف',
            'amount': 'المبلغ (ج.م)',
            'method': 'طريقة الدفع',
            'notes': 'ملاحظات'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=50, anchor='center')
        self.tree.column('date', width=100, anchor='center')
        self.tree.column('currency', width=80, anchor='center')
        self.tree.column('foreign_amount', width=120, anchor='center')
        self.tree.column('rate', width=100, anchor='center')
        self.tree.column('amount', width=130, anchor='center')
        self.tree.column('method', width=120, anchor='center')
        self.tree.column('notes', width=200, anchor='e')

        self.tree.tag_configure('odd_row', background='#f8f9fa')
        self.tree.tag_configure('even_row', background='#ffffff')

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', lambda e: self.edit_payment())

        self.load_payments()

    def load_payments(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        payments = get_trader_payments(self.trader_id)
        for i, p in enumerate(payments):
            tag = 'odd_row' if i % 2 == 0 else 'even_row'
            curr = p.get('currency', 'نقدي') or 'نقدي'
            foreign = p.get('foreign_amount', 0) or 0
            rate = p.get('exchange_rate', 0) or 0

            # إظهار المبلغ الأجنبي فقط إذا كانت العملة أجنبية
            if curr == 'نقدي':
                foreign_str = ''
                rate_str = ''
            else:
                symbol = '€' if curr == 'يورو' else '$'
                foreign_str = f"{foreign:,.2f} {symbol}"
                rate_str = f"{rate:,.2f}"

            self.tree.insert('', 'end', iid=str(p['id']), tags=(tag,), values=(
                p['id'],
                p.get('date', ''),
                curr,
                foreign_str,
                rate_str,
                f"{p.get('amount', 0):,.2f}",
                p.get('payment_method', '') or '',
                p.get('notes', '') or '',
            ))

    def on_select(self, event):
        sel = self.tree.selection()
        self.selected_payment_id = int(sel[0]) if sel else None

    def edit_payment(self):
        if self.selected_payment_id is None:
            messagebox.showwarning("تنبيه", "اختر دفعة أولاً")
            return
        TraderPaymentWindow(self.parent_page,
                             payment_id=self.selected_payment_id)
        self.window.after(500, self.load_payments)

    def delete_payment(self):
        if self.selected_payment_id is None:
            messagebox.showwarning("تنبيه", "اختر دفعة أولاً")
            return

        if not messagebox.askyesno("تأكيد",
                                    "⚠️ سيتم حذف الدفعة والمصروف المرتبط بها.\n"
                                    "هل أنت متأكد؟"):
            return

        if delete_trader_payment(self.selected_payment_id):
            messagebox.showinfo("نجاح",
                                 "تم الحذف من الدفعات والمصروفات")
            self.selected_payment_id = None
            self.load_payments()

            # تحديث صفحة المصروفات
            if self.parent_page.controller:
                try:
                    paids_page = self.parent_page.controller.pages.get('paidsPage')
                    if paids_page and hasattr(paids_page, 'load_all_paids'):
                        paids_page.load_all_paids()
                except:
                    pass
# ═══════════════════════════════════════════════════════════════════
# دالة التصدير الشامل — كل التقارير في ورقة واحدة متتالية
# ═══════════════════════════════════════════════════════════════════
def export_comprehensive_excel(start_date, end_date, file_path,
                                 visa_discount=2.5):

    """
    تصدير تقرير شامل يحتوي على كل التقارير في ورقة واحدة متتالية.
    """
    try:
        from openpyxl import Workbook
        from openpyxl.styles import (Font, Alignment, PatternFill,
                                     Border, Side)
    except ImportError:
        messagebox.showerror("خطأ", "مكتبة openpyxl غير مثبتة")
        return False

    try:
        # ═══════════════════════════════════════════════════════
        # إعدادات مشتركة
        # ═══════════════════════════════════════════════════════
        COMPANY_INFO = [
            "شركة الزهراء للبرمجيات - المحاسب / حسام الدين سيد ربيع",
            "أنظمة إدارة متكاملة | جودة واحترافية",
            "📞 01003796984 | ✉️ hussein@zahraa-soft.com",
        ]

        COLOR_HEADER_BG = "2C3E50"
        COLOR_HEADER_FG = "FFFFFF"
        COLOR_SECTION_BG = "8B4513"
        COLOR_TOTAL_BG = "2874A6"
        COLOR_TOTAL_FG = "FFFFFF"
        COLOR_ALT_ROW = "F8F9FA"
        COLOR_FOOTER = "7F8C8D"

        def style_cell(cell, bg=None, fg="000000", bold=False,
                       size=11, align='center', border=True):
            cell.font = Font(name='Arial', size=size, bold=bold, color=fg)
            if bg:
                cell.fill = PatternFill("solid", fgColor=bg)
            cell.alignment = Alignment(horizontal=align, vertical='center',
                                        wrap_text=True)
            if border:
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

        def style_title(ws, row, text, num_cols, bg=COLOR_HEADER_BG,
                         fg=COLOR_HEADER_FG, size=14, height=30):
            ws.merge_cells(start_row=row, start_column=1,
                            end_row=row, end_column=num_cols)
            cell = ws.cell(row=row, column=1, value=text)
            cell.font = Font(name='Arial', size=size, bold=True, color=fg)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[row].height = height
            return row + 1

        def add_footer(ws, row, num_cols):
            row += 1
            for i, line in enumerate(COMPANY_INFO):
                ws.merge_cells(start_row=row, start_column=1,
                                end_row=row, end_column=num_cols)
                cell = ws.cell(row=row, column=1, value=line)
                cell.font = Font(name='Arial', size=10,
                                  bold=(i == 0),
                                  color=COLOR_FOOTER if i > 0 else "2C3E50")
                cell.alignment = Alignment(horizontal='center',
                                            vertical='center')
                row += 1
            return row

        def add_separator(ws, row, num_cols, height=10):
            """إضافة فاصل بين التقارير"""
            ws.row_dimensions[row].height = height
            return row + 1

        # ═══════════════════════════════════════════════════════
        # إنشاء الملف — ورقة واحدة
        # ═══════════════════════════════════════════════════════
        wb = Workbook()
        ws = wb.active
        ws.title = "التقرير الشامل"
        ws.sheet_view.rightToLeft = True

        # ─── عرض الأعمدة (مشترك لكل التقارير) ───
        col_widths = {
            'A': 6,   # #
            'B': 14,  # التاريخ
            'C': 38,  # البيان / التكلفة
            'D': 14,  # القيمة / يورو
            'E': 14,  # يورو / دولار
            'F': 16,  # دولار / جنيه
            'G': 14,  # جنيه / فيزا
            'H': 14,  # فيزا / القسم
            'I': 14,  # المكان
            'J': 16,  # مصدر الدفع
        }
        for col, w in col_widths.items():
            ws.column_dimensions[col].width = w

        row = 1

        # ═══════════════════════════════════════════════════════════
        # عنوان رئيسي: التقرير الشامل
        # ═══════════════════════════════════════════════════════════
        row = style_title(ws, row,
                           "📊 التقرير الشامل",
                           10, bg="1A5276", fg="F1C40F", size=18, height=40)
        row = style_title(ws, row,
                           f"من {start_date} إلى {end_date}",
                           10, bg="1A5276", fg="FFFFFF", size=13, height=25)
        row = add_separator(ws, row, 10, 15)

        # ═══════════════════════════════════════════════════════════
        # ═══════ التقرير 1: المبيعات اليومية - bazar ═══════
        # ═══════════════════════════════════════════════════════════
        row = style_title(ws, row,
                           "📊 تقرير المبيعات اليومية - المكان: bazar",
                           10, bg="8B4513", fg="FFFFFF", size=14, height=30)

        # رؤوس الأعمدة
        headers_sales = ['#', 'التاريخ', 'التكلفة',
                          'يورو (€)', 'دولار ($)', 'جنيه (ج.م)', 'فيزا',
                          '', '', '']
        for col_idx, h in enumerate(headers_sales, 1):
            cell = ws.cell(row=row, column=col_idx, value=h)
            if h:
                style_cell(cell, bg="34495E", fg="FFFFFF", bold=True, size=11)
        ws.row_dimensions[row].height = 25
        row += 1

        bazar_rows = get_sales_data(start_date, end_date, "bazar")
        bazar_totals = get_sales_totals(start_date, end_date, "bazar")

        for i, r in enumerate(bazar_rows):
            values = [i + 1, r.get('date', ''), r.get('cost', 0),
                      r.get('euro', 0), r.get('usd', 0),
                      r.get('egp', 0), r.get('visa', 0), '', '', '']
            for col_idx, val in enumerate(values, 1):
                cell = ws.cell(row=row, column=col_idx, value=val)
                bg = COLOR_ALT_ROW if i % 2 == 0 else "FFFFFF"
                style_cell(cell, bg=bg)
            row += 1

        # الإجمالي
        total_values = ['', 'الإجمالي',
                        bazar_totals.get('cost', 0),
                        bazar_totals.get('euro', 0),
                        bazar_totals.get('usd', 0),
                        bazar_totals.get('egp', 0),
                        bazar_totals.get('visa', 0), '', '', '']
        for col_idx, val in enumerate(total_values, 1):
            cell = ws.cell(row=row, column=col_idx, value=val)
            style_cell(cell, bg=COLOR_TOTAL_BG, fg=COLOR_TOTAL_FG,
                       bold=True, size=12)
        ws.row_dimensions[row].height = 28
        row += 1

        # فاصل
        row = add_separator(ws, row, 10, 20)

        # ═══════════════════════════════════════════════════════════
        # ═══════ التقرير 2: المبيعات اليومية - beach ═══════
        # ═══════════════════════════════════════════════════════════
        row = style_title(ws, row,
                           "📊 تقرير المبيعات اليومية - المكان: beach",
                           10, bg="8B4513", fg="FFFFFF", size=14, height=30)

        for col_idx, h in enumerate(headers_sales, 1):
            cell = ws.cell(row=row, column=col_idx, value=h)
            if h:
                style_cell(cell, bg="34495E", fg="FFFFFF", bold=True, size=11)
        ws.row_dimensions[row].height = 25
        row += 1

        beach_rows = get_sales_data(start_date, end_date, "beach")
        beach_totals = get_sales_totals(start_date, end_date, "beach")

        for i, r in enumerate(beach_rows):
            values = [i + 1, r.get('date', ''), r.get('cost', 0),
                      r.get('euro', 0), r.get('usd', 0),
                      r.get('egp', 0), r.get('visa', 0), '', '', '']
            for col_idx, val in enumerate(values, 1):
                cell = ws.cell(row=row, column=col_idx, value=val)
                bg = COLOR_ALT_ROW if i % 2 == 0 else "FFFFFF"
                style_cell(cell, bg=bg)
            row += 1

        total_values = ['', 'الإجمالي',
                        beach_totals.get('cost', 0),
                        beach_totals.get('euro', 0),
                        beach_totals.get('usd', 0),
                        beach_totals.get('egp', 0),
                        beach_totals.get('visa', 0), '', '', '']
        for col_idx, val in enumerate(total_values, 1):
            cell = ws.cell(row=row, column=col_idx, value=val)
            style_cell(cell, bg=COLOR_TOTAL_BG, fg=COLOR_TOTAL_FG,
                       bold=True, size=12)
        ws.row_dimensions[row].height = 28
        row += 1

        # فاصل
        row = add_separator(ws, row, 10, 20)

        # ═══════════════════════════════════════════════════════════
        # ═══════ التقرير 3: الحساب الختامي ═══════
        # ═══════════════════════════════════════════════════════════
        row = style_title(ws, row,
                           "📊 الحساب الختامي",
                           10, bg="8B4513", fg="FFFFFF", size=14, height=30)

        # ═══ رؤوس الحساب الختامي ═══
        # دمج A:B لخلية "البيان"
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=2)
        cell = ws.cell(row=row, column=1, value="البيان")
        style_cell(cell, bg="34495E", fg="FFFFFF", bold=True, size=11)

        # ✅ الرؤوس في الأعمدة C, D, E, F, G
        headers_fs_clean = ['يورو', 'دولار', 'جنيه', 'فيزا', 'الإجمالي']
        for i, h in enumerate(headers_fs_clean):
            col_idx = 3 + i
            cell = ws.cell(row=row, column=col_idx, value=h)
            style_cell(cell, bg="34495E", fg="FFFFFF", bold=True, size=11)

        ws.row_dimensions[row].height = 25
        row += 1

        # ─── حساب الحساب الختامي ───
        month = end_date[:7]
        next_month = get_next_month(month)

        carried = get_total_carried_balance(month)
        c_euro = carried.get('euro', 0.0)
        c_usd = carried.get('usd', 0.0)
        c_egp = carried.get('egp', 0.0)
        euro_rate = carried.get('euro_rate', 51.0)
        usd_rate = carried.get('usd_rate', 48.0)

        next_carried = get_total_carried_balance(next_month)
        n_euro = next_carried.get('euro', 0.0)
        n_usd = next_carried.get('usd', 0.0)
        n_egp = next_carried.get('egp', 0.0)

        sales_data = get_sales_by_place(start_date, end_date)
        bazar = sales_data.get('bazar', {'euro': 0, 'usd': 0, 'egp': 0, 'visa': 0})
        beach = sales_data.get('beach', {'euro': 0, 'usd': 0, 'egp': 0, 'visa': 0})
        other = sales_data.get('other', {'euro': 0, 'usd': 0, 'egp': 0, 'visa': 0})

        t_euro = bazar['euro'] + beach['euro'] + other['euro'] + c_euro
        t_usd = bazar['usd'] + beach['usd'] + other['usd'] + c_usd
        t_egp = bazar['egp'] + beach['egp'] + other['egp'] + c_egp
        t_visa = bazar['visa'] + beach['visa'] + other['visa']

        visa_factor = 1 - (visa_discount / 100)

        e_conv = t_euro * euro_rate
        u_conv = t_usd * usd_rate
        v_net = t_visa * visa_factor
        total_egp_conv = t_egp + e_conv + u_conv + v_net

        exp = get_expenses_by_department(start_date, end_date)
        p_cur = exp.get('purchases', {'egp': 0, 'euro': 0, 'usd': 0})
        e_cur = exp.get('expenses', {'egp': 0, 'euro': 0, 'usd': 0})
        o_cur = exp.get('other', {'egp': 0, 'euro': 0, 'usd': 0})

        p_total = p_cur['egp'] + p_cur['euro'] * euro_rate + p_cur['usd'] * usd_rate
        e_total = e_cur['egp'] + e_cur['euro'] * euro_rate + e_cur['usd'] * usd_rate

        net_sales = total_egp_conv - p_total
        sub_total = net_sales - e_total
        carried_total = n_egp + n_euro * euro_rate + n_usd * usd_rate
        final_net = sub_total - carried_total

        # صفوف الحساب الختامي
        fs_rows = [
            ('رصيد مرحل من الشهر السابق', c_euro, c_usd, c_egp, '', '', 'normal'),
            ('مبيعات البازار', bazar['euro'], bazar['usd'], bazar['egp'],
             bazar['visa'], '', 'normal'),
            ('مبيعات البيتش', beach['euro'], beach['usd'], beach['egp'],
             beach['visa'], '', 'normal'),
            ('إجمالي (بما في ذلك الرصيد)', t_euro, t_usd, t_egp, t_visa,
             '', 'total'),
            ('سعر الصرف / الخصم', euro_rate, usd_rate, '',
             f'خصم {visa_discount:.1f}%', '', 'normal'),

            ('المبيعات المحولة (جنيه)', e_conv, u_conv, t_egp, v_net,
             total_egp_conv, 'total'),
            ('المشتريات', p_cur['euro'], p_cur['usd'], p_cur['egp'], '',
             p_total, 'normal'),
            ('صافي المبيعات', '', '', '', '', net_sales, 'total'),
            ('المصاريف', e_cur['euro'], e_cur['usd'], e_cur['egp'], '',
             e_total, 'normal'),
            ('صافي (صافي المبيعات - المصاريف)', '', '', '', '', sub_total,
             'total'),
            ('رصيد مرحل للشهر التالي', n_euro, n_usd, n_egp, '', carried_total,
             'normal'),
            ('⭐ الصافي النهائي ⭐', '', '', '', '', final_net,
             'positive' if final_net >= 0 else 'negative'),
        ]

        for r in fs_rows:
            label, euro, usd, egp, visa, total, tag = r

            # ═══ خلية البيان (مدمجة A:B) ═══
            ws.merge_cells(start_row=row, start_column=1,
                           end_row=row, end_column=2)
            cell = ws.cell(row=row, column=1, value=label)

            # تحديد الألوان حسب tag
            if tag == 'total':
                bg, fg = COLOR_TOTAL_BG, COLOR_TOTAL_FG
                b, sz = True, 11
            elif tag == 'positive':
                bg, fg = "27AE60", "FFFFFF"
                b, sz = True, 12
            elif tag == 'negative':
                bg, fg = "E74C3C", "FFFFFF"
                b, sz = True, 12
            else:
                bg, fg = None, "000000"
                b, sz = False, 11

            # تنسيق خلية A فقط (الدمج سيطبّق التنسيق على B تلقائيًا)
            style_cell(cell, bg=bg, fg=fg, bold=b, size=sz, align='right')

            # باقي الأعمدة
            values = [round(euro, 2) if isinstance(euro, (int, float)) else euro,
                      round(usd, 2) if isinstance(usd, (int, float)) else usd,
                      round(egp, 2) if isinstance(egp, (int, float)) else egp,
                      round(visa, 2) if isinstance(visa, (int, float)) else visa,
                      round(total, 2) if isinstance(total, (int, float)) else total,
                      '', '', '']

            for col_offset, val in enumerate(values, 0):
                col_idx = 3 + col_offset
                cell = ws.cell(row=row, column=col_idx, value=val)
                if tag == 'total':
                    style_cell(cell, bg=COLOR_TOTAL_BG, fg=COLOR_TOTAL_FG,
                               bold=True, size=11)
                elif tag == 'positive':
                    style_cell(cell, bg="27AE60", fg="FFFFFF",
                               bold=True, size=12)
                elif tag == 'negative':
                    style_cell(cell, bg="E74C3C", fg="FFFFFF",
                               bold=True, size=12)
                else:
                    style_cell(cell)

            ws.row_dimensions[row].height = 25
            row += 1

        # فاصل
        row = add_separator(ws, row, 10, 20)

        # ═══════════════════════════════════════════════════════════
        # ═══════ التقرير 4: المصروفات ═══════
        # ═══════════════════════════════════════════════════════════
        row = style_title(ws, row,
                           "💸 تقرير المصروفات "
                           "(المكان: الكل | مصدر الدفع: الكل | القسم: الكل)",
                           10, bg="8B4513", fg="FFFFFF", size=14, height=30)

        # ═══ رؤوس المصروفات: 9 أعمدة (بدون مصدر الدفع والمرسل) ═══
        headers4 = ['#', 'التاريخ', 'البيان', 'القيمة',
                    'يورو', 'دولار', 'القسم', 'المكان', 'الصورة']
        for col_idx, h in enumerate(headers4, 1):
            cell = ws.cell(row=row, column=col_idx, value=h)
            style_cell(cell, bg="34495E", fg="FFFFFF", bold=True, size=11)
        ws.row_dimensions[row].height = 25
        row += 1

        paids_rows = search_paids(None, start_date, end_date,
                                  "الكل", "الكل", "الكل")

        for i, r in enumerate(paids_rows):
            # ✅ استخراج اسم الصورة من المسار الكامل
            img_path = r.get('image_path') or ''
            img_name = os.path.basename(img_path) if img_path else ''

            values = [i + 1, r.get('date', ''), r.get('description', ''),
                      r.get('value', 0), r.get('euro', 0),
                      r.get('usd', 0), r.get('department', ''),
                      r.get('place', ''), img_name]

            for col_idx, val in enumerate(values, 1):
                cell = ws.cell(row=row, column=col_idx, value=val)
                bg = COLOR_ALT_ROW if i % 2 == 0 else "FFFFFF"
                style_cell(cell, bg=bg,
                           align='right' if col_idx == 3 else 'center')
            row += 1

        # ═══ الإجمالي ═══
        total_val = sum(r.get('value', 0) or 0 for r in paids_rows)
        total_euro = sum(r.get('euro', 0) or 0 for r in paids_rows)
        total_usd = sum(r.get('usd', 0) or 0 for r in paids_rows)

        # دمج A:C لخلية "الإجمالي"
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=3)
        cell = ws.cell(row=row, column=1,
                       value=f"الإجمالي ({len(paids_rows)} قيد)")
        style_cell(cell, bg=COLOR_TOTAL_BG, fg=COLOR_TOTAL_FG,
                   bold=True, size=12, align='right')

        # القيمة
        cell = ws.cell(row=row, column=4, value=total_val)
        style_cell(cell, bg=COLOR_TOTAL_BG, fg=COLOR_TOTAL_FG,
                   bold=True, size=12)

        # يورو
        cell = ws.cell(row=row, column=5, value=total_euro)
        style_cell(cell, bg=COLOR_TOTAL_BG, fg=COLOR_TOTAL_FG,
                   bold=True, size=12)

        # دولار
        cell = ws.cell(row=row, column=6, value=total_usd)
        style_cell(cell, bg=COLOR_TOTAL_BG, fg=COLOR_TOTAL_FG,
                   bold=True, size=12)

        # ✅ 3 أعمدة فارغة (القسم، المكان، الصورة) - بدلاً من 4
        for col_idx in range(7, 10):
            cell = ws.cell(row=row, column=col_idx, value='')
            style_cell(cell, bg=COLOR_TOTAL_BG, fg=COLOR_TOTAL_FG,
                       bold=True, size=12)

        ws.row_dimensions[row].height = 28
        row += 1

        # ═══════════════════════════════════════════════════════════
        # التذييل النهائي
        # ═══════════════════════════════════════════════════════════
        row = add_separator(ws, row, 10, 20)
        row = style_title(ws, row,
                           "═══════════════════════════════════════",
                           10, bg="ECF0F1", fg="7F8C8D", size=11, height=15)
        add_footer(ws, row - 1, 10)

        # ═══════════════════════════════════════════════════════
        # الحفظ
        # ═══════════════════════════════════════════════════════
        wb.save(file_path)
        return True

    except Exception as e:
        import traceback
        traceback.print_exc()
        messagebox.showerror("خطأ", f"فشل التصدير: {e}")
        return False

# ═══════════════════════════════════════════════════════════════════
# COMPREHENSIVE REPORT PAGE (التقرير الشامل)
# ═══════════════════════════════════════════════════════════════════
class ComprehensiveReportPage(tk.Frame):
    """صفحة التقرير الشامل — تصدير عدة تقارير في ملف Excel واحد"""

    def __init__(self, parent, controller):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller

        # 1) البانر
        add_program_banner(self)

        # 2) شريط الحالة
        self.status_bar = tk.Label(self, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#2c3e50', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # 3) منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        # 4) العنوان
        tk.Label(content, text="📊 التقرير الشامل",
                 font=('Arial', 22, 'bold'),
                 bg='#ecf0f1', fg='#1a5276').pack(pady=(20, 10))

        tk.Label(content,
                 text="تصدير عدة تقارير في ملف Excel واحد",
                 font=('Arial', 12, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=(0, 20))

        # ═══════════════════════════════════════════════════════
        # فلتر الفترة
        # ═══════════════════════════════════════════════════════
        filter_frame = tk.LabelFrame(content, text="  📅  اختر المدة  ",
                                      font=('Arial', 13, 'bold'),
                                      bg='#ecf0f1', fg='#1a5276',
                                      padx=20, pady=20)
        filter_frame.pack(fill='x', padx=40, pady=(0, 20))

        row1 = tk.Frame(filter_frame, bg='#ecf0f1')
        row1.pack(fill='x', pady=(0, 10))

        tk.Label(row1, text="من تاريخ:",
                 font=('Arial', 13, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(side='right', padx=10)

        self.start_date_entry = DateEntry(row1, width=15,
                                           date_pattern='yyyy-mm-dd',
                                           font=('Arial', 12),
                                           justify='center')
        self.start_date_entry.pack(side='right', padx=10)
        today = datetime.today()
        self.start_date_entry.set_date(today.replace(day=1))

        tk.Label(row1, text="إلى تاريخ:",
                 font=('Arial', 13, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(side='right', padx=10)

        self.end_date_entry = DateEntry(row1, width=15,
                                         date_pattern='yyyy-mm-dd',
                                         font=('Arial', 12),
                                         justify='center')
        self.end_date_entry.pack(side='right', padx=10)
        self.end_date_entry.set_date(today)
        # ─── نسبة خصم الفيزا ───
        row_visa = tk.Frame(filter_frame, bg='#ecf0f1')
        row_visa.pack(fill='x', pady=(10, 0))

        tk.Label(row_visa, text="نسبة خصم الفيزا (%):",
                 font=('Arial', 13, 'bold'),
                 bg='#ecf0f1', fg='#8e44ad').pack(side='right', padx=10)

        self.visa_discount_entry = tk.Entry(row_visa, width=10,
                                            font=('Arial', 13, 'bold'),
                                            justify='center',
                                            bg='#f4ecf7')
        self.visa_discount_entry.pack(side='right', padx=10)
        self.visa_discount_entry.insert(0, "2.5")

        # ═══════════════════════════════════════════════════════
        # فترات سريعة
        # ═══════════════════════════════════════════════════════
        row2 = tk.Frame(filter_frame, bg='#ecf0f1')
        row2.pack(fill='x', pady=(0, 10))

        tk.Label(row2, text="فترات سريعة:",
                 font=('Arial', 12, 'bold'),
                 bg='#ecf0f1', fg='#34495e').pack(side='right', padx=10)

        tk.Button(row2, text="هذا الشهر",
                  command=self.set_this_month,
                  bg='#16a085', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=6).pack(side='right', padx=3)

        tk.Button(row2, text="الشهر السابق",
                  command=self.set_prev_month,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=6).pack(side='right', padx=3)

        tk.Button(row2, text="آخر 30 يوم",
                  command=self.set_last_30_days,
                  bg='#3498db', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=6).pack(side='right', padx=3)

        # ═══════════════════════════════════════════════════════
        # محتوى التقرير
        # ═══════════════════════════════════════════════════════
        report_info = tk.LabelFrame(content,
                                     text="  📋  محتوى التقرير الشامل  ",
                                     font=('Arial', 13, 'bold'),
                                     bg='#ecf0f1', fg='#8B4513',
                                     padx=20, pady=20)
        report_info.pack(fill='x', padx=40, pady=(0, 20))

        reports_list = [
            ("1️⃣", "تقرير المبيعات اليومية - المكان: bazar",
             "كل المبيعات اليومية لمكان bazar (تكلفة + يورو + دولار + جنيه + فيزا)"),
            ("2️⃣", "تقرير المبيعات اليومية - المكان: beach",
             "كل المبيعات اليومية لمكان beach"),
            ("3️⃣", "الحساب الختامي (Final Statement)",
             "المبيعات + المشتريات + المصاريف + الصافي النهائي"),
            ("4️⃣", "تقرير المصروفات (Expenses)",
             "كل المصروفات (المكان: الكل | مصدر الدفع: الكل | القسم: الكل)"),
        ]

        for num, name, desc in reports_list:
            row_r = tk.Frame(report_info, bg='#ecf0f1')
            row_r.pack(fill='x', pady=5)

            tk.Label(row_r, text=num,
                     font=('Arial', 16),
                     bg='#ecf0f1').pack(side='right', padx=10)

            tk.Label(row_r, text=name,
                     font=('Arial', 12, 'bold'),
                     bg='#ecf0f1', fg='#2c3e50',
                     anchor='e').pack(side='right', padx=5)

            tk.Label(row_r, text=desc,
                     font=('Arial', 10, 'italic'),
                     bg='#ecf0f1', fg='#7f8c8d',
                     anchor='e').pack(side='right', padx=5)

        # ═══════════════════════════════════════════════════════
        # زر التصدير
        # ═══════════════════════════════════════════════════════
        btn_frame = tk.Frame(content, bg='#ecf0f1')
        btn_frame.pack(pady=(10, 30))

        tk.Button(btn_frame,
                  text="💾  تصدير التقرير الشامل إلى Excel",
                  command=self.export_report,
                  bg='#27ae60', fg='white',
                  font=('Arial', 15, 'bold'),
                  padx=40, pady=15).pack()

        tk.Label(content,
                 text="💡 سيتم إنشاء ملف Excel يحتوي على 4 أوراق (Sheets) + التذييل",
                 font=('Arial', 10, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=(0, 20))

    # ═══════════════════════════════════════════════════════════
    def set_this_month(self):
        today = datetime.today()
        self.start_date_entry.set_date(today.replace(day=1))
        self.end_date_entry.set_date(today)

    def set_prev_month(self):
        from datetime import timedelta
        today = datetime.today()
        first_current = today.replace(day=1)
        last_prev = first_current - timedelta(days=1)
        first_prev = last_prev.replace(day=1)
        self.start_date_entry.set_date(first_prev)
        self.end_date_entry.set_date(last_prev)

    def set_last_30_days(self):
        from datetime import timedelta
        today = datetime.today()
        self.start_date_entry.set_date(today - timedelta(days=29))
        self.end_date_entry.set_date(today)

    # ═══════════════════════════════════════════════════════════
    def export_report(self):
        start = self.start_date_entry.get().strip()
        end = self.end_date_entry.get().strip()

        if not start or not end:
            messagebox.showwarning("تنبيه", "اختر تاريخ البداية والنهاية")
            return

        try:
            d1 = datetime.strptime(start, '%Y-%m-%d')
            d2 = datetime.strptime(end, '%Y-%m-%d')
            if d2 < d1:
                messagebox.showerror("خطأ",
                                     "تاريخ النهاية يجب أن يكون بعد البداية")
                return
        except:
            messagebox.showerror("خطأ", "صيغة التاريخ غير صحيحة")
            return

        # ─── اختيار مسار الحفظ ───
        from tkinter import filedialog
        default_name = f"التقرير_الشامل_{start}_{end}.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_name,
            title="حفظ التقرير الشامل"
        )

        if not file_path:
            return

        # ─── التصدير ───
        self.status_bar.config(text="⏳ جارٍ التصدير...")
        self.update_idletasks()

        # قراءة نسبة الخصم
        try:
            visa_discount = float(self.visa_discount_entry.get().strip() or 2.5)
        except ValueError:
            visa_discount = 2.5

        success = export_comprehensive_excel(start, end, file_path, visa_discount)

        if success:
            self.status_bar.config(text=f"✅ تم التصدير: {file_path}")
            messagebox.showinfo("نجاح",
                                 f"✅ تم إنشاء التقرير الشامل\n\n"
                                 f"📄 الملف: {file_path}\n\n"
                                 f"📊 يحتوي على:\n"
                                 f"   1. مبيعات bazar\n"
                                 f"   2. مبيعات beach\n"
                                 f"   3. الحساب الختامي\n"
                                 f"   4. المصروفات")

            if messagebox.askyesno("فتح الملف", "هل تريد فتح الملف الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass
        else:
            self.status_bar.config(text="❌ فشل التصدير")

# ═══════════════════════════════════════════════════════════════════
# ABOUT PAGE (عن الشركة)
# ═══════════════════════════════════════════════════════════════════
class AboutPage(tk.Frame):
    """صفحة التعريف بالشركة والمسؤول والبرامج"""

    def __init__(self, parent, controller=None):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller

        # 1) الشريط العلوي
        add_program_banner(self)

        # 2) منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        # ═══════════════════════════════════════════════════════
        # العنوان الرئيسي
        # ═══════════════════════════════════════════════════════
        tk.Label(content, text="🏢 عن الشركة",
                 font=('Arial', 22, 'bold'),
                 bg='#ecf0f1', fg='#1a5276').pack(pady=(20, 10))

        # ═══════════════════════════════════════════════════════
        # الشعار الكبير
        # ═══════════════════════════════════════════════════════
        logo_frame = tk.Frame(content, bg='#ecf0f1')
        logo_frame.pack(pady=(10, 20))

        logo_displayed = False
        if os.path.exists(COMPANY_LOGO_PATH):
            try:
                logo_img = Image.open(COMPANY_LOGO_PATH)
                if logo_img.mode != 'RGBA':
                    logo_img = logo_img.convert('RGBA')
                # تكبير الشعار إلى 220×220
                logo_img.thumbnail((220, 220), Image.Resampling.LANCZOS)
                logo_photo = ImageTk.PhotoImage(logo_img)
                lbl = tk.Label(logo_frame, image=logo_photo,
                               bg='#ecf0f1')
                lbl.image = logo_photo  # لمنع حذفها من الذاكرة
                lbl.pack()
                logo_displayed = True
            except Exception as e:
                print(f"⚠️ [فشل تحميل الشعار]: {e}")

        if not logo_displayed:
            tk.Label(logo_frame, text="☀️",
                     font=('Arial', 80),
                     bg='#ecf0f1', fg='#1a5276').pack()

        # ─── اسم الشركة تحت الشعار ───
        tk.Label(content,
                 text="شركة الزهراء للبرمجيات",
                 font=('Arial', 24, 'bold'),
                 bg='#ecf0f1', fg='#1a5276').pack(pady=(0, 5))

        tk.Label(content,
                 text="Zahraa Software Company",
                 font=('Arial', 14, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=(0, 25))

        # ═══════════════════════════════════════════════════════
        # قسم: عن الشركة
        # ═══════════════════════════════════════════════════════
        company_section = tk.LabelFrame(
            content, text="  🏢  عن الشركة  ",
            font=('Arial', 14, 'bold'),
            bg='#ecf0f1', fg='#1a5276',
            padx=25, pady=20)
        company_section.pack(fill='x', padx=40, pady=(0, 20))

        company_info = (
            "شركة الزهراء للبرمجيات هي شركة رائدة في مجال تصميم وتطوير "
            "الأنظمة والبرامج المحاسبية والإدارية المتكاملة.\n\n"
            "نقدم حلولاً برمجية احترافية تلبي احتياجات الشركات والمؤسسات "
            "والمطاعم والمحلات التجارية، مع التركيز على:\n\n"
            "    ✅  جودة البرمجة والاحترافية في التنفيذ\n"
            "    ✅  سهولة الاستخدام وبساطة الواجهات\n"
            "    ✅  الدعم الفني المستمر بعد التسليم\n"
            "    ✅  التخصيص حسب احتياجات كل عميل\n"
        )
        tk.Label(company_section, text=company_info,
                 font=('Arial', 12),
                 bg='#ecf0f1', fg='#2c3e50',
                 justify='right', anchor='e',
                 wraplength=900).pack(fill='x')

        # ═══════════════════════════════════════════════════════
        # قسم: المسؤول عن الشركة
        # ═══════════════════════════════════════════════════════
        manager_section = tk.LabelFrame(
            content, text="  👨‍💼  المسؤول عن الشركة  ",
            font=('Arial', 14, 'bold'),
            bg='#ecf0f1', fg='#c0392b',
            padx=25, pady=20)
        manager_section.pack(fill='x', padx=40, pady=(0, 20))

        # ─── الاسم ───
        tk.Label(manager_section,
                 text="المحاسب / حسام الدين سيد ربيع",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#1a5276',
                 anchor='center').pack(fill='x', pady=(0, 5))

        # ─── المنصب ───
        tk.Label(manager_section,
                 text="المدير التنفيذي ومصمم البرامج",
                 font=('Arial', 12, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d',
                 anchor='center').pack(fill='x', pady=(0, 15))

        # ─── إطار بيانات التواصل ───
        contact_frame = tk.Frame(manager_section, bg='#ecf0f1')
        contact_frame.pack(fill='x')

        contacts = [
            ("📞  الهاتف / واتساب:", "01003796984"),
            ("✉️  البريد الإلكتروني:", "hussein@zahraa-soft.com"),
            ("📍  العنوان:", "مصر"),
        ]

        for label_text, value_text in contacts:
            row = tk.Frame(contact_frame, bg='#ecf0f1')
            row.pack(pady=3)

            # القيمة على اليسار
            tk.Label(row, text=value_text,
                     font=('Arial', 12, 'bold'),
                     bg='#ecf0f1', fg='#1a5276',
                     anchor='w').pack(side='left', padx=(0, 10))

            # التسمية على اليمين
            tk.Label(row, text=label_text,
                     font=('Arial', 12, 'bold'),
                     bg='#ecf0f1', fg='#2c3e50',
                     anchor='e').pack(side='right', padx=(10, 0))

        # ═══════════════════════════════════════════════════════
        # قسم: البرامج المصممة
        # ═══════════════════════════════════════════════════════
        programs_section = tk.LabelFrame(
            content, text="  💻  البرامج المصممة من الشركة  ",
            font=('Arial', 14, 'bold'),
            bg='#ecf0f1', fg='#16a085',
            padx=25, pady=20)
        programs_section.pack(fill='x', padx=40, pady=(0, 20))

        # ─── قائمة البرامج ───
        programs = [
            ("☀️", "Sunrise System",
             "نظام متكامل لإدارة المبيعات والخزينة والأرباح"),
            ("🍽️", "برنامج حسابات المطاعم",
             "نظام متكامل لإدارة المطاعم والمبيعات والمصاريف والمخزون"),
            ("🏪", "برنامج إدارة المحلات التجارية",
             "نظام شامل لإدارة المبيعات، المخزون، الفواتير، والعملاء"),
            ("📦", "برنامج إدارة المخازن",
             "نظام متطور لتتبع المخزون والجرد الشهري والتنبيهات"),
            ("💰", "برنامج الحسابات العامة",
             "نظام محاسبي متكامل للقيود والتقارير المالية"),
            ("👥", "برنامج إدارة الموارد البشرية",
             "نظام لإدارة العاملين، الرواتب، السلف، والحضور"),
            ("📊", "برنامج التقارير والتحليلات",
             "نظام تقارير ذكي مع رسوم بيانية وتصدير Excel"),
        ]

        for icon, name, desc in programs:
            row = tk.Frame(programs_section, bg='#ecf0f1')
            row.pack(fill='x', pady=6)

            # إطار ثابت العرض للأيقونة (لضمان المحاذاة)
            icon_holder = tk.Frame(row, bg='#ecf0f1', width=60, height=40)
            icon_holder.pack(side='right', padx=(0, 10))
            icon_holder.pack_propagate(False)

            tk.Label(icon_holder, text=icon,
                     font=('Arial', 20),
                     bg='#ecf0f1').pack(expand=True)

            # النصوص
            text_frame = tk.Frame(row, bg='#ecf0f1')
            text_frame.pack(side='right', fill='x', expand=True)

            tk.Label(text_frame, text=name,
                     font=('Arial', 13, 'bold'),
                     bg='#ecf0f1', fg='#2c3e50',
                     anchor='e').pack(fill='x')

            tk.Label(text_frame, text=desc,
                     font=('Arial', 11),
                     bg='#ecf0f1', fg='#7f8c8d',
                     anchor='e').pack(fill='x')

        # ═══════════════════════════════════════════════════════
        # زر معرّف الجهاز
        # ═══════════════════════════════════════════════════════
        machine_section = tk.LabelFrame(
            content, text="  🔑  معلومات الجهاز  ",
            font=('Arial', 12, 'bold'),
            bg='#ecf0f1', fg='#8e44ad',
            padx=15, pady=15)
        machine_section.pack(fill='x', padx=40, pady=(0, 20))

        tk.Button(machine_section,
                  text="🔑 الحصول على معرّف هذا الجهاز",
                  command=self._show_machine_id,
                  bg='#3498db', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8).pack()

        tk.Label(machine_section,
                 text="💡 يمكنك استخدام هذا المعرّف للتواصل مع الدعم الفني",
                 font=('Arial', 9, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=(8, 0))

        # ═══════════════════════════════════════════════════════
        # التذييل
        # ═══════════════════════════════════════════════════════
        footer = tk.Frame(content, bg='#2c3e50', height=60)
        footer.pack(fill='x', pady=(20, 0))
        footer.pack_propagate(False)

        tk.Label(footer,
                 text="© 2025 شركة الزهراء للبرمجيات — جميع الحقوق محفوظة",
                 font=('Arial', 11, 'bold'),
                 bg='#2c3e50', fg='#f1c40f').pack(pady=20)

    # ═══════════════════════════════════════════════════════════
    def _show_machine_id(self):
        """عرض معرّف الجهاز"""
        try:
            import subprocess
            import hashlib

            # حساب معرّف الجهاز
            try:
                cpu = subprocess.check_output(
                    'wmic cpu get ProcessorId',
                    shell=True, stderr=subprocess.DEVNULL
                ).decode().split()
                cpu_id = cpu[1] if len(cpu) > 1 else 'unknown'
            except:
                cpu_id = 'unknown'

            try:
                board = subprocess.check_output(
                    'wmic baseboard get SerialNumber',
                    shell=True, stderr=subprocess.DEVNULL
                ).decode().split()
                board_id = board[1] if len(board) > 1 else 'unknown'
            except:
                board_id = 'unknown'

            raw = f"{cpu_id}|{board_id}"
            hashed = hashlib.sha256(raw.encode()).hexdigest()[:16].upper()
            machine_id = '-'.join([hashed[i:i+4] for i in range(0, 16, 4)])

            msg = (
                f"🔑 معرّف هذا الجهاز:\n\n"
                f"{machine_id}\n\n"
                f"انسخه وأرسله للدعم الفني\n"
                f"📞 01003796984\n"
                f"✉️ hussein@zahraa-soft.com"
            )

            # نافذة عرض منسقة
            win = tk.Toplevel(self.winfo_toplevel())
            win.title("معرّف الجهاز")
            win.geometry("500x320")
            win.grab_set()
            win.configure(bg='#ecf0f1')

            tk.Label(win, text="🔑 معرّف هذا الجهاز",
                     font=('Arial', 16, 'bold'),
                     bg='#1a5276', fg='#f1c40f').pack(fill='x', ipady=15)

            tk.Label(win, text=machine_id,
                     font=('Consolas', 16, 'bold'),
                     bg='#fff9e6', fg='#c0392b',
                     relief='solid', bd=2).pack(pady=20, padx=30, ipady=15)

            tk.Label(win,
                     text="انسخه وأرسله للدعم الفني للحصول على الترخيص",
                     font=('Arial', 11, 'italic'),
                     bg='#ecf0f1', fg='#7f8c8d').pack(pady=5)

            btn_frame = tk.Frame(win, bg='#ecf0f1')
            btn_frame.pack(pady=15)

            def copy_id():
                win.clipboard_clear()
                win.clipboard_append(machine_id)
                messagebox.showinfo("تم النسخ",
                                     "✅ تم نسخ معرّف الجهاز")

            tk.Button(btn_frame, text="📋 نسخ",
                      command=copy_id,
                      bg='#3498db', fg='white',
                      font=('Arial', 11, 'bold'),
                      padx=20, pady=8).pack(side='left', padx=5)

            tk.Button(btn_frame, text="❌ إغلاق",
                      command=win.destroy,
                      bg='#e74c3c', fg='white',
                      font=('Arial', 11, 'bold'),
                      padx=20, pady=8).pack(side='left', padx=5)

        except Exception as e:
            messagebox.showerror("خطأ", f"تعذر الحصول على معرّف الجهاز: {e}")

    # ═══════════════════════════════════════════════════════════
    def load_traders(self):
        """يُستدعى تلقائيًا من MainApp (لا يحتاج تحديث)"""
        pass

# =============================================
# 3. الصفحة الرئيسية (التحكم)
# =============================================
# ═══════════════════════════════════════════════════════════════════
# MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════
class MainApp(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title("☀️ Sunrise System")
        self.geometry("1280x720")

        # ═══════════════════════════════════════════════════════
        # 1) الإعلان المتحرك
        # ═══════════════════════════════════════════════════════
        self.ad_full_text = COMPANY_AD * 6
        self.ad_display_length = AD_DISPLAY_LENGTH
        self.ad_index = 0
        self.is_ad_mode = True
        self.scroll_ad()

        # ═══════════════════════════════════════════════════════
        # 2) القائمة الجانبية
        # ═══════════════════════════════════════════════════════
        sidebar = tk.Frame(self, bg=SIDEBAR_BG, width=220)
        sidebar.pack(side='left', fill='y')
        sidebar.pack_propagate(False)

        # ─── العنوان ───
        title_frame = tk.Frame(sidebar, bg=SIDEBAR_BG, height=70)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text="📋  القائمة الرئيسية",
                 font=('Arial', 14, 'bold'),
                 bg=SIDEBAR_BG, fg=SIDEBAR_ACCENT).pack(pady=20)

        # ─── خط ذهبي ───
        tk.Frame(sidebar, bg=SIDEBAR_ACCENT, height=2).pack(fill='x', padx=15)

        # ─── قاموس الأزرار ───
        self.menu_buttons = {}
        self.active_page_key = None

        # ─── دالة إنشاء زر ───
        def make_menu_button(text, command, page_key=None):
            btn = tk.Button(
                sidebar,
                text=f"  {text}",
                command=command,
                bg=SIDEBAR_BTN_BG,
                fg=SIDEBAR_FG,
                font=('Arial', 11, 'bold'),
                relief='flat',
                bd=0,
                activebackground=SIDEBAR_BTN_HOVER,
                activeforeground='#1a5276',
                cursor='hand2',
                anchor='w',
                padx=10,
                pady=7,
            )
            btn.pack(pady=2, padx=10, fill='x')

            def on_enter(e):
                if self.active_page_key != page_key:
                    btn.config(bg=SIDEBAR_BTN_HOVER, fg='#1a5276')

            def on_leave(e):
                if self.active_page_key != page_key:
                    btn.config(bg=SIDEBAR_BTN_BG, fg=SIDEBAR_FG)

            btn.bind('<Enter>', on_enter)
            btn.bind('<Leave>', on_leave)

            if page_key:
                self.menu_buttons[page_key] = btn

            return btn

        # ═══════════════════════════════════════════════════════
        # 3) إنشاء الأزرار (نفس صفحات البرنامج الأصلي)
        # ═══════════════════════════════════════════════════════
        make_menu_button("📚  المبيعات",
                         lambda: self.show_page(salesPage),
                         page_key='salesPage')

        make_menu_button("💸  المصروفات",
                         lambda: self.show_page(paidsPage),
                         page_key='paidsPage')

        make_menu_button("📊  تقرير المبيعات عن مدة",
                         lambda: self.show_page(report_sales_page),
                         page_key='report_sales_page')

        make_menu_button("💰  تقرير رصيد النقدية",
                         lambda: self.show_page(CashBalancePage),
                         page_key='CashBalancePage')

        make_menu_button("📋  البيان الختامي",
                         lambda: self.show_page(FinalStatementPage),
                         page_key='FinalStatementPage')

        make_menu_button("💵  ترحيل الأرصدة",
                         lambda: self.show_page(CarriedBalancePage),
                         page_key='CarriedBalancePage')

        make_menu_button("🏦  الخزينة العامة",
                         lambda: self.show_page(GeneralTreasuryPage),
                         page_key='GeneralTreasuryPage')

        make_menu_button("🔄  مصاريف دورية",
                         lambda: self.show_page(RecurringExpensesPage),
                         page_key='RecurringExpensesPage')

        make_menu_button("💎  إدارة الأرباح",
                         lambda: self.show_page(ProfitTransferPage),
                         page_key='ProfitTransferPage')

        make_menu_button(
            "💰  الاموال المضافة",
            lambda: self.show_page(OwnerMoneyPage),
            page_key='OwnerMoneyPage'
        )
        make_menu_button(
            "🧑‍💼  حسابات التجار",
            lambda: self.show_page(TradersPage),
            page_key='TradersPage'
        )
        make_menu_button(
            "📊  التقرير الشامل",
            lambda: self.show_page(ComprehensiveReportPage),
            page_key='ComprehensiveReportPage'
        )

        make_menu_button(
            "ℹ️  عن الشركة",
            lambda: self.show_page(AboutPage),
            page_key='AboutPage'
        )
        # ═══════════════════════════════════════════════════════
        # 4) إطار عرض الصفحات
        # ═══════════════════════════════════════════════════════
        self.container = tk.Frame(self, bg='#ecf0f1')
        self.container.pack(side='right', fill='both', expand=True)

        self.pages = {}
        self.show_page(salesPage)
        # ✅ إغلاق شاشة التحميل
        close_splash()  # إغلاق فوري
        # ═══════════════════════════════════════════════════════
        # 5) تنسيق الجداول (Treeview)
        # ═══════════════════════════════════════════════════════
        style = ttk.Style()
        style.theme_use('clam')

        style.configure("Treeview.Heading",
                        font=('Arial', 11, 'bold'),
                        background='#2c3e50',
                        foreground='white',
                        relief='raised',
                        borderwidth=1)
        style.map("Treeview.Heading",
                  background=[('active', '#34495e')])

        style.configure("Treeview",
                        font=('Arial', 10),
                        rowheight=28,
                        background='white',
                        fieldbackground='white',
                        foreground='black',
                        borderwidth=0,
                        relief='flat')

        style.map("Treeview",
                  background=[('selected', '#3498db')],
                  foreground=[('selected', 'white')])

    # ═══════════════════════════════════════════════════════════
    # دوال التمرير والتمييز والعرض
    # ═══════════════════════════════════════════════════════════
    def scroll_ad(self):
        if not self.is_ad_mode:
            return

        global BANNER_AD_LABELS
        BANNER_AD_LABELS = [lbl for lbl in BANNER_AD_LABELS
                            if lbl.winfo_exists()]

        full_text = self.ad_full_text
        start = self.ad_index
        end = start + self.ad_display_length

        if end >= len(full_text):
            self.ad_index = 0
            start = self.ad_index
            end = start + self.ad_display_length

        display_part = full_text[start:end]

        for lbl in BANNER_AD_LABELS:
            try:
                lbl.config(text=display_part)
            except tk.TclError:
                pass

        self.ad_index += 1
        self.after(AD_SPEED, self.scroll_ad)

    def _highlight_active_button(self, active_key):
        self.active_page_key = active_key
        for key, btn in self.menu_buttons.items():
            try:
                if key == active_key:
                    btn.config(bg=SIDEBAR_BTN_HOVER, fg='#1a5276')
                else:
                    btn.config(bg=SIDEBAR_BTN_BG, fg=SIDEBAR_FG)
            except tk.TclError:
                pass

    def show_page(self, page_class):
        page_name = page_class.__name__

        if page_name not in self.pages:
            self.pages[page_name] = page_class(parent=self.container,
                                                controller=self)
            self.pages[page_name].pack(fill='both', expand=True)
        else:
            self.pages[page_name].pack(fill='both', expand=True)

        for name, page in self.pages.items():
            if name != page_name:
                page.pack_forget()

        self._highlight_active_button(page_name)

# =============================================
# 4. تشغيل التطبيق
# =============================================
if __name__ == '__main__':
    init_db()
    migrate_add_currency_to_trader_payments()   # ← أضف هذا
    app = MainApp()
    app.mainloop()